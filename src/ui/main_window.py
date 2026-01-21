from __future__ import annotations

from datetime import datetime
from pathlib import Path
import os
import sys
import time
import traceback
import tkinter as tk
from tkinter import filedialog
import customtkinter as ctk

import styles
from core.state import (
    AppPaths,
    AppState,
    SW_MAJOR_RELEASES,
    SW_MINOR_RELEASES,
    ME_VERSIONS,
    VEHICLE_NUMBERS,
    load_vehicle_catalog,
    save_state,
)
from services.canoe import (
    CANoeInstallation,
    connect_canoe,
    discover_canoe_installations,
    get_logging_block_status,
    is_canoe_running,
    load_canoe_config,
    open_canoe_installation,
    wait_for_process as _wait_for_process,
    _normalize_path_key,
    _get_active_canoe,
    _extract_major_from_text,
    _major_from_hint,
    _prog_id_exists,
    _attach_running_canoe,
    _spawn_canoe_instance,
)

class MainWindow(ctk.CTk):
    """
    Main application window.
    Handles UI, CANoe connection, logging start/stop,
    and operator comments tagged with timestamp.
    """

    def __init__(self, *, paths: AppPaths, state: AppState) -> None:
        super().__init__()

        self.paths = paths
        self.is_recording = False
        self.canoe = None  # COM object
        self.last_meas_running: bool | None = None  # last known Measurement.Running

        # Comment/log resolution state
        self.comment_file_path: Path | None = None
        self._current_log_folder: Path | None = None
        self._current_prefix: str | None = None  # e.g. "R300RC1_VEH123_tag_"
        self._resolve_tries: int = 0
        self._record_start_wallclock: float | None = None  # wall clock when Start was pressed
        self._comment_metadata_written: bool = False  # ensures metadata header written once

        # --- Window ---
        self.app_title_base = "anSWer Logging Hub"
        self.title(self.app_title_base)
        self.geometry("1160x760")
        self.minsize(820, 620)
        self.configure(bg=styles.Palette.BG)

        # --- Tk state variables (initialized from paths/state) ---
        self.canoe_config = tk.StringVar(value=str(paths.cfg_file) if paths.cfg_file else "")
        self.log_tag = tk.StringVar(value=state.tag or "")
        initial_sw_rel = (state.sw_rel or "").strip()
        major_release, minor_release = self._split_sw_release(initial_sw_rel)
        self.sw_major_var = tk.StringVar(value=major_release)
        self.sw_minor_var = tk.StringVar(value=minor_release)
        self.sw_rel = tk.StringVar(value=self._compose_sw_release(major_release, minor_release))
        self.me_version_var = tk.StringVar(value=state.me_version or (ME_VERSIONS[0] if ME_VERSIONS else ""))
        self.vehicle_id = tk.StringVar(value=state.vehicle_id or "")
        self.vehicle_id.trace_add("write", lambda *_: self._update_titles_with_release())
        self.vehicle_catalog = load_vehicle_catalog(self.paths.root)
        self._vehicle_label_to_id: dict[str, str] = {}
        self.canoe_install_var = tk.StringVar(value="")
        self._installations_by_label: dict[str, CANoeInstallation] = {}
        self._exec_key_to_label: dict[str, str] = {}
        self._initial_canoe_exec = state.canoe_exec or ""
        default_log_dir = (state.log_dir or "").strip() or str(self.paths.log_dir)
        self.log_dir_var = tk.StringVar(value=default_log_dir)
        self._record_timer_var = tk.StringVar(value="Recording time: --:--:--.---")
        self._camera_mode_var = tk.StringVar(value="Camera mode: --")
        self._ethernet_status_var = tk.StringVar(value="Ethernet: --")
        self._flexray_status_var = tk.StringVar(value="Flexray: --")
        self._ethernet_drops_var = tk.StringVar(value="Ethernet drops: --")
        self._flexray_drops_var = tk.StringVar(value="Flexray drops: --")
        self._hint_popup: ctk.CTkToplevel | None = None
        self._theme_mode: str = "neutral"
        self._theme_cards: list[ctk.CTkFrame] = []
        self._theme_entries: list[ctk.CTkEntry] = []
        self._theme_menus: list[ctk.CTkOptionMenu] = []
        self._theme_textboxes: list[ctk.CTkTextbox] = []

        # Build UI
        self._build_body()
        self._update_titles_with_release()
        self._refresh_canoe_installations(preferred_exec=self._initial_canoe_exec or None)
        self._install_exception_hooks()

        # Focus window
        self.after(0, self.focus_set)
        self.update_idletasks()
        self.lift()

        # Start periodic polling of CANoe measurement state
        self.after(500, self._sync_measurement_ui)
        self.after(1500, self._process_poll_tick)

    def _install_exception_hooks(self) -> None:
        def handle_exception(exc_type, exc_value, exc_tb):
            self._log_exception(exc_type, exc_value, exc_tb)
            if self._prev_excepthook is not None:
                try:
                    self._prev_excepthook(exc_type, exc_value, exc_tb)
                except Exception:
                    pass

        self._prev_excepthook = getattr(sys, "excepthook", None)
        sys.excepthook = handle_exception

    def _log_exception(self, exc_type, exc_value, exc_tb) -> None:
        trace = "".join(traceback.format_exception(exc_type, exc_value, exc_tb)).strip()
        self._debug_log(f"Unhandled exception:\n{trace}")
        self._set_status("❌ Python error (see debug log)", tone="danger")

    def report_callback_exception(self, exc_type, exc_value, exc_tb) -> None:
        self._log_exception(exc_type, exc_value, exc_tb)

    # -------------------- UI building --------------------
    def _build_body(self) -> None:
        """
        Build a responsive card-based layout:
        - Hero row with title and status pill
        - CANoe connection card
        - Session metadata card
        - Measurement controls
        - Comment notebook
        """
        pad_x, pad_y = styles.Metrics.PAD_X, styles.Metrics.PAD_Y
        column_gap = max(8, pad_x // 2)

        self.body = ctk.CTkScrollableFrame(
            self,
            fg_color=styles.Palette.BG,
            scrollbar_fg_color=styles.Palette.BG,
            scrollbar_button_color=styles.Palette.BG,
            scrollbar_button_hover_color=styles.Palette.BG,
        )
        self.body.pack(fill="both", expand=True, padx=max(4, pad_x // 3), pady=max(4, pad_y // 3))

        self.body.grid_columnconfigure((0, 1), weight=1, uniform="col", minsize=360)
        self.body.grid_rowconfigure(4, weight=1)
        self.body.grid_rowconfigure(6, weight=1)

        # ---- Hero row ----
        hero = styles.card(self.body)
        hero.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, pad_y))
        hero.grid_columnconfigure(0, weight=1)

        hero_header = ctk.CTkFrame(hero, fg_color="transparent")
        hero_header.grid(row=0, column=0, sticky="w", padx=pad_x, pady=(pad_y, 2))
        hero_header.grid_columnconfigure(0, weight=0)

        self.hero_title_label = ctk.CTkLabel(hero_header, text=self._app_title())
        styles.style_label(self.hero_title_label, kind="title")
        self.hero_title_label.grid(row=0, column=0, sticky="w")

        hero_hint = self._create_hint_icon(
            hero_header,
            "Control CANoe sessions, tag logs, and capture operator context.",
        )
        hero_hint.grid(row=0, column=1, sticky="w", padx=(6, 0))

        self.status = ctk.CTkLabel(
            hero,
            text="",
            width=180,
            height=30,
            corner_radius=styles.Metrics.RADIUS_LG,
            anchor="center",
            font=styles.Fonts.BUTTON,
        )
        self.status.grid(row=0, column=1, rowspan=2, sticky="e", padx=pad_x, pady=pad_y)
        self._set_status("Not connected", tone="muted")

        # ---- CANoe connection card ----
        connect_card = styles.card(self.body)
        connect_card.grid(row=1, column=0, sticky="nsew", padx=(0, column_gap), pady=(0, pad_y))
        connect_card.grid_columnconfigure(0, weight=1)

        conn_header = ctk.CTkFrame(connect_card, fg_color="transparent")
        conn_header.grid(row=0, column=0, sticky="w", padx=pad_x, pady=(pad_y, 2))

        conn_title = ctk.CTkLabel(conn_header, text="CANoe session")
        styles.style_label(conn_title, kind="section")
        conn_title.grid(row=0, column=0, sticky="w")

        conn_hint = self._create_hint_icon(
            conn_header,
            "Select any installed CANoe version and keep the desired configuration ready.",
        )
        conn_hint.grid(row=0, column=1, sticky="w", padx=(6, 0))

        install_label = ctk.CTkLabel(connect_card, text="CANoe installation")
        styles.style_label(install_label, kind="hint")
        install_label.grid(row=1, column=0, sticky="w", padx=pad_x, pady=(0, 6))

        install_row = ctk.CTkFrame(connect_card, fg_color="transparent")
        install_row.grid(row=2, column=0, sticky="ew", padx=pad_x, pady=(0, pad_y))
        install_row.grid_columnconfigure(0, weight=1)
        install_row.grid_columnconfigure(1, weight=0)

        self.install_dropdown = ctk.CTkOptionMenu(
            install_row,
            variable=self.canoe_install_var,
            values=["Searching..."],
            command=self._on_canoe_version_change,
        )
        self.install_dropdown.set("Searching...")
        styles.style_option_menu(self.install_dropdown, roundness="md")
        self.install_dropdown.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self._theme_menus.append(self.install_dropdown)

        self.btn_refresh_install = ctk.CTkButton(
            install_row,
            text="Refresh",
            command=lambda: self._refresh_canoe_installations(
                preferred_exec=self._selected_canoe_exec_string() or None
            ),
        )
        styles.style_button(self.btn_refresh_install, variant="neutral", size="sm", roundness="md")
        self.btn_refresh_install.grid(row=0, column=1, sticky="e")

        button_row = ctk.CTkFrame(connect_card, fg_color="transparent")
        button_row.grid(row=3, column=0, sticky="ew", padx=pad_x, pady=(0, pad_y))
        button_row.grid_columnconfigure(0, weight=1)

        self.btn_launch = ctk.CTkButton(
            button_row,
            text="Launch CANoe",
            command=self._open_or_connect_canoe,
        )
        styles.style_button(self.btn_launch, variant="primary", size="lg", roundness="lg")
        self.btn_launch.grid(row=0, column=0, sticky="ew")

        cfg_label = ctk.CTkLabel(connect_card, text="Configuration file")
        styles.style_label(cfg_label, kind="hint")
        cfg_label.grid(row=4, column=0, sticky="w", padx=pad_x, pady=(pad_y, 6))

        row_cfg = ctk.CTkFrame(connect_card, fg_color="transparent")
        row_cfg.grid(row=5, column=0, sticky="ew", padx=pad_x, pady=(0, pad_y))
        row_cfg.grid_columnconfigure(0, weight=1)
        row_cfg.grid_columnconfigure(1, weight=0)

        entry_cfg = ctk.CTkEntry(row_cfg, textvariable=self.canoe_config)
        styles.style_entry(entry_cfg, roundness="md")
        entry_cfg.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self._theme_entries.append(entry_cfg)

        btn_browse = ctk.CTkButton(row_cfg, text="Browse…", command=self._choose_cfg)
        styles.style_button(btn_browse, variant="neutral", size="sm", roundness="md")
        btn_browse.grid(row=0, column=1, sticky="e")

        # ---- Session metadata ----
        session_card = styles.card(self.body)
        session_card.grid(row=1, column=1, sticky="nsew", padx=(column_gap, 0), pady=(0, pad_y))
        session_card.grid_columnconfigure(0, weight=1)

        session_header = ctk.CTkFrame(session_card, fg_color="transparent")
        session_header.grid(row=0, column=0, sticky="w", padx=pad_x, pady=(pad_y, 2))

        session_title = ctk.CTkLabel(session_header, text="Session metadata")
        styles.style_label(session_title, kind="section")
        session_title.grid(row=0, column=0, sticky="w")

        session_hint = self._create_hint_icon(
            session_header,
            "Keep naming consistent for every measurement.",
        )
        session_hint.grid(row=0, column=1, sticky="w", padx=(6, 0))

        info_frame = ctk.CTkFrame(session_card, fg_color="transparent")
        info_frame.grid(row=1, column=0, sticky="ew", padx=pad_x, pady=(0, pad_y))
        info_frame.grid_columnconfigure(1, weight=1)
        field_gap_x = 12

        lbl_tag = ctk.CTkLabel(info_frame, text="Recording tag")
        styles.style_label(lbl_tag, kind="hint")
        lbl_tag.grid(row=0, column=0, sticky="w", padx=(0, field_gap_x), pady=(0, 6))
        entry_tag = ctk.CTkEntry(info_frame, textvariable=self.log_tag)
        styles.style_entry(entry_tag, roundness="md")
        entry_tag.grid(row=0, column=1, sticky="ew", pady=(0, 6))
        self._theme_entries.append(entry_tag)

        lbl_vehicle = ctk.CTkLabel(info_frame, text="Vehicle")
        styles.style_label(lbl_vehicle, kind="hint")
        lbl_vehicle.grid(row=1, column=0, sticky="w", padx=(0, field_gap_x), pady=(0, 6))

        if self.vehicle_catalog:
            self._vehicle_label_to_id = {
                self._format_vehicle_option_label(veh_id, model): veh_id
                for veh_id, model in self.vehicle_catalog.items()
            }
            vehicle_values = list(self._vehicle_label_to_id.keys())
            default_vehicle_id = self._resolve_initial_vehicle_id()
            default_label = self._format_vehicle_option_label(
                default_vehicle_id,
                self.vehicle_catalog.get(default_vehicle_id, ""),
            )
            if default_label not in self._vehicle_label_to_id:
                self._vehicle_label_to_id = {
                    default_label: default_vehicle_id,
                    **self._vehicle_label_to_id,
                }
                vehicle_values = list(self._vehicle_label_to_id.keys())

            self.vehicle_dropdown_var = tk.StringVar(value=default_label)
            entry_vehicle = ctk.CTkOptionMenu(
                info_frame,
                variable=self.vehicle_dropdown_var,
                values=vehicle_values,
                command=self._on_vehicle_dropdown_change,
            )
            entry_vehicle.set(default_label)
            styles.style_option_menu(entry_vehicle, roundness="md")
        else:
            entry_vehicle = ctk.CTkEntry(info_frame, textvariable=self.vehicle_id)
            styles.style_entry(entry_vehicle, roundness="md")
        entry_vehicle.grid(row=1, column=1, sticky="ew", pady=(0, 6))
        if isinstance(entry_vehicle, ctk.CTkOptionMenu):
            self._theme_menus.append(entry_vehicle)
        else:
            self._theme_entries.append(entry_vehicle)

        lbl_release = ctk.CTkLabel(info_frame, text="SW release")
        styles.style_label(lbl_release, kind="hint")
        lbl_release.grid(row=2, column=0, sticky="w", padx=(0, field_gap_x), pady=(0, 6))

        release_row = ctk.CTkFrame(info_frame, fg_color="transparent")
        release_row.grid(row=2, column=1, sticky="ew", pady=(0, 6))
        release_row.grid_columnconfigure((0, 1), weight=1)

        major_dropdown = ctk.CTkOptionMenu(
            release_row,
            variable=self.sw_major_var,
            values=SW_MAJOR_RELEASES,
            command=self._on_sw_release_change,
        )
        major_dropdown.set(self.sw_major_var.get())
        styles.style_option_menu(major_dropdown, roundness="md")
        major_dropdown.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self._theme_menus.append(major_dropdown)

        minor_dropdown = ctk.CTkOptionMenu(
            release_row,
            variable=self.sw_minor_var,
            values=SW_MINOR_RELEASES,
            command=self._on_sw_release_change,
        )
        minor_dropdown.set(self.sw_minor_var.get())
        styles.style_option_menu(minor_dropdown, roundness="md")
        minor_dropdown.grid(row=0, column=1, sticky="ew")
        self._theme_menus.append(minor_dropdown)

        lbl_me_version = ctk.CTkLabel(info_frame, text="ME version")
        styles.style_label(lbl_me_version, kind="hint")
        lbl_me_version.grid(row=3, column=0, sticky="w", padx=(0, field_gap_x), pady=(0, 6))

        me_version_menu = ctk.CTkOptionMenu(
            info_frame,
            variable=self.me_version_var,
            values=ME_VERSIONS,
        )
        me_version_menu.set(self.me_version_var.get())
        styles.style_option_menu(me_version_menu, roundness="md")
        me_version_menu.grid(row=3, column=1, sticky="ew", pady=(0, 6))
        self._theme_menus.append(me_version_menu)

        # ---- Measurement controls ----
        action_card = styles.card(self.body)
        action_card.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, pad_y))
        action_card.grid_columnconfigure(0, weight=1)

        action_title = ctk.CTkLabel(action_card, text="Measurement control")
        styles.style_label(action_title, kind="section")
        action_title.grid(row=0, column=0, sticky="w", padx=pad_x, pady=(pad_y, 4))

        log_dir_row = ctk.CTkFrame(action_card, fg_color="transparent")
        log_dir_row.grid(row=1, column=0, sticky="ew", padx=pad_x, pady=(0, 6))
        log_dir_row.grid_columnconfigure(1, weight=1)

        log_dir_label = ctk.CTkLabel(log_dir_row, text="Log folder")
        styles.style_label(log_dir_label, kind="body")
        log_dir_label.grid(row=0, column=0, sticky="w", padx=(0, pad_x // 2))

        log_dir_entry = ctk.CTkEntry(log_dir_row, textvariable=self.log_dir_var)
        styles.style_entry(log_dir_entry, roundness="md")
        log_dir_entry.grid(row=0, column=1, sticky="ew", padx=(0, pad_x // 2))
        log_dir_entry.bind("<FocusOut>", self._on_log_dir_entry_commit)
        log_dir_entry.bind("<Return>", self._on_log_dir_entry_commit)
        self._theme_entries.append(log_dir_entry)

        browse_log_btn = ctk.CTkButton(
            log_dir_row,
            text="Browse",
            command=self._choose_log_dir,
            width=100,
        )
        styles.style_button(browse_log_btn, variant="neutral", size="sm", roundness="md")
        browse_log_btn.grid(row=0, column=2, sticky="e")

        action_hint = ctk.CTkLabel(
            action_card,
            text="Start/stop CANoe logging once metadata is ready.",
        )
        styles.style_label(action_hint, kind="hint")
        action_hint.grid(row=2, column=0, sticky="w", padx=pad_x, pady=(0, 4))

        action_btn_row = ctk.CTkFrame(action_card, fg_color="transparent")
        action_btn_row.grid(row=3, column=0, sticky="ew", padx=pad_x, pady=(pad_y, pad_y // 2))
        action_btn_row.grid_columnconfigure(0, weight=9)
        action_btn_row.grid_columnconfigure(1, weight=1)

        self.btn_record = ctk.CTkButton(
            action_btn_row,
            text="Start recording",
            command=self._on_start_stop_click,
            state="disabled",
        )
        styles.style_button(self.btn_record, variant="neutral", size="lg", roundness="lg")
        self.btn_record.grid(row=0, column=0, sticky="ew", padx=(0, pad_x // 2))

        self.btn_discard = ctk.CTkButton(
            action_btn_row,
            text="Discard recording",
            command=self._on_discard_click,
            state="disabled",
        )
        styles.style_button(self.btn_discard, variant="neutral", size="lg", roundness="lg")
        self.btn_discard.grid(row=0, column=1, sticky="ew", padx=(pad_x // 2, 0))

        # ---- Recording status ----
        self.status_card = styles.card(self.body)
        self.status_card.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, pad_y))
        self.status_card.grid_columnconfigure(0, weight=1)

        status_title = ctk.CTkLabel(self.status_card, text="Recording status")
        styles.style_label(status_title, kind="section")
        status_title.grid(row=0, column=0, sticky="w", padx=pad_x, pady=(pad_y, 4))

        status_row = ctk.CTkFrame(self.status_card, fg_color="transparent")
        status_row.grid(row=1, column=0, sticky="ew", padx=pad_x, pady=(0, pad_y))
        status_row.grid_columnconfigure(0, weight=1)
        status_row.grid_columnconfigure(1, weight=1)
        status_row.grid_columnconfigure(2, weight=1)
        status_row.grid_columnconfigure(3, weight=1)
        status_row.grid_rowconfigure(0, weight=1)
        status_row.grid_rowconfigure(1, weight=0)

        self.record_timer_label = ctk.CTkLabel(
            status_row,
            textvariable=self._record_timer_var,
            anchor="center",
        )
        styles.style_label(self.record_timer_label, kind="body")
        self.record_timer_label.grid(row=0, column=0, sticky="nsew")

        self.camera_mode_label = ctk.CTkLabel(
            status_row,
            textvariable=self._camera_mode_var,
            anchor="center",
        )
        styles.style_label(self.camera_mode_label, kind="body")
        self.camera_mode_label.configure(font=styles.Fonts.BODY_BOLD)
        self.camera_mode_label.grid(row=0, column=1, sticky="nsew")

        self.ethernet_status_label = ctk.CTkLabel(
            status_row,
            textvariable=self._ethernet_status_var,
            anchor="center",
        )
        styles.style_label(self.ethernet_status_label, kind="body")
        self.ethernet_status_label.configure(font=styles.Fonts.BODY_BOLD)
        self.ethernet_status_label.grid(row=0, column=2, sticky="nsew")

        self.flexray_status_label = ctk.CTkLabel(
            status_row,
            textvariable=self._flexray_status_var,
            anchor="center",
        )
        styles.style_label(self.flexray_status_label, kind="body")
        self.flexray_status_label.configure(font=styles.Fonts.BODY_BOLD)
        self.flexray_status_label.grid(row=0, column=3, sticky="nsew")

        self.ethernet_drops_label = ctk.CTkLabel(
            status_row,
            textvariable=self._ethernet_drops_var,
            anchor="center",
        )
        styles.style_label(self.ethernet_drops_label, kind="body")
        self.ethernet_drops_label.grid(row=1, column=2, sticky="nsew", pady=(2, 0))

        self.flexray_drops_label = ctk.CTkLabel(
            status_row,
            textvariable=self._flexray_drops_var,
            anchor="center",
        )
        styles.style_label(self.flexray_drops_label, kind="body")
        self.flexray_drops_label.grid(row=1, column=3, sticky="nsew", pady=(2, 0))

        # ---- Comment workspace ----
        self.comment_card = styles.card(self.body)
        self.comment_card.grid(row=4, column=0, columnspan=2, sticky="nsew")
        self.comment_card.grid_columnconfigure(0, weight=1)
        self.comment_card.grid_rowconfigure(1, weight=1)

        comment_header = ctk.CTkFrame(self.comment_card, fg_color="transparent")
        comment_header.grid(row=0, column=0, sticky="w", padx=pad_x, pady=(pad_y, 2))

        comment_title = ctk.CTkLabel(comment_header, text="Operator notes")
        styles.style_label(comment_title, kind="section")
        comment_title.grid(row=0, column=0, sticky="w")

        comment_hint = self._create_hint_icon(
            comment_header,
            "Capture context while the measurement runs. Enter = save, Shift+Enter = newline.",
        )
        comment_hint.grid(row=0, column=1, sticky="w", padx=(6, 0))

        self.comment_box = ctk.CTkTextbox(self.comment_card, height=80, wrap="word")
        styles.style_textbox(self.comment_box, roundness="lg")
        self.comment_box.grid(row=1, column=0, sticky="nsew", padx=pad_x, pady=(0, pad_y))
        self.comment_box.bind("<Return>", self._on_comment_enter)
        self._theme_textboxes.append(self.comment_box)

        self.btn_save_comment = ctk.CTkButton(
            self.comment_card,
            text="Save comment",
            command=self._on_save_comment_click,
            state="disabled",
        )
        styles.style_button(self.btn_save_comment, variant="neutral", size="md", roundness="md")
        self.btn_save_comment.grid(row=2, column=0, sticky="e", padx=pad_x, pady=(0, pad_y // 2))

        # ---- Debug console toggle ----
        debug_toggle_row = ctk.CTkFrame(self.body, fg_color="transparent")
        debug_toggle_row.grid(row=5, column=0, columnspan=2, sticky="ew", padx=pad_x, pady=(pad_y // 2, pad_y // 2))
        debug_toggle_row.grid_columnconfigure(0, weight=1)
        debug_toggle_row.grid_columnconfigure(1, weight=0)

        self.debug_toggle_btn = ctk.CTkButton(
            debug_toggle_row,
            text="? Debug log",
            command=self._toggle_debug_panel,
            width=170,
        )
        styles.style_button(self.debug_toggle_btn, variant="neutral", size="sm", roundness="md")
        self.debug_toggle_btn.grid(row=0, column=0, sticky="ew")

        self.btn_copy_debug = ctk.CTkButton(
            debug_toggle_row,
            text="Copy",
            width=80,
            command=self._copy_debug_log,
        )
        styles.style_button(self.btn_copy_debug, variant="neutral", size="sm", roundness="md")
        self.btn_copy_debug.grid(row=0, column=1, sticky="e", padx=(6, 0))

        self.debug_card = styles.card(self.body)
        self.debug_card.grid(row=6, column=0, columnspan=2, sticky="nsew", pady=(pad_y // 2, 0))
        self.debug_card.grid_rowconfigure(1, weight=1)
        self.debug_card.grid_columnconfigure(0, weight=1)

        debug_header = ctk.CTkFrame(self.debug_card, fg_color="transparent")
        debug_header.grid(row=0, column=0, sticky="ew", padx=pad_x, pady=(pad_y, 4))
        debug_header.grid_columnconfigure(0, weight=1)

        debug_title = ctk.CTkLabel(debug_header, text="Debug log")
        styles.style_label(debug_title, kind="section")
        debug_title.grid(row=0, column=0, sticky="w")

        self.btn_clear_debug = ctk.CTkButton(
            debug_header,
            text="Clear",
            width=80,
            command=self._clear_debug_log,
        )
        styles.style_button(self.btn_clear_debug, variant="neutral", size="sm", roundness="md")
        self.btn_clear_debug.grid(row=0, column=1, sticky="e")

        self.debug_text = ctk.CTkTextbox(self.debug_card, height=110, wrap="word")
        styles.style_textbox(self.debug_text, roundness="md")
        self.debug_text.grid(row=1, column=0, sticky="nsew", padx=pad_x, pady=(0, pad_y))
        self.debug_text.configure(state="disabled")
        self._theme_textboxes.append(self.debug_text)
        self.debug_panel_visible = True
        self._toggle_debug_panel(force_state=False)

        self._theme_cards = [
            hero,
            connect_card,
            session_card,
            action_card,
            self.status_card,
            self.comment_card,
            self.debug_card,
        ]
        self._apply_overall_theme("neutral")

    def _create_hint_icon(self, master, text: str) -> ctk.CTkButton:
        """
        Compact info icon that shows contextual help on click.
        """
        btn = ctk.CTkButton(
            master,
            text="i",
            width=20,
            height=20,
            corner_radius=styles.Metrics.RADIUS_SM,
            fg_color="transparent",
            hover_color=styles.Palette.CARD_BORDER,
            text_color=styles.Palette.MUTED,
            font=styles.Fonts.CAPTION,
        )
        btn.configure(border_width=0, width=20, height=20)
        btn.bind("<ButtonPress-1>", lambda _e, msg=text, widget=btn: self._show_hint_tooltip(msg, widget))
        btn.bind("<ButtonRelease-1>", lambda _e: self._hide_hint_tooltip())
        btn.bind("<Leave>", lambda _e: self._hide_hint_tooltip())
        return btn

    def _show_hint_tooltip(self, text: str, widget: tk.Widget) -> None:
        self._hide_hint_tooltip()
        popup = ctk.CTkToplevel(self)
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)
        popup.configure(fg_color=styles.Palette.CARD_DARK)
        label = ctk.CTkLabel(
            popup,
            text=text,
            wraplength=260,
            justify="left",
            font=styles.Fonts.BODY,
        )
        label.pack(padx=10, pady=8)
        popup.update_idletasks()
        width = popup.winfo_reqwidth()
        height = popup.winfo_reqheight()
        x = widget.winfo_rootx()
        y = widget.winfo_rooty() - height - 8
        if y < 0:
            y = widget.winfo_rooty() + widget.winfo_height() + 8
        popup.geometry(f"{width}x{height}+{int(x)}+{int(y)}")
        self._hint_popup = popup

    def _hide_hint_tooltip(self) -> None:
        if self._hint_popup is not None:
            try:
                self._hint_popup.destroy()
            except Exception:
                pass
            self._hint_popup = None

    def _set_status(self, text: str, tone: str = "muted") -> None:
        """
        Update the status pill with contextual color coding.
        """
        palette = {
            "muted": (styles.Palette.STATUS_MUTED_BG, styles.Palette.MUTED),
            "success": (styles.Palette.STATUS_SUCCESS_BG, styles.Palette.SUCCESS),
            "warning": (styles.Palette.STATUS_WARNING_BG, styles.Palette.WARNING),
            "danger": (styles.Palette.STATUS_DANGER_BG, styles.Palette.DANGER),
            "info": (styles.Palette.STATUS_INFO_BG, styles.Palette.PRIMARY),
        }
        fg_color, text_color = palette.get(tone, palette["muted"])
        label = getattr(self, "status", None)
        if label is not None:
            label.configure(text=text, fg_color=fg_color, text_color=text_color)

    def _apply_overall_theme(self, mode: str) -> None:
        if mode == self._theme_mode:
            return

        bg = styles.Palette.BG
        card = styles.Palette.CARD_DARK
        card_alt = styles.Palette.CARD_DARK_ALT
        border = styles.Palette.CARD_BORDER
        input_bg = styles.Palette.INPUT_BG
        input_border = styles.Palette.INPUT_BORDER

        self.configure(bg=bg)
        if getattr(self, "body", None) is not None:
            self.body.configure(fg_color=bg)

        for card_frame in self._theme_cards:
            card_frame.configure(fg_color=(card, card_alt), border_color=border)

        for entry in self._theme_entries:
            entry.configure(fg_color=(input_bg, input_bg), border_color=input_border)

        menu_hover = styles._darken_hex(input_bg, 0.08)
        for menu in self._theme_menus:
            menu.configure(
                fg_color=(input_bg, input_bg),
                button_color=input_bg,
                button_hover_color=menu_hover,
                dropdown_fg_color=card_alt,
            )

        for textbox in self._theme_textboxes:
            textbox.configure(fg_color=(input_bg, input_bg), border_color=input_border)

        self._theme_mode = mode

    def _debug_log(self, message: str) -> None:
        """
        Append a timestamped debug line to the on-screen console and stdout.
        """
        timestamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}] {message}"
        print(f"[DEBUG] {line}")
        widget = getattr(self, "debug_text", None)
        if widget is None:
            return
        widget.configure(state="normal")
        widget.insert("end", f"{line}\n")
        widget.see("end")
        widget.configure(state="disabled")

    def _clear_debug_log(self) -> None:
        widget = getattr(self, "debug_text", None)
        if widget is None:
            return
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.configure(state="disabled")
        self._debug_log("Debug log cleared")

    def _copy_debug_log(self) -> None:
        widget = getattr(self, "debug_text", None)
        if widget is None:
            return
        text = widget.get("1.0", "end").strip()
        if not text:
            self._set_status("⚠️ Debug log is empty", tone="warning")
            return
        try:
            self.clipboard_clear()
            self.clipboard_append(text)
            self.update_idletasks()
            self._set_status("✅ Debug log copied", tone="success")
        except Exception as exc:
            self._set_status(f"❌ Copy failed: {exc}", tone="danger")

    def _toggle_debug_panel(self, force_state: bool | None = None) -> None:
        if not hasattr(self, "debug_card"):
            return
        current = getattr(self, "debug_panel_visible", False)
        new_state = (not current) if force_state is None else bool(force_state)
        self.debug_panel_visible = new_state
        if new_state:
            self.debug_card.grid()
            if hasattr(self, "debug_toggle_btn"):
                self.debug_toggle_btn.configure(text="▼ Debug log")
        else:
            self.debug_card.grid_remove()
            if hasattr(self, "debug_toggle_btn"):
                self.debug_toggle_btn.configure(text="▶ Debug log")

    def _compose_sw_release(self, major: str, minor: str) -> str:
        return f"{(major or '').strip().upper()}{(minor or '').strip().upper()}".strip()

    def _split_sw_release(self, combined: str) -> tuple[str, str]:
        default_major = SW_MAJOR_RELEASES[0] if SW_MAJOR_RELEASES else ""
        default_minor = SW_MINOR_RELEASES[0] if SW_MINOR_RELEASES else ""
        raw = (combined or "").strip().upper()
        major = default_major
        minor = default_minor
        for candidate in SW_MAJOR_RELEASES:
            cand_upper = candidate.strip().upper()
            if cand_upper and raw.startswith(cand_upper):
                major = candidate
                remainder = raw[len(cand_upper):]
                if remainder in SW_MINOR_RELEASES:
                    minor = remainder
                break
        return major or "", minor or ""

    def _vehicle_descriptor(self, vehicle_id: str | None = None) -> str:
        vid = (vehicle_id or self.vehicle_id.get() or "").strip()
        return (self.vehicle_catalog.get(vid) or "").strip() if self.vehicle_catalog else ""

    def _vehicle_model_tag(self, include_id_fallback: bool = True) -> str:
        """
        Build a compact vehicle token like 'XC60_Veh6'.
        Includes descriptor + fleet number; falls back to vehicle ID if needed.
        """
        vehicle_id = (self.vehicle_id.get() or "").strip()
        desc = self._vehicle_descriptor(vehicle_id)
        number_tag = self._vehicle_number_tag(vehicle_id)
        sanitize = lambda value: (value or "").replace(" ", "_")

        parts: list[str] = []
        if desc:
            parts.append(sanitize(desc))
        if number_tag:
            parts.append(number_tag)
        elif include_id_fallback and vehicle_id:
            parts.append(sanitize(vehicle_id))

        return "_".join([p for p in parts if p])

    def _app_title(self) -> str:
        rel = (self.sw_rel.get() or "").strip()
        model = self._vehicle_model_tag(include_id_fallback=False)
        parts = [self.app_title_base]
        if rel:
            parts.append(rel)
        if model:
            parts.append(model)
        return " - ".join(parts)

    def _update_titles_with_release(self) -> None:
        title_text = self._app_title()
        self.title(title_text)
        if hasattr(self, "hero_title_label"):
            self.hero_title_label.configure(text=title_text)

    def _on_sw_release_change(self, _selection: str | None = None) -> None:
        self.sw_rel.set(self._compose_sw_release(self.sw_major_var.get(), self.sw_minor_var.get()))
        self._persist_state_snapshot()
        self._update_titles_with_release()

    def _resolve_initial_vehicle_id(self) -> str:
        current = (self.vehicle_id.get() or "").strip()
        if current and (not self.vehicle_catalog or current in self.vehicle_catalog):
            return current
        if self.vehicle_catalog:
            fallback = next(iter(self.vehicle_catalog.keys()))
            self.vehicle_id.set(fallback)
            return fallback
        return current

    def _format_vehicle_option_label(self, vehicle_id: str, descriptor: str | None) -> str:
        vid = (vehicle_id or "").strip()
        desc = (descriptor or "").strip()
        number = VEHICLE_NUMBERS.get(vid.upper()) if vid else None

        parts: list[str] = []
        if number is not None:
            parts.append(f"Vehicle {number}")
            if vid:
                parts.append(vid)
        elif vid:
            parts.append(vid)

        if desc:
            parts.append(desc)

        return " - ".join(parts)

    def _on_vehicle_dropdown_change(self, selection: str) -> None:
        vehicle_id = self._vehicle_label_to_id.get(selection, selection)
        self.vehicle_id.set((vehicle_id or "").strip())
        self._persist_state_snapshot()
        self._update_titles_with_release()

    def _vehicle_number_tag(self, vehicle_id: str) -> str:
        normalized = (vehicle_id or "").strip().upper()
        number = VEHICLE_NUMBERS.get(normalized)
        return f"Veh{number}" if number is not None else ""

    def _vehicle_prefix_component(self) -> str:
        """
        Token used in file/folder names. Prefer descriptor + fleet number (e.g. XC60_Veh6).
        Falls back to vehicle ID if descriptor/number are missing.
        """
        return self._vehicle_model_tag(include_id_fallback=True)

    def _gather_state(self) -> AppState:
        return AppState(
            cfg_file=self.canoe_config.get(),
            tag=self.log_tag.get(),
            sw_rel=self.sw_rel.get(),
            me_version=self.me_version_var.get(),
            vehicle_id=self.vehicle_id.get(),
            canoe_exec=self._selected_canoe_exec_string() or "",
            log_dir=self.log_dir_var.get(),
        )

    def _persist_state_snapshot(self) -> None:
        save_state(state=self._gather_state(), paths=self.paths)

    # -------------------- Polling / UI sync --------------------
    def _read_sysvar_value(self, fieldname: str) -> str | None:
        if self.canoe is None:
            return None
        parts = [part for part in (fieldname or "").split("::") if part]
        if len(parts) < 2:
            return None
        var_name = parts[-1]
        namespace_chain = parts[:-1]
        try:
            ns = self.canoe.System.Namespaces.Item(namespace_chain[0])
            for child in namespace_chain[1:]:
                ns = ns.Namespaces.Item(child)
            var = ns.Variables.Item(var_name)
            value = var.Value
        except Exception:
            return None
        if value is None:
            return None
        return str(value)

    @staticmethod
    def _is_expected_status(value: str | None, expected: int) -> bool:
        if value is None:
            return False
        try:
            return int(str(value).strip()) == expected
        except ValueError:
            return False

    def _sync_measurement_ui(self) -> None:
        """
        Poll CANoe measurement state and update:
        - Record button text/style
        - Save comment button enabled/disabled
        - Status label
        """
        try:
            running = bool(self.canoe.Measurement.Running) if self.canoe else False
        except Exception:
            # CANoe died / COM broke
            running = False
            self.canoe = None
            self._update_launch_button_state()

        if running != self.last_meas_running:
            self.last_meas_running = running
            self.is_recording = running

            if running:
                # Measurement running
                self.btn_record.configure(text="Stop recording", state="normal")
                styles.style_button(self.btn_record, variant="danger")
                self.btn_discard.configure(state="normal")
                styles.style_button(self.btn_discard, variant="danger")
                self.btn_save_comment.configure(state="normal")
                styles.style_button(self.btn_save_comment, variant="primary")
                self._set_status("▶️ Measurement running", tone="success")
            else:
                # Measurement not running
                self.btn_save_comment.configure(state="disabled")
                styles.style_button(self.btn_save_comment, variant="neutral")
                self.btn_discard.configure(state="disabled")
                styles.style_button(self.btn_discard, variant="neutral")

                if self.canoe is None:
                    # Lost CANoe
                    self.btn_record.configure(text="Start recording", state="disabled")
                    styles.style_button(self.btn_record, variant="neutral")
                    self.btn_discard.configure(state="disabled")
                    styles.style_button(self.btn_discard, variant="neutral")
                    self._set_status("❌ Not connected", tone="danger")
                else:
                    # Connected but idle
                    self.btn_record.configure(text="Start recording", state="normal")
                    styles.style_button(self.btn_record, variant="success")
                    self.btn_discard.configure(state="disabled")
                    styles.style_button(self.btn_discard, variant="neutral")
                    self._set_status("⏹ Measurement stopped", tone="info")

        elapsed_display = self._format_measurement_timestamp() if running else "--:--:--.---"
        self._record_timer_var.set(f"Recording time: {elapsed_display}")
        camera_mode = self._read_sysvar_value("anSWer_SysVal::Camera_Mode")
        self._camera_mode_var.set(f"Camera mode: {camera_mode or '--'}")
        ethernet_status = self._read_sysvar_value("anSWer_SysVal::Network_Status::Ethernet")
        self._ethernet_status_var.set(f"Ethernet: {ethernet_status or '--'}")
        flexray_status = self._read_sysvar_value("anSWer_SysVal::Network_Status::Flexray")
        self._flexray_status_var.set(f"Flexray: {flexray_status or '--'}")
        ethernet_drops = self._read_sysvar_value("anSWer_SysVal::Network_Status::Ethernet_Drops")
        self._ethernet_drops_var.set(f"Ethernet drops: {ethernet_drops or '--'}")
        flexray_drops = self._read_sysvar_value("anSWer_SysVal::Network_Status::Flexray_Drops")
        self._flexray_drops_var.set(f"Flexray drops: {flexray_drops or '--'}")

        camera_ok = self._is_expected_status(camera_mode, 4)
        ethernet_ok = self._is_expected_status(ethernet_status, 1)
        flexray_ok = self._is_expected_status(flexray_status, 1)

        camera_color = styles.Palette.CHILL_GREEN_TEXT if camera_ok else styles.Palette.CHILL_RED_TEXT
        ethernet_color = styles.Palette.CHILL_GREEN_TEXT if ethernet_ok else styles.Palette.CHILL_RED_TEXT
        flexray_color = styles.Palette.CHILL_GREEN_TEXT if flexray_ok else styles.Palette.CHILL_RED_TEXT

        self.camera_mode_label.configure(text_color=camera_color)
        self.ethernet_status_label.configure(text_color=ethernet_color)
        self.flexray_status_label.configure(text_color=flexray_color)
        try:
            eth_drops_val = int(str(ethernet_drops).strip())
        except Exception:
            eth_drops_val = 0
        eth_drops_color = styles.Palette.CHILL_RED_TEXT if eth_drops_val > 0 else styles.Palette.TEXT
        self.ethernet_drops_label.configure(text_color=eth_drops_color)
        try:
            drops_val = int(str(flexray_drops).strip())
        except Exception:
            drops_val = 0
        drops_color = styles.Palette.CHILL_RED_TEXT if drops_val > 0 else styles.Palette.TEXT
        self.flexray_drops_label.configure(text_color=drops_color)

        overall_ok = camera_ok and ethernet_ok and flexray_ok
        if running and overall_ok:
            status_bg = styles.Palette.CHILL_GREEN_BG
        elif running:
            status_bg = styles.Palette.CHILL_RED_BG
        else:
            status_bg = styles.Palette.CARD_DARK
        self.status_card.configure(fg_color=(status_bg, status_bg))

        # schedule next poll
        self.after(500, self._sync_measurement_ui)

    # -------------------- File dialog --------------------
    def _choose_cfg(self) -> None:
        """File picker to choose a CANoe .cfg file."""
        path = filedialog.askopenfilename(
            title="Select CANoe .cfg",
            filetypes=[("CANoe config", "*.cfg"), ("All files", "*.*")],
        )
        if path:
            self.canoe_config.set(path)

    def _choose_log_dir(self) -> None:
        """Folder picker to choose where logs should be stored."""
        path = filedialog.askdirectory(
            title="Select log output folder",
            mustexist=True,
        )
        if path:
            self.log_dir_var.set(path)
            self._persist_state_snapshot()

    def _on_log_dir_entry_commit(self, _event: tk.Event | None = None) -> None:
        """Persist state after manual edits to the log directory."""
        self._persist_state_snapshot()

    def _resolve_log_root(self) -> Path | None:
        raw = (self.log_dir_var.get() or "").strip()
        if not raw:
            return None
        expanded = os.path.expandvars(raw)
        try:
            return Path(expanded).expanduser()
        except Exception:
            return None

    def _selected_canoe_installation(self) -> CANoeInstallation | None:
        label = self.canoe_install_var.get()
        return self._installations_by_label.get(label)

    def _selected_canoe_exec_path(self) -> Path | None:
        installation = self._selected_canoe_installation()
        return installation.exec_path if installation else None

    def _selected_canoe_exec_string(self) -> str | None:
        selected = self._selected_canoe_exec_path()
        return str(selected) if selected else None

    def _on_canoe_version_change(self, _selection: str | None = None) -> None:
        if self._selected_canoe_installation():
            self._persist_state_snapshot()
        self._update_launch_button_state()

    def _refresh_canoe_installations(self, preferred_exec: str | None = None) -> None:
        installs = discover_canoe_installations()
        self._installations_by_label = {inst.label: inst for inst in installs}
        self._exec_key_to_label = {
            _normalize_path_key(inst.exec_path): inst.label for inst in installs
        }
        self._debug_log(f"Installation refresh: discovered {len(installs)} entries.")

        values = list(self._installations_by_label.keys())
        previous_value = self.canoe_install_var.get()

        if not values:
            placeholder = "No CANoe installation found"
            self.install_dropdown.configure(values=[placeholder], state="disabled")
            self.install_dropdown.set(placeholder)
            self.btn_launch.configure(state="disabled")
            self.canoe_install_var.set(placeholder)
            self._debug_log("Installation refresh yielded no results; dropdown disabled.")
            return

        target_label = None
        if preferred_exec:
            preferred_key = _normalize_path_key(preferred_exec)
            target_label = self._exec_key_to_label.get(preferred_key)
        if target_label is None and previous_value in self._installations_by_label:
            target_label = previous_value
        if target_label is None:
            target_label = values[0]

        self.install_dropdown.configure(values=values, state="normal")
        self.install_dropdown.set(target_label)
        self.canoe_install_var.set(target_label)
        self.btn_launch.configure(state="normal")
        self._update_launch_button_state()
        self._debug_log(f"Installation refresh selected '{target_label}'.")

        if target_label != previous_value:
            self._persist_state_snapshot()
        else:
            self._update_launch_button_state()

    def _selected_canoe_is_running(self) -> bool:
        exe = self._selected_canoe_exec_path()
        if exe is None:
            return False
        return is_canoe_running(exe)

    def _active_canoe_version(self) -> str | None:
        for prog_id in ("CANoe.Application", "CANoe.Application.1"):
            instance = _get_active_canoe(prog_id)
            if instance is None:
                continue
            try:
                version = str(instance.Version)
            except Exception as exc:
                self._debug_log(f"Found CANoe via {prog_id} but failed to read Version: {exc!r}")
                continue
            self._debug_log(f"Found active CANoe via {prog_id}: {version}")
            return version
        self._debug_log("No active CANoe COM object detected while probing versions.")
        return None

    def _update_launch_button_state(self) -> None:
        if not hasattr(self, "btn_launch"):
            return

        installation = self._selected_canoe_installation()
        if installation is None:
            self.btn_launch.configure(state="disabled", text="Launch CANoe")
            styles.style_button(self.btn_launch, variant="neutral", size="lg", roundness="lg")
            return

        if self.canoe is not None:
            self.btn_launch.configure(state="disabled", text="Connected")
            styles.style_button(self.btn_launch, variant="success", size="lg", roundness="lg")
            return

        if self._selected_canoe_is_running():
            self.btn_launch.configure(state="normal", text="Connect CANoe")
            styles.style_button(self.btn_launch, variant="success", size="lg", roundness="lg")
        else:
            self.btn_launch.configure(state="normal", text="Launch CANoe")
            styles.style_button(self.btn_launch, variant="primary", size="lg", roundness="lg")

    def _process_poll_tick(self) -> None:
        if not self.winfo_exists():
            return
        self._update_launch_button_state()
        self.after(1500, self._process_poll_tick)

    # -------------------- Connect / load CANoe --------------------
    def _open_or_connect_canoe(self) -> None:
        """
        Single entry point to either connect to the running CANoe instance
        that matches the dropdown or launch & connect the selected installation.
        """
        installation = self._selected_canoe_installation()
        if installation is None:
            self._set_status("Select a CANoe installation first", tone="danger")
            styles.style_button(self.btn_launch, variant="danger", size="lg", roundness="lg")
            return

        self._debug_log(
            f"Open/connect requested -> label='{installation.label}', exec='{installation.exec_path}', prog_id='{installation.prog_id}'"
        )

        if self.canoe is not None:
            self._set_status("Already connected to CANoe.", tone="info")
            styles.style_button(self.btn_launch, variant="success", size="lg", roundness="lg")
            self._debug_log("Skipping action: already connected to CANoe COM instance.")
            return

        running_version = self._active_canoe_version()
        self._debug_log(f"Detected active CANoe version: {running_version or 'none'}")

        if self._selected_canoe_is_running():
            if running_version:
                self._set_status(f"Detected CANoe {running_version} running. Connecting...", tone="info")
                self._debug_log(f"Matching CANoe process already running (version {running_version}).")
            self._connect_selected_canoe()
            return

        if is_canoe_running():
            if running_version:
                self._set_status(
                    f"CANoe {running_version} is running but {installation.label} is selected. Launching the selected version.",
                    tone="warning",
                )
                self._debug_log(
                    f"Another CANoe instance reported version {running_version}; proceeding to launch selected installation."
                )
            else:
                self._set_status("Another CANoe instance is running. Launching the selected version.", tone="warning")
                self._debug_log("Another CANoe instance is running without detectable version; attempting launch anyway.")

        self._debug_log("Attempting to launch selected CANoe installation.")
        self._launch_selected_canoe()

        if self._selected_canoe_is_running():
            self._debug_log("Launch appears successful; trying to connect via COM.")
            self._connect_selected_canoe()
        else:
            self._debug_log("Post-launch check did not find the selected CANoe process.")

    def _launch_selected_canoe(self) -> None:
        """
        Launch the selected CANoe installation. Does not establish COM connection.
        """
        installation = self._selected_canoe_installation()
        if installation is None:
            self._set_status("Select a CANoe installation first", tone="danger")
            styles.style_button(self.btn_launch, variant="danger", size="lg", roundness="lg")
            self._debug_log("Launch aborted: no installation selected.")
            return

        executable = installation.exec_path
        self._debug_log(f"Launching CANoe from '{executable}'.")

        if self._selected_canoe_is_running():
            self._set_status("CANoe already running. Ready to connect.", tone="info")
            styles.style_button(self.btn_launch, variant="success", size="lg", roundness="lg")
            self._update_launch_button_state()
            self._debug_log("Launch skipped: selected CANoe process already running.")
            if self.canoe is None:
                self._debug_log("Launch path: attempting COM connect to running CANoe.")
                self._connect_selected_canoe()
            return

        try:
            launched = open_canoe_installation(executable)
        except Exception as e:
            self._set_status(f"Launch failed: {e}", tone="danger")
            styles.style_button(self.btn_launch, variant="danger", size="lg", roundness="lg")
            self._debug_log(f"Exception while launching CANoe: {e!r}")
            return

        if not launched:
            self._set_status("Cannot find the selected CANoe executable", tone="danger")
            styles.style_button(self.btn_launch, variant="danger", size="lg", roundness="lg")
            self._debug_log("Launch failed: executable did not start (open_canoe_installation returned False).")
            return

        _wait_for_process(executable)
        self._debug_log("Launch command sent, waiting for process detection finished.")

        if self._selected_canoe_is_running():
            self._set_status("CANoe launched. Connect when ready.", tone="success")
            styles.style_button(self.btn_launch, variant="success", size="lg", roundness="lg")
            self._debug_log("Selected CANoe process detected after launch.")
            if self.canoe is None:
                self._debug_log("Launch path: attempting COM connect after successful launch.")
                self._connect_selected_canoe()
        else:
            self._set_status("Launch command sent but process not detected.", tone="warning")
            styles.style_button(self.btn_launch, variant="danger", size="lg", roundness="lg")
            self._debug_log("Launch uncertainty: process not detected after wait.")

        self._update_launch_button_state()

    def _connect_selected_canoe(self) -> None:
        """
        Connect to a running CANoe instance that matches the selected installation.
        """
        installation = self._selected_canoe_installation()
        if installation is None:
            self._set_status("Select a CANoe installation first", tone="danger")
            styles.style_button(self.btn_launch, variant="danger", size="lg", roundness="lg")
            self._debug_log("Connect aborted: no installation selected.")
            return

        if self.canoe is not None:
            self._set_status("Already connected to CANoe.", tone="info")
            styles.style_button(self.btn_launch, variant="success", size="lg", roundness="lg")
            self._debug_log("Connect skipped: already holding CANoe COM reference.")
            return

        if not self._selected_canoe_is_running():
            self._set_status("Launch the selected CANoe before connecting.", tone="warning")
            styles.style_button(self.btn_launch, variant="danger", size="lg", roundness="lg")
            self._update_launch_button_state()
            self._debug_log("Connect aborted: selected CANoe executable not detected in process list.")
            return

        target_prog_id = installation.prog_id
        expected_major = _major_from_hint(installation.version_hint)
        self._debug_log(
            f"Connecting via COM -> target_prog_id='{target_prog_id}', expected_major={expected_major}, exec='{installation.exec_path}'"
        )

        def matches_installation(canoe_obj) -> bool:
            if expected_major is None:
                return True
            try:
                version_text = str(canoe_obj.Version)
            except Exception:
                self._debug_log("COM candidate rejected: could not read Version from instance.")
                return False
            actual_major = _extract_major_from_text(version_text)
            if actual_major != expected_major:
                self._debug_log(
                    f"COM candidate rejected: expected major {expected_major}, got {actual_major} (raw '{version_text}')."
                )
            return actual_major == expected_major

        prog_id_candidates: list[str] = []
        if expected_major is not None:
            for suffix in (
                f"{expected_major}",
                f"{expected_major}.0",
                f"{expected_major:02d}",
                f"{expected_major:02d}.0",
            ):
                candidate = f"CANoe.Application.{suffix}"
                if _prog_id_exists(candidate):
                    prog_id_candidates.append(candidate)

        for prog_id in (target_prog_id, "CANoe.Application", "CANoe.Application.1"):
            if prog_id and prog_id not in prog_id_candidates:
                prog_id_candidates.append(prog_id)

        self._debug_log(f"COM attach candidates: {prog_id_candidates}")
        self.canoe = _attach_running_canoe(prog_id_candidates, matches_installation, timeout=15.0)
        attached_from_running = self.canoe is not None

        if self.canoe is None:
            self._debug_log("COM attach failed: trying to spawn a matching CANoe COM server.")
            self.canoe = _spawn_canoe_instance(prog_id_candidates, matches_installation, timeout=20.0)
            if self.canoe is not None:
                self._debug_log("Spawned new CANoe instance via COM and obtained handle.")

        if self.canoe is None:
            self._set_status("Cannot connect to the selected CANoe version", tone="danger")
            styles.style_button(self.btn_launch, variant="danger", size="lg", roundness="lg")
            self._debug_log("COM attach failed: unable to spawn or attach to CANoe instance.")
            return
        elif attached_from_running:
            self._debug_log("Attached to already-running CANoe instance via COM.")

        cfg = self.canoe_config.get().strip()
        if cfg:
            try:
                load_canoe_config(self.canoe, cfg)
            except Exception as e:
                self._set_status(f"Connected but failed to load cfg: {e}", tone="warning")
                self._debug_log(f"Connected but failed to load cfg '{cfg}': {e!r}")
            else:
                self._set_status(f"Connected to CANoe {self.canoe.Version}", tone="success")
                self._debug_log(f"Connected and loaded cfg '{cfg}'.")
        else:
            self._set_status("Connected to CANoe", tone="success")
            self._debug_log("Connected without cfg load request.")

        styles.style_button(self.btn_launch, variant="success", size="lg", roundness="lg")
        self._debug_log("COM connection complete; ready for measurement operations.")

        self.last_meas_running = None
        self._update_launch_button_state()
        self._sync_measurement_ui()

    # -------------------- Timestamp helpers --------------------
    def _format_measurement_timestamp(self) -> str:
        """
        Timestamp relative to measurement start.
        Prefer CANoe's Measurement.GetTime(); fallback to our wall-clock delta.
        """
        seconds_float: float | None = None
        if self.canoe is not None:
            try:
                seconds_float = float(self.canoe.Measurement.GetTime())
            except Exception:
                seconds_float = None
        if seconds_float is None:
            if self._record_start_wallclock is not None:
                seconds_float = max(0.0, time.time() - self._record_start_wallclock)
            else:
                seconds_float = 0.0
        return self._format_seconds(seconds_float)

    @staticmethod
    def _format_seconds(elapsed_seconds: float) -> str:
        total_ms = int(elapsed_seconds * 1000.0)
        ms = total_ms % 1000
        total_sec = total_ms // 1000
        s = total_sec % 60
        total_min = total_sec // 60
        m = total_min % 60
        h = total_min // 60
        return f"{h:02d}:{m:02d}:{s:02d}.{ms:03d}"

    def _reset_current_session_state(self) -> None:
        self.comment_file_path = None
        self._current_log_folder = None
        self._current_prefix = None
        self._record_start_wallclock = None
        self._comment_metadata_written = False

    def _resolve_current_log_suffix(self) -> str | None:
        folder = self._current_log_folder
        prefix = self._current_prefix
        start_ts = self._record_start_wallclock
        if folder is None or prefix is None or start_ts is None:
            return None

        ignore_ext = {".txt", ".avi", ".tmp"}
        best_path = None
        best_mtime = -1.0

        try:
            for p in folder.iterdir():
                if not p.is_file():
                    continue
                if not p.name.startswith(prefix):
                    continue
                if p.suffix.lower() in ignore_ext:
                    continue
                mtime = p.stat().st_mtime
                if mtime + 1.0 < start_ts:
                    continue
                if mtime > best_mtime:
                    best_mtime = mtime
                    best_path = p
        except Exception:
            return None

        if best_path is None:
            return None

        stem = best_path.stem
        if not stem.startswith(prefix):
            return None
        suffix = stem[len(prefix):]
        return suffix or None

    def _delete_current_log_files(self) -> tuple[int, int, bool]:
        """
        Delete files belonging to the current log run.
        Returns (deleted_count, failed_count, removed_folder).
        """
        folder = self._current_log_folder
        prefix = self._current_prefix
        start_ts = self._record_start_wallclock
        if folder is None or prefix is None or start_ts is None:
            return 0, 0, False

        suffix = self._resolve_current_log_suffix()
        deleted = 0
        failed = 0

        try:
            for p in folder.iterdir():
                if not p.is_file():
                    continue
                name = p.name
                if suffix is not None:
                    match_main = name.startswith(f"{prefix}{suffix}")
                    match_video = name.startswith(f"_{prefix}{suffix}_")
                    if not (match_main or match_video):
                        continue
                else:
                    if not (name.startswith(prefix) or name.startswith(f"_{prefix}")):
                        continue
                    try:
                        mtime = p.stat().st_mtime
                    except Exception:
                        continue
                    if mtime + 1.0 < start_ts:
                        continue

                try:
                    p.unlink()
                    deleted += 1
                except Exception:
                    failed += 1
        except Exception:
            failed += 1

        removed_folder = False
        try:
            if folder.exists() and not any(folder.iterdir()):
                folder.rmdir()
                removed_folder = True
        except Exception:
            removed_folder = False

        return deleted, failed, removed_folder

    def _on_discard_click(self) -> None:
        """
        Stop the measurement and delete the files from the current log run.
        """
        if self.canoe is None or not self.is_recording:
            self._set_status("⚠️ No active recording to discard", tone="warning")
            return

        try:
            self.canoe.Measurement.Stop()
        except Exception as e:
            self._set_status(f"❌ Discard failed to stop measurement: {e}", tone="danger")
            return

        time.sleep(0.5)
        deleted, failed, removed_folder = self._delete_current_log_files()
        self._reset_current_session_state()

        if failed:
            self._set_status(f"⚠️ Discarded with {failed} delete error(s)", tone="warning")
        elif deleted:
            folder_note = " and removed empty folder" if removed_folder else ""
            self._set_status(f"🗑️ Discarded {deleted} file(s){folder_note}", tone="success")
        else:
            self._set_status("⚠️ No files found to discard", tone="warning")

    # -------------------- Comment file name resolution --------------------
    def _schedule_comment_filename_resolution(self) -> None:
        """
        Start polling the log folder to resolve the final comment_file_path.
        We need to wait for CANoe to actually create the log files with the
        real {MeasurementStart} timestamp in the filename.
        """
        self._resolve_tries = 0
        self.after(500, self._try_resolve_comment_filename_poll)

    def _try_resolve_comment_filename_poll(self) -> None:
        """
        Poll loop to resolve comment_file_path.
        If resolved: announce and stop.
        If not resolved after N tries: fallback to a wall-clock timestamp.
        """
        if self._current_log_folder is None or self._current_prefix is None:
            return

        resolved = self._try_resolve_comment_filename_once()
        if resolved:
            self._set_status(f"📝 Comments → {self.comment_file_path.name}", tone="info")
            return

        self._resolve_tries += 1
        if self._resolve_tries <= 30:  # ~15 s total (30 * 0.5s)
            self.after(500, self._try_resolve_comment_filename_poll)
        else:
            # Fallback: build a unique name from wall clock
            ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            self.comment_file_path = (
                self._current_log_folder
                / f"{self._current_prefix}{ts}.txt"
            ).resolve()
            self._set_status(
                f"⚠️ Could not resolve MeasurementStart; using {self.comment_file_path.name}",
                tone="warning",
            )

    def _try_resolve_comment_filename_once(self) -> bool:
        """
        Single attempt:
        - Look into the log folder.
        - Find files that start with current prefix (e.g. 'R300RC1_VEH123_tag_').
        - Ignore .txt/.avi/.tmp (we only care about CANoe log containers).
        - Keep only files created/modified after we pressed Start (this run),
          so we don't accidentally reuse an old session's timestamp.
        - Pick the newest one that matches.
        - Extract its {MeasurementStart} suffix and build the .txt filename.
        """
        folder = self._current_log_folder
        prefix = self._current_prefix
        start_ts = self._record_start_wallclock
        if folder is None or prefix is None or start_ts is None:
            return False

        ignore_ext = {".txt", ".avi", ".tmp"}

        best_path = None
        best_mtime = -1.0

        try:
            for p in folder.iterdir():
                if not p.is_file():
                    continue
                name = p.name
                if not name.startswith(prefix):
                    continue
                if p.suffix.lower() in ignore_ext:
                    continue

                mtime = p.stat().st_mtime

                # Only consider files that look "new enough":
                # allow +1s tolerance in case file time is slightly earlier
                if mtime + 1.0 < start_ts:
                    continue

                if mtime > best_mtime:
                    best_mtime = mtime
                    best_path = p
        except Exception:
            return False

        if best_path is None:
            return False

        # Extract suffix after the prefix, from the stem (no extension)
        suffix = best_path.stem[len(prefix):]
        if not suffix:
            return False

        new_path = (folder / f"{prefix}{suffix}.txt").resolve()
        if self.comment_file_path and self.comment_file_path != new_path:
            try:
                if self.comment_file_path.exists():
                    new_path.parent.mkdir(parents=True, exist_ok=True)
                    self.comment_file_path.replace(new_path)
            except Exception:
                # Keep the original file if renaming fails.
                return True

        self.comment_file_path = new_path
        return True

    # -------------------- Comment save --------------------
    def _on_save_comment_click(self) -> None:
        """
        Take the text in the comment box, prepend the measurement timestamp,
        and append it to the resolved .txt comment log.
        Format per line: [HH:MM:SS.mmm] comment text
        """
        if self.canoe is None or not self.is_recording:
            self._set_status("❌ Cannot save comment (not recording)", tone="danger")
            return

        comment = self.comment_box.get("1.0", "end").strip()
        if not comment:
            self._set_status("⚠️ Empty comment, not saved", tone="warning")
            return

        if self.comment_file_path is None:
            self._set_status("❌ Comment file not yet resolved", tone="warning")
            return

        try:
            ts = self._format_measurement_timestamp()
            line = f"[{ts}] {comment}\n"

            self.comment_file_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.comment_file_path, "a", encoding="utf-8") as f:
                f.write(line)

            self.comment_box.delete("1.0", "end")
            self._set_status(f"💬 Comment saved to {self.comment_file_path.name}", tone="success")
        except Exception as e:
            self._set_status(f"❌ Could not save comment: {e}", tone="danger")

    def _on_comment_enter(self, event: tk.Event) -> str | None:
        """
        Intercept Enter presses in the comment box.
        Plain Enter submits; Shift+Enter stays multi-line.
        """
        state = int(getattr(event, "state", 0))
        if state & 0x0001:
            return None
        self._on_save_comment_click()
        return "break"

    def _write_comment_metadata(self) -> None:
        if self.comment_file_path is None or self._comment_metadata_written:
            return

        vehicle_id = (self.vehicle_id.get() or "").strip()
        vehicle_model = self._vehicle_descriptor(vehicle_id)
        vehicle_number = VEHICLE_NUMBERS.get(vehicle_id.upper()) if vehicle_id else None

        def fallback(value: str | None) -> str:
            return (value or "").strip() or "--"

        lines = [
            "Recording metadata",
            f"Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"Title: {fallback(self._app_title())}",
            f"SW release: {fallback(self.sw_rel.get())}",
            f"ME version: {fallback(self.me_version_var.get())}",
            f"Recording tag: {fallback(self.log_tag.get())}",
            f"Vehicle model: {fallback(vehicle_model)}",
            f"Vehicle plate/ID: {fallback(vehicle_id)}",
            f"Vehicle number: {vehicle_number if vehicle_number is not None else '--'}",
            "",
            "Operator comments:",
        ]

        try:
            self.comment_file_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.comment_file_path, "w", encoding="utf-8") as f:
                f.write("\n".join(lines) + "\n")
            self._comment_metadata_written = True
        except Exception:
            pass

    # -------------------- Start / Stop logic --------------------
    def _on_start_stop_click(self) -> None:
        """
        If measurement is running -> Stop it.
        If measurement is not running -> configure output paths, Start it,
        and prepare comment file naming.
        UI colors/text are handled by _sync_measurement_ui().
        """
        self._debug_log("Start/Stop pressed.")
        if self.canoe is None:
            self._set_status("❌ Not connected", tone="danger")
            self._debug_log("Start/Stop aborted: no CANoe COM connection.")
            return

        try:
            currently_running = bool(self.canoe.Measurement.Running)
        except Exception as e:
            self._set_status(f"❌ Lost CANoe connection: {e}", tone="danger")
            self._debug_log(f"Start/Stop failed: cannot read Measurement.Running ({e!r}).")
            self.canoe = None
            self._update_launch_button_state()
            return

        # -------- STOP CASE --------
        if currently_running:
            try:
                self.canoe.Measurement.Stop()
                self._debug_log("Stop requested via CANoe.Measurement.Stop().")
            except Exception as e:
                self._set_status(f"❌ Stop failed: {e}", tone="danger")
                self._debug_log(f"Stop failed: {e!r}")
                return

            # Clear current session state
            self._reset_current_session_state()
            self._debug_log("Stop completed; session state cleared.")
            return

        # -------- START CASE --------

        # 1) Persist UI state to disk
        self._persist_state_snapshot()
        self._debug_log("State snapshot saved.")

        # 2) Prepare output dirs / naming
        log_root = self._resolve_log_root()
        if log_root is None:
            self._set_status("Log directory is not configured.", tone="danger")
            self._debug_log("Start aborted: log directory not configured.")
            return
        if not log_root.exists():
            self._set_status(f"Log directory does not exist: {log_root}", tone="danger")
            self._debug_log(f"Start aborted: log directory does not exist ({log_root}).")
            return
        if not log_root.is_dir():
            self._set_status(f"Log path is not a folder: {log_root}", tone="danger")
            self._debug_log(f"Start aborted: log directory is not a folder ({log_root}).")
            return

        try:
            resolved_root = log_root.resolve()
        except Exception:
            resolved_root = log_root
        release_dir = resolved_root / f"{self.sw_rel.get()}"
        release_dir.mkdir(parents=True, exist_ok=True)

        date = datetime.now().strftime("%Y-%m-%d")
        # sanitize components for filenames
        safe_tag = (self.log_tag.get() or "").replace(" ", "_")
        rel = (self.sw_rel.get() or "").replace(" ", "_")
        veh = self._vehicle_prefix_component()

        # base string CANoe will expand. CANoe will replace {MeasurementStart}
        # Include only the vehicle number tag in the logging "title"/prefix.
        # Example: R300RC1_Veh1_myCase_{MeasurementStart}
        base_prefix_parts = [rel]
        if veh:
            base_prefix_parts.append(veh)
        if safe_tag:
            base_prefix_parts.append(safe_tag)
        base_prefix = "_".join([p for p in base_prefix_parts if p])

        log_name = f"{base_prefix}_{{MeasurementStart}}"
        log_folder_name = base_prefix  # keep folder grouped by SW+Vehicle(+Tag)

        log_folder = release_dir / f"{rel}_{date}" / log_folder_name
        log_folder.mkdir(parents=True, exist_ok=True)
        self._debug_log(f"Log folder resolved to {log_folder}")

        # Save info so we can later resolve the comment file name
        self._current_log_folder = log_folder
        self._current_prefix = f"{base_prefix}_"
        self._record_start_wallclock = time.time()
        self._comment_metadata_written = False

        # Create the comment file immediately with a wall-clock suffix.
        ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        self.comment_file_path = (log_folder / f"{base_prefix}_{ts}.txt").resolve()
        self._write_comment_metadata()
        self._debug_log(f"Comment file initialized at {self.comment_file_path}")

        try:
            # 3) Configure CANoe logging blocks to point at <log_folder>/<log_name>.ext
            logging_collection = self.canoe.Configuration.OnlineSetup.LoggingCollection
            try:
                for i in range(logging_collection.Count):
                    log_block = logging_collection.Item(i + 1)
                    original_name_split = log_block.FullName.split(".")
                    file_extension = original_name_split[-1]
                    log_block.FullName = str((log_folder / f"{log_name}.{file_extension}").resolve())
                self._debug_log(f"Logging blocks updated: {logging_collection.Count}")
            except Exception as e:
                self._set_status(f"⚠️ Could not set logging blocks: {e}", tone="warning")
                self._debug_log(f"Logging block update failed: {e!r}")

            # 4) Configure CANoe video captures to same folder
            video_config = self.canoe.Configuration.OnlineSetup.VideoWindows
            for i in range(video_config.Count):
                vw = video_config.Item(i + 1)
                video_name = vw.Name
                vw.RecordFile = str((log_folder / f"_{log_name}_{video_name}.avi").resolve())
            self._debug_log(f"Video windows updated: {video_config.Count}")

            # 5) Start CANoe measurement
            self.canoe.Measurement.Start()
            time.sleep(0.5)
            self._debug_log("CANoe.Measurement.Start() invoked.")

        except Exception as e:
            self._set_status(f"❌ Error on logging setup/start: {e}", tone="danger")
            self._debug_log(f"Start failed: {e!r}")
            return

        # 6) After CANoe starts, resolve the actual filename suffix CANoe used
        self._debug_log("Scheduling comment filename resolution.")
        self._schedule_comment_filename_resolution()

    # -------------------- Debug helper --------------------
    def _check_logging(self) -> None:
        """
        Helper to display whether logging blocks exist in the loaded configuration.
        """
        if self.canoe is None:
            self._set_status("❌ Not connected", tone="danger")
            return
        response = get_logging_block_status(self.canoe)
        if response.startswith("✅"):
            tone = "success"
        elif response.startswith("⚠️"):
            tone = "warning"
        elif response.startswith("❌"):
            tone = "danger"
        else:
            tone = "info"
        self._set_status(f"{response}", tone=tone)
