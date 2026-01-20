"""
ui/app.py — GUI bootstrap for the CANoe Logging Tool.
Fully anglicized identifiers and English docstrings/comments.
"""

from __future__ import annotations

from dataclasses import dataclass, asdict, is_dataclass
from datetime import datetime
from pathlib import Path
from typing import NamedTuple
import json
import os
import re
import time
import tkinter as tk
from tkinter import filedialog
import subprocess
import sys
import psutil
import win32com.client
import win32api
import pythoncom
import winreg
import customtkinter as ctk

import styles

SW_MAJOR_RELEASES = [
    "R120",
    "R200",
    "R300",
    "R310",
    "R320",
    "R400",
    "R410",
    "R420",
    "R500",
    "R510",
] 

SW_MINOR_RELEASES = [
    "RX0",
    "RX1",
    "RX2",
    "RX3",
    "RX4",
    "RX5",
    "RX6",
    "RX7",
    "RX8",
    "RX9",
    "RX10",
    "RC0",
    "RC1",
    "RC2",
    "RC3",
    "RC4",
    "RC5",
    "RC6",
    "RC7",
    "RC8",
    "RC9",
    "RC10",
]

VEHICLE_NUMBERS: dict[str, int] = {
    "YJA55E": 1,
    "SOF03C": 3,
    "RWA50U": 4,
    "MUD01W": 5,
    "JUD79J": 6,
}

# ------------------------------------ #
# ----------- STATE STORE  ----------- #
# ------------------------------------ #

@dataclass
class AppState:
    """
    Minimal app state you want to persist between sessions.
    """
    cfg_file: str = ""   # Path to CANoe .cfg to open
    tag: str = ""        # Recording tag
    sw_rel: str = ""     # e.g. "R300RC1"
    vehicle_id: str = "" # e.g. VIN or fleet code
    canoe_exec: str = "" # path to preferred CANoe executable
    log_dir: str = ""    # base directory for log output

    @staticmethod
    def load(paths: "AppPaths") -> "AppState":
        """
        Load state from disk. Returns default state if the file does not exist
        or cannot be parsed.
        """
        p = paths.state_file
        if not p.exists():
            return AppState()
        try:
            raw = json.loads(p.read_text(encoding="utf-8"))
            if not isinstance(raw, dict):
                return AppState()
            known = {k: v for k, v in raw.items() if k in AppState().__dict__.keys()}
            return AppState(**known)
        except Exception:
            return AppState()

    def save(self, state: "AppState", paths: "AppPaths") -> None:
        """
        Persist current state to disk as JSON.
        """
        p = paths.state_file
        data = asdict(state) if is_dataclass(state) else state.__dict__
        p.parent.mkdir(parents=True, exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)


def load_state(paths: "AppPaths") -> AppState:
    return AppState.load(paths)


def save_state(state: AppState, paths: "AppPaths") -> None:
    state.save(state, paths)


def load_vehicle_catalog(app_root: Path) -> dict[str, str]:
    """
    Load the list of known vehicles from data/vehicles.json.
    Returns a mapping of vehicle ID -> descriptor (e.g. model).
    """
    vehicles_path = app_root / "data" / "vehicles.json"
    if not vehicles_path.exists():
        return {}
    try:
        raw = json.loads(vehicles_path.read_text(encoding="utf-8"))
    except Exception:
        return {}
    if not isinstance(raw, dict):
        return {}

    catalog: dict[str, str] = {}
    for vehicle_id, descriptor in raw.items():
        key = (vehicle_id or "").strip()
        value = (descriptor or "").strip()
        if key:
            catalog[key] = value
    return catalog


# ------------------------------------ #
# -------------- PATHS --------------- #
# ------------------------------------ #

class AppPaths(NamedTuple):
    root: Path
    cfg_file: Path | None
    log_dir: Path
    data_dir: Path
    state_file: Path


def app_root_dir() -> Path:
    """Return the folder where the app (script or EXE) lives."""
    if getattr(sys, "frozen", False):  # PyInstaller bundle
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def user_data_dir(app_name: str = "anSWer-Logging-Hub") -> Path:
    """
    A writable per-user data directory.
    On Windows we try APPDATA / LOCALAPPDATA.
    Fallback is local dir.
    """
    if os.name == "nt":
        base = os.environ.get("APPDATA") or os.environ.get("LOCALAPPDATA") or app_root_dir()
        return Path(base) / app_name
    return app_root_dir() / "data"


def discover_paths() -> AppPaths:
    """
    Build all relevant paths (logs folder, persisted state file, etc.).
    Try to restore last cfg_file if possible.
    """
    project_root = app_root_dir()

    log_dir = (project_root / "../Logs").resolve()
    log_dir.mkdir(parents=True, exist_ok=True)

    data_dir = user_data_dir()
    data_dir.mkdir(parents=True, exist_ok=True)

    state_file = data_dir / "state.json"

    cfg_path = None
    if state_file.exists():
        try:
            state_data = json.loads(state_file.read_text(encoding="utf-8"))
            maybe_cfg = state_data.get("cfg_file", "")
            if maybe_cfg and Path(maybe_cfg).exists():
                cfg_path = Path(maybe_cfg)
        except Exception:
            pass

    if cfg_path is None:
        default_cfg = (project_root / "SPA1_anSWer_SysVal" / "SPA1_anSWer_SysVal.cfg").resolve()
        cfg_path = default_cfg if default_cfg.exists() else None

    return AppPaths(
        root=project_root,
        cfg_file=cfg_path,
        log_dir=log_dir,
        data_dir=data_dir,
        state_file=state_file,
    )


# ------------------------------------ #
# -------------- CANOE --------------- #
# ------------------------------------ #

@dataclass(frozen=True)
class CANoeInstallation:
    label: str
    exec_path: Path
    version_hint: tuple[int, ...]
    prog_id: str | None = None


def _normalize_path_key(value: str | Path) -> str:
    return str(Path(value).resolve(strict=False)).lower()


def _extract_version_hint(text: str) -> tuple[int, ...]:
    match = re.search(r"(\d+(?:\.\d+)*)", text)
    if not match:
        return tuple()
    return tuple(int(part) for part in match.group(1).split("."))


def _major_from_hint(hint: tuple[int, ...]) -> int | None:
    return hint[0] if hint else None


def _extract_major_from_text(text: str) -> int | None:
    hint = _extract_version_hint(text)
    return _major_from_hint(hint)


def _file_version_hint(executable: Path) -> tuple[int, ...]:
    try:
        info = win32api.GetFileVersionInfo(str(executable), "\\")
    except Exception:
        return tuple()
    ms = info.get("FileVersionMS", 0)
    ls = info.get("FileVersionLS", 0)
    if not (ms or ls):
        return tuple()
    return (ms >> 16, ms & 0xFFFF, ls >> 16, ls & 0xFFFF)


def _extract_executable_from_command(command: str) -> Path | None:
    command = (command or "").strip()
    if not command:
        return None
    if command.startswith('"'):
        end = command.find('"', 1)
        if end == -1:
            return None
        raw_path = command[1:end]
    else:
        parts = command.split()
        if not parts:
            return None
        raw_path = parts[0]
    expanded = os.path.expandvars(raw_path)
    try:
        return Path(expanded)
    except Exception:
        return None


_PROG_ID_EXEC_CACHE: dict[str, Path | None] = {}


def _prog_id_executable(prog_id: str) -> Path | None:
    prog_id = prog_id or ""
    if not prog_id:
        return None
    if prog_id in _PROG_ID_EXEC_CACHE:
        return _PROG_ID_EXEC_CACHE[prog_id]
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{prog_id}\\CLSID") as clsid_key:
            clsid, _ = winreg.QueryValueEx(clsid_key, None)
    except OSError:
        _PROG_ID_EXEC_CACHE[prog_id] = None
        return None
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"CLSID\\{clsid}\\LocalServer32") as server_key:
            command, _ = winreg.QueryValueEx(server_key, None)
    except OSError:
        _PROG_ID_EXEC_CACHE[prog_id] = None
        return None
    exe = _extract_executable_from_command(command)
    _PROG_ID_EXEC_CACHE[prog_id] = exe
    return exe


def _prog_id_targets_exec(prog_id: str, executable: Path) -> bool:
    target = _prog_id_executable(prog_id)
    if target is None:
        return False
    return _normalize_path_key(target) == _normalize_path_key(executable)


def _resolve_prog_id_for_installation(executable: Path, major_hint: int | None) -> str | None:
    candidates: list[str] = []
    if major_hint is not None:
        suffixes = {
            f"{major_hint}",
            f"{major_hint}.0",
            f"{major_hint:02d}",
            f"{major_hint:02d}.0",
        }
        for suffix in suffixes:
            candidates.append(f"CANoe.Application.{suffix}")
    candidates.extend(
        [
            "CANoe.Application",
            "CANoe.Application.1",
        ]
    )

    seen: set[str] = set()
    for prog_id in candidates:
        if not prog_id or prog_id in seen:
            continue
        seen.add(prog_id)
        if _prog_id_targets_exec(prog_id, executable):
            return prog_id
    return None


def _prog_id_exists(prog_id: str) -> bool:
    clsid_from_progid = getattr(pythoncom, "CLSIDFromProgID", None)
    if clsid_from_progid is None:
        return True
    try:
        clsid_from_progid(prog_id)
    except pythoncom.com_error:
        return False
    except Exception:
        return False
    return True


def _format_canoe_label(folder_name: str, bits: str) -> str:
    base = folder_name.replace("_", " ").strip() or "CANoe"
    return f"{base} ({bits})" if bits else base


def _iter_dirs(path: Path) -> list[Path]:
    try:
        return [p for p in path.iterdir() if p.is_dir()]
    except Exception:
        return []


def _candidate_directories(root: Path, max_depth: int = 2) -> list[Path]:
    queue: list[tuple[Path, int]] = [(root, 0)]
    seen: set[Path] = set()
    candidates: list[Path] = []

    while queue:
        current, depth = queue.pop(0)
        for child in _iter_dirs(current):
            if child in seen:
                continue
            seen.add(child)
            lowered = child.name.lower()
            if "canoe" in lowered:
                candidates.append(child)
            if depth < max_depth and ("vector" in lowered or "canoe" in lowered):
                queue.append((child, depth + 1))

    return candidates


def _installation_from_dir(directory: Path) -> CANoeInstallation | None:
    candidates = [
        (directory / "Exec64" / "CANoe64.exe", "64-bit"),
        (directory / "Exec32" / "CANoe32.exe", "32-bit"),
        (directory / "CANoe64.exe", "64-bit"),
        (directory / "CANoe32.exe", "32-bit"),
    ]
    for candidate, bits in candidates:
        if not candidate.exists():
            continue

        label = _format_canoe_label(directory.name, bits)
        version_hint = _extract_version_hint(directory.name)
        if not version_hint:
            version_hint = _extract_version_hint(candidate.stem)
        if not version_hint:
            version_hint = _file_version_hint(candidate)

        major = _major_from_hint(version_hint)
        prog_id = _resolve_prog_id_for_installation(candidate, major)

        return CANoeInstallation(
            label=label,
            exec_path=candidate,
            version_hint=version_hint,
            prog_id=prog_id,
        )
    return None


def _candidate_roots() -> list[Path]:
    raw_roots = [
        os.environ.get("VECTOR_CANOE_HOME"),
        os.environ.get("VECTOR_CANOE_ROOT"),
        os.environ.get("ProgramFiles"),
        os.environ.get("ProgramW6432"),
        os.environ.get("ProgramFiles(x86)"),
        r"C:\Program Files\Vector",
        r"C:\Program Files",
        r"C:\Program Files (x86)",
    ]
    roots: list[Path] = []
    for raw in raw_roots:
        if not raw:
            continue
        path = Path(raw)
        if path.exists() and path not in roots:
            roots.append(path)
    return roots


def discover_canoe_installations() -> list[CANoeInstallation]:
    installs: list[CANoeInstallation] = []
    seen_execs: set[str] = set()

    for root in _candidate_roots():
        if "canoe" in root.name.lower():
            inst = _installation_from_dir(root)
            if inst:
                key = _normalize_path_key(inst.exec_path)
                if key not in seen_execs:
                    installs.append(inst)
                    seen_execs.add(key)

        for directory in _candidate_directories(root):
            inst = _installation_from_dir(directory)
            if not inst:
                continue
            key = _normalize_path_key(inst.exec_path)
            if key in seen_execs:
                continue
            installs.append(inst)
            seen_execs.add(key)

    installs.sort(key=lambda inst: inst.label.lower())
    installs.sort(key=lambda inst: inst.version_hint, reverse=True)
    return installs

def connect_canoe(prog_id: str | None = None, *, new_instance: bool = False):
    """
    Get a COM handle to a CANoe instance.
    If prog_id is provided (e.g., CANoe.Application.15) we only attempt that server.
    """
    target = prog_id or "CANoe.Application"
    factory = win32com.client.DispatchEx if new_instance else win32com.client.Dispatch
    return factory(target)


def _get_active_canoe(prog_id: str) -> object | None:
    try:
        return win32com.client.GetActiveObject(prog_id)
    except Exception:
        return None


def _wait_for_process(executable: Path, timeout: float = 20.0) -> None:
    deadline = time.time() + timeout
    while time.time() < deadline:
        if is_canoe_running(executable):
            return
        time.sleep(0.5)


def _attach_running_canoe(prog_ids: list[str], matcher, timeout: float = 15.0):
    deadline = time.time() + timeout
    while time.time() < deadline:
        for prog_id in prog_ids:
            candidate = _get_active_canoe(prog_id)
            if candidate and matcher(candidate):
                return candidate
        time.sleep(0.5)
    return None


def _spawn_canoe_instance(prog_ids: list[str], matcher, timeout: float = 10.0):
    if not prog_ids:
        return None
    deadline = time.time() + timeout
    while time.time() < deadline:
        for prog_id in prog_ids:
            try:
                candidate = connect_canoe(prog_id=prog_id, new_instance=True)
            except Exception:
                continue
            if candidate and matcher(candidate):
                return candidate
        time.sleep(0.5)
    return None


def load_canoe_config(canoe, cfg_file: str | Path) -> None:
    """
    Load a .cfg file into the running CANoe instance
    if it's not already open.
    """
    current = getattr(canoe.Configuration, "FullName", "")
    if current and Path(current).resolve() != Path(cfg_file).resolve():
        canoe.Open(str(cfg_file))


def open_canoe_installation(executable: Path) -> bool:
    """
    Ensure the selected CANoe executable is running; launch it if it's not.
    Return True if the process is running or successfully launched.
    """
    exe = Path(executable)
    if not exe.exists():
        return False

    if not is_canoe_running(exe):
        subprocess.Popen([str(exe)])
    return True


def get_logging_block_status(canoe) -> str:
    """Check whether there are logging blocks configured."""
    logging_config = canoe.Configuration.OnlineSetup.LoggingCollection
    if logging_config.Count == 0:
        return "❌ No logging blocks found"
    return "✅ Logging block(s) found"


def is_canoe_running(executable: str | Path | None = None) -> bool:
    """Check process list to see if CANoe is running (optionally matching an executable)."""
    exec_key = _normalize_path_key(executable) if executable else None
    for proc in psutil.process_iter(["name", "exe"]):
        try:
            name = proc.info.get("name") or ""
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
        if name and "canoe" in name.lower():
            if exec_key is None:
                return True
            exe_path = proc.info.get("exe")
            if exe_path:
                try:
                    if _normalize_path_key(exe_path) == exec_key:
                        return True
                except Exception:
                    continue
    return False


# ------------------------------------ #
# ----------- MAIN WINDOW ------------ #
# ------------------------------------ #

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
        self._log_dir_hint_var = tk.StringVar()
        self._update_log_dir_hint()
        self.log_dir_var.trace_add("write", lambda *_: self._on_log_dir_var_changed())
        self._record_timer_var = tk.StringVar(value="Recording time: --:--:--.---")
        self._camera_mode_var = tk.StringVar(value="Camera mode: --")
        self._ethernet_status_var = tk.StringVar(value="Ethernet: --")
        self._flexray_status_var = tk.StringVar(value="Flexray: --")
        self._hint_popup: ctk.CTkToplevel | None = None
        self._theme_mode: str = "neutral"
        self._theme_cards: list[ctk.CTkFrame] = []
        self._theme_subcards: list[ctk.CTkFrame] = []
        self._theme_entries: list[ctk.CTkEntry] = []
        self._theme_menus: list[ctk.CTkOptionMenu] = []
        self._theme_textboxes: list[ctk.CTkTextbox] = []

        # Build UI
        self._build_body()
        self._update_titles_with_release()
        self._refresh_canoe_installations(preferred_exec=self._initial_canoe_exec or None)

        # Focus window
        self.after(0, self.focus_set)
        self.update_idletasks()
        self.lift()

        # Start periodic polling of CANoe measurement state
        self.after(500, self._sync_measurement_ui)
        self.after(1500, self._process_poll_tick)

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

        self.body = ctk.CTkScrollableFrame(self, fg_color=styles.Palette.BG)
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

        # ---- Measurement controls ----
        action_card = styles.card(self.body)
        action_card.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, pad_y))
        action_card.grid_columnconfigure(0, weight=1)

        action_title = ctk.CTkLabel(action_card, text="Measurement control")
        styles.style_label(action_title, kind="section")
        action_title.grid(row=0, column=0, sticky="w", padx=pad_x, pady=(pad_y, 4))

        action_hint = ctk.CTkLabel(
            action_card,
            text="Start/stop CANoe logging once metadata is ready.",
        )
        styles.style_label(action_hint, kind="hint")
        action_hint.grid(row=1, column=0, sticky="w", padx=pad_x, pady=(0, 4))

        action_btn_row = ctk.CTkFrame(action_card, fg_color="transparent")
        action_btn_row.grid(row=2, column=0, sticky="ew", padx=pad_x, pady=(pad_y, pad_y // 2))
        action_btn_row.grid_columnconfigure(0, weight=7)
        action_btn_row.grid_columnconfigure(1, weight=3)

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

        log_dir_row = ctk.CTkFrame(action_card, fg_color="transparent")
        log_dir_row.grid(row=4, column=0, sticky="ew", padx=pad_x, pady=(pad_y // 2, pad_y // 2))
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

        log_hint = ctk.CTkLabel(
            action_card,
            textvariable=self._log_dir_hint_var,
        )
        styles.style_label(log_hint, kind="hint")
        log_hint.grid(row=5, column=0, sticky="w", padx=pad_x, pady=(0, pad_y))

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

        self.record_timer_label = ctk.CTkLabel(
            status_row,
            textvariable=self._record_timer_var,
            anchor="center",
        )
        styles.style_label(self.record_timer_label, kind="body")
        self.record_timer_label.grid(row=0, column=0, sticky="nsew")

        self.camera_mode_box = ctk.CTkFrame(
            status_row,
            corner_radius=styles.Metrics.RADIUS_MD,
            border_width=1,
            border_color=styles.Palette.CARD_BORDER,
            fg_color=(styles.Palette.CARD_DARK, styles.Palette.CARD_DARK_ALT),
        )
        self.camera_mode_box.grid(row=0, column=1, sticky="nsew", padx=(4, 4), pady=4)
        self.camera_mode_box.grid_columnconfigure(0, weight=1)
        self.camera_mode_box.grid_rowconfigure(0, weight=1)
        self.camera_mode_label = ctk.CTkLabel(
            self.camera_mode_box,
            textvariable=self._camera_mode_var,
            anchor="center",
        )
        styles.style_label(self.camera_mode_label, kind="body")
        self.camera_mode_label.configure(font=styles.Fonts.BODY_BOLD)
        self.camera_mode_label.grid(row=0, column=0, sticky="nsew")

        self.ethernet_status_box = ctk.CTkFrame(
            status_row,
            corner_radius=styles.Metrics.RADIUS_MD,
            border_width=1,
            border_color=styles.Palette.CARD_BORDER,
            fg_color=(styles.Palette.CARD_DARK, styles.Palette.CARD_DARK_ALT),
        )
        self.ethernet_status_box.grid(row=0, column=2, sticky="nsew", padx=(4, 4), pady=4)
        self.ethernet_status_box.grid_columnconfigure(0, weight=1)
        self.ethernet_status_box.grid_rowconfigure(0, weight=1)
        self.ethernet_status_label = ctk.CTkLabel(
            self.ethernet_status_box,
            textvariable=self._ethernet_status_var,
            anchor="center",
        )
        styles.style_label(self.ethernet_status_label, kind="body")
        self.ethernet_status_label.configure(font=styles.Fonts.BODY_BOLD)
        self.ethernet_status_label.grid(row=0, column=0, sticky="nsew")

        self.flexray_status_box = ctk.CTkFrame(
            status_row,
            corner_radius=styles.Metrics.RADIUS_MD,
            border_width=1,
            border_color=styles.Palette.CARD_BORDER,
            fg_color=(styles.Palette.CARD_DARK, styles.Palette.CARD_DARK_ALT),
        )
        self.flexray_status_box.grid(row=0, column=3, sticky="nsew", padx=(4, 4), pady=4)
        self.flexray_status_box.grid_columnconfigure(0, weight=1)
        self.flexray_status_box.grid_rowconfigure(0, weight=1)
        self.flexray_status_label = ctk.CTkLabel(
            self.flexray_status_box,
            textvariable=self._flexray_status_var,
            anchor="center",
        )
        styles.style_label(self.flexray_status_label, kind="body")
        self.flexray_status_label.configure(font=styles.Fonts.BODY_BOLD)
        self.flexray_status_label.grid(row=0, column=0, sticky="nsew")

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
        self.debug_toggle_btn = ctk.CTkButton(
            self.body,
            text="▶ Debug log",
            command=self._toggle_debug_panel,
            width=170,
        )
        styles.style_button(self.debug_toggle_btn, variant="neutral", size="sm", roundness="md")
        self.debug_toggle_btn.grid(row=5, column=0, columnspan=2, sticky="ew", padx=pad_x, pady=(pad_y // 2, pad_y // 2))

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
        self._theme_subcards = [
            self.camera_mode_box,
            self.ethernet_status_box,
            self.flexray_status_box,
        ]
        self._apply_overall_theme("neutral")

    def _create_hint_icon(self, master, text: str) -> ctk.CTkButton:
        """
        Compact info icon that shows contextual help on hover.
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
        btn.bind("<Enter>", lambda _e, msg=text, widget=btn: self._show_hint_tooltip(msg, widget))
        btn.bind("<Leave>", lambda _e: self._hide_hint_tooltip())
        btn.bind("<ButtonPress>", lambda _e: self._hide_hint_tooltip())
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

        if mode == "ok":
            bg = styles.Palette.OK_BG
            card = styles.Palette.OK_CARD
            card_alt = styles.Palette.OK_CARD_ALT
            border = styles.Palette.OK_BORDER
            inner = styles.Palette.OK_INNER
            input_bg = styles.Palette.OK_INNER
            input_border = styles.Palette.OK_BORDER
        elif mode == "nok":
            bg = styles.Palette.NOK_BG
            card = styles.Palette.NOK_CARD
            card_alt = styles.Palette.NOK_CARD_ALT
            border = styles.Palette.NOK_BORDER
            inner = styles.Palette.NOK_INNER
            input_bg = styles.Palette.NOK_INNER
            input_border = styles.Palette.NOK_BORDER
        else:
            bg = styles.Palette.BG
            card = styles.Palette.CARD_DARK
            card_alt = styles.Palette.CARD_DARK_ALT
            border = styles.Palette.CARD_BORDER
            inner = styles.Palette.CARD_DARK_ALT
            input_bg = styles.Palette.INPUT_BG
            input_border = styles.Palette.INPUT_BORDER

        self.configure(bg=bg)
        if getattr(self, "body", None) is not None:
            self.body.configure(fg_color=bg)

        for card_frame in self._theme_cards:
            card_frame.configure(fg_color=(card, card_alt), border_color=border)

        for subcard in self._theme_subcards:
            subcard.configure(fg_color=(inner, inner), border_color=border)

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

        camera_ok = self._is_expected_status(camera_mode, 4)
        ethernet_ok = self._is_expected_status(ethernet_status, 1)
        flexray_ok = self._is_expected_status(flexray_status, 1)

        camera_color = styles.Palette.CHILL_GREEN_TEXT if camera_ok else styles.Palette.CHILL_RED_TEXT
        ethernet_color = styles.Palette.CHILL_GREEN_TEXT if ethernet_ok else styles.Palette.CHILL_RED_TEXT
        flexray_color = styles.Palette.CHILL_GREEN_TEXT if flexray_ok else styles.Palette.CHILL_RED_TEXT

        self.camera_mode_label.configure(text_color=camera_color)
        self.ethernet_status_label.configure(text_color=ethernet_color)
        self.flexray_status_label.configure(text_color=flexray_color)

        overall_ok = camera_ok and ethernet_ok and flexray_ok
        if running:
            theme = "ok" if overall_ok else "nok"
        else:
            theme = "neutral"
        self._apply_overall_theme(theme)

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

    def _on_log_dir_var_changed(self, *_args: object) -> None:
        self._update_log_dir_hint()

    def _update_log_dir_hint(self) -> None:
        path = (self.log_dir_var.get() or "").strip() or "Not set"
        self._log_dir_hint_var.set(f"Log output directory → {path}")

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

        self.comment_file_path = (folder / f"{prefix}{suffix}.txt").resolve()
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

    # -------------------- Start / Stop logic --------------------
    def _on_start_stop_click(self) -> None:
        """
        If measurement is running -> Stop it.
        If measurement is not running -> configure output paths, Start it,
        and prepare comment file naming.
        UI colors/text are handled by _sync_measurement_ui().
        """
        if self.canoe is None:
            self._set_status("❌ Not connected", tone="danger")
            return

        try:
            currently_running = bool(self.canoe.Measurement.Running)
        except Exception as e:
            self._set_status(f"❌ Lost CANoe connection: {e}", tone="danger")
            self.canoe = None
            self._update_launch_button_state()
            return

        # -------- STOP CASE --------
        if currently_running:
            try:
                self.canoe.Measurement.Stop()
            except Exception as e:
                self._set_status(f"❌ Stop failed: {e}", tone="danger")
                return

            # Clear current session state
            self._reset_current_session_state()
            return

        # -------- START CASE --------

        # 1) Persist UI state to disk
        self._persist_state_snapshot()

        # 2) Prepare output dirs / naming
        log_root = self._resolve_log_root()
        if log_root is None:
            self._set_status("Log directory is not configured.", tone="danger")
            return
        if not log_root.exists():
            self._set_status(f"Log directory does not exist: {log_root}", tone="danger")
            return
        if not log_root.is_dir():
            self._set_status(f"Log path is not a folder: {log_root}", tone="danger")
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

        # Save info so we can later resolve the comment file name
        self._current_log_folder = log_folder
        self._current_prefix = f"{base_prefix}_"
        self.comment_file_path = None
        self._record_start_wallclock = time.time()
        self._comment_metadata_written = False

        try:
            # 3) Configure CANoe logging blocks to point at <log_folder>/<log_name>.ext
            logging_collection = self.canoe.Configuration.OnlineSetup.LoggingCollection
            try:
                for i in range(logging_collection.Count):
                    log_block = logging_collection.Item(i + 1)
                    original_name_split = log_block.FullName.split(".")
                    file_extension = original_name_split[-1]
                    log_block.FullName = str((log_folder / f"{log_name}.{file_extension}").resolve())
            except Exception as e:
                self._set_status(f"⚠️ Could not set logging blocks: {e}", tone="warning")

            # 4) Configure CANoe video captures to same folder
            video_config = self.canoe.Configuration.OnlineSetup.VideoWindows
            for i in range(video_config.Count):
                vw = video_config.Item(i + 1)
                video_name = vw.Name
                vw.RecordFile = str((log_folder / f"_{log_name}_{video_name}.avi").resolve())

            # 5) Start CANoe measurement
            self.canoe.Measurement.Start()
            time.sleep(0.5)

        except Exception as e:
            self._set_status(f"❌ Error on logging setup/start: {e}", tone="danger")
            return

        # 6) After CANoe starts, resolve the actual filename suffix CANoe used
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


# ------------------------------------ #
# --------------- APP ---------------- #
# ------------------------------------ #

def run() -> int:
    styles.apply_global(appearance="dark", color_theme="dark-blue")

    paths = discover_paths()
    state = load_state(paths)

    app = MainWindow(paths=paths, state=state)
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(run())
