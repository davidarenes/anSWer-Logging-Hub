from __future__ import annotations

from dataclasses import dataclass, asdict, is_dataclass
from pathlib import Path
from typing import NamedTuple
import json
import os
import sys

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

SW_RELEASE_TYPES = [
    "RC",
    "RX",
]

SW_RELEASE_MINORS = [
    "0",
    "1",
    "2",
    "3",
    "4",
    "5",
    "6",
    "7",
    "8",
    "9",
    "10",
]

ME_VERSIONS = [
    "2.0",
    "2.1",
    "4.2",
]

VEHICLE_NUMBERS: dict[str, int] = {
    "YJA55E": 1,
    "SOF03C": 3,
    "RWA50U": 4,
    "MUD01W": 5,
    "JUD79J": 6,
}


@dataclass
class AppState:
    """
    Minimal app state you want to persist between sessions.
    """
    cfg_file: str = ""   # Path to CANoe .cfg to open
    tag: str = ""        # Recording tag
    sw_rel: str = ""     # e.g. "R300RC1"
    me_version: str = "" # e.g. "2.0"
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
    return Path(__file__).resolve().parent.parent


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
