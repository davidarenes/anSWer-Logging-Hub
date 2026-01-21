from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import os
import re
import subprocess
import time
import psutil
import win32com.client
import win32api
import pythoncom
import winreg


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


def wait_for_process(executable: Path, timeout: float = 20.0) -> None:
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
