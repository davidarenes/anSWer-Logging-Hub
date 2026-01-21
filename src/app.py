"""
ui/app.py - GUI bootstrap for the CANoe Logging Tool.
"""

from __future__ import annotations

import styles
from core.state import discover_paths, load_state
from ui.main_window import MainWindow


def run() -> int:
    styles.apply_global(appearance="dark", color_theme="dark-blue")

    paths = discover_paths()
    state = load_state(paths)

    app = MainWindow(paths=paths, state=state)
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(run())
