from __future__ import annotations

import sys
from pathlib import Path

from .ui.main_window import run_app


def find_project_root() -> Path:
    executable_dir = Path(sys.executable).resolve().parent
    if (executable_dir / "templates").exists():
        return executable_dir

    current = Path(__file__).resolve()
    for parent in current.parents:
        if (parent / "templates").exists():
            return parent
    return Path.cwd()


def main() -> int:
    return run_app(find_project_root())


if __name__ == "__main__":
    sys.exit(main())
