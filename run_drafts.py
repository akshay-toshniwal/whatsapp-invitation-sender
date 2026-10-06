#!/usr/bin/env python3
"""
Cross-platform entrypoint for WhatsApp draft creation.

Works on Windows, macOS, and Linux:

  python setup_env.py
  python run_drafts.py --file guests.xlsx --from-sr 1 --to-sr 300
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

from platform_paths import ROOT, venv_python

SCRIPT = ROOT / "playwright" / "create_drafts.py"


def main() -> int:
    if not SCRIPT.exists():
        print(f"Missing script: {SCRIPT}", file=sys.stderr)
        return 1

    py = venv_python()
    if not py.exists():
        print("Virtual environment not found. Run this first:\n  python setup_env.py")
        return 1

    return subprocess.call([str(py), str(SCRIPT), *sys.argv[1:]], cwd=str(ROOT))


if __name__ == "__main__":
    raise SystemExit(main())
