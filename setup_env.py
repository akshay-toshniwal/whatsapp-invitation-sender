#!/usr/bin/env python3
"""
One-time setup for Windows, macOS, and Linux.

  python setup_env.py

Creates .venv, installs dependencies, and downloads Chromium for Playwright.
"""

from __future__ import annotations

import sys

from platform_paths import REQUIREMENTS, ROOT, ensure_venv_python, run


def main() -> int:
    if not REQUIREMENTS.exists():
        print(f"Missing {REQUIREMENTS}", file=sys.stderr)
        return 1

    py = ensure_venv_python()
    run([str(py), "-m", "pip", "install", "--upgrade", "pip"])
    run([str(py), "-m", "pip", "install", "-r", str(REQUIREMENTS)])
    run([str(py), "-m", "playwright", "install", "chromium"])

    print(
        f"""
Setup complete on this machine.

Next:
  python run_drafts.py --file your-guests.xlsx --from-sr 1 --to-sr 300

Project folder: {ROOT}
"""
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
