#!/usr/bin/env python3
"""Cross-platform helpers for venv + project paths (Windows / macOS / Linux)."""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
VENV_DIR = ROOT / ".venv"
REQUIREMENTS = ROOT / "requirements.txt"


def venv_python() -> Path:
    if os.name == "nt":
        return VENV_DIR / "Scripts" / "python.exe"
    return VENV_DIR / "bin" / "python"


def ensure_venv_python() -> Path:
    py = venv_python()
    if py.exists():
        return py
    print("Creating virtual environment in .venv …")
    subprocess.check_call([sys.executable, "-m", "venv", str(VENV_DIR)])
    if not py.exists():
        raise SystemExit(f"Could not find venv Python at {py}")
    return py


def run(cmd: list[str], **kwargs) -> None:
    print(">", " ".join(cmd))
    subprocess.check_call(cmd, **kwargs)
