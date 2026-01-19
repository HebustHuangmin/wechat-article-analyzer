#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Sequential runner: main.py then excel_analyzer.py with a pause."""
import subprocess
import time
import sys
from pathlib import Path

# Use the project venv python if available
VENV_PY = Path(__file__).parent / ".venv" / "Scripts" / "python.exe"
PYTHON = str(VENV_PY) if VENV_PY.exists() else sys.executable

ROOT = Path(__file__).parent
MAIN = ROOT / "main.py"
ANALYZER = ROOT / "excel_analyzer.py"


def run_cmd(label, script_path):
    print(f"[RUN] {label}: {script_path}")
    result = subprocess.run([PYTHON, str(script_path)], cwd=ROOT)
    if result.returncode != 0:
        raise SystemExit(f"{label} failed with exit code {result.returncode}")


def main():
    run_cmd("main.py", MAIN)
    print("[INFO] main.py finished, sleeping 5 seconds...")
    time.sleep(5)
    run_cmd("excel_analyzer.py", ANALYZER)
    print("[DONE] All tasks completed.")


if __name__ == "__main__":
    main()
