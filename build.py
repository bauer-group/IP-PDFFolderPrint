#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build script for pdf_folder_print.exe
=======================================
Creates a standalone Windows EXE with all dependencies bundled.
The EXE contains both CLI (Rich) and GUI (tkinter) in a single file:
  • Double-click         →  GUI
  • With CLI arguments   →  Command line with Rich output

  python build.py              → Build the EXE (local, with venv)
  python build.py --ci         → Build the EXE (CI, uses system Python)
  python build.py --clean      → Remove build artifacts
  python build.py --clean-all  → Remove everything (incl. venv, dist)
  python build.py --check      → Check dependencies only
  python build.py --rebuild-venv → Recreate venv and build

Requirements: Python 3.10+ with pip
Output:       dist/pdf_folder_print.exe
"""

import argparse
import shutil
import subprocess
import sys
import venv
from pathlib import Path

# ── Configuration ─────────────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).parent.resolve()
VENV_DIR = SCRIPT_DIR / ".venv"
DIST_DIR = SCRIPT_DIR / "dist"
BUILD_DIR = SCRIPT_DIR / "build"
SPEC_FILE = SCRIPT_DIR / "pdf_folder_print.spec"
MAIN_SCRIPT = SCRIPT_DIR / "src" / "pdf_folder_print.py"
EXE_NAME = "pdf_folder_print"

DEPENDENCIES = ["pymupdf", "Pillow", "pywin32", "rich", "pyinstaller"]

if sys.platform == "win32":
    VENV_PYTHON = VENV_DIR / "Scripts" / "python.exe"
    VENV_PIP = VENV_DIR / "Scripts" / "pip.exe"
else:
    VENV_PYTHON = VENV_DIR / "bin" / "python"
    VENV_PIP = VENV_DIR / "bin" / "pip"


def log(msg: str, icon: str = "→"):
    print(f"  {icon} {msg}")


def run(cmd: list[str], desc: str, check: bool = True) -> subprocess.CompletedProcess:
    log(desc, "⏳")
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if check and result.returncode != 0:
            log(f"FAILED (exit {result.returncode})", "❌")
            if result.stdout.strip():
                print(f"\n--- stdout (last 2000 chars) ---\n{result.stdout[-2000:]}")
            if result.stderr.strip():
                print(f"\n--- stderr (last 2000 chars) ---\n{result.stderr[-2000:]}")
            sys.exit(1)
        return result
    except FileNotFoundError:
        log(f"Command not found: {cmd[0]}", "❌")
        sys.exit(1)


def ensure_venv():
    if VENV_PYTHON.exists():
        log("Virtual environment exists", "✅")
        return
    log("Creating virtual environment...")
    try:
        venv.create(str(VENV_DIR), with_pip=True, clear=True)
    except Exception as e:
        log(f"venv creation failed: {e}", "❌")
        sys.exit(1)
    if not VENV_PYTHON.exists():
        log(f"Python not found in venv: {VENV_PYTHON}", "❌")
        sys.exit(1)
    log("Virtual environment created", "✅")


def install_deps(ci: bool = False):
    if ci:
        run(
            [sys.executable, "-m", "pip", "install", "-r", str(SCRIPT_DIR / "requirements.txt")],
            "Installing from requirements.txt (CI)",
        )
    else:
        run([str(VENV_PIP), "install", "--upgrade", "pip"], "Upgrading pip")
        run(
            [str(VENV_PIP), "install", "--upgrade"] + DEPENDENCIES,
            f"Installing: {', '.join(DEPENDENCIES)}",
        )
    log("All dependencies installed", "✅")


def check_main_script():
    if not MAIN_SCRIPT.exists():
        log(f"Main script not found: {MAIN_SCRIPT}", "❌")
        log("build.py and pdf_folder_print.py must be in the same directory.", "💡")
        sys.exit(1)
    log(f"Main script: {MAIN_SCRIPT.name}", "✅")


def build_exe(ci: bool = False):
    if ci:
        pyinstaller_cmd = [sys.executable, "-m", "PyInstaller"]
    else:
        pyinstaller = VENV_DIR / ("Scripts" if sys.platform == "win32" else "bin") / "pyinstaller"
        if sys.platform == "win32":
            pyinstaller = pyinstaller.with_suffix(".exe")

        if pyinstaller.exists():
            pyinstaller_cmd = [str(pyinstaller)]
        else:
            pyinstaller_cmd = [str(VENV_PYTHON), "-m", "PyInstaller"]

    cmd = pyinstaller_cmd + [
        "--onefile",
        "--console",
        "--name", EXE_NAME,
        "--clean",
        "--noconfirm",
        "--distpath", str(DIST_DIR),
        "--workpath", str(BUILD_DIR),
        "--specpath", str(SCRIPT_DIR),
        "--hidden-import=win32print",
        "--hidden-import=win32ui",
        "--hidden-import=win32api",
        "--hidden-import=pywintypes",
        "--hidden-import=PIL",
        "--hidden-import=PIL.Image",
        "--hidden-import=PIL.ImageWin",
        "--hidden-import=fitz",
        "--hidden-import=rich",
        "--hidden-import=rich.console",
        "--hidden-import=rich.table",
        "--hidden-import=rich.panel",
        "--hidden-import=rich.tree",
        "--hidden-import=rich.progress",
        "--hidden-import=rich.text",
        "--hidden-import=rich.box",
        "--collect-all", "pymupdf",
        "--collect-all", "rich",
        str(MAIN_SCRIPT),
    ]

    run(cmd, "PyInstaller build (1-2 minutes)...")

    exe_path = DIST_DIR / f"{EXE_NAME}.exe"
    if exe_path.exists():
        size_mb = exe_path.stat().st_size / (1024 * 1024)
        log(f"EXE created: {exe_path} ({size_mb:.1f} MB)", "✅")
    else:
        log("EXE was not created – check build log", "❌")
        sys.exit(1)


def clean_build():
    removed = []
    if BUILD_DIR.exists():
        shutil.rmtree(BUILD_DIR)
        removed.append("build/")
    if SPEC_FILE.exists():
        SPEC_FILE.unlink()
        removed.append(".spec")
    log(f"Cleaned: {', '.join(removed)}" if removed else "Nothing to clean", "🧹")


def clean_all():
    for d in [BUILD_DIR, DIST_DIR, VENV_DIR]:
        if d.exists():
            shutil.rmtree(d)
            log(f"Removed: {d.name}/", "🧹")
    if SPEC_FILE.exists():
        SPEC_FILE.unlink()
        log("Removed: .spec", "🧹")


def main():
    parser = argparse.ArgumentParser(description=f"Build script for {EXE_NAME}.exe")
    parser.add_argument("--clean", action="store_true", help="Remove build artifacts")
    parser.add_argument("--clean-all", action="store_true", help="Remove everything (incl. venv/dist)")
    parser.add_argument("--check", action="store_true", help="Check dependencies only")
    parser.add_argument("--rebuild-venv", action="store_true", help="Recreate venv from scratch")
    parser.add_argument("--ci", action="store_true", help="CI mode: skip venv, use system Python")
    args = parser.parse_args()

    print()
    print("═" * 60)
    print(f"  🔨 Build: {EXE_NAME}.exe")
    print("═" * 60)
    print()

    if args.clean_all:
        clean_all()
        log("Done.", "✅")
        return

    if args.clean:
        clean_build()
        log("Done.", "✅")
        return

    if args.rebuild_venv and VENV_DIR.exists():
        log("Removing existing venv...")
        shutil.rmtree(VENV_DIR)

    log(f"Python: {sys.version.split()[0]} ({sys.executable})")
    check_main_script()

    print()
    log("Preparing environment...", "📦")
    if args.ci:
        log("CI mode: using system Python", "🔧")
        install_deps(ci=True)
    else:
        ensure_venv()
        install_deps()

    if args.check:
        print()
        log("Check OK – ready to build.", "✅")
        return

    print()
    log("Building EXE...", "🔨")
    build_exe(ci=args.ci)

    print()
    clean_build()

    print()
    print("═" * 60)
    print("  ✅ BUILD COMPLETE")
    print()
    print(f"  EXE:  dist\\{EXE_NAME}.exe")
    print()
    print("  Usage:")
    print("    Double-click                        →  GUI")
    print(f'    {EXE_NAME}.exe "C:\\path\\to\\folder"  →  CLI')
    print(f"    {EXE_NAME}.exe --dry-run             →  CLI preview")
    print(f"    {EXE_NAME}.exe --list-printers")
    print("═" * 60)
    print()


if __name__ == "__main__":
    main()
