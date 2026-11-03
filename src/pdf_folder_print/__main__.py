#!/usr/bin/env python3
"""
Entry point for `python -m pdf_folder_print` and the PyInstaller EXE.

Two modes:
  • Double-click / no arguments  →  GUI (tkinter)
  • With arguments               →  CLI (Rich console)
"""

import os
import sys

# ── Windows console UTF-8 ─────────────────────────────────────────────
if os.name == "nt":
    os.system("")
    try:
        import ctypes

        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        ctypes.windll.kernel32.SetConsoleCP(65001)
    except Exception:
        pass
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def main() -> None:
    # Absolute imports (no relative ones) because this module is executed as
    # the top-level script "__main__" by PyInstaller, where the package
    # context does not exist.
    from pdf_folder_print.cli import cli_main
    from pdf_folder_print.gui import gui_main

    if len(sys.argv) > 1 and "--gui" not in sys.argv:
        cli_main()
    else:
        gui_main()


if __name__ == "__main__":
    main()
