# PDF Folder Print

**Structured PDF batch printing from folder trees on Windows.**

Recursively collects all PDFs from a folder and its subfolders, then prints them in a deterministic order with proper page scaling — exactly like hitting "Print → Fit to Page" in Adobe Reader, but for entire folder trees in one go.

No Adobe Reader, no SumatraPDF, no PDF viewer of any kind required.

---

## The Problem

Many business applications (payroll software, ERP exports, accounting tools) dump monthly output into folder structures like this:

```
2026_March/
├── Payslips/
│   ├── Doe_Jane.pdf
│   └── Smith_John.pdf
├── Tax_Reports/
│   └── Summary.pdf
├── Benefits/
│   ├── Health.pdf
│   └── Pension.pdf
└── Overview.pdf
```

Printing these manually means: open each PDF, check orientation, hit print, repeat — across dozens of files and subfolders. Every month.

## The Solution

```
pdf_folder_print.exe "P:\Exports\2026_March"
```

That's it. Every PDF gets printed in order, correctly scaled, with automatic portrait/landscape detection per page.

---

## Features

- **Structured print order** — root folder first, then subfolders alphabetically, each folder's PDFs sorted alphabetically
- **Fit-to-Page scaling** — undersized pages scale up, oversized pages scale down, aspect ratio preserved, content centered
- **Auto orientation** — portrait and landscape detected per page by comparing normal vs. rotated 90° placement and keeping whichever fills the page better (works even for mixed-orientation PDFs)
- **Cancel anytime** — GUI button (or `Esc`) / `Ctrl+C` in the CLI; the spool job of the current file is aborted cleanly
- **Print report** — optional CSV or JSON report (one row per file: timestamp, pages, status, error, duration)
- **Printer status check** — the spooler is queried before printing; offline/paused/no-toner printers warn before jobs are queued (live indicator in the GUI)
- **Retries** — failed files can be retried automatically with backoff (`--retries`, GUI field)
- **Natural sort** — optional file order `file2.pdf` before `file10.pdf` (`--natural-sort`, GUI checkbox)
- **GUI remembers settings** — folder, printer, delay and options are persisted per user (`%APPDATA%\PDFFolderPrint\settings.json`)
- **Dry-run mode** — preview what would be printed, including page counts, dimensions, and orientation, without sending anything to the printer; unreadable PDFs are flagged and fail the run
- **GUI and CLI** — double-click for a graphical interface, pass arguments for command-line automation
- **Rich CLI output** — colored progress bars, folder trees, summary tables via the [Rich](https://github.com/Textualize/rich) library
- **Standalone EXE** — bundle as a single `.exe` with no runtime dependencies via PyInstaller
- **No external viewer** — renders PDFs internally via PyMuPDF, prints directly through the Windows GDI spooler
- **Printer-native DPI** — renders at the printer's own resolution for optimal quality (no bitmap rescaling)

---

## Quick Start

### Option A: Run with Python

```bash
pip install -e .
pdf-folder-print "C:\path\to\folder"
# oder äquivalent:
python -m pdf_folder_print "C:\path\to\folder"
```

### Option B: Build standalone EXE

```bash
python build.py
dist\pdf_folder_print.exe "C:\path\to\folder"
```

### Option C: Just double-click

Double-click `pdf_folder_print.exe` (or run without arguments) to launch the GUI.

---

## Usage

### CLI

```bash
# Print all PDFs to the default printer
pdf_folder_print.exe "C:\Exports\March"

# Dry-run: list all files with page counts, dimensions, and orientation
pdf_folder_print.exe --dry-run "C:\Exports\March"

# Print to a specific printer
pdf_folder_print.exe --printer "HP LaserJet Pro MFP" "C:\Exports\March"

# Skip the confirmation prompt (for automation / scheduled tasks)
pdf_folder_print.exe --yes "C:\Exports\March"

# Retry failed files up to 2 times, write a CSV report
pdf_folder_print.exe --yes --retries 2 --report print_report.csv "C:\Exports\March"

# JSON report instead of CSV
pdf_folder_print.exe --yes --report report.json --report-format json "C:\Exports\March"

# Natural file order (file2.pdf before file10.pdf)
pdf_folder_print.exe --natural-sort "C:\Exports\March"

# Override render DPI (default: printer-native)
pdf_folder_print.exe --dpi 300 "C:\Exports\March"

# Adjust delay between print jobs (default: 0, spooler handles pacing)
pdf_folder_print.exe --delay 2.0 "C:\Exports\March"

# List all available printers
pdf_folder_print.exe --list-printers

# Show the version
pdf_folder_print.exe --version

# Force GUI mode
pdf_folder_print.exe --gui
```

### CLI Output

The CLI uses [Rich](https://github.com/Textualize/rich) for structured, colored terminal output:

- **Summary panel** with folder, printer, mode, and file count
- **Folder tree** showing the complete file structure before printing
- **Progress bar** with spinner, ETA, and percentage
- **Per-file status** with page counts, dimensions, and orientation (portrait/landscape color-coded)
- **Result table** with final summary

### `--list-printers`

Renders a table of all available Windows printers with the default printer highlighted.

### GUI

The GUI (German interface) provides the same functionality with a visual interface:

- **Folder browser** with path entry (press `Enter` to run a preview)
- **Printer dropdown** with **live status indicator** (`✓ ready` / `⚠ offline, paper out, ...`) — problems are confirmed before a print run starts
- **Delay** between print jobs and **Retries** for failed files
- **Natural sort** checkbox for `file2.pdf`-before-`file10.pdf` ordering
- **Report** checkbox with CSV/JSON format and optional report folder (default: next to the printed folder)
- **Cancel button** (or `Esc`) — stops after the current file, the spool job is aborted cleanly
- **Dedicated Preview button** (dry-run) + Print button with confirmation dialog
- **Progress bar** with file counter and status line
- **Dark-themed log window** with syntax highlighting
- **Threaded execution** — UI stays responsive during printing
- **Settings persistence** — folder, printer, delay and options survive restarts (`%APPDATA%\PDFFolderPrint\settings.json`)

---

## How Scaling Works

The scaling logic mirrors Adobe Reader's "Fit" option:

1. **Read PDF page dimensions** from metadata (in PDF points, 1pt = 1/72 inch)
2. **Render the page to a bitmap** — zoom factor computed so the rasterized bitmap matches the printer's native resolution (or a manual `--dpi` override)
3. **Detect the best orientation** — the fit math compares two placements: the bitmap as-is vs. rotated 90°. Whichever yields the higher scale factor (= better page fill) is used. This achieves the same result as switching the printer's orientation, but works on all pywin32 versions without DEVMODE/`ResetDC()`
4. **Calculate uniform scale factor** — `min(printable_width / bitmap_width, printable_height / bitmap_height)` — this scales both up and down
5. **Center on page** — offset to place the content in the middle of the printable area

This means:
- An A5 PDF on an A4 printer → scaled **up** to fill the page
- An A3 PDF on an A4 printer → scaled **down** to fit
- A landscape PDF on portrait paper → the bitmap is rotated 90° so the landscape content fills the page
- Mixed-orientation PDFs → the orientation is decided independently for every page

---

## Print Order

The tool guarantees a deterministic, reproducible print order:

```
Given:
  root/
  ├── zebra.pdf          ← printed 1st (root folder, alphabetical)
  ├── alpha.pdf          ← printed 2nd
  ├── Invoices/
  │   ├── 002.pdf        ← printed 3rd (first subfolder "Invoices", alpha)
  │   └── 001.pdf        ← printed 4th
  └── Reports/
      └── summary.pdf    ← printed 5th (second subfolder "Reports")
```

1. **Root folder PDFs first** — sorted alphabetically (case-insensitive)
2. **Subfolders in alphabetical order** — each subfolder's PDFs sorted alphabetically
3. **One level deep only** — sub-subfolders are not traversed (by design, to match typical export structures)

---

## Building the EXE

The `build.py` script handles everything:

```bash
# Full build (creates venv, installs deps, runs PyInstaller)
python build.py

# Check dependencies without building
python build.py --check

# Rebuild from scratch (recreate venv)
python build.py --rebuild-venv

# Clean build artifacts (keep dist/ and venv/)
python build.py --clean

# Clean everything (remove dist/, venv/, build/)
python build.py --clean-all
```

The resulting `dist/pdf_folder_print.exe` is a single file (~110-120 MB) that runs on any Windows machine without Python installed.

### Build Requirements

- Python 3.10+
- Windows (the build uses `pywin32` which is Windows-only)
- Internet access (for pip to download dependencies)

The EXE and GUI window icons (`assets/icon.ico`, `assets/icon.png`) are the
BAUER GROUP corporate icon, generated from the Corporate Identity project
(`CorporateIdentity-BAUERGROUP`). The multi-size `.ico` was produced from
`bauer-group-icon-512.png` and is committed so the build does not depend on
that repository.

---

## Dependencies

| Package    | Purpose                                        |
|------------|------------------------------------------------|
| `pymupdf`  | PDF rendering (rasterizes pages to bitmaps)    |
| `Pillow`   | Image handling (`PIL.ImageWin` for GDI output) |
| `pywin32`  | Windows printer API (GDI DC, DEVMODE, spooler) |
| `rich`     | CLI output (tables, trees, progress bars)      |
| `tkinter`  | GUI (bundled with Python, no install needed)   |

For building the standalone EXE:

| Package       | Purpose                         |
|---------------|----------------------------------|
| `pyinstaller` | Bundles Python + deps into EXE  |

---

## Requirements

- **OS:** Windows 10/11 (uses Win32 GDI printing API)
- **Python:** 3.10+ (for running from source or building)
- **Printer:** Any printer accessible via Windows (local, network, virtual)

---

## Limitations

- **Windows only** — relies on the Win32 GDI printing API (`win32ui`, `win32print`)
- **One subfolder level** — does not recurse into sub-subfolders (intentional for structured export folders)
- **Raster-based printing** — PDFs are rasterized to bitmaps before printing (at printer-native DPI, so quality is equivalent to any viewer, but vector data is not preserved in the spooler)
- **No duplex/staple control** — uses the printer's default tray and duplex settings (can be changed in printer preferences before printing)

---

## Drag & Drop

Drop a folder onto `pdf_folder_print.exe` (in Explorer) — it starts in CLI mode with that folder and shows the confirmation prompt before printing. Double-click (no arguments) always opens the GUI.

---

## Development

```bash
pip install -r requirements-dev.txt
python -m pytest                        # run unit tests
python -m ruff check src tests build.py # lint
```

The package layout lives under `src/pdf_folder_print/`:

| Module        | Purpose                                                |
|---------------|--------------------------------------------------------|
| `version.py`  | single source of truth for the version                 |
| `core.py`     | file collection, PDF info, natural sort (no win32)     |
| `scaling.py`  | pure fit-to-page / orientation math (no win32)         |
| `printer.py`  | Windows GDI printing, printer status                   |
| `cli.py`      | Rich console interface                                 |
| `gui.py`      | tkinter interface (primary UI, settings persistence)   |
| `report.py`   | optional CSV/JSON print report                         |

Dependency versions are pinned in `pyproject.toml` / `requirements*.txt`
for reproducible builds. Runtime deps in `requirements.txt`, build deps
(`pyinstaller`) in `requirements-build.txt`, dev deps in
`requirements-dev.txt`.

---

## License

MIT
