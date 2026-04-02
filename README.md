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
- **Auto orientation** — portrait and landscape detected per page, printer switched via Windows DEVMODE (works even for mixed-orientation PDFs)
- **Dry-run mode** — preview what would be printed, including page counts, dimensions, and orientation, without sending anything to the printer
- **GUI and CLI** — double-click for a graphical interface, pass arguments for command-line automation
- **Rich CLI output** — colored progress bars, folder trees, summary tables via the [Rich](https://github.com/Textualize/rich) library
- **Standalone EXE** — bundle as a single `.exe` with no runtime dependencies via PyInstaller
- **No external viewer** — renders PDFs internally via PyMuPDF, prints directly through the Windows GDI spooler
- **Printer-native DPI** — renders at the printer's own resolution for optimal quality (no bitmap rescaling)

---

## Quick Start

### Option A: Run with Python

```bash
pip install pymupdf Pillow pywin32 rich
python src/pdf_folder_print.py "C:\path\to\folder"
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

# Override render DPI (default: printer-native)
pdf_folder_print.exe --dpi 300 "C:\Exports\March"

# Adjust delay between print jobs (default: 1s)
pdf_folder_print.exe --delay 2.0 "C:\Exports\March"

# List all available printers
pdf_folder_print.exe --list-printers

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

The GUI provides the same functionality with a visual interface:

- **Folder browser** with path entry
- **Printer dropdown** (auto-detects all installed printers, default pre-selected)
- **Dry-run checkbox** + dedicated Preview button
- **Progress bar** with file counter
- **Dark-themed log window** with syntax highlighting
- **Threaded execution** — UI stays responsive during printing
- **Confirmation dialog** before actual printing

---

## How Scaling Works

The scaling logic mirrors Adobe Reader's "Fit" option:

1. **Read PDF page dimensions** from metadata (in PDF points, 1pt = 1/72 inch)
2. **Detect orientation** — if width > height → landscape, else portrait
3. **Switch printer orientation** via `DEVMODE.Orientation` + `ResetDC()` (only when orientation changes)
4. **Re-read printer dimensions** — after `ResetDC`, the printable area (HORZRES/VERTRES) reflects the new orientation
5. **Calculate uniform scale factor** — `min(printable_width / pdf_width, printable_height / pdf_height)` — this scales both up and down
6. **Render at target resolution** — zoom factor computed so the rasterized bitmap matches the target pixel dimensions exactly (no GDI bitmap rescaling)
7. **Center on page** — offset to place the content in the middle of the printable area

This means:
- An A5 PDF on an A4 printer → scaled **up** to fill the page
- An A3 PDF on an A4 printer → scaled **down** to fit
- A landscape PDF → printer switches to landscape, then fit-to-page within that orientation
- Mixed-orientation PDFs → each page triggers orientation switch as needed

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

The resulting `dist/pdf_folder_print.exe` is a single file (~35-45 MB) that runs on any Windows machine without Python installed.

### Build Requirements

- Python 3.10+
- Windows (the build uses `pywin32` which is Windows-only)
- Internet access (for pip to download dependencies)

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

## License

MIT
