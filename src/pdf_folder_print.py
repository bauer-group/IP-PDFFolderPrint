#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF Folder Print – Structured PDF Batch Printing
==================================================
Recursively prints all PDFs from a folder and its subfolders
in a structured, deterministic order:
  1. PDFs in the root folder (alphabetically)
  2. Subfolders alphabetically, each with its PDFs (alphabetically)

Page scaling (Fit-to-Page, like Adobe Reader):
  - Undersized pages are scaled up, oversized pages scaled down
  - Aspect ratio is preserved, content is centered
  - Orientation is handled per page by rotating the rendered
    bitmap when that yields better page utilization

Two modes:
  • Double-click / no arguments  →  GUI (tkinter)
  • With arguments               →  CLI (Rich console)

Dependencies:
  pip install pymupdf Pillow pywin32 rich

Build standalone EXE:
  python build.py
"""

import argparse
import os
import sys
import threading
import time
from pathlib import Path

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

# ── Dependencies ──────────────────────────────────────────────────────
_missing = []
try:
    import fitz
except ImportError:
    _missing.append("pymupdf")
try:
    from PIL import Image, ImageWin
except ImportError:
    _missing.append("Pillow")
try:
    import win32print
    import win32ui
except ImportError:
    _missing.append("pywin32")
try:
    from rich.console import Console
    from rich.table import Table
    from rich.panel import Panel
    from rich.tree import Tree
    from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn, TaskProgressColumn, TimeRemainingColumn
    from rich.text import Text
    from rich import box
except ImportError:
    _missing.append("rich")

if _missing:
    msg = (
        f"Missing packages: {', '.join(_missing)}\n"
        f"Install: pip install {' '.join(_missing)}"
    )
    print(f"ERROR: {msg}", file=sys.stderr)
    try:
        import tkinter as tk
        from tkinter import messagebox
        _r = tk.Tk()
        _r.withdraw()
        messagebox.showerror("Missing Dependencies", msg)
    except Exception:
        pass
    sys.exit(1)


# ══════════════════════════════════════════════════════════════════════
#  CORE LOGIC
# ══════════════════════════════════════════════════════════════════════

PDF_EXTENSIONS = {".pdf"}
FALLBACK_DPI = 300


def collect_pdfs(root: Path) -> list[tuple[str, Path]]:
    """
    Collect PDFs: root folder first, then subfolders alphabetically.
    Returns: list of (group_name, pdf_path) tuples.
    """
    results: list[tuple[str, Path]] = []

    root_pdfs = sorted(
        [f for f in root.iterdir() if f.is_file() and f.suffix.lower() in PDF_EXTENSIONS],
        key=lambda p: p.name.lower(),
    )
    for pdf in root_pdfs:
        results.append(("(Root)", pdf))

    subdirs = sorted(
        [d for d in root.iterdir() if d.is_dir()],
        key=lambda d: d.name.lower(),
    )
    for subdir in subdirs:
        sub_pdfs = sorted(
            [f for f in subdir.iterdir() if f.is_file() and f.suffix.lower() in PDF_EXTENSIONS],
            key=lambda p: p.name.lower(),
        )
        for pdf in sub_pdfs:
            results.append((subdir.name, pdf))

    return results


def get_available_printers() -> list[str]:
    return [
        p[2]
        for p in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        )
    ]


def get_default_printer() -> str:
    return win32print.GetDefaultPrinter()


def get_pdf_info(pdf_path: Path) -> tuple[int, float, float]:
    """Page count and first page size in mm."""
    try:
        doc = fitz.open(str(pdf_path))
        n = len(doc)
        if n > 0:
            r = doc[0].rect
            w_mm = r.width / 72 * 25.4
            h_mm = r.height / 72 * 25.4
        else:
            w_mm, h_mm = 0, 0
        doc.close()
        return n, w_mm, h_mm
    except Exception:
        return 0, 0, 0


def print_pdf_file(
    pdf_path: Path,
    printer_name: str,
    dpi_override: int | None = None,
) -> tuple[bool, int, str]:
    """
    Render PDF and print via Windows spooler.

    Orientation handling per page (no DEVMODE/ResetDC needed):
      1. Render the PDF page to a bitmap
      2. Compare two options: normal vs. rotated 90°
      3. Pick whichever yields a higher scale factor (= better page fill)
      4. Fit-to-Page, center, print

    This gives the same result as switching printer orientation
    but works on all pywin32 versions without ResetDC.

    Returns: (success, page_count, error_message).
    """
    try:
        doc = fitz.open(str(pdf_path))
    except Exception as e:
        return False, 0, f"Cannot open PDF: {e}"

    page_count = len(doc)
    if page_count == 0:
        doc.close()
        return True, 0, ""

    # Create printer DC
    try:
        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer_name)
    except Exception as e:
        doc.close()
        return False, 0, f"Printer error: {e}"

    # Printer dimensions (fixed, portrait orientation)
    printer_w_px = hdc.GetDeviceCaps(8)    # HORZRES
    printer_h_px = hdc.GetDeviceCaps(10)   # VERTRES
    printer_dpi_x = hdc.GetDeviceCaps(88) or FALLBACK_DPI
    printer_dpi_y = hdc.GetDeviceCaps(90) or FALLBACK_DPI

    printed_page = 0

    try:
        hdc.StartDoc(pdf_path.name)

        for page_num in range(page_count):
            page = doc[page_num]
            pdf_w_pt = page.rect.width
            pdf_h_pt = page.rect.height
            if pdf_w_pt <= 0 or pdf_h_pt <= 0:
                continue

            # ── Render PDF page to bitmap ──────────────────────
            if dpi_override:
                zoom = dpi_override / 72.0
            else:
                # Render at printer-native DPI (based on longer axis)
                zoom_x = printer_w_px / pdf_w_pt
                zoom_y = printer_h_px / pdf_h_pt
                zoom = max(zoom_x, zoom_y)  # render large enough for either orientation

            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # ── Determine best orientation ─────────────────────
            # Try normal and rotated, pick whichever fills more of the page
            img_w, img_h = img.size

            scale_normal = min(printer_w_px / img_w, printer_h_px / img_h)
            scale_rotated = min(printer_w_px / img_h, printer_h_px / img_w)

            if scale_rotated > scale_normal:
                # Rotating gives better page utilization → rotate 90° CCW
                img = img.transpose(Image.Transpose.ROTATE_90)
                img_w, img_h = img.size
                scale = scale_rotated
            else:
                scale = scale_normal

            # ── Fit-to-Page ────────────────────────────────────
            out_w = int(img_w * scale)
            out_h = int(img_h * scale)

            # Center on printable area
            x = (printer_w_px - out_w) // 2
            y = (printer_h_px - out_h) // 2

            hdc.StartPage()
            dib = ImageWin.Dib(img)
            dib.draw(hdc.GetHandleOutput(), (x, y, x + out_w, y + out_h))
            hdc.EndPage()
            printed_page = page_num + 1

        hdc.EndDoc()

    except Exception as e:
        doc.close()
        try:
            hdc.AbortDoc()
        except Exception:
            pass
        hdc.DeleteDC()
        return False, page_count, f"Print error page {printed_page + 1}: {e}"

    finally:
        try:
            hdc.DeleteDC()
        except Exception:
            pass

    doc.close()
    return True, page_count, ""


# ══════════════════════════════════════════════════════════════════════
#  CLI MODE (Rich)
# ══════════════════════════════════════════════════════════════════════

def cli_main():
    parser = argparse.ArgumentParser(
        description="PDF Folder Print – Structured PDF batch printing",
    )
    parser.add_argument("folder", nargs="?", help="Path to the folder containing PDFs")
    parser.add_argument("--gui", action="store_true", help="Launch GUI mode")
    parser.add_argument("--dry-run", action="store_true", help="List files without printing")
    parser.add_argument("--printer", default=None, help="Printer name (default: system default)")
    parser.add_argument("--dpi", type=int, default=None, help="Render DPI override (default: printer-native)")
    parser.add_argument("--delay", type=float, default=1.0, help="Delay between print jobs in seconds")
    parser.add_argument("--list-printers", action="store_true", help="List available printers and exit")
    args = parser.parse_args()

    console = Console()

    if args.gui:
        gui_main()
        return

    if args.list_printers:
        default = get_default_printer()
        table = Table(title="Available Printers", box=box.ROUNDED, title_style="bold cyan")
        table.add_column("Printer", style="white")
        table.add_column("Status", style="dim")
        for name in get_available_printers():
            if name == default:
                table.add_row(f"[bold]{name}[/bold]", "[green]● default[/green]")
            else:
                table.add_row(name, "")
        console.print()
        console.print(table)
        console.print()
        sys.exit(0)

    if not args.folder:
        parser.print_help()
        sys.exit(1)

    root = Path(args.folder)
    if not root.is_dir():
        console.print(f"[bold red]ERROR:[/] Folder not found: {root}")
        sys.exit(1)

    if args.printer:
        available = get_available_printers()
        if args.printer not in available:
            console.print(f"[bold red]ERROR:[/] Printer [bold]'{args.printer}'[/] not found.\n")
            for p in available:
                console.print(f"  • {p}")
            sys.exit(1)
        printer_name = args.printer
    else:
        printer_name = get_default_printer()

    pdfs = collect_pdfs(root)
    if not pdfs:
        console.print("[yellow]No PDF files found.[/]")
        sys.exit(0)

    groups: dict[str, list[Path]] = {}
    for group, pdf_path in pdfs:
        groups.setdefault(group, []).append(pdf_path)

    dpi_label = f"{args.dpi} (manual)" if args.dpi else "printer-native"
    mode_label = "[yellow]DRY-RUN[/yellow] (preview only)" if args.dry_run else "[green bold]PRINT[/green bold]"

    info_table = Table(box=None, show_header=False, padding=(0, 2))
    info_table.add_column("Key", style="dim", no_wrap=True)
    info_table.add_column("Value", style="bold")
    info_table.add_row("Folder", root.name)
    info_table.add_row("Printer", printer_name)
    info_table.add_row("Render DPI", dpi_label)
    info_table.add_row("Scaling", "Fit-to-Page + auto orientation")
    info_table.add_row("Mode", mode_label)
    info_table.add_row("Files", f"{len(pdfs)} in {len(groups)} groups")

    console.print()
    console.print(Panel(info_table, title="[bold]PDF Folder Print[/]", border_style="cyan", expand=False))

    tree = Tree(f"[bold cyan]{root.name}[/]")
    for gname, gpaths in groups.items():
        if gname == "(Root)":
            for p in gpaths:
                tree.add(f"[dim]{p.name}[/]")
        else:
            branch = tree.add(f"[yellow]{gname}/[/] [dim]({len(gpaths)} files)[/]")
            for p in gpaths:
                branch.add(f"[dim]{p.name}[/]")

    console.print()
    console.print(tree)
    console.print()

    if not args.dry_run:
        if not console.input(f"[bold]Start printing {len(pdfs)} files? [y/N][/] ").strip().lower() in ("y", "yes", "j", "ja"):
            console.print("[dim]Cancelled.[/]")
            sys.exit(0)
        console.print()

    ok_count = 0
    fail_count = 0
    total_pages = 0
    current_group = None

    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(bar_width=30),
        TaskProgressColumn(),
        TextColumn("•"),
        TimeRemainingColumn(),
        console=console,
        transient=False,
    ) as progress:

        task = progress.add_task("[cyan]Processing...", total=len(pdfs))

        for i, (group, pdf_path) in enumerate(pdfs):
            if group != current_group:
                current_group = group
                progress.console.print(f"\n[bold yellow]📁 {group}[/]")

            if args.dry_run:
                n, w, h = get_pdf_info(pdf_path)
                if n > 0:
                    orient = "landscape" if w > h else "portrait"
                    orient_color = "magenta" if w > h else "blue"
                    progress.console.print(
                        f"   [dim]📄[/] {pdf_path.name}  "
                        f"[dim]→[/]  {n} page{'s' if n != 1 else ''}  "
                        f"[dim]{w:.0f}×{h:.0f}mm[/]  "
                        f"[{orient_color}]{orient}[/]"
                    )
                else:
                    progress.console.print(f"   [dim]📄[/] {pdf_path.name}  [red]→ unreadable[/]")
                ok_count += 1
            else:
                progress.update(task, description=f"[cyan]{pdf_path.name}")
                ok, pages, err = print_pdf_file(pdf_path, printer_name, args.dpi)
                if ok:
                    progress.console.print(
                        f"   [green]✓[/] {pdf_path.name}  "
                        f"[dim]{pages} page{'s' if pages != 1 else ''}[/]"
                    )
                    ok_count += 1
                    total_pages += pages
                else:
                    progress.console.print(f"   [bold red]✗[/] {pdf_path.name}  [red]{err}[/]")
                    fail_count += 1
                if i < len(pdfs) - 1 and args.delay > 0:
                    time.sleep(args.delay)

            progress.update(task, advance=1)

        progress.update(task, description="[green]Done")

    console.print()

    if args.dry_run:
        result_table = Table(box=box.ROUNDED, border_style="yellow", title="[bold]Dry-Run Summary[/]")
        result_table.add_column("", style="dim")
        result_table.add_column("", style="bold")
        result_table.add_row("Files", f"{ok_count}")
        result_table.add_row("Groups", f"{len(groups)}")
        result_table.add_row("Status", "[yellow]Preview only – nothing printed[/]")
        console.print(result_table)
    else:
        style = "green" if fail_count == 0 else "red"
        result_table = Table(box=box.ROUNDED, border_style=style, title="[bold]Print Summary[/]")
        result_table.add_column("", style="dim")
        result_table.add_column("", style="bold")
        result_table.add_row("Printed", f"[green]{ok_count} files ({total_pages} pages)[/]")
        if fail_count:
            result_table.add_row("Failed", f"[bold red]{fail_count} files[/]")
        result_table.add_row("Printer", printer_name)
        result_table.add_row("Status", "[green]✓ Complete[/]" if fail_count == 0 else "[red]✗ Errors occurred[/]")
        console.print(result_table)

    console.print()

    if not args.dry_run:
        console.input("[dim]Press Enter to exit...[/]")

    sys.exit(1 if fail_count > 0 else 0)


# ══════════════════════════════════════════════════════════════════════
#  GUI MODE (tkinter)
# ══════════════════════════════════════════════════════════════════════

def gui_main():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    root = tk.Tk()
    root.title("PDF Folder Print")
    root.geometry("820x680")
    root.minsize(700, 550)
    root.configure(bg="#f5f5f5")

    try:
        root.iconbitmap(default="")
    except Exception:
        pass

    style = ttk.Style()
    style.theme_use("clam")

    BG = "#f5f5f5"
    FRAME_BG = "#ffffff"
    ACCENT = "#1a56db"
    SUCCESS = "#16a34a"
    ERROR = "#dc2626"

    style.configure("TFrame", background=BG)
    style.configure("Card.TFrame", background=FRAME_BG, relief="solid", borderwidth=1)
    style.configure("TLabel", background=BG, font=("Segoe UI", 10))
    style.configure("Header.TLabel", background=BG, font=("Segoe UI", 14, "bold"))
    style.configure("Card.TLabel", background=FRAME_BG, font=("Segoe UI", 10))
    style.configure("CardBold.TLabel", background=FRAME_BG, font=("Segoe UI", 10, "bold"))
    style.configure("Status.TLabel", background=BG, font=("Segoe UI", 10))
    style.configure("Accent.TButton", font=("Segoe UI", 11, "bold"), padding=(20, 10))
    style.configure("Secondary.TButton", font=("Segoe UI", 10), padding=(12, 6))

    # ── State ──────────────────────────────────────────────────
    folder_var = tk.StringVar()
    printer_var = tk.StringVar()
    is_running = threading.Event()

    printers = get_available_printers()
    default_printer = get_default_printer()
    printer_var.set(default_printer)

    # ── Header ─────────────────────────────────────────────────
    header_frame = ttk.Frame(root)
    header_frame.pack(fill="x", padx=20, pady=(15, 5))
    ttk.Label(header_frame, text="🖨️ PDF Folder Print", style="Header.TLabel").pack(side="left")
    ttk.Label(header_frame, text="Structured PDF Batch Printing", style="TLabel", foreground="#888").pack(side="right")

    # ── Settings card ──────────────────────────────────────────
    settings_outer = ttk.Frame(root)
    settings_outer.pack(fill="x", padx=20, pady=(10, 5))
    settings = ttk.Frame(settings_outer, style="Card.TFrame")
    settings.pack(fill="x", ipady=8)

    # Folder
    row_folder = ttk.Frame(settings, style="Card.TFrame")
    row_folder.pack(fill="x", padx=15, pady=(10, 5))
    ttk.Label(row_folder, text="📁 Folder:", style="CardBold.TLabel").pack(side="left")
    btn_browse = ttk.Button(row_folder, text="Browse...", style="Secondary.TButton", command=lambda: browse_folder())
    btn_browse.pack(side="right", padx=(5, 0))
    entry_folder = ttk.Entry(row_folder, textvariable=folder_var, font=("Segoe UI", 10))
    entry_folder.pack(side="right", fill="x", expand=True, padx=(10, 5))

    # Printer
    row_printer = ttk.Frame(settings, style="Card.TFrame")
    row_printer.pack(fill="x", padx=15, pady=(5, 10))
    ttk.Label(row_printer, text="🖨️ Printer:", style="CardBold.TLabel").pack(side="left")
    combo_printer = ttk.Combobox(row_printer, textvariable=printer_var, values=printers, state="readonly", font=("Segoe UI", 10), width=50)
    combo_printer.pack(side="right", fill="x", expand=True, padx=(10, 0))

    # ── Buttons ────────────────────────────────────────────────
    btn_frame = ttk.Frame(root)
    btn_frame.pack(fill="x", padx=20, pady=8)
    btn_print = ttk.Button(btn_frame, text="▶  Print", style="Accent.TButton", command=lambda: start_print())
    btn_print.pack(side="left")
    btn_preview = ttk.Button(btn_frame, text="🔍 Preview", style="Secondary.TButton", command=lambda: start_preview())
    btn_preview.pack(side="left", padx=(10, 0))
    status_label = ttk.Label(btn_frame, text="", style="Status.TLabel")
    status_label.pack(side="right")

    # ── Progress ───────────────────────────────────────────────
    progress_var = tk.DoubleVar(value=0)
    progress = ttk.Progressbar(root, variable=progress_var, maximum=100)
    progress.pack(fill="x", padx=20, pady=(0, 5))

    # ── Log ────────────────────────────────────────────────────
    log_frame = ttk.Frame(root)
    log_frame.pack(fill="both", expand=True, padx=20, pady=(0, 15))
    log_text = tk.Text(log_frame, font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4", insertbackground="#d4d4d4", selectbackground="#264f78", relief="flat", wrap="word", state="disabled", padx=10, pady=8)
    log_text.pack(fill="both", expand=True, side="left")
    scrollbar = ttk.Scrollbar(log_frame, command=log_text.yview)
    scrollbar.pack(fill="y", side="right")
    log_text.configure(yscrollcommand=scrollbar.set)

    log_text.tag_configure("header", foreground="#569cd6", font=("Consolas", 9, "bold"))
    log_text.tag_configure("group", foreground="#dcdcaa")
    log_text.tag_configure("success", foreground="#6a9955")
    log_text.tag_configure("error", foreground="#f44747")
    log_text.tag_configure("info", foreground="#9cdcfe")
    log_text.tag_configure("dim", foreground="#808080")

    # ── Helpers ────────────────────────────────────────────────
    def log_msg(text: str, tag: str = ""):
        log_text.configure(state="normal")
        log_text.insert("end", text + "\n", (tag,) if tag else ())
        log_text.see("end")
        log_text.configure(state="disabled")

    def log_clear():
        log_text.configure(state="normal")
        log_text.delete("1.0", "end")
        log_text.configure(state="disabled")

    def set_status(text: str, color: str = "#888"):
        status_label.configure(text=text, foreground=color)

    def set_running(running: bool):
        s = "disabled" if running else "normal"
        if running:
            is_running.set()
        else:
            is_running.clear()
        btn_print.configure(state=s)
        btn_preview.configure(state=s)
        btn_browse.configure(state=s)
        combo_printer.configure(state=s if running else "readonly")

    def browse_folder():
        initial = folder_var.get() or ""
        folder = filedialog.askdirectory(title="Select folder", initialdir=initial if os.path.isdir(initial) else "")
        if folder:
            folder_var.set(folder)

    def validate() -> Path | None:
        folder = folder_var.get().strip()
        if not folder:
            messagebox.showwarning("No folder", "Please select a folder first.")
            return None
        p = Path(folder)
        if not p.is_dir():
            messagebox.showerror("Folder not found", f"Folder does not exist:\n{p}")
            return None
        return p

    # ── Preview (always dry-run) ───────────────────────────────
    def start_preview():
        root_path = validate()
        if not root_path:
            return
        log_clear()
        set_running(True)
        progress_var.set(0)
        thread = threading.Thread(target=run_print_job, args=(root_path, printer_var.get(), True), daemon=True)
        thread.start()

    # ── Print (always real) ────────────────────────────────────
    def start_print():
        root_path = validate()
        if not root_path:
            return
        printer = printer_var.get()
        if not messagebox.askyesno(
            "Start printing?",
            f"Print all PDFs in\n\n{root_path.name}\n\nto printer\n\n{printer}\n\n?",
        ):
            return
        log_clear()
        set_running(True)
        progress_var.set(0)
        thread = threading.Thread(target=run_print_job, args=(root_path, printer, False), daemon=True)
        thread.start()

    # ── Worker ─────────────────────────────────────────────────
    def run_print_job(root_path: Path, printer: str, dry_run: bool):
        try:
            _run_print_job_inner(root_path, printer, dry_run)
        except Exception as e:
            root.after(0, lambda: log_msg(f"\n⚠ Unexpected error: {e}", "error"))
        finally:
            root.after(0, lambda: set_running(False))

    def _run_print_job_inner(root_path: Path, printer: str, dry_run: bool):
        pdfs = collect_pdfs(root_path)
        if not pdfs:
            root.after(0, lambda: log_msg("No PDF files found.", "error"))
            root.after(0, lambda: set_status("No files", ERROR))
            return

        groups: dict[str, int] = {}
        for g, _ in pdfs:
            groups[g] = groups.get(g, 0) + 1

        mode = "PREVIEW" if dry_run else "PRINTING"
        root.after(0, lambda: log_msg(f"{'═' * 60}", "dim"))
        root.after(0, lambda: log_msg(f"  Folder:      {root_path.name}", "header"))
        root.after(0, lambda: log_msg(f"  Printer:     {printer}", "header"))
        root.after(0, lambda: log_msg(f"  Mode:        {mode}", "header"))
        root.after(0, lambda: log_msg(f"  Scaling:     Fit-to-Page + auto orientation", "header"))
        root.after(0, lambda: log_msg(f"  Files:       {len(pdfs)} in {len(groups)} groups", "header"))
        root.after(0, lambda: log_msg(f"{'═' * 60}", "dim"))

        for gname, count in groups.items():
            root.after(0, lambda g=gname, c=count: log_msg(f"  📁 {g}: {c} file(s)", "info"))

        root.after(0, lambda: log_msg(""))
        root.after(0, lambda: set_status(f"0 / {len(pdfs)}...", ACCENT))

        current_group = None
        ok_count = 0
        fail_count = 0
        total_pages = 0

        for i, (group, pdf_path) in enumerate(pdfs):
            if group != current_group:
                current_group = group
                root.after(0, lambda g=group: log_msg(f"\n{'─' * 40}", "dim"))
                root.after(0, lambda g=group: log_msg(f"📁 {g}", "group"))
                root.after(0, lambda g=group: log_msg(f"{'─' * 40}", "dim"))

            if dry_run:
                n, w, h = get_pdf_info(pdf_path)
                if n > 0:
                    orient = "landscape" if w > h else "portrait"
                    info = f"{n} page{'s' if n != 1 else ''}, {w:.0f}×{h:.0f}mm ({orient})"
                else:
                    info = "?"
                root.after(0, lambda p=pdf_path.name, inf=info: log_msg(f"  📄 {p}  →  {inf}", "info"))
                ok_count += 1
            else:
                root.after(0, lambda p=pdf_path.name: log_msg(f"  🖨️ {p}...", "info"))
                ok, pages, err = print_pdf_file(pdf_path, printer)
                if ok:
                    root.after(0, lambda pg=pages: log_msg(f"     ✓ {pg} page{'s' if pg != 1 else ''}", "success"))
                    ok_count += 1
                    total_pages += pages
                else:
                    root.after(0, lambda e=err: log_msg(f"     ✗ {e}", "error"))
                    fail_count += 1
                if i < len(pdfs) - 1:
                    time.sleep(1.0)

            pct = (i + 1) / len(pdfs) * 100
            root.after(0, lambda v=pct: progress_var.set(v))
            root.after(0, lambda idx=i + 1, t=len(pdfs): set_status(f"{idx} / {t}...", ACCENT))

        root.after(0, lambda: log_msg(f"\n{'═' * 60}", "dim"))
        if dry_run:
            root.after(0, lambda: log_msg(f"  🔍 PREVIEW: {ok_count} files found", "header"))
            root.after(0, lambda: set_status(f"Preview: {ok_count} files", ACCENT))
        else:
            summary = f"✅ {ok_count} files ({total_pages} pages) printed"
            if fail_count:
                summary += f", {fail_count} failed"
            root.after(0, lambda: log_msg(f"  {summary}", "success" if fail_count == 0 else "error"))
            color = SUCCESS if fail_count == 0 else ERROR
            root.after(0, lambda: set_status(f"Done – {ok_count} OK, {fail_count} failed", color))
        root.after(0, lambda: log_msg(f"{'═' * 60}", "dim"))
        root.after(0, lambda: progress_var.set(100))

    # ── Welcome ────────────────────────────────────────────────
    log_msg("PDF Folder Print", "header")
    log_msg("Structured PDF batch printing from folder trees.", "header")
    log_msg("", "")
    log_msg("1. Select a folder containing PDFs (with subfolders)", "info")
    log_msg("2. Verify the target printer", "info")
    log_msg("3. Click 'Preview' to inspect or 'Print' to start", "info")
    log_msg("", "")
    log_msg("Print order: root folder first, then subfolders (A→Z)", "dim")
    log_msg("Scaling:     Fit-to-Page with auto portrait/landscape", "dim")

    root.mainloop()


# ══════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════

def main():
    if len(sys.argv) > 1 and "--gui" not in sys.argv:
        cli_main()
    else:
        gui_main()


if __name__ == "__main__":
    main()
