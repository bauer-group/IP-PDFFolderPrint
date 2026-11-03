"""
GUI mode (tkinter) – the primary user interface.

All functionality is available from the window: folder selection, printer
(with live status), delay/retries, natural sort, optional CSV/JSON report,
preview, printing and cancelling.

``PrintApp`` lives at module level (not nested) so it can be instantiated in
tests. Worker threads never touch tkinter directly – they push
``(kind, ...)`` messages into a ``queue.Queue`` that the UI polls via
``root.after``.
"""

from __future__ import annotations

import json
import os
import queue
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .core import collect_pdfs, get_pdf_info
from .printer import get_available_printers, get_default_printer, print_pdf_file, printer_status
from .report import PrintReport, default_report_path
from .version import __version__


def _asset_path(name: str) -> Path:
    """Locate bundled assets (works from source checkout and frozen EXE)."""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent)) / "assets" / name
    # source layout: src/pdf_folder_print/gui.py → project root/assets
    return Path(__file__).resolve().parents[2] / "assets" / name


def _settings_file() -> Path:
    base = os.environ.get("APPDATA") or str(Path.home() / "AppData" / "Roaming")
    return Path(base) / "PDFFolderPrint" / "settings.json"


def load_settings() -> dict:
    """Last used GUI settings; errors are silently ignored."""
    try:
        return json.loads(_settings_file().read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_settings(data: dict) -> None:
    """Persist GUI settings; errors are silently ignored (cosmetic feature)."""
    try:
        f = _settings_file()
        f.parent.mkdir(parents=True, exist_ok=True)
        f.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


class PrintApp:
    """Main application window."""

    BG = "#f5f5f5"
    FRAME_BG = "#ffffff"
    ACCENT = "#FF8500"  # BAUER GROUP brand orange-500
    ACCENT_DARK = "#EA6D00"  # orange-600 (hover)
    SUCCESS = "#16a34a"
    WARN = "#b45309"
    ERROR = "#dc2626"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.ui_queue: queue.Queue = queue.Queue()
        self.is_running = threading.Event()
        self.cancel_event = threading.Event()

        self.root.title(f"PDF Folder Print {__version__}")
        self.root.geometry("860x720")
        self.root.minsize(740, 580)
        self.root.configure(bg=self.BG)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self._set_window_icon()
        self._setup_style()
        self._build_ui()
        self._show_welcome()
        self.refresh_printer_status()
        self.root.bind("<Escape>", self._on_escape)
        self.root.after(100, self._poll_queue)

    # ── Setup ──────────────────────────────────────────────────
    def _set_window_icon(self):
        try:
            icon = tk.PhotoImage(file=str(_asset_path("icon.png")))
            self.root.iconphoto(True, icon)  # keep a reference to avoid GC
            self._icon_ref = icon
        except Exception:
            pass  # cosmetic only – never block startup over an icon

    def _setup_style(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame", background=self.BG)
        style.configure("Card.TFrame", background=self.FRAME_BG, relief="solid", borderwidth=1)
        style.configure("TLabel", background=self.BG, font=("Segoe UI", 10))
        style.configure("Header.TLabel", background=self.BG, font=("Segoe UI", 14, "bold"))
        style.configure("Card.TLabel", background=self.FRAME_BG, font=("Segoe UI", 10))
        style.configure("Card.TCheckbutton", background=self.FRAME_BG, font=("Segoe UI", 10))
        style.map("Card.TCheckbutton", background=[("active", self.FRAME_BG)])
        style.configure("CardBold.TLabel", background=self.FRAME_BG, font=("Segoe UI", 10, "bold"))
        style.configure("CardDim.TLabel", background=self.FRAME_BG, font=("Segoe UI", 9), foreground="#888888")
        style.configure("Status.TLabel", background=self.BG, font=("Segoe UI", 10))
        style.configure("Accent.TButton", font=("Segoe UI", 11, "bold"), padding=(20, 10))
        style.configure(
            "Accent.TButton",
            background=self.ACCENT,
            foreground="white",
        )
        style.map(
            "Accent.TButton",
            background=[("active", self.ACCENT_DARK), ("disabled", "#FFE0BF")],
            foreground=[("disabled", "#AAAAAA")],
        )
        style.configure("Secondary.TButton", font=("Segoe UI", 10), padding=(12, 6))
        style.configure(
            "Horizontal.TProgressbar",
            troughcolor=self.FRAME_BG,
            background=self.ACCENT,
            lightcolor=self.ACCENT,
            darkcolor=self.ACCENT,
        )

    def _build_ui(self):
        # ── State (pre-filled from last session) ───────────────
        saved = load_settings()
        self.folder_var = tk.StringVar(value=saved.get("folder", ""))
        self.printer_var = tk.StringVar()
        self.delay_var = tk.StringVar(value=str(saved.get("delay", "1.0")))
        self.retries_var = tk.StringVar(value=str(saved.get("retries", "0")))
        self.natural_var = tk.BooleanVar(value=bool(saved.get("natural_sort", False)))
        self.report_var = tk.BooleanVar(value=bool(saved.get("report", False)))
        self.report_fmt_var = tk.StringVar(value=str(saved.get("report_format", "csv")))
        self.report_dir_var = tk.StringVar(value=str(saved.get("report_dir", "")))

        printers = get_available_printers()
        default_printer = get_default_printer()
        self.printer_var.set(default_printer)
        if saved.get("printer") in printers:
            self.printer_var.set(saved["printer"])

        # ── Header ─────────────────────────────────────────────
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill="x", padx=20, pady=(15, 5))
        ttk.Label(header_frame, text="🖨️ PDF Folder Print", style="Header.TLabel").pack(side="left")
        ttk.Label(
            header_frame,
            text=f"v{__version__} · BAUER GROUP",
            style="TLabel",
            foreground="#888",
        ).pack(side="right")

        # ── Settings card ──────────────────────────────────────
        settings_outer = ttk.Frame(self.root)
        settings_outer.pack(fill="x", padx=20, pady=(10, 5))
        settings = ttk.Frame(settings_outer, style="Card.TFrame")
        settings.pack(fill="x", ipady=8)

        # Folder
        row_folder = ttk.Frame(settings, style="Card.TFrame")
        row_folder.pack(fill="x", padx=15, pady=(10, 5))
        ttk.Label(row_folder, text="📁 Folder:", style="CardBold.TLabel").pack(side="left")
        self.btn_browse = ttk.Button(
            row_folder, text="Browse...", style="Secondary.TButton", command=self.browse_folder
        )
        self.btn_browse.pack(side="right", padx=(5, 0))
        entry_folder = ttk.Entry(row_folder, textvariable=self.folder_var, font=("Segoe UI", 10))
        entry_folder.pack(side="right", fill="x", expand=True, padx=(10, 5))
        entry_folder.bind("<Return>", lambda _e: self.start_preview())

        # Printer (with live status indicator)
        row_printer = ttk.Frame(settings, style="Card.TFrame")
        row_printer.pack(fill="x", padx=15, pady=(5, 2))
        ttk.Label(row_printer, text="🖨️ Printer:", style="CardBold.TLabel").pack(side="left")
        self.printer_status_label = ttk.Label(row_printer, text="", style="Card.TLabel")
        self.printer_status_label.pack(side="right")
        self.combo_printer = ttk.Combobox(
            row_printer,
            textvariable=self.printer_var,
            values=printers,
            state="readonly",
            font=("Segoe UI", 10),
            width=46,
        )
        self.combo_printer.pack(side="right", fill="x", expand=True, padx=(10, 8))
        self.combo_printer.bind("<<ComboboxSelected>>", lambda _e: self.refresh_printer_status())

        # Options: delay · retries · natural sort
        row_options = ttk.Frame(settings, style="Card.TFrame")
        row_options.pack(fill="x", padx=15, pady=(4, 2))
        ttk.Label(row_options, text="⏱️ Delay:", style="CardBold.TLabel").pack(side="left")
        ttk.Spinbox(
            row_options,
            textvariable=self.delay_var,
            from_=0.0,
            to=10.0,
            increment=0.5,
            width=5,
            font=("Segoe UI", 10),
        ).pack(side="left", padx=(6, 2))
        ttk.Label(row_options, text="s", style="CardDim.TLabel").pack(side="left")

        ttk.Label(row_options, text="Retries:", style="CardBold.TLabel").pack(side="left", padx=(18, 0))
        ttk.Spinbox(
            row_options,
            textvariable=self.retries_var,
            from_=0,
            to=5,
            increment=1,
            width=4,
            font=("Segoe UI", 10),
        ).pack(side="left", padx=(6, 0))

        ttk.Checkbutton(
            row_options,
            text="Natural sort (file2 before file10)",
            variable=self.natural_var,
            style="Card.TCheckbutton",
        ).pack(side="left", padx=(18, 0))

        # Report: enable + format + directory
        row_report = ttk.Frame(settings, style="Card.TFrame")
        row_report.pack(fill="x", padx=15, pady=(4, 10))
        ttk.Checkbutton(
            row_report,
            text="📄 Report:",
            variable=self.report_var,
            style="Card.TCheckbutton",
            command=self._toggle_report_fields,
        ).pack(side="left")
        self.report_fmt_combo = ttk.Combobox(
            row_report,
            textvariable=self.report_fmt_var,
            values=["csv", "json"],
            state="readonly",
            width=5,
            font=("Segoe UI", 10),
        )
        self.report_fmt_combo.pack(side="left", padx=(6, 10))
        self.btn_report_dir = ttk.Button(
            row_report, text="Folder...", style="Secondary.TButton", command=self.browse_report_dir, width=10
        )
        self.btn_report_dir.pack(side="right", padx=(5, 0))
        self.entry_report_dir = ttk.Entry(
            row_report, textvariable=self.report_dir_var, font=("Segoe UI", 9)
        )
        self.entry_report_dir.pack(side="right", fill="x", expand=True, padx=(10, 4))
        ttk.Label(
            row_report, text="empty = next to the printed folder", style="CardDim.TLabel"
        ).pack(side="right", padx=(0, 4))

        # ── Buttons ────────────────────────────────────────────
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill="x", padx=20, pady=8)
        self.btn_print = ttk.Button(btn_frame, text="▶  Print", style="Accent.TButton", command=self.start_print)
        self.btn_print.pack(side="left")
        self.btn_preview = ttk.Button(
            btn_frame, text="🔍 Preview", style="Secondary.TButton", command=self.start_preview
        )
        self.btn_preview.pack(side="left", padx=(10, 0))
        self.btn_cancel = ttk.Button(
            btn_frame, text="✖ Cancel", style="Secondary.TButton", command=self.cancel_job, state="disabled"
        )
        self.btn_cancel.pack(side="left", padx=(10, 0))
        self.status_label = ttk.Label(btn_frame, text="", style="Status.TLabel")
        self.status_label.pack(side="right")

        # ── Progress ───────────────────────────────────────────
        self.progress_var = tk.DoubleVar(value=0)
        progress = ttk.Progressbar(self.root, variable=self.progress_var, maximum=100)
        progress.pack(fill="x", padx=20, pady=(0, 5))

        # ── Log ────────────────────────────────────────────────
        log_frame = ttk.Frame(self.root)
        log_frame.pack(fill="both", expand=True, padx=20, pady=(0, 15))
        self.log_text = tk.Text(
            log_frame,
            font=("Consolas", 9),
            bg="#1e1e1e",
            fg="#d4d4d4",
            insertbackground="#d4d4d4",
            selectbackground="#264f78",
            relief="flat",
            wrap="word",
            state="disabled",
            padx=10,
            pady=8,
        )
        self.log_text.pack(fill="both", expand=True, side="left")
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(fill="y", side="right")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self.log_text.tag_configure("header", foreground="#569cd6", font=("Consolas", 9, "bold"))
        self.log_text.tag_configure("group", foreground="#dcdcaa")
        self.log_text.tag_configure("success", foreground="#6a9955")
        self.log_text.tag_configure("error", foreground="#f44747")
        self.log_text.tag_configure("info", foreground="#9cdcfe")
        self.log_text.tag_configure("warn", foreground="#e5c07b")
        self.log_text.tag_configure("dim", foreground="#808080")

        self._toggle_report_fields()

    def _show_welcome(self):
        self.log_msg("PDF Folder Print", "header")
        self.log_msg("Structured PDF batch printing from folder trees.", "header")
        self.log_msg("", "")
        self.log_msg("1. Select a folder containing PDFs (with subfolders)", "info")
        self.log_msg("2. Verify the target printer", "info")
        self.log_msg("3. Click 'Preview' to inspect or 'Print' to start", "info")
        self.log_msg("", "")
        self.log_msg("Print order: root folder first, then subfolders (A→Z)", "dim")
        self.log_msg("Scaling:     Fit-to-Page with auto portrait/landscape", "dim")
        self.log_msg("Cancel:      stops after the current file (button or Esc)", "dim")

    # ── Printer status ─────────────────────────────────────────
    def refresh_printer_status(self):
        """Show live spooler status next to the printer dropdown."""
        name = self.printer_var.get()
        if not name:
            self.printer_status_label.configure(text="", foreground="#888")
            return
        problems = printer_status(name)
        if problems:
            self.printer_status_label.configure(
                text="⚠ " + ", ".join(problems), foreground=self.WARN
            )
        else:
            self.printer_status_label.configure(text="✓ ready", foreground=self.SUCCESS)

    # ── Queue-driven UI updates ────────────────────────────────
    def _poll_queue(self):
        try:
            while True:
                msg = self.ui_queue.get_nowait()
                kind = msg[0]
                if kind == "log":
                    _, text, tag = msg
                    self.log_msg(text, tag)
                elif kind == "status":
                    _, text, color = msg
                    self.set_status(text, color)
                elif kind == "progress":
                    _, value = msg
                    self.progress_var.set(value)
                elif kind == "running":
                    _, running = msg
                    self.set_running(running)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    # ── Helpers ────────────────────────────────────────────────
    def log_msg(self, text: str, tag: str = ""):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n", (tag,) if tag else ())
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def set_status(self, text: str, color: str = "#888"):
        self.status_label.configure(text=text, foreground=color)

    def set_running(self, running: bool):
        s = "disabled" if running else "normal"
        if running:
            self.is_running.set()
        else:
            self.is_running.clear()
            self.cancel_event.clear()
        self.btn_print.configure(state=s)
        self.btn_preview.configure(state=s)
        self.btn_browse.configure(state=s)
        self.btn_report_dir.configure(state=s)
        self.btn_cancel.configure(state="normal" if running else "disabled")
        self.combo_printer.configure(state=s if running else "readonly")
        self.report_fmt_combo.configure(state=s if running else "readonly")
        if not running:
            self.refresh_printer_status()

    def cancel_job(self):
        """Request cancellation – takes effect after the current file."""
        if not self.is_running.is_set():
            return
        self.cancel_event.set()
        self.btn_cancel.configure(state="disabled")
        self.set_status("Cancelling after current file...", self.ERROR)

    def _on_escape(self, _event=None):
        if self.is_running.is_set():
            self.cancel_job()

    def _on_close(self):
        save_settings(self._current_settings())
        self.root.destroy()

    def browse_folder(self):
        initial = self.folder_var.get() or ""
        valid_dir = initial and Path(initial).is_dir()
        folder = filedialog.askdirectory(title="Select folder", initialdir=valid_dir or "")
        if folder:
            self.folder_var.set(folder)

    def browse_report_dir(self):
        initial = self.report_dir_var.get() or self.folder_var.get() or ""
        valid_dir = initial and Path(initial).is_dir()
        folder = filedialog.askdirectory(title="Report folder", initialdir=valid_dir or "")
        if folder:
            self.report_dir_var.set(folder)

    def _toggle_report_fields(self):
        combo_state = "readonly" if self.report_var.get() else "disabled"
        entry_state = "normal" if self.report_var.get() else "disabled"
        self.report_fmt_combo.configure(state=combo_state)
        self.entry_report_dir.configure(state=entry_state)
        self.btn_report_dir.configure(state=entry_state)

    def validate(self) -> Path | None:
        folder = self.folder_var.get().strip()
        if not folder:
            messagebox.showwarning("No folder", "Please select a folder first.")
            return None
        p = Path(folder)
        if not p.is_dir():
            messagebox.showerror("Folder not found", f"Folder does not exist:\n{p}")
            return None
        return p

    # ── Actions ────────────────────────────────────────────────
    def _read_delay(self) -> float:
        try:
            return max(0.0, float(self.delay_var.get().replace(",", ".")))
        except ValueError:
            return 1.0

    def _read_retries(self) -> int:
        try:
            return max(0, min(5, int(self.retries_var.get())))
        except ValueError:
            return 0

    def _report_dir(self) -> Path | None:
        d = self.report_dir_var.get().strip()
        if d and Path(d).is_dir():
            return Path(d)
        return None

    def _current_settings(self) -> dict:
        return {
            "folder": self.folder_var.get().strip(),
            "printer": self.printer_var.get(),
            "delay": self.delay_var.get(),
            "retries": self.retries_var.get(),
            "natural_sort": bool(self.natural_var.get()),
            "report": bool(self.report_var.get()),
            "report_format": self.report_fmt_var.get(),
            "report_dir": self.report_dir_var.get().strip(),
        }

    def _start_job(self, dry_run: bool):
        root_path = self.validate()
        if not root_path:
            return
        printer = self.printer_var.get()

        # Fail fast on an unusable printer instead of queueing everything
        if not dry_run:
            problems = printer_status(printer)
            if problems and not messagebox.askyesno(
                "Printer problem?",
                f"The printer reports:\n\n{', '.join(problems)}\n\nStart printing anyway?",
            ):
                return
            if not messagebox.askyesno(
                "Start printing?",
                f"Print all PDFs in\n\n{root_path.name}\n\nto printer\n\n{printer}\n\n?",
            ):
                return

        save_settings(self._current_settings())
        self.log_clear()
        self.set_running(True)
        self.progress_var.set(0)
        threading.Thread(
            target=self._run_job,
            args=(
                root_path,
                printer,
                dry_run,
                self._read_delay(),
                self._read_retries(),
                self.natural_var.get(),
                bool(self.report_var.get()),
                self._report_dir(),
            ),
            daemon=True,
        ).start()

    def start_preview(self):
        self._start_job(dry_run=True)

    def start_print(self):
        self._start_job(dry_run=False)

    # ── Worker (runs in background thread) ─────────────────────
    def _run_job(
        self,
        root_path: Path,
        printer: str,
        dry_run: bool,
        delay: float,
        retries: int,
        natural: bool,
        report: bool,
        report_dir: Path | None,
    ):
        self.cancel_event.clear()
        report_obj = None
        try:
            if report:
                fmt = self.report_fmt_var.get()
                if fmt not in ("csv", "json"):
                    fmt = "csv"
                path = (
                    Path(report_dir) / f"pdf_print_report_{time.strftime('%Y%m%d-%H%M%S')}.{fmt}"
                    if report_dir
                    else default_report_path(root_path, fmt=fmt)
                )
                report_obj = PrintReport(path, fmt=fmt)
                self.ui_queue.put(("log", f"  Report:      {path}", "header"))
            self._job_inner(root_path, printer, dry_run, delay, retries, natural, report_obj)
        except Exception as e:  # unexpected failure – surface it in the log
            self.ui_queue.put(("log", f"\n⚠ Unexpected error: {e}", "error"))
        finally:
            if report_obj is not None:
                try:
                    report_obj.close()
                    self.ui_queue.put(("log", f"Report written: {report_obj.path}", "dim"))
                except Exception as e:
                    self.ui_queue.put(("log", f"⚠ Report could not be written: {e}", "error"))
            self.ui_queue.put(("running", False))

    def _job_inner(
        self,
        root_path: Path,
        printer: str,
        dry_run: bool,
        delay: float,
        retries: int,
        natural: bool,
        report: PrintReport | None,
    ):
        q = self.ui_queue
        pdfs = collect_pdfs(root_path, natural=natural)
        if not pdfs:
            q.put(("log", "No PDF files found.", "error"))
            q.put(("status", "No files", self.ERROR))
            return

        groups: dict[str, int] = {}
        for g, _ in pdfs:
            groups[g] = groups.get(g, 0) + 1

        mode = "PREVIEW" if dry_run else "PRINTING"
        q.put(("log", "═" * 60, "dim"))
        q.put(("log", f"  Folder:      {root_path.name}", "header"))
        q.put(("log", f"  Printer:     {printer}", "header"))
        q.put(("log", f"  Mode:        {mode}", "header"))
        q.put(("log", "  Scaling:     Fit-to-Page + auto orientation", "header"))
        q.put(("log", f"  Order:       {'natural' if natural else 'alphabetical'}", "header"))
        q.put(("log", f"  Files:       {len(pdfs)} in {len(groups)} groups", "header"))
        q.put(("log", "═" * 60, "dim"))
        for gname, count in groups.items():
            q.put(("log", f"  📁 {gname}: {count} file(s)", "info"))
        q.put(("log", "", ""))
        q.put(("status", f"0 / {len(pdfs)}...", self.ACCENT))

        current_group = None
        ok_count = 0
        fail_count = 0
        total_pages = 0
        cancelled = False

        for i, (group, pdf_path) in enumerate(pdfs):
            if self.cancel_event.is_set():
                cancelled = True
                q.put(("log", f"\n⚠ Cancelled – {i} of {len(pdfs)} files processed", "error"))
                q.put(("status", f"Cancelled at {i} / {len(pdfs)}", self.ERROR))
                break

            if group != current_group:
                current_group = group
                q.put(("log", "\n" + "─" * 40, "dim"))
                q.put(("log", f"📁 {group}", "group"))
                q.put(("log", "─" * 40, "dim"))

            if dry_run:
                n, w, h = get_pdf_info(pdf_path)
                if n > 0:
                    orient = "landscape" if w > h else "portrait"
                    info = f"{n} page{'s' if n != 1 else ''}, {w:.0f}×{h:.0f}mm ({orient})"
                    ok_count += 1
                    if report is not None:
                        report.add(pdf_path.name, group, n, "dry-run")
                else:
                    info = "? (unreadable)"
                    fail_count += 1
                    if report is not None:
                        report.add(pdf_path.name, group, 0, "unreadable", error="cannot open PDF")
                q.put(
                    (
                        "log",
                        f"  📄 {pdf_path.name}  →  {info}",
                        "info" if fail_count == 0 or n > 0 else "error",
                    )
                )
            else:
                q.put(("log", f"  🖨️ {pdf_path.name}...", "info"))
                attempt = 0
                t0 = time.perf_counter()
                while True:
                    ok, pages, err = print_pdf_file(pdf_path, printer)
                    dur_ms = int((time.perf_counter() - t0) * 1000)
                    if ok or attempt >= retries or err == "cancelled":
                        break
                    attempt += 1
                    q.put(("log", f"     ↻ retry {attempt}/{retries}", "warn"))
                    time.sleep(min(2**attempt, 5))

                if ok:
                    q.put(("log", f"     ✓ {pages} page{'s' if pages != 1 else ''}", "success"))
                    ok_count += 1
                    total_pages += pages
                    if report is not None:
                        report.add(pdf_path.name, group, pages, "ok", duration_ms=dur_ms)
                elif err == "cancelled":
                    cancelled = True
                    q.put(("log", "     ✗ cancelled", "error"))
                    if report is not None:
                        report.add(
                            pdf_path.name, group, pages, "cancelled", error="aborted by user", duration_ms=dur_ms
                        )
                    break
                else:
                    q.put(("log", f"     ✗ {err}", "error"))
                    fail_count += 1
                    if report is not None:
                        report.add(pdf_path.name, group, pages, "failed", error=err, duration_ms=dur_ms)
                if i < len(pdfs) - 1 and delay > 0:
                    time.sleep(delay)

            pct = (i + 1) / len(pdfs) * 100
            q.put(("progress", pct))
            q.put(("status", f"{i + 1} / {len(pdfs)}...", self.ACCENT))

        q.put(("log", "\n" + "═" * 60, "dim"))
        if cancelled:
            q.put(("log", f"  ⚠ ABORTED – {ok_count} files printed, {fail_count} failed", "error"))
        elif dry_run:
            q.put(("log", f"  🔍 PREVIEW: {ok_count} files found", "header"))
            q.put(("status", f"Preview: {ok_count} files", self.ACCENT))
        else:
            summary = f"✅ {ok_count} files ({total_pages} pages) printed"
            if fail_count:
                summary += f", {fail_count} failed"
            q.put(("log", f"  {summary}", "success" if fail_count == 0 else "error"))
            color = self.SUCCESS if fail_count == 0 else self.ERROR
            q.put(("status", f"Done – {ok_count} OK, {fail_count} failed", color))
        q.put(("log", "═" * 60, "dim"))
        q.put(("progress", 100))

    # ── Log window ─────────────────────────────────────────────
    def log_clear(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")


def gui_main() -> None:
    root = tk.Tk()
    PrintApp(root)
    root.mainloop()
