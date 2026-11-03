"""
GUI mode (tkinter) – the primary user interface.

All functionality is available from the window: folder selection, printer
(with live status), delay/retries, sort order, optional CSV/JSON report,
preview, printing and cancelling.

Layout philosophy:
  • everything is one aligned form grid (labels right-aligned, one column)
  • options are grouped and labeled, not scattered
  • progress/cancel elements only appear while a job runs
  • the log is the result area

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

    LABEL_COL = 130  # px width of the form's label column

    def __init__(self, root: tk.Tk):
        self.root = root
        self.ui_queue: queue.Queue = queue.Queue()
        self.is_running = threading.Event()
        self.cancel_event = threading.Event()

        self.root.title(f"PDF Folder Print {__version__}")
        self.root.geometry("900x760")
        self.root.minsize(760, 620)
        self.root.configure(bg=self.BG)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self._set_window_icon()
        self._setup_style()
        self._build_ui()
        self._show_welcome()
        self.refresh_printer_status()
        self._toggle_report_fields()
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
        style.configure("Cell.TFrame", background=self.FRAME_BG, relief="flat", borderwidth=0)
        style.configure("TLabel", background=self.BG, font=("Segoe UI", 10))
        style.configure("Header.TLabel", background=self.BG, font=("Segoe UI", 14, "bold"))
        style.configure("Card.TLabel", background=self.FRAME_BG, font=("Segoe UI", 10))
        style.configure(
            "CardLabel.TLabel", background=self.FRAME_BG, font=("Segoe UI", 10), foreground="#333333"
        )
        style.configure(
            "CardDim.TLabel", background=self.FRAME_BG, font=("Segoe UI", 9), foreground="#999999"
        )
        style.configure(
            "CardSection.TLabel",
            background=self.FRAME_BG,
            font=("Segoe UI", 9, "bold"),
            foreground="#999999",
        )
        style.configure("Card.TCheckbutton", background=self.FRAME_BG, font=("Segoe UI", 10))
        style.map("Card.TCheckbutton", background=[("active", self.FRAME_BG)])
        style.configure("Card.TRadiobutton", background=self.FRAME_BG, font=("Segoe UI", 10))
        style.map("Card.TRadiobutton", background=[("active", self.FRAME_BG)])
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

    # ── Small construction helpers ─────────────────────────────
    def _form_label(self, parent, text: str, row: int) -> ttk.Label:
        lbl = ttk.Label(parent, text=text, style="CardLabel.TLabel", anchor="e")
        lbl.grid(row=row, column=0, sticky="ne", padx=(14, 8), pady=6)
        return lbl

    def _section_label(self, parent, text: str, row: int) -> ttk.Label:
        lbl = ttk.Label(parent, text=text.upper(), style="CardSection.TLabel")
        lbl.grid(row=row, column=0, columnspan=3, sticky="w", padx=14, pady=(10, 2))
        return lbl

    def _build_ui(self):
        # ── State (pre-filled from last session) ───────────────
        saved = load_settings()
        self.folder_var = tk.StringVar(value=saved.get("folder", ""))
        self.printer_var = tk.StringVar()
        self.delay_var = tk.StringVar(value=str(saved.get("delay", "1.0")))
        self.retries_var = tk.StringVar(value=str(saved.get("retries", "0")))
        self.sort_var = tk.StringVar(value="natural" if saved.get("natural_sort") else "alpha")
        self.report_var = tk.BooleanVar(value=bool(saved.get("report", False)))
        self.report_fmt_var = tk.StringVar(value=str(saved.get("report_format", "csv")))
        self.report_dir_var = tk.StringVar(value=str(saved.get("report_dir", "")))

        printers = get_available_printers()
        default_printer = get_default_printer()
        self.printer_var.set(default_printer)
        if saved.get("printer") in printers:
            self.printer_var.set(saved["printer"])

        # ── Header ─────────────────────────────────────────────
        header = ttk.Frame(self.root)
        header.pack(fill="x", padx=24, pady=(16, 10))
        ttk.Label(header, text="PDF Folder Print", style="Header.TLabel").pack(side="left")
        ttk.Label(
            header, text=f"v{__version__} · BAUER GROUP", style="TLabel", foreground="#999999"
        ).pack(side="right", padx=(8, 0))
        ttk.Label(
            header, text="Strukturierter PDF-Stapeldruck", style="TLabel", foreground="#999999"
        ).pack(side="right", padx=(8, 0))
        header_line = tk.Frame(header, height=2, bg=self.ACCENT)
        header_line.pack(fill="x", side="bottom", pady=(6, 0))

        # ── Settings card (form grid) ──────────────────────────
        card_outer = ttk.Frame(self.root)
        card_outer.pack(fill="x", padx=24, pady=(4, 8))
        card = ttk.Frame(card_outer, style="Card.TFrame")
        card.pack(fill="x", ipady=6)
        card.columnconfigure(1, weight=1)

        row = 0

        # Section: Source & Destination
        self._section_label(card, "Quelle & Ziel", row)
        row += 1

        self._form_label(card, "Ordner", row)
        folder_cell = ttk.Frame(card, style="Cell.TFrame")
        folder_cell.grid(row=row, column=1, columnspan=2, sticky="ew", padx=(0, 14), pady=6)
        folder_cell.columnconfigure(0, weight=1)
        self.entry_folder = ttk.Entry(folder_cell, textvariable=self.folder_var, font=("Segoe UI", 10))
        self.entry_folder.grid(row=0, column=0, sticky="ew")
        self.entry_folder.bind("<Return>", lambda _e: self.start_preview())
        self.btn_browse = ttk.Button(
            folder_cell, text="Durchsuchen…", style="Secondary.TButton", command=self.browse_folder, width=12
        )
        self.btn_browse.grid(row=0, column=1, padx=(8, 0))
        row += 1

        self._form_label(card, "Drucker", row)
        printer_cell = ttk.Frame(card, style="Cell.TFrame")
        printer_cell.grid(row=row, column=1, columnspan=2, sticky="ew", padx=(0, 14), pady=6)
        printer_cell.columnconfigure(0, weight=1)
        self.combo_printer = ttk.Combobox(
            printer_cell,
            textvariable=self.printer_var,
            values=printers,
            state="readonly",
            font=("Segoe UI", 10),
        )
        self.combo_printer.grid(row=0, column=0, sticky="ew")
        self.combo_printer.bind("<<ComboboxSelected>>", lambda _e: self.refresh_printer_status())
        self.printer_status_label = ttk.Label(printer_cell, text="", style="Card.TLabel", anchor="e")
        self.printer_status_label.grid(row=0, column=1, padx=(8, 0))
        row += 1

        # Section: Options
        self._section_label(card, "Optionen", row)
        row += 1

        self._form_label(card, "Pause / Wiederholungen", row)
        timing_cell = ttk.Frame(card, style="Cell.TFrame")
        timing_cell.grid(row=row, column=1, columnspan=2, sticky="w", padx=(0, 14), pady=6)
        self.spin_delay = ttk.Spinbox(
            timing_cell, textvariable=self.delay_var, from_=0.0, to=10.0, increment=0.5, width=5, font=("Segoe UI", 10)
        )
        self.spin_delay.grid(row=0, column=0)
        ttk.Label(timing_cell, text="Sek. zwischen Jobs (Delay)", style="CardDim.TLabel").grid(
            row=0, column=1, padx=(6, 16)
        )
        self.spin_retries = ttk.Spinbox(
            timing_cell, textvariable=self.retries_var, from_=0, to=5, increment=1, width=5, font=("Segoe UI", 10)
        )
        self.spin_retries.grid(row=0, column=2)
        ttk.Label(timing_cell, text="Wiederholungen bei Fehler", style="CardDim.TLabel").grid(
            row=0, column=3, padx=(6, 0)
        )
        timing_cell.columnconfigure((0, 1, 2, 3), pad=4)
        row += 1

        self._form_label(card, "Sortierung", row)
        sort_cell = ttk.Frame(card, style="Cell.TFrame")
        sort_cell.grid(row=row, column=1, columnspan=2, sticky="w", padx=(0, 14), pady=6)
        self._sort_radios = [
            ttk.Radiobutton(
                sort_cell, text="Alphabetisch (A→Z)", variable=self.sort_var, value="alpha", style="Card.TRadiobutton"
            ),
            ttk.Radiobutton(
                sort_cell,
                text="Natürlich (file2 vor file10)",
                variable=self.sort_var,
                value="natural",
                style="Card.TRadiobutton",
            ),
        ]
        self._sort_radios[0].pack(side="left")
        self._sort_radios[1].pack(side="left", padx=(14, 0))
        row += 1

        # Section: Report
        self._section_label(card, "Report", row)
        row += 1

        self._form_label(card, "Report schreiben", row)
        report_cell = ttk.Frame(card, style="Cell.TFrame")
        report_cell.grid(row=row, column=1, columnspan=2, sticky="ew", padx=(0, 14), pady=6)
        report_cell.columnconfigure(2, weight=1)
        report_cb = ttk.Checkbutton(
            report_cell,
            text="CSV-/JSON-Zusammenfassung pro Datei",
            variable=self.report_var,
            style="Card.TCheckbutton",
            command=self._toggle_report_fields,
        )
        report_cb.grid(row=0, column=0)
        self.report_fmt_combo = ttk.Combobox(
            report_cell,
            textvariable=self.report_fmt_var,
            values=["csv", "json"],
            state="disabled",
            width=6,
            font=("Segoe UI", 10),
        )
        self.report_fmt_combo.grid(row=0, column=1, padx=(14, 10))
        self.entry_report_dir = ttk.Entry(report_cell, textvariable=self.report_dir_var, font=("Segoe UI", 10))
        self.entry_report_dir.grid(row=0, column=2, sticky="ew")
        self.btn_report_dir = ttk.Button(
            report_cell, text="Ordner…", style="Secondary.TButton", command=self.browse_report_dir, width=9
        )
        self.btn_report_dir.grid(row=0, column=3, padx=(8, 0))
        self.report_hint = ttk.Label(report_cell, text="", style="CardDim.TLabel")
        self.report_hint.grid(row=1, column=0, columnspan=4, sticky="w")
        row += 1

        # ── Action bar ─────────────────────────────────────────
        action = ttk.Frame(self.root)
        action.pack(fill="x", padx=24, pady=(8, 4))
        self.btn_print = ttk.Button(action, text="▶  Drucken", style="Accent.TButton", command=self.start_print)
        self.btn_print.pack(side="left")
        self.btn_preview = ttk.Button(
            action, text="🔍 Vorschau", style="Secondary.TButton", command=self.start_preview
        )
        self.btn_preview.pack(side="left", padx=(10, 0))
        self.btn_cancel = ttk.Button(
            action, text="✖ Abbrechen", style="Secondary.TButton", command=self.cancel_job, state="disabled"
        )
        self.btn_cancel.pack(side="left", padx=(10, 0))
        self.status_label = ttk.Label(action, text="Bereit.", style="Status.TLabel")
        self.status_label.pack(side="right")

        # ── Progress (hidden while idle) ───────────────────────
        self.progress_frame = ttk.Frame(self.root)
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(self.progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", padx=24, pady=(2, 0))
        # hidden until a job starts (set_running shows it)

        # ── Log ────────────────────────────────────────────────
        log_frame = ttk.Frame(self.root)
        log_frame.pack(fill="both", expand=True, padx=24, pady=(4, 16))
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
            padx=12,
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

    def _show_welcome(self):
        self.log_msg("PDF Folder Print", "header")
        self.log_msg("Strukturierter PDF-Stapeldruck aus Ordnerbäumen.", "header")
        self.log_msg("", "")
        self.log_msg("1. Ordner mit PDFs auswählen (inkl. Unterordner)", "info")
        self.log_msg("2. Ziel-Drucker prüfen", "info")
        self.log_msg("3. 'Vorschau' zur Kontrolle oder 'Drucken' zum Start", "info")
        self.log_msg("", "")
        self.log_msg("Reihenfolge: Stammordner zuerst, dann Unterordner (A→Z)", "dim")
        self.log_msg("Skalierung:  Fit-to-Page mit automatischer Ausrichtung", "dim")
        self.log_msg("Abbrechen:   stoppt nach der aktuellen Datei (Button oder Esc)", "dim")

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
            self.printer_status_label.configure(text="✓ bereit", foreground=self.SUCCESS)

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
            self.progress_frame.pack(fill="x", before=self.log_text.master)
            self.progress_var.set(0)
        else:
            self.is_running.clear()
            self.cancel_event.clear()
            self.progress_frame.pack_forget()
        self.btn_print.configure(state=s)
        self.btn_preview.configure(state=s)
        self.btn_browse.configure(state=s)
        self.btn_report_dir.configure(state=s)
        self.entry_report_dir.configure(state=s if not self.report_var.get() else "normal")
        self.btn_cancel.configure(state="normal" if running else "disabled")
        self.combo_printer.configure(state=s if running else "readonly")
        self.report_fmt_combo.configure(
            state=(s if running else "readonly") if self.report_var.get() else "disabled"
        )
        self.spin_delay.configure(state=s if running else "normal")
        self.spin_retries.configure(state=s if running else "normal")
        for radio in self._sort_radios:
            radio.configure(state=s)
        if not running:
            self.refresh_printer_status()

    def cancel_job(self):
        """Request cancellation – takes effect after the current file."""
        if not self.is_running.is_set():
            return
        self.cancel_event.set()
        self.btn_cancel.configure(state="disabled")
        self.set_status("Abbruch nach aktueller Datei…", self.ERROR)

    def _on_escape(self, _event=None):
        if self.is_running.is_set():
            self.cancel_job()

    def _on_close(self):
        save_settings(self._current_settings())
        self.root.destroy()

    def browse_folder(self):
        initial = self.folder_var.get() or ""
        valid_dir = initial and Path(initial).is_dir()
        folder = filedialog.askdirectory(title="Ordner auswählen", initialdir=valid_dir or "")
        if folder:
            self.folder_var.set(folder)

    def browse_report_dir(self):
        initial = self.report_dir_var.get() or self.folder_var.get() or ""
        valid_dir = initial and Path(initial).is_dir()
        folder = filedialog.askdirectory(title="Report-Ordner", initialdir=valid_dir or "")
        if folder:
            self.report_dir_var.set(folder)

    def _toggle_report_fields(self):
        if self.report_var.get():
            self.report_fmt_combo.configure(state="readonly")
            self.entry_report_dir.configure(state="normal")
            self.btn_report_dir.configure(state="normal")
            self.report_hint.configure(
                text="leer = neben dem gedruckten Ordner", foreground="#999999"
            )
        else:
            self.report_fmt_combo.configure(state="disabled")
            self.entry_report_dir.configure(state="disabled")
            self.btn_report_dir.configure(state="disabled")
            self.report_hint.configure(text="")

    def validate(self) -> Path | None:
        folder = self.folder_var.get().strip()
        if not folder:
            messagebox.showwarning("Kein Ordner", "Bitte zuerst einen Ordner auswählen.")
            return None
        p = Path(folder)
        if not p.is_dir():
            messagebox.showerror("Ordner nicht gefunden", f"Der Ordner existiert nicht:\n{p}")
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
            "natural_sort": self.sort_var.get() == "natural",
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
                "Drucker-Problem?",
                f"Der Drucker meldet:\n\n{', '.join(problems)}\n\nTrotzdem drucken?",
            ):
                return
            if not messagebox.askyesno(
                "Druck starten?",
                f"Alle PDFs aus\n\n{root_path.name}\n\nauf Drucker\n\n{printer}\n\ndrucken?",
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
                self.sort_var.get() == "natural",
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
            self.ui_queue.put(("log", f"\n⚠ Unerwarteter Fehler: {e}", "error"))
        finally:
            if report_obj is not None:
                try:
                    report_obj.close()
                    self.ui_queue.put(("log", f"Report geschrieben: {report_obj.path}", "dim"))
                except Exception as e:
                    self.ui_queue.put(("log", f"⚠ Report konnte nicht geschrieben werden: {e}", "error"))
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
            q.put(("log", "Keine PDF-Dateien gefunden.", "error"))
            q.put(("status", "Keine Dateien", self.ERROR))
            return

        groups: dict[str, int] = {}
        for g, _ in pdfs:
            groups[g] = groups.get(g, 0) + 1

        mode = "VORSCHAU" if dry_run else "DRUCKEN"
        q.put(("log", "═" * 60, "dim"))
        q.put(("log", f"  Ordner:      {root_path.name}", "header"))
        q.put(("log", f"  Drucker:     {printer}", "header"))
        q.put(("log", f"  Modus:       {mode}", "header"))
        q.put(("log", "  Skalierung:  Fit-to-Page + automatische Ausrichtung", "header"))
        q.put(("log", f"  Reihenfolge: {'natürlich' if natural else 'alphabetisch'}", "header"))
        q.put(("log", f"  Dateien:     {len(pdfs)} in {len(groups)} Gruppen", "header"))
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
                q.put(("log", f"\n⚠ Abgebrochen – {i} von {len(pdfs)} Dateien verarbeitet", "error"))
                q.put(("status", f"Abgebrochen bei {i} / {len(pdfs)}", self.ERROR))
                break

            if group != current_group:
                current_group = group
                q.put(("log", "\n" + "─" * 40, "dim"))
                q.put(("log", f"📁 {group}", "group"))
                q.put(("log", "─" * 40, "dim"))

            if dry_run:
                n, w, h = get_pdf_info(pdf_path)
                if n > 0:
                    orient = "quer" if w > h else "hoch"
                    info = f"{n} Seite{'n' if n != 1 else ''}, {w:.0f}×{h:.0f}mm ({orient})"
                    ok_count += 1
                    if report is not None:
                        report.add(pdf_path.name, group, n, "dry-run")
                else:
                    info = "? (unlesbar)"
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
                    q.put(("log", f"     ↻ Wiederholung {attempt}/{retries}", "warn"))
                    time.sleep(min(2**attempt, 5))

                if ok:
                    q.put(
                        (
                            "log",
                            f"     ✓ {pages} Seite{'n' if pages != 1 else ''}",
                            "success",
                        )
                    )
                    ok_count += 1
                    total_pages += pages
                    if report is not None:
                        report.add(pdf_path.name, group, pages, "ok", duration_ms=dur_ms)
                elif err == "cancelled":
                    cancelled = True
                    q.put(("log", "     ✗ abgebrochen", "error"))
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
            q.put(("status", f"{i + 1} / {len(pdfs)}…", self.ACCENT))

        q.put(("log", "\n" + "═" * 60, "dim"))
        if cancelled:
            q.put(
                (
                    "log",
                    f"  ⚠ ABGEBROCHEN – {ok_count} Dateien gedruckt, {fail_count} fehlgeschlagen",
                    "error",
                )
            )
        elif dry_run:
            q.put(("log", f"  🔍 VORSCHAU: {ok_count} Dateien gefunden", "header"))
            q.put(("status", f"Vorschau: {ok_count} Dateien", self.ACCENT))
        else:
            summary = f"✅ {ok_count} Dateien ({total_pages} Seiten) gedruckt"
            if fail_count:
                summary += f", {fail_count} fehlgeschlagen"
            q.put(("log", f"  {summary}", "success" if fail_count == 0 else "error"))
            color = self.SUCCESS if fail_count == 0 else self.ERROR
            q.put(("status", f"Fertig – {ok_count} OK, {fail_count} fehlgeschlagen", color))
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

