"""
GUI mode (tkinter).

The window logic lives in the ``PrintApp`` class. Worker threads never touch
tkinter directly – they push ``(kind, ...)`` messages into a ``queue.Queue``
that the UI polls via ``root.after``.
"""

from __future__ import annotations

import queue
import sys
import threading
import time
from pathlib import Path

from .core import collect_pdfs, get_pdf_info
from .printer import get_available_printers, get_default_printer, print_pdf_file


def _asset_path(name: str) -> Path:
    """Locate bundled assets (works from source checkout and frozen EXE)."""
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent)) / "assets" / name
    # source layout: src/pdf_folder_print/gui.py → project root/assets
    return Path(__file__).resolve().parents[2] / "assets" / name


def gui_main() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    class PrintApp:
        """Main application window."""

        BG = "#f5f5f5"
        FRAME_BG = "#ffffff"
        ACCENT = "#FF8500"  # BAUER GROUP brand orange-500
        SUCCESS = "#16a34a"
        ERROR = "#dc2626"

        def __init__(self, root: tk.Tk):
            self.root = root
            self.ui_queue: queue.Queue = queue.Queue()
            self.is_running = threading.Event()

            self.root.title("PDF Folder Print")
            self.root.geometry("820x680")
            self.root.minsize(700, 550)
            self.root.configure(bg=self.BG)

            self._set_window_icon()
            self._setup_style()
            self._build_ui()
            self._show_welcome()
            self.root.after(100, self._poll_queue)

        # ── Setup ──────────────────────────────────────────────
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
            style.configure("CardBold.TLabel", background=self.FRAME_BG, font=("Segoe UI", 10, "bold"))
            style.configure("Status.TLabel", background=self.BG, font=("Segoe UI", 10))
            style.configure("Accent.TButton", font=("Segoe UI", 11, "bold"), padding=(20, 10))
            style.map(
                "Accent.TButton",
                background=[("active", "#EA6D00"), ("disabled", "#FFE0BF")],  # orange-600 / dimmed
                foreground=[("disabled", "#AAAAAA")],
            )
            style.configure("Accent.TButton", background=self.ACCENT, foreground="white")
            style.configure("Secondary.TButton", font=("Segoe UI", 10), padding=(12, 6))
            style.configure(
                "Horizontal.TProgressbar",
                troughcolor=self.FRAME_BG,
                background=self.ACCENT,
                lightcolor=self.ACCENT,
                darkcolor=self.ACCENT,
            )

        def _build_ui(self):
            # State
            self.folder_var = tk.StringVar()
            self.printer_var = tk.StringVar()
            self.printer_var.set(get_default_printer())

            printers = get_available_printers()

            # Header
            header_frame = ttk.Frame(self.root)
            header_frame.pack(fill="x", padx=20, pady=(15, 5))
            ttk.Label(header_frame, text="🖨️ PDF Folder Print", style="Header.TLabel").pack(side="left")
            ttk.Label(header_frame, text="Structured PDF Batch Printing", style="TLabel", foreground="#888").pack(
                side="right"
            )

            # Settings card
            settings_outer = ttk.Frame(self.root)
            settings_outer.pack(fill="x", padx=20, pady=(10, 5))
            settings = ttk.Frame(settings_outer, style="Card.TFrame")
            settings.pack(fill="x", ipady=8)

            row_folder = ttk.Frame(settings, style="Card.TFrame")
            row_folder.pack(fill="x", padx=15, pady=(10, 5))
            ttk.Label(row_folder, text="📁 Folder:", style="CardBold.TLabel").pack(side="left")
            self.btn_browse = ttk.Button(
                row_folder, text="Browse...", style="Secondary.TButton", command=self.browse_folder
            )
            self.btn_browse.pack(side="right", padx=(5, 0))
            entry_folder = ttk.Entry(row_folder, textvariable=self.folder_var, font=("Segoe UI", 10))
            entry_folder.pack(side="right", fill="x", expand=True, padx=(10, 5))

            row_printer = ttk.Frame(settings, style="Card.TFrame")
            row_printer.pack(fill="x", padx=15, pady=(5, 10))
            ttk.Label(row_printer, text="🖨️ Printer:", style="CardBold.TLabel").pack(side="left")
            self.combo_printer = ttk.Combobox(
                row_printer,
                textvariable=self.printer_var,
                values=printers,
                state="readonly",
                font=("Segoe UI", 10),
                width=50,
            )
            self.combo_printer.pack(side="right", fill="x", expand=True, padx=(10, 0))

            # Delay between print jobs
            row_delay = ttk.Frame(settings, style="Card.TFrame")
            row_delay.pack(fill="x", padx=15, pady=(5, 10))
            ttk.Label(row_delay, text="⏱️ Delay:", style="CardBold.TLabel").pack(side="left")
            self.delay_var = tk.StringVar(value="1.0")
            delay_spin = ttk.Spinbox(
                row_delay,
                textvariable=self.delay_var,
                from_=0.0,
                to=10.0,
                increment=0.5,
                width=6,
                font=("Segoe UI", 10),
            )
            delay_spin.pack(side="right", padx=(10, 0))
            ttk.Label(
                row_delay,
                text="seconds between print jobs",
                style="Card.TLabel",
                foreground="#888",
            ).pack(side="right")

            # Buttons
            btn_frame = ttk.Frame(self.root)
            btn_frame.pack(fill="x", padx=20, pady=8)
            self.btn_print = ttk.Button(btn_frame, text="▶  Print", style="Accent.TButton", command=self.start_print)
            self.btn_print.pack(side="left")
            self.btn_preview = ttk.Button(
                btn_frame, text="🔍 Preview", style="Secondary.TButton", command=self.start_preview
            )
            self.btn_preview.pack(side="left", padx=(10, 0))
            self.status_label = ttk.Label(btn_frame, text="", style="Status.TLabel")
            self.status_label.pack(side="right")

            # Progress
            self.progress_var = tk.DoubleVar(value=0)
            progress = ttk.Progressbar(self.root, variable=self.progress_var, maximum=100)
            progress.pack(fill="x", padx=20, pady=(0, 5))

            # Log
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
            self.log_text.tag_configure("dim", foreground="#808080")

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

        # ── Queue-driven UI updates ────────────────────────────
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

        # ── Helpers ────────────────────────────────────────────
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
            self.btn_print.configure(state=s)
            self.btn_preview.configure(state=s)
            self.btn_browse.configure(state=s)
            self.combo_printer.configure(state=s if running else "readonly")

        def browse_folder(self):

            initial = self.folder_var.get() or ""
            valid_dir = initial and Path(initial).is_dir()
            folder = filedialog.askdirectory(title="Select folder", initialdir=valid_dir or "")
            if folder:
                self.folder_var.set(folder)

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

        # ── Actions ────────────────────────────────────────────
        def _read_delay(self) -> float:
            try:
                return max(0.0, float(self.delay_var.get().replace(",", ".")))
            except ValueError:
                return 1.0

        def start_preview(self):
            root_path = self.validate()
            if not root_path:
                return
            self.log_clear()
            self.set_running(True)
            self.progress_var.set(0)
            threading.Thread(
                target=self._run_job,
                args=(root_path, self.printer_var.get(), True, self._read_delay()),
                daemon=True,
            ).start()

        def start_print(self):

            root_path = self.validate()
            if not root_path:
                return
            printer = self.printer_var.get()
            if not messagebox.askyesno(
                "Start printing?",
                f"Print all PDFs in\n\n{root_path.name}\n\nto printer\n\n{printer}\n\n?",
            ):
                return
            self.log_clear()
            self.set_running(True)
            self.progress_var.set(0)
            threading.Thread(
                target=self._run_job,
                args=(root_path, printer, False, self._read_delay()),
                daemon=True,
            ).start()

        # ── Worker (runs in background thread) ─────────────────
        def _run_job(self, root_path: Path, printer: str, dry_run: bool, delay: float):
            try:
                self._job_inner(root_path, printer, dry_run, delay)
            except Exception as e:  # unexpected failure – surface it in the log
                self.ui_queue.put(("log", f"\n⚠ Unexpected error: {e}", "error"))
            finally:
                self.ui_queue.put(("running", False))

        def _job_inner(self, root_path: Path, printer: str, dry_run: bool, delay: float):
            q = self.ui_queue
            pdfs = collect_pdfs(root_path)
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

            for i, (group, pdf_path) in enumerate(pdfs):
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
                    else:
                        info = "? (unreadable)"
                        fail_count += 1
                    q.put(("log", f"  📄 {pdf_path.name}  →  {info}", "info" if fail_count == 0 or n > 0 else "error"))
                else:
                    q.put(("log", f"  🖨️ {pdf_path.name}...", "info"))
                    ok, pages, err = print_pdf_file(pdf_path, printer)
                    if ok:
                        q.put(("log", f"     ✓ {pages} page{'s' if pages != 1 else ''}", "success"))
                        ok_count += 1
                        total_pages += pages
                    else:
                        q.put(("log", f"     ✗ {err}", "error"))
                        fail_count += 1
                    if i < len(pdfs) - 1 and delay > 0:
                        time.sleep(delay)

                pct = (i + 1) / len(pdfs) * 100
                q.put(("progress", pct))
                q.put(("status", f"{i + 1} / {len(pdfs)}...", self.ACCENT))

            q.put(("log", "\n" + "═" * 60, "dim"))
            if dry_run:
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

        # ── Log window ─────────────────────────────────────────
        def log_clear(self):
            self.log_text.configure(state="normal")
            self.log_text.delete("1.0", "end")
            self.log_text.configure(state="disabled")

    root = tk.Tk()
    PrintApp(root)
    root.mainloop()
