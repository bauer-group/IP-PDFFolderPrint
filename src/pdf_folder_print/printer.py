"""
Windows printer access: printer enumeration and GDI batch printing.

Everything here requires pywin32 and only works on Windows.
"""

from __future__ import annotations

import contextlib
from pathlib import Path

import fitz
import win32print
import win32ui
from PIL import Image, ImageWin

from .scaling import fit_page, render_zoom

FALLBACK_DPI = 300


def get_available_printers() -> list[str]:
    return [
        p[2]
        for p in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        )
    ]


def get_default_printer() -> str:
    return win32print.GetDefaultPrinter()


# win32 PRINTER_STATUS_* flags that clearly indicate the printer cannot
# take jobs right now. Flags not listed here are treated as "unknown" and
# never block, because they are vendor-dependent.
_PRINTER_STATUS_FLAGS = {
    0x00000001: "paused",
    0x00000002: "error",
    0x00000004: "pending deletion",
    0x00000008: "paper jam",
    0x00000010: "paper out",
    0x00000040: "paper problem",
    0x00000080: "offline",
    0x00000100: "door open",
    0x00000200: "user intervention required",
    0x00000400: "out of memory",
    0x00008000: "no toner",
}


def printer_status(printer_name: str) -> list[str]:
    """
    Query a printer's problem flags via the spooler.

    Returns a list of human-readable problem descriptions, e.g.
    ``["offline", "paper out"]``. Empty list means the printer reports no
    problems (or only vendor-specific flags we deliberately ignore).
    """
    try:
        handle = win32print.OpenPrinter(printer_name)
        try:
            info = win32print.GetPrinter(handle, 2)
        finally:
            win32print.ClosePrinter(handle)
    except Exception as e:
        return [f"status not queryable ({e})"]

    status = int(info.get("Status", 0) or 0)
    return [label for bit, label in _PRINTER_STATUS_FLAGS.items() if status & bit]


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
    printer_w_px = hdc.GetDeviceCaps(8)  # HORZRES
    printer_h_px = hdc.GetDeviceCaps(10)  # VERTRES

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
            zoom = render_zoom(pdf_w_pt, pdf_h_pt, printer_w_px, printer_h_px, dpi_override)
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # ── Fit-to-Page with best orientation ─────────────
            out_w, out_h, x, y, rotated = fit_page(
                img.width, img.height, printer_w_px, printer_h_px
            )
            if rotated:
                # Rotating gives better page utilization → rotate 90° CCW
                img = img.transpose(Image.Transpose.ROTATE_90)

            hdc.StartPage()
            dib = ImageWin.Dib(img)
            dib.draw(hdc.GetHandleOutput(), (x, y, x + out_w, y + out_h))
            hdc.EndPage()
            printed_page = page_num + 1

        hdc.EndDoc()

    except KeyboardInterrupt:
        # Ctrl+C / GUI cancel: abort the spool job, don't leave a half-printed doc
        with contextlib.suppress(Exception):
            doc.close()
        with contextlib.suppress(Exception):
            hdc.AbortDoc()
        return False, printed_page, "cancelled"

    except Exception as e:
        doc.close()
        with contextlib.suppress(Exception):
            hdc.AbortDoc()
        hdc.DeleteDC()
        return False, page_count, f"Print error page {printed_page + 1}: {e}"

    finally:
        with contextlib.suppress(Exception):
            hdc.DeleteDC()

    doc.close()
    return True, page_count, ""
