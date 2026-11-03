"""
Pure scaling / orientation math for Fit-to-Page printing.

No PyMuPDF, no Windows dependencies – fully unit-testable.
"""

from __future__ import annotations


def render_zoom(
    pdf_w_pt: float,
    pdf_h_pt: float,
    printer_w_px: int,
    printer_h_px: int,
    dpi_override: int | None = None,
) -> float:
    """
    PyMuPDF zoom factor for rasterizing a PDF page.

    - With ``dpi_override``: render at that fixed DPI.
    - Otherwise: render at the printer's native resolution, using the zoom
      that covers the *longer* axis so the bitmap is large enough for either
      orientation.
    """
    if dpi_override:
        return dpi_override / 72.0
    zoom_x = printer_w_px / pdf_w_pt
    zoom_y = printer_h_px / pdf_h_pt
    return max(zoom_x, zoom_y)


def fit_page(
    img_w: int,
    img_h: int,
    printer_w_px: int,
    printer_h_px: int,
) -> tuple[int, int, int, int, bool]:
    """
    Fit-to-Page geometry (mirrors Adobe Reader's "Fit" option).

    Compares normal vs. 90°-rotated placement and keeps whichever yields a
    higher scale factor (= better page fill).

    Returns:
        (out_w, out_h, x, y, rotated) – target bitmap size, top-left offset
        centered on the printable area, and whether the bitmap must be
        rotated 90° before drawing.
    """
    scale_normal = min(printer_w_px / img_w, printer_h_px / img_h)
    scale_rotated = min(printer_w_px / img_h, printer_h_px / img_w)

    rotated = scale_rotated > scale_normal
    scale = scale_rotated if rotated else scale_normal

    # After a 90° rotation the bitmap's width/height are swapped
    eff_w, eff_h = (img_h, img_w) if rotated else (img_w, img_h)
    out_w = int(eff_w * scale)
    out_h = int(eff_h * scale)
    x = (printer_w_px - out_w) // 2
    y = (printer_h_px - out_h) // 2
    return out_w, out_h, x, y, rotated
