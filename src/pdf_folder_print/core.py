"""
Core file collection and PDF info helpers.

This module is intentionally free of Windows dependencies (only PyMuPDF)
so it can be unit-tested without pywin32.
"""

from __future__ import annotations

import re
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:  # pragma: no cover - startup dependency check handles this
    fitz = None

PDF_EXTENSIONS = {".pdf"}


def natural_sort_key(name: str) -> tuple:
    """
    Sort key that treats digit runs numerically ("file2.pdf" < "file10.pdf").

    Non-digit parts are compared case-insensitively, so the result is a
    drop-in replacement for plain alphabetical sorting.
    """
    parts = re.split(r"(\d+)", name)
    return tuple(int(p) if p.isdigit() else p.lower() for p in parts)


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
