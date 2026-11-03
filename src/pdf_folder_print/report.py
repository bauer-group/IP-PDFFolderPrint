"""
Optional print report (CSV or JSON).

One row per processed file: timestamp, file, group, pages, status, error,
duration. CSV is written incrementally (so a killed process still leaves a
usable report); JSON is written on close.
"""

from __future__ import annotations

import csv
import json
import time
from pathlib import Path

FIELDS = ["timestamp", "file", "group", "pages", "status", "error", "duration_ms"]

SUPPORTED_FORMATS = ("csv", "json")


def default_report_path(root: Path, fmt: str = "csv") -> Path:
    """Next to the printed folder: pdf_print_report_<timestamp>.<fmt>."""
    stamp = time.strftime("%Y%m%d-%H%M%S")
    return Path(root).parent / f"pdf_print_report_{stamp}.{fmt}"


class PrintReport:
    """Collects one row per printed file and persists it as CSV or JSON."""

    def __init__(self, path: Path, fmt: str = "csv"):
        if fmt not in SUPPORTED_FORMATS:
            raise ValueError(f"Unsupported report format: {fmt!r} (use {', '.join(SUPPORTED_FORMATS)})")
        self.path = Path(path)
        self.fmt = fmt
        self.rows: list[dict] = []
        self._file = None
        self._writer = None
        if fmt == "csv":
            self.path.parent.mkdir(parents=True, exist_ok=True)
            # utf-8-sig so Excel opens umlauts correctly
            self._file = self.path.open("w", newline="", encoding="utf-8-sig")
            self._writer = csv.DictWriter(self._file, fieldnames=FIELDS)
            self._writer.writeheader()

    def add(
        self,
        file: str,
        group: str,
        pages: int,
        status: str,
        error: str = "",
        duration_ms: int = 0,
    ) -> None:
        row = {
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            "file": file,
            "group": group,
            "pages": pages,
            "status": status,
            "error": error,
            "duration_ms": duration_ms,
        }
        self.rows.append(row)
        if self._writer:
            self._writer.writerow(row)
            self._file.flush()

    def close(self) -> None:
        if self._file is not None:
            self._file.close()
            self._file = None
            self._writer = None
        elif self.fmt == "json":
            self.path.parent.mkdir(parents=True, exist_ok=True)
            payload = {
                "tool": "pdf-folder-print",
                "generated": time.strftime("%Y-%m-%d %H:%M:%S"),
                "files": self.rows,
            }
            self.path.write_text(
                json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8"
            )

    def __enter__(self) -> PrintReport:
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()
