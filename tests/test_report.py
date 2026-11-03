"""Tests for the print report module."""

import csv
import json

from pdf_folder_print.report import PrintReport, default_report_path


def test_csv_report_written_incrementally(tmp_path):
    path = tmp_path / "report.csv"
    with PrintReport(path, fmt="csv") as report:
        report.add("a.pdf", "(Root)", 3, "ok", duration_ms=120)
        # file must already contain the row (crash-safe)
        with path.open(encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
        assert len(rows) == 1
        assert rows[0]["file"] == "a.pdf"
        assert rows[0]["status"] == "ok"

    report.add("b.pdf", "Invoices", 1, "failed", error="boom", duration_ms=50)  # after close -> no-op for csv
    with path.open(encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    assert len(rows) == 1


def test_csv_report_multiple_rows(tmp_path):
    path = tmp_path / "report.csv"
    with PrintReport(path) as report:
        report.add("a.pdf", "(Root)", 3, "ok")
        report.add("b.pdf", "Sub", 2, "failed", error="printer offline")

    with path.open(encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    assert [r["status"] for r in rows] == ["ok", "failed"]
    assert rows[1]["error"] == "printer offline"


def test_json_report_written_on_close(tmp_path):
    path = tmp_path / "report.json"
    report = PrintReport(path, fmt="json")
    report.add("a.pdf", "(Root)", 3, "ok", duration_ms=99)
    assert not path.exists()  # JSON is buffered until close
    report.close()

    data = json.loads(path.read_text(encoding="utf-8"))
    assert data["tool"] == "pdf-folder-print"
    assert len(data["files"]) == 1
    assert data["files"][0]["pages"] == 3
    assert data["files"][0]["duration_ms"] == 99


def test_invalid_format_rejected(tmp_path):
    import pytest

    with pytest.raises(ValueError):
        PrintReport(tmp_path / "x.xml", fmt="xml")


def test_default_report_path_next_to_root(tmp_path):
    folder = tmp_path / "March"
    folder.mkdir()
    p = default_report_path(folder, fmt="csv")
    assert p.parent == tmp_path
    assert p.suffix == ".csv"
    assert p.name.startswith("pdf_print_report_")
