"""Tests for the CLI: version, printer listing, and error exit codes.

win32 printer access is monkeypatched so the tests run without hardware.
"""

import sys

import pytest

from pdf_folder_print import __version__, cli


@pytest.fixture
def fake_printers(monkeypatch):
    monkeypatch.setattr(cli, "get_available_printers", lambda: ["HP LaserJet", "Canon MX920"])
    monkeypatch.setattr(cli, "get_default_printer", lambda: "HP LaserJet")


def run_cli(monkeypatch, *argv):
    monkeypatch.setattr(sys, "argv", ["pdf_folder_print", *argv])
    cli.cli_main()


def test_version(capsys, monkeypatch):
    with pytest.raises(SystemExit) as exc:
        run_cli(monkeypatch, "--version")
    assert exc.value.code == 0
    assert f"pdf_folder_print {__version__}" in capsys.readouterr().out


def test_list_printers(fake_printers, capsys, monkeypatch):
    with pytest.raises(SystemExit) as exc:
        run_cli(monkeypatch, "--list-printers")
    assert exc.value.code == 0
    out = capsys.readouterr().out
    assert "HP LaserJet" in out
    assert "Canon MX920" in out


def test_missing_folder_arg_exits_1(fake_printers, monkeypatch):
    with pytest.raises(SystemExit) as exc:
        run_cli(monkeypatch)
    assert exc.value.code == 1


def test_nonexistent_folder_exits_1(fake_printers, monkeypatch):
    with pytest.raises(SystemExit) as exc:
        run_cli(monkeypatch, r"C:\definitely\not\a\real\folder")
    assert exc.value.code == 1


def test_unknown_printer_exits_1(fake_printers, monkeypatch, tmp_path):
    with pytest.raises(SystemExit) as exc:
        run_cli(monkeypatch, "--printer", "Ghost Printer", str(tmp_path))
    assert exc.value.code == 1


def test_empty_folder_exits_0(fake_printers, monkeypatch, tmp_path):
    with pytest.raises(SystemExit) as exc:
        run_cli(monkeypatch, "--dry-run", str(tmp_path))
    assert exc.value.code == 0
