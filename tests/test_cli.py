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


class _RecordingConsole:
    """Stands in for rich Console.input to detect confirmation prompts."""

    def __init__(self):
        self.called = 0

    def __call__(self, *args, **kwargs):
        self.called += 1
        return "n"


def test_yes_skips_confirmation(fake_printers, monkeypatch, tmp_path):
    import fitz

    doc = fitz.open()
    doc.new_page()
    doc.save(tmp_path / "a.pdf")
    doc.close()

    calls = []
    monkeypatch.setattr(cli, "print_pdf_file", lambda *a, **kw: calls.append(a) or (True, 1, ""))
    prompt = _RecordingConsole()
    # patch Console.input on the class used by cli
    monkeypatch.setattr(cli.Console, "input", prompt)

    with pytest.raises(SystemExit) as exc:
        run_cli(monkeypatch, "--yes", "--delay", "0", str(tmp_path))
    assert exc.value.code == 0
    assert prompt.called == 0
    assert len(calls) == 1


def test_confirmation_prompt_asked_without_yes(fake_printers, monkeypatch, tmp_path):
    import fitz

    doc = fitz.open()
    doc.new_page()
    doc.save(tmp_path / "a.pdf")
    doc.close()

    prompt = _RecordingConsole()
    monkeypatch.setattr(cli.Console, "input", prompt)

    with pytest.raises(SystemExit) as exc:
        run_cli(monkeypatch, "--delay", "0", str(tmp_path))
    assert exc.value.code == 0
    assert prompt.called == 1  # answered "n" → cancelled before printing


def test_dry_run_unreadable_pdf_exits_1(fake_printers, monkeypatch, tmp_path):
    (tmp_path / "broken.pdf").write_bytes(b"not a pdf at all")

    with pytest.raises(SystemExit) as exc:
        run_cli(monkeypatch, "--dry-run", str(tmp_path))
    assert exc.value.code == 1


def test_dry_run_valid_pdf_exits_0(fake_printers, monkeypatch, tmp_path):
    import fitz

    doc = fitz.open()
    doc.new_page()
    doc.save(tmp_path / "good.pdf")
    doc.close()

    with pytest.raises(SystemExit) as exc:
        run_cli(monkeypatch, "--dry-run", str(tmp_path))
    assert exc.value.code == 0
