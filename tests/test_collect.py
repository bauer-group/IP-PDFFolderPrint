"""Tests for collect_pdfs: deterministic order, filtering, single-level traversal."""

from pathlib import Path

from pdf_folder_print.core import collect_pdfs


def make_tree(base: Path, *rel_paths: str) -> Path:
    for rel in rel_paths:
        p = base / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(b"%PDF-1.4 test")
    return base


def test_root_first_then_subfolders_alphabetical(tmp_path):
    root = make_tree(
        tmp_path,
        "zebra.pdf",
        "alpha.pdf",
        "Reports/summary.pdf",
        "Invoices/002.pdf",
        "Invoices/001.pdf",
    )

    result = collect_pdfs(root)
    names = [(group, p.name) for group, p in result]

    assert names == [
        ("(Root)", "alpha.pdf"),
        ("(Root)", "zebra.pdf"),
        ("Invoices", "001.pdf"),
        ("Invoices", "002.pdf"),
        ("Reports", "summary.pdf"),
    ]


def test_case_insensitive_sorting(tmp_path):
    root = make_tree(
        tmp_path,
        "Banana.pdf",
        "apple.pdf",
        "Cherry.pdf",
    )

    result = collect_pdfs(root)
    names = [p.name for _, p in result]

    assert names == ["apple.pdf", "Banana.pdf", "Cherry.pdf"]


def test_uppercase_pdf_extension_included(tmp_path):
    root = make_tree(
        tmp_path,
        "a.PDF",
        "b.Pdf",
        "c.pdf",
    )

    result = collect_pdfs(root)
    assert len(result) == 3


def test_non_pdf_files_ignored(tmp_path):
    root = make_tree(
        tmp_path,
        "doc.txt",
        "image.png",
        "spreadsheet.xlsx",
        "keep.pdf",
    )

    result = collect_pdfs(root)
    names = [p.name for _, p in result]
    assert names == ["keep.pdf"]


def test_single_level_only(tmp_path):
    root = make_tree(
        tmp_path,
        "top.pdf",
        "Sub/inner.pdf",
        "Sub/nested/deeper.pdf",
    )

    result = collect_pdfs(root)
    names = [(g, p.name) for g, p in result]

    # "nested" is a sub-subfolder and must not be traversed
    assert names == [("(Root)", "top.pdf"), ("Sub", "inner.pdf")]


def test_empty_directory(tmp_path):
    assert collect_pdfs(tmp_path) == []


def test_empty_subfolder_yields_no_group(tmp_path):
    root = make_tree(tmp_path, "only.pdf", "EmptyFolder/.gitkeep")

    result = collect_pdfs(root)
    names = [p.name for _, p in result]

    assert names == ["only.pdf"]
