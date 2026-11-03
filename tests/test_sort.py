"""Tests for natural_sort_key."""

from pdf_folder_print.core import natural_sort_key


def test_numbers_sort_numerically():
    names = ["file10.pdf", "file2.pdf", "file1.pdf"]
    assert sorted(names, key=natural_sort_key) == ["file1.pdf", "file2.pdf", "file10.pdf"]


def test_zero_padded_numbers_unchanged():
    names = ["002.pdf", "001.pdf", "010.pdf"]
    assert sorted(names, key=natural_sort_key) == ["001.pdf", "002.pdf", "010.pdf"]


def test_mixed_digit_and_text():
    names = ["page10", "page2", "pageX", "page1"]
    assert sorted(names, key=natural_sort_key) == ["page1", "page2", "page10", "pageX"]


def test_case_insensitive_text():
    assert natural_sort_key("apple") < natural_sort_key("Banana")


def test_no_digits_falls_back_to_alphabetical():
    names = ["b.pdf", "a.pdf", "c.pdf"]
    assert sorted(names, key=natural_sort_key) == ["a.pdf", "b.pdf", "c.pdf"]


def test_multiple_digit_runs():
    names = ["r2d10", "r2d2", "r10d1"]
    assert sorted(names, key=natural_sort_key) == ["r2d2", "r2d10", "r10d1"]
