"""Package metadata smoke tests."""

from pdf_folder_print import __version__


def test_version_is_semver_like():
    parts = __version__.split(".")
    assert len(parts) == 3
    assert all(p.isdigit() for p in parts)
