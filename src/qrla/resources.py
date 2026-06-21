"""Bundled static assets shipped with the package."""

from __future__ import annotations

from importlib import resources
from pathlib import Path

TEMPLATE_FILENAME = "qual_coding_template.xlsx"


def bundled_template_path() -> str:
    """Return the filesystem path to the bundled canonical workbook template."""
    ref = resources.files("qrla").joinpath("templates", TEMPLATE_FILENAME)
    with resources.as_file(ref) as path:
        return str(path)


def copy_bundled_template(destination: Path) -> Path:
    """Copy the bundled template workbook to ``destination``."""
    ref = resources.files("qrla").joinpath("templates", TEMPLATE_FILENAME)
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_bytes(ref.read_bytes())
    return destination
