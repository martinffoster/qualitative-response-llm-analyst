from pathlib import Path

from qrla.resources import TEMPLATE_FILENAME, bundled_template_path, copy_bundled_template


def test_bundled_template_path_exists() -> None:
    path = Path(bundled_template_path())
    assert path.is_file()
    assert path.name == TEMPLATE_FILENAME


def test_copy_bundled_template(tmp_path: Path) -> None:
    destination = tmp_path / "workbook.xlsx"
    copy_bundled_template(destination)
    assert destination.is_file()
    assert destination.stat().st_size > 0
