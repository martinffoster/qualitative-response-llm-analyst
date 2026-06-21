from qrla.resources import bundled_template_path
from qrla.validation import validate_workbook


def test_template_validates_against_itself() -> None:
    template = bundled_template_path()
    valid, issues = validate_workbook(template, template)
    assert valid, issues
