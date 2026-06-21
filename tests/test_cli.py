from pathlib import Path

from typer.testing import CliRunner

from qrla.cli import app

runner = CliRunner()


def test_cli_help() -> None:
    result = runner.invoke(app, ["--help"])
    assert result.exit_code == 0
    assert "validate" in result.stdout


def test_init_template(tmp_path: Path) -> None:
    output = tmp_path / "survey.xlsx"
    result = runner.invoke(app, ["init-template", str(output)])
    assert result.exit_code == 0
    assert output.is_file()
    assert "Wrote template" in result.stdout


def test_validate_bundled_template(tmp_path: Path) -> None:
    workbook = tmp_path / "survey.xlsx"
    runner.invoke(app, ["init-template", str(workbook)])
    result = runner.invoke(app, ["validate", str(workbook)])
    assert result.exit_code == 0
    assert "Validation passed" in result.stdout
