import typer
from pathlib import Path

from .validation import validate_workbook
from .themes import generate_theme_id, normalize_label
from .stage1_discovery import discover as discover_stage1
from .stage1_review import review_themes as review_themes_stage1
from .stage2_assignment import assign_themes as assign_stage2
from .stage2_review_assignments import create_review_assignments_sheet
from .stage3_summary import summarize as summarize_stage3

app = typer.Typer()


@app.command()
def validate(
    workbook: str = typer.Argument(..., help="Path to workbook to validate."),
    template: str = typer.Option("templates/qual_coding_template.xlsx", help="Path to template workbook."),
):
    """Validate a workbook against the canonical template."""
    wb_path = Path(workbook)
    tpl_path = Path(template)
    if not wb_path.exists():
        typer.echo(f"workbook not found: {wb_path}")
        raise typer.Exit(code=2)
    if not tpl_path.exists():
        typer.echo(f"template not found: {tpl_path}")
        raise typer.Exit(code=2)
    valid, issues = validate_workbook(str(wb_path), str(tpl_path))
    if valid:
        typer.echo("Validation passed — workbook complies with template.")
    else:
        typer.echo("Validation failed:")
        for it in issues:
            typer.echo(f" - {it}")
        raise typer.Exit(code=1)


@app.command()
def id_example(
    survey: str = typer.Option("bai", help="Survey short name, e.g., 'bai'"),
    question: str = typer.Option("cb02", help="Question short code, e.g., 'cb02'"),
    label: str = typer.Option("Executive support", help="Theme label to normalize"),
    seq: int = typer.Option(1, help="Sequence number"),
):
    """Print an example generated theme id."""
    tid = generate_theme_id(survey, question, label, seq)
    typer.echo(tid)


if __name__ == "__main__":
    app()

@app.command()
def discover(
    workbook: str = typer.Argument(..., help="Path to workbook to process."),
    question_id: str = typer.Option(..., "--question-id", help="Question ID to process (e.g., SURVEY_2025_Q01)."),
    model: str = typer.Option(..., help="OpenRouter model name to use."),
    sample_size: int | None = typer.Option(None, help="Number of responses to sample for discovery. If not specified, all responses are used."),
    temperature: float = typer.Option(0.2, help="Sampling temperature for the model."),
    max_tokens: int = typer.Option(32000, help="Maximum output tokens. Should be less than model's context window minus input tokens."),
    max_themes: int = typer.Option(30, help="Maximum number of themes to discover (default: 30)."),
    output: str = typer.Option(None, help="Output workbook path; defaults to input with .discovered.<model>.xlsx suffix."),
    verbose: int = typer.Option(0, "--verbose", "-v", count=True, help="Verbosity level. -v for stats, -vv to also show prompts and LLM responses."),
    context_column: str | None = typer.Option(None, "--context-column", help="Column name in 'question' sheet containing custom context (domain knowledge, coding guidelines, etc.)."),
):
    """Run Stage 1 Theme Discovery on a workbook, writing a new workbook with candidate themes/groups and audit."""
    out = discover_stage1(
        workbook,
        question_id=question_id,
        model=model,
        sample_size=sample_size,
        temperature=temperature,
        max_tokens=max_tokens,
        max_themes=max_themes,
        output=output,
        verbose=verbose,
        context_column=context_column,
    )
    typer.echo(f"Discovery complete. Wrote: {out}")


@app.command()
def review_themes(
    workbook: str = typer.Argument(..., help="Path to workbook to process."),
    question_id: str = typer.Option(..., "--question-id", help="Question ID to process (e.g., SURVEY_2025_Q01)."),
    model: str = typer.Option(..., help="OpenRouter model name to use."),
    sample_size: int | None = typer.Option(None, help="Number of responses to sample for review. If not specified, all responses are used."),
    temperature: float = typer.Option(0.2, help="Sampling temperature for the model."),
    max_tokens: int = typer.Option(32000, help="Maximum output tokens. Should be less than model's context window minus input tokens."),
    max_new_themes: int = typer.Option(30, help="Maximum number of new themes to suggest (default: 30)."),
    output: str = typer.Option(None, help="Output workbook path; defaults to input with .reviewed.<model>.xlsx suffix."),
    verbose: int = typer.Option(0, "--verbose", "-v", count=True, help="Verbosity level. -v for stats, -vv to also show prompts and LLM responses."),
    context_column: str | None = typer.Option(None, "--context-column", help="Column name in 'question' sheet containing custom context (domain knowledge, coding guidelines, etc.)."),
):
    """Run Stage 1.5 Theme Review on a workbook, reviewing existing themes and suggesting additions/retirements."""
    out = review_themes_stage1(
        workbook,
        question_id=question_id,
        model=model,
        sample_size=sample_size,
        temperature=temperature,
        max_tokens=max_tokens,
        max_new_themes=max_new_themes,
        output=output,
        verbose=verbose,
        context_column=context_column,
    )
    typer.echo(f"Review complete. Wrote: {out}")


@app.command()
def assign(
    workbook: str = typer.Argument(..., help="Path to workbook to process."),
    question_id: str = typer.Option(..., "--question-id", help="Question ID to process (e.g., SURVEY_2025_Q01)."),
    model: str = typer.Option(..., help="OpenRouter model name to use."),
    max_themes: int = typer.Option(5, help="Maximum number of themes to assign per response (max: 50, default: 5). Warning if >20."),
    sample_size: int | None = typer.Option(None, help="Number of responses to process (optional; if not specified, all responses are used)."),
    temperature: float = typer.Option(0.2, help="Sampling temperature for the model."),
    max_tokens: int = typer.Option(32000, help="Maximum output tokens."),
    response_ids: str | None = typer.Option(None, "--response-ids", help="Comma-separated list of response IDs to process. If specified, only these responses are processed (overrides sample_size)."),
    skip_existing: bool = typer.Option(False, "--skip-existing", help="Skip responses that already have assignments for this model (checks if model_X_theme_1 column has a value)."),
    verbose: int = typer.Option(0, "--verbose", "-v", count=True, help="Verbosity level. -v for progress and stats, -vv to also show prompts."),
    context_column: str | None = typer.Option(None, "--context-column", help="Column name in 'question' sheet containing custom context (domain knowledge, coding guidelines, etc.)."),
):
    """Run Stage 2 Theme Assignment on a workbook, assigning themes to each response."""
    # Validate max_themes
    if max_themes > 50:
        typer.echo(f"Error: max_themes cannot exceed 50 (got {max_themes})", err=True)
        raise typer.Exit(code=1)
    elif max_themes > 20:
        typer.echo(f"Warning: max_themes={max_themes} is high. This will create {max_themes * 2} columns per model.", err=True)
    
    out = assign_stage2(
        workbook,
        question_id=question_id,
        model=model,
        max_themes=max_themes,
        sample_size=sample_size,
        temperature=temperature,
        max_tokens=max_tokens,
        response_ids=response_ids,
        skip_existing=skip_existing,
        verbose=verbose,
        context_column=context_column,
    )
    typer.echo(f"Assignment complete. Updated: {out}")


@app.command()
def summarize(
    workbook: str = typer.Argument(..., help="Path to workbook to process."),
    question_id: str = typer.Option(..., "--question-id", help="Question ID to process (e.g., SURVEY_2025_Q01)."),
    chart: str = typer.Option(None, "--chart", help="Chart type: 'stackedbar' (requires groups), 'bar' (ignores groups), or 'auto' (auto-detect). If not specified, no chart is created."),
    verbose: int = typer.Option(0, "--verbose", "-v", count=True, help="Verbosity level. -v shows model detection, theme counts, and consistency warnings."),
):
    """Create summary sheets for each model's theme assignments to facilitate human review."""
    if chart and chart not in ("stackedbar", "bar", "auto"):
        typer.echo(f"Error: --chart must be one of: 'stackedbar', 'bar', or 'auto'", err=True)
        raise typer.Exit(code=1)
    
    out = summarize_stage3(
        workbook,
        question_id=question_id,
        chart_type=chart,
        verbose=verbose,
    )
    typer.echo(f"Summary sheets created. Updated: {out}")


@app.command("review-assignments")
def review_assignments(
    workbook: str = typer.Argument(..., help="Path to workbook"),
    question_id: str = typer.Option(..., "--question-id", help="Question ID to process"),
    verbose: int = typer.Option(0, "-v", "--verbose", count=True, help="Verbose output"),
) -> None:
    """Create review_assignments sheet for easier assignment review."""
    out = create_review_assignments_sheet(
        workbook=workbook,
        question_id=question_id,
        verbose=verbose,
    )
    typer.echo(f"Review assignments sheet created. Updated: {out}")
