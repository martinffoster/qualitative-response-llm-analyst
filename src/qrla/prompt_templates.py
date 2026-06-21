"""Load and render LLM prompt templates from the prompts package directory."""

from __future__ import annotations

from importlib import resources
from typing import Any, Mapping

PROMPT_STAGE1_DISCOVERY = "stage1_discovery.txt"
PROMPT_STAGE1_REVIEW = "stage1_review.txt"
PROMPT_STAGE2_ASSIGNMENT = "stage2_assignment.txt"


def _read_template(name: str) -> str:
    """Read a prompt template file bundled with the package."""
    return resources.files("qrla.prompts").joinpath(name).read_text(encoding="utf-8")


def _strip_metadata_comments(text: str) -> str:
    """Remove leading # comment lines (metadata header) from a template file."""
    lines = text.splitlines()
    while lines and lines[0].startswith("#"):
        lines.pop(0)
    while lines and not lines[0].strip():
        lines.pop(0)
    return "\n".join(lines)


def render_prompt(template_name: str, **variables: Any) -> str:
    """Load a template and substitute ``{placeholders}`` with the given variables."""
    raw = _strip_metadata_comments(_read_template(template_name))
    return raw.format_map(_SafeFormatDict(variables))


def optional_context_line(label: str, value: str | None) -> str:
    """Format an optional context line, or return empty string if value is missing."""
    if not value:
        return ""
    return f"{label}: {value}\n"


def format_responses_block(responses: list[Mapping[str, Any]]) -> str:
    """Format response rows as bullet lines for inclusion in a prompt."""
    lines: list[str] = []
    for rec in responses:
        rid = rec.get("response_id") or rec.get("id") or ""
        txt = str(rec.get("response_text", "") or "").strip()
        if txt:
            lines.append(f"- [{rid}] {txt}")
    return "\n".join(lines)


class _SafeFormatDict(dict):
    """Return empty string for missing format keys instead of raising KeyError."""

    def __missing__(self, key: str) -> str:
        return ""
