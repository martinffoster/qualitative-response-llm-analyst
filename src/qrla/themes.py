from dataclasses import dataclass
import re


def normalize_label(label: str) -> str:
    """Normalize a human label to short snake_case.

    - lowercase
    - remove non-alphanumeric (except spaces and underscores)
    - replace whitespace with underscore
    - collapse multiple underscores
    - strip leading/trailing underscores
    """
    if not label:
        return ""
    s = label.lower()
    s = re.sub(r"[^a-z0-9_\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    s = s.replace(" ", "_")
    s = re.sub(r"_+", "_", s)
    s = s.strip("_")
    return s


def generate_theme_id(survey: str, question: str, label: str, seq: int) -> str:
    """Generate a short snake_case theme id.

    Example: bai_cb02_exec_support_001
    """
    survey_part = normalize_label(survey)[:20]
    question_part = normalize_label(question)[:20]
    label_part = normalize_label(label)[:40]
    seq_part = f"{seq:03d}"
    parts = [p for p in (survey_part, question_part, label_part) if p]
    base = "_".join(parts)
    if not base:
        base = "theme"
    return f"{base}_{seq_part}"


@dataclass
class Theme:
    theme_id: str
    theme_label: str
    theme_definition: str | None = None
    theme_group_id: str | None = None
    status: str = "candidate"

    def as_dict(self) -> dict:
        return {
            "theme_id": self.theme_id,
            "theme_label": self.theme_label,
            "theme_definition": self.theme_definition,
            "theme_group_id": self.theme_group_id,
            "status": self.status,
        }


