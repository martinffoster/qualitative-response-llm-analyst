# System Specification (Overview)

This document summarizes the core constraints and operating assumptions for agents and contributors.

- You are working in this repo. All behavior must comply with this spec.
- Do not invent new columns, rename columns, or remove audit fields unless explicitly instructed.
- Code in Python 3.11+, follow PEP 8, and use type hints. Prefer small, contained edits.
- Do not refactor existing code unless necessary and agreed with the user.
- Do not commit survey workbooks under `data/`; they may contain respondent PII.
- Use OpenRouter for LLM access. Secrets live in `.env`.

For implementation details, see the project `README.md` and related modules under `src/`.

Naming conventions:

- Project/distribution: `qualitative-response-llm-analyst`
- Python package: `qrla`
- CLI entry point: `qrla`
