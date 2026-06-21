# Agents Guide

The following rules govern agent and contributor behavior in this repository. Treat these as enforceable policies.

- You are working in this repo. All behavior must comply with `docs/spec/overview.md`.
- Do not invent new columns, rename columns, or remove audit fields unless instructed.
- We'll be coding in Python 3.11+, PEP8, with type hints. Make small contained changes.
- No refactoring existing code unless necessary and agreed with the user.
- Do not commit survey workbooks under `data/`; they may contain respondent PII.
- Use OpenRouter for LLM access. Secrets live in `.env`.

See the full specification for context and invariants:

- Spec: ../spec/overview.md

Documentation hygiene for agents:

- When updating guidance or schemas, keep the root `README.md` and `docs/` in sync.
- The root `README.md` is the friendly entry point; `docs/` is the canonical source of truth.
- If there is a discrepancy, update both places in the same edit.

Naming:

- Project/distribution name: `qualitative-response-llm-analyst`
- Python import package: `qrla`
- CLI entry point: `qrla`


