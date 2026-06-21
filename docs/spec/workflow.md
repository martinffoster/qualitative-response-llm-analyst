# Workflow (End-to-End)

This document summarizes the stage-by-stage workflow. Keep this file and the root `README.md` synchronized.

- Stage 0 — Ingestion: normalize raw survey exports to the template structure.
- Stage 1 — Theme discovery/maintenance: curate `theme_groups` and `theme_catalog` (limit active themes ~30; candidates flagged).
- Stage 2 — Theme assignment: assign up to N=5 themes per response per model; store per-model columns; compute proposed finals.
- Stage 3 — Human review & audit: confirm `final_theme_*`, set `response_quotable`, record corrections.
- Stage 4 — Export & analysis: flatten for dashboards; aggregate by theme and theme_group. **Not implemented** (`qrla export` is planned but not yet available).

LLM access via OpenRouter. See `docs/spec/overview.md` for constraints.
