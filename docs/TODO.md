# TODO

Outstanding documentation and product gaps for the public repo.

## Unimplemented features

- **`qrla export` (Stage 4)** — Flatten coded workbook data to CSV or Parquet for dashboards and downstream analysis. Referenced in the workflow spec; no CLI command yet.

## Example data

- **Synthetic example workbooks** — Add committed example files under `examples/` (not `data/`) so users can run `qrla validate` and later-stage commands without their own survey data or API keys where possible. Likely a small staged set:
  - Minimal ingest: question + responses only (for discovery demos)
  - Post-discovery: candidate themes populated (for review/assign demos)
  - Post-assignment: fake multi-model columns (for `review-assignments` and `summarize` offline)
- Use obviously synthetic response IDs and generic question domains; never commit real survey responses.
