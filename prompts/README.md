# Prompt templates

LLM prompt templates live in [`src/qrla/prompts/`](../src/qrla/prompts/).

| File | Stage | CLI command |
| ---- | ----- | ----------- |
| `stage1_discovery.txt` | Theme discovery | `qrla discover` |
| `stage1_review.txt` | Theme review | `qrla review-themes` |
| `stage2_assignment.txt` | Theme assignment | `qrla assign` |

Each file starts with `#` metadata comments (version, stage, task type). The loader in `prompt_templates.py` strips those before rendering.

Templates use `{placeholder}` syntax. Workbook data (responses, theme catalog, etc.) is formatted in Python and passed in as blocks such as `{responses_block}` and `{theme_catalog_block}`.

To change LLM behaviour, edit the `.txt` file and bump the `# version:` comment. Prompt hashes in `model_runs` will reflect the change.
