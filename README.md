# Qualitative Response LLM Analyst

Tools for **coding qualitative survey responses** using a hybrid of **human review** and **multi-model LLM assistance** via [OpenRouter](https://openrouter.ai/).

The workflow supports:

- Discovering and maintaining **themes** and **theme groups**
- Assigning up to **five themes per response** (configurable)
- Flagging **candidate themes** for human approval
- Producing **auditable coded datasets** in Excel workbooks
- Summarising model assignments for human-in-the-loop review

All behaviour follows the specification in [`docs/spec/overview.md`](./docs/spec/overview.md).

## Repository structure

```
.
├── docs/
│   ├── agents/          Agent and contributor guidance
│   └── spec/            Workbook schema, workflow, constraints
├── prompts/             LLM prompt templates (see prompts/README.md)
├── src/qrla/            Python package (`qrla` CLI)
│   └── prompts/         Bundled .txt templates loaded at runtime
├── templates/
│   └── qual_coding_template.xlsx   Canonical workbook structure
├── .env.example         Environment variable template
└── README.md
```

Survey workbooks live under `data/` (gitignored). Copy the template to create a new question workbook.

## Requirements

- Python 3.11+
- An [OpenRouter](https://openrouter.ai/) API key

## Installation

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -e .
```

Or install dependencies directly:

```bash
pip install pandas openpyxl python-dotenv httpx rich typer
```

## Configuration

Copy `.env.example` to `.env` and set your API key:

```bash
cp .env.example .env
# Edit .env — at minimum set OPENROUTER_API_KEY
```

Never commit `.env` files.

## Workbook model

Each question is one Excel workbook with these sheets:

| Sheet             | Purpose                                              |
| ----------------- | ---------------------------------------------------- |
| `question`        | Survey question text, context, and metadata          |
| `theme_groups`    | High-level conceptual clusters                       |
| `theme_catalog`   | Individual themes and definitions                    |
| `responses_coded` | One row per response; model outputs and human finals |
| `model_runs`      | Model metadata and parameters for audit              |

Copy [`templates/qual_coding_template.xlsx`](./templates/qual_coding_template.xlsx) and keep sheet names and column headers unchanged. Full schema details: [`docs/spec/schemas.md`](./docs/spec/schemas.md).

## CLI commands

The `qrla` command implements the end-to-end workflow:

| Command              | Stage | Description                                      |
| -------------------- | ----- | ------------------------------------------------ |
| `qrla validate`        | —     | Check a workbook against the template            |
| `qrla discover`        | 1     | Propose themes from response samples             |
| `qrla review-themes`   | 1.5   | Validate, extend, or retire existing themes      |
| `qrla assign`          | 2     | Assign themes to each response                   |
| `qrla review-assignments` | 2.5 | Create a tidy review sheet for filtering      |
| `qrla summarize`       | 3     | Per-model summary sheets (optional charts)       |
| `qrla export`          | 4     | **Not implemented** — flatten workbook to CSV/Parquet |

### Quick start

```bash
# Validate workbook structure
qrla validate data/my-survey-Q01-themes.xlsx

# Stage 1 — discover themes from responses
qrla discover data/my-survey-Q01-themes.xlsx \
  --question-id SURVEY_2025_Q01 \
  --model openai/gpt-4o \
  --max-themes 30 \
  -v

# Stage 2 — assign themes (run once per model for multi-model comparison)
qrla assign data/my-survey-Q01-themes.discovered.openai_gpt_4o.xlsx \
  --question-id SURVEY_2025_Q01 \
  --model anthropic/claude-3.5-haiku \
  -v

# Stage 3 — summarise assignments for human review
qrla summarize data/my-survey-Q01-themes.discovered.openai_gpt_4o.xlsx \
  --question-id SURVEY_2025_Q01 \
  --chart auto \
  -v
```

Common options:

- `--question-id` — must match a row in the `question` sheet
- `--model` — OpenRouter model id (e.g. `openai/gpt-4o`, `anthropic/claude-3.5-sonnet`)
- `--max-themes` — cap on themes discovered or assigned per response
- `--context-column` — optional column on the `question` sheet with domain-specific coding guidance
- `-v` / `-vv` — progress stats; `-vv` also prints prompts and raw LLM output

See [`docs/spec/workflow.md`](./docs/spec/workflow.md) for the full stage-by-stage process.

## Model guidance

Typical choices (via OpenRouter):

- **Discovery / theme review:** capable frontier models (e.g. GPT-4o, Claude Sonnet)
- **Assignment:** faster, cheaper models (e.g. GPT-4o mini, Gemini Flash, Claude Haiku)

Run several models on assignment and compare results before final human coding.

## Theme status values

| Status            | Meaning                          |
| ----------------- | -------------------------------- |
| `candidate`       | Proposed by LLM; needs review    |
| `candidate-add`   | Suggested new theme              |
| `candidate-retire`| Suggested retirement             |
| `active`          | Approved; used in assignment     |
| `retired`         | Historical; excluded from prompts|

## Documentation

- [Specification overview](./docs/spec/overview.md)
- [Workflow](./docs/spec/workflow.md)
- [Schemas](./docs/spec/schemas.md)
- [Agents guide](./docs/agents/README.md)
- [TODO](./docs/TODO.md)

## Contributing

1. Fork and branch (e.g. `feature/stage2-improvements`)
2. Follow PEP 8; use type hints and docstrings
3. Keep `README.md` and `docs/` in sync when changing behaviour
4. Do not commit survey data or API keys

## License

BSD 2-Clause License — see [LICENSE](./LICENSE).

Maintained by Martin Foster and collaborators.
