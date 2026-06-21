# Qualitative Response LLM Analyst

Tools for **coding qualitative survey responses** using a hybrid of **human review** and **multi-model LLM assistance**.

LLMs via [OpenRouter](https://openrouter.ai/).

The workflow supports:

- Discovering and maintaining **themes** and **theme groups**
- Assigning up to **five themes per response** (configurable)
- Flagging **candidate themes** for human approval
- Producing **auditable coded datasets** in Excel workbooks
- Summarising model assignments for human-in-the-loop review

All behaviour follows the specification in the [project overview](https://github.com/martinffoster/qualitative-response-llm-analyst/blob/main/docs/spec/overview.md).

## History and Purpose

First version ~July 2023; with major improvements 2024, 2025.  

Used to support the qualitative analysis of:
* [Business Agility Insititute](https://businessagility.institute/)'s Business Agility Report [2023](https://businessagility.institute/learn/2023-business-agility-report/751),[2024](https://businessagility.institute/learn/2024-business-agility-report/754),[2025](https://businessagility.institute/learn/the-2025-business-agility-report/758). [Personal site cache](https://netlog.net/publications/business-agility-report/)
* [Organisational Design Forum Practitioner Survey 2026](https://organizationdesignforum.org/odf-research/). [Personal site cache](https://netlog.net/publications/odf-global-practitioner-report-2026/)

## Repository structure

```
.
├── docs/
│   ├── agents/          Agent and contributor guidance
│   └── spec/            Workbook schema, workflow, constraints
├── prompts/             Authoring notes for LLM prompt templates
├── src/qrla/            Python package (`qrla` CLI)
│   ├── prompts/         Bundled .txt templates loaded at runtime
│   └── templates/       Bundled canonical Excel workbook template
├── .env.example         Environment variable template
└── README.md
```

Survey workbooks live under `data/` (gitignored). Create a new workbook from the bundled template (see below).

## Requirements

- Python 3.11+
- An [OpenRouter](https://openrouter.ai/) API key

## Installation

From PyPI:

```bash
pip install qualitative-response-llm-analyst
```

For local development:

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -e ".[dev]"
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

Create a workbook from the bundled template:

```bash
qrla init-template data/my-survey-Q01-themes.xlsx
```

Keep sheet names and column headers unchanged. Full schema details: [schemas](https://github.com/martinffoster/qualitative-response-llm-analyst/blob/main/docs/spec/schemas.md).

## CLI commands

The `qrla` command implements the end-to-end workflow:

| Command              | Stage | Description                                      |
| -------------------- | ----- | ------------------------------------------------ |
| `qrla init-template`   | —     | Copy the bundled canonical workbook template     |
| `qrla validate`        | —     | Check a workbook against the template            |
| `qrla discover`        | 1     | Propose themes from response samples             |
| `qrla review-themes`   | 1.5   | Validate, extend, or retire existing themes      |
| `qrla assign`          | 2     | Assign themes to each response                   |
| `qrla review-assignments` | 2.5 | Create a tidy review sheet for filtering      |
| `qrla summarize`       | 3     | Per-model summary sheets (optional charts)       |
| `qrla export`          | 4     | **Not implemented** — flatten workbook to CSV/Parquet |

### Quick start

```bash
# Create and validate a new workbook
qrla init-template data/my-survey-Q01-themes.xlsx
qrla validate data/my-survey-Q01-themes.xlsx

# Stage 1 — discover themes from responses
qrla discover data/my-survey-Q01-themes.xlsx \
  --question-id SURVEY_2025_Q01 \
  --model openai/gpt-5.5 \
  --max-themes 30 \
  -v

# Stage 2 — assign themes (run once per model for multi-model comparison)
qrla assign data/my-survey-Q01-themes.discovered.openai_gpt_5_5.xlsx \
  --question-id SURVEY_2025_Q01 \
  --model openai/gpt-5.4-mini \
  -v

# Stage 3 — summarise assignments for human review
qrla summarize data/my-survey-Q01-themes.discovered.openai_gpt_5_5.xlsx \
  --question-id SURVEY_2025_Q01 \
  --chart auto \
  -v
```

Common options:

- `--question-id` — must match a row in the `question` sheet
- `--model` — OpenRouter model id (e.g. `openai/gpt-5.5`, `anthropic/claude-sonnet-4.6`)
- `--max-themes` — cap on themes discovered or assigned per response
- `--context-column` — optional column on the `question` sheet with domain-specific coding guidance
- `-v` / `-vv` — progress stats; `-vv` also prints prompts and raw LLM output

See the [workflow spec](https://github.com/martinffoster/qualitative-response-llm-analyst/blob/main/docs/spec/workflow.md) for the full stage-by-stage process.

## Model guidance

Typical choices (via OpenRouter — pass the model id to `--model`):

- **Discovery / theme review:** capable frontier models — e.g. `openai/gpt-5.5`, `anthropic/claude-sonnet-4.6`, `google/gemini-3.1-pro-preview`, `x-ai/grok-4.3`, `mistralai/mistral-large-2512`
- **Assignment:** faster, cheaper models — e.g. `openai/gpt-5.4-mini`, `anthropic/claude-haiku-4.5`, `google/gemini-3.5-flash`, `x-ai/grok-build-0.1`, `mistralai/mistral-small-2603`

Run several assignment models and compare results before final human coding.

## Theme status values

| Status            | Meaning                          |
| ----------------- | -------------------------------- |
| `candidate`       | Proposed by LLM; needs review    |
| `candidate-add`   | Suggested new theme              |
| `candidate-retire`| Suggested retirement             |
| `active`          | Approved; used in assignment     |
| `retired`         | Historical; excluded from prompts|

## Documentation

- [Specification overview](https://github.com/martinffoster/qualitative-response-llm-analyst/blob/main/docs/spec/overview.md)
- [Workflow](https://github.com/martinffoster/qualitative-response-llm-analyst/blob/main/docs/spec/workflow.md)
- [Schemas](https://github.com/martinffoster/qualitative-response-llm-analyst/blob/main/docs/spec/schemas.md)
- [Agents guide](https://github.com/martinffoster/qualitative-response-llm-analyst/blob/main/docs/agents/README.md)
- [TODO](https://github.com/martinffoster/qualitative-response-llm-analyst/blob/main/docs/TODO.md)

## Contributing

1. Fork and branch (e.g. `feature/stage2-improvements`)
2. Follow PEP 8; use type hints and docstrings
3. Keep `README.md` and `docs/` in sync when changing behaviour
4. Do not commit survey data or API keys
5. Run tests locally with `pytest` before opening a PR (CI runs the same checks on GitHub Actions)

## Releasing

Production releases use [PyPI trusted publishing](https://docs.pypi.org/trusted-publishers/). TestPyPI uses the same mechanism for dry runs before the first real upload.

### One-time setup

1. Create an account at [test.pypi.org](https://test.pypi.org/) (separate from production PyPI).
2. Enable **2FA** on both TestPyPI and production PyPI.
3. Add a **trusted publisher** on each site (Account settings → Publishing, or project settings after the first upload):
   - **TestPyPI:** workflow `test-release.yml`, environment `testpypi` (optional but recommended)
   - **Production PyPI:** workflow `release.yml`, environment `pypi` (optional but recommended)
4. In GitHub: **Settings → Environments** — create `testpypi` and `pypi` if you want approval gates before publish.

### Dry run on TestPyPI

Use this before tagging a production release.

1. Ensure `version` in `pyproject.toml` is the version you want to test (TestPyPI allows re-upload only if you bump the version or delete the file).
2. Push your branch to GitHub.
3. Open **Actions → Test Release → Run workflow** and start the run.
4. Install from TestPyPI (dependencies still come from production PyPI):

```bash
python -m venv .venv-test
source .venv-test/bin/activate   # Windows: .venv-test\Scripts\activate
pip install -i https://test.pypi.org/simple/ \
  --extra-index-url https://pypi.org/simple/ \
  qualitative-response-llm-analyst
```

5. Smoke-test the install:

```bash
qrla --help
qrla init-template /tmp/test-workbook.xlsx
qrla validate /tmp/test-workbook.xlsx
```

6. When satisfied, update [CHANGELOG.md](./CHANGELOG.md), bump `version` in `pyproject.toml`, commit, then tag and push:

```bash
git tag v0.1.1
git push origin v0.1.1
```

That triggers **Release** (`release.yml`), which publishes to [PyPI](https://pypi.org/project/qualitative-response-llm-analyst/) and creates a [GitHub Release](https://github.com/martinffoster/qualitative-response-llm-analyst/releases) with notes and attached sdist/wheel.

Install from production PyPI:

```bash
pip install qualitative-response-llm-analyst
```

### Git tag vs GitHub Release

A **git tag** (`v0.1.0`) triggers the release workflow and PyPI upload. A **GitHub Release** is the page users see under [Releases](https://github.com/martinffoster/qualitative-response-llm-analyst/releases) with notes and downloadable assets. From `v0.1.1` onward, `release.yml` creates both automatically.

### Backfill an existing tag (e.g. `v0.1.0`)

If you tagged before GitHub Releases were configured, create the release once without re-publishing to PyPI:

```bash
gh release create v0.1.0 \
  --title "0.1.0" \
  --notes-file CHANGELOG.md \
  --verify-tag
```

Or in the web UI: **Releases → Draft a new release → Choose tag `v0.1.0` → paste notes → Publish release**.

**Note:** TestPyPI is ephemeral — do not rely on packages staying there long term. It exists to validate packaging and install before the real release.

## License

BSD 2-Clause License — see [LICENSE](./LICENSE).

Maintained by Martin Foster and collaborators.
