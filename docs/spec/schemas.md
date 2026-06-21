# Schemas (Sheets and Columns)

Keep this file and the root `README.md` synchronized. Column names are authoritative.

## question
- question_id
- question_text
- question_context
- year
- notes

## theme_groups
- theme_group_id
- theme_group_label
- theme_group_definition
- question_id
- seen_years
- status (active|retired|candidate|candidate-add|candidate-retire)
- last_edited_by
- last_edited_at
- notes

## theme_catalog
- theme_id
- theme_group_id
- theme_label
- theme_definition
- question_id
- seen_years
- status (active|retired|candidate|candidate-add|candidate-retire)
- last_edited_by
- last_edited_at
- replaced_prior_theme_ids

**Status values:**
- `active`: Human-approved theme in use
- `candidate`: LLM-proposed theme awaiting human review (from `qrla discover`)
- `candidate-add`: LLM-suggested new theme from `qrla review-themes`, awaiting human approval
- `candidate-retire`: LLM-suggested retirement from `qrla review-themes`, awaiting human approval to move to `retired`
- `retired`: Human-approved retirement (no longer used)

## responses_coded
Identity
- question_id
- response_id
- response_text

Per-model assignments (repeat block per model)
- model_<name>_theme_1 .. model_<name>_theme_5
- (optional) model_<name>_truncated

Final human-approved
- final_theme_1 .. final_theme_5

Flags / QA / editorial
- final_theme_overflow_flag
- needs_review_flag
- response_quotable

Audit
- correction_version
- corrected_by
- corrected_at
- notes

Derived (optional)
- final_theme_group_1 .. final_theme_group_5

## model_runs
- model_run_id
- model_name
- model_version
- provider
- temperature
- timestamp_utc
- prompt_template_id
- task_type (theme_suggestion|theme_assignment|group_suggestion)
- notes
