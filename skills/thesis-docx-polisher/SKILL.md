---
name: thesis-docx-polisher
description: Polish and lightly revise academic thesis DOCX files via an OpenAI-compatible API, with paragraph-level diff marking in Word (insertions in red, deletions with strikethrough). Use when the user asks to analyze and polish chapters in a thesis/paper/report .docx and wants reusable command-based processing with configurable API URL/key/model.
---

# Thesis DOCX Polisher

Use this skill to process thesis-like `.docx` files by calling an OpenAI-compatible API paragraph-by-paragraph, then writing a new `.docx` with visual diffs.

## Workflow

1. Confirm input file path and desired chapter range.
2. Run `scripts/polish_docx_via_api.py` with API config.
3. Verify output `.docx` exists and report statistics from script output.

## Run command

```bash
python ".claude/skills/thesis-docx-polisher/scripts/polish_docx_via_api.py" \
  --input "<input.docx>" \
  --output "<output.docx>" \
  --base-url "<openai-compatible-base-url>" \
  --api-key "<api-key>" \
  --model "<model-name>" \
  --start-chapter 1 \
  --end-chapter 5 \
  --chapter-style-key "Heading 1" \
  --reference-title "参考文献"
```

## Parameter rules

- `--start-chapter` and `--end-chapter` are based on first-level heading sequence.
- You can customize heading detection with `--chapter-style-key` (default `Heading 1`).
- `--end-chapter 0` means process from start chapter to the `--reference-title` text.
- Non-body paragraphs are skipped automatically by style keywords (`--skip-style-keys`).
- Diff rendering:
  - inserted text: red font
  - deleted text: strikethrough
  - unchanged text: default style/color

## Recommended reliability options

Use these when API endpoint is unstable:

```bash
--retries 2 --retry-backoff 1.5 --timeout 180 --fail-log "failures.jsonl"
```

## Read references when needed

- API/schema and prompt contract: `references/api_and_prompt.md`
- DOCX behavior and chapter detection details: `references/docx_workflow_notes.md`

## Script

Primary script: `scripts/polish_docx_via_api.py`
