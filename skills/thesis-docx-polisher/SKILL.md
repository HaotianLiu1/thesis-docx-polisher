---
name: thesis-docx-polisher
description: Polish and lightly revise academic thesis DOCX files via an OpenAI-compatible API, with paragraph-level diff marking in Word (insertions in red, deletions with strikethrough). Use when the user asks to analyze and polish chapters in a thesis/paper/report .docx and wants reusable command-based processing with configurable API URL/key/model.
---

# Thesis DOCX Polisher

Use this skill to process thesis-like `.docx` files by calling an OpenAI-compatible API in paragraph batches, then writing a new `.docx` with visual diffs.

## Workflow

1. Confirm input and output file paths, plus API config (`base-url/api-key/model`).
2. If the user did not specify any of these, ask follow-up questions before execution:
   - start chapter (`--start-chapter`)
   - end chapter (`--end-chapter`)
   - paragraphs per API call (`--paragraphs-per-call`)
3. For `--paragraphs-per-call`, suggest **3~8** by default, and warn that values too large may cause timeout, rate-limit amplification, or JSON parse fallback for an entire batch.
4. Run `scripts/polish_docx_via_api.py`.
5. Verify output `.docx` exists and report script stats (candidate/edited/errors/http_calls).

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
  --paragraphs-per-call 5 \
  --chapter-style-key "Heading 1" \
  --reference-title "参考文献"
```

## Parameter rules

- `--start-chapter` and `--end-chapter` are based on first-level heading sequence.
- `--end-chapter 0` means process from start chapter to the `--reference-title` text.
- `--paragraphs-per-call` controls batch size per API request.
  - `1` keeps legacy behavior (one paragraph per call).
  - Recommended: `3~8` for lower request count with manageable response size.
- You can customize heading detection with `--chapter-style-key` (default `Heading 1`).
- Non-body paragraphs are skipped automatically by style keywords (`--skip-style-keys`).
- Diff rendering:
  - inserted text: red font
  - deleted text: strikethrough
  - unchanged text: default style/color

## Recommended reliability options

Use these when API endpoint is unstable:

```bash
--retries 2 --retry-backoff 1.5 --timeout 180 --sleep 0.3 --fail-log "failures.jsonl"
```

Notes:
- Retries cover HTTP `429` and `5xx` plus request exceptions.
- If response has `Retry-After`, script waits by that value first.
- Fallback is paragraph-level: invalid/missing item id or empty `revised_text` keeps original text.

## Read references when needed

- API/schema and prompt contract: `references/api_and_prompt.md`
- DOCX behavior and chapter detection details: `references/docx_workflow_notes.md`

## Script

Primary script: `scripts/polish_docx_via_api.py`
