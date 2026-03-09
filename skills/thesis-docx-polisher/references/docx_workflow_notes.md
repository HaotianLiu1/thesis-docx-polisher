# DOCX Workflow Notes

## TOC

- Chapter range detection
- Paragraph filtering policy
- Batch processing semantics
- Diff rendering semantics
- Stability and retry behavior

## Chapter range detection

- Primary chapter anchors are `Heading 1` paragraphs in order.
- `--start-chapter` and `--end-chapter` refer to this order.
- `--end-chapter 0` means process until paragraph whose text is `参考文献`, or file end if missing.

## Paragraph filtering policy

The script skips paragraphs likely not body text based on `--skip-style-keys` (comma separated). Default keys are:

- `Heading`
- `标题-图`
- `标题-表格`
- `标题-无编号`
- `参考文献`
- `TOC`

It also skips very short text (`len < --min-len`).

## Batch processing semantics

- Candidate paragraphs are collected first, then split by `--paragraphs-per-call`.
- Each batch is sent as `items=[{id,text}, ...]`.
- Returned results are mapped by `id` to avoid ordering issues.
- `--limit` is applied on candidate count before batching.

Fallback behavior:
- Whole batch fallback to original text on non-JSON response.
- Per-item fallback on missing/invalid `id` or empty `revised_text`.

## Diff rendering semantics

Paragraph-level old/new diff is produced with token alignment:

- equal chunk: keep default style
- insert chunk: set font color red (`RGB FF0000`)
- delete chunk: set strike (`font.strike=True`)

This gives a git-like visual comparison inside Word while keeping one final paragraph.

## Stability and retry behavior

- Retry on request exceptions, HTTP `429`, and HTTP `5xx`.
- `Retry-After` header is honored when present.
- Exponential backoff fallback: `retry_backoff * 2^attempt`.
- `--sleep` throttles between batches.
- `--fail-log` writes JSONL entries for API errors and fallback statuses.

Example `--fail-log` lines:

```json
{"paragraph_index": 645, "status": "api_error", "error": "HTTPError: 502 Server Error"}
{"paragraph_index": 646, "status": "id_missing"}
```
