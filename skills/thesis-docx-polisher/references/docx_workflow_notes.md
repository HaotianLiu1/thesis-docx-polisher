# DOCX Workflow Notes

## TOC

- Chapter range detection
- Paragraph filtering policy
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

## Diff rendering semantics

Paragraph-level old/new diff is produced with token alignment:

- equal chunk: keep default style
- insert chunk: set font color red (`RGB FF0000`)
- delete chunk: set strike (`font.strike=True`)

This gives a git-like visual comparison inside Word while keeping one final paragraph.

## Stability and retry behavior

- Retry on exceptions and HTTP 5xx.
- Exponential backoff: `retry_backoff * 2^attempt`.
- `--fail-log` writes JSONL entries with paragraph index and error.

Example `--fail-log` line:

```json
{"paragraph_index": 645, "error": "HTTPError: 502 Server Error"}
```
