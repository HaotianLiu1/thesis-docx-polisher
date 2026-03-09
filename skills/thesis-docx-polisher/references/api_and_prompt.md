# API and Prompt Contract

## OpenAI-compatible endpoint

The script calls:

- `POST <base_url>/v1/chat/completions`
- Header: `Authorization: Bearer <api_key>`
- Header: `Content-Type: application/json`

Payload shape:

```json
{
  "model": "gpt-5.4",
  "temperature": 0.2,
  "messages": [
    {"role": "system", "content": "..."},
    {"role": "user", "content": "<paragraph_text>"}
  ]
}
```

## Required model response format

The script expects JSON (or fenced JSON) in this format:

```json
{"need_edit": true, "revised_text": "..."}
```

Rules:

- If `need_edit=false` and `revised_text` unchanged, paragraph remains unchanged.
- If parser cannot extract valid JSON, paragraph is treated as unchanged.

## Prompt intent

The system prompt enforces:

1. Detect AI-like heavy tone and awkward/unfluent phrasing.
2. Apply minimal edits only.
3. Prefer basic/common connectors and direct wording.
4. Keep logic coherent and transitions natural.
5. Preserve technical terms, numbering, formulas, figure/table/citation markers.

## Environment variables

Script supports env vars as defaults:

- `OPENAI_BASE_URL`
- `OPENAI_API_KEY`
- `OPENAI_MODEL`

CLI options override env vars.
