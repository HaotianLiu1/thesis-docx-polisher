# API and Prompt Contract

## OpenAI-compatible endpoint

The script calls:

- `POST <base_url>/v1/chat/completions`
- Header: `Authorization: Bearer <api_key>`
- Header: `Content-Type: application/json`

## Batch payload shape

The script now sends paragraph batches in one request (size controlled by `--paragraphs-per-call`):

```json
{
  "model": "gpt-5.4",
  "temperature": 0.2,
  "messages": [
    {"role": "system", "content": "..."},
    {
      "role": "user",
      "content": "... {\"items\":[{\"id\":0,\"text\":\"段落A\"},{\"id\":1,\"text\":\"段落B\"}]} ..."
    }
  ]
}
```

Notes:
- `id` is the local index within the current batch (`0..batch_size-1`).
- `--paragraphs-per-call 1` preserves legacy one-paragraph-per-request behavior.

## Required model response format

Model must return JSON in this contract:

```json
{
  "items": [
    {"id": 0, "need_edit": true, "revised_text": "..."},
    {"id": 1, "need_edit": false, "revised_text": "..."}
  ]
}
```

Rules:
- Each returned item maps back by `id`, not by order.
- If `need_edit=false` or `revised_text` equals original, paragraph remains unchanged.

## Fallback / fault tolerance policy

- Whole-response parse failure (`non-JSON`): all paragraphs in that batch fallback to original text.
- Missing/invalid `id`: only that paragraph falls back.
- Empty `revised_text`: only that paragraph falls back.
- Fallback events are written to `--fail-log` with `status`.

## Retry / rate limit behavior

- Retries include request exceptions, HTTP `429`, and HTTP `5xx`.
- Backoff: `retry_backoff * 2^attempt`.
- If response has `Retry-After`, it is used as wait time before exponential fallback.
- `--sleep` applies between batches for lightweight client-side throttling.

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
