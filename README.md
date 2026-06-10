# mcp.office

Combined MCP server for Euro-Office DocumentServer workflows focused on:

- Word editor integration
- PDF editing/form workflows
- document conversion
- force-save and session info commands
- headless prompt-based document editing (no UI)

## Why combined MCP

Euro-Office DocumentServer exposes one command and conversion stack for multiple
formats. A single MCP server gives one integration point while still supporting
both Word and PDF use cases.

## Tools

- `check_health`: probes `/healthcheck` on DocumentServer
- `build_editor_config`: generates editor payload for `word`, `cell`, `slide`, or `pdf`
- `convert_document`: calls conversion endpoint
- `command_force_save`: triggers force save for an active document key
- `command_info`: fetches session info for a document key
- `describe_capabilities`: quick summary of Word/PDF support
- `edit_document_from_prompt`: non-visual editing for DOCX and PDF

## Environment variables

- `EURO_OFFICE_BASE_URL` (default: `http://localhost:8080`)
- `EURO_OFFICE_TIMEOUT_SECONDS` (default: `20`)
- `EURO_OFFICE_COMMAND_PATH` (default: `/coauthoring/CommandService.ashx`)
- `EURO_OFFICE_CONVERT_PATH` (default: `/ConvertService.ashx`)
- `EURO_OFFICE_JWT_SECRET` (optional, enables signed requests)
- `EURO_OFFICE_JWT_HEADER` (default: `Authorization`)
- `EURO_OFFICE_JWT_PREFIX` (default: `Bearer`)

## Local run

```bash
pip install -r requirements.txt
python server.py
```

Health checks:

- `GET /healthz`
- `GET /readyz`

## Docker

```bash
docker build -t ghcr.io/gdmkonsult/mcp.office:main .
docker run --rm -p 8000:8000 \
  -e EURO_OFFICE_BASE_URL=http://documentserver:8080 \
  -e EURO_OFFICE_JWT_SECRET=replace-me \
  ghcr.io/gdmkonsult/mcp.office:main
```

## Notes

- The MCP focuses on API-level editing orchestration, not UI automation.
- Your host application still needs to serve files and callback endpoints.
- Some environments use different command/convert paths; configure via env vars.
- Headless DOCX editing supports prompt actions like replace, append, prepend, and delete paragraph containing.
- Headless PDF editing supports either form-field updates (Field=Value prompt lines) or text-rewrite fallback that creates a new simplified PDF.
