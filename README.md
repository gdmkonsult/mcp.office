# mcp.office

Headless MCP server for document editing:

- Word document editing (DOCX)
- PDF form filling and text editing
- Non-visual, prompt-based operations
- No backend server required

## Tools

- `edit_document_from_prompt`: non-visual editing for DOCX and PDF
- `describe_capabilities`: quick summary of supported operations

## Environment variables

None required. The MCP is fully self-contained.

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
docker run --rm -p 8000:8000 ghcr.io/gdmkonsult/mcp.office:main
```

## Usage

### Edit a Word document

```python
import asyncio
from mcp import ClientSession
from mcp.client.streamable_http import streamablehttp_client

async def main():
    async with streamablehttp_client("http://localhost:8000") as (read, write, _):
        async with ClientSession(read, write) as session:
            await session.initialize()
            result = await session.call_tool("edit_document_from_prompt", {
                "input_path": "/tmp/document.docx",
                "output_path": "/tmp/document_edited.docx",
                "prompt": 'replace "Draft" with "Final"'
            })
            print(result)

asyncio.run(main())
```

### Edit a PDF form

```python
result = await session.call_tool("edit_document_from_prompt", {
    "input_path": "/tmp/form.pdf",
    "output_path": "/tmp/form_filled.pdf",
    "prompt": "FirstName=John\nLastName=Doe"
})
```

## Prompt syntax

### DOCX

- `replace "old" with "new"` — replace text
- `append paragraph "text"` — add paragraph at end
- `prepend paragraph "text"` — add paragraph at start
- `delete paragraph containing "text"` — remove matching line

### PDF

- `FieldName=Value` — fill form fields (one per line)
- If no form fields exist, falls back to text rewrite using above DOCX syntax

## Notes

- DOCX editing uses python-docx
- PDF form editing uses PyPDF
- PDF text rewriting uses reportlab and may lose original formatting
- Files can be local paths or accessed from mounted volumes in Kubernetes

