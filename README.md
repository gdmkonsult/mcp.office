# mcp.document-edit

Headless MCP server for document editing (DOCX and PDF) with no backend required.

## Tier 1 Tools — Comprehensive Document Editing

**Core editing:**
- `edit_document_from_prompt`: non-visual editing for DOCX and PDF using natural language prompts
- `format_text`: bold, italic, color, font size, alignment
- `add_list` / `apply_list_formatting`: bullet and numbered lists

**Tables:**
- `create_table`, `edit_table_cell`, `add_table_row`, `delete_table_row`

**Document inspection & search:**
- `get_document_structure`: metadata and document outline
- `search_text`: find text in paragraphs and tables

**Page elements:**
- `add_header`, `add_footer`

**PDF support:**
- PDF form field filling and text editing

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
docker build -t ghcr.io/gdmkonsult/mcp.document-edit:main .
docker run --rm -p 8000:8000 ghcr.io/gdmkonsult/mcp.document-edit:main
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

