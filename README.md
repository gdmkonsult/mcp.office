# mcp.document-edit

Headless MCP server for editing Word documents (DOCX) and PDF files. No UI, no desktop application, no backend required. Designed to be called by AI assistants to create, modify, and inspect documents programmatically.

Deployed as part of the GDM AI platform at `https://ai.gdm.se/api/v1/mcp/document-edit`.

---

## Tools

### `edit_document_from_prompt`

Natural-language editing for DOCX and PDF files. Parses a prompt into discrete actions and applies them to the document.

| Parameter | Type | Required | Description |
|-----------|------|:--------:|-------------|
| `input_path` | `string` | ✅ | Path to input file (`.docx` or `.pdf`) |
| `output_path` | `string` | ✅ | Path to write the edited file |
| `prompt` | `string` | ✅ | Natural-language edit instructions (see syntax below) |

**DOCX prompt syntax:**
```
replace "old text" with "new text"
append paragraph "New content at the end"
prepend paragraph "Content at the top"
delete paragraph containing "text to remove"
```

**PDF form filling syntax** (one field per line):
```
First Name=John
Last Name=Smith
Email=john@example.com
Date=2026-06-11
```

For PDFs without form fields, text is extracted and rewritten using reportlab. Note: original layout is not preserved in this fallback mode.

---

### `format_text`

Apply formatting to a paragraph in a DOCX file. All formatting parameters are optional — only specified ones are applied.

| Parameter | Type | Required | Description |
|-----------|------|:--------:|-------------|
| `input_path` | `string` | ✅ | Path to input DOCX file |
| `output_path` | `string` | ✅ | Path to write the formatted file |
| `paragraph_index` | `integer` | ✅ | Index of paragraph to format (0-based) |
| `bold` | `boolean` | | Set bold on/off |
| `italic` | `boolean` | | Set italic on/off |
| `font_size` | `integer` | | Font size in points |
| `color` | `string` | | Hex color code, e.g. `"#FF0000"` |
| `alignment` | `string` | | `"left"`, `"center"`, `"right"`, or `"justify"` |

---

### `add_list`

Append a bulleted or numbered list to a DOCX document.

| Parameter | Type | Required | Description |
|-----------|------|:--------:|-------------|
| `input_path` | `string` | ✅ | Path to input DOCX file |
| `output_path` | `string` | ✅ | Path to write the output file |
| `items` | `list[string]` | ✅ | List of items to add |
| `list_type` | `string` | | `"bullet"` (default) or `"number"` |

---

### `apply_list_formatting`

Apply list style to existing paragraphs by index.

| Parameter | Type | Required | Description |
|-----------|------|:--------:|-------------|
| `input_path` | `string` | ✅ | Path to input DOCX file |
| `output_path` | `string` | ✅ | Path to write the output file |
| `paragraph_indices` | `list[integer]` | ✅ | Paragraph indices to convert (0-based) |
| `list_type` | `string` | | `"bullet"` (default) or `"number"` |

---

### `create_table`

Add a new table to a DOCX document, optionally pre-filled with data.

| Parameter | Type | Required | Description |
|-----------|------|:--------:|-------------|
| `input_path` | `string` | ✅ | Path to input DOCX file |
| `output_path` | `string` | ✅ | Path to write the output file |
| `rows` | `integer` | ✅ | Number of rows |
| `cols` | `integer` | ✅ | Number of columns |
| `data` | `list[list[string]]` | | 2D array of cell values |

---

### `edit_table_cell`

Set the text content of a single table cell.

| Parameter | Type | Required | Description |
|-----------|------|:--------:|-------------|
| `input_path` | `string` | ✅ | Path to input DOCX file |
| `output_path` | `string` | ✅ | Path to write the output file |
| `table_index` | `integer` | ✅ | Index of the table (0-based) |
| `row` | `integer` | ✅ | Row index (0-based) |
| `col` | `integer` | ✅ | Column index (0-based) |
| `text` | `string` | ✅ | New cell text |

---

### `add_table_row`

Append a row to an existing table.

| Parameter | Type | Required | Description |
|-----------|------|:--------:|-------------|
| `input_path` | `string` | ✅ | Path to input DOCX file |
| `output_path` | `string` | ✅ | Path to write the output file |
| `table_index` | `integer` | ✅ | Index of the table (0-based) |
| `row_data` | `list[string]` | | Cell values for the new row |

---

### `delete_table_row`

Delete a row from an existing table.

| Parameter | Type | Required | Description |
|-----------|------|:--------:|-------------|
| `input_path` | `string` | ✅ | Path to input DOCX file |
| `output_path` | `string` | ✅ | Path to write the output file |
| `table_index` | `integer` | ✅ | Index of the table (0-based) |
| `row` | `integer` | ✅ | Row index to delete (0-based) |

---

### `get_document_structure`

Inspect a DOCX file — returns structure and metadata without modifying the file.

| Parameter | Type | Required | Description |
|-----------|------|:--------:|-------------|
| `input_path` | `string` | ✅ | Path to DOCX file |

**Response includes:**
- `paragraph_count`, `table_count`
- `paragraphs`: list of `{ index, text (first 100 chars), style }`
- `tables`: list of `{ index, rows, cols }`
- `core_properties`: `{ title, author, subject }`

---

### `search_text`

Case-insensitive search across all paragraphs and table cells in a DOCX file.

| Parameter | Type | Required | Description |
|-----------|------|:--------:|-------------|
| `input_path` | `string` | ✅ | Path to DOCX file |
| `search_term` | `string` | ✅ | Text to search for |

Returns matches with type (`paragraph` or `table`), location (index or table/row/col), and a text excerpt.

---

### `add_header`

Set header text on the first section of a DOCX document.

| Parameter | Type | Required | Description |
|-----------|------|:--------:|-------------|
| `input_path` | `string` | ✅ | Path to input DOCX file |
| `output_path` | `string` | ✅ | Path to write the output file |
| `text` | `string` | ✅ | Header text |

---

### `add_footer`

Set footer text on the first section of a DOCX document.

| Parameter | Type | Required | Description |
|-----------|------|:--------:|-------------|
| `input_path` | `string` | ✅ | Path to input DOCX file |
| `output_path` | `string` | ✅ | Path to write the output file |
| `text` | `string` | ✅ | Footer text |

---

### `describe_capabilities`

Returns server metadata and a list of all available tools. No parameters.

---

## Response Format

All tools return a JSON object.

**Success:**
```json
{
  "ok": true,
  "output_path": "/path/to/output.docx"
}
```

**Error:**
```json
{
  "ok": false,
  "error": "Paragraph index 12 out of range"
}
```

All indices (paragraphs, tables, rows, columns) are **0-based**.

---

## Supported File Types

| Format | Support |
|--------|---------|
| `.docx` | Full — all tools |
| `.pdf` (form) | Form field filling via `edit_document_from_prompt` |
| `.pdf` (plain) | Text extraction and rewrite via `edit_document_from_prompt` (layout not preserved) |

---

## Usage Examples

### Via MCP gateway (curl)

```bash
# List available tools
curl https://ai.gdm.se/api/v1/mcp/document-edit \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer YOUR_API_KEY" \
  -d '{
    "jsonrpc": "2.0",
    "id": 1,
    "method": "tools/list",
    "params": {}
  }'

# Inspect a document
curl https://ai.gdm.se/api/v1/mcp/document-edit \
  -H "Content-Type: application/json" \
  -H "Authorization: Bearer YOUR_API_KEY" \
  -d '{
    "jsonrpc": "2.0",
    "id": 2,
    "method": "tools/call",
    "params": {
      "name": "get_document_structure",
      "arguments": {
        "input_path": "/data/contract.docx"
      }
    }
  }'
```

### Via Python (MCP SDK)

```python
import asyncio
from mcp import ClientSession
from mcp.client.streamable_http import streamablehttp_client

async def main():
    async with streamablehttp_client(
        "https://ai.gdm.se/api/v1/mcp/document-edit",
        headers={"Authorization": "Bearer YOUR_API_KEY"},
    ) as (read, write, _):
        async with ClientSession(read, write) as session:
            await session.initialize()

            # Edit a document using a prompt
            result = await session.call_tool(
                "edit_document_from_prompt",
                arguments={
                    "input_path": "/data/report.docx",
                    "output_path": "/data/report_edited.docx",
                    "prompt": 'replace "Draft" with "Final"\nappend paragraph "Approved by: Jane Doe"',
                },
            )
            print(result)

            # Create a table
            result = await session.call_tool(
                "create_table",
                arguments={
                    "input_path": "/data/report_edited.docx",
                    "output_path": "/data/report_table.docx",
                    "rows": 3,
                    "cols": 2,
                    "data": [["Name", "Value"], ["Item A", "100"], ["Item B", "200"]],
                },
            )
            print(result)

asyncio.run(main())
```

---

## Local Development

```bash
pip install -r requirements.txt
python server.py
```

Health checks:
- `GET /healthz` → `{"status": "ok"}`
- `GET /readyz` → `{"status": "ready"}`

MCP endpoint: `http://localhost:8000/mcp`

## Docker

```bash
docker build -t ghcr.io/gdmkonsult/mcp.document-edit:main .
docker run --rm -p 8000:8000 ghcr.io/gdmkonsult/mcp.document-edit:main
```

---

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| `mcp` | 1.27.0 | FastMCP server framework |
| `python-docx` | 1.2.0 | DOCX reading and writing |
| `pypdf` | 6.0.0 | PDF form field reading and filling |
| `reportlab` | 4.4.4 | PDF text rewriting fallback |
| `starlette` | 1.2.1 | HTTP server for health probes and MCP endpoint |
| `uvicorn` | 0.45.0 | ASGI server |

