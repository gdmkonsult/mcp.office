import contextlib
import os
import re
from pathlib import Path
from typing import Any, Optional

from mcp.server.fastmcp import FastMCP
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from starlette.applications import Starlette
from starlette.responses import JSONResponse
from starlette.routing import Mount, Route
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

mcp = FastMCP(
    "DocumentEdit",
    host="0.0.0.0",
    stateless_http=True,
    json_response=True,
)


def _ensure_parent(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)


def _parse_prompt_actions(prompt: str) -> list[dict[str, str]]:
    actions: list[dict[str, str]] = []

    patterns = [
        (r"replace\s+\"(?P<old>.+?)\"\s+with\s+\"(?P<new>.+?)\"", "replace"),
        (r"append\s+paragraph\s+\"(?P<text>.+?)\"", "append_paragraph"),
        (r"prepend\s+paragraph\s+\"(?P<text>.+?)\"", "prepend_paragraph"),
        (r"delete\s+paragraph\s+containing\s+\"(?P<contains>.+?)\"", "delete_paragraph"),
    ]

    for line in [item.strip() for item in prompt.splitlines() if item.strip()]:
        matched = False
        for pattern, kind in patterns:
            m = re.search(pattern, line, flags=re.IGNORECASE)
            if m:
                action = {"type": kind}
                action.update({k: v for k, v in m.groupdict().items() if v is not None})
                actions.append(action)
                matched = True
                break
        if not matched and "=" in line:
            # Allow PDF form filling with "Field Name=Value" lines.
            name, value = line.split("=", 1)
            if name.strip():
                actions.append(
                    {
                        "type": "set_form_field",
                        "field": name.strip(),
                        "value": value.strip(),
                    }
                )

    return actions


def _edit_docx_from_actions(input_path: Path, output_path: Path, actions: list[dict[str, str]]) -> dict[str, Any]:
    doc = Document(str(input_path))
    summary: list[str] = []

    for action in actions:
        kind = action.get("type")
        if kind == "replace":
            old = action.get("old", "")
            new = action.get("new", "")
            replaced = 0
            for paragraph in doc.paragraphs:
                if old and old in paragraph.text:
                    paragraph.text = paragraph.text.replace(old, new)
                    replaced += 1
            summary.append(f"replace:{old}->{new} paragraphs:{replaced}")

        elif kind == "append_paragraph":
            text = action.get("text", "")
            doc.add_paragraph(text)
            summary.append("append_paragraph")

        elif kind == "prepend_paragraph":
            text = action.get("text", "")
            first = doc.paragraphs[0] if doc.paragraphs else doc.add_paragraph("")
            inserted = first.insert_paragraph_before(text)
            inserted.style = first.style
            summary.append("prepend_paragraph")

        elif kind == "delete_paragraph":
            contains = action.get("contains", "")
            deleted = 0
            for paragraph in list(doc.paragraphs):
                if contains and contains in paragraph.text:
                    element = paragraph._element
                    parent = element.getparent()
                    if parent is not None:
                        parent.remove(element)
                        deleted += 1
            summary.append(f"delete_paragraph:{contains} count:{deleted}")

    _ensure_parent(output_path)
    doc.save(str(output_path))

    return {
        "ok": True,
        "output_path": str(output_path),
        "summary": summary,
    }


def _edit_pdf_form_from_actions(input_path: Path, output_path: Path, actions: list[dict[str, str]]) -> dict[str, Any]:
    field_updates = {
        action["field"]: action["value"]
        for action in actions
        if action.get("type") == "set_form_field"
    }
    if not field_updates:
        return {
            "ok": False,
            "error": "No PDF form field actions found. Use lines like FieldName=Value in prompt.",
        }

    reader = PdfReader(str(input_path))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    try:
        writer.update_page_form_field_values(writer.pages[0], field_updates)
    except Exception as exc:
        return {
            "ok": False,
            "error": f"Failed to apply PDF form values: {exc}",
        }

    _ensure_parent(output_path)
    with open(output_path, "wb") as f:
        writer.write(f)

    return {
        "ok": True,
        "output_path": str(output_path),
        "updated_fields": list(field_updates.keys()),
    }


def _rewrite_pdf_text_from_actions(input_path: Path, output_path: Path, actions: list[dict[str, str]]) -> dict[str, Any]:
    reader = PdfReader(str(input_path))
    text = "\n".join((page.extract_text() or "") for page in reader.pages)
    lines = [line for line in text.splitlines() if line.strip()]

    summary: list[str] = []
    for action in actions:
        kind = action.get("type")
        if kind == "replace":
            old = action.get("old", "")
            new = action.get("new", "")
            lines = [line.replace(old, new) for line in lines]
            summary.append(f"replace:{old}->{new}")
        elif kind == "delete_paragraph":
            contains = action.get("contains", "")
            before = len(lines)
            lines = [line for line in lines if contains not in line]
            summary.append(f"delete_lines:{before - len(lines)}")
        elif kind in {"append_paragraph", "prepend_paragraph"}:
            text_value = action.get("text", "")
            if kind == "append_paragraph":
                lines.append(text_value)
            else:
                lines.insert(0, text_value)
            summary.append(kind)

    _ensure_parent(output_path)
    c = canvas.Canvas(str(output_path), pagesize=A4)
    width, height = A4
    y = height - 50
    for line in lines:
        if y < 50:
            c.showPage()
            y = height - 50
        c.drawString(40, y, line[:150])
        y -= 16
    c.save()

    return {
        "ok": True,
        "output_path": str(output_path),
        "summary": summary,
        "warning": "Original PDF layout is not preserved in text-rewrite fallback mode.",
    }


@mcp.tool()
async def describe_capabilities() -> dict:
    """Describe what this MCP supports for document editing."""
    return {
        "server": "mcp.document-edit",
        "capabilities": {
            "edit_docx": "Headless editing of Word documents (.docx) using python-docx",
            "format_text": "Format text in DOCX (bold, italic, color, size, alignment)",
            "lists": "Create and format bullet/numbered lists",
            "tables": "Create tables, edit cells, add/delete rows",
            "inspect": "Get document structure and metadata",
            "search": "Search for text in documents",
            "headers_footers": "Add headers and footers",
            "edit_pdf_forms": "PDF form field filling using PyPDF",
            "edit_pdf_text": "PDF text rewriting using reportlab (fallback for non-form PDFs)",
        },
        "tools": [
            "edit_document_from_prompt",
            "format_text",
            "add_list",
            "apply_list_formatting",
            "create_table",
            "edit_table_cell",
            "add_table_row",
            "delete_table_row",
            "get_document_structure",
            "search_text",
            "add_header",
            "add_footer",
        ],
        "headless_only": True,
        "backend_independent": True,
    }


@mcp.tool()
async def edit_document_from_prompt(input_path: str, output_path: str, prompt: str) -> dict:
    """Headless edit of DOCX or PDF (forms) based on a natural-language prompt.

    Prompt syntax currently supports:
    - replace "old" with "new"
    - append paragraph "text"
    - prepend paragraph "text"
    - delete paragraph containing "text"
    - PDF form filling with lines like: FieldName=Value
    """
    src = Path(input_path)
    dst = Path(output_path)

    if not src.exists():
        return {"ok": False, "error": f"Input file does not exist: {input_path}"}

    actions = _parse_prompt_actions(prompt)
    if not actions:
        return {
            "ok": False,
            "error": "No supported actions parsed from prompt.",
            "hint": "For DOCX: replace/append/prepend/delete paragraph commands. For PDF forms: Field=Value lines.",
        }

    suffix = src.suffix.lower()
    if suffix == ".docx":
        result = _edit_docx_from_actions(src, dst, actions)
        result["document_type"] = "docx"
        result["actions"] = actions
        return result

    if suffix == ".pdf":
        has_form_actions = any(action.get("type") == "set_form_field" for action in actions)
        if has_form_actions:
            result = _edit_pdf_form_from_actions(src, dst, actions)
        else:
            result = _rewrite_pdf_text_from_actions(src, dst, actions)

        if not result.get("ok") and "No /AcroForm" in str(result.get("error", "")):
            result = _rewrite_pdf_text_from_actions(src, dst, actions)
        result["document_type"] = "pdf"
        result["actions"] = actions
        return result

    return {
        "ok": False,
        "error": f"Unsupported extension: {suffix}",
        "supported": [".docx", ".pdf"],
    }


# ===== TIER 1 TOOLS =====


@mcp.tool()
async def format_text(
    input_path: str,
    output_path: str,
    paragraph_index: int,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    font_size: Optional[int] = None,
    color: Optional[str] = None,
    alignment: Optional[str] = None,
) -> dict:
    """Format text in a specific paragraph (DOCX only).
    
    Args:
        input_path: Path to input DOCX file
        output_path: Path to output DOCX file
        paragraph_index: Index of paragraph to format (0-based)
        bold: Set bold (True/False)
        italic: Set italic (True/False)
        font_size: Font size in points
        color: Hex color code (e.g. "#FF0000" for red)
        alignment: "left", "center", "right", or "justify"
    """
    src = Path(input_path)
    dst = Path(output_path)

    if not src.exists():
        return {"ok": False, "error": f"Input file does not exist: {input_path}"}

    if src.suffix.lower() != ".docx":
        return {"ok": False, "error": "format_text only supports DOCX files"}

    try:
        doc = Document(str(src))

        if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
            return {"ok": False, "error": f"Paragraph index {paragraph_index} out of range"}

        para = doc.paragraphs[paragraph_index]

        # Apply formatting to all runs in the paragraph
        for run in para.runs:
            if bold is not None:
                run.bold = bold
            if italic is not None:
                run.italic = italic
            if font_size is not None:
                run.font.size = Pt(font_size)
            if color is not None:
                try:
                    rgb = RGBColor(
                        int(color[1:3], 16),
                        int(color[3:5], 16),
                        int(color[5:7], 16),
                    )
                    run.font.color.rgb = rgb
                except (ValueError, IndexError):
                    return {"ok": False, "error": f"Invalid color format: {color}"}

        # Apply alignment to paragraph
        if alignment is not None:
            align_map = {
                "left": WD_PARAGRAPH_ALIGNMENT.LEFT,
                "center": WD_PARAGRAPH_ALIGNMENT.CENTER,
                "right": WD_PARAGRAPH_ALIGNMENT.RIGHT,
                "justify": WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
            }
            if alignment.lower() not in align_map:
                return {"ok": False, "error": f"Invalid alignment: {alignment}"}
            para.alignment = align_map[alignment.lower()]

        _ensure_parent(dst)
        doc.save(str(dst))

        return {
            "ok": True,
            "output_path": str(dst),
            "paragraph_index": paragraph_index,
            "formatting_applied": {
                "bold": bold,
                "italic": italic,
                "font_size": font_size,
                "color": color,
                "alignment": alignment,
            },
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


@mcp.tool()
async def add_list(
    input_path: str,
    output_path: str,
    items: list[str],
    list_type: str = "bullet",
) -> dict:
    """Add a bulleted or numbered list to DOCX document.
    
    Args:
        input_path: Path to input DOCX file
        output_path: Path to output DOCX file
        items: List of items to add
        list_type: "bullet" for bullets, "number" for numbered list
    """
    src = Path(input_path)
    dst = Path(output_path)

    if not src.exists():
        return {"ok": False, "error": f"Input file does not exist: {input_path}"}

    if src.suffix.lower() != ".docx":
        return {"ok": False, "error": "add_list only supports DOCX files"}

    try:
        doc = Document(str(src))

        style = "List Bullet" if list_type.lower() == "bullet" else "List Number"

        for item in items:
            doc.add_paragraph(item, style=style)

        _ensure_parent(dst)
        doc.save(str(dst))

        return {
            "ok": True,
            "output_path": str(dst),
            "list_type": list_type,
            "items_added": len(items),
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


@mcp.tool()
async def apply_list_formatting(
    input_path: str,
    output_path: str,
    paragraph_indices: list[int],
    list_type: str = "bullet",
) -> dict:
    """Apply list formatting to existing paragraphs.
    
    Args:
        input_path: Path to input DOCX file
        output_path: Path to output DOCX file
        paragraph_indices: List of paragraph indices to format (0-based)
        list_type: "bullet" for bullets, "number" for numbered list
    """
    src = Path(input_path)
    dst = Path(output_path)

    if not src.exists():
        return {"ok": False, "error": f"Input file does not exist: {input_path}"}

    if src.suffix.lower() != ".docx":
        return {"ok": False, "error": "apply_list_formatting only supports DOCX files"}

    try:
        doc = Document(str(src))
        style = "List Bullet" if list_type.lower() == "bullet" else "List Number"

        formatted_count = 0
        for idx in paragraph_indices:
            if 0 <= idx < len(doc.paragraphs):
                doc.paragraphs[idx].style = style
                formatted_count += 1

        _ensure_parent(dst)
        doc.save(str(dst))

        return {
            "ok": True,
            "output_path": str(dst),
            "list_type": list_type,
            "paragraphs_formatted": formatted_count,
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


@mcp.tool()
async def create_table(
    input_path: str,
    output_path: str,
    rows: int,
    cols: int,
    data: Optional[list[list[str]]] = None,
) -> dict:
    """Create a table in DOCX document.
    
    Args:
        input_path: Path to input DOCX file
        output_path: Path to output DOCX file
        rows: Number of rows
        cols: Number of columns
        data: 2D list of cell values (optional)
    """
    src = Path(input_path)
    dst = Path(output_path)

    if not src.exists():
        return {"ok": False, "error": f"Input file does not exist: {input_path}"}

    if src.suffix.lower() != ".docx":
        return {"ok": False, "error": "create_table only supports DOCX files"}

    try:
        doc = Document(str(src))
        table = doc.add_table(rows=rows, cols=cols)
        table.style = "Table Grid"

        if data:
            for r_idx, row_data in enumerate(data[:rows]):
                for c_idx, cell_value in enumerate(row_data[:cols]):
                    table.rows[r_idx].cells[c_idx].text = str(cell_value)

        _ensure_parent(dst)
        doc.save(str(dst))

        return {
            "ok": True,
            "output_path": str(dst),
            "table_rows": rows,
            "table_cols": cols,
            "data_filled": bool(data),
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


@mcp.tool()
async def edit_table_cell(
    input_path: str,
    output_path: str,
    table_index: int,
    row: int,
    col: int,
    text: str,
) -> dict:
    """Edit a single table cell in DOCX document.
    
    Args:
        input_path: Path to input DOCX file
        output_path: Path to output DOCX file
        table_index: Index of table (0-based)
        row: Row index (0-based)
        col: Column index (0-based)
        text: New cell text
    """
    src = Path(input_path)
    dst = Path(output_path)

    if not src.exists():
        return {"ok": False, "error": f"Input file does not exist: {input_path}"}

    if src.suffix.lower() != ".docx":
        return {"ok": False, "error": "edit_table_cell only supports DOCX files"}

    try:
        doc = Document(str(src))

        if table_index < 0 or table_index >= len(doc.tables):
            return {"ok": False, "error": f"Table index {table_index} out of range"}

        table = doc.tables[table_index]

        if row < 0 or row >= len(table.rows):
            return {"ok": False, "error": f"Row {row} out of range"}

        if col < 0 or col >= len(table.columns):
            return {"ok": False, "error": f"Column {col} out of range"}

        table.rows[row].cells[col].text = text

        _ensure_parent(dst)
        doc.save(str(dst))

        return {
            "ok": True,
            "output_path": str(dst),
            "table": table_index,
            "cell": {"row": row, "col": col},
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


@mcp.tool()
async def add_table_row(
    input_path: str,
    output_path: str,
    table_index: int,
    row_data: Optional[list[str]] = None,
) -> dict:
    """Add a row to a table in DOCX document.
    
    Args:
        input_path: Path to input DOCX file
        output_path: Path to output DOCX file
        table_index: Index of table (0-based)
        row_data: Optional list of cell values for the new row
    """
    src = Path(input_path)
    dst = Path(output_path)

    if not src.exists():
        return {"ok": False, "error": f"Input file does not exist: {input_path}"}

    if src.suffix.lower() != ".docx":
        return {"ok": False, "error": "add_table_row only supports DOCX files"}

    try:
        doc = Document(str(src))

        if table_index < 0 or table_index >= len(doc.tables):
            return {"ok": False, "error": f"Table index {table_index} out of range"}

        table = doc.tables[table_index]
        new_row = table.add_row()

        if row_data:
            for col_idx, cell_value in enumerate(row_data[: len(table.columns)]):
                new_row.cells[col_idx].text = str(cell_value)

        _ensure_parent(dst)
        doc.save(str(dst))

        return {
            "ok": True,
            "output_path": str(dst),
            "table": table_index,
            "row_added": True,
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


@mcp.tool()
async def delete_table_row(
    input_path: str,
    output_path: str,
    table_index: int,
    row: int,
) -> dict:
    """Delete a row from a table in DOCX document.
    
    Args:
        input_path: Path to input DOCX file
        output_path: Path to output DOCX file
        table_index: Index of table (0-based)
        row: Row index to delete (0-based)
    """
    src = Path(input_path)
    dst = Path(output_path)

    if not src.exists():
        return {"ok": False, "error": f"Input file does not exist: {input_path}"}

    if src.suffix.lower() != ".docx":
        return {"ok": False, "error": "delete_table_row only supports DOCX files"}

    try:
        doc = Document(str(src))

        if table_index < 0 or table_index >= len(doc.tables):
            return {"ok": False, "error": f"Table index {table_index} out of range"}

        table = doc.tables[table_index]

        if row < 0 or row >= len(table.rows):
            return {"ok": False, "error": f"Row {row} out of range"}

        tbl = table._tbl
        tr = table.rows[row]._tr
        tbl.remove(tr)

        _ensure_parent(dst)
        doc.save(str(dst))

        return {
            "ok": True,
            "output_path": str(dst),
            "table": table_index,
            "row_deleted": row,
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


@mcp.tool()
async def get_document_structure(input_path: str) -> dict:
    """Get structure and metadata of a DOCX document.
    
    Args:
        input_path: Path to DOCX file
    """
    src = Path(input_path)

    if not src.exists():
        return {"ok": False, "error": f"Input file does not exist: {input_path}"}

    if src.suffix.lower() != ".docx":
        return {"ok": False, "error": "get_document_structure only supports DOCX files"}

    try:
        doc = Document(str(src))

        paragraphs = [
            {"index": i, "text": p.text[:100], "style": p.style.name}
            for i, p in enumerate(doc.paragraphs)
        ]

        tables = [
            {
                "index": i,
                "rows": len(t.rows),
                "cols": len(t.columns),
            }
            for i, t in enumerate(doc.tables)
        ]

        core_props = doc.core_properties

        return {
            "ok": True,
            "file_path": str(src),
            "paragraph_count": len(doc.paragraphs),
            "table_count": len(doc.tables),
            "paragraphs": paragraphs,
            "tables": tables,
            "core_properties": {
                "title": core_props.title,
                "author": core_props.author,
                "subject": core_props.subject,
            },
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


@mcp.tool()
async def search_text(input_path: str, search_term: str) -> dict:
    """Search for text in a DOCX document.
    
    Args:
        input_path: Path to DOCX file
        search_term: Text to search for
    """
    src = Path(input_path)

    if not src.exists():
        return {"ok": False, "error": f"Input file does not exist: {input_path}"}

    if src.suffix.lower() != ".docx":
        return {"ok": False, "error": "search_text only supports DOCX files"}

    try:
        doc = Document(str(src))

        results = []

        # Search in paragraphs
        for p_idx, para in enumerate(doc.paragraphs):
            if search_term.lower() in para.text.lower():
                results.append({
                    "type": "paragraph",
                    "index": p_idx,
                    "text": para.text[:150],
                })

        # Search in tables
        for t_idx, table in enumerate(doc.tables):
            for r_idx, row in enumerate(table.rows):
                for c_idx, cell in enumerate(row.cells):
                    if search_term.lower() in cell.text.lower():
                        results.append({
                            "type": "table",
                            "table_index": t_idx,
                            "row": r_idx,
                            "col": c_idx,
                            "text": cell.text[:150],
                        })

        return {
            "ok": True,
            "search_term": search_term,
            "results_count": len(results),
            "results": results,
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


@mcp.tool()
async def add_header(
    input_path: str,
    output_path: str,
    text: str,
) -> dict:
    """Add header text to a DOCX document.
    
    Args:
        input_path: Path to input DOCX file
        output_path: Path to output DOCX file
        text: Header text to add
    """
    src = Path(input_path)
    dst = Path(output_path)

    if not src.exists():
        return {"ok": False, "error": f"Input file does not exist: {input_path}"}

    if src.suffix.lower() != ".docx":
        return {"ok": False, "error": "add_header only supports DOCX files"}

    try:
        doc = Document(str(src))
        section = doc.sections[0]
        header = section.header
        header_para = header.paragraphs[0]
        header_para.text = text

        _ensure_parent(dst)
        doc.save(str(dst))

        return {
            "ok": True,
            "output_path": str(dst),
            "header_text": text,
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


@mcp.tool()
async def add_footer(
    input_path: str,
    output_path: str,
    text: str,
) -> dict:
    """Add footer text to a DOCX document.
    
    Args:
        input_path: Path to input DOCX file
        output_path: Path to output DOCX file
        text: Footer text to add
    """
    src = Path(input_path)
    dst = Path(output_path)

    if not src.exists():
        return {"ok": False, "error": f"Input file does not exist: {input_path}"}

    if src.suffix.lower() != ".docx":
        return {"ok": False, "error": "add_footer only supports DOCX files"}

    try:
        doc = Document(str(src))
        section = doc.sections[0]
        footer = section.footer
        footer_para = footer.paragraphs[0]
        footer_para.text = text

        _ensure_parent(dst)
        doc.save(str(dst))

        return {
            "ok": True,
            "output_path": str(dst),
            "footer_text": text,
        }
    except Exception as e:
        return {"ok": False, "error": str(e)}


# ===== HTTP HEALTH ENDPOINTS =====


async def healthz(_request):
    return JSONResponse({"status": "ok"})


async def readyz(_request):
    return JSONResponse({"status": "ready"})


@contextlib.asynccontextmanager
async def lifespan(app: Starlette):
    async with mcp.session_manager.run():
        yield


app = Starlette(
    routes=[
        Route("/healthz", healthz),
        Route("/readyz", readyz),
        Mount("/", app=mcp.streamable_http_app()),
    ],
    lifespan=lifespan,
)

if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)
