import contextlib
import hashlib
import os
import re
import time
from pathlib import Path
from typing import Any

import httpx
import jwt
from mcp.server.fastmcp import FastMCP
from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from starlette.applications import Starlette
from starlette.responses import JSONResponse
from starlette.routing import Mount, Route
from docx import Document

mcp = FastMCP(
    "Office",
    host="0.0.0.0",
    stateless_http=True,
    json_response=True,
)


def _base_url() -> str:
    return os.getenv("EURO_OFFICE_BASE_URL", "http://localhost:8080").rstrip("/")


def _timeout_seconds() -> float:
    return float(os.getenv("EURO_OFFICE_TIMEOUT_SECONDS", "20"))


def _command_path() -> str:
    return os.getenv("EURO_OFFICE_COMMAND_PATH", "/coauthoring/CommandService.ashx")


def _convert_path() -> str:
    return os.getenv("EURO_OFFICE_CONVERT_PATH", "/ConvertService.ashx")


def _safe_path(path: str) -> str:
    if not path.startswith("/"):
        return f"/{path}"
    return path


def _jwt_header_name() -> str:
    return os.getenv("EURO_OFFICE_JWT_HEADER", "Authorization")


def _jwt_prefix() -> str:
    return os.getenv("EURO_OFFICE_JWT_PREFIX", "Bearer")


def _jwt_secret() -> str | None:
    return os.getenv("EURO_OFFICE_JWT_SECRET")


def _build_token(payload: dict[str, Any]) -> str | None:
    secret = _jwt_secret()
    if not secret:
        return None

    now = int(time.time())
    envelope = {
        "iat": now,
        "exp": now + 300,
        "payload": payload,
    }
    return jwt.encode(envelope, secret, algorithm="HS256")


def _auth_headers(payload: dict[str, Any] | None = None) -> dict[str, str]:
    headers = {"Content-Type": "application/json"}
    token = _build_token(payload or {})

    if token:
        prefix = _jwt_prefix().strip()
        value = f"{prefix} {token}" if prefix else token
        headers[_jwt_header_name()] = value

    return headers


async def _post_json(path: str, body: dict[str, Any]) -> dict[str, Any]:
    url = f"{_base_url()}{_safe_path(path)}"
    headers = _auth_headers(body)

    async with httpx.AsyncClient(timeout=_timeout_seconds()) as client:
        response = await client.post(url, json=body, headers=headers)
        response.raise_for_status()
        return response.json()


async def _get(path: str) -> dict[str, Any]:
    url = f"{_base_url()}{_safe_path(path)}"
    async with httpx.AsyncClient(timeout=_timeout_seconds()) as client:
        response = await client.get(url)
        response.raise_for_status()
        return {"status_code": response.status_code, "body": response.text}


def _extension(filename: str) -> str:
    parts = filename.rsplit(".", 1)
    if len(parts) < 2:
        return "docx"
    return parts[1].lower()


def _default_key(document_url: str, callback_url: str, filename: str) -> str:
    source = f"{document_url}|{callback_url}|{filename}"
    return hashlib.sha256(source.encode("utf-8")).hexdigest()[:32]


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
async def check_health() -> dict:
    """Check Euro-Office DocumentServer health endpoint."""
    try:
        result = await _get("/healthcheck")
        return {
            "ok": True,
            "base_url": _base_url(),
            "healthcheck": result,
        }
    except Exception as exc:
        return {
            "ok": False,
            "base_url": _base_url(),
            "error": str(exc),
        }


@mcp.tool()
async def build_editor_config(
    document_url: str,
    callback_url: str,
    filename: str,
    document_type: str = "word",
    mode: str = "edit",
    user_id: str = "mcp-user",
    user_name: str = "MCP User",
    autosave: bool = True,
    can_download: bool = True,
    can_print: bool = True,
    key: str | None = None,
    lang: str = "en",
) -> dict:
    """Build a Euro-Office editor config for Word, PDF, and other docs.

    Args:
        document_url: Public URL the editor can download from.
        callback_url: Your app callback URL for save/forcesave events.
        filename: Document filename including extension.
        document_type: editor type: word, cell, slide, or pdf.
        mode: edit or view.
        user_id: Stable user ID.
        user_name: Display name.
        autosave: Enable auto-save in editor.
        can_download: Allow downloading from editor UI.
        can_print: Allow printing from editor UI.
        key: Optional stable doc key. Auto-generated if omitted.
        lang: Editor language code, for example en or sv.
    """
    document_type = document_type.lower().strip()
    if document_type not in {"word", "cell", "slide", "pdf"}:
        return {"error": f"Unsupported document_type: {document_type}"}

    editor_mode = mode.lower().strip()
    if editor_mode not in {"edit", "view"}:
        return {"error": f"Unsupported mode: {mode}"}

    final_key = key or _default_key(document_url, callback_url, filename)

    config = {
        "document": {
            "fileType": _extension(filename),
            "key": final_key,
            "title": filename,
            "url": document_url,
            "permissions": {
                "edit": editor_mode == "edit",
                "download": can_download,
                "print": can_print,
            },
        },
        "documentType": document_type,
        "editorConfig": {
            "callbackUrl": callback_url,
            "lang": lang,
            "mode": editor_mode,
            "user": {
                "id": user_id,
                "name": user_name,
            },
            "customization": {
                "autosave": autosave,
            },
        },
    }

    return {
        "base_url": _base_url(),
        "editor_config": config,
        "notes": [
            "Use this payload in your integration frontend when opening Euro-Office editor.",
            "Keep document.key stable for the same logical document revision to preserve collaboration state.",
        ],
    }


@mcp.tool()
async def convert_document(
    file_url: str,
    file_type: str,
    output_type: str,
    title: str = "document",
    key: str | None = None,
    lang: str = "en",
) -> dict:
    """Convert a document using Euro-Office conversion endpoint.

    Args:
        file_url: Source file URL reachable from DocumentServer.
        file_type: Source extension without dot, for example docx or pdf.
        output_type: Target extension without dot, for example pdf or docx.
        title: Optional title used by conversion service.
        key: Optional conversion key. Auto-generated if omitted.
        lang: Language code for conversion context.
    """
    final_key = key or hashlib.sha256(f"{file_url}|{file_type}|{output_type}".encode("utf-8")).hexdigest()[:32]

    body = {
        "async": False,
        "filetype": file_type.lower().lstrip("."),
        "key": final_key,
        "outputtype": output_type.lower().lstrip("."),
        "title": title,
        "url": file_url,
        "lang": lang,
    }

    try:
        data = await _post_json(_convert_path(), body)
        return {
            "ok": True,
            "request": body,
            "response": data,
            "convert_path": _convert_path(),
        }
    except Exception as exc:
        return {
            "ok": False,
            "request": body,
            "convert_path": _convert_path(),
            "error": str(exc),
        }


@mcp.tool()
async def command_force_save(document_key: str, userdata: str | None = None) -> dict:
    """Trigger a force-save command for an open document key."""
    body: dict[str, Any] = {
        "c": "forcesave",
        "key": document_key,
    }

    if userdata:
        body["userdata"] = userdata

    try:
        data = await _post_json(_command_path(), body)
        return {
            "ok": True,
            "request": body,
            "response": data,
            "command_path": _command_path(),
        }
    except Exception as exc:
        return {
            "ok": False,
            "request": body,
            "command_path": _command_path(),
            "error": str(exc),
        }


@mcp.tool()
async def command_info(document_key: str) -> dict:
    """Get document session info for a document key."""
    body = {
        "c": "info",
        "key": document_key,
    }

    try:
        data = await _post_json(_command_path(), body)
        return {
            "ok": True,
            "request": body,
            "response": data,
            "command_path": _command_path(),
        }
    except Exception as exc:
        return {
            "ok": False,
            "request": body,
            "command_path": _command_path(),
            "error": str(exc),
        }


@mcp.tool()
async def describe_capabilities() -> dict:
    """Describe what this MCP supports for Word and PDF workflows."""
    return {
        "server": "mcp.office",
        "base_url": _base_url(),
        "supports": {
            "word": [
                "build_editor_config(document_type='word')",
                "convert_document(..., output_type='docx'/'pdf'/'odt')",
                "command_force_save",
                "command_info",
                "edit_document_from_prompt(input_path=..., output_path=..., prompt=...)",
            ],
            "pdf": [
                "build_editor_config(document_type='pdf')",
                "convert_document(..., output_type='pdf'/'docx')",
                "command_force_save",
                "command_info",
                "edit_document_from_prompt(input_path=..., output_path=..., prompt='Field=Value')",
            ],
        },
        "combined": True,
        "headless_editing": True,
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
