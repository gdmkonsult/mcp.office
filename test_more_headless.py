from pathlib import Path
import asyncio

from pypdf import PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

import server


async def run_tests() -> None:
    root = Path("test-files")
    root.mkdir(exist_ok=True)

    # 1) Invalid prompt handling
    r1 = await server.edit_document_from_prompt(
        "test-files/headless-source.docx",
        "test-files/out-invalid.docx",
        "please improve this",
    )
    print("T1_OK", r1.get("ok"))
    print("T1_HAS_ERROR", bool(r1.get("error")))

    # 2) Unsupported extension handling
    (root / "dummy.txt").write_text("hello")
    r2 = await server.edit_document_from_prompt(
        "test-files/dummy.txt",
        "test-files/out.txt",
        'replace "a" with "b"',
    )
    print("T2_OK", r2.get("ok"))
    print("T2_UNSUPPORTED", "Unsupported extension" in str(r2.get("error")))

    # 3) PDF text rewrite fallback verification
    src_pdf = root / "plain.pdf"
    out_pdf = root / "plain-edited.pdf"
    c = canvas.Canvas(str(src_pdf), pagesize=A4)
    c.drawString(50, 800, "Acme AB contract for 2026")
    c.drawString(50, 780, "Second line marker")
    c.save()

    prompt = "\n".join(
        [
            'replace "Acme AB" with "Eneo AB"',
            'append paragraph "Approved by legal"',
            'delete paragraph containing "Second line marker"',
        ]
    )
    r3 = await server.edit_document_from_prompt(str(src_pdf), str(out_pdf), prompt)
    print("T3_OK", r3.get("ok"))
    print("T3_WARNING_PRESENT", bool(r3.get("warning")))

    text = "\n".join((p.extract_text() or "") for p in PdfReader(str(out_pdf)).pages)
    print("T3_HAS_ENEO", "Eneo AB" in text)
    print("T3_HAS_APPROVED", "Approved by legal" in text)
    print("T3_HAS_DELETED_LINE", "Second line marker" in text)


if __name__ == "__main__":
    asyncio.run(run_tests())
