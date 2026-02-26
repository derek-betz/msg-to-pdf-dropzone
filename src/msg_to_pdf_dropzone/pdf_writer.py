from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from html import escape
from pathlib import Path

from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

from .models import EmailRecord

BLANK_LINE_PATTERN = re.compile(r"\n\s*\n")
SCRIPT_TAG_PATTERN = re.compile(r"(?is)<script[^>]*>.*?</script>")
HTML_BODY_PATTERN = re.compile(r"(?is)<body[^>]*>(?P<body>.*)</body>")


def _format_body_blocks(body: str) -> list[str]:
    normalized = (body or "").replace("\r\n", "\n").replace("\r", "\n").strip()
    if not normalized:
        return ["(No body text.)"]
    return [block.strip() for block in BLANK_LINE_PATTERN.split(normalized) if block.strip()]


def _as_paragraph_text(value: str) -> str:
    return escape(value).replace("\n", "<br/>")


def _format_sent_value(sent_at: datetime) -> str:
    try:
        local_time = sent_at.astimezone()
    except Exception:
        local_time = sent_at
    return local_time.strftime("%Y-%m-%d %H:%M:%S %Z").strip()


def _extract_body_html_fragment(html_body: str) -> str:
    cleaned = SCRIPT_TAG_PATTERN.sub("", html_body or "").strip()
    if not cleaned:
        return ""
    match = HTML_BODY_PATTERN.search(cleaned)
    if match:
        return match.group("body").strip()
    return cleaned


def _build_attachment_list_html(attachment_names: list[str]) -> str:
    if not attachment_names:
        return ""
    items = "".join(f"<li>{escape(name)}</li>" for name in attachment_names)
    return f"""
    <section class="section attachments">
      <h3>Attachments</h3>
      <ul>{items}</ul>
    </section>
    """


def _build_body_fragment(record: EmailRecord) -> str:
    html_fragment = _extract_body_html_fragment(record.html_body)
    if html_fragment:
        return html_fragment
    plain = escape(record.body or "(No body text.)")
    return f'<pre class="plain-body">{plain}</pre>'


def build_email_html_document(record: EmailRecord) -> str:
    body_fragment = _build_body_fragment(record)
    attachments_html = _build_attachment_list_html(record.attachment_names)
    sent_value = escape(_format_sent_value(record.sent_at))

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <title>{escape(record.subject or "No Subject")}</title>
  <style>
    :root {{
      color-scheme: light;
    }}
    body {{
      margin: 0;
      padding: 24px;
      background: #ffffff;
      color: #111827;
      font-family: Calibri, "Segoe UI", Arial, sans-serif;
      font-size: 11pt;
      line-height: 1.35;
    }}
    .meta {{
      border: 1px solid #d1d5db;
      border-radius: 6px;
      margin-bottom: 16px;
      overflow: hidden;
    }}
    .meta-row {{
      display: table;
      width: 100%;
      border-bottom: 1px solid #e5e7eb;
    }}
    .meta-row:last-child {{
      border-bottom: none;
    }}
    .meta-label {{
      display: table-cell;
      width: 86px;
      padding: 8px 10px;
      background: #f3f4f6;
      font-weight: 700;
      vertical-align: top;
    }}
    .meta-value {{
      display: table-cell;
      padding: 8px 10px;
      vertical-align: top;
      word-break: break-word;
    }}
    .section {{
      margin-top: 16px;
    }}
    .section h3 {{
      margin: 0 0 8px;
      font-size: 11pt;
      color: #1f2937;
    }}
    .attachments ul {{
      margin: 0;
      padding-left: 20px;
    }}
    .body {{
      border-top: 1px solid #d1d5db;
      padding-top: 14px;
    }}
    .plain-body {{
      margin: 0;
      white-space: pre-wrap;
      word-break: break-word;
      font-family: Calibri, "Segoe UI", Arial, sans-serif;
    }}
    img {{
      max-width: 100%;
      height: auto;
    }}
    table {{
      max-width: 100%;
      border-collapse: collapse;
    }}
  </style>
</head>
<body>
  <section class="meta">
    <div class="meta-row">
      <div class="meta-label">Subject</div>
      <div class="meta-value">{escape(record.subject or "No Subject")}</div>
    </div>
    <div class="meta-row">
      <div class="meta-label">Sent</div>
      <div class="meta-value">{sent_value}</div>
    </div>
    <div class="meta-row">
      <div class="meta-label">From</div>
      <div class="meta-value">{escape(record.sender or "(empty)")}</div>
    </div>
    <div class="meta-row">
      <div class="meta-label">To</div>
      <div class="meta-value">{escape(record.to or "(empty)")}</div>
    </div>
    <div class="meta-row">
      <div class="meta-label">Cc</div>
      <div class="meta-value">{escape(record.cc or "(empty)")}</div>
    </div>
  </section>
  {attachments_html}
  <section class="section body">
    {body_fragment}
  </section>
</body>
</html>
"""


def _find_edge_executable() -> Path | None:
    candidates = [
        Path("C:/Program Files/Microsoft/Edge/Application/msedge.exe"),
        Path("C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def _print_web_document_via_edge(input_path: Path, output_path: Path) -> bool:
    if os.name != "nt":
        return False

    edge_path = _find_edge_executable()
    if edge_path is None:
        return False

    uri = "file:///" + str(input_path.resolve()).replace("\\", "/")
    command = [
        str(edge_path),
        "--headless",
        "--disable-gpu",
        "--run-all-compositor-stages-before-draw",
        "--virtual-time-budget=10000",
        f"--print-to-pdf={output_path}",
        uri,
    ]
    try:
        result = subprocess.run(command, capture_output=True, text=True, timeout=120, check=False)
        if result.returncode != 0:
            return False
        return output_path.exists() and output_path.stat().st_size > 0
    except Exception:
        return False


def _try_write_pdf_via_outlook_and_edge(msg_path: Path, output_path: Path) -> bool:
    if os.name != "nt":
        return False

    if os.environ.get("MSG_TO_PDF_DISABLE_OUTLOOK_EDGE", "").strip() == "1":
        return False

    if not msg_path.exists():
        return False

    temp_dir = Path(tempfile.mkdtemp(prefix="msg-to-pdf-mht-"))
    mht_path = temp_dir / "email.mht"
    try:
        command = [
            sys.executable,
            "-m",
            "msg_to_pdf_dropzone.outlook_mhtml_worker",
            str(msg_path),
            str(mht_path),
        ]
        result = subprocess.run(command, capture_output=True, text=True, timeout=120, check=False)
        if result.returncode != 0:
            return False
        return _print_web_document_via_edge(mht_path, output_path)
    except Exception:
        return False
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def _try_write_pdf_via_edge_html(html_document: str, output_path: Path) -> bool:
    if os.name != "nt":
        return False

    temp_dir = Path(tempfile.mkdtemp(prefix="msg-to-pdf-html-"))
    html_path = temp_dir / "email.html"
    html_path.write_text(html_document, encoding="utf-8")
    try:
        return _print_web_document_via_edge(html_path, output_path)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def _write_pdf_via_reportlab(record: EmailRecord, output_path: Path) -> None:
    styles = getSampleStyleSheet()
    normal_style = styles["Normal"]
    body_style = ParagraphStyle(
        "EmailBody",
        parent=normal_style,
        leading=14,
        fontSize=10,
    )

    document = SimpleDocTemplate(
        str(output_path),
        pagesize=LETTER,
        leftMargin=50,
        rightMargin=50,
        topMargin=50,
        bottomMargin=50,
    )

    story = [
        Paragraph("<b>Email Export</b>", styles["Title"]),
        Spacer(1, 14),
        Paragraph(f"<b>Subject:</b> {_as_paragraph_text(record.subject or 'No Subject')}", normal_style),
        Spacer(1, 6),
        Paragraph(f"<b>Sent:</b> {_as_paragraph_text(_format_sent_value(record.sent_at))}", normal_style),
        Spacer(1, 6),
        Paragraph(f"<b>From:</b> {_as_paragraph_text(record.sender or '(empty)')}", normal_style),
        Spacer(1, 6),
        Paragraph(f"<b>To:</b> {_as_paragraph_text(record.to or '(empty)')}", normal_style),
        Spacer(1, 6),
        Paragraph(f"<b>Cc:</b> {_as_paragraph_text(record.cc or '(empty)')}", normal_style),
        Spacer(1, 12),
    ]

    if record.attachment_names:
        story.append(Paragraph("<b>Attachments:</b>", normal_style))
        story.append(Spacer(1, 4))
        for attachment_name in record.attachment_names:
            story.append(Paragraph(f"- {_as_paragraph_text(attachment_name)}", body_style))
            story.append(Spacer(1, 2))
        story.append(Spacer(1, 10))

    story.append(Paragraph("<b>Body</b>", styles["Heading3"]))
    story.append(Spacer(1, 6))
    for block in _format_body_blocks(record.body):
        story.append(Paragraph(_as_paragraph_text(block), body_style))
        story.append(Spacer(1, 10))

    document.build(story)


def write_email_pdf(record: EmailRecord, output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if _try_write_pdf_via_outlook_and_edge(record.source_path, output_path):
        return output_path

    html_document = build_email_html_document(record)
    if _try_write_pdf_via_edge_html(html_document, output_path):
        return output_path

    _write_pdf_via_reportlab(record, output_path)
    return output_path
