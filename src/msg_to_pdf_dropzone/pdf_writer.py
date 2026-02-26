from __future__ import annotations

import re
from pathlib import Path
from xml.sax.saxutils import escape

from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

from .models import EmailRecord

BLANK_LINE_PATTERN = re.compile(r"\n\s*\n")


def _format_body_blocks(body: str) -> list[str]:
    normalized = (body or "").replace("\r\n", "\n").replace("\r", "\n").strip()
    if not normalized:
        return ["(No body text.)"]
    return [block.strip() for block in BLANK_LINE_PATTERN.split(normalized) if block.strip()]


def _as_paragraph_text(value: str) -> str:
    return escape(value).replace("\n", "<br/>")


def write_email_pdf(record: EmailRecord, output_path: Path) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)

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
        Paragraph(f"<b>Sent:</b> {_as_paragraph_text(record.sent_at.isoformat())}", normal_style),
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
    return output_path
