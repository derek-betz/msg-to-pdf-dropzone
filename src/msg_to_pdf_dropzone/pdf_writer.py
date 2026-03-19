from __future__ import annotations

from collections.abc import Mapping
import base64
from dataclasses import dataclass, field
import os
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from html import escape, unescape
from pathlib import Path
from time import perf_counter

from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer

from .models import EmailRecord, InlineImageAsset
from .task_events import TaskEventSink, TaskMetaValue, default_task_id_for_path, emit_task_event

BLANK_LINE_PATTERN = re.compile(r"\n\s*\n")
SCRIPT_TAG_PATTERN = re.compile(r"(?is)<script[^>]*>.*?</script>")
HTML_BODY_PATTERN = re.compile(r"(?is)<body[^>]*>(?P<body>.*)</body>")
HTML_TAG_PATTERN = re.compile(r"(?is)<[^>]+>")
IMG_TAG_PATTERN = re.compile(r"(?is)<img\b[^>]*>")
SRC_ATTR_PATTERN = re.compile(
    r"""(?is)\bsrc\s*=\s*(?:(?P<quote>["'])(?P<quoted>.*?)(?P=quote)|(?P<bare>[^\s>]+))"""
)
WIDTH_ATTR_PATTERN = re.compile(r"""(?is)\bwidth\s*=\s*(?:"([^"]+)"|'([^']+)'|([^\s>]+))""")
HEIGHT_ATTR_PATTERN = re.compile(r"""(?is)\bheight\s*=\s*(?:"([^"]+)"|'([^']+)'|([^\s>]+))""")
STYLE_ATTR_PATTERN = re.compile(r"""(?is)\bstyle\s*=\s*(?:"([^"]*)"|'([^']*)')""")
CSS_WIDTH_PATTERN = re.compile(r"(?is)\bwidth\s*:\s*([^;]+)")
CSS_HEIGHT_PATTERN = re.compile(r"(?is)\bheight\s*:\s*([^;]+)")
SIGNATURE_MARKER_PATTERN = re.compile(
    r"(?i)\b(?:thanks|thank you|regards|best regards|kind regards|sincerely|cheers|best)\b"
)
RENDER_STRATEGY_FIDELITY = "fidelity"
RENDER_STRATEGY_FAST = "fast"
REMOTE_IMAGE_SCHEMES = ("http://", "https://")


@dataclass(slots=True)
class PdfWriteDiagnostics:
    pipeline: str = ""
    total_seconds: float = 0.0
    stage_seconds: dict[str, float] = field(default_factory=dict)
    image_metrics: dict[str, int] = field(default_factory=dict)


def _empty_image_metrics() -> dict[str, int]:
    return {
        "total_images": 0,
        "cid_resolved": 0,
        "cid_unresolved": 0,
        "signature_small_dropped": 0,
        "remote_dropped": 0,
    }


def _get_render_strategy() -> str:
    raw_value = os.environ.get("MSG_TO_PDF_RENDER_STRATEGY", "").strip().lower()
    if raw_value == RENDER_STRATEGY_FAST:
        return RENDER_STRATEGY_FAST
    return RENDER_STRATEGY_FIDELITY


def _allow_remote_images() -> bool:
    raw_value = os.environ.get("MSG_TO_PDF_ALLOW_REMOTE_IMAGES", "").strip().lower()
    return raw_value in {"1", "true", "yes", "on"}


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


def _normalize_cid(value: str) -> str:
    cid = (value or "").strip().strip("<>").strip()
    if cid.lower().startswith("cid:"):
        cid = cid[4:].strip()
    return cid


def _build_inline_image_lookup(inline_images: list[InlineImageAsset]) -> dict[str, InlineImageAsset]:
    lookup: dict[str, InlineImageAsset] = {}
    for asset in inline_images:
        cid_key = _normalize_cid(asset.cid).lower()
        if cid_key:
            lookup[cid_key] = asset
    return lookup


def _parse_numeric_dimension(value: str) -> int | None:
    text = (value or "").strip()
    if not text:
        return None
    match = re.search(r"(\d+(?:\.\d+)?)", text)
    if match is None:
        return None
    number = float(match.group(1))
    if number <= 0:
        return None
    return int(round(number))


def _extract_dimension_value(attribute_match: re.Match[str] | None) -> int | None:
    if attribute_match is None:
        return None
    for value in attribute_match.groups():
        if value:
            return _parse_numeric_dimension(value)
    return None


def _extract_img_dimensions(tag_html: str) -> tuple[int | None, int | None]:
    width = _extract_dimension_value(WIDTH_ATTR_PATTERN.search(tag_html))
    height = _extract_dimension_value(HEIGHT_ATTR_PATTERN.search(tag_html))

    style_match = STYLE_ATTR_PATTERN.search(tag_html)
    if style_match is not None:
        style_value = style_match.group(1) or style_match.group(2) or ""
        if width is None:
            css_width = CSS_WIDTH_PATTERN.search(style_value)
            if css_width is not None:
                width = _parse_numeric_dimension(css_width.group(1))
        if height is None:
            css_height = CSS_HEIGHT_PATTERN.search(style_value)
            if css_height is not None:
                height = _parse_numeric_dimension(css_height.group(1))

    return width, height


def _is_small_signature_image(width: int | None, height: int | None, size_bytes: int) -> bool:
    max_dimension = max((value for value in (width, height) if value is not None), default=None)
    area = width * height if width is not None and height is not None else None
    size_gate = size_bytes <= 20 * 1024
    dimension_gate = (
        (max_dimension is not None and max_dimension <= 180)
        or (area is not None and area <= 24_000)
    )
    return size_gate and dimension_gate


def _is_content_image(width: int | None, height: int | None, size_bytes: int) -> bool:
    max_dimension = max((value for value in (width, height) if value is not None), default=None)
    area = width * height if width is not None and height is not None else None
    return (
        size_bytes >= 40 * 1024
        or (max_dimension is not None and max_dimension >= 260)
        or (area is not None and area >= 40_000)
    )


def _build_data_uri(asset: InlineImageAsset) -> str:
    encoded = base64.b64encode(asset.data).decode("ascii")
    return f"data:{asset.mime_type};base64,{encoded}"


def _extract_img_src(tag_html: str) -> tuple[str, tuple[int, int] | None, str]:
    match = SRC_ATTR_PATTERN.search(tag_html)
    if match is None:
        return "", None, '"'
    src_value = match.group("quoted")
    if src_value is None:
        src_value = match.group("bare") or ""
    quote = match.group("quote") or '"'
    return unescape(src_value.strip()), match.span(), quote


def _replace_img_src(tag_html: str, source_span: tuple[int, int], quote: str, new_src: str) -> str:
    replacement = f"src={quote}{new_src}{quote}"
    return f"{tag_html[:source_span[0]]}{replacement}{tag_html[source_span[1]:]}"


def _estimate_data_uri_bytes(data_uri: str) -> int:
    comma_index = data_uri.find(",")
    if comma_index < 0:
        return 0
    header = data_uri[:comma_index].lower()
    payload = re.sub(r"\s+", "", data_uri[comma_index + 1 :])
    if ";base64" in header:
        padding = payload.count("=")
        return max(0, (len(payload) * 3) // 4 - padding)
    return len(payload.encode("utf-8"))


def _html_to_signature_text(fragment: str) -> str:
    normalized = re.sub(r"(?is)<br\s*/?>", "\n", fragment)
    normalized = re.sub(
        r"(?is)</?(?:p|div|li|tr|td|th|table|section|article|header|footer|h[1-6])[^>]*>",
        "\n",
        normalized,
    )
    normalized = HTML_TAG_PATTERN.sub(" ", normalized)
    normalized = unescape(normalized).replace("\r", "\n")
    normalized = re.sub(r"[ \t\f\v]+", " ", normalized)
    normalized = re.sub(r"\n{3,}", "\n\n", normalized)
    return normalized.lower()


def _is_signature_zone(html_fragment: str, image_start: int) -> bool:
    if not html_fragment:
        return False
    if image_start < int(len(html_fragment) * 0.35):
        return False
    prefix = html_fragment[max(0, image_start - 2400) : image_start]
    marker_window = _html_to_signature_text(prefix)[-700:]
    return SIGNATURE_MARKER_PATTERN.search(marker_window) is not None


def _normalize_img_tag(
    tag_html: str,
    *,
    inline_lookup: dict[str, InlineImageAsset],
    signature_zone: bool,
    allow_remote: bool,
    image_metrics: dict[str, int],
) -> str:
    src, source_span, quote = _extract_img_src(tag_html)
    if source_span is None or not src:
        return tag_html

    src_lower = src.lower()
    width, height = _extract_img_dimensions(tag_html)

    if src_lower.startswith(REMOTE_IMAGE_SCHEMES):
        if allow_remote:
            return tag_html
        image_metrics["remote_dropped"] += 1
        return ""

    if src_lower.startswith("cid:"):
        cid_key = _normalize_cid(src).lower()
        asset = inline_lookup.get(cid_key)
        if asset is None:
            image_metrics["cid_unresolved"] += 1
            return ""

        if (
            signature_zone
            and _is_small_signature_image(width, height, asset.size_bytes)
            and not _is_content_image(width, height, asset.size_bytes)
        ):
            image_metrics["signature_small_dropped"] += 1
            return ""

        image_metrics["cid_resolved"] += 1
        return _replace_img_src(tag_html, source_span, quote, _build_data_uri(asset))

    if src_lower.startswith("data:") and signature_zone:
        data_size = _estimate_data_uri_bytes(src)
        if (
            _is_small_signature_image(width, height, data_size)
            and not _is_content_image(width, height, data_size)
        ):
            image_metrics["signature_small_dropped"] += 1
            return ""

    return tag_html


def _rewrite_html_images(record: EmailRecord, html_fragment: str) -> tuple[str, dict[str, int]]:
    image_metrics = _empty_image_metrics()
    if not html_fragment:
        return html_fragment, image_metrics

    inline_lookup = _build_inline_image_lookup(record.inline_images)
    allow_remote = _allow_remote_images()

    output_parts: list[str] = []
    previous_end = 0
    for match in IMG_TAG_PATTERN.finditer(html_fragment):
        output_parts.append(html_fragment[previous_end : match.start()])
        tag_html = match.group(0)
        image_metrics["total_images"] += 1
        rewritten = _normalize_img_tag(
            tag_html,
            inline_lookup=inline_lookup,
            signature_zone=_is_signature_zone(html_fragment, match.start()),
            allow_remote=allow_remote,
            image_metrics=image_metrics,
        )
        output_parts.append(rewritten)
        previous_end = match.end()

    output_parts.append(html_fragment[previous_end:])
    return "".join(output_parts), image_metrics


def _prepare_html_body_fragment(
    record: EmailRecord,
    diagnostics: PdfWriteDiagnostics | None = None,
) -> str:
    html_fragment = _extract_body_html_fragment(record.html_body)
    if not html_fragment:
        if diagnostics is not None:
            diagnostics.image_metrics.clear()
            diagnostics.image_metrics.update(_empty_image_metrics())
        return ""

    rewritten, image_metrics = _rewrite_html_images(record, html_fragment)
    if diagnostics is not None:
        diagnostics.image_metrics.clear()
        diagnostics.image_metrics.update(image_metrics)
    return rewritten


def _build_header_rows_html(record: EmailRecord) -> str:
    rows = [
        ("From:", record.sender or "(empty)"),
        ("Sent:", _format_sent_value(record.sent_at)),
        ("To:", record.to or "(empty)"),
        ("Cc:", record.cc or "(empty)"),
        ("Subject:", record.subject or "No Subject"),
    ]
    if record.attachment_names:
        rows.append(("Attachments:", "\n".join(record.attachment_names)))

    return "".join(
        (
            '<div class="header-row">'
            f'<div class="header-label">{escape(label)}</div>'
            f'<div class="header-value">{_as_paragraph_text(value)}</div>'
            "</div>"
        )
        for label, value in rows
    )


def _build_body_fragment(record: EmailRecord, diagnostics: PdfWriteDiagnostics | None = None) -> str:
    html_fragment = _prepare_html_body_fragment(record, diagnostics=diagnostics)
    if html_fragment:
        return html_fragment
    if diagnostics is not None:
        diagnostics.image_metrics.clear()
        diagnostics.image_metrics.update(_empty_image_metrics())
    plain = escape(record.body or "(No body text.)")
    return f'<pre class="plain-body">{plain}</pre>'


def build_email_html_document(
    record: EmailRecord,
    *,
    diagnostics: PdfWriteDiagnostics | None = None,
) -> str:
    body_fragment = _build_body_fragment(record, diagnostics=diagnostics)
    header_rows_html = _build_header_rows_html(record)

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
    .message-header {{
      margin: 0 0 14px;
      padding: 0 0 12px;
      border-bottom: 1px solid #d1d5db;
    }}
    .header-row {{
      display: grid;
      grid-template-columns: 88px minmax(0, 1fr);
      column-gap: 12px;
      align-items: start;
      margin: 0 0 6px;
    }}
    .header-row:last-child {{
      margin-bottom: 0;
    }}
    .header-label {{
      font-weight: 700;
    }}
    .header-value {{
      word-break: break-word;
      white-space: pre-wrap;
    }}
    .body {{
      padding-top: 2px;
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
  <section class="message-header">
    {header_rows_html}
  </section>
  <section class="body">
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
        "--no-pdf-header-footer",
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
        Paragraph(f"<b>From:</b> {_as_paragraph_text(record.sender or '(empty)')}", normal_style),
        Spacer(1, 6),
        Paragraph(f"<b>Sent:</b> {_as_paragraph_text(_format_sent_value(record.sent_at))}", normal_style),
        Spacer(1, 6),
        Paragraph(f"<b>To:</b> {_as_paragraph_text(record.to or '(empty)')}", normal_style),
        Spacer(1, 6),
        Paragraph(f"<b>Cc:</b> {_as_paragraph_text(record.cc or '(empty)')}", normal_style),
        Spacer(1, 6),
        Paragraph(f"<b>Subject:</b> {_as_paragraph_text(record.subject or 'No Subject')}", normal_style),
        Spacer(1, 12),
    ]

    if record.attachment_names:
        attachments_text = "<br/>".join(_as_paragraph_text(name) for name in record.attachment_names)
        story.append(Paragraph(f"<b>Attachments:</b> {attachments_text}", normal_style))
        story.append(Spacer(1, 12))

    for block in _format_body_blocks(record.body):
        story.append(Paragraph(_as_paragraph_text(block), body_style))
        story.append(Spacer(1, 10))

    document.build(story)


def write_email_pdf(
    record: EmailRecord,
    output_path: Path,
    *,
    diagnostics: PdfWriteDiagnostics | None = None,
    event_sink: TaskEventSink | None = None,
    task_id: str | None = None,
    event_meta: Mapping[str, TaskMetaValue] | None = None,
) -> Path:
    started_at = perf_counter()
    resolved_task_id = task_id or default_task_id_for_path(record.source_path)
    if diagnostics is not None:
        diagnostics.pipeline = ""
        diagnostics.total_seconds = 0.0
        diagnostics.stage_seconds.clear()
        diagnostics.image_metrics.clear()
        diagnostics.image_metrics.update(_empty_image_metrics())

    output_path.parent.mkdir(parents=True, exist_ok=True)
    render_strategy = _get_render_strategy()
    prefer_cid_html = bool(record.inline_images) and bool(record.html_body.strip())
    if diagnostics is not None:
        diagnostics.stage_seconds["render_strategy_fast"] = (
            1.0 if render_strategy == RENDER_STRATEGY_FAST else 0.0
        )
        diagnostics.stage_seconds["auto_cid_html"] = 1.0 if prefer_cid_html else 0.0

    if render_strategy == RENDER_STRATEGY_FIDELITY:
        emit_task_event(
            event_sink,
            task_id=resolved_task_id,
            stage="pipeline_selected",
            file_name=record.source_path.name,
            pipeline="outlook_edge",
            meta=_merge_event_meta(event_meta, {"outputName": output_path.name}),
        )
        stage_started_at = perf_counter()
        if _try_write_pdf_via_outlook_and_edge(record.source_path, output_path):
            if diagnostics is not None:
                diagnostics.pipeline = "outlook_edge"
                diagnostics.stage_seconds["outlook_edge"] = perf_counter() - stage_started_at
                diagnostics.total_seconds = perf_counter() - started_at
            return output_path
        if diagnostics is not None:
            diagnostics.stage_seconds["outlook_edge"] = perf_counter() - stage_started_at
    elif diagnostics is not None:
        diagnostics.stage_seconds["outlook_edge"] = 0.0

    stage_started_at = perf_counter()
    html_document = build_email_html_document(record, diagnostics=diagnostics)
    if diagnostics is not None:
        diagnostics.stage_seconds["build_html"] = perf_counter() - stage_started_at

    emit_task_event(
        event_sink,
        task_id=resolved_task_id,
        stage="pipeline_selected",
        file_name=record.source_path.name,
        pipeline="edge_html",
        meta=_merge_event_meta(event_meta, {"outputName": output_path.name}),
    )
    stage_started_at = perf_counter()
    if _try_write_pdf_via_edge_html(html_document, output_path):
        if diagnostics is not None:
            diagnostics.pipeline = "edge_html"
            diagnostics.stage_seconds["edge_html"] = perf_counter() - stage_started_at
            diagnostics.total_seconds = perf_counter() - started_at
        return output_path
    if diagnostics is not None:
        diagnostics.stage_seconds["edge_html"] = perf_counter() - stage_started_at

    emit_task_event(
        event_sink,
        task_id=resolved_task_id,
        stage="pipeline_selected",
        file_name=record.source_path.name,
        pipeline="reportlab",
        meta=_merge_event_meta(event_meta, {"outputName": output_path.name}),
    )
    stage_started_at = perf_counter()
    _write_pdf_via_reportlab(record, output_path)
    if diagnostics is not None:
        diagnostics.pipeline = "reportlab_fast" if render_strategy == RENDER_STRATEGY_FAST else "reportlab"
        diagnostics.stage_seconds["reportlab"] = perf_counter() - stage_started_at
        diagnostics.total_seconds = perf_counter() - started_at
    return output_path


def _merge_event_meta(
    base_meta: Mapping[str, TaskMetaValue] | None,
    extra_meta: Mapping[str, TaskMetaValue] | None,
) -> dict[str, TaskMetaValue] | None:
    merged: dict[str, TaskMetaValue] = {}
    if base_meta:
        merged.update(base_meta)
    if extra_meta:
        merged.update(extra_meta)
    return merged or None
