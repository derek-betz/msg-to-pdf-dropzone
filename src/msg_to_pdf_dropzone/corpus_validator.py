from __future__ import annotations

import argparse
import json
import os
import re
import unicodedata
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from statistics import mean

from pypdf import PdfReader

from .converter import ConversionResult, convert_msg_files
from .pdf_writer import RENDER_STRATEGY_FAST, RENDER_STRATEGY_FIDELITY

DEFAULT_CASES_DIR = Path("emails-for-testing")
DEFAULT_OUTPUT_ROOT = Path(".local-corpus-profiles")
HEADER_LABELS = ("From", "Sent", "To", "Cc", "Subject", "Attachments")
HEADER_LABEL_PATTERN = re.compile(
    r"^\s*(From|Sent|To|Cc|Subject|Attachments)\s*:\s*(.*)$",
    re.IGNORECASE,
)
INLINE_HEADER_LABEL_PATTERN = re.compile(r"(From|Sent|To|Cc|Subject|Attachments)\s*:\s*", re.IGNORECASE)
TOKEN_PATTERN = re.compile(r"[a-z0-9@._+-]+")
DATE_ISO_PATTERN = re.compile(r"\b(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2})\b")
DATE_WORD_PATTERN = re.compile(
    r"\b(?P<month_name>January|February|March|April|May|June|July|August|September|October|"
    r"November|December)\s+(?P<day>\d{1,2}),\s+(?P<year>\d{4})\b",
    re.IGNORECASE,
)
TIME_PATTERN = re.compile(r"\b(?P<hour>\d{1,2}):(?P<minute>\d{2})(?::(?P<second>\d{2}))?\s*(?P<ampm>AM|PM)?\b")
ATTACHMENT_NAME_PATTERN = re.compile(r"[^;\n]+?\.[A-Za-z0-9]{2,8}")
MONTH_LOOKUP = {
    "january": 1,
    "february": 2,
    "march": 3,
    "april": 4,
    "may": 5,
    "june": 6,
    "july": 7,
    "august": 8,
    "september": 9,
    "october": 10,
    "november": 11,
    "december": 12,
}
HEADER_ORDER_INDEX = {label: index for index, label in enumerate(HEADER_LABELS)}
HEADER_CONTINUATION_LABELS = {"To", "Cc", "Subject", "Attachments"}


@dataclass(slots=True)
class CasePair:
    case_id: str
    case_dir: Path
    msg_path: Path
    golden_pdf_path: Path


@dataclass(slots=True)
class PdfSnapshot:
    path: Path
    page_count: int
    file_size_bytes: int
    extracted_text: str
    normalized_text: str
    header_fields: dict[str, str]
    header_order: list[str]
    body_text: str
    body_anchors: list[str]


@dataclass(slots=True)
class ComparisonIssue:
    code: str
    message: str


@dataclass(slots=True)
class CaseValidationResult:
    case_id: str
    msg_path: str
    golden_pdf_path: str
    generated_pdf_path: str | None
    pipeline: str
    passed: bool
    hard_failures: list[dict[str, str]] = field(default_factory=list)
    warnings: list[dict[str, str]] = field(default_factory=list)
    infos: list[dict[str, str]] = field(default_factory=list)
    header_order_ok: bool = True
    golden_header_order: list[str] = field(default_factory=list)
    generated_header_order: list[str] = field(default_factory=list)
    golden_headers: dict[str, str] = field(default_factory=dict)
    generated_headers: dict[str, str] = field(default_factory=dict)
    attachment_names: list[str] = field(default_factory=list)
    missing_attachments: list[str] = field(default_factory=list)
    body_anchor_count: int = 0
    body_anchor_hits: int = 0
    body_anchor_hit_ratio: float = 1.0
    page_count_golden: int = 0
    page_count_generated: int = 0
    file_size_golden: int = 0
    file_size_generated: int = 0
    conversion_errors: list[str] = field(default_factory=list)


def _sort_case_dirs(paths: list[Path]) -> list[Path]:
    def _sort_key(path: Path) -> tuple[int, int | str]:
        name = path.name
        if name.isdigit():
            return (0, int(name))
        return (1, name.lower())

    return sorted(paths, key=_sort_key)


def _single_match(case_dir: Path, pattern: str, label: str) -> Path:
    matches = sorted(path for path in case_dir.glob(pattern) if path.is_file())
    if len(matches) != 1:
        raise ValueError(
            f"Case '{case_dir.name}' must contain exactly one {label}; found {len(matches)}."
        )
    return matches[0]


def discover_case_pairs(cases_dir: Path, case_id: str | None = None) -> list[CasePair]:
    cases_dir = cases_dir.resolve()
    if not cases_dir.exists():
        raise FileNotFoundError(f"Cases directory does not exist: {cases_dir}")

    if case_id is not None:
        selected_dir = cases_dir / case_id
        if not selected_dir.exists() or not selected_dir.is_dir():
            raise FileNotFoundError(f"Requested case does not exist: {selected_dir}")
        case_dirs = [selected_dir]
    else:
        case_dirs = _sort_case_dirs([path for path in cases_dir.iterdir() if path.is_dir()])
        if not case_dirs:
            raise ValueError(f"No case directories found in: {cases_dir}")

    pairs: list[CasePair] = []
    for case_dir in case_dirs:
        pairs.append(
            CasePair(
                case_id=case_dir.name,
                case_dir=case_dir,
                msg_path=_single_match(case_dir, "*.msg", ".msg file"),
                golden_pdf_path=_single_match(case_dir, "*.pdf", "golden .pdf file"),
            )
        )
    return pairs


def _normalize_text(value: str) -> str:
    text = unicodedata.normalize("NFKC", value or "")
    text = text.replace("\u00A0", " ").replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t\f\v]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _collapse_spaced_header_labels(value: str) -> str:
    collapsed = value
    for label in HEADER_LABELS:
        letters = r"\s*".join(re.escape(character) for character in label)
        pattern = re.compile(rf"\b{letters}\s*:", re.IGNORECASE)
        collapsed = pattern.sub(f"{label}:", collapsed)
    return collapsed


def _normalize_inline_text(value: str) -> str:
    return re.sub(r"\s+", " ", _normalize_text(value)).strip().lower()


def _tokenize(value: str) -> set[str]:
    tokens: set[str] = set()
    for token in TOKEN_PATTERN.findall(_normalize_inline_text(value)):
        tokens.add(token)
        if "@" in token:
            local_part = token.split("@", 1)[0]
            tokens.add(local_part)
            tokens.update(part for part in re.split(r"[._-]+", local_part) if part)
    return tokens


def _extract_header_and_body_inline(extracted_text: str) -> tuple[dict[str, str], list[str], str]:
    normalized = _collapse_spaced_header_labels(_normalize_text(extracted_text))
    candidate_window = normalized[:4000]
    matches = list(INLINE_HEADER_LABEL_PATTERN.finditer(candidate_window))
    if not matches:
        return {}, [], normalized

    header_matches: list[re.Match[str]] = []
    last_order = -1
    for match in matches:
        label = match.group(1).title()
        order_index = HEADER_ORDER_INDEX[label]
        if not header_matches:
            if label != "From":
                continue
            header_matches.append(match)
            last_order = order_index
            continue
        if order_index < last_order:
            break
        header_matches.append(match)
        last_order = order_index

    if len(header_matches) <= 1:
        return {}, [], normalized

    headers: dict[str, str] = {}
    order: list[str] = []
    for index, match in enumerate(header_matches):
        label = match.group(1).title()
        start = match.end()
        end = header_matches[index + 1].start() if index + 1 < len(header_matches) else len(candidate_window)
        value = candidate_window[start:end].strip()
        headers[label] = value
        order.append(label)

    body_start = header_matches[-1].end() + len(headers[order[-1]])
    body_text = normalized[body_start:].strip()
    return headers, order, body_text


def _extract_header_and_body(extracted_text: str) -> tuple[dict[str, str], list[str], str]:
    normalized = _collapse_spaced_header_labels(_normalize_text(extracted_text))
    lines = normalized.splitlines()
    headers: dict[str, str] = {}
    header_order: list[str] = []
    current_label: str | None = None
    body_start_index = len(lines)
    header_started = False

    for index, line in enumerate(lines):
        stripped = re.sub(r"\s+", " ", line).strip()
        label_match = HEADER_LABEL_PATTERN.match(stripped)
        if label_match is None and not header_started:
            continue

        if label_match is not None:
            header_started = True
            current_label = label_match.group(1).title()
            if current_label not in header_order:
                header_order.append(current_label)
            headers[current_label] = label_match.group(2).strip()
            continue

        if stripped == "":
            current_label = None
            if header_started:
                body_start_index = index + 1
                break
            continue

        if current_label in HEADER_CONTINUATION_LABELS:
            headers[current_label] = f"{headers[current_label]} {stripped}".strip()
            continue

        body_start_index = index
        break

    body_text = "\n".join(lines[body_start_index:]).strip()
    if len(header_order) <= 1 and len(list(INLINE_HEADER_LABEL_PATTERN.finditer(normalized[:4000]))) > 1:
        inline_headers, inline_order, inline_body = _extract_header_and_body_inline(extracted_text)
        if len(inline_order) > len(header_order):
            return inline_headers, inline_order, inline_body
    return headers, header_order, body_text


def _extract_attachment_names(value: str) -> list[str]:
    if not value.strip():
        return []
    matches = [match.group(0).strip(" ;") for match in ATTACHMENT_NAME_PATTERN.finditer(value)]
    if matches:
        unique_matches: list[str] = []
        seen: set[str] = set()
        for name in matches:
            normalized_name = _normalize_inline_text(name)
            if normalized_name not in seen:
                unique_matches.append(name)
                seen.add(normalized_name)
        return unique_matches

    parts = []
    for raw_part in re.split(r"[\n;]+", value):
        part = raw_part.strip()
        if part:
            parts.append(part)
    return parts


def _extract_date_token(value: str) -> tuple[int, int, int] | None:
    normalized = _normalize_text(value)
    iso_match = DATE_ISO_PATTERN.search(normalized)
    if iso_match is not None:
        return (
            int(iso_match.group("year")),
            int(iso_match.group("month")),
            int(iso_match.group("day")),
        )

    word_match = DATE_WORD_PATTERN.search(normalized)
    if word_match is None:
        return None
    month_number = MONTH_LOOKUP[word_match.group("month_name").lower()]
    return (
        int(word_match.group("year")),
        month_number,
        int(word_match.group("day")),
    )


def _extract_time_minutes(value: str) -> int | None:
    normalized = _normalize_text(value)
    match = TIME_PATTERN.search(normalized)
    if match is None:
        return None

    hour = int(match.group("hour"))
    minute = int(match.group("minute"))
    ampm = (match.group("ampm") or "").upper()
    if ampm == "AM" and hour == 12:
        hour = 0
    elif ampm == "PM" and hour < 12:
        hour += 12
    if hour > 23:
        return None
    return hour * 60 + minute


def _sent_values_match(expected: str, actual: str) -> bool:
    expected_date = _extract_date_token(expected)
    actual_date = _extract_date_token(actual)
    if expected_date is not None and actual_date is not None and expected_date != actual_date:
        return False

    expected_minutes = _extract_time_minutes(expected)
    actual_minutes = _extract_time_minutes(actual)
    if expected_minutes is not None and actual_minutes is not None:
        delta = abs(expected_minutes - actual_minutes)
        delta = min(delta, 24 * 60 - delta)
        return delta <= 2

    if expected_date is not None and actual_date is not None:
        return True
    return _normalize_inline_text(expected) in _normalize_inline_text(actual)


def _field_values_match(label: str, expected: str, actual: str) -> bool:
    if not expected.strip():
        return True
    if not actual.strip():
        return False
    if label == "Sent":
        return _sent_values_match(expected, actual)

    expected_normalized = _normalize_inline_text(expected)
    actual_normalized = _normalize_inline_text(actual)
    if expected_normalized in actual_normalized:
        return True
    if label == "Attachments":
        expected_compact = re.sub(r"\s+", "", expected_normalized)
        actual_compact = re.sub(r"\s+", "", actual_normalized)
        if expected_compact and expected_compact in actual_compact:
            return True

    expected_tokens = {token for token in _tokenize(expected) if len(token) >= 3}
    actual_tokens = _tokenize(actual)
    if expected_tokens and expected_tokens.issubset(actual_tokens):
        return True

    expected_name_tokens = [token for token in TOKEN_PATTERN.findall(_normalize_inline_text(expected)) if len(token) >= 2]
    if len(expected_name_tokens) >= 2:
        first = expected_name_tokens[0]
        last = expected_name_tokens[-1]
        aliases = {
            f"{first[0]}{last}",
            f"{first}{last}",
            f"{first}.{last}",
            f"{first}_{last}",
        }
        return any(alias in actual_tokens for alias in aliases)
    return False


def _extract_body_anchors(body_text: str) -> list[str]:
    candidates: list[tuple[int, str]] = []
    seen: set[str] = set()
    for raw_line in _normalize_text(body_text).splitlines():
        line = re.sub(r"\s+", " ", raw_line).strip()
        if len(line) < 20:
            continue
        if HEADER_LABEL_PATTERN.match(line):
            continue
        if re.fullmatch(r"[-=_* .]{4,}", line):
            continue
        normalized_line = _normalize_inline_text(line)
        if normalized_line in seen:
            continue
        candidates.append((len(line), line))
        seen.add(normalized_line)

    longest = sorted(candidates, key=lambda item: (-item[0], item[1].lower()))
    return [line for _length, line in longest[:5]]


def _anchor_matches(anchor: str, normalized_generated_text: str) -> bool:
    normalized_anchor = _normalize_inline_text(anchor)
    if normalized_anchor in normalized_generated_text:
        return True

    anchor_tokens = [token for token in TOKEN_PATTERN.findall(normalized_anchor) if len(token) >= 3]
    if not anchor_tokens:
        return False

    generated_tokens = _tokenize(normalized_generated_text)
    hit_count = sum(1 for token in anchor_tokens if token in generated_tokens)
    return (hit_count / len(anchor_tokens)) >= 0.75


def analyze_pdf(pdf_path: Path) -> PdfSnapshot:
    reader = PdfReader(str(pdf_path))
    extracted_text = "\n".join(page.extract_text() or "" for page in reader.pages).strip()
    if not extracted_text:
        raise ValueError(f"No extractable text found in PDF: {pdf_path}")

    header_fields, header_order, body_text = _extract_header_and_body(extracted_text)
    return PdfSnapshot(
        path=pdf_path,
        page_count=len(reader.pages),
        file_size_bytes=pdf_path.stat().st_size,
        extracted_text=extracted_text,
        normalized_text=_normalize_inline_text(extracted_text),
        header_fields=header_fields,
        header_order=header_order,
        body_text=body_text,
        body_anchors=_extract_body_anchors(body_text),
    )


def _issue_dict(issue: ComparisonIssue) -> dict[str, str]:
    return {"code": issue.code, "message": issue.message}


def _compare_header_fields(
    golden: PdfSnapshot,
    generated: PdfSnapshot,
) -> tuple[list[ComparisonIssue], list[str]]:
    failures: list[ComparisonIssue] = []
    missing_attachments: list[str] = []

    for label in HEADER_LABELS:
        expected_value = golden.header_fields.get(label, "").strip()
        actual_value = generated.header_fields.get(label, "").strip()
        if label == "Attachments":
            expected_attachments = _extract_attachment_names(expected_value)
            if expected_attachments:
                actual_attachments_text = actual_value
                for attachment_name in expected_attachments:
                    if not _field_values_match(label, attachment_name, actual_attachments_text):
                        missing_attachments.append(attachment_name)
                if missing_attachments:
                    failures.append(
                        ComparisonIssue(
                            code="attachment_missing",
                            message=(
                                "Generated PDF is missing golden attachment name(s): "
                                + ", ".join(missing_attachments)
                            ),
                        )
                    )
            continue

        if expected_value and not _field_values_match(label, expected_value, actual_value):
            failures.append(
                ComparisonIssue(
                    code=f"missing_header_{label.lower()}",
                    message=f"Generated PDF is missing or mismatching header field '{label}'.",
                )
            )
            continue
    return failures, missing_attachments


def compare_snapshots(
    golden: PdfSnapshot,
    generated: PdfSnapshot,
    *,
    pipeline: str,
    conversion_errors: list[str] | None = None,
) -> CaseValidationResult:
    failures: list[ComparisonIssue] = []
    warnings: list[ComparisonIssue] = []
    infos: list[ComparisonIssue] = []
    conversion_errors = conversion_errors or []

    header_order_ok = True
    generated_positions = [HEADER_ORDER_INDEX[label] for label in generated.header_order if label in HEADER_ORDER_INDEX]
    if generated_positions != sorted(generated_positions):
        header_order_ok = False
        failures.append(
            ComparisonIssue(
                code="header_order",
                message="Generated PDF metadata header order does not follow the expected Outlook-style order.",
            )
        )

    header_failures, missing_attachments = _compare_header_fields(golden, generated)
    failures.extend(header_failures)

    anchor_hits = 0
    if golden.body_anchors:
        for anchor in golden.body_anchors:
            if _anchor_matches(anchor, generated.normalized_text):
                anchor_hits += 1
        anchor_ratio = anchor_hits / len(golden.body_anchors)
        if anchor_ratio < 0.8:
            failures.append(
                ComparisonIssue(
                    code="body_anchor_low",
                    message=(
                        f"Generated PDF matched {anchor_hits}/{len(golden.body_anchors)} golden body anchors."
                    ),
                )
            )
    else:
        anchor_ratio = 1.0

    page_delta = abs(golden.page_count - generated.page_count)
    if page_delta > 2:
        failures.append(
            ComparisonIssue(
                code="page_count",
                message=(
                    f"Page count drift is too large: golden={golden.page_count}, generated={generated.page_count}."
                ),
            )
        )
    elif page_delta >= 1:
        infos.append(
            ComparisonIssue(
                code="page_count_info",
                message=(
                    f"Page count drift detected: golden={golden.page_count}, generated={generated.page_count}."
                ),
            )
        )

    if pipeline == "reportlab":
        warnings.append(
            ComparisonIssue(
                code="pipeline_reportlab",
                message="Generated PDF used the ReportLab fallback pipeline.",
            )
        )

    return CaseValidationResult(
        case_id=golden.path.parent.name,
        msg_path="",
        golden_pdf_path=str(golden.path),
        generated_pdf_path=str(generated.path),
        pipeline=pipeline,
        passed=not failures,
        hard_failures=[_issue_dict(issue) for issue in failures],
        warnings=[_issue_dict(issue) for issue in warnings],
        infos=[_issue_dict(issue) for issue in infos],
        header_order_ok=header_order_ok,
        golden_header_order=golden.header_order,
        generated_header_order=generated.header_order,
        golden_headers=dict(golden.header_fields),
        generated_headers=dict(generated.header_fields),
        attachment_names=_extract_attachment_names(golden.header_fields.get("Attachments", "")),
        missing_attachments=missing_attachments,
        body_anchor_count=len(golden.body_anchors),
        body_anchor_hits=anchor_hits,
        body_anchor_hit_ratio=round(anchor_ratio, 4),
        page_count_golden=golden.page_count,
        page_count_generated=generated.page_count,
        file_size_golden=golden.file_size_bytes,
        file_size_generated=generated.file_size_bytes,
        conversion_errors=list(conversion_errors),
    )


def validate_case_pair(case_pair: CasePair, generated_root: Path) -> CaseValidationResult:
    generated_root.mkdir(parents=True, exist_ok=True)
    case_output_dir = generated_root / case_pair.case_id
    case_output_dir.mkdir(parents=True, exist_ok=True)

    conversion = convert_msg_files([case_pair.msg_path], case_output_dir)
    if not conversion.converted_files:
        return CaseValidationResult(
            case_id=case_pair.case_id,
            msg_path=str(case_pair.msg_path),
            golden_pdf_path=str(case_pair.golden_pdf_path),
            generated_pdf_path=None,
            pipeline=conversion.file_timing_records[0].pipeline if conversion.file_timing_records else "",
            passed=False,
            hard_failures=[
                _issue_dict(
                    ComparisonIssue(
                        code="conversion_failure",
                        message=(
                            f"Case conversion failed: {'; '.join(conversion.errors) or 'no PDF generated'}."
                        ),
                    )
                )
            ],
            conversion_errors=list(conversion.errors),
        )

    generated_pdf_path = conversion.converted_files[0]
    pipeline = conversion.file_timing_records[0].pipeline if conversion.file_timing_records else ""
    try:
        golden_snapshot = analyze_pdf(case_pair.golden_pdf_path)
        generated_snapshot = analyze_pdf(generated_pdf_path)
    except Exception as exc:
        return CaseValidationResult(
            case_id=case_pair.case_id,
            msg_path=str(case_pair.msg_path),
            golden_pdf_path=str(case_pair.golden_pdf_path),
            generated_pdf_path=str(generated_pdf_path),
            pipeline=pipeline,
            passed=False,
            hard_failures=[
                _issue_dict(
                    ComparisonIssue(
                        code="text_extraction_failure",
                        message=f"PDF analysis failed: {exc}",
                    )
                )
            ],
            conversion_errors=list(conversion.errors),
        )

    comparison = compare_snapshots(
        golden_snapshot,
        generated_snapshot,
        pipeline=pipeline,
        conversion_errors=conversion.errors,
    )
    comparison.case_id = case_pair.case_id
    comparison.msg_path = str(case_pair.msg_path)
    comparison.golden_pdf_path = str(case_pair.golden_pdf_path)
    comparison.generated_pdf_path = str(generated_pdf_path)
    comparison.pipeline = pipeline
    comparison.conversion_errors = list(conversion.errors)
    return comparison


def _group_issue_counts(results: list[CaseValidationResult]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for result in results:
        for issue in result.hard_failures + result.warnings + result.infos:
            counts[issue["code"]] = counts.get(issue["code"], 0) + 1
    return dict(sorted(counts.items(), key=lambda item: (-item[1], item[0])))


def _group_case_ids_by_issue(results: list[CaseValidationResult]) -> dict[str, list[str]]:
    grouped: dict[str, list[str]] = {}
    for result in results:
        for issue in result.hard_failures + result.warnings + result.infos:
            grouped.setdefault(issue["code"], []).append(result.case_id)

    return {
        code: sorted(case_ids, key=lambda value: (0, int(value)) if value.isdigit() else (1, value.lower()))
        for code, case_ids in sorted(grouped.items(), key=lambda item: item[0])
    }


def _write_markdown_report(summary: dict[str, object], markdown_path: Path) -> None:
    results = summary["cases"]
    lines = [
        "# MSG Golden Corpus Validation Report",
        "",
        f"- Generated: `{summary['generated_at']}`",
        f"- Cases dir: `{summary['cases_dir']}`",
        f"- Render strategy: `{summary['render_strategy']}`",
        f"- Total cases: `{summary['case_count']}`",
        f"- Passed: `{summary['passed_count']}`",
        f"- Failed: `{summary['failed_count']}`",
        f"- Cases with warnings: `{summary['warning_case_count']}`",
        f"- Cases with info: `{summary['info_case_count']}`",
        "",
        "## Case Results",
        "",
        "| Case | Status | Pipeline | Hard failures | Warnings | Info | Anchor hit ratio |",
        "| --- | --- | --- | ---: | ---: | ---: | ---: |",
    ]

    for case in results:
        status = "PASS" if case["passed"] else "FAIL"
        lines.append(
            f"| {case['case_id']} | {status} | {case['pipeline'] or '-'} | "
            f"{len(case['hard_failures'])} | {len(case['warnings'])} | {len(case['infos'])} | "
            f"{case['body_anchor_hit_ratio']:.2f} |"
        )

    lines.extend(["", "## Issue Groups", ""])
    issue_counts = summary["issue_counts"]
    if issue_counts:
        lines.extend(
            [
                "| Issue code | Count | Cases |",
                "| --- | ---: | --- |",
            ]
        )
        issue_case_ids = summary["issue_case_ids"]
        for code, count in issue_counts.items():
            case_ids = ", ".join(issue_case_ids.get(code, []))
            lines.append(f"| {code} | {count} | {case_ids} |")
    else:
        lines.append("- No hard failures or warnings.")

    failed_cases = [case for case in results if not case["passed"]]
    if failed_cases:
        lines.extend(["", "## Failed Cases", ""])
        for case in failed_cases:
            lines.append(f"### Case {case['case_id']}")
            lines.append(f"- Pipeline: `{case['pipeline'] or '-'}`")
            lines.append(f"- Message: `{Path(case['msg_path']).name}`")
            lines.append(f"- Generated PDF: `{case['generated_pdf_path'] or '-'}`")
            for issue in case["hard_failures"]:
                lines.append(f"- FAIL `{issue['code']}`: {issue['message']}")
            for issue in case["warnings"]:
                lines.append(f"- WARN `{issue['code']}`: {issue['message']}")
            for issue in case["infos"]:
                lines.append(f"- INFO `{issue['code']}`: {issue['message']}")
            lines.append("")

    warning_only_cases = [case for case in results if case["passed"] and case["warnings"]]
    if warning_only_cases:
        lines.extend(["", "## Warning Cases", ""])
        for case in warning_only_cases:
            lines.append(f"### Case {case['case_id']}")
            lines.append(f"- Pipeline: `{case['pipeline'] or '-'}`")
            lines.append(f"- Message: `{Path(case['msg_path']).name}`")
            lines.append(f"- Generated PDF: `{case['generated_pdf_path'] or '-'}`")
            for issue in case["warnings"]:
                lines.append(f"- WARN `{issue['code']}`: {issue['message']}")
            lines.append("")

    info_only_cases = [case for case in results if case["passed"] and case["infos"]]
    if info_only_cases:
        lines.extend(["", "## Info Cases", ""])
        for case in info_only_cases:
            lines.append(f"### Case {case['case_id']}")
            lines.append(f"- Pipeline: `{case['pipeline'] or '-'}`")
            lines.append(f"- Message: `{Path(case['msg_path']).name}`")
            lines.append(f"- Generated PDF: `{case['generated_pdf_path'] or '-'}`")
            for issue in case["infos"]:
                lines.append(f"- INFO `{issue['code']}`: {issue['message']}")
            lines.append("")

    markdown_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def validate_corpus(
    cases_dir: Path,
    output_root: Path,
    *,
    render_strategy: str = RENDER_STRATEGY_FIDELITY,
    case_id: str | None = None,
    fail_on_warnings: bool = False,
) -> tuple[dict[str, object], Path, Path]:
    case_pairs = discover_case_pairs(cases_dir, case_id=case_id)

    output_root = output_root.resolve()
    output_root.mkdir(parents=True, exist_ok=True)
    run_stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    run_dir = output_root / f"validation-{run_stamp}"
    run_dir.mkdir(parents=True, exist_ok=True)
    generated_root = run_dir / "generated"
    generated_root.mkdir(parents=True, exist_ok=True)

    previous_strategy = os.environ.get("MSG_TO_PDF_RENDER_STRATEGY")
    os.environ["MSG_TO_PDF_RENDER_STRATEGY"] = render_strategy

    try:
        results = [validate_case_pair(case_pair, generated_root) for case_pair in case_pairs]
    finally:
        if previous_strategy is None:
            os.environ.pop("MSG_TO_PDF_RENDER_STRATEGY", None)
        else:
            os.environ["MSG_TO_PDF_RENDER_STRATEGY"] = previous_strategy

    passed_count = sum(1 for result in results if result.passed)
    warning_case_count = sum(1 for result in results if result.warnings)
    info_case_count = sum(1 for result in results if result.infos)
    failed_count = len(results) - passed_count
    issue_counts = _group_issue_counts(results)
    issue_case_ids = _group_case_ids_by_issue(results)
    anchor_ratios = [result.body_anchor_hit_ratio for result in results]
    strict_failed_count = failed_count + warning_case_count if fail_on_warnings else failed_count

    summary: dict[str, object] = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "cases_dir": str(cases_dir.resolve()),
        "output_dir": str(run_dir),
        "render_strategy": render_strategy,
        "selected_case": case_id,
        "fail_on_warnings": fail_on_warnings,
        "case_count": len(results),
        "passed_count": passed_count,
        "failed_count": failed_count,
        "warning_case_count": warning_case_count,
        "info_case_count": info_case_count,
        "strict_failed_count": strict_failed_count,
        "avg_anchor_hit_ratio": round(mean(anchor_ratios), 4) if anchor_ratios else 0.0,
        "issue_counts": issue_counts,
        "issue_case_ids": issue_case_ids,
        "cases": [asdict(result) for result in results],
    }

    json_path = run_dir / "summary.json"
    markdown_path = run_dir / "summary.md"
    json_path.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    _write_markdown_report(summary, markdown_path)
    return summary, json_path, markdown_path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Validate paired .msg/.pdf golden corpus cases.")
    parser.add_argument(
        "--cases-dir",
        type=Path,
        default=DEFAULT_CASES_DIR,
        help=f"Directory containing numbered case folders (default: {DEFAULT_CASES_DIR})",
    )
    parser.add_argument(
        "--output-root",
        type=Path,
        default=DEFAULT_OUTPUT_ROOT,
        help=f"Directory where reports are written (default: {DEFAULT_OUTPUT_ROOT})",
    )
    parser.add_argument(
        "--render-strategy",
        choices=[RENDER_STRATEGY_FIDELITY, RENDER_STRATEGY_FAST],
        default=RENDER_STRATEGY_FIDELITY,
        help="Render strategy for validation runs (default: fidelity).",
    )
    parser.add_argument(
        "--case",
        dest="case_id",
        help="Optional numbered case folder to validate in isolation.",
    )
    parser.add_argument(
        "--fail-on-warnings",
        action="store_true",
        help="Return a non-zero exit code when any case has warnings.",
    )
    args = parser.parse_args(argv)

    summary, json_path, markdown_path = validate_corpus(
        args.cases_dir,
        args.output_root,
        render_strategy=args.render_strategy,
        case_id=args.case_id,
        fail_on_warnings=args.fail_on_warnings,
    )
    print(f"Wrote JSON report: {json_path}")
    print(f"Wrote Markdown report: {markdown_path}")
    print(
        f"Validation summary: passed={summary['passed_count']}, "
        f"failed={summary['failed_count']}, warnings={summary['warning_case_count']}"
    )
    if summary["failed_count"]:
        return 1
    if args.fail_on_warnings and summary["warning_case_count"]:
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
