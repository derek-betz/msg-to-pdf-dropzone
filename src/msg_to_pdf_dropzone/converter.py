from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field
import inspect
from pathlib import Path
from time import perf_counter

from .msg_parser import parse_msg_file
from .pdf_writer import PdfWriteDiagnostics, write_email_pdf
from .task_events import TaskEventSink, TaskMetaValue, default_task_id_for_path, emit_task_event
from .thread_logic import build_pdf_filename, get_latest_thread_dates, make_unique_path

MAX_FILES_PER_BATCH = 10


class ConversionError(Exception):
    pass


@dataclass(slots=True)
class FileTimingRecord:
    file_name: str
    parse_seconds: float
    filename_seconds: float
    pdf_seconds: float
    total_seconds: float
    pipeline: str = ""
    success: bool = True
    error: str = ""
    stage_seconds: dict[str, float] = field(default_factory=dict)
    image_metrics: dict[str, int] = field(default_factory=dict)


@dataclass(slots=True)
class ConversionResult:
    requested_count: int = 0
    converted_files: list[Path] = field(default_factory=list)
    skipped_files: list[Path] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    parse_seconds: float = 0.0
    thread_group_seconds: float = 0.0
    write_seconds: float = 0.0
    total_seconds: float = 0.0
    timing_lines: list[str] = field(default_factory=list)
    file_timing_records: list[FileTimingRecord] = field(default_factory=list)


def _format_seconds(value: float) -> str:
    return f"{value:.2f}s"


def _supports_keyword_argument(function_object: object, argument_name: str) -> bool:
    try:
        signature = inspect.signature(function_object)
    except (TypeError, ValueError):
        return False
    return argument_name in signature.parameters


def _resolve_task_id(
    source_path: Path,
    task_ids_by_source_path: Mapping[Path, str] | None,
) -> str:
    resolved = source_path.resolve()
    if task_ids_by_source_path is not None:
        mapped = task_ids_by_source_path.get(resolved)
        if mapped:
            return mapped
    return default_task_id_for_path(resolved)


def _normalize_event_pipeline_name(pipeline: str) -> str | None:
    if pipeline == "reportlab_fast":
        return "reportlab"
    if pipeline in {"outlook_edge", "edge_html", "reportlab"}:
        return pipeline
    return None


def _resolve_event_meta(
    source_path: Path,
    batch_meta_by_source_path: Mapping[Path, Mapping[str, TaskMetaValue]] | None,
) -> dict[str, TaskMetaValue]:
    resolved = source_path.resolve()
    if batch_meta_by_source_path is None:
        return {}
    raw_meta = batch_meta_by_source_path.get(resolved)
    if raw_meta is None:
        return {}
    return dict(raw_meta)


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


def convert_msg_files(
    msg_paths: list[Path],
    output_dir: Path,
    *,
    event_sink: TaskEventSink | None = None,
    task_ids_by_source_path: Mapping[Path, str] | None = None,
    batch_meta_by_source_path: Mapping[Path, Mapping[str, TaskMetaValue]] | None = None,
) -> ConversionResult:
    started_at = perf_counter()
    result = ConversionResult(requested_count=len(msg_paths))
    if not msg_paths:
        raise ConversionError("No .msg files were provided.")
    if len(msg_paths) > MAX_FILES_PER_BATCH:
        raise ConversionError(f"Select up to {MAX_FILES_PER_BATCH} files per batch.")

    email_records = []
    parse_durations: dict[Path, float] = {}
    timing_lines: list[str] = []
    supports_diagnostics = _supports_keyword_argument(write_email_pdf, "diagnostics")
    supports_event_sink = _supports_keyword_argument(write_email_pdf, "event_sink")
    supports_task_id = _supports_keyword_argument(write_email_pdf, "task_id")
    supports_event_meta = _supports_keyword_argument(write_email_pdf, "event_meta")

    for msg_path in msg_paths:
        if msg_path.suffix and msg_path.suffix.lower() != ".msg":
            result.skipped_files.append(msg_path)
            result.errors.append(f"Skipped non-msg file: {msg_path.name}")
            continue

        task_id = _resolve_task_id(msg_path, task_ids_by_source_path)
        event_meta = _resolve_event_meta(msg_path, batch_meta_by_source_path)
        emit_task_event(
            event_sink,
            task_id=task_id,
            stage="parse_started",
            file_name=msg_path.name,
            meta=event_meta,
        )
        parse_started_at = perf_counter()
        try:
            record = parse_msg_file(msg_path)
            elapsed = perf_counter() - parse_started_at
            parse_durations[msg_path] = elapsed
            email_records.append(record)
        except Exception as exc:  # pragma: no cover - external parser failure branch
            elapsed = perf_counter() - parse_started_at
            parse_durations[msg_path] = elapsed
            result.skipped_files.append(msg_path)
            result.errors.append(f"Failed to parse {msg_path.name}: {exc}")
            timing_lines.append(
                f"{msg_path.name}: parse failed after {_format_seconds(elapsed)}"
            )
            emit_task_event(
                event_sink,
                task_id=task_id,
                stage="failed",
                file_name=msg_path.name,
                success=False,
                error=str(exc),
                meta=event_meta,
            )

    result.parse_seconds = sum(parse_durations.values())

    if not email_records:
        result.total_seconds = perf_counter() - started_at
        timing_lines.insert(
            0,
            (
                f"Total {_format_seconds(result.total_seconds)} "
                f"(parse {_format_seconds(result.parse_seconds)})"
            ),
        )
        result.timing_lines = timing_lines
        return result

    thread_started_at = perf_counter()
    latest_thread_dates = get_latest_thread_dates(email_records)
    result.thread_group_seconds = perf_counter() - thread_started_at

    for record in email_records:
        naming_seconds = 0.0
        write_seconds = 0.0
        diagnostics = PdfWriteDiagnostics()
        task_id = _resolve_task_id(record.source_path, task_ids_by_source_path)
        event_meta = _resolve_event_meta(record.source_path, batch_meta_by_source_path)
        try:
            naming_started_at = perf_counter()
            file_name = build_pdf_filename(record.subject, latest_thread_dates[record.thread_key])
            output_path = make_unique_path(output_dir / file_name)
            naming_seconds = perf_counter() - naming_started_at
            emit_task_event(
                event_sink,
                task_id=task_id,
                stage="filename_built",
                file_name=record.source_path.name,
                meta=_merge_event_meta(event_meta, {"outputName": output_path.name}),
            )

            write_started_at = perf_counter()
            emit_task_event(
                event_sink,
                task_id=task_id,
                stage="pdf_pipeline_started",
                file_name=record.source_path.name,
                meta=_merge_event_meta(event_meta, {"outputName": output_path.name}),
            )
            write_kwargs: dict[str, object] = {}
            if supports_diagnostics:
                write_kwargs["diagnostics"] = diagnostics
            if supports_event_sink:
                write_kwargs["event_sink"] = event_sink
            if supports_task_id:
                write_kwargs["task_id"] = task_id
            if supports_event_meta:
                write_kwargs["event_meta"] = event_meta
            write_email_pdf(record, output_path, **write_kwargs)
            write_seconds = perf_counter() - write_started_at
            result.write_seconds += write_seconds
            result.converted_files.append(output_path)
            pipeline_name = _normalize_event_pipeline_name(diagnostics.pipeline)
            emit_task_event(
                event_sink,
                task_id=task_id,
                stage="pdf_written",
                file_name=record.source_path.name,
                pipeline=pipeline_name,
                meta=_merge_event_meta(
                    event_meta,
                    {
                        "outputName": output_path.name,
                        "outputPath": str(output_path),
                    },
                ),
            )
            emit_task_event(
                event_sink,
                task_id=task_id,
                stage="deliver_started",
                file_name=record.source_path.name,
                pipeline=pipeline_name,
                meta=_merge_event_meta(
                    event_meta,
                    {
                        "outputName": output_path.name,
                        "outputDir": str(output_dir),
                        "outputDirLabel": output_dir.name or str(output_dir),
                    },
                ),
            )
            emit_task_event(
                event_sink,
                task_id=task_id,
                stage="complete",
                file_name=record.source_path.name,
                pipeline=pipeline_name,
                success=True,
                meta=_merge_event_meta(
                    event_meta,
                    {
                        "outputName": output_path.name,
                        "outputPath": str(output_path),
                    },
                ),
            )

            parse_seconds = parse_durations.get(record.source_path, 0.0)
            file_total = parse_seconds + naming_seconds + write_seconds
            line = (
                f"{record.source_path.name}: parse {_format_seconds(parse_seconds)}, "
                f"filename {_format_seconds(naming_seconds)}, "
                f"pdf {_format_seconds(write_seconds)}, "
                f"total {_format_seconds(file_total)}"
            )
            if diagnostics.pipeline:
                line = f"{line} [pipeline {diagnostics.pipeline}]"

            stage_bits = []
            for stage_name, stage_value in diagnostics.stage_seconds.items():
                stage_bits.append(f"{stage_name} {_format_seconds(stage_value)}")
            if stage_bits:
                line = f"{line} ({', '.join(stage_bits)})"
            image_bits = ", ".join(
                f"{name} {value}" for name, value in diagnostics.image_metrics.items()
            )
            if image_bits:
                line = f"{line} [images {image_bits}]"
            timing_lines.append(line)
            result.file_timing_records.append(
                FileTimingRecord(
                    file_name=record.source_path.name,
                    parse_seconds=parse_seconds,
                    filename_seconds=naming_seconds,
                    pdf_seconds=write_seconds,
                    total_seconds=file_total,
                    pipeline=diagnostics.pipeline,
                    success=True,
                    error="",
                    stage_seconds=dict(diagnostics.stage_seconds),
                    image_metrics=dict(diagnostics.image_metrics),
                )
            )
        except Exception as exc:  # pragma: no cover - external PDF writer failure branch
            result.skipped_files.append(record.source_path)
            result.errors.append(f"Failed to convert {record.source_path.name}: {exc}")
            parse_seconds = parse_durations.get(record.source_path, 0.0)
            failure_total = parse_seconds + naming_seconds + write_seconds
            timing_lines.append(
                f"{record.source_path.name}: conversion failed after {_format_seconds(failure_total)}"
            )
            result.file_timing_records.append(
                FileTimingRecord(
                    file_name=record.source_path.name,
                    parse_seconds=parse_seconds,
                    filename_seconds=naming_seconds,
                    pdf_seconds=write_seconds,
                    total_seconds=failure_total,
                    pipeline="",
                    success=False,
                    error=str(exc),
                    stage_seconds=dict(diagnostics.stage_seconds),
                    image_metrics=dict(diagnostics.image_metrics),
                )
            )
            emit_task_event(
                event_sink,
                task_id=task_id,
                stage="failed",
                file_name=record.source_path.name,
                pipeline=_normalize_event_pipeline_name(diagnostics.pipeline),
                success=False,
                error=str(exc),
                meta=event_meta,
            )

    result.total_seconds = perf_counter() - started_at
    timing_lines.insert(
        0,
        (
            f"Total {_format_seconds(result.total_seconds)} "
            f"(parse {_format_seconds(result.parse_seconds)}, "
            f"thread {_format_seconds(result.thread_group_seconds)}, "
            f"pdf {_format_seconds(result.write_seconds)})"
        ),
    )
    result.timing_lines = timing_lines
    return result
