from __future__ import annotations

import argparse
import json
import math
import os
import shutil
from datetime import datetime
from pathlib import Path
from statistics import mean
from time import perf_counter

from .converter import ConversionResult, FileTimingRecord, convert_msg_files
from .msg_parser import parse_msg_file
from .pdf_writer import RENDER_STRATEGY_FAST, RENDER_STRATEGY_FIDELITY

DEFAULT_EMAILS_DIR = Path("emails-for-testing")
DEFAULT_OUTPUT_ROOT = Path(".local-corpus-profiles")
DEFAULT_RUN_COUNT = 3


def _percentile(values: list[float], percentile: float) -> float:
    if not values:
        return 0.0
    ordered = sorted(values)
    index = max(0, min(len(ordered) - 1, math.ceil(percentile * len(ordered)) - 1))
    return ordered[index]


def _build_aggregate(metrics: list[dict[str, object]]) -> dict[str, object]:
    total_values = [float(item["convert_total_seconds"]) for item in metrics]
    parse_values = [float(item["parse_seconds"]) for item in metrics]
    write_values = [float(item["write_seconds"]) for item in metrics]
    failure_count = sum(1 for item in metrics if not bool(item["convert_ok"]))
    image_metric_names = (
        "total_images",
        "cid_resolved",
        "cid_unresolved",
        "signature_small_dropped",
        "remote_dropped",
    )
    image_aggregate: dict[str, dict[str, float]] = {}
    for name in image_metric_names:
        values = [
            int((item.get("image_metrics") or {}).get(name, 0))  # type: ignore[union-attr]
            for item in metrics
        ]
        image_aggregate[name] = {
            "sum": int(sum(values)),
            "avg": round(mean(values), 4) if values else 0.0,
        }

    return {
        "count": len(metrics),
        "failure_count": failure_count,
        "total_seconds": {
            "min": round(min(total_values), 4) if total_values else 0.0,
            "avg": round(mean(total_values), 4) if total_values else 0.0,
            "p95": round(_percentile(total_values, 0.95), 4) if total_values else 0.0,
        },
        "parse_seconds": {
            "min": round(min(parse_values), 4) if parse_values else 0.0,
            "avg": round(mean(parse_values), 4) if parse_values else 0.0,
            "p95": round(_percentile(parse_values, 0.95), 4) if parse_values else 0.0,
        },
        "write_seconds": {
            "min": round(min(write_values), 4) if write_values else 0.0,
            "avg": round(mean(write_values), 4) if write_values else 0.0,
            "p95": round(_percentile(write_values, 0.95), 4) if write_values else 0.0,
        },
        "image_metrics": image_aggregate,
    }


def _run_single_file(path: Path, output_dir: Path) -> dict[str, object]:
    parse_started_at = perf_counter()
    parse_ok = True
    parse_error = ""
    attachment_count = 0
    html_length = 0
    body_length = 0
    thread_key = ""
    subject = ""
    try:
        record = parse_msg_file(path)
        attachment_count = len(record.attachment_names)
        html_length = len(record.html_body)
        body_length = len(record.body)
        thread_key = record.thread_key
        subject = record.subject
    except Exception as exc:  # pragma: no cover - external parser failures
        parse_ok = False
        parse_error = str(exc)
    parse_seconds = perf_counter() - parse_started_at

    convert_started_at = perf_counter()
    conversion: ConversionResult = convert_msg_files([path], output_dir)
    convert_wall_seconds = perf_counter() - convert_started_at

    file_timing = conversion.file_timing_records[0] if conversion.file_timing_records else None
    if file_timing is None:
        file_timing = FileTimingRecord(
            file_name=path.name,
            parse_seconds=conversion.parse_seconds,
            filename_seconds=0.0,
            pdf_seconds=conversion.write_seconds,
            total_seconds=conversion.total_seconds,
            pipeline="",
            success=len(conversion.converted_files) > 0,
            error="; ".join(conversion.errors),
            stage_seconds={},
        )

    for output_file in conversion.converted_files:
        try:
            output_file.unlink()
        except Exception:
            pass

    return {
        "file_name": path.name,
        "size_bytes": path.stat().st_size,
        "parse_ok": parse_ok,
        "parse_error": parse_error,
        "parse_seconds": round(parse_seconds, 4),
        "subject": subject,
        "thread_key": thread_key,
        "body_length": body_length,
        "html_length": html_length,
        "attachment_count": attachment_count,
        "convert_ok": file_timing.success and len(conversion.converted_files) > 0,
        "convert_total_seconds": round(conversion.total_seconds, 4),
        "convert_wall_seconds": round(convert_wall_seconds, 4),
        "write_seconds": round(file_timing.pdf_seconds, 4),
        "filename_seconds": round(file_timing.filename_seconds, 4),
        "pipeline": file_timing.pipeline,
        "stage_seconds": {name: round(value, 4) for name, value in file_timing.stage_seconds.items()},
        "image_metrics": {name: int(value) for name, value in file_timing.image_metrics.items()},
        "errors": conversion.errors,
        "timing_lines": conversion.timing_lines,
    }


def _write_markdown_report(summary: dict[str, object], markdown_path: Path) -> None:
    aggregate = summary["aggregate"]
    runs = summary["runs"]

    lines = [
        "# MSG Corpus Profiling Report",
        "",
        f"- Generated: `{summary['generated_at']}`",
        f"- Corpus: `{summary['emails_dir']}`",
        f"- Strategy: `{summary['render_strategy']}`",
        f"- Runs: `{summary['run_count']}`",
        f"- Files per run: `{summary['file_count']}`",
        "",
        "## Aggregate Metrics",
        "",
        "| Metric | Min | Avg | P95 |",
        "| --- | ---: | ---: | ---: |",
        (
            f"| Total (s) | {aggregate['total_seconds']['min']:.4f} | "
            f"{aggregate['total_seconds']['avg']:.4f} | {aggregate['total_seconds']['p95']:.4f} |"
        ),
        (
            f"| Parse (s) | {aggregate['parse_seconds']['min']:.4f} | "
            f"{aggregate['parse_seconds']['avg']:.4f} | {aggregate['parse_seconds']['p95']:.4f} |"
        ),
        (
            f"| Write (s) | {aggregate['write_seconds']['min']:.4f} | "
            f"{aggregate['write_seconds']['avg']:.4f} | {aggregate['write_seconds']['p95']:.4f} |"
        ),
        f"| Failures | {aggregate['failure_count']} | - | - |",
        "",
        "## Image Metrics (Aggregate)",
        "",
        "| Metric | Sum | Avg per file |",
        "| --- | ---: | ---: |",
        (
            f"| total_images | {aggregate['image_metrics']['total_images']['sum']} | "
            f"{aggregate['image_metrics']['total_images']['avg']:.4f} |"
        ),
        (
            f"| cid_resolved | {aggregate['image_metrics']['cid_resolved']['sum']} | "
            f"{aggregate['image_metrics']['cid_resolved']['avg']:.4f} |"
        ),
        (
            f"| cid_unresolved | {aggregate['image_metrics']['cid_unresolved']['sum']} | "
            f"{aggregate['image_metrics']['cid_unresolved']['avg']:.4f} |"
        ),
        (
            f"| signature_small_dropped | {aggregate['image_metrics']['signature_small_dropped']['sum']} | "
            f"{aggregate['image_metrics']['signature_small_dropped']['avg']:.4f} |"
        ),
        (
            f"| remote_dropped | {aggregate['image_metrics']['remote_dropped']['sum']} | "
            f"{aggregate['image_metrics']['remote_dropped']['avg']:.4f} |"
        ),
        "",
        "## Per-Run Summary",
        "",
        "| Run | Converted | Failures | Avg Total (s) | P95 Total (s) |",
        "| --- | ---: | ---: | ---: | ---: |",
    ]

    for run in runs:
        converted = sum(1 for item in run["files"] if item["convert_ok"])
        failures = sum(1 for item in run["files"] if not item["convert_ok"])
        totals = [float(item["convert_total_seconds"]) for item in run["files"]]
        lines.append(
            f"| {run['run_index']} | {converted} | {failures} | "
            f"{mean(totals):.4f} | {_percentile(totals, 0.95):.4f} |"
        )

    lines.extend(
        [
            "",
            "## Acceptance Snapshot",
            "",
            f"- Parse failures: `{summary['acceptance']['parse_failures']}`",
            f"- Conversion failures: `{summary['acceptance']['conversion_failures']}`",
            f"- Conversion p95 total: `{summary['acceptance']['conversion_p95_total_seconds']:.4f}s`",
            (
                f"- UI-ready target (<2s) for drop dispatch is validated by automated app drop tests "
                f"(not measured by this corpus profiler)."
            ),
        ]
    )

    markdown_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def profile_corpus(
    emails_dir: Path,
    output_root: Path,
    *,
    runs: int = DEFAULT_RUN_COUNT,
    render_strategy: str = RENDER_STRATEGY_FIDELITY,
) -> tuple[dict[str, object], Path, Path]:
    emails_dir = emails_dir.resolve()
    if not emails_dir.exists():
        raise FileNotFoundError(f"Emails directory does not exist: {emails_dir}")

    files = sorted(path for path in emails_dir.glob("*.msg") if path.is_file())
    if not files:
        raise ValueError(f"No .msg files found in: {emails_dir}")

    output_root = output_root.resolve()
    output_root.mkdir(parents=True, exist_ok=True)
    run_stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    run_dir = output_root / f"profile-{run_stamp}"
    run_dir.mkdir(parents=True, exist_ok=True)

    previous_strategy = os.environ.get("MSG_TO_PDF_RENDER_STRATEGY")
    os.environ["MSG_TO_PDF_RENDER_STRATEGY"] = render_strategy

    runs_data: list[dict[str, object]] = []
    flattened_metrics: list[dict[str, object]] = []
    try:
        for run_index in range(1, runs + 1):
            run_output_dir = run_dir / f"run-{run_index}-pdf-output"
            run_output_dir.mkdir(parents=True, exist_ok=True)

            run_rows: list[dict[str, object]] = []
            for file_path in files:
                row = _run_single_file(file_path, run_output_dir)
                run_rows.append(row)
                flattened_metrics.append(row)

            runs_data.append(
                {
                    "run_index": run_index,
                    "files": run_rows,
                }
            )
            shutil.rmtree(run_output_dir, ignore_errors=True)
    finally:
        if previous_strategy is None:
            os.environ.pop("MSG_TO_PDF_RENDER_STRATEGY", None)
        else:
            os.environ["MSG_TO_PDF_RENDER_STRATEGY"] = previous_strategy

    aggregate = _build_aggregate(flattened_metrics)
    parse_failures = sum(1 for item in flattened_metrics if not item["parse_ok"])
    conversion_failures = sum(1 for item in flattened_metrics if not item["convert_ok"])

    summary: dict[str, object] = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "emails_dir": str(emails_dir),
        "output_dir": str(run_dir),
        "render_strategy": render_strategy,
        "run_count": runs,
        "file_count": len(files),
        "aggregate": aggregate,
        "acceptance": {
            "parse_failures": parse_failures,
            "conversion_failures": conversion_failures,
            "conversion_p95_total_seconds": aggregate["total_seconds"]["p95"],
            "drop_ui_ready_target_seconds": 2.0,
            "drop_ui_ready_measured_by": "automated app drop responsiveness tests",
        },
        "runs": runs_data,
    }

    json_path = run_dir / "summary.json"
    markdown_path = run_dir / "summary.md"
    json_path.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    _write_markdown_report(summary, markdown_path)
    return summary, json_path, markdown_path


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Profile local .msg corpus parse/conversion timings.")
    parser.add_argument(
        "--emails-dir",
        type=Path,
        default=DEFAULT_EMAILS_DIR,
        help=f"Directory containing .msg files (default: {DEFAULT_EMAILS_DIR})",
    )
    parser.add_argument(
        "--output-root",
        type=Path,
        default=DEFAULT_OUTPUT_ROOT,
        help=f"Directory where reports are written (default: {DEFAULT_OUTPUT_ROOT})",
    )
    parser.add_argument(
        "--runs",
        type=int,
        default=DEFAULT_RUN_COUNT,
        help=f"How many repeated corpus runs to execute (default: {DEFAULT_RUN_COUNT})",
    )
    parser.add_argument(
        "--render-strategy",
        choices=[RENDER_STRATEGY_FIDELITY, RENDER_STRATEGY_FAST],
        default=RENDER_STRATEGY_FIDELITY,
        help="Render strategy for profiling comparisons (default: fidelity).",
    )
    args = parser.parse_args(argv)

    if args.runs <= 0:
        raise ValueError("--runs must be >= 1")

    summary, json_path, markdown_path = profile_corpus(
        args.emails_dir,
        args.output_root,
        runs=args.runs,
        render_strategy=args.render_strategy,
    )
    print(f"Wrote JSON report: {json_path}")
    print(f"Wrote Markdown report: {markdown_path}")
    print(
        "Acceptance snapshot: "
        f"parse_failures={summary['acceptance']['parse_failures']}, "
        f"conversion_failures={summary['acceptance']['conversion_failures']}, "
        f"p95_total={summary['acceptance']['conversion_p95_total_seconds']:.4f}s"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
