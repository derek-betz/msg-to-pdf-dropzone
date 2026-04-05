from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from contextlib import ExitStack
from dataclasses import asdict
from datetime import datetime
import json
import os
from pathlib import Path
import re
import socket
import subprocess
import sys
import threading
import time
from typing import Any

import httpx

from .corpus_validator import (
    analyze_pdf,
    compare_snapshots,
    discover_case_pairs,
)
from .pdf_writer import RENDER_STRATEGY_FAST, RENDER_STRATEGY_FIDELITY

DEFAULT_CASES_DIR = Path("emails-for-testing")
DEFAULT_OUTPUT_ROOT = Path(".local-browser-run")
DEFAULT_HOST = "127.0.0.1"
OUTPUT_NAME_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}_.+\.pdf$")
NUMERIC_SUFFIX_PATTERN = re.compile(r" \(\d+\)\.pdf$", re.IGNORECASE)
EXPECTED_STAGE_ORDER = [
    "drop_received",
    "outlook_extract_started",
    "files_accepted",
    "output_folder_selected",
    "parse_started",
    "filename_built",
    "pdf_pipeline_started",
    "pipeline_selected",
    "pdf_written",
    "deliver_started",
    "complete",
    "failed",
]
EXPECTED_STAGE_INDEX = {stage: index for index, stage in enumerate(EXPECTED_STAGE_ORDER)}


def _pick_free_port(host: str) -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind((host, 0))
        return int(sock.getsockname()[1])


def _wait_for_server(base_url: str, *, timeout_seconds: float) -> None:
    deadline = time.time() + timeout_seconds
    last_error = "server did not become healthy"
    with httpx.Client(timeout=httpx.Timeout(5.0, connect=2.0)) as client:
        while time.time() < deadline:
            try:
                response = client.get(f"{base_url}/api/health")
                response.raise_for_status()
                if response.json().get("ok") is True:
                    return
                last_error = f"unexpected health payload: {response.text}"
            except Exception as exc:
                last_error = str(exc)
            time.sleep(0.5)
    raise RuntimeError(f"Timed out waiting for browser server at {base_url}: {last_error}")


def _event_listener(
    base_url: str,
    events: list[dict[str, Any]],
    connected: threading.Event,
    stop_requested: threading.Event,
    errors: list[str],
) -> None:
    try:
        timeout = httpx.Timeout(connect=5.0, read=None, write=30.0, pool=30.0)
        with httpx.Client(timeout=timeout) as client:
            with client.stream("GET", f"{base_url}/api/events") as response:
                response.raise_for_status()
                connected.set()
                for raw_line in response.iter_lines():
                    if stop_requested.is_set():
                        break
                    if not raw_line:
                        continue
                    line = raw_line.strip()
                    if not line or line.startswith(":") or line.startswith("retry:"):
                        continue
                    if not line.startswith("data:"):
                        continue
                    payload = json.loads(line[5:].strip())
                    if isinstance(payload, dict):
                        events.append(payload)
    except Exception as exc:
        errors.append(str(exc))
        connected.set()


def summarize_task_events(events: list[dict[str, Any]], expected_task_ids: list[str]) -> dict[str, Any]:
    stage_events = [event for event in events if isinstance(event, dict) and event.get("stage")]
    events_by_task_id: dict[str, list[dict[str, Any]]] = defaultdict(list)
    pipeline_counts: Counter[str] = Counter()
    output_paths_by_task_id: dict[str, str] = {}
    output_names_by_task_id: dict[str, str] = {}
    failures: list[dict[str, str]] = []
    order_issues: list[dict[str, str]] = []
    missing_terminal_tasks: list[str] = []

    for event in stage_events:
        task_id = str(event.get("taskId", ""))
        if not task_id:
            continue
        events_by_task_id[task_id].append(event)
        meta = event.get("meta") or {}
        output_name = meta.get("outputName")
        if isinstance(output_name, str) and output_name:
            output_names_by_task_id[task_id] = output_name
        output_path = meta.get("outputPath")
        if isinstance(output_path, str) and output_path:
            output_paths_by_task_id[task_id] = output_path

    completed_count = 0
    failed_count = 0
    pipelines_by_task_id: dict[str, str] = {}
    for task_id in expected_task_ids:
        task_events = events_by_task_id.get(task_id, [])
        last_stage_index = -1
        terminal_stage = ""
        pipeline_name = ""
        for event in task_events:
            stage = str(event.get("stage", ""))
            stage_index = EXPECTED_STAGE_INDEX.get(stage, last_stage_index)
            if stage_index < last_stage_index:
                order_issues.append(
                    {
                        "taskId": task_id,
                        "stage": stage,
                        "message": f"Stage '{stage}' arrived after a later stage.",
                    }
                )
            last_stage_index = max(last_stage_index, stage_index)
            event_pipeline = event.get("pipeline")
            if isinstance(event_pipeline, str) and event_pipeline:
                pipeline_name = event_pipeline
            if stage in {"complete", "failed"}:
                terminal_stage = stage
        if pipeline_name:
            pipelines_by_task_id[task_id] = pipeline_name
            pipeline_counts[pipeline_name] += 1
        if terminal_stage == "complete":
            completed_count += 1
        elif terminal_stage == "failed":
            failed_count += 1
            failure_event = next(
                (event for event in reversed(task_events) if event.get("stage") == "failed"),
                None,
            )
            failures.append(
                {
                    "taskId": task_id,
                    "fileName": str((failure_event or {}).get("fileName", "")),
                    "error": str((failure_event or {}).get("error", "")),
                }
            )
        else:
            missing_terminal_tasks.append(task_id)

    output_names = [output_names_by_task_id[task_id] for task_id in expected_task_ids if task_id in output_names_by_task_id]
    numeric_suffix_outputs = sorted(name for name in output_names if NUMERIC_SUFFIX_PATTERN.search(name))

    return {
        "taskCount": len(expected_task_ids),
        "stageEventCount": len(stage_events),
        "completedCount": completed_count,
        "failedCount": failed_count,
        "pipelineCounts": dict(sorted(pipeline_counts.items())),
        "outputNamesByTaskId": output_names_by_task_id,
        "outputPathsByTaskId": output_paths_by_task_id,
        "allOutputNamesMatchPattern": bool(output_names) and all(OUTPUT_NAME_PATTERN.match(name) for name in output_names),
        "numericSuffixOutputs": numeric_suffix_outputs,
        "failures": failures,
        "missingTerminalTasks": sorted(missing_terminal_tasks),
        "orderIssues": order_issues,
        "pipelinesByTaskId": pipelines_by_task_id,
    }


def validate_generated_outputs(
    *,
    case_pairs: list[Any],
    case_by_task_id: dict[str, Any],
    event_summary: dict[str, Any],
) -> dict[str, Any]:
    output_paths_by_task_id: dict[str, str] = event_summary["outputPathsByTaskId"]
    pipelines_by_task_id: dict[str, str] = event_summary["pipelinesByTaskId"]
    results = []

    for case_pair in case_pairs:
        task_id = next(task_id for task_id, mapped_case in case_by_task_id.items() if mapped_case.case_id == case_pair.case_id)
        output_path_value = output_paths_by_task_id.get(task_id)
        pipeline = pipelines_by_task_id.get(task_id, "")
        if not output_path_value:
            results.append(
                {
                    "case_id": case_pair.case_id,
                    "msg_path": str(case_pair.msg_path),
                    "golden_pdf_path": str(case_pair.golden_pdf_path),
                    "generated_pdf_path": None,
                    "pipeline": pipeline,
                    "passed": False,
                    "hard_failures": [
                        {
                            "code": "missing_output",
                            "message": "Browser validation did not emit an outputPath for this case.",
                        }
                    ],
                    "warnings": [],
                    "infos": [],
                }
            )
            continue

        generated_pdf_path = Path(output_path_value)
        try:
            golden_snapshot = analyze_pdf(case_pair.golden_pdf_path)
            generated_snapshot = analyze_pdf(generated_pdf_path)
            comparison = compare_snapshots(golden_snapshot, generated_snapshot, pipeline=pipeline)
            comparison.case_id = case_pair.case_id
            comparison.msg_path = str(case_pair.msg_path)
            comparison.golden_pdf_path = str(case_pair.golden_pdf_path)
            comparison.generated_pdf_path = str(generated_pdf_path)
            comparison.pipeline = pipeline
            results.append(asdict(comparison))
        except Exception as exc:
            results.append(
                {
                    "case_id": case_pair.case_id,
                    "msg_path": str(case_pair.msg_path),
                    "golden_pdf_path": str(case_pair.golden_pdf_path),
                    "generated_pdf_path": str(generated_pdf_path),
                    "pipeline": pipeline,
                    "passed": False,
                    "hard_failures": [
                        {
                            "code": "analysis_failure",
                            "message": f"Browser-generated PDF analysis failed: {exc}",
                        }
                    ],
                    "warnings": [],
                    "infos": [],
                }
            )

    passed_count = sum(1 for result in results if result["passed"])
    warning_case_count = sum(1 for result in results if result["warnings"])
    info_case_count = sum(1 for result in results if result["infos"])
    failed_count = len(results) - passed_count
    return {
        "caseCount": len(results),
        "passedCount": passed_count,
        "failedCount": failed_count,
        "warningCaseCount": warning_case_count,
        "infoCaseCount": info_case_count,
        "cases": results,
    }


def _write_markdown(summary: dict[str, Any], output_path: Path) -> None:
    event_summary = summary["eventSummary"]
    validation_summary = summary["validationSummary"]
    lines = [
        "# Browser Golden Corpus Validation",
        "",
        f"- Generated: `{summary['generatedAt']}`",
        f"- Cases dir: `{summary['casesDir']}`",
        f"- Browser URL: `{summary['baseUrl']}`",
        f"- Render strategy: `{summary['renderStrategy']}`",
        f"- Total `.msg` files processed: `{summary['msgFileCount']}`",
        f"- Converted: `{summary['convertedCount']}`",
        f"- Failed tasks: `{event_summary['failedCount']}`",
        f"- Stage events: `{event_summary['stageEventCount']}`",
        f"- Pipelines: `{json.dumps(event_summary['pipelineCounts'], sort_keys=True)}`",
        f"- Semantic validation: `{validation_summary['passedCount']} passed / {validation_summary['failedCount']} failed / {validation_summary['warningCaseCount']} warnings`",
        "",
        "## Event Checks",
        "",
        f"- Missing terminal tasks: `{len(event_summary['missingTerminalTasks'])}`",
        f"- Order issues: `{len(event_summary['orderIssues'])}`",
        f"- Output names match pattern: `{event_summary['allOutputNamesMatchPattern']}`",
        f"- Numeric suffix outputs: `{len(event_summary['numericSuffixOutputs'])}`",
        "",
        "## Case Results",
        "",
        "| Case | Status | Pipeline | Hard failures | Warnings | Info |",
        "| --- | --- | --- | ---: | ---: | ---: |",
    ]
    for case in validation_summary["cases"]:
        status = "PASS" if case["passed"] else "FAIL"
        lines.append(
            f"| {case['case_id']} | {status} | {case.get('pipeline') or '-'} | "
            f"{len(case.get('hard_failures', []))} | {len(case.get('warnings', []))} | {len(case.get('infos', []))} |"
        )
    output_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def run_browser_validation(
    *,
    cases_dir: Path,
    output_root: Path,
    host: str,
    port: int | None,
    base_url: str | None,
    timeout_seconds: float,
    render_strategy: str,
) -> tuple[dict[str, Any], Path, Path]:
    case_pairs = discover_case_pairs(cases_dir)
    msg_file_count = len(case_pairs)
    run_stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    output_root = output_root.resolve()
    output_root.mkdir(parents=True, exist_ok=True)
    run_dir = output_root / f"browser-validation-{run_stamp}"
    run_dir.mkdir(parents=True, exist_ok=True)
    generated_root = run_dir / "generated"
    generated_root.mkdir(parents=True, exist_ok=True)

    server_process: subprocess.Popen[str] | None = None
    server_log = None
    server_err_log = None
    resolved_port = port if port is not None else _pick_free_port(host)
    resolved_base_url = base_url or f"http://{host}:{resolved_port}"

    try:
        if base_url is None:
            server_log = (run_dir / "server.log").open("w", encoding="utf-8")
            server_err_log = (run_dir / "server.err.log").open("w", encoding="utf-8")
            server_env = dict(os.environ)
            server_env["MSG_TO_PDF_RENDER_STRATEGY"] = render_strategy
            server_process = subprocess.Popen(
                [
                    sys.executable,
                    "-m",
                    "msg_to_pdf_dropzone",
                    "--host",
                    host,
                    "--port",
                    str(resolved_port),
                    "--no-browser",
                ],
                stdout=server_log,
                stderr=server_err_log,
                cwd=str(Path.cwd()),
                env=server_env,
                text=True,
            )
            _wait_for_server(resolved_base_url, timeout_seconds=timeout_seconds)

        stream_events: list[dict[str, Any]] = []
        stream_errors: list[str] = []
        connected = threading.Event()
        stop_requested = threading.Event()
        listener = threading.Thread(
            target=_event_listener,
            args=(resolved_base_url, stream_events, connected, stop_requested, stream_errors),
            daemon=True,
        )
        listener.start()
        if not connected.wait(timeout=10):
            raise RuntimeError("Timed out waiting for the browser event stream to connect.")
        if stream_errors:
            raise RuntimeError(stream_errors[0])

        client_timeout = httpx.Timeout(connect=10.0, read=timeout_seconds, write=timeout_seconds, pool=30.0)
        with httpx.Client(timeout=client_timeout) as client:
            with ExitStack() as stack:
                upload_files = []
                for case_pair in case_pairs:
                    handle = stack.enter_context(case_pair.msg_path.open("rb"))
                    upload_files.append(
                        ("files", (case_pair.msg_path.name, handle, "application/vnd.ms-outlook"))
                    )
                upload_response = client.post(f"{resolved_base_url}/api/upload", files=upload_files)
            upload_response.raise_for_status()
            upload_payload = upload_response.json()
            accepted_items = upload_payload["accepted"]
            if len(accepted_items) != len(case_pairs):
                raise RuntimeError(
                    f"Expected {len(case_pairs)} accepted uploads but got {len(accepted_items)}."
                )

            convert_response = client.post(
                f"{resolved_base_url}/api/convert",
                json={
                    "ids": [item["id"] for item in accepted_items],
                    "output_dir": str(generated_root),
                },
            )
            convert_response.raise_for_status()
            convert_payload = convert_response.json()

        expected_task_ids = [str(item["taskId"]) for item in accepted_items]
        case_by_task_id = {
            str(item["taskId"]): case_pair
            for item, case_pair in zip(accepted_items, case_pairs, strict=True)
        }

        deadline = time.time() + timeout_seconds
        while time.time() < deadline:
            event_summary = summarize_task_events(stream_events, expected_task_ids)
            if (
                event_summary["completedCount"] + event_summary["failedCount"] >= len(expected_task_ids)
                and not event_summary["missingTerminalTasks"]
            ):
                break
            time.sleep(0.5)
        else:
            raise RuntimeError("Timed out waiting for terminal browser task events.")

        stop_requested.set()
        listener.join(timeout=2)
        event_summary = summarize_task_events(stream_events, expected_task_ids)
        validation_summary = validate_generated_outputs(
            case_pairs=case_pairs,
            case_by_task_id=case_by_task_id,
            event_summary=event_summary,
        )

        summary: dict[str, Any] = {
            "generatedAt": datetime.now().astimezone().isoformat(timespec="seconds"),
            "casesDir": str(cases_dir.resolve()),
            "outputDir": str(run_dir),
            "baseUrl": resolved_base_url,
            "host": host,
            "port": resolved_port,
            "renderStrategy": render_strategy,
            "goldenCorpusFileCount": sum(1 for _ in cases_dir.resolve().rglob("*") if _.is_file()),
            "msgFileCount": msg_file_count,
            "acceptedCount": len(accepted_items),
            "requestedCount": int(convert_payload.get("requestedCount", len(accepted_items))),
            "convertedCount": len(convert_payload.get("convertedFiles", [])),
            "conversionErrors": list(convert_payload.get("errors", [])),
            "skippedFiles": list(convert_payload.get("skippedFiles", [])),
            "timingLines": list(convert_payload.get("timingLines", [])),
            "eventSummary": event_summary,
            "validationSummary": validation_summary,
        }
        json_path = run_dir / "summary.json"
        markdown_path = run_dir / "summary.md"
        json_path.write_text(json.dumps(summary, indent=2), encoding="utf-8")
        _write_markdown(summary, markdown_path)
        return summary, json_path, markdown_path
    finally:
        if server_process is not None:
            server_process.terminate()
            try:
                server_process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                server_process.kill()
                server_process.wait(timeout=10)
        if server_log is not None:
            server_log.close()
        if server_err_log is not None:
            server_err_log.close()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Run the canonical golden corpus through the browser API and validate the generated PDFs."
    )
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
    parser.add_argument("--host", default=DEFAULT_HOST)
    parser.add_argument("--port", type=int, default=None, help="Optional server port when starting a local server.")
    parser.add_argument("--base-url", help="Use an already-running browser server instead of starting one.")
    parser.add_argument(
        "--timeout-seconds",
        type=float,
        default=300.0,
        help="Timeout for server startup and corpus completion (default: 300).",
    )
    parser.add_argument(
        "--render-strategy",
        choices=[RENDER_STRATEGY_FIDELITY, RENDER_STRATEGY_FAST],
        default=RENDER_STRATEGY_FIDELITY,
        help="Render strategy for the browser server when this command starts it (default: fidelity).",
    )
    args = parser.parse_args(argv)

    summary, json_path, markdown_path = run_browser_validation(
        cases_dir=args.cases_dir,
        output_root=args.output_root,
        host=args.host,
        port=args.port,
        base_url=args.base_url,
        timeout_seconds=args.timeout_seconds,
        render_strategy=args.render_strategy,
    )
    print(f"Wrote JSON report: {json_path}")
    print(f"Wrote Markdown report: {markdown_path}")
    print(
        f"Browser validation summary: processed={summary['msgFileCount']}, "
        f"converted={summary['convertedCount']}, "
        f"event_failures={summary['eventSummary']['failedCount']}, "
        f"validator_failed={summary['validationSummary']['failedCount']}, "
        f"validator_warnings={summary['validationSummary']['warningCaseCount']}"
    )
    if summary["eventSummary"]["failedCount"] or summary["eventSummary"]["missingTerminalTasks"]:
        return 1
    if summary["validationSummary"]["failedCount"]:
        return 2
    if summary["validationSummary"]["warningCaseCount"]:
        return 3
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
