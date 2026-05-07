from __future__ import annotations

from msg_to_pdf_dropzone.browser_validation import summarize_task_events


def test_summarize_task_events_reports_outputs_and_sequence_issues() -> None:
    task_ids = ["task-1", "task-2"]
    events = [
        {"taskId": "task-1", "stage": "drop_received"},
        {"taskId": "task-1", "stage": "files_accepted"},
        {"taskId": "task-1", "stage": "output_folder_selected"},
        {"taskId": "task-1", "stage": "parse_started"},
        {
            "taskId": "task-1",
            "stage": "complete",
            "pipeline": "outlook_edge",
            "meta": {
                "outputName": "2026-04-04_Sample.pdf",
                "outputPath": r"C:\temp\2026-04-04_Sample.pdf",
            },
        },
        {"taskId": "task-2", "stage": "drop_received"},
        {"taskId": "task-2", "stage": "pdf_written"},
        {"taskId": "task-2", "stage": "parse_started"},
    ]

    summary = summarize_task_events(events, task_ids)

    assert summary["taskCount"] == 2
    assert summary["completedCount"] == 1
    assert summary["failedCount"] == 0
    assert summary["pipelineCounts"] == {"outlook_edge": 1}
    assert summary["allOutputNamesMatchPattern"] is True
    assert summary["outputPathsByTaskId"]["task-1"].endswith("2026-04-04_Sample.pdf")
    assert summary["missingTerminalTasks"] == ["task-2"]
    assert summary["orderIssues"] == [
        {
            "taskId": "task-2",
            "stage": "parse_started",
            "message": "Stage 'parse_started' arrived after a later stage.",
        }
    ]


def test_summarize_task_events_validates_selected_filename_style() -> None:
    task_ids = ["task-1", "task-2"]
    events = [
        {
            "taskId": "task-1",
            "stage": "complete",
            "pipeline": "edge_html",
            "meta": {"outputName": "Jane Smith_Project Update.pdf"},
        },
        {
            "taskId": "task-2",
            "stage": "complete",
            "pipeline": "edge_html",
            "meta": {"outputName": "2026-04-04_Project Update.pdf"},
        },
    ]

    sender_style_summary = summarize_task_events(events, task_ids, filename_style="sender_subject")
    default_style_summary = summarize_task_events(events, task_ids)

    assert sender_style_summary["filenameStyle"] == "sender_subject"
    assert sender_style_summary["allOutputNamesMatchPattern"] is False
    assert sender_style_summary["styleMismatchedOutputs"] == ["2026-04-04_Project Update.pdf"]
    assert default_style_summary["filenameStyle"] == "date_subject"
    assert default_style_summary["styleMismatchedOutputs"] == ["Jane Smith_Project Update.pdf"]


def test_summarize_task_events_allows_unique_numeric_suffixes() -> None:
    task_ids = ["task-1"]
    events = [
        {
            "taskId": "task-1",
            "stage": "complete",
            "pipeline": "edge_html",
            "meta": {"outputName": "2026-04-04_Jane Smith_Project Update (2).pdf"},
        }
    ]

    summary = summarize_task_events(events, task_ids, filename_style="date_sender_subject")

    assert summary["allOutputNamesMatchPattern"] is True
    assert summary["numericSuffixOutputs"] == ["2026-04-04_Jane Smith_Project Update (2).pdf"]
