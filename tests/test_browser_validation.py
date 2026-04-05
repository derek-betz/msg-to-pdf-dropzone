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
