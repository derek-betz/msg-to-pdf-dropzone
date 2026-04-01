from __future__ import annotations

import threading
import time
import tkinter as tk
from pathlib import Path
from time import perf_counter
from tkinter import filedialog, messagebox, ttk
from typing import Callable
from uuid import uuid4

from tkinterdnd2 import COPY, DND_ALL, TkinterDnD

from .converter import MAX_FILES_PER_BATCH, ConversionError, ConversionResult, convert_msg_files
from .drop_helpers import is_supported_msg_candidate, parse_drop_paths, wait_for_materialized_file
from .outlook_selection import extract_selected_outlook_messages, is_likely_outlook_drop
from .task_events import (
    TaskEventSink,
    TaskMetaValue,
    TaskStage,
    default_task_id_for_path,
    emit_task_event,
)
from .theater_host import TheaterController

DROP_UI_READY_TARGET_SECONDS = 2.0
DEFAULT_DROP_MATERIALIZATION_TIMEOUT_SECONDS = 0.10
HEARTBEAT_INTERVAL_MS = 50


class MsgToPdfApp:
    def __init__(
        self,
        root: TkinterDnD.Tk,
        event_sink: TaskEventSink | None = None,
        theater_controller: TheaterController | None = None,
    ) -> None:
        self.root = root
        self.selected_files: list[Path] = []
        self.temp_outlook_files: set[Path] = set()
        self._event_sink = event_sink
        self._theater_controller = theater_controller

        self.root.title("MSG to PDF Dropzone")
        self.root.geometry("920x680")
        self.root.minsize(820, 620)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.status_var = tk.StringVar(value="Ready for drop. Add up to 25 Outlook .msg files.")
        self.status_detail_var = tk.StringVar(
            value="Drag from File Explorer or Classic Outlook, then choose an output folder."
        )
        self.selection_count_var = tk.StringVar(value="0 selected")
        self.selection_hint_var = tk.StringVar(value="No files queued yet.")
        self._busy = False
        self._drop_dispatch_seconds = 0.0
        self._drop_slow_warning_count = 0

        self._heartbeat_last_tick = perf_counter()
        self._heartbeat_job_id: str | None = None
        self._active_operation_name = ""
        self._active_operation_max_stall_seconds = 0.0
        self._last_operation_name = ""
        self._last_operation_max_stall_seconds = 0.0

        self._build_ui()
        self._fit_window_to_content()
        self._update_theater_button()
        self._schedule_heartbeat()
        if self._should_open_theater_on_launch():
            self.root.after(0, self._open_theater_on_launch)

    def _build_ui(self) -> None:
        self._apply_styles()

        container = ttk.Frame(self.root, padding=18, style="App.TFrame")
        container.pack(fill=tk.BOTH, expand=True)

        hero = ttk.Frame(container, padding=18, style="Card.TFrame")
        hero.pack(fill=tk.X, pady=(0, 14))

        hero_header = ttk.Frame(hero, style="Card.TFrame")
        hero_header.pack(fill=tk.X)

        title_block = ttk.Frame(hero_header, style="Card.TFrame")
        title_block.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Label(
            title_block,
            text="Convert Outlook Emails to Clean PDFs",
            style="HeroTitle.TLabel",
        ).pack(anchor=tk.W)
        ttk.Label(
            title_block,
            text=(
                "Drag in .msg files or a Classic Outlook selection. "
                "Each email becomes one PDF with thread-aware naming."
            ),
            style="HeroSubtitle.TLabel",
        ).pack(anchor=tk.W, pady=(6, 0))

        ttk.Label(
            hero_header,
            textvariable=self.selection_count_var,
            style="Badge.TLabel",
            anchor=tk.CENTER,
            width=16,
        ).pack(side=tk.RIGHT, padx=(14, 0))

        self.drop_zone = tk.Label(
            hero,
            text="Drop Outlook .msg files here\n\nClassic Outlook drag is supported.",
            relief=tk.GROOVE,
            borderwidth=2,
            bg="#f5f7fb",
            fg="#16324f",
            padx=16,
            pady=28,
            font=("Segoe UI Semibold", 12),
            justify=tk.CENTER,
        )
        self.drop_zone_default_bg = self.drop_zone.cget("bg")
        self.drop_zone.pack(fill=tk.X, pady=(16, 10))
        self.drop_zone.drop_target_register(DND_ALL)
        self.drop_zone.dnd_bind("<<DropEnter>>", self._on_drop_enter)
        self.drop_zone.dnd_bind("<<DropPosition>>", self._on_drop_position)
        self.drop_zone.dnd_bind("<<DropLeave>>", self._on_drop_leave)
        self.drop_zone.dnd_bind("<<Drop>>", self._on_drop)

        meta_row = ttk.Frame(hero, style="Card.TFrame")
        meta_row.pack(fill=tk.X)
        ttk.Label(meta_row, text="Batch limit: 25 files", style="Meta.TLabel").pack(side=tk.LEFT)
        ttk.Label(meta_row, text="Choose the output folder when converting", style="Meta.TLabel").pack(
            side=tk.LEFT, padx=(18, 0)
        )
        ttk.Label(meta_row, text="Conversion keeps the validated PDF pipeline", style="Meta.TLabel").pack(
            side=tk.RIGHT
        )

        body = ttk.Frame(container, style="App.TFrame")
        body.pack(fill=tk.BOTH, expand=True)

        queue_card = ttk.Frame(body, padding=14, style="Card.TFrame")
        queue_card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        queue_header = ttk.Frame(queue_card, style="Card.TFrame")
        queue_header.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(queue_header, text="Queued Messages", style="SectionTitle.TLabel").pack(side=tk.LEFT)
        ttk.Label(queue_header, textvariable=self.selection_hint_var, style="Muted.TLabel").pack(side=tk.RIGHT)

        list_frame = ttk.Frame(queue_card, style="Card.TFrame")
        list_frame.pack(fill=tk.BOTH, expand=True)

        self.file_listbox = tk.Listbox(
            list_frame,
            selectmode=tk.EXTENDED,
            height=12,
            font=("Consolas", 10),
            bg="#fbfcfe",
            bd=0,
            highlightthickness=0,
            activestyle="none",
            selectbackground="#dbe8ff",
            selectforeground="#10233d",
        )
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.configure(yscrollcommand=scrollbar.set)

        action_card = ttk.Frame(body, padding=14, style="CardAccent.TFrame")
        action_card.pack(side=tk.RIGHT, fill=tk.Y, padx=(14, 0))

        ttk.Label(action_card, text="Next Step", style="SectionTitleBlue.TLabel").pack(anchor=tk.W)
        ttk.Label(
            action_card,
            text="Review the queue, then convert the batch and choose where the PDFs should be saved.",
            style="PanelBody.TLabel",
            wraplength=200,
            justify=tk.LEFT,
        ).pack(anchor=tk.W, pady=(6, 14))

        self.convert_button = ttk.Button(
            action_card,
            text="Convert to PDF",
            command=self._convert,
            style="Primary.TButton",
        )
        self.convert_button.pack(fill=tk.X)
        self.add_button = ttk.Button(action_card, text="Add Files", command=self._choose_files)
        self.add_button.pack(fill=tk.X, pady=(12, 0))
        self.remove_button = ttk.Button(action_card, text="Remove Selected", command=self._remove_selected)
        self.remove_button.pack(fill=tk.X, pady=(8, 0))
        self.clear_button = ttk.Button(action_card, text="Clear Queue", command=self._clear_files)
        self.clear_button.pack(fill=tk.X, pady=(8, 0))
        self.theater_button = ttk.Button(action_card, text="Open Theater", command=self._toggle_theater)
        self.theater_button.pack(fill=tk.X, pady=(8, 0))

        ttk.Separator(action_card, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=14)
        ttk.Label(action_card, text="What You Get", style="SectionTitleBlue.TLabel").pack(anchor=tk.W)
        ttk.Label(
            action_card,
            text=(
                "Clean PDF headers, no print timestamp stamp, and semantic validation against the golden corpus."
            ),
            style="PanelBody.TLabel",
            wraplength=200,
            justify=tk.LEFT,
        ).pack(anchor=tk.W, pady=(6, 0))

        status_card = ttk.Frame(container, padding=12, style="Status.TFrame")
        status_card.pack(fill=tk.X, pady=(14, 0))
        ttk.Label(status_card, text="Status", style="SectionTitleStatus.TLabel").pack(anchor=tk.W)
        ttk.Label(status_card, textvariable=self.status_var, style="StatusMain.TLabel").pack(
            anchor=tk.W, pady=(6, 2)
        )
        ttk.Label(
            status_card,
            textvariable=self.status_detail_var,
            style="StatusDetail.TLabel",
            wraplength=720,
            justify=tk.LEFT,
        ).pack(anchor=tk.W)
        self.progress = ttk.Progressbar(status_card, mode="indeterminate", length=180)
        self.progress.pack(fill=tk.X, pady=(10, 0))

        self._refresh_file_list()

    def _apply_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        self.root.configure(bg="#e9eef5")
        style.configure("App.TFrame", background="#e9eef5")
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("CardAccent.TFrame", background="#edf4ff")
        style.configure("Status.TFrame", background="#f4f7fb")
        style.configure("HeroTitle.TLabel", background="#ffffff", foreground="#10233d", font=("Segoe UI", 16, "bold"))
        style.configure("HeroSubtitle.TLabel", background="#ffffff", foreground="#4c6078", font=("Segoe UI", 10))
        style.configure("SectionTitle.TLabel", background="#ffffff", foreground="#17304c", font=("Segoe UI", 10, "bold"))
        style.configure(
            "SectionTitleBlue.TLabel", background="#edf4ff", foreground="#17304c", font=("Segoe UI", 10, "bold")
        )
        style.configure(
            "SectionTitleStatus.TLabel", background="#f4f7fb", foreground="#17304c", font=("Segoe UI", 10, "bold")
        )
        style.configure("PanelBody.TLabel", background="#edf4ff", foreground="#304761", font=("Segoe UI", 9))
        style.configure("StatusMain.TLabel", background="#f4f7fb", foreground="#13263d", font=("Segoe UI", 10, "bold"))
        style.configure("StatusDetail.TLabel", background="#f4f7fb", foreground="#5f7084", font=("Segoe UI", 9))
        style.configure("Muted.TLabel", background="#ffffff", foreground="#617287", font=("Segoe UI", 9))
        style.configure("Meta.TLabel", background="#ffffff", foreground="#617287", font=("Segoe UI", 9))
        style.configure(
            "Badge.TLabel",
            background="#16324f",
            foreground="#ffffff",
            font=("Segoe UI", 10, "bold"),
            padding=(10, 6),
        )
        style.configure("TButton", font=("Segoe UI", 9))
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"))

    def _fit_window_to_content(self) -> None:
        try:
            self.root.update_idletasks()
            requested_width = self.root.winfo_reqwidth()
            requested_height = self.root.winfo_reqheight()
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
        except Exception:
            return

        horizontal_margin = 80
        vertical_margin = 100
        max_width = max(700, screen_width - horizontal_margin)
        max_height = max(520, screen_height - vertical_margin)

        target_width = min(max(920, requested_width + 24), max_width)
        target_height = min(max(680, requested_height + 24), max_height)

        min_width = min(max(820, requested_width), max_width)
        min_height = min(max(620, requested_height), max_height)

        self.root.geometry(f"{target_width}x{target_height}")
        self.root.minsize(min_width, min_height)

    def _set_status(self, message: str, detail: str | None = None) -> None:
        status_var = getattr(self, "status_var", None)
        if status_var is not None:
            status_var.set(message)
        if detail is not None:
            status_detail_var = getattr(self, "status_detail_var", None)
            if status_detail_var is not None:
                status_detail_var.set(detail)

    def _update_selection_summary(self) -> None:
        count = len(getattr(self, "selected_files", []))
        selection_count_var = getattr(self, "selection_count_var", None)
        if selection_count_var is not None:
            selection_count_var.set(f"{count} selected")

        selection_hint_var = getattr(self, "selection_hint_var", None)
        if selection_hint_var is not None:
            if count == 0:
                selection_hint_var.set("No files queued yet.")
            elif count == 1:
                selection_hint_var.set("Ready to convert one email.")
            else:
                selection_hint_var.set(f"Ready to convert {count} emails.")

    def _build_completion_dialog_message(self, result: ConversionResult, output_path: Path) -> str:
        converted_count = len(result.converted_files)
        skipped_count = len(result.skipped_files)
        lines = [
            f"Converted {converted_count} of {result.requested_count} file(s).",
            f"Saved to: {output_path}",
            f"Total time: {result.total_seconds:.2f}s",
        ]

        if skipped_count:
            lines.append(f"Skipped: {skipped_count}")
        if result.errors:
            preview = result.errors[:3]
            lines.append("")
            lines.append("Issues:")
            lines.extend(f"- {item}" for item in preview)
            remaining = len(result.errors) - len(preview)
            if remaining > 0:
                lines.append(f"- {remaining} more issue(s) not shown")

        return "\n".join(lines)

    def _build_completion_status_detail(self, result: ConversionResult, output_path: Path) -> str:
        detail_parts = [
            f"Output folder: {output_path}",
            f"Drop dispatch {self._drop_dispatch_seconds:.2f}s",
            f"UI stall max {int(self._last_operation_max_stall_seconds * 1000)}ms",
        ]
        if result.timing_lines:
            detail_parts.append("Diagnostics available in timing records for this run.")
        if result.errors:
            detail_parts.append(f"{len(result.errors)} issue(s) were reported during conversion.")
        return " | ".join(detail_parts)

    def _set_busy(self, busy: bool, status_message: str | None = None) -> None:
        self._busy = busy
        state = tk.DISABLED if busy else tk.NORMAL
        self.add_button.configure(state=state)
        self.remove_button.configure(state=state)
        self.clear_button.configure(state=state)
        self.convert_button.configure(state=state)
        self.file_listbox.configure(state=state)

        cursor = "watch" if busy else ""
        self.root.configure(cursor=cursor)
        self.drop_zone.configure(cursor=cursor)
        progress = getattr(self, "progress", None)
        if progress is not None:
            if busy:
                progress.start(10)
            else:
                progress.stop()

        if status_message:
            self._set_status(status_message)

    def _schedule_heartbeat(self) -> None:
        self._heartbeat_last_tick = perf_counter()
        self._heartbeat_job_id = self.root.after(HEARTBEAT_INTERVAL_MS, self._on_heartbeat_tick)

    def _on_heartbeat_tick(self) -> None:
        now = perf_counter()
        interval_seconds = HEARTBEAT_INTERVAL_MS / 1000.0
        elapsed = now - self._heartbeat_last_tick
        self._heartbeat_last_tick = now

        if self._active_operation_name:
            stall_seconds = max(0.0, elapsed - interval_seconds)
            if stall_seconds > self._active_operation_max_stall_seconds:
                self._active_operation_max_stall_seconds = stall_seconds

        self._update_theater_button()
        self._heartbeat_job_id = self.root.after(HEARTBEAT_INTERVAL_MS, self._on_heartbeat_tick)

    def _begin_operation_diagnostics(self, operation_name: str) -> None:
        self._active_operation_name = operation_name
        self._active_operation_max_stall_seconds = 0.0

    def _finish_operation_diagnostics(self) -> float:
        self._last_operation_name = self._active_operation_name
        self._last_operation_max_stall_seconds = self._active_operation_max_stall_seconds
        self._active_operation_name = ""
        self._active_operation_max_stall_seconds = 0.0
        return self._last_operation_max_stall_seconds

    def _run_in_background(
        self,
        *,
        operation_name: str,
        status_message: str,
        work: Callable[[], object],
        on_success: Callable[[object], None],
        on_error: Callable[[Exception], None] | None = None,
    ) -> bool:
        if self._busy:
            self._set_status("Please wait for the current operation to finish.")
            return False

        self._begin_operation_diagnostics(operation_name)
        self._set_busy(True, status_message)

        def _worker() -> None:
            result: object | None = None
            error: Exception | None = None
            try:
                result = work()
            except Exception as exc:  # pragma: no cover - defensive UI worker safety
                error = exc

            def _complete() -> None:
                self._set_busy(False)
                self._finish_operation_diagnostics()
                if error is None:
                    on_success(result)
                    return
                if on_error is not None:
                    on_error(error)
                    return
                messagebox.showerror("Operation Error", str(error))
                self._set_status("Operation failed.", str(error))

            try:
                self.root.after(0, _complete)
            except Exception:
                pass

        threading.Thread(target=_worker, daemon=True).start()
        return True

    def _choose_files(self) -> None:
        if self._busy:
            self._set_status("Please wait for the current operation to finish.")
            return
        selected = filedialog.askopenfilenames(
            title="Select Outlook .msg files",
            filetypes=[("Outlook message", "*.msg")],
        )
        self._add_files([Path(path) for path in selected], materialization_timeout_seconds=0.0)

    def _remove_selected(self) -> None:
        if self._busy:
            self._set_status("Please wait for the current operation to finish.")
            return
        selected_indexes = sorted(self.file_listbox.curselection(), reverse=True)
        if not selected_indexes:
            return
        for index in selected_indexes:
            removed_path = self.selected_files[index]
            del self.selected_files[index]
            self._delete_temp_file_if_managed(removed_path)
        self._refresh_file_list()

    def _clear_files(self) -> None:
        if self._busy:
            self._set_status("Please wait for the current operation to finish.")
            return
        self._delete_managed_temp_files(self.selected_files)
        self.selected_files.clear()
        self._refresh_file_list()

    def _on_drop_enter(self, _: object) -> str:
        if not self._busy:
            self.drop_zone.configure(bg="#e3eaef")
        return COPY

    def _on_drop_position(self, _: object) -> str:
        return COPY

    def _on_drop_leave(self, _: object) -> None:
        self.drop_zone.configure(bg=self.drop_zone_default_bg)

    def _on_drop(self, event: object) -> str:
        drop_started_at = perf_counter()
        if self._busy:
            self._set_status("Please wait for the current operation to finish.")
            return COPY

        try:
            self.drop_zone.configure(bg=self.drop_zone_default_bg)
            raw_data = getattr(event, "data", "")
            sourcetypes = getattr(event, "sourcetypes", ())
            likely_outlook_drop = is_likely_outlook_drop(raw_data, sourcetypes)

            candidates = parse_drop_paths(raw_data, self.root.tk.splitlist)
            # Outlook drops often include token payloads that are not real files.
            # Avoid blocking the UI waiting on materialization for those tokens.
            materialization_timeout_seconds = (
                0.0 if likely_outlook_drop else DEFAULT_DROP_MATERIALIZATION_TIMEOUT_SECONDS
            )
            added_count = self._add_files(
                candidates,
                is_temp_outlook_file=False,
                show_empty_status=False,
                materialization_timeout_seconds=materialization_timeout_seconds,
                emit_drop_received=True,
            )
            if added_count > 0:
                return COPY

            if likely_outlook_drop or not candidates:
                remaining_slots = MAX_FILES_PER_BATCH - len(self.selected_files)
                if remaining_slots <= 0:
                    messagebox.showwarning(
                        "File Limit Reached",
                        f"Only {MAX_FILES_PER_BATCH} files can be converted at once.",
                    )
                    return COPY

                def _work() -> object:
                    started_at = time.perf_counter()
                    files = extract_selected_outlook_messages(remaining_slots)
                    return files, time.perf_counter() - started_at

                def _success(result: object) -> None:
                    outlook_temp_files: list[Path] = []
                    extract_seconds = 0.0
                    if isinstance(result, tuple) and len(result) == 2:
                        files, elapsed = result
                        if isinstance(files, list):
                            outlook_temp_files = [path for path in files if isinstance(path, Path)]
                        if isinstance(elapsed, (float, int)):
                            extract_seconds = float(elapsed)
                    elif isinstance(result, list):
                        outlook_temp_files = [path for path in result if isinstance(path, Path)]

                    existing_before = {path.resolve() for path in self.selected_files}
                    outlook_added = self._add_files(
                        outlook_temp_files,
                        is_temp_outlook_file=True,
                        show_empty_status=False,
                        emit_files_accepted=False,
                    )
                    accepted_paths = [
                        path
                        for path in self.selected_files
                        if path.resolve() not in existing_before
                    ]
                    self._emit_stage_for_paths(accepted_paths, "outlook_extract_started")
                    self._emit_stage_for_paths(accepted_paths, "files_accepted")
                    if outlook_added > 0:
                        stall_ms = int(self._last_operation_max_stall_seconds * 1000)
                        self._set_status(
                            f"Added {outlook_added} message(s) from Classic Outlook.",
                            f"Extraction took {extract_seconds:.2f}s. Maximum UI stall observed: {stall_ms}ms.",
                        )
                    else:
                        stall_ms = int(self._last_operation_max_stall_seconds * 1000)
                        self._set_status(
                            "No new .msg files were added from the Outlook selection.",
                            f"Outlook check took {extract_seconds:.2f}s. Maximum UI stall observed: {stall_ms}ms.",
                        )

                def _error(exc: Exception) -> None:
                    self._set_status("Failed to read the Outlook selection.", str(exc))

                self._run_in_background(
                    operation_name="outlook_selection",
                    status_message="Reading selected message(s) from Classic Outlook...",
                    work=_work,
                    on_success=_success,
                    on_error=_error,
                )
                return COPY

            self._set_status("No new .msg files added.")
            return COPY
        finally:
            self._drop_dispatch_seconds = perf_counter() - drop_started_at
            if self._drop_dispatch_seconds > DROP_UI_READY_TARGET_SECONDS:
                self._drop_slow_warning_count += 1

    def _add_files(
        self,
        candidates: list[Path],
        *,
        is_temp_outlook_file: bool = False,
        show_empty_status: bool = True,
        materialization_timeout_seconds: float = 2.0,
        emit_drop_received: bool = False,
        emit_files_accepted: bool = True,
    ) -> int:
        msg_candidates = []
        available_slots = MAX_FILES_PER_BATCH - len(self.selected_files)
        for path in candidates:
            if available_slots > 0 and len(msg_candidates) >= available_slots:
                break
            candidate = wait_for_materialized_file(path, timeout_seconds=materialization_timeout_seconds)
            if not candidate.exists():
                continue
            if candidate.exists() and candidate.is_dir():
                continue
            if is_supported_msg_candidate(candidate):
                msg_candidates.append(candidate)

        existing = {path.resolve() for path in self.selected_files}
        unique_candidates = []
        for path in msg_candidates:
            resolved = path.resolve()
            if resolved not in existing:
                unique_candidates.append(resolved)
                existing.add(resolved)

        if not unique_candidates:
            if show_empty_status:
                self._set_status("No new .msg files added.")
            return 0

        accepted = unique_candidates[: max(available_slots, 0)]
        ignored_count = len(unique_candidates) - len(accepted)

        self.selected_files.extend(accepted)
        if is_temp_outlook_file:
            self.temp_outlook_files.update(accepted)
        self._refresh_file_list()
        if emit_drop_received:
            self._emit_stage_for_paths(accepted, "drop_received")
        if emit_files_accepted:
            self._emit_stage_for_paths(accepted, "files_accepted")

        if ignored_count > 0:
            messagebox.showwarning(
                "File Limit Reached",
                f"Only {MAX_FILES_PER_BATCH} files can be converted at once. "
                f"Ignored {ignored_count} file(s).",
            )

        return len(accepted)

    def _refresh_file_list(self) -> None:
        self.file_listbox.delete(0, tk.END)
        for path in self.selected_files:
            self.file_listbox.insert(tk.END, str(path))
        self._update_selection_summary()
        self._set_status(
            f"Queue has {len(self.selected_files)} of {MAX_FILES_PER_BATCH} allowed files.",
            "Choose more files, remove any you do not want, then click Convert to PDF.",
        )

    def _convert(self) -> None:
        if self._busy:
            self._set_status("Please wait for the current operation to finish.")
            return

        if not self.selected_files:
            messagebox.showwarning("No Files", "Add at least one .msg file first.")
            return

        output_dir = filedialog.askdirectory(title="Choose where to save converted PDFs")
        if not output_dir:
            self._set_status("Conversion canceled.", "No output folder was selected.")
            return

        selected_files = list(self.selected_files)
        output_path = Path(output_dir)
        batch_meta_by_source_path = self._build_batch_meta_by_path(selected_files)
        event_sink = self._current_event_sink()
        for path in selected_files:
            batch_meta = batch_meta_by_source_path[path.resolve()]
            emit_task_event(
                event_sink,
                task_id=self._task_id_for_path(path),
                stage="output_folder_selected",
                file_name=path.name,
                meta={
                    **batch_meta,
                    "outputDir": str(output_path),
                    "outputDirLabel": output_path.name or str(output_path),
                },
            )
        task_ids_by_source_path = {
            path.resolve(): self._task_id_for_path(path)
            for path in selected_files
        }

        def _work() -> object:
            return convert_msg_files(
                selected_files,
                output_path,
                event_sink=self._current_event_sink(),
                task_ids_by_source_path=task_ids_by_source_path,
                batch_meta_by_source_path=batch_meta_by_source_path,
            )

        def _success(result: object) -> None:
            if not isinstance(result, ConversionResult):
                messagebox.showerror("Conversion Error", "Unexpected conversion result type.")
                self._set_status("Conversion failed.", "Unexpected result type returned by converter.")
                return

            conversion_result = result
            dialog_message = self._build_completion_dialog_message(conversion_result, output_path)
            status_message = (
                f"Converted {len(conversion_result.converted_files)} of "
                f"{conversion_result.requested_count} file(s)."
            )
            status_detail = self._build_completion_status_detail(conversion_result, output_path)

            if conversion_result.converted_files:
                messagebox.showinfo("Conversion Complete", dialog_message)
            else:
                messagebox.showwarning("No Files Converted", dialog_message)
            self._set_status(status_message, status_detail)

        def _error(exc: Exception) -> None:
            if isinstance(exc, ConversionError):
                messagebox.showerror("Conversion Error", str(exc))
                self._set_status("Conversion failed.", str(exc))
                return
            messagebox.showerror("Conversion Error", f"Unexpected failure: {exc}")
            self._set_status("Conversion failed.", str(exc))

        self._run_in_background(
            operation_name="convert",
            status_message="Converting...",
            work=_work,
            on_success=_success,
            on_error=_error,
        )

    def _on_close(self) -> None:
        if self._busy:
            messagebox.showinfo(
                "Please Wait",
                "An operation is still running. Wait for it to finish before closing.",
            )
            return
        if self._heartbeat_job_id is not None:
            try:
                self.root.after_cancel(self._heartbeat_job_id)
            except Exception:
                pass
            self._heartbeat_job_id = None
        self._delete_managed_temp_files(self.temp_outlook_files)
        theater_controller = getattr(self, "_theater_controller", None)
        if theater_controller is not None:
            theater_controller.shutdown()
        self.root.destroy()

    def _delete_managed_temp_files(self, paths: list[Path] | set[Path]) -> None:
        for path in list(paths):
            self._delete_temp_file_if_managed(path)

    def _delete_temp_file_if_managed(self, path: Path) -> None:
        if path not in self.temp_outlook_files:
            return
        try:
            if path.exists():
                path.unlink()
        except Exception:
            pass
        self.temp_outlook_files.discard(path)

    def _task_id_for_path(self, path: Path) -> str:
        return default_task_id_for_path(path)

    def _emit_stage_for_paths(
        self,
        paths: list[Path],
        stage: TaskStage,
        *,
        success: bool | None = None,
        error: str | None = None,
        meta: dict[str, TaskMetaValue] | None = None,
    ) -> None:
        event_sink = self._current_event_sink()
        if event_sink is None:
            return
        for path in paths:
            emit_task_event(
                event_sink,
                task_id=self._task_id_for_path(path),
                stage=stage,
                file_name=path.name,
                success=success,
                error=error,
                meta=meta,
            )

    def _build_batch_meta_by_path(self, paths: list[Path]) -> dict[Path, dict[str, TaskMetaValue]]:
        batch_id = f"msg-batch-{uuid4().hex[:12]}"
        batch_size = len(paths)
        return {
            path.resolve(): {
                "batchId": batch_id,
                "batchSize": batch_size,
                "batchIndex": index + 1,
            }
            for index, path in enumerate(paths)
        }

    def _current_event_sink(self) -> TaskEventSink | None:
        theater_controller = getattr(self, "_theater_controller", None)
        if theater_controller is not None and theater_controller.event_sink is not None:
            return theater_controller.event_sink
        return getattr(self, "_event_sink", None)

    def _should_open_theater_on_launch(self) -> bool:
        theater_controller = getattr(self, "_theater_controller", None)
        if theater_controller is None:
            return False
        return theater_controller.should_open_on_launch()

    def _open_theater_on_launch(self) -> None:
        theater_controller = getattr(self, "_theater_controller", None)
        if theater_controller is None:
            return
        theater_controller.open()
        self._update_theater_button()

    def _toggle_theater(self) -> None:
        theater_controller = getattr(self, "_theater_controller", None)
        if theater_controller is None:
            self._set_status("Theater integration is unavailable.")
            return

        if theater_controller.is_open:
            theater_controller.close()
            theater_controller.set_persisted_open(False)
            self._set_status("Theater closed.")
            self._update_theater_button()
            return

        if not theater_controller.open():
            self._set_status("Theater assets are unavailable.", "Build the theater runtime first.")
            self._update_theater_button()
            return

        theater_controller.set_persisted_open(True)
        self._set_status("Theater opened.")
        self._update_theater_button()

    def _update_theater_button(self) -> None:
        button = getattr(self, "theater_button", None)
        if button is None:
            return
        theater_controller = getattr(self, "_theater_controller", None)
        if theater_controller is None:
            button.configure(text="Theater Unavailable")
            return
        button.configure(text="Close Theater" if theater_controller.is_open else "Open Theater")


def main() -> None:
    root = TkinterDnD.Tk()
    theater_controller = TheaterController()
    MsgToPdfApp(root, theater_controller=theater_controller)
    root.mainloop()
