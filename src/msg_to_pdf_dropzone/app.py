from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from tkinterdnd2 import DND_FILES, TkinterDnD

from .converter import MAX_FILES_PER_BATCH, ConversionError, convert_msg_files


class MsgToPdfApp:
    def __init__(self, root: TkinterDnD.Tk) -> None:
        self.root = root
        self.selected_files: list[Path] = []

        self.root.title("MSG to PDF Dropzone")
        self.root.geometry("780x520")
        self.root.minsize(700, 440)

        self.status_var = tk.StringVar(value="Drop up to 10 .msg files.")

        self._build_ui()

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=14)
        container.pack(fill=tk.BOTH, expand=True)

        title = ttk.Label(
            container,
            text="Drop Outlook .msg files and convert to PDF",
            font=("Segoe UI", 13, "bold"),
        )
        title.pack(anchor=tk.W, pady=(0, 8))

        subtitle = ttk.Label(
            container,
            text="Each .msg becomes one PDF. Output filename is prefixed with latest thread date.",
        )
        subtitle.pack(anchor=tk.W, pady=(0, 12))

        self.drop_zone = tk.Label(
            container,
            text="Drag and drop .msg files here",
            relief=tk.GROOVE,
            borderwidth=2,
            bg="#f3f5f8",
            padx=10,
            pady=36,
            font=("Segoe UI", 11),
        )
        self.drop_zone.pack(fill=tk.X, pady=(0, 12))
        self.drop_zone.drop_target_register(DND_FILES)
        self.drop_zone.dnd_bind("<<Drop>>", self._on_drop)

        list_frame = ttk.Frame(container)
        list_frame.pack(fill=tk.BOTH, expand=True)

        self.file_listbox = tk.Listbox(
            list_frame,
            selectmode=tk.EXTENDED,
            height=14,
            font=("Consolas", 10),
        )
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.configure(yscrollcommand=scrollbar.set)

        button_row = ttk.Frame(container)
        button_row.pack(fill=tk.X, pady=(12, 4))

        ttk.Button(button_row, text="Add Files", command=self._choose_files).pack(side=tk.LEFT)
        ttk.Button(button_row, text="Remove Selected", command=self._remove_selected).pack(
            side=tk.LEFT, padx=(8, 0)
        )
        ttk.Button(button_row, text="Clear", command=self._clear_files).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(button_row, text="Convert to PDF", command=self._convert).pack(side=tk.RIGHT)

        status_label = ttk.Label(container, textvariable=self.status_var)
        status_label.pack(fill=tk.X, pady=(8, 0))

    def _choose_files(self) -> None:
        selected = filedialog.askopenfilenames(
            title="Select Outlook .msg files",
            filetypes=[("Outlook message", "*.msg")],
        )
        self._add_files([Path(path) for path in selected])

    def _remove_selected(self) -> None:
        selected_indexes = sorted(self.file_listbox.curselection(), reverse=True)
        if not selected_indexes:
            return
        for index in selected_indexes:
            del self.selected_files[index]
        self._refresh_file_list()

    def _clear_files(self) -> None:
        self.selected_files.clear()
        self._refresh_file_list()

    def _on_drop(self, event: object) -> None:
        raw_data = getattr(event, "data", "")
        path_strings = self.root.tk.splitlist(raw_data)
        candidates = [Path(value) for value in path_strings]
        self._add_files(candidates)

    def _add_files(self, candidates: list[Path]) -> None:
        msg_candidates = []
        for path in candidates:
            if path.suffix.lower() == ".msg":
                msg_candidates.append(path)

        existing = {path.resolve() for path in self.selected_files}
        unique_candidates = []
        for path in msg_candidates:
            resolved = path.resolve()
            if resolved not in existing:
                unique_candidates.append(resolved)
                existing.add(resolved)

        if not unique_candidates:
            self.status_var.set("No new .msg files added.")
            return

        available_slots = MAX_FILES_PER_BATCH - len(self.selected_files)
        accepted = unique_candidates[: max(available_slots, 0)]
        ignored_count = len(unique_candidates) - len(accepted)

        self.selected_files.extend(accepted)
        self._refresh_file_list()

        if ignored_count > 0:
            messagebox.showwarning(
                "File Limit Reached",
                f"Only {MAX_FILES_PER_BATCH} files can be converted at once. "
                f"Ignored {ignored_count} file(s).",
            )

    def _refresh_file_list(self) -> None:
        self.file_listbox.delete(0, tk.END)
        for path in self.selected_files:
            self.file_listbox.insert(tk.END, str(path))
        self.status_var.set(f"Selected {len(self.selected_files)} of {MAX_FILES_PER_BATCH} allowed files.")

    def _convert(self) -> None:
        if not self.selected_files:
            messagebox.showwarning("No Files", "Add at least one .msg file first.")
            return

        output_dir = filedialog.askdirectory(title="Choose where to save converted PDFs")
        if not output_dir:
            self.status_var.set("Conversion canceled (no output folder selected).")
            return

        self.status_var.set("Converting...")
        self.root.update_idletasks()

        try:
            result = convert_msg_files(self.selected_files, Path(output_dir))
        except ConversionError as exc:
            messagebox.showerror("Conversion Error", str(exc))
            self.status_var.set(str(exc))
            return

        summary = [f"Converted {len(result.converted_files)} of {result.requested_count} file(s)."]
        if result.errors:
            summary.append("")
            summary.append("Issues:")
            summary.extend(result.errors)

        if result.converted_files:
            messagebox.showinfo("Conversion Complete", "\n".join(summary))
        else:
            messagebox.showwarning("No Files Converted", "\n".join(summary))
        self.status_var.set(summary[0])


def main() -> None:
    root = TkinterDnD.Tk()
    MsgToPdfApp(root)
    root.mainloop()
