from __future__ import annotations

import os
import sys
import tkinter as tk
import zipfile
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from tkcalendar import DateEntry

from .core import find_matching_sheets, generate_test_result, today_text


class ReleaseNoteApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Release Note To Test Result")
        self.root.geometry("640x320")
        self.root.resizable(False, False)

        self.release_type_var = tk.StringVar(value="UAT")
        self.request_date_var = tk.StringVar(value=today_text())
        self.input_file_var = tk.StringVar()
        self.sheet_name_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Select a Release Note file to get started.")
        self.available_sheets: list[str] = []

        self.project_root = self._resolve_runtime_root()
        self.output_dir = self.project_root / "output" / "generated"
        self.taiga_config_path = self.project_root / "taiga.local.json"

        self._build_ui()

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=20)
        container.pack(fill="both", expand=True)
        container.columnconfigure(1, weight=1)

        ttk.Label(container, text="Release Type").grid(row=0, column=0, sticky="w", pady=8)
        ttk.Combobox(
            container,
            textvariable=self.release_type_var,
            state="readonly",
            values=["SIT", "UAT", "PROD"],
        ).grid(row=0, column=1, sticky="ew", pady=8)

        ttk.Label(container, text="Request Date").grid(row=1, column=0, sticky="w", pady=8)
        self.date_picker = DateEntry(
            container,
            textvariable=self.request_date_var,
            date_pattern="dd-mm-yyyy",
            firstweekday="monday",
            width=18,
        )
        self.date_picker.grid(
            row=1, column=1, sticky="ew", pady=8
        )
        self.date_picker.bind("<Button-1>", self._show_date_picker, add="+")
        self.date_picker.bind("<FocusIn>", self._show_date_picker, add="+")
        self.date_picker.bind("<FocusOut>", self._hide_date_picker_on_focus_out, add="+")
        self.root.bind_all("<ButtonPress-1>", self._hide_date_picker_on_global_click, add="+")

        ttk.Label(container, text="Input File").grid(row=2, column=0, sticky="w", pady=8)
        input_frame = ttk.Frame(container)
        input_frame.grid(row=2, column=1, sticky="ew", pady=8)
        input_frame.columnconfigure(0, weight=1)
        ttk.Entry(input_frame, textvariable=self.input_file_var).grid(row=0, column=0, sticky="ew")
        ttk.Button(input_frame, text="Browse...", command=self._pick_input_file).grid(
            row=0, column=1, padx=(8, 0)
        )

        ttk.Label(container, text="Sheet").grid(row=3, column=0, sticky="w", pady=8)
        self.sheet_combo = ttk.Combobox(
            container,
            textvariable=self.sheet_name_var,
            state="readonly",
        )
        self.sheet_combo.grid(row=3, column=1, sticky="ew", pady=8)

        ttk.Button(
            container,
            text="Generate Test Result",
            command=self._generate,
        ).grid(row=4, column=1, sticky="e", pady=(20, 8))

        ttk.Label(
            container,
            textvariable=self.status_var,
            foreground="#1f4e79",
            wraplength=560,
            justify="left",
        ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(16, 0))

    def _pick_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select Release Note .xlsx file",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not file_path:
            return

        validation_error = self._validate_input_file(Path(file_path))
        if validation_error:
            messagebox.showerror("Invalid file", validation_error)
            self.input_file_var.set("")
            self.sheet_combo["values"] = []
            self.sheet_name_var.set("")
            self.status_var.set("The selected file is not a valid .xlsx workbook.")
            return

        self.input_file_var.set(file_path)
        try:
            sheets = find_matching_sheets(Path(file_path))
        except Exception as exc:
            messagebox.showerror("Invalid file", str(exc))
            self.status_var.set("The selected file is invalid or cannot be read.")
            return

        if not sheets:
            messagebox.showerror(
                "No matching sheet",
                "No sheet with the required columns was found in this workbook.",
            )
            self.sheet_combo["values"] = []
            self.sheet_name_var.set("")
            self.status_var.set("No valid sheet was found.")
            return

        self.available_sheets = sheets
        self.sheet_combo["values"] = sheets
        self.sheet_name_var.set(sheets[0])
        self.status_var.set(
            f"Found {len(sheets)} valid sheet(s). Select a sheet and click Generate."
        )

    def _generate(self) -> None:
        input_file = self.input_file_var.get().strip()
        sheet_name = self.sheet_name_var.get().strip()
        release_type = self.release_type_var.get().strip()
        request_date = self.request_date_var.get().strip()

        if not input_file:
            messagebox.showwarning("Missing input", "Please select a Release Note file.")
            return
        validation_error = self._validate_input_file(Path(input_file))
        if validation_error:
            messagebox.showerror("Invalid file", validation_error)
            return
        if not sheet_name:
            messagebox.showwarning("Missing sheet", "Please select a sheet to generate.")
            return
        if not request_date:
            messagebox.showwarning("Missing date", "Please select a date for the output file name.")
            return
        if not self._validate_request_date(request_date):
            messagebox.showwarning(
                "Invalid date",
                "The selected date is invalid. Please use the date picker.",
            )
            return

        try:
            output_path = generate_test_result(
                input_path=Path(input_file),
                sheet_name=sheet_name,
                release_type=release_type,
                request_date=request_date,
                output_dir=self.output_dir,
                taiga_config_path=self.taiga_config_path,
            )
        except Exception as exc:
            messagebox.showerror("Generate failed", str(exc))
            self.status_var.set("Generation failed. Check the error message for details.")
            return

        taiga_note = ""
        if not self.taiga_config_path.exists():
            taiga_note = "\nNo taiga.local.json file was found, so Status and QC PIC were left blank."

        self.status_var.set(f"Output created: {output_path}{taiga_note}")
        messagebox.showinfo("Success", f"Output created successfully:\n{output_path}{taiga_note}")
        self._open_output_folder(output_path.parent)

    def _resolve_runtime_root(self) -> Path:
        if getattr(sys, "frozen", False):
            repo_root = Path.home() / "SC-NEW" / "release-note-test-result-tool"
            if repo_root.exists():
                return repo_root
            return Path(sys.executable).resolve().parent
        return Path(__file__).resolve().parents[2]

    def _open_output_folder(self, output_dir: Path) -> None:
        try:
            os.startfile(str(output_dir))
        except OSError:
            pass

    def _show_date_picker(self, _event: object) -> None:
        try:
            self.date_picker.drop_down()
        except Exception:
            pass

    def _hide_date_picker_on_focus_out(self, _event: object) -> None:
        self.root.after(50, self._hide_date_picker_if_needed)

    def _hide_date_picker_on_global_click(self, event: object) -> None:
        self._hide_date_picker_if_needed(getattr(event, "widget", None))

    def _hide_date_picker_if_needed(
        self,
        clicked_widget: object | None = None,
    ) -> None:
        try:
            top_cal = self.date_picker._top_cal
            if not self.date_picker._calendar.winfo_ismapped():
                return

            if clicked_widget is not None:
                clicked_path = str(clicked_widget)
                if clicked_widget == self.date_picker or clicked_path.startswith(str(top_cal)):
                    return
                top_cal.withdraw()
                self.date_picker.state(["!pressed"])
                return

            focused_widget = self.root.focus_get()
            if focused_widget is None:
                top_cal.withdraw()
                self.date_picker.state(["!pressed"])
                return

            focused_path = str(focused_widget)
            if focused_widget == self.date_picker or focused_path.startswith(str(top_cal)):
                return

            top_cal.withdraw()
            self.date_picker.state(["!pressed"])
        except Exception:
            pass

    def _validate_input_file(self, input_path: Path) -> str:
        if not input_path.exists():
            return "The selected file does not exist."
        if input_path.suffix.casefold() != ".xlsx":
            return "Only .xlsx files are supported."
        if not zipfile.is_zipfile(input_path):
            return "The selected .xlsx file is invalid or corrupted."
        return ""

    def _validate_request_date(self, value: str) -> bool:
        try:
            datetime.strptime(value, "%d-%m-%Y")
            return True
        except ValueError:
            return False
