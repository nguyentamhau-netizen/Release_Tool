from __future__ import annotations

import os
import sys
import threading
import json
import tkinter as tk
import zipfile
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from tkcalendar import Calendar

from .taiga import TaigaConfig
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
        self.date_popup: tk.Toplevel | None = None
        self.calendar_widget: Calendar | None = None

        self.project_root = self._resolve_runtime_root()
        self.output_dir = self.project_root / "output" / "generated"
        self.taiga_config_path = self.project_root / "taiga.local.json"
        self.prefs_path = self.project_root / "user_prefs.json"

        self._load_prefs()
        self._build_ui()

    def _load_prefs(self) -> None:
        if self.prefs_path.exists():
            try:
                prefs = json.loads(self.prefs_path.read_text("utf-8"))
                if "release_type" in prefs:
                    self.release_type_var.set(prefs["release_type"])
                if "input_file" in prefs:
                    self.input_file_var.set(prefs["input_file"])
                    # Auto-load sheets
                    try:
                        sheets = find_matching_sheets(Path(prefs["input_file"]))
                        if sheets:
                            self.available_sheets = sheets
                            self.sheet_combo["values"] = sheets
                            if "sheet_name" in prefs and prefs["sheet_name"] in sheets:
                                self.sheet_name_var.set(prefs["sheet_name"])
                            else:
                                self.sheet_name_var.set(sheets[0])
                    except Exception:
                        pass
            except Exception:
                pass

    def _save_prefs(self) -> None:
        prefs = {
            "release_type": self.release_type_var.get(),
            "input_file": self.input_file_var.get(),
            "sheet_name": self.sheet_name_var.get(),
        }
        try:
            self.prefs_path.write_text(json.dumps(prefs), "utf-8")
        except Exception:
            pass
            
    def _open_settings(self):
        popup = tk.Toplevel(self.root)
        popup.title("Taiga Settings")
        popup.geometry("400x450")
        popup.resizable(False, False)
        popup.transient(self.root)
        popup.grab_set()

        try:
            config = TaigaConfig.from_path(self.taiga_config_path)
        except Exception:
            config = TaigaConfig("", "", "", "", tuple())

        frame = ttk.Frame(popup, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Base URL:").grid(row=0, column=0, sticky="w", pady=4)
        base_url_var = tk.StringVar(value=config.base_url)
        ttk.Entry(frame, textvariable=base_url_var, width=40).grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(frame, text="Project Slug:").grid(row=1, column=0, sticky="w", pady=4)
        slug_var = tk.StringVar(value=config.project_slug)
        ttk.Entry(frame, textvariable=slug_var, width=40).grid(row=1, column=1, sticky="w", pady=4)

        ttk.Label(frame, text="Username:").grid(row=2, column=0, sticky="w", pady=4)
        user_var = tk.StringVar(value=config.username)
        ttk.Entry(frame, textvariable=user_var, width=40).grid(row=2, column=1, sticky="w", pady=4)

        ttk.Label(frame, text="Password:").grid(row=3, column=0, sticky="w", pady=4)
        pass_var = tk.StringVar(value=config.password)
        ttk.Entry(frame, textvariable=pass_var, width=40, show="*").grid(row=3, column=1, sticky="w", pady=4)

        ttk.Label(frame, text="QC Names (1 per line):").grid(row=4, column=0, columnspan=2, sticky="w", pady=(12, 4))
        qc_text = tk.Text(frame, height=10, width=45)
        qc_text.grid(row=5, column=0, columnspan=2, sticky="w")
        qc_text.insert("1.0", "\n".join(config.qc_names))

        def save_settings():
            names = [n.strip() for n in qc_text.get("1.0", tk.END).split("\n") if n.strip()]
            new_config = TaigaConfig(
                base_url=base_url_var.get().strip(),
                project_slug=slug_var.get().strip(),
                username=user_var.get().strip(),
                password=pass_var.get(),
                qc_names=tuple(names)
            )
            try:
                new_config.save(self.taiga_config_path)
                messagebox.showinfo("Success", "Settings saved!", parent=popup)
                popup.destroy()
            except Exception as e:
                messagebox.showerror("Error", str(e), parent=popup)

        ttk.Button(frame, text="Save Settings", command=save_settings).grid(row=6, column=0, columnspan=2, pady=(16, 0))

    def _show_logs(self, logs):
        popup = tk.Toplevel(self.root)
        popup.title("Generate Logs")
        popup.geometry("800x400")
        
        text = tk.Text(popup, wrap="word")
        text.pack(fill="both", expand=True, padx=10, pady=10)
        
        for log in logs:
            text.insert(tk.END, f"[{log.status}] Row {log.row_no} | {log.item_type} | Ref {log.ref_id} | {log.message}\n")
        text.config(state="disabled")

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=20)
        container.pack(fill="both", expand=True)
        container.columnconfigure(1, weight=1)
        style = ttk.Style(self.root)
        style.configure(
            "DatePicker.TEntry",
            fieldbackground="white",
            foreground="black",
        )
        style.map(
            "DatePicker.TEntry",
            fieldbackground=[("readonly", "white")],
            foreground=[("readonly", "black")],
        )

        ttk.Label(container, text="Release Type").grid(row=0, column=0, sticky="w", pady=8)
        ttk.Combobox(
            container,
            textvariable=self.release_type_var,
            state="readonly",
            values=["SIT", "UAT", "PROD"],
        ).grid(row=0, column=1, sticky="ew", pady=8)

        ttk.Label(container, text="Request Date").grid(row=1, column=0, sticky="w", pady=8)
        date_frame = ttk.Frame(container)
        date_frame.grid(row=1, column=1, sticky="ew", pady=8)
        date_frame.columnconfigure(0, weight=1)
        self.date_picker = ttk.Entry(
            date_frame,
            textvariable=self.request_date_var,
            state="readonly",
            width=18,
            style="DatePicker.TEntry",
        )
        self.date_picker.grid(row=0, column=0, sticky="ew")
        self.date_picker.bind("<Button-1>", self._show_date_picker, add="+")
        ttk.Button(date_frame, text="Pick Date", command=lambda: self._show_date_picker(None)).grid(
            row=0, column=1, padx=(8, 0)
        )
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

        btn_frame = ttk.Frame(container)
        btn_frame.grid(row=4, column=1, sticky="e", pady=(20, 8))
        
        ttk.Button(btn_frame, text="⚙️ Settings", command=self._open_settings).pack(side="left", padx=8)
        self.generate_btn = ttk.Button(btn_frame, text="Generate Test Result", command=self._generate)
        self.generate_btn.pack(side="left")
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(container, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=6, column=0, columnspan=2, sticky="ew", pady=8)
        self.progress_bar.grid_remove()  # Hide initially

        ttk.Label(
            container,
            textvariable=self.status_var,
            foreground="#1f4e79",
            wraplength=560,
            justify="left",
        ).grid(row=6, column=0, columnspan=2, sticky="w", pady=(16, 0))

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
        self._save_prefs()
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

        self._save_prefs()
        self.generate_btn.config(state="disabled")
        self.progress_bar.grid()
        self.progress_var.set(0)
        self.status_var.set("Generating... Please wait.")
        
        def run():
            try:
                def progress(completed, total):
                    pct = (completed / total) * 100
                    self.root.after(0, lambda: self.progress_var.set(pct))
                    self.root.after(0, lambda: self.status_var.set(f"Generating... {completed}/{total} processed."))
                    
                output_path, taiga_logs = generate_test_result(
                    input_path=Path(input_file),
                    sheet_name=sheet_name,
                    release_type=release_type,
                    request_date=request_date,
                    output_dir=self.output_dir,
                    taiga_config_path=self.taiga_config_path,
                    progress_callback=progress
                )
                self.root.after(0, lambda: self._on_generate_success(output_path, taiga_logs))
            except Exception as exc:
                self.root.after(0, lambda: self._on_generate_error(str(exc)))

        threading.Thread(target=run, daemon=True).start()

    def _on_generate_success(self, output_path, taiga_logs):
        self.generate_btn.config(state="normal")
        self.progress_bar.grid_remove()
        
        taiga_note = ""
        if not self.taiga_config_path.exists():
            taiga_note = "\\nNo taiga.local.json file was found, so Status and QC PIC were left blank."

        self.status_var.set(f"Output created: {output_path}{taiga_note}")
        
        # Build success window with log button
        success_popup = tk.Toplevel(self.root)
        success_popup.title("Success")
        success_popup.geometry("500x200")
        success_popup.transient(self.root)
        
        ttk.Label(success_popup, text="Output created successfully:", font=("", 10, "bold")).pack(pady=(20, 5))
        ttk.Label(success_popup, text=str(output_path), wraplength=450).pack(pady=5)
        
        btn_frame = ttk.Frame(success_popup)
        btn_frame.pack(pady=20)
        
        ttk.Button(btn_frame, text="Open Folder", command=lambda: [self._open_output_folder(output_path.parent), success_popup.destroy()]).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="View Logs", command=lambda: self._show_logs(taiga_logs)).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Close", command=success_popup.destroy).pack(side="left", padx=10)

    def _on_generate_error(self, error_msg):
        self.generate_btn.config(state="normal")
        self.progress_bar.grid_remove()
        messagebox.showerror("Generate failed", error_msg)
        self.status_var.set("Generation failed. Check the error message for details.")

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
            if self.date_popup is not None and self.date_popup.winfo_exists():
                return

            popup = tk.Toplevel(self.root)
            popup.withdraw()
            popup.overrideredirect(True)
            popup.transient(self.root)

            current_date = datetime.strptime(self.request_date_var.get(), "%d-%m-%Y")
            calendar = Calendar(
                popup,
                selectmode="day",
                date_pattern="dd-mm-yyyy",
                firstweekday="monday",
                year=current_date.year,
                month=current_date.month,
                day=current_date.day,
            )
            calendar.pack()
            calendar.bind("<<CalendarSelected>>", self._on_calendar_selected)

            self.date_popup = popup
            self.calendar_widget = calendar

            x = self.date_picker.winfo_rootx()
            y = self.date_picker.winfo_rooty() + self.date_picker.winfo_height()
            popup.geometry(f"+{x}+{y}")
            popup.deiconify()
            popup.lift()
        except Exception:
            pass

    def _on_calendar_selected(self, _event: object) -> None:
        if self.calendar_widget is None:
            return
        self.request_date_var.set(self.calendar_widget.get_date())
        self._close_date_picker()

    def _hide_date_picker_on_global_click(self, event: object) -> None:
        self._hide_date_picker_if_needed(getattr(event, "widget", None))

    def _hide_date_picker_if_needed(
        self,
        clicked_widget: object | None = None,
    ) -> None:
        try:
            if self.date_popup is None or not self.date_popup.winfo_exists():
                return

            if clicked_widget is not None:
                clicked_path = str(clicked_widget)
                if clicked_widget == self.date_picker or clicked_path.startswith(str(self.date_popup)):
                    return
                self._close_date_picker()
                return

            focused_widget = self.root.focus_get()
            if focused_widget is None:
                self._close_date_picker()
                return

            focused_path = str(focused_widget)
            if focused_widget == self.date_picker or focused_path.startswith(str(self.date_popup)):
                return

            self._close_date_picker()
        except Exception:
            pass

    def _close_date_picker(self) -> None:
        try:
            if self.date_popup is not None and self.date_popup.winfo_exists():
                self.date_popup.destroy()
        except Exception:
            pass
        self.date_popup = None
        self.calendar_widget = None

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
