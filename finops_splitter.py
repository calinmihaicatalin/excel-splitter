"""FinOps Excel Splitter: split an Excel workbook into files by department and owner."""
from __future__ import annotations

import copy
import math
import re
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter


def safe_name(value: object) -> str:
    """Return a filesystem-safe name for an output folder or workbook."""
    name = re.sub(r'[\\/:*?"<>|]', "_", str(value)).strip().strip(".")
    return (name[:80] or "Unnamed")


def copy_style(source, target) -> None:
    if source.has_style:
        target.font = copy.copy(source.font)
        target.fill = copy.copy(source.fill)
        target.border = copy.copy(source.border)
        target.alignment = copy.copy(source.alignment)
        target.number_format = source.number_format
        target.protection = copy.copy(source.protection)


def copy_cell(source, target) -> None:
    target.value = source.value
    copy_style(source, target)
    if source.hyperlink:
        target._hyperlink = copy.copy(source.hyperlink)


def copy_sheet_header(source, target, header_row: int) -> None:
    """Copy the header/template area and sheet layout into an output sheet."""
    for key, dimension in source.column_dimensions.items():
        target.column_dimensions[key].width = dimension.width
        target.column_dimensions[key].hidden = dimension.hidden
    for row, dimension in source.row_dimensions.items():
        if row <= header_row:
            target.row_dimensions[row].height = dimension.height
            target.row_dimensions[row].hidden = dimension.hidden
    for cells in source.iter_rows(min_row=1, max_row=header_row):
        for cell in cells:
            copy_cell(cell, target.cell(cell.row, cell.column))
    for merged in source.merged_cells.ranges:
        if merged.max_row <= header_row:
            target.merge_cells(str(merged))
    target.freeze_panes = source.freeze_panes
    target.sheet_view.showGridLines = source.sheet_view.showGridLines
    target.sheet_view.zoomScale = source.sheet_view.zoomScale
    if source.sheet_properties.tabColor:
        target.sheet_properties.tabColor = copy.copy(source.sheet_properties.tabColor)
    target.page_margins = copy.copy(source.page_margins)
    target.page_setup = copy.copy(source.page_setup)
    target.print_title_rows = source.print_title_rows


class Splitter(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("FinOps Excel Splitter")
        self.geometry("760x580")
        self.minsize(620, 480)
        self.path = tk.StringVar()
        self.output = tk.StringVar()
        self.header = tk.IntVar(value=5)
        self.folder_column = tk.StringVar()
        self.owner_column = tk.StringVar()
        self.frames: dict[str, pd.DataFrame] = {}
        self.workbook = None
        self.departments: dict[str, tk.BooleanVar] = {}
        self._ui()

    def _ui(self):
        root = ttk.Frame(self, padding=16)
        root.pack(fill="both", expand=True)
        ttk.Label(root, text="FinOps Excel Splitter", font=("Segoe UI", 17, "bold")).pack(anchor="w")
        ttk.Label(root, text="Create formatted workbooks for each owner, grouped by department.").pack(anchor="w", pady=(0, 14))

        file_row = ttk.Frame(root)
        file_row.pack(fill="x", pady=4)
        ttk.Entry(file_row, textvariable=self.path).pack(side="left", fill="x", expand=True)
        ttk.Button(file_row, text="Browse…", command=self._browse).pack(side="left", padx=(8, 0))

        config = ttk.Frame(root)
        config.pack(fill="x", pady=6)
        ttk.Label(config, text="Header row:").pack(side="left")
        ttk.Spinbox(config, from_=1, to=100, width=5, textvariable=self.header).pack(side="left", padx=(6, 18))
        ttk.Label(config, text="Output folder:").pack(side="left")
        ttk.Entry(config, textvariable=self.output).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(config, text="Browse…", command=self._browse_output).pack(side="left")
        ttk.Button(root, text="Load Workbook", command=self._load).pack(anchor="w", pady=8)

        columns = ttk.Frame(root)
        columns.pack(fill="x", pady=6)
        ttk.Label(columns, text="Department / folder column:").grid(row=0, column=0, sticky="w")
        self.dept_box = ttk.Combobox(columns, textvariable=self.folder_column, state="disabled")
        self.dept_box.grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Label(columns, text="Owner / file column:").grid(row=1, column=0, sticky="w", pady=5)
        self.owner_box = ttk.Combobox(columns, textvariable=self.owner_column, state="disabled")
        self.owner_box.grid(row=1, column=1, sticky="ew", padx=8, pady=5)
        columns.columnconfigure(1, weight=1)
        self.dept_box.bind("<<ComboboxSelected>>", lambda _event: self._show_departments())

        dept_bar = ttk.Frame(root)
        dept_bar.pack(fill="x", pady=(8, 3))
        ttk.Label(dept_bar, text="Departments to include:").pack(side="left")
        ttk.Button(dept_bar, text="Select all", command=lambda: self._toggle(True)).pack(side="right", padx=3)
        ttk.Button(dept_bar, text="Select none", command=lambda: self._toggle(False)).pack(side="right", padx=3)
        self.dept_canvas = tk.Canvas(root, height=130, highlightthickness=0)
        scroll = ttk.Scrollbar(root, orient="vertical", command=self.dept_canvas.yview)
        self.dept_canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.dept_canvas.pack(fill="x")
        self.dept_frame = ttk.Frame(self.dept_canvas)
        self.dept_canvas.create_window((0, 0), window=self.dept_frame, anchor="nw")
        self.dept_frame.bind("<Configure>", lambda _e: self.dept_canvas.configure(scrollregion=self.dept_canvas.bbox("all")))

        self.run_button = ttk.Button(root, text="Run Split", command=self._run, state="disabled")
        self.run_button.pack(anchor="w", pady=10)
        self.progress = ttk.Progressbar(root, mode="determinate")
        self.progress.pack(fill="x", pady=3)
        self.status = tk.StringVar(value="Choose a template workbook to begin.")
        ttk.Label(root, textvariable=self.status, wraplength=700).pack(anchor="w", pady=4)

    def _browse(self):
        path = filedialog.askopenfilename(filetypes=[("Excel workbooks", "*.xlsx *.xlsm")])
        if path:
            self.path.set(path)
            if not self.output.get():
                self.output.set(str(Path(path).parent / "output"))

    def _browse_output(self):
        path = filedialog.askdirectory()
        if path:
            self.output.set(path)

    def _load(self):
        path = Path(self.path.get().strip())
        if not path.is_file():
            messagebox.showerror("Invalid workbook", "Select an existing .xlsx or .xlsm workbook.")
            return
        try:
            row = self.header.get()
            self.workbook = load_workbook(path, keep_vba=path.suffix.lower() == ".xlsm")
            raw = pd.read_excel(path, sheet_name=None, header=row - 1)
            self.frames = {}
            for name, frame in raw.items():
                frame.dropna(how="all", inplace=True)
                frame.columns = [str(c).strip() if str(c).strip() and not str(c).lower().startswith("unnamed:") else f"Unnamed_{i+1}" for i, c in enumerate(frame.columns)]
                self.frames[name] = frame
            columns = next((list(df.columns) for df in self.frames.values() if not df.empty), [])
            if not columns:
                raise ValueError("No data columns were found. Check the header row number.")
            for combo in (self.dept_box, self.owner_box):
                combo["values"] = columns
                combo.configure(state="readonly")
            self.status.set(f"Loaded {len(self.frames)} sheet(s). Select the department and owner columns.")
            self.run_button.configure(state="normal")
            self._show_departments()
        except Exception as exc:
            messagebox.showerror("Load error", str(exc))

    def _show_departments(self):
        column = self.folder_column.get()
        values = set()
        if column:
            for frame in self.frames.values():
                if column in frame:
                    values.update(str(v).strip() for v in frame[column].dropna() if str(v).strip())
        previous = {name: var.get() for name, var in self.departments.items()}
        for widget in self.dept_frame.winfo_children():
            widget.destroy()
        self.departments = {name: tk.BooleanVar(value=previous.get(name, True)) for name in sorted(values)}
        for index, (name, var) in enumerate(self.departments.items()):
            ttk.Checkbutton(self.dept_frame, text=name, variable=var).grid(row=index // 3, column=index % 3, sticky="w", padx=5, pady=2)
        self.dept_canvas.configure(scrollregion=self.dept_canvas.bbox("all"))

    def _toggle(self, value):
        for variable in self.departments.values():
            variable.set(value)

    def _run(self):
        dept_col, owner_col = self.folder_column.get(), self.owner_column.get()
        selected = [name for name, var in self.departments.items() if var.get()]
        if not dept_col or not owner_col or not selected:
            messagebox.showerror("Missing configuration", "Choose both columns and at least one department.")
            return
        output = Path(self.output.get().strip())
        if not str(output):
            messagebox.showerror("Missing output", "Choose an output folder.")
            return
        self.run_button.configure(state="disabled")
        self.progress.configure(value=0, maximum=1)
        threading.Thread(target=self._split, args=(dept_col, owner_col, selected, output), daemon=True).start()

    def _split(self, dept_col, owner_col, selected, output):
        try:
            header = self.header.get()
            grouped: dict[str, dict[str, dict[str, pd.DataFrame]]] = {dept: {} for dept in selected}
            for sheet, frame in self.frames.items():
                if dept_col not in frame or owner_col not in frame:
                    continue
                clean_dept = frame[dept_col].astype(str).str.strip()
                clean_owner = frame[owner_col].astype(str).str.strip()
                for dept in selected:
                    subset = frame[clean_dept == dept]
                    for owner in subset[owner_col].dropna().astype(str).str.strip().unique():
                        if owner and owner.lower() not in ("nan", "none"):
                            grouped[dept].setdefault(owner, {})[sheet] = subset[clean_owner.loc[subset.index] == owner]
            jobs = [(dept, owner, sheets) for dept, owners in grouped.items() for owner, sheets in owners.items()]
            if not jobs:
                raise ValueError("No matching rows found for the selected departments and columns.")
            output.mkdir(parents=True, exist_ok=True)
            for index, (dept, owner, sheets) in enumerate(jobs, 1):
                folder = output / safe_name(dept)
                folder.mkdir(parents=True, exist_ok=True)
                new_book = Workbook()
                new_book.remove(new_book.active)
                for sheet, frame in sheets.items():
                    if sheet not in self.workbook.sheetnames:
                        continue
                    source = self.workbook[sheet]
                    target = new_book.create_sheet(title=sheet)
                    copy_sheet_header(source, target, header)
                    headers = {str(source.cell(header, col).value).strip(): col for col in range(1, source.max_column + 1) if source.cell(header, col).value is not None}
                    for row_index, (_, record) in enumerate(frame.iterrows(), start=header + 1):
                        template_row = header + 1
                        if template_row in source.row_dimensions:
                            target.row_dimensions[row_index].height = source.row_dimensions[template_row].height
                        for name, col in headers.items():
                            value = record.get(name)
                            if isinstance(value, float) and math.isnan(value):
                                value = None
                            cell = target.cell(row_index, col, value)
                            copy_style(source.cell(template_row, col), cell)
                    last_col = get_column_letter(max(headers.values(), default=1))
                    target.auto_filter.ref = f"A{header}:{last_col}{header}"
                    target.freeze_panes = target.cell(header + 1, 1)
                new_book.save(folder / f"{safe_name(owner)}.xlsx")
                self.after(0, lambda n=index: self.progress.configure(value=n))
            self.after(0, lambda: self.status.set(f"Complete: created {len(jobs)} workbook(s) in {output}"))
            self.after(0, lambda: messagebox.showinfo("Split complete", f"Created {len(jobs)} workbook(s).\n\nOutput folder:\n{output}"))
        except Exception as exc:
            self.after(0, lambda: self.status.set(f"Error: {exc}"))
            self.after(0, lambda: messagebox.showerror("Split error", str(exc)))
        finally:
            self.after(0, lambda: self.run_button.configure(state="normal"))


if __name__ == "__main__":
    Splitter().mainloop()
