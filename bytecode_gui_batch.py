#!/usr/bin/env python3
"""
Bytecode GUI Batch – Select columns in CSV/XLSX and process to new file.

What it does:
- Choose a .xlsx or .csv file
- Detect header row (first row)
- Let you select one or more columns to process
- For each selected column, creates <ColumnName>_Processed with:
    ✅【 <MATH-BOLD(header-before-first-colon)>】: <rest>
- Saves to a new file in the same folder (or Save As)

Notes:
- CSV is saved as UTF-8 with BOM and CRLF (Excel-friendly)
- XLSX is saved via openpyxl (keeps other sheets as-is only for the active one? -> We only process the active sheet)
"""

import os, sys, csv, re, unicodedata
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Optional import for XLSX (only required when handling .xlsx)
try:
    from openpyxl import load_workbook, Workbook
    OPENPYXL_AVAILABLE = True
except Exception:
    OPENPYXL_AVAILABLE = False

# ---------- Transform helpers (same rule you approved) ----------

SMART_SPACES = {
    "\u00A0": " ",  # NBSP
    "\u2007": " ",  # figure space
    "\u202F": " ",  # narrow NBSP
}
ZERO_WIDTH_BIDI = re.compile(r"[\u200B\u200C\u200D\u2060\uFEFF\u200E\u200F\u202A-\u202E\u2066-\u2069]")
CTRL = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]")

def _to_math_bold(s: str) -> str:
    """A..Z→𝐀..𝐙, a..z→𝐚..𝐳; others unchanged."""
    out = []
    for ch in s:
        o = ord(ch)
        if 65 <= o <= 90:          # A-Z
            out.append(chr(0x1D400 + (o - 65)))
        elif 97 <= o <= 122:       # a-z
            out.append(chr(0x1D41A + (o - 97)))
        else:
            out.append(ch)
    return "".join(out)

def transform_line(line: str) -> str:
    """Apply header→math-bold + ✅【 … 】:  wrapping. Keep body content."""
    if line is None:
        return ""
    if not isinstance(line, str):
        line = str(line)

    s = unicodedata.normalize("NFKC", line)
    s = s.translate(str.maketrans(SMART_SPACES))
    s = ZERO_WIDTH_BIDI.sub("", s)
    s = CTRL.sub("", s)

    if ":" not in s:
        # No colon → just math-bold whole line with wrapper
        return f"✅【 {_to_math_bold(s.strip())}】"

    head, body = s.split(":", 1)
    head = head.strip()
    bold_head = _to_math_bold(head)
    return f"✅【 {bold_head}】: {body.lstrip()}"

# ---------- File IO helpers ----------

def read_csv(path: str):
    """Return headers(list[str]), rows(list[list[str]]). Assumes first row is header."""
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        rows = list(reader)
    if not rows:
        return [], []
    headers = rows[0]
    data = rows[1:]
    return headers, data

def write_csv(path: str, headers, rows):
    """Write CSV as UTF-8 BOM + CRLF."""
    # Write BOM
    with open(path, "wb") as fbin:
        fbin.write(b"\xef\xbb\xbf")
    # Append actual CSV
    with open(path, "ab", buffering=0) as fab:
        writer = csv.writer(fab, lineterminator="\r\n")
        writer.writerow(headers)
        for r in rows:
            writer.writerow(r)

def read_xlsx(path: str):
    """Return headers(list[str]), rows(list[list[Any]]), and workbook & ws for saving."""
    if not OPENPYXL_AVAILABLE:
        raise RuntimeError("openpyxl not installed. Install with: pip install openpyxl")
    wb = load_workbook(path)
    ws = wb.active
    max_row = ws.max_row
    max_col = ws.max_column
    if max_row < 1:
        return [], [], wb, ws
    headers = []
    for c in range(1, max_col + 1):
        headers.append(ws.cell(row=1, column=c).value if ws.cell(row=1, column=c).value is not None else "")
    rows = []
    for r in range(2, max_row + 1):
        row = []
        for c in range(1, max_col + 1):
            row.append(ws.cell(row=r, column=c).value)
        rows.append(row)
    return headers, rows, wb, ws

def write_xlsx_processed(path_out: str, wb, ws, headers, rows):
    """Overwrite only active sheet with provided headers/rows, keep workbook object."""
    # Clear the active sheet (simple approach: create new sheet and remove old)
    title = ws.title
    wb.remove(ws)
    new_ws = wb.create_sheet(title=title, index=0)
    # Write header
    for c, h in enumerate(headers, start=1):
        new_ws.cell(row=1, column=c, value=h)
    # Write rows
    for r_idx, row in enumerate(rows, start=2):
        for c_idx, val in enumerate(row, start=1):
            new_ws.cell(row=r_idx, column=c_idx, value=val)
    wb.save(path_out)

# ---------- GUI ----------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bytecode Batch – Column Processor")
        self.geometry("820x540")

        self.file_path = None
        self.file_type = None  # "csv" or "xlsx"
        self.headers = []
        self.rows = []
        self.wb = None
        self.ws = None

        # Top controls
        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=10)

        ttk.Button(top, text="Open CSV/XLSX", command=self.open_file).pack(side="left")
        self.lbl_file = ttk.Label(top, text="No file loaded")
        self.lbl_file.pack(side="left", padx=10)

        # Column select
        frame_cols = ttk.LabelFrame(self, text="Select columns to process (header row)")
        frame_cols.pack(fill="both", expand=True, padx=10, pady=10)

        self.listbox = tk.Listbox(frame_cols, selectmode=tk.MULTIPLE, height=12)
        self.listbox.pack(fill="both", expand=True, padx=8, pady=8)

        # Bottom buttons
        bottom = ttk.Frame(self)
        bottom.pack(fill="x", padx=10, pady=10)

        ttk.Button(bottom, text="Process & Save", command=self.process_and_save).pack(side="left", padx=4)
        ttk.Button(bottom, text="Preview First Row", command=self.preview_first).pack(side="left", padx=4)

        self.status = ttk.Label(self, text="Ready")
        self.status.pack(fill="x", padx=10, pady=(0,10))

    # ----- Actions -----

    def open_file(self):
        path = filedialog.askopenfilename(
            title="Open CSV or XLSX",
            filetypes=[("Excel Workbook", "*.xlsx"), ("CSV (Comma delimited)", "*.csv"), ("All files", "*.*")]
        )
        if not path:
            return

        ext = os.path.splitext(path)[1].lower()
        try:
            if ext == ".csv":
                headers, rows = read_csv(path)
                self.file_type = "csv"
                self.headers, self.rows = headers, rows
                self.wb = self.ws = None
            elif ext == ".xlsx":
                if not OPENPYXL_AVAILABLE:
                    messagebox.showerror("Missing dependency", "Install openpyxl: pip install openpyxl")
                    return
                headers, rows, wb, ws = read_xlsx(path)
                self.file_type = "xlsx"
                self.headers, self.rows = headers, rows
                self.wb, self.ws = wb, ws
            else:
                messagebox.showerror("Unsupported", "Please choose a .csv or .xlsx file.")
                return
        except Exception as e:
            messagebox.showerror("Error opening file", str(e))
            return

        self.file_path = path
        self.lbl_file.config(text=os.path.basename(path))
        self.populate_listbox()
        self.status.config(text=f"Loaded {len(self.rows)} rows, {len(self.headers)} columns")

    def populate_listbox(self):
        self.listbox.delete(0, tk.END)
        for i, h in enumerate(self.headers):
            label = str(h) if h is not None else f"Column {i+1}"
            self.listbox.insert(tk.END, label)

    def get_selected_columns(self):
        idxs = list(self.listbox.curselection())
        return idxs

    def preview_first(self):
        idxs = self.get_selected_columns()
        if not idxs:
            messagebox.showwarning("No selection", "Select at least one column.")
            return
        if not self.rows:
            messagebox.showwarning("No rows", "The file has no data rows.")
            return
        row0 = self.rows[0][:]
        # Create preview row
        preview_pairs = []
        for c in idxs:
            val = row0[c] if c < len(row0) else ""
            preview_pairs.append((self.headers[c], transform_line(val)))
        # Show a simple preview modal
        txt = []
        for name, processed in preview_pairs:
            txt.append(f"[{name}] -> {processed}")
        messagebox.showinfo("Preview (first data row)", "\n\n".join(txt))

    def process_and_save(self):
        if not self.file_path:
            messagebox.showwarning("No file", "Open a CSV/XLSX first.")
            return
        idxs = self.get_selected_columns()
        if not idxs:
            messagebox.showwarning("No selection", "Select at least one column to process.")
            return

        # Build new headers (append _Processed columns)
        new_headers = self.headers[:]
        for c in idxs:
            base = str(self.headers[c]) if self.headers[c] is not None else f"Column{c+1}"
            new_headers.append(f"{base}_Processed")

        # Build new rows
        new_rows = []
        for r in self.rows:
            row_out = r[:]
            for c in idxs:
                val = r[c] if c < len(r) else ""
                row_out.append(transform_line(val))
            new_rows.append(row_out)

        # Decide default output path
        base, ext = os.path.splitext(self.file_path)
        default_out = f"{base}_processed{ext}"

        # Ask where to save
        if self.file_type == "csv":
            out_path = filedialog.asksaveasfilename(
                title="Save Processed CSV",
                defaultextension=".csv",
                initialfile=os.path.basename(default_out),
                filetypes=[("CSV (UTF-8 BOM)", "*.csv"), ("All files", "*.*")]
            )
            if not out_path:
                return
            try:
                write_csv(out_path, new_headers, new_rows)
                self.status.config(text=f"Saved CSV: {out_path}")
                messagebox.showinfo("Saved", f"Processed CSV saved:\n{out_path}")
            except Exception as e:
                messagebox.showerror("Save error", str(e))
        else:
            out_path = filedialog.asksaveasfilename(
                title="Save Processed XLSX",
                defaultextension=".xlsx",
                initialfile=os.path.basename(default_out),
                filetypes=[("Excel Workbook", "*.xlsx"), ("All files", "*.*")]
            )
            if not out_path:
                return
            try:
                write_xlsx_processed(out_path, self.wb, self.ws, new_headers, new_rows)
                self.status.config(text=f"Saved XLSX: {out_path}")
                messagebox.showinfo("Saved", f"Processed XLSX saved:\n{out_path}")
            except Exception as e:
                messagebox.showerror("Save error", str(e))


if __name__ == "__main__":
    app = App()
    app.mainloop()
