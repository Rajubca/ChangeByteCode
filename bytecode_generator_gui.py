#!/usr/bin/env python3
"""
Bytecode Generator – GUI MVP (Step 2)

Tkinter app using the same pipeline from CLI.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import hashlib, base64
from bytecode_generator import process_text, finalize_text_bytes  # import from Step 1

class BytecodeGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bytecode Generator – GUI MVP")
        self.geometry("900x600")

        # ====== Top Frame: Input ======
        frm_top = ttk.LabelFrame(self, text="Input")
        frm_top.pack(fill="both", expand=True, padx=8, pady=4)

        self.txt_input = tk.Text(frm_top, height=8, wrap="word")
        self.txt_input.pack(fill="both", expand=True, padx=4, pady=4)

        # ====== Middle Frame: Options ======
        frm_opts = ttk.LabelFrame(self, text="Options")
        frm_opts.pack(fill="x", padx=8, pady=4)

        ttk.Label(frm_opts, text="Mode:").grid(row=0, column=0, sticky="w")
        self.cmb_mode = ttk.Combobox(frm_opts, values=["plain","ascii","html","csv","json"], state="readonly")
        self.cmb_mode.set("plain")
        self.cmb_mode.grid(row=0, column=1, padx=4, pady=2)

        ttk.Label(frm_opts, text="Normalize:").grid(row=0, column=2, sticky="w")
        self.cmb_norm = ttk.Combobox(frm_opts, values=["NFKC","NFC","NFD","NFKD","none"], state="readonly")
        self.cmb_norm.set("NFKC")
        self.cmb_norm.grid(row=0, column=3, padx=4, pady=2)

        self.var_bom = tk.BooleanVar(value=False)
        self.var_neutralize = tk.BooleanVar(value=True)
        self.var_newline = tk.StringVar(value="lf")

        ttk.Checkbutton(frm_opts, text="Excel Neutralize", variable=self.var_neutralize).grid(row=0, column=4, padx=4)
        ttk.Checkbutton(frm_opts, text="UTF-8 BOM", variable=self.var_bom).grid(row=0, column=5, padx=4)

        ttk.Label(frm_opts, text="Newline:").grid(row=0, column=6, sticky="w")
        self.cmb_newline = ttk.Combobox(frm_opts, values=["lf","crlf"], state="readonly", textvariable=self.var_newline)
        self.cmb_newline.grid(row=0, column=7, padx=4)

        ttk.Button(frm_opts, text="Process", command=self.process).grid(row=0, column=8, padx=8)

        # ====== Bottom Frame: Output ======
        frm_out = ttk.LabelFrame(self, text="Output")
        frm_out.pack(fill="both", expand=True, padx=8, pady=4)

        self.txt_output = tk.Text(frm_out, height=10, wrap="word")
        self.txt_output.pack(fill="both", expand=True, padx=4, pady=4)

        # Metadata
        self.lbl_meta = ttk.Label(frm_out, text="Bytes: 0 | SHA-256: -")
        self.lbl_meta.pack(anchor="w", padx=4, pady=2)

        # Buttons
        frm_btns = ttk.Frame(frm_out)
        frm_btns.pack(fill="x", pady=2)
        ttk.Button(frm_btns, text="Copy Output", command=self.copy_output).pack(side="left", padx=4)
        ttk.Button(frm_btns, text="Export", command=self.export_file).pack(side="left", padx=4)

    def process(self):
        raw = self.txt_input.get("1.0", "end-1c")
        mode = self.cmb_mode.get()
        norm = self.cmb_norm.get()
        if norm == "none":
            norm = "NONE"

        out_text, _ = process_text(
            raw=raw,
            mode=mode,
            normalize=norm,
            csv_neutralize=self.var_neutralize.get()
        )
        data = finalize_text_bytes(out_text, newline=self.var_newline.get(), bom=self.var_bom.get())
        sha = hashlib.sha256(data).hexdigest()
        self.txt_output.delete("1.0","end")
        self.txt_output.insert("1.0", out_text)
        self.lbl_meta.config(text=f"Bytes: {len(data)} | SHA-256: {sha[:16]}...")

        # Keep raw bytes for export
        self.last_bytes = data
        self.last_text = out_text

    def copy_output(self):
        if hasattr(self, "last_text"):
            self.clipboard_clear()
            self.clipboard_append(self.last_text)
            self.update()
            messagebox.showinfo("Copied","Output text copied to clipboard.")

    def export_file(self):
        if not hasattr(self, "last_bytes"):
            messagebox.showerror("Error","No processed output yet.")
            return
        fpath = filedialog.asksaveasfilename(defaultextension=".txt",
            filetypes=[("Text","*.txt"),("CSV","*.csv"),("HTML","*.html"),("All","*.*")])
        if not fpath: return
        with open(fpath,"wb") as f:
            f.write(self.last_bytes)
        messagebox.showinfo("Saved", f"Exact bytes written to {fpath}")

if __name__=="__main__":
    app = BytecodeGUI()
    app.mainloop()
