#!/usr/bin/env python3
"""
Bytecode GUI – Minimal Header→Unicode Bold Wrapper

What it does
- Paste input on top, click Process.
- It normalizes (NFKC), cleans invisibles, then:
  * Splits on the FIRST ':'
  * Converts the header to Unicode math BOLD letters (A–Z/a–z)
  * Wraps as:  ✅【 <BOLD HEADER>】: <body>
- Shows exact processed text in the bottom box.
- Buttons: Process, Copy Output, Save .txt

No external deps (stdlib only). Works on Windows/macOS/Linux.
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import re, unicodedata

# ===== Unicode helpers =====
SMART_SPACES = {
    "\u00A0": " ",  # NBSP
    "\u2007": " ",  # figure space
    "\u202F": " ",  # narrow NBSP
}
ZERO_WIDTH_BIDI = re.compile(r"[\u200B\u200C\u200D\u2060\uFEFF\u200E\u200F\u202A-\u202E\u2066-\u2069]")
CTRL = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]")

def nfkc_clean(s: str) -> str:
    s = unicodedata.normalize("NFKC", s)
    s = s.translate(str.maketrans(SMART_SPACES))
    s = ZERO_WIDTH_BIDI.sub("", s)
    s = CTRL.sub("", s)
    return s

# A..Z → 𝐀..𝐙  and  a..z → 𝐚..𝐳 ; others unchanged
_defA = ord('A'); _defa = ord('a')

def to_math_bold(text: str) -> str:
    out = []
    for ch in text:
        o = ord(ch)
        if 65 <= o <= 90:      # A-Z
            out.append(chr(0x1D400 + (o - 65)))
        elif 97 <= o <= 122:   # a-z
            out.append(chr(0x1D41A + (o - 97)))
        else:
            out.append(ch)
    return ''.join(out)

def transform_line(line: str) -> str:
    if not line:
        return ""
    s = nfkc_clean(line)
    if ":" not in s:
        return f"✅【 {to_math_bold(s.strip())}】"  # no colon case
    head, body = s.split(":", 1)
    head = head.strip()
    bold_head = to_math_bold(head)
    return f"✅【 {bold_head}】: {body.lstrip()}"

# ===== GUI =====
class BytecodeGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bytecode GUI – Header Bold Wrapper")
        self.geometry("920x600")

        # Input frame
        frm_in = ttk.LabelFrame(self, text="Input (paste here)")
        frm_in.pack(fill="both", expand=True, padx=10, pady=8)
        self.txt_in = tk.Text(frm_in, height=10, wrap="word")
        self.txt_in.pack(fill="both", expand=True, padx=8, pady=8)

        # Buttons
        frm_btn = ttk.Frame(self)
        frm_btn.pack(fill="x", padx=10)
        ttk.Button(frm_btn, text="Process", command=self.process).pack(side="left", padx=4)
        ttk.Button(frm_btn, text="Copy Output", command=self.copy_output).pack(side="left", padx=4)
        ttk.Button(frm_btn, text="Save .txt", command=self.save_txt).pack(side="left", padx=4)

        # Output frame
        frm_out = ttk.LabelFrame(self, text="Output (exact text to paste)")
        frm_out.pack(fill="both", expand=True, padx=10, pady=8)
        self.txt_out = tk.Text(frm_out, height=12, wrap="word")
        self.txt_out.pack(fill="both", expand=True, padx=8, pady=8)

        # Status bar
        self.status = ttk.Label(self, text="Ready")
        self.status.pack(fill="x", padx=10, pady=(0,8))

        self.last_text = ""

    def process(self):
        raw = self.txt_in.get("1.0", "end-1c")
        out = transform_line(raw)
        self.txt_out.delete("1.0", "end")
        self.txt_out.insert("1.0", out)
        self.last_text = out
        self.status.config(text=f"Chars: {len(out)}")

    def copy_output(self):
        if not self.last_text:
            messagebox.showwarning("Nothing to copy", "Click Process first.")
            return
        self.clipboard_clear()
        self.clipboard_append(self.last_text)
        self.update()
        messagebox.showinfo("Copied", "Output copied to clipboard.")

    def save_txt(self):
        if not self.last_text:
            messagebox.showwarning("Nothing to save", "Click Process first.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text","*.txt"), ("All","*.*")])
        if not path:
            return
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            f.write(self.last_text)
        messagebox.showinfo("Saved", f"Saved to {path}")

if __name__ == "__main__":
    app = BytecodeGUI()
    app.mainloop()
