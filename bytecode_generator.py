#!/usr/bin/env python3
"""
Bytecode Generator – CLI MVP (Step 1)

Goals:
- Deterministic, cross‑platform safe text transformation
- Presets: plain UTF‑8, ASCII‑only, HTML‑safe, CSV/Excel‑safe, JSON‑safe
- Normalization: NFC/NFKC/None (default NFKC)
- Cleaning: remove zero‑width/BiDi controls, replace smart quotes/dashes, NBSP→space
- Excel safety: neutralize formula prefixes, CRLF endings, optional BOM
- Verification: byte length, SHA‑256, UTF‑8 hex, Base64

Usage examples:
  # 1) Plain UTF‑8 (default NFKC) from inline text
  python bytecode_generator_cli.py --mode plain --text "✅【 𝐔𝐋𝐓𝐈𝐌𝐀𝐓𝐄 】"

  # 2) ASCII‑only (emoji dropped, math bold flattened) reading from stdin
  type NUL | set /p="Your text here" | python bytecode_generator_cli.py -m ascii
  # or Linux/macOS:
  echo "Your text here" | python3 bytecode_generator_cli.py -m ascii

  # 3) HTML‑safe export to file
  python bytecode_generator_cli.py -m html --text "a < b & c > d" -o out.html

  # 4) CSV/Excel‑safe single cell
  python bytecode_generator_cli.py -m csv --text "=1+2" --csv-delim , --bom --newline crlf -o row.csv

  # 5) CSV/Excel‑safe multi‑cell row (split on ||)
  python bytecode_generator_cli.py -m csv --csv-values "SKU001||=1+2||Cool Chair, Black" -o row.csv

  # 6) JSON‑safe
  python bytecode_generator_cli.py -m json --text "Line\nBreak\tTab\"Quote\""

Note: Only standard library is used.
"""
from __future__ import annotations
import argparse
import base64
import csv
import hashlib
import html
import io
import json
import re
import sys
import unicodedata
from typing import List, Tuple

# =========================
# Character maps & regexes
# =========================
SMART_PUNCT_MAP = {
    # Quotes
    "\u2018": "'",  # ‘
    "\u2019": "'",  # ’
    "\u201C": '"',  # “
    "\u201D": '"',  # ”
    "\u2032": "'",  # ′ prime
    "\u2033": '"',  # ″ double prime
    # Dashes
    "\u2012": "-",  # figure dash
    "\u2013": "-",  # en dash
    "\u2014": "-",  # em dash
    "\u2212": "-",  # minus sign
    # Ellipsis
    "\u2026": "...",
    # Spaces
    "\u00A0": " ",   # NBSP
    "\u2007": " ",   # figure space
    "\u202F": " ",   # narrow NBSP
}

# Zero-width & BiDi controls to remove
ZW_BIDI_PATTERN = re.compile(
    "[\u200B\u200C\u200D\u2060\uFEFF\u200E\u200F\u202A\u202B\u202C\u202D\u202E\u2066\u2067\u2068\u2069]"
)

# C0/C1 controls (except \t, \n, \r) to strip
CTRL_PATTERN = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F]")

# Excel formula injection risky leading chars
EXCEL_RISKY_PREFIX = tuple(["=", "+", "-", "@", "\t", "\r"])

# =========================
# Normalization & Cleaning
# =========================

def normalize_text(s: str, form: str = "NFKC") -> str:
    if form.upper() in {"NFC", "NFD", "NFKC", "NFKD"}:
        return unicodedata.normalize(form.upper(), s)
    return s


def replace_smart_punct(s: str) -> str:
    return s.translate(str.maketrans(SMART_PUNCT_MAP))


def strip_zero_width_and_controls(s: str) -> Tuple[str, int, int]:
    before = len(s)
    s = ZW_BIDI_PATTERN.sub("", s)
    after_zw = len(s)
    s = CTRL_PATTERN.sub("", s)
    after_ctrl = len(s)
    return s, (before - after_zw), (after_zw - after_ctrl)


def ascii_fold(s: str) -> str:
    """Best-effort ASCII folding without external deps.
    Removes diacritics; drops un-representable codepoints.
    """
    # NFKD -> remove combining marks
    s = unicodedata.normalize("NFKD", s)
    out = []
    for ch in s:
        if unicodedata.combining(ch):
            continue
        cp = ord(ch)
        if cp < 128:
            out.append(ch)
        else:
            # handle a few common symbols
            if ch == "°":
                out.append(" degrees")
            # else drop
    return "".join(out)


# =========================
# Preset transformers
# =========================

def preset_plain(s: str) -> str:
    return s

def mark_header_bold(s: str, mode: str) -> str:
    if ":" not in s:
        return s
    head, tail = s.split(":", 1)
    if mode == "html":
        return f"<strong>{head}</strong>:{tail}"
    else:
        # plain, ascii, csv, json → markdown style
        return f"**{head}**:{tail}"


def preset_ascii(s: str) -> str:
    return ascii_fold(s)


def preset_html(s: str, br: bool = False) -> str:
    esc = html.escape(s, quote=True)
    return esc.replace("\n", "<br>\n") if br else esc


def neutralize_excel_cell(cell: str, strategy: str = "apostrophe") -> str:
    """Prevent formula execution when pasted/opened in Excel.
    strategy: 'apostrophe' -> prefix '  |  'space' -> prefix space
    """
    if not cell:
        return cell
    if cell.startswith(EXCEL_RISKY_PREFIX):
        return ("'" + cell) if strategy == "apostrophe" else (" " + cell)
    return cell


def csv_quote_cell(cell: str, quotechar: str = '"') -> str:
    """RFC4180 quoting for a single cell (no delimiter knowledge needed)."""
    need_quote = any(c in cell for c in ["\n", "\r", quotechar, ","]) or cell.startswith(" ") or cell.endswith(" ")
    if need_quote:
        return quotechar + cell.replace(quotechar, quotechar * 2) + quotechar
    return cell


def preset_csv_row(
    values: List[str],
    delimiter: str = ",",
    neutralize: bool = True,
    strategy: str = "apostrophe",
    quotechar: str = '"',
) -> str:
    safe_vals = []
    for v in values:
        v2 = neutralize_excel_cell(v, strategy=strategy) if neutralize else v
        safe_vals.append(csv_quote_cell(v2, quotechar=quotechar))
    return delimiter.join(safe_vals)


def preset_json(s: str) -> str:
    # Represent as a JSON string literal (not a full object)
    return json.dumps(s, ensure_ascii=False)


# =========================
# Encoding / finalization
# =========================

def finalize_text_bytes(text: str, newline: str = "lf", bom: bool = False, encoding: str = "utf-8") -> bytes:
    if newline.lower() == "crlf":
        text = text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    elif newline.lower() == "lf":
        text = text.replace("\r\n", "\n").replace("\r", "\n")
    else:
        raise ValueError("newline must be 'lf' or 'crlf'")

    data = text.encode(encoding)
    if bom and encoding.lower().replace("-", "") == "utf8":
        data = b"\xef\xbb\xbf" + data
    return data


# =========================
# Pipeline
# =========================

def process_text(
    raw: str,
    mode: str = "plain",
    normalize: str = "NFKC",
    replace_smart: bool = True,
    strip_zw: bool = True,
    strip_ctrl: bool = True,
    html_br: bool = False,
    csv_values: List[str] | None = None,
    csv_delim: str = ",",
    csv_neutralize: bool = True,
    csv_strategy: str = "apostrophe",
    quotechar: str = '"',
) -> Tuple[str, dict]:
    # Normalize
    s = normalize_text(raw, normalize)

    # Smart punct
    if replace_smart:
        s = replace_smart_punct(s)

    # Strip zero-width & controls
    zw_count = 0
    ctrl_count = 0
    if strip_zw or strip_ctrl:
        s2, zw_removed, ctrl_removed = strip_zero_width_and_controls(s)
        s = s2
        zw_count += zw_removed
        ctrl_count += ctrl_removed

    # Preset
    if mode == "plain":
        out = preset_plain(s)
    elif mode == "ascii":
        out = preset_ascii(s)
    elif mode == "html":
        out = preset_html(s, br=html_br)
    elif mode == "csv":
        if csv_values is not None:
            out = preset_csv_row(csv_values, delimiter=csv_delim, neutralize=csv_neutralize, strategy=csv_strategy, quotechar=quotechar)
        else:
            # single-cell CSV as a row
            out = preset_csv_row([s], delimiter=csv_delim, neutralize=csv_neutralize, strategy=csv_strategy, quotechar=quotechar)
    elif mode == "json":
        out = preset_json(s)
    else:
        raise ValueError(f"Unknown mode: {mode}")
    
    # Preset
    if mode in ("plain", "ascii", "html", "csv", "json"):
        s = mark_header_bold(s, mode)


    meta = {
        "zw_removed": zw_count,
        "ctrl_removed": ctrl_count,
    }
    return out, meta


# =========================
# CLI
# =========================

def main():
    p = argparse.ArgumentParser(description="Bytecode Generator – CLI MVP")
    p.add_argument("-m", "--mode", choices=["plain", "ascii", "html", "csv", "json"], default="plain")
    p.add_argument("--text", help="Input text (if omitted, read from stdin)")
    p.add_argument("--csv-values", help="CSV mode: multi‑cell row, fields separated by '||'")
    p.add_argument("--normalize", choices=["none", "NFC", "NFD", "NFKC", "NFKD"], default="NFKC")
    p.add_argument("--no-smart", action="store_true", help="Do not replace smart quotes/dashes")
    p.add_argument("--keep-zw", action="store_true", help="Keep zero‑width/BiDi controls")
    p.add_argument("--keep-ctrl", action="store_true", help="Keep C0/C1 control chars")

    # HTML options
    p.add_argument("--html-br", action="store_true", help="HTML mode: convert newlines to <br>")

    # CSV/Excel options
    p.add_argument("--csv-delim", default=",", help="CSV delimiter (default ,)")
    p.add_argument("--quotechar", default='"', help='CSV quote character (default ")')

    p.add_argument("--no-neutralize", action="store_true", help="CSV: do not neutralize Excel formulas")
    p.add_argument("--neutralize-strategy", choices=["apostrophe", "space"], default="apostrophe")

    # Output & encoding
    p.add_argument("-o", "--out", help="Write output bytes to file (exact bytes)")
    p.add_argument("--newline", choices=["lf", "crlf"], default="lf")
    p.add_argument("--bom", action="store_true", help="Prefix UTF‑8 BOM (useful for Excel)")
    p.add_argument("--encoding", default="utf-8", help="Text encoding (default utf-8)")

    # Display controls
    p.add_argument("--show-hex", action="store_true", help="Print full hex dump of bytes")
    p.add_argument("--show-base64", action="store_true", help="Print Base64 of bytes")

    args = p.parse_args()

    # Acquire input
    if args.csv_values and args.text:
        sys.exit("Use either --text or --csv-values, not both.")

    if args.csv_values:
        values = args.csv_values.split("||")
        raw = ""  # unused
    else:
        if args.text is not None:
            raw = args.text
        else:
            raw = sys.stdin.read()

    norm = args.normalize
    if norm == "none":
        norm = "NONE"

    out_text, meta = process_text(
        raw=raw,
        mode=args.mode,
        normalize=norm,
        replace_smart=(not args.no_smart),
        strip_zw=(not args.keep_zw),
        strip_ctrl=(not args.keep_ctrl),
        html_br=args.html_br,
        csv_values=(args.csv_values.split("||") if args.csv_values else None),
        csv_delim=args.csv_delim,
        csv_neutralize=(not args.no_neutralize),
        csv_strategy=args.neutralize_strategy,
        quotechar=args.quotechar,
    )

    # Finalize bytes (newline, encoding, BOM)
    data = finalize_text_bytes(out_text, newline=args.newline, bom=args.bom, encoding=args.encoding)

    # Stats
    sha = hashlib.sha256(data).hexdigest()
    print("=== OUTPUT TEXT (preview) ===")
    # Preview should avoid printing binary nonsense – we print the logical text before encoding/newline changes
    preview = out_text
    if len(preview) > 1000:
        preview = preview[:1000] + "\n...[truncated]"
    print(preview)

    print("\n=== BYTE STATS ===")
    print(f"Bytes: {len(data)} | Encoding: {args.encoding} | Newline: {args.newline} | BOM: {args.bom}")
    print(f"SHA-256: {sha}")

    if args.show_hex:
        print("\n=== HEX (UTF-8 bytes) ===")
        print(data.hex())

    if args.show_base64:
        print("\n=== Base64 (UTF-8 bytes) ===")
        print(base64.b64encode(data).decode("ascii"))

    # Write file if requested (exact bytes)
    if args.out:
        with open(args.out, "wb") as f:
            f.write(data)
        print(f"\n[Saved exact bytes to {args.out}]")


if __name__ == "__main__":
    main()
