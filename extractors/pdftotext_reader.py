#!/usr/bin/env python3
"""
Convert PDF(s) from "Chase_history" to text files in "texts" using xpdf/poppler pdftotext.

Defaults:
- Workspace: C:\Users\chris\Documents\2425_Delmar\Budget
- Input:     Chase_history\*.pdf   (file, folder, or wildcard accepted)
- Output:    texts
- Suffix:    -raw.txt
- pdftotext: C:\Program Files\xpdf-tools\bin64\pdftotext.exe

Examples:
  python convert_pdftotext.py
  python convert_pdftotext.py --input "Chase_history\2025-07-12.pdf"
  python convert_pdftotext.py --input "Chase_history\*.pdf" --recurse --overwrite
  python convert_pdftotext.py --pdftotext "C:\Tools\poppler\bin\pdftotext.exe"
"""
import argparse
import glob
import os
import shutil
import subprocess
import sys
from pathlib import Path

DEFAULT_WORKSPACE = r"C:\Users\chris\Documents\2425_Delmar\Budget"
DEFAULT_INPUT     = r"Chase_history\*.pdf"
DEFAULT_OUTPUT    = r"texts"
DEFAULT_SUFFIX    = "-raw.txt"
DEFAULT_PDFTOTEXT = r"C:\Program Files\xpdf-tools\bin64\pdftotext.exe"

def resolve_path(path_str: str, root: Path) -> Path:
    p = Path(path_str)
    if p.is_absolute():
        return p
    return (root / p)

def resolve_existing(path_str: str, root: Path) -> Path:
    p = resolve_path(path_str, root)
    try:
        return p.resolve(strict=True)
    except FileNotFoundError:
        return p

def ensure_pdftotext(path_str: str) -> Path:
    # Add .exe if omitted
    if not path_str.lower().endswith(".exe"):
        candidate = path_str + ".exe"
    else:
        candidate = path_str

    exe_path = Path(candidate)
    if exe_path.is_absolute():
        if exe_path.exists():
            return exe_path
        raise FileNotFoundError(f"pdftotext not found at: {exe_path}")
    # Search PATH
    found = shutil.which(candidate) or shutil.which(path_str)
    if not found:
        raise FileNotFoundError(
            "pdftotext not found on PATH and no valid absolute path provided. "
            "Install Poppler/Xpdf or pass --pdftotext with a full path."
        )
    return Path(found)

def collect_pdfs(input_arg: str, workspace: Path, recurse: bool) -> list[Path]:
    resolved = resolve_existing(input_arg, workspace)

    files: list[Path] = []

    # If input is a directory, collect *.pdf from it
    if resolved.exists() and resolved.is_dir():
        if recurse:
            for p in resolved.rglob("*.pdf"):
                if p.is_file():
                    files.append(p)
        else:
            for p in resolved.glob("*.pdf"):
                if p.is_file():
                    files.append(p)
        return files

    # If input is an existing file
    if resolved.exists() and resolved.is_file():
        if resolved.suffix.lower() != ".pdf":
            raise ValueError(f"Input file is not a PDF: {resolved}")
        return [resolved]

    # Otherwise, treat as wildcard (glob) relative to workspace
    # e.g., "Chase_history\*.pdf" (can include subfolders if user uses ** and sets recurse)
    glob_pattern = str(resolved)
    if recurse and "**" not in glob_pattern:
        # If user asked for recurse but did not provide **, expand into **/*.pdf from the parent dir
        parent = Path(glob_pattern).parent
        pattern = Path(parent) / "**" / Path(glob_pattern).name
        glob_pattern = str(pattern)

    for path_str in glob.glob(glob_pattern, recursive=recurse):
        p = Path(path_str)
        if p.is_file() and p.suffix.lower() == ".pdf":
            files.append(p)

    return files

def convert_one(pdftotext: Path, src_pdf: Path, out_txt: Path, layout: bool, nopgbrk: bool) -> None:
    args = [str(pdftotext)]
    if layout:
        args.append("-layout")
    if nopgbrk:
        args.append("-nopgbrk")
    args.extend(["-enc", "UTF-8"])
    args.extend([str(src_pdf), str(out_txt)])

    proc = subprocess.run(args, capture_output=True, text=True)
    if proc.returncode != 0:
        raise RuntimeError(
            f"pdftotext failed for {src_pdf} (code {proc.returncode}).\n"
            f"STDOUT: {proc.stdout}\nSTDERR: {proc.stderr}"
        )
    if not out_txt.exists():
        raise RuntimeError(f"pdftotext reported success but no output produced: {out_txt}")
    if out_txt.stat().st_size < 3:
        print(f"WARNING: Very small output ({out_txt.stat().st_size} bytes): {out_txt}")

def main(argv=None):
    parser = argparse.ArgumentParser(description="Convert PDF(s) to text with pdftotext.")
    parser.add_argument("--workspace", default=DEFAULT_WORKSPACE, help="Root folder (default: %(default)s)")
    parser.add_argument("--input",     default=DEFAULT_INPUT,     help="File/folder/wildcard (default: %(default)s)")
    parser.add_argument("--output",    default=DEFAULT_OUTPUT,    help="Output folder (default: %(default)s)")
    parser.add_argument("--suffix",    default=DEFAULT_SUFFIX,    help="Output suffix (default: %(default)s)")
    parser.add_argument("--pdftotext", default=DEFAULT_PDFTOTEXT, help="Path to pdftotext.exe or in PATH")
    parser.add_argument("--recurse",   action="store_true",       help="Recurse into subfolders (for folder or wildcard)")
    parser.add_argument("--overwrite", action="store_true",       help="Overwrite existing outputs")
    parser.add_argument("--no-layout", dest="layout", action="store_false", help="Disable -layout")
    parser.add_argument("--keep-pgbrk", dest="nopgbrk", action="store_false", help="Keep page breaks (disable -nopgbrk)")

    args = parser.parse_args(argv)

    ws = Path(args.workspace)
    if not ws.exists():
        print(f"ERROR: Workspace not found: {ws}", file=sys.stderr)
        return 2

    out_dir = resolve_path(args.output, ws)
    out_dir.mkdir(parents=True, exist_ok=True)

    try:
        pdftotext_path = ensure_pdftotext(args.pdftotext)
    except FileNotFoundError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 3

    pdfs = collect_pdfs(args.input, ws, args.recurse)
    if not pdfs:
        print(f"WARNING: No PDFs found for input: {args.input}")
        return 0

    converted = skipped = failed = 0
    for pdf in pdfs:
        try:
            base = pdf.stem
            out_txt = out_dir / f"{base}{args.suffix}"
            if out_txt.exists() and not args.overwrite:
                skipped += 1
                continue
            convert_one(pdftotext_path, pdf, out_txt, layout=args.layout, nopgbrk=args.nopgbrk)
            print(f"OK: {pdf.name} -> {out_txt.name}")
            converted += 1
        except Exception as ex:
            print(f"FAILED: {pdf}\n  {ex}", file=sys.stderr)
            failed += 1

    print(f"Done. Converted: {converted}  Skipped: {skipped}  Failed: {failed}")
    return 0 if failed == 0 else 1

if __name__ == "__main__":
    raise SystemExit(main())
