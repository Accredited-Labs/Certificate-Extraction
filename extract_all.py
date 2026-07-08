"""
Extract structured fields from every Certificate of Calibration in
'Cleaned Certs Setting' (all formats), reusing the validated field logic.

Readers (scalable, parallel):
  .pdf          pdfplumber text (+ PyMuPDF/Tesseract OCR fallback)
  .doc          antiword -w 0  (pipes/[pic] stripped)
  .docx/.docm   python-docx

All paths feed the same emit_tokens -> extract_fields pipeline (validated:
antiword .doc matches the Word-COM output field-for-field).

Outputs (resumable):
  eli_all_extract.jsonl   one JSON row per cert, appended as completed
  eli_all_extract.json    consolidated
  eli_all_extract.xlsx    spreadsheet

Usage:
  python extract_all.py --limit 50        # test
  python extract_all.py                    # full run
  python extract_all.py --resume           # continue from jsonl
"""
import os
import re
import io
import json
import html
import shutil
import zipfile
import tempfile
import argparse
import subprocess
from collections import Counter
from concurrent.futures import ProcessPoolExecutor, as_completed

import pdfplumber
from eli_word_extractor import (emit_tokens, build_tokens, extract_fields,
                                extract_standards, FIELDS, clean_val, L, collapse)
from eli_pdf_extractor import fix_spacing, extract_address, extract_standards_any

HERE = os.path.dirname(os.path.abspath(__file__))
DEST = os.path.join(HERE, "Cleaned Certs Setting")
SCRATCH = r"C:\Users\yangy\AppData\Local\Temp\claude\C--Users-yangy-Desktop-Accredited-Labs-Repo\cd1fa42c-733d-41a7-860a-e3950dfa67d3\scratchpad\extract"
ANTIWORD = r"C:\Program Files\Git\mingw64\bin\antiword.exe"
OUT_JSONL = os.path.join(HERE, "eli_all_extract.jsonl")
OUT_JSON = os.path.join(HERE, "eli_all_extract.json")
OUT_XLSX = os.path.join(HERE, "eli_all_extract.xlsx")
CID_RE = re.compile(r"\(cid:\d+\)")


def _pdf_text(data):
    out = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for p in pdf.pages:
            out.append(p.extract_text() or "")
    t = "\n".join(out)
    if len(t.strip()) >= 120:
        return CID_RE.sub("", t), "pdf-text"
    import fitz
    import pytesseract
    from PIL import Image
    d = fitz.open(stream=data, filetype="pdf")
    o = []
    try:
        for page in d:
            pix = page.get_pixmap(dpi=300)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            o.append(pytesseract.image_to_string(img))
    finally:
        d.close()
    return CID_RE.sub("", "\n".join(o)), "pdf-ocr"


def _docx(data):
    import docx
    doc = docx.Document(io.BytesIO(data))
    paras = [p.text for p in doc.paragraphs]
    tables = [[[c.text for c in row.cells] for row in t.rows] for t in doc.tables]
    return paras, tables


def _docx_xml_text(data):
    """Text from a .docx/.docm via raw XML (works for macro-enabled .docm that
    python-docx rejects). Paragraph -> newline, table cell -> tab."""
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        xml = z.read("word/document.xml").decode("utf-8", "ignore")
    xml = re.sub(r"</w:p>", "\n", xml)
    xml = re.sub(r"</w:tc>", "\t", xml)
    xml = re.sub(r"<w:tab[ /]", "\t", xml)
    return html.unescape(re.sub(r"<[^>]+>", "", xml))


def _doc_text(full):
    fd, tmp = tempfile.mkstemp(suffix=".doc", dir=SCRATCH)
    os.close(fd)
    try:
        shutil.copy2(L(full), L(tmp))
        r = subprocess.run([ANTIWORD, "-w", "0", tmp], capture_output=True, timeout=120)
        return r.stdout.decode("utf-8", "ignore").replace("|", " ").replace("[pic]", " ")
    finally:
        try:
            os.remove(tmp)
        except Exception:
            pass


def _tokens_from_text(text):
    toks = []
    for ln in text.split("\n"):
        toks.extend(emit_tokens(ln))
    return toks


def extract_file(rel):
    full = os.path.join(DEST, rel)
    ext = os.path.splitext(rel)[1].lower()
    row = {f: "" for f in FIELDS}
    row.update({"file": rel, "source_format": ext.lstrip("."), "num_standards": 0,
                "standards": [], "error": ""})
    try:
        if ext == ".pdf":
            data = open(L(full), "rb").read()
            text, src = _pdf_text(data)
            text = fix_spacing(text)
            f = extract_fields(_tokens_from_text(text))
            addr = extract_address(text)
            if addr:
                f["customer_address"] = addr
            stds = extract_standards_any(text)
            row["source_format"] = src
        elif ext == ".docx":
            data = open(L(full), "rb").read()
            paras, tables = _docx(data)
            tokens = build_tokens(paras, tables)
            f = extract_fields(tokens)
            # text for standards: paragraphs + collapsed (de-duped) table cells, one per line
            parts = list(paras)
            for t in tables:
                for r in t:
                    parts.extend(collapse([c.replace("\n", " ").strip() for c in r]))
            stds = extract_standards_any("\n".join(parts))
        elif ext == ".docm":
            data = open(L(full), "rb").read()
            text = fix_spacing(_docx_xml_text(data))
            f = extract_fields(_tokens_from_text(text))
            addr = extract_address(text)
            if addr:
                f["customer_address"] = addr
            stds = extract_standards_any(text)
        elif ext == ".doc":
            text = fix_spacing(_doc_text(full))
            f = extract_fields(_tokens_from_text(text))
            addr = extract_address(text)
            if addr:
                f["customer_address"] = addr
            stds = extract_standards_any(text)
        else:
            row["error"] = "unsupported ext"
            return row
        for k in FIELDS:
            row[k] = f.get(k, "")
        row["standards"] = stds
        row["num_standards"] = len(stds)
    except Exception as e:
        row["error"] = f"{type(e).__name__}: {e}"
    return row


def walk_dest():
    rels = []
    for root, _, fns in os.walk(L(DEST)):
        for fn in fns:
            if os.path.splitext(fn)[1].lower() in (".pdf", ".doc", ".docx", ".docm"):
                full = os.path.join(root, fn)
                rels.append(os.path.relpath(full, L(DEST)))
    rels.sort()
    return rels


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--workers", type=int, default=10)
    ap.add_argument("--resume", action="store_true")
    args = ap.parse_args()

    os.makedirs(SCRATCH, exist_ok=True)
    rels = walk_dest()
    if args.limit:
        rels = rels[: args.limit]

    done = set()
    if args.resume and os.path.exists(OUT_JSONL):
        for line in open(OUT_JSONL, encoding="utf-8"):
            try:
                done.add(json.loads(line)["file"])
            except Exception:
                pass
        print(f"Resume: {len(done)} already done.", flush=True)
    todo = [r for r in rels if r not in done]
    print(f"Extracting {len(todo)} / {len(rels)} certs with {args.workers} workers...", flush=True)

    results = []
    if args.resume and os.path.exists(OUT_JSONL):
        results = [json.loads(l) for l in open(OUT_JSONL, encoding="utf-8")]

    mode = "a" if (args.resume and os.path.exists(OUT_JSONL)) else "w"
    n = 0
    with open(OUT_JSONL, mode, encoding="utf-8") as jl, ProcessPoolExecutor(max_workers=args.workers) as ex:
        futs = {ex.submit(extract_file, r): r for r in todo}
        for fut in as_completed(futs):
            row = fut.result()
            jl.write(json.dumps(row, ensure_ascii=False) + "\n")
            jl.flush()
            results.append(row)
            n += 1
            if n % 1000 == 0 or n == len(todo):
                print(f"  extracted {n}/{len(todo)}", flush=True)

    # consolidate
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)

    # xlsx
    import openpyxl
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("certs")
    cols = ["file", "source_format"] + FIELDS + ["num_standards", "standards", "error"]
    ws.append(cols)
    for r in results:
        out = []
        for c in cols:
            if c == "standards":
                v = "; ".join(f"{s['std']}:{s['description']}({s['date_due']})" for s in r.get("standards", []))
            else:
                v = r.get(c, "")
            out.append(clean_val(v) if isinstance(v, str) else v)
        ws.append(out)
    wb.save(OUT_XLSX)

    # summary
    errs = [r for r in results if r.get("error")]
    byfmt = Counter(r["source_format"] for r in results)
    print("\n" + "=" * 56 + "\nEXTRACTION SUMMARY\n" + "=" * 56, flush=True)
    print(f"  total rows: {len(results)}   errors: {len(errs)}", flush=True)
    print(f"  by source: {dict(byfmt)}", flush=True)
    keyflds = ["customer_name", "cert_number", "date_cal", "due_cal", "manufacturer",
               "model", "nomenclature", "serial_number", "procedure", "certified_by", "num_standards"]
    print("  fill rates:", flush=True)
    for fld in keyflds:
        filled = sum(1 for r in results if r.get(fld))
        print(f"    {fld:16} {filled}/{len(results)}  ({100*filled//max(1,len(results))}%)", flush=True)
    print(f"\n  wrote {OUT_JSON} , {OUT_XLSX} , {OUT_JSONL}", flush=True)


if __name__ == "__main__":
    main()
