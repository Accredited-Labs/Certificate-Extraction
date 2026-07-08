"""
Classify every file under 'ELI CERTS 2024 - 2026' by reading its content/header,
then copy the real Certificates of Calibration into 'Cleaned Certs Setting'
(preserving each file's relative sub-path).

Readers (content-based, fast):
  .pdf            -> PyMuPDF (fitz) text of page 1   (no OCR; scanned -> "Scanned/Empty")
  .docx / .docm   -> unzip word/document.xml, strip tags
  .doc            -> antiword (copied to a short scratch path first)
  others          -> "Other (<ext>)"  (zip/xlsx/jpg/...)

Outputs:
  classification.json   full {path, ext, doc_type} list + summary
  copies of certs into 'Cleaned Certs Setting/<same subpath>'

Usage:
  python classify_and_copy.py --limit 300          # test on first 300
  python classify_and_copy.py                       # full run
  python classify_and_copy.py --no-copy             # classify only
  python classify_and_copy.py --use-cache           # reuse classification.json, just copy
"""
import os
import re
import io
import sys
import json
import shutil
import zipfile
import tempfile
import argparse
import subprocess
from collections import Counter
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor, as_completed

HERE = os.path.dirname(os.path.abspath(__file__))
ARCHIVE = os.path.join(HERE, "ELI CERTS 2024 - 2026")
DEST = os.path.join(HERE, "Cleaned Certs Setting")
CLASS_JSON = os.path.join(HERE, "classification.json")
SCRATCH = r"C:\Users\yangy\AppData\Local\Temp\claude\C--Users-yangy-Desktop-Accredited-Labs-Repo\cd1fa42c-733d-41a7-860a-e3950dfa67d3\scratchpad\classify"
ANTIWORD = r"C:\Program Files\Git\mingw64\bin\antiword.exe"

CERT_TYPE = "Certificate of Calibration"
OTHER_EXT = {".zip", ".xlsx", ".xls", ".xlsm", ".jpg", ".jpeg", ".png", ".lnk",
             ".lbx", ".dot", ".rtf", ".msg", ".eml"}


def L(p):
    p = os.path.abspath(p)
    if os.name == "nt" and not p.startswith("\\\\?\\"):
        return "\\\\?\\" + p
    return p


def classify_text(t):
    h = (t or "").lower()
    if "packing slip" in h:
        return "Packing Slip"
    if "calibration data report" in h:
        return "Calibration Data Report"
    if "certificate of calibration" in h:
        return "Certificate of Calibration"
    if re.search(r"on[\-\s]?site\s+sheet", h):
        return "On-Site Sheet"
    if not (t or "").strip():
        return "Scanned/Empty"
    return "Other/Unknown"


def _read_pdf(full):
    import fitz
    data = open(L(full), "rb").read()
    d = fitz.open(stream=data, filetype="pdf")
    try:
        txt = d[0].get_text("text") if d.page_count else ""
    finally:
        d.close()
    return txt[:2000]


def _read_docx(full):
    b = open(L(full), "rb").read()
    with zipfile.ZipFile(io.BytesIO(b)) as z:
        xml = z.read("word/document.xml").decode("utf-8", "ignore")
    return re.sub(r"<[^>]+>", " ", xml)[:3000]


def _read_doc(full):
    fd, tmp = tempfile.mkstemp(suffix=".doc", dir=SCRATCH)
    os.close(fd)
    try:
        shutil.copy2(L(full), L(tmp))
        r = subprocess.run([ANTIWORD, tmp], capture_output=True, timeout=90)
        return r.stdout.decode("utf-8", "ignore")[:2000]
    finally:
        try:
            os.remove(tmp)
        except Exception:
            pass


def classify_one(full):
    """Top-level worker for ProcessPoolExecutor. Returns (full, ext, doc_type)."""
    ext = os.path.splitext(full)[1].lower()
    try:
        if ext == ".pdf":
            dt = classify_text(_read_pdf(full))
        elif ext in (".docx", ".docm"):
            dt = classify_text(_read_docx(full))
        elif ext == ".doc":
            dt = classify_text(_read_doc(full))
        elif ext in OTHER_EXT:
            dt = f"Other ({ext})"
        else:
            dt = f"Other ({ext or 'noext'})"
    except Exception as e:
        dt = f"ERROR: {type(e).__name__}"
    return full, ext, dt


def walk_archive():
    files = []
    for root, _, fns in os.walk(L(ARCHIVE)):
        for fn in fns:
            files.append(os.path.join(root, fn))
    files.sort()
    return files


def relpath(full):
    return os.path.relpath(full, L(ARCHIVE))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--workers", type=int, default=10)
    ap.add_argument("--copy-workers", type=int, default=16)
    ap.add_argument("--no-copy", action="store_true")
    ap.add_argument("--use-cache", action="store_true")
    args = ap.parse_args()

    os.makedirs(SCRATCH, exist_ok=True)

    # -------- classify (or load cache) --------
    if args.use_cache and os.path.exists(CLASS_JSON):
        results = json.load(open(CLASS_JSON, encoding="utf-8"))["files"]
        print(f"Loaded {len(results)} classifications from cache.", flush=True)
    else:
        files = walk_archive()
        if args.limit:
            files = files[: args.limit]
        total = len(files)
        print(f"Classifying {total} files with {args.workers} workers...", flush=True)
        results = []
        done = 0
        with ProcessPoolExecutor(max_workers=args.workers) as ex:
            futs = {ex.submit(classify_one, f): f for f in files}
            for fut in as_completed(futs):
                full, ext, dt = fut.result()
                results.append({"file": relpath(full), "ext": ext, "doc_type": dt})
                done += 1
                if done % 1000 == 0 or done == total:
                    print(f"  classified {done}/{total}", flush=True)
        results.sort(key=lambda r: r["file"])
        by_type = Counter(r["doc_type"] for r in results)
        with open(CLASS_JSON, "w", encoding="utf-8") as f:
            json.dump({"total": len(results), "by_type": dict(by_type), "files": results},
                      f, indent=2)
        print(f"\nSaved classification.json", flush=True)

    by_type = Counter(r["doc_type"] for r in results)
    print("\n" + "=" * 56 + "\nCLASSIFICATION SUMMARY\n" + "=" * 56, flush=True)
    for t, c in by_type.most_common():
        print(f"  {t:<30} {c}", flush=True)

    if args.no_copy:
        return

    # -------- copy certs --------
    certs = [r for r in results if r["doc_type"] == CERT_TYPE]
    print(f"\nCopying {len(certs)} '{CERT_TYPE}' files into {DEST} ...", flush=True)
    os.makedirs(L(DEST), exist_ok=True)

    def copy_one(rel):
        src = os.path.join(ARCHIVE, rel)
        dst = os.path.join(DEST, rel)
        try:
            if os.path.exists(L(dst)) and os.path.getsize(L(dst)) == os.path.getsize(L(src)):
                return "skip"
            os.makedirs(L(os.path.dirname(dst)), exist_ok=True)
            shutil.copy2(L(src), L(dst))
            return "ok"
        except Exception as e:
            return f"err: {e}"

    ok = skip = err = 0
    done = 0
    with ThreadPoolExecutor(max_workers=args.copy_workers) as ex:
        futs = [ex.submit(copy_one, r["file"]) for r in certs]
        for fut in as_completed(futs):
            s = fut.result()
            if s == "ok":
                ok += 1
            elif s == "skip":
                skip += 1
            else:
                err += 1
                if err <= 15:
                    print("  " + s, flush=True)
            done += 1
            if done % 2000 == 0 or done == len(certs):
                print(f"  copied {done}/{len(certs)}  (new={ok} skip={skip} err={err})", flush=True)

    print("\n" + "=" * 56 + "\nDONE\n" + "=" * 56, flush=True)
    print(f"  certificates copied (new): {ok}", flush=True)
    print(f"  already present:           {skip}", flush=True)
    print(f"  errors:                    {err}", flush=True)
    print(f"  destination:               {DEST}", flush=True)


if __name__ == "__main__":
    main()
