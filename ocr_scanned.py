"""OCR the 'Scanned/Empty' PDFs, re-classify, and copy any that are
'Certificate of Calibration' into Cleaned Certs Setting. Reports all results."""
import os
import json
import shutil
from collections import Counter

import classify_and_copy as C

HERE = os.path.dirname(os.path.abspath(__file__))
data = json.load(open(os.path.join(HERE, "classification.json"), encoding="utf-8"))["files"]
scanned = [f for f in data if f["doc_type"] == "Scanned/Empty"]


def ocr_pdf(full):
    import fitz
    import pytesseract
    from PIL import Image
    raw = open(C.L(full), "rb").read()
    d = fitz.open(stream=raw, filetype="pdf")
    out = []
    try:
        for page in d:
            pix = page.get_pixmap(dpi=300)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            out.append(pytesseract.image_to_string(img))
    finally:
        d.close()
    return "\n".join(out)


results = []
for i, f in enumerate(scanned, 1):
    full = os.path.join(C.ARCHIVE, f["file"])
    try:
        if f["ext"] == ".pdf":
            txt = ocr_pdf(full)
            dt = C.classify_text(txt)
            method = "ocr"
        else:
            # .doc/.docx that returned empty -> try Word COM as a last resort
            dt = "OCR-NA (" + f["ext"] + ")"
            method = "skip"
    except Exception as e:
        dt = f"OCR-ERR: {type(e).__name__}"
        method = "err"
    results.append({**f, "ocr_type": dt, "method": method})
    if i % 10 == 0 or i == len(scanned):
        print(f"  ocr {i}/{len(scanned)}", flush=True)

print("\nOCR re-classification:", dict(Counter(r["ocr_type"] for r in results)))

# copy the ones that are now Certificate of Calibration
certs = [r for r in results if r["ocr_type"] == "Certificate of Calibration"]
copied = 0
for r in certs:
    src = os.path.join(C.ARCHIVE, r["file"])
    dst = os.path.join(C.DEST, r["file"])
    os.makedirs(C.L(os.path.dirname(dst)), exist_ok=True)
    if not (os.path.exists(C.L(dst)) and os.path.getsize(C.L(dst)) == os.path.getsize(C.L(src))):
        shutil.copy2(C.L(src), C.L(dst))
        copied += 1
print(f"\nScanned PDFs that are Certificate of Calibration: {len(certs)}  (copied {copied})")

with open(os.path.join(HERE, "ocr_scanned_results.json"), "w", encoding="utf-8") as fh:
    json.dump(results, fh, indent=2)
print("Wrote ocr_scanned_results.json")
