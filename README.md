# ELI Certificate Pipeline

Download calibration certificates from SharePoint (Microsoft Graph), separate the
real certificates from other documents, and extract their data into a
spreadsheet / JSON.

## Workflow
| Step | Script | What it does |
|------|--------|--------------|
| 1. Download | `download_eli_certs.py` | Pull files from a SharePoint folder via MS Graph (device-code sign-in) |
| 2. Classify & copy | `classify_and_copy.py` | Read every file; copy only real "Certificate of Calibration" docs into a clean folder |
| 2b. OCR (optional) | `ocr_scanned.py` | OCR image-only PDFs to recover scanned certs |
| 3–4. Extract | `extract_all.py` | Parse each cert into structured fields + standards → `.xlsx` / `.json` |

`extract_all.py` depends on the shared parsers **`eli_word_extractor.py`** (DOC/DOCX)
and **`eli_pdf_extractor.py`** (PDF) — keep all three in the same folder.

## Requirements
Python 3.10+:
    pip install msal requests pdfplumber PyMuPDF python-docx openpyxl pytesseract Pillow

External tools on your PATH:
- **antiword** — reads legacy `.doc` files
- **Tesseract OCR** — for scanned PDFs

## Configure
Set the constants near the top of the scripts before running:
- `download_eli_certs.py`: `TENANT_ID`, `CLIENT_ID`, `SHAREPOINT_HOST`, `SITE_PATH`, `FOLDER_PATH`
- `classify_and_copy.py` / `extract_all.py`: `ARCHIVE`, `DEST`, `SCRATCH`, `ANTIWORD` paths

## Usage
    python download_eli_certs.py     # downloads the archive (sign in when prompted)
    python classify_and_copy.py      # -> Cleaned Certs Setting/ + classification.json
    python ocr_scanned.py            # optional: recover scanned certs
    python extract_all.py            # -> eli_all_extract.{xlsx,json,jsonl}

## Notes
- Certificate files and extracted output are **not** included (proprietary + large).
- `.msal_token_cache.json` stores auth tokens — keep it git-ignored; never commit it.
