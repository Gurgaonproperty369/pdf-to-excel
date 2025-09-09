# PDF → Excel converter (Flask)

Small Flask app to convert PDF tables into Excel workbooks.

## Features
- Upload PDF via web UI
- Attempts extraction via Camelot → Tabula → pdfplumber
- Outputs a multi-sheet .xlsx (one sheet per table)
- Simple UI, easy to deploy

## Requirements & external dependencies
- Python 3.10+
- Java (required by tabula-py)
- Ghostscript & system libraries (required for Camelot):
  - On Debian/Ubuntu: `sudo apt install ghostscript python3-tk build-essential libpoppler-cpp-dev`
  - For Camelot, also install `tk` & `ghostscript`. Follow Camelot installation docs.
- pip packages: see `requirements.txt`

## Installation
1. Clone repo
2. (Optional) Create virtualenv:
