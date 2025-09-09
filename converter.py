# converter.py
import os
import tempfile
import pandas as pd
import pdfplumber

# Try Camelot and Tabula only if installed
try:
    import camelot
    HAS_CAMELOT = True
except Exception:
    HAS_CAMELOT = False

try:
    import tabula
    HAS_TABULA = True
except Exception:
    HAS_TABULA = False


def _write_dfs_to_excel(dfs, excel_path):
    """Write list-of-DataFrames to an Excel workbook, each DF on its own sheet."""
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        for i, df in enumerate(dfs):
            sheet_name = f"Table_{i+1}"
            # ensure sheet name length <= 31 and safe
            sheet_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return excel_path


def extract_with_camelot(pdf_path, pages="1-end", flavor="lattice"):
    if not HAS_CAMELOT:
        raise RuntimeError("Camelot not available in environment.")
    # flavor: 'lattice' or 'stream'
    tables = camelot.read_pdf(pdf_path, pages=pages, flavor=flavor)
    dfs = [t.df for t in tables]
    return dfs


def extract_with_tabula(pdf_path, pages="all", multiple_tables=True):
    if not HAS_TABULA:
        raise RuntimeError("Tabula not available in environment.")
    # tabula returns list of dfs when multiple_tables True
    dfs = tabula.read_pdf(pdf_path, pages=pages, multiple_tables=multiple_tables)
    return dfs


def extract_with_pdfplumber(pdf_path, pages=None):
    """pdfplumber extraction: tries to extract tables across pages."""
    dfs = []
    with pdfplumber.open(pdf_path) as pdf:
        page_count = len(pdf.pages)
        page_range = pages or range(1, page_count + 1)
        # if pages passed as string like "1-3" or "1,3" handle outside
        # Here assume pages is an iterable of ints
        for p in page_range:
            try:
                page = pdf.pages[p - 1]
            except IndexError:
                continue
            # pdfplumber's extract_tables returns list of tables (list of lists)
            page_tables = page.extract_tables()
            for table in page_tables:
                # convert to pandas df
                try:
                    df = pd.DataFrame(table)
                    # try to detect header row: if first row has unique values, treat it as header
                    # simple heuristic: no duplicate in first row and no all-none
                    first_row = df.iloc[0].tolist()
                    if len(set(first_row)) == len(first_row):
                        df.columns = first_row
                        df = df.drop(df.index[0]).reset_index(drop=True)
                    dfs.append(df)
                except Exception:
                    continue
    return dfs


def parse_pages_arg(pages_arg, page_count=None):
    """Support pages param like 'all', '1-3', '1,3,5' or integer."""
    if not pages_arg or pages_arg in ("all", "All", "ALL"):
        return None
    pages_arg = str(pages_arg).strip()
    if "-" in pages_arg:
        start, end = pages_arg.split("-", 1)
        return range(int(start), int(end) + 1)
    if "," in pages_arg:
        parts = [int(x.strip()) for x in pages_arg.split(",") if x.strip().isdigit()]
        return parts
    if pages_arg.isdigit():
        return [int(pages_arg)]
    return None


def pdf_to_excel(pdf_path, excel_out_path, method="auto", pages="all", camelot_flavor="lattice"):
    """
    Convert pdf to excel, try methods in order:
     - camelot (if method=='camelot' or method=='auto' and available)
     - tabula (if available)
     - pdfplumber (fallback)
    """
    pages_arg = parse_pages_arg(pages)
    dfs = []
    errors = []

    if method in ("auto", "camelot") and HAS_CAMELOT:
        try:
            pages_str = pages if isinstance(pages, str) else "1-end"
            dfs = extract_with_camelot(pdf_path, pages=pages_str, flavor=camelot_flavor)
            if dfs:
                return _write_dfs_to_excel(dfs, excel_out_path), "camelot"
        except Exception as e:
            errors.append(f"Camelot error: {e}")

    if method in ("auto", "tabula") and HAS_TABULA:
        try:
            pages_str = pages if isinstance(pages, str) else "all"
            dfs = extract_with_tabula(pdf_path, pages=pages_str)
            if dfs:
                return _write_dfs_to_excel(dfs, excel_out_path), "tabula"
        except Exception as e:
            errors.append(f"Tabula error: {e}")

    # fallback to pdfplumber
    try:
        pages_iter = pages_arg
        dfs = extract_with_pdfplumber(pdf_path, pages=pages_iter)
        if dfs:
            return _write_dfs_to_excel(dfs, excel_out_path), "pdfplumber"
        else:
            # If no tables detected, try a fallback: try to extract text and put it into a sheet
            with open(pdf_path, "rb") as f:
                pass
            # simple fallback: create df with the whole text
            text_rows = []
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        for line in text.split("\n"):
                            text_rows.append([line])
            if text_rows:
                df = pd.DataFrame(text_rows, columns=["Extracted text"])
                return _write_dfs_to_excel([df], excel_out_path), "text_fallback"
    except Exception as e:
        errors.append(f"pdfplumber error: {e}")

    # If everything failed
    raise RuntimeError("No tables found or extraction failed. Errors: " + " | ".join(errors))

