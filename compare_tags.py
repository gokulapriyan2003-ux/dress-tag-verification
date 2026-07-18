"""
Dress Tag vs Master Sheet Verifier
-----------------------------------
Extracts per-SKU fields from a multi-tag PDF (dress/garment tags laid out
N-per-row) and compares them against a master Excel sheet, producing a
match/mismatch report.

Usage:
    python compare_tags.py <tag_pdf> <master_xlsx> [--sheet SHEET_NAME] [--out report.xlsx]
"""

import os
import re
import sys
import argparse
import pdfplumber
import openpyxl
import pandas as pd
import urllib.request


# ---------------------------------------------------------------------------
# 1. PDF EXTRACTION
# ---------------------------------------------------------------------------

# Each of these labels appears once per tag, repeated N times per printed line
# (N = number of tags side-by-side in that row of the sheet). We split each
# line on the label text itself to recover the N individual values in order.
LABELS = [
    "Style:",
    "Product:",
    "Fit:",
    "Color:",
    "Category:",
    "Manufactured On:",
    "Net Quantity:",
    "SKU Code:",
    "SIZE :",
    "MRP:",
    "Qty:",
]

BARCODE_RE = re.compile(r"^\d{8,14}$")          # standalone barcode line
CM_RE = re.compile(r"^\(\d+(\.\d+)?CM\)$")       # e.g. (71.12CM)


def split_repeated_label(line: str, label: str):
    """Split a line like 'Label: A Label: B Label: C' into ['A','B','C']."""
    parts = line.split(label)
    parts = [p.strip() for p in parts if p.strip() != ""]
    return parts


def extract_pdf_tags(pdf_path: str) -> pd.DataFrame:
    field_lists = {lbl: [] for lbl in LABELS}
    barcodes = []
    cm_sizes = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for raw_line in text.split("\n"):
                line = raw_line.strip()
                if not line:
                    continue

                matched_label = None
                for lbl in LABELS:
                    if line.startswith(lbl):
                        matched_label = lbl
                        break

                if matched_label:
                    field_lists[matched_label].extend(
                        split_repeated_label(line, matched_label)
                    )
                    continue

                # barcode / cm-size lines contain several space-separated tokens
                tokens = line.split()
                if all(BARCODE_RE.match(t) for t in tokens) and tokens:
                    barcodes.extend(tokens)
                elif all(CM_RE.match(t) for t in tokens) and tokens:
                    cm_sizes.extend(tokens)

    counts = {k: len(v) for k, v in field_lists.items()}
    counts["Barcode"] = len(barcodes)
    counts["CM"] = len(cm_sizes)
    n_tags = counts.get("SKU Code:")

    # Sanity check: every field should appear exactly once per tag.
    mismatched = {k: v for k, v in counts.items() if v != n_tags}
    if mismatched:
        print("WARNING: field counts don't all line up 1:1 with tag count "
              f"({n_tags}). Counts: {counts}", file=sys.stderr)

    rows = []
    for i in range(n_tags):
        def get(lbl):
            lst = field_lists[lbl]
            return lst[i] if i < len(lst) else None

        mrp_raw = get("MRP:")
        mrp_val = None
        if mrp_raw:
            m = re.search(r"[\d,]+\.?\d*", mrp_raw.replace("₹", ""))
            if m:
                mrp_val = float(m.group().replace(",", ""))

        qty_raw = get("Qty:")
        qty_val = int(qty_raw) if qty_raw and qty_raw.isdigit() else qty_raw

        rows.append({
            "Style": get("Style:"),
            "Product": get("Product:"),
            "Fit": get("Fit:"),
            "Color": get("Color:"),
            "Category": get("Category:"),
            "Manufactured On": get("Manufactured On:"),
            "Net Quantity": get("Net Quantity:"),
            "SKU": get("SKU Code:"),
            "Size": get("SIZE :"),
            "Size(CM)": cm_sizes[i] if i < len(cm_sizes) else None,
            "Barcode": barcodes[i] if i < len(barcodes) else None,
            "MRP": mrp_val,
            "Qty": qty_val,
        })

    return pd.DataFrame(rows)


# ---------------------------------------------------------------------------
# 2. EXCEL MASTER SHEET EXTRACTION
# ---------------------------------------------------------------------------

def extract_excel_master(xlsx_path: str, sheet_name: str = None) -> pd.DataFrame:
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    sheet_name = sheet_name or wb.sheetnames[0]
    ws = wb[sheet_name]

    rows = [list(r) for r in ws.iter_rows(values_only=True)]

    # Find the header row: the row that contains something like "SKU"
    header_idx = None
    for i, row in enumerate(rows):
        cells = [str(c).strip().upper() if c else "" for c in row]
        if any("SKU" in c for c in cells):
            header_idx = i
            break

    if header_idx is None:
        raise ValueError("Could not find a header row containing 'SKU' in the sheet.")

    header = [str(c).strip() if c else "" for c in rows[header_idx]]
    data_rows = []
    for row in rows[header_idx + 1:]:
        # stop at a blank row or a "TOTAL" row
        first_cell = str(row[0]).strip().upper() if row[0] else ""
        if first_cell == "" and all(c is None for c in row):
            continue
        if first_cell == "TOTAL":
            break
        if row[0] is None:
            continue
        data_rows.append(row)

    df = pd.DataFrame(data_rows, columns=header[:len(data_rows[0])] if data_rows else header)
    # Trim to only named columns
    df = df[[c for c in header if c]]
    return df


# ---------------------------------------------------------------------------
# 3. COMPARISON
# ---------------------------------------------------------------------------

def normalize_sku(x):
    if x is None:
        return ""
    return str(x).strip().upper()


def normalize_text(x):
    if x is None:
        return ""
    return str(x).strip().upper()


def normalize_size(x):
    if x is None:
        return ""
    s = str(x).strip().upper()

    if "/" in s:
        s = s.split("/")[0].strip()

    if s.endswith("UK"):
        s = s[:-2].strip()

    if (s.startswith("K") or s.startswith("Y")) and len(s) >= 2 and s[1:].isdigit():
        s = s[1:]

    if s.isdigit():
        s = str(int(s))

    if len(s) == 3 and s[0].isalpha() and s[1:].isdigit():
        s = str(int(s[1:]))

    size_map = {
        "SML": "S",
        "SMALL": "S",
        "MED": "M",
        "MEDIUM": "M",
        "LAR": "L",
        "LARGE": "L",
        "XLR": "XL",
        "EXTRA LARGE": "XL",
        "2XLR": "2XL",
        "XXL": "2XL",
        "2XL": "2XL",
        "3XLR": "3XL",
        "XXXL": "3XL",
        "3XL": "3XL",
        "4XLR": "4XL",
        "XXXXL": "4XL",
        "4XL": "4XL",
        "5XLR": "5XL",
        "5XL": "5XL",
    }
    return size_map.get(s, s)


def format_size_as_tag(size_str):
    if not size_str:
        return size_str
    s = str(size_str).strip().upper()

    if (s.startswith("K") or s.startswith("Y")) and len(s) >= 2 and s[1:].isdigit():
        digits_str = s[1:]
        if len(digits_str) == 1:
            digits_str = "0" + digits_str
        return digits_str + "UK"

    if len(s) == 3 and s[0].isalpha() and s[1:].isdigit():
        return str(int(s[1:]))

    size_map = {
        "SML": "S",
        "MED": "M",
        "LAR": "L",
        "XLR": "XL",
        "2XLR": "2XL",
        "XXL": "2XL",
        "3XLR": "3XL",
        "XXXL": "3XL",
        "4XLR": "4XL",
        "XXXXL": "4XL",
        "5XLR": "5XL",
    }
    return size_map.get(s, s)


color_map = {
    "OGY": "OYSTER GRAY", "MTG": "MINT GREEN", "CNG": "CHARCOAL GRAY",
    "VIO": "VINTAGE INDIGO", "BLK": "BLACK", "NVY": "NAVY", "WHT": "WHITE",
    "BRD": "BERRY RED", "WLG": "WOODLAND GRAY", "ASH": "ASH", "YLW": "YELLOW",
    "CFK": "CHILI FLAKES", "BTE": "BEETLE", "IRG": "IRON GRAY", "ASC": "ASSORTED",
    "IGM": "IRON GRAY MELANGE", "TRN": "TREKKING GREEN", "DPK": "DUSTY PINK",
    "MYE": "MYRTLE", "BLE": "BLUE STONE", "GNB": "GREEN BOG", "NTC": "NORTH ATLANTIC",
    "GSE": "GREEN SMOKE", "SYM": "SMOKEY OLIVE", "CHA": "CHOCOLATE",
    "CTM": "CHOCOLATE TRUFFLE", "NYB": "NAVY BLUE", "LIB": "LIGHT BLUE",
    "SKO": "SMOKEY OLIVE", "OFW": "OFF WHITE", "OBR": "OX BLOOD RED",
    "DDC": "DUSTY DEEP CHARCOAL", "CBL": "CARBON BLACK", "GFT": "GULF COAST",
    "IRON GREY": "IRON GRAY"
}


def normalize_color(x):
    if not x:
        return ""
    c = str(x).strip().upper().replace("GREY", "GRAY")
    for suffix in [" PRO", " PLUS", " PREMIUM"]:
        if c.endswith(suffix):
            c = c[:-len(suffix)].strip()
    return color_map.get(c, c)


def normalize_number(x):
    if x is None or x == "":
        return None
    try:
        return round(float(str(x).replace(",", "").replace("₹", "").strip()), 2)
    except ValueError:
        return str(x).strip().upper()


def find_col(df, *candidates):
    """Find a column in df whose name loosely matches one of the candidates."""
    cols_upper = {str(c).upper().strip(): c for c in df.columns}
    for cand in candidates:
        cand_u = cand.upper().strip()
        if cand_u in cols_upper:
            return cols_upper[cand_u]
    for cand in candidates:
        cand_u = cand.upper().strip()
        for cu, orig in cols_upper.items():
            if cand_u in cu:
                # Exclude columns containing dates/locations for MRP
                if cand_u == "MRP" and any(x in cu for x in ["DATE", "LOCATION", "ACTIVE"]):
                    continue
                return orig
    # Fallback to match anything if no clean match found
    for cand in candidates:
        cand_u = cand.upper().strip()
        for cu, orig in cols_upper.items():
            if cand_u in cu:
                return orig
    return None


def get_updated_mrp(pdf_style, pdf_sku, gsheet_dfs):
    if not gsheet_dfs:
        return None

    style_clean = str(pdf_style).strip().upper() if pdf_style else ""
    sku_clean = str(pdf_sku).strip().upper() if pdf_sku else ""

    # 1. Search in DT FINAL MRP (matching against Column H (8th column, index 7))
    df_dt = gsheet_dfs.get("DT FINAL MRP")
    if df_dt is not None and len(df_dt.columns) > 7:
        col_h = df_dt.columns[7]
        # Match exact style
        match = df_dt[df_dt[col_h].astype(str).str.strip().str.upper() == style_clean]
        if not match.empty:
            mrp_val = match.iloc[0].get("MRP")
            if pd.notna(mrp_val):
                return mrp_val

        # Fallback: base style match (W209/01 -> W209)
        if "/" in style_clean:
            base_style = style_clean.split("/")[0]
            match = df_dt[df_dt[col_h].astype(str).str.strip().str.upper() == base_style]
            if not match.empty:
                mrp_val = match.iloc[0].get("MRP")
                if pd.notna(mrp_val):
                    return mrp_val

    # 2. Search in New MRP 26-27 (matching against Column I (9th column, index 8))
    df_new = gsheet_dfs.get("New MRP 26-27")
    if df_new is not None and len(df_new.columns) > 8:
        col_i = df_new.columns[8]
        # Match exact style
        match = df_new[df_new[col_i].astype(str).str.strip().str.upper() == style_clean]
        if not match.empty:
            mrp_val = match.iloc[0].get("MRP")
            if pd.notna(mrp_val):
                return mrp_val

        # Fallback: base style match (W209/01 -> W209)
        if "/" in style_clean:
            base_style = style_clean.split("/")[0]
            match = df_new[df_new[col_i].astype(str).str.strip().str.upper() == base_style]
            if not match.empty:
                mrp_val = match.iloc[0].get("MRP")
                if pd.notna(mrp_val):
                    return mrp_val

    return None


def extract_sku_details(sku_str):
    sku = str(sku_str).strip().upper()
    n = len(sku)

    rules = {
        11: (2, 4, 0),
        12: (2, 4, 0),
        13: (2, 5, 0),
        14: (2, 4, 2),
        15: (2, 4, 3),
        16: (2, 5, 3),
        17: (2, 8, 0),
        18: (2, 4, 6) if sku.endswith(("2PK", "3PK")) else (3, 6, 3),
    }

    if n not in rules:
        return None, None, None

    if n == 12:
        body = sku[2:]
        if len(body) >= 5 and body[:2].isalpha() and body[2:5].isdigit():
            style_len = 5
        else:
            style_len = 4
        style = body[:style_len]
        if len(style) >= 3 and style[1:3] == "OR":
            style = style[1:]
        color = body[style_len:style_len+3]
        size = body[style_len+3:]
        return style, color, size

    if n == 15:
        apparel_sizes = {"MED", "LAR", "XLR", "2XL", "3XL", "4XL", "5XL", "SML"}
        if sku[-3:] in apparel_sizes:
            style = sku[2:-6]
            if len(style) >= 3 and style[1:3] == "OR":
                style = style[1:]
            color = sku[-6:-3]
            size = sku[-3:]
            return style, color, size

    start_remove, style_len, end_remove = rules[n]
    body = sku[start_remove:]
    if end_remove:
        body = body[:-end_remove]

    style = body[:style_len]
    if len(style) >= 3 and style[1:3] == "OR":
        style = style[1:]
    color = body[style_len:style_len+3]
    size = body[-3:]

    return style, color, size


def extract_style_and_size_from_sku(sku_str):
    style, color, size = extract_sku_details(sku_str)
    return style, size





def compare(pdf_df: pd.DataFrame, excel_df: pd.DataFrame, gsheet_dfs: dict) -> pd.DataFrame:
    sku_col = find_col(excel_df, "SKU CODE", "SKU")
    barcode_col = find_col(excel_df, "BARCODE", "BAR CODE", "EAN", "GTIN")
    mrp_col = find_col(excel_df, "MRP")
    size_col = find_col(excel_df, "SIZE")
    color_col = find_col(excel_df, "COLOUR", "COLOR")
    qty_col = find_col(excel_df, "TAG QTY", "QTY")

    if sku_col is None:
        raise ValueError("Could not find an SKU column in the Excel sheet.")

    excel_idx = {normalize_sku(row[sku_col]): row for _, row in excel_df.iterrows()}

    field_map = [
        ("SKU", sku_col, normalize_sku),
        ("Barcode", barcode_col, normalize_text),
        ("MRP", mrp_col, normalize_number),
        ("Size", size_col, normalize_size),
        ("Color", color_col, normalize_color),
        ("Qty", qty_col, normalize_number),
    ]

    report_rows = []
    matched_excel_skus = set()

    for _, tag in pdf_df.iterrows():
        pdf_sku_norm = normalize_sku(tag["SKU"])
        excel_row = excel_idx.get(pdf_sku_norm)

        if excel_row is None:
            report_rows.append({
                "SKU": tag["SKU"],
                "Field": "SKU",
                "PDF Value": tag["SKU"],
                "Excel Value": None,
                "Status": "❌ Not found in Excel",
            })
            continue

        matched_excel_skus.add(pdf_sku_norm)

        for field_name, excel_col, norm_fn in field_map:
            pdf_val = tag.get(field_name)

            if field_name == "MRP":
                excel_val = get_updated_mrp(tag.get("Style"), tag.get("SKU"), gsheet_dfs)
                if excel_val is None:
                    excel_val = excel_row.get(excel_col) if excel_col else None
            elif field_name == "Size":
                excel_sku = excel_row.get(sku_col)
                _, _, extracted_size = extract_sku_details(excel_sku)
                excel_val = format_size_as_tag(extracted_size) if extracted_size else None
            elif field_name == "Color":
                if excel_col and pd.notna(excel_row.get(excel_col)):
                    excel_val = excel_row.get(excel_col)
                else:
                    excel_sku = excel_row.get(sku_col)
                    _, extracted_color, _ = extract_sku_details(excel_sku)
                    excel_val = color_map.get(extracted_color, extracted_color) if extracted_color else None
            else:
                if excel_col is None:
                    continue
                excel_val = excel_row.get(excel_col)

            pdf_norm = norm_fn(pdf_val)
            excel_norm = norm_fn(excel_val)

            status = "✅ Match" if pdf_norm == excel_norm else "❌ Mismatch"
            report_rows.append({
                "SKU": tag["SKU"],
                "Field": field_name,
                "PDF Value": pdf_val,
                "Excel Value": excel_val,
                "Status": status,
            })

    return pd.DataFrame(report_rows)


def main():
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass

    ap = argparse.ArgumentParser(description="Compare dress-tag PDF against master Excel sheet")
    ap.add_argument("pdf", nargs="?", default=None)
    ap.add_argument("xlsx", nargs="?", default=None)
    ap.add_argument("--sheet", default=None)
    ap.add_argument("--out", default="tag_comparison_report.xlsx")
    args = ap.parse_args()

    script_dir = os.path.dirname(os.path.abspath(__file__))

    pdf_path = args.pdf
    if not pdf_path:
        pdfs = [f for f in os.listdir(script_dir) if f.lower().endswith(".pdf") and not f.startswith("~$")]
        if not pdfs:
            print("Error: No PDF file found in the script directory.", file=sys.stderr)
            sys.exit(1)
        pdf_path = os.path.join(script_dir, pdfs[0])
    elif not os.path.isabs(pdf_path):
        pdf_path = os.path.join(script_dir, pdf_path)

    xlsx_path = args.xlsx
    if not xlsx_path:
        xlsxs = [f for f in os.listdir(script_dir) if f.lower().endswith(".xlsx") and not f.startswith("~$") and not any(x in f.lower() for x in ["report", "google", "explore", "comparison"])]
        if not xlsxs:
            print("Error: No Excel (.xlsx) file found in the script directory.", file=sys.stderr)
            sys.exit(1)
        xlsx_path = os.path.join(script_dir, xlsxs[0])
    elif not os.path.isabs(xlsx_path):
        xlsx_path = os.path.join(script_dir, xlsx_path)

    out_path = args.out
    if not os.path.isabs(out_path):
        out_path = os.path.join(script_dir, out_path)

    print(f"Using PDF: {pdf_path}")
    print(f"Using Excel: {xlsx_path}")

    # Download updated MRP Google Sheet
    gsheet_dfs = {}
    gsheet_url = "https://docs.google.com/spreadsheets/d/1Q7nboN_Rezl807J0naA0QczTyoAQ6WM-KNmp_F26n5M/export?format=xlsx"
    gsheet_path = os.path.join(script_dir, "google_sheet_mrp.xlsx")
    print("\nDownloading updated MRP Google Sheet...")
    try:
        req = urllib.request.Request(gsheet_url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req, timeout=15) as response:
            with open(gsheet_path, "wb") as f:
                f.write(response.read())
        print("Download successful.")
        xls = pd.ExcelFile(gsheet_path)
        for sheet_name in xls.sheet_names:
            gsheet_dfs[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
    except Exception as e:
        print(f"Warning: Could not download/read updated MRP Google Sheet ({e}). Using local Excel MRP values as fallback.", file=sys.stderr)

    pdf_df = extract_pdf_tags(pdf_path)
    excel_df = extract_excel_master(xlsx_path, args.sheet)

    print(f"Extracted {len(pdf_df)} tags from PDF.")
    print(f"Extracted {len(excel_df)} rows from Excel.")

    report_df = compare(pdf_df, excel_df, gsheet_dfs)

    n_mismatch = (report_df["Status"] != "✅ Match").sum()
    n_total = len(report_df)
    print(f"\n{n_total - n_mismatch}/{n_total} field checks passed.")
    if n_mismatch:
        print(f"{n_mismatch} issues found (showing first 50, see full report in Excel):")
        print(report_df[report_df["Status"] != "✅ Match"].head(50).to_string(index=False))

    try:
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            pdf_df.to_excel(writer, sheet_name="PDF_Extracted", index=False)
            excel_df.to_excel(writer, sheet_name="Excel_Master", index=False)
            report_df.to_excel(writer, sheet_name="Comparison_Report", index=False)

        # Color the report sheet
        wb = openpyxl.load_workbook(out_path)
        ws = wb["Comparison_Report"]
        from openpyxl.styles import PatternFill
        green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        status_col_idx = report_df.columns.get_loc("Status") + 1
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=status_col_idx)
            fill = green if "Match" in str(cell.value) and "Mis" not in str(cell.value) and "Not found" not in str(cell.value) else red
            for col_idx in range(1, ws.max_column + 1):
                ws.cell(row=row_idx, column=col_idx).fill = fill
        wb.save(out_path)

        print(f"\nReport saved to {out_path}")
    except PermissionError:
        print(f"\nERROR: Permission denied when writing to '{out_path}'.\n"
              "Please make sure the file is closed in Microsoft Excel or other programs and run the script again.", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
