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
    "Lot No:",
    "Product:",
    "Fit:",
    "Color:",
    "Category:",
    "Manufactured On:",
    "MFD :",
    "Net Quantity:",
    "Net Qty:",
    "HSN Code:",
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
    descriptions = []
    total_mrps = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            raw_lines = [l.strip() for l in text.split("\n") if l.strip()]
            for idx_line, line in enumerate(raw_lines):
                matched_label = None
                for lbl in LABELS:
                    if line.startswith(lbl):
                        matched_label = lbl
                        break

                if matched_label:
                    if matched_label == "Lot No:" and idx_line > 0:
                        prev_line = raw_lines[idx_line - 1]
                        lots_count = len(split_repeated_label(line, "Lot No:"))
                        if "Lot No:" not in prev_line and not any(prev_line.startswith(x) for x in LABELS):
                            parts = [p.strip() for p in prev_line.split("  ") if p.strip()]
                            if len(parts) == lots_count:
                                descriptions.extend(parts)
                            else:
                                chunk_len = max(1, len(prev_line) // lots_count)
                                single_desc = prev_line[:chunk_len].strip()
                                descriptions.extend([single_desc] * lots_count)

                    field_lists[matched_label].extend(
                        split_repeated_label(line, matched_label)
                    )
                    continue

                # Check for Total MRP line
                tot_tokens = line.split()
                tot_mrps_line = []
                for t in tot_tokens:
                    m_tot = re.search(r"₹?\s*([\d,]+\.?\d*)\s*/-\s*\(\s*\d+\s*Nos?\s*\)", t, re.IGNORECASE)
                    if m_tot:
                        try:
                            tot_mrps_line.append(float(m_tot.group(1).replace(",", "")))
                        except ValueError:
                            pass
                if tot_mrps_line:
                    total_mrps.extend(tot_mrps_line)
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

        style_raw = get("Style:") or get("Lot No:")
        net_qty_raw = get("Net Quantity:") or get("Net Qty:")
        mfd_raw = get("Manufactured On:") or get("MFD :")
        sku_val = get("SKU Code:")
        desc_val = get("Product:") or (descriptions[i] if i < len(descriptions) else None)
        size_val = get("SIZE :")
        if not size_val and sku_val:
            _, _, ext_sz = extract_sku_details(sku_val)
            if ext_sz:
                size_val = format_size_as_tag(ext_sz)

        rows.append({
            "Style": style_raw,
            "Product": desc_val,
            "Description": desc_val,
            "Fit": get("Fit:"),
            "Color": get("Color:"),
            "Category": get("Category:"),
            "Manufactured On": mfd_raw,
            "Net Quantity": net_qty_raw,
            "SKU": sku_val,
            "Size": size_val,
            "Size(CM)": cm_sizes[i] if i < len(cm_sizes) else None,
            "Barcode": barcodes[i] if i < len(barcodes) else None,
            "MRP": mrp_val,
            "Total MRP": total_mrps[i] if i < len(total_mrps) else None,
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
    "ACL": "ACID LIME",
    "ACID LIME": "ACID LIME",
    "AQM": "AQUA MELANGE",
    "AQUA MELANGE": "AQUA MELANGE",
    "AQO": "AQUA OCEAN",
    "AQUA OCEAN": "AQUA OCEAN",
    "BPK": "BABY PINK",
    "BABY PINK": "BABY PINK",
    "BLF": "NO",
    "NO": "NO",
    "BFM": "BAY LEAF MELANGE",
    "BAY LEAF MELANGE": "BAY LEAF MELANGE",
    "BLK": "BLACK",
    "BLACK": "BLACK",
    "BLM": "BLACK MELANGE",
    "BLACK MELANGE": "BLACK MELANGE",
    "BCT": "BLACK CURRENT",
    "BLACK CURRENT": "BLACK CURRENT",
    "BCM": "BLACK CURRENT MELANGE",
    "BLACK CURRENT MELANGE": "BLACK CURRENT MELANGE",
    "BBV": "BLUE BERRY VIOLET",
    "BLUE BERRY VIOLET": "BLUE BERRY VIOLET",
    "BMB": "BOMBAY BLUE",
    "BOMBAY BLUE": "BOMBAY BLUE",
    "BEK": "BUBBLE PINK",
    "BUBBLE PINK": "BUBBLE PINK",
    "BUG": "BURGUNDY",
    "BURGUNDY": "BURGUNDY",
    "BGM": "BURGUNDY MELANGE",
    "BURGUNDY MELANGE": "BURGUNDY MELANGE",
    "CBN": "CARBON",
    "CARBON": "CARBON",
    "CBM": "CARBON MELANGE",
    "CARBON MELANGE": "CARBON MELANGE",
    "CYW": "CASTRO YELLOW",
    "CASTRO YELLOW": "CASTRO YELLOW",
    "CGY": "CITY GRAY",
    "CITY GREY": "CITY GRAY",
    "CITY GRAY": "CITY GRAY",
    "CRD": "CORAL RED",
    "CORAL RED": "CORAL RED",
    "CRM": "CORAL RED-MELANGE",
    "CORAL RED-MELANGE": "CORAL RED-MELANGE",
    "DPE": "DARK PURPLE",
    "DARK PURPLE": "DARK PURPLE",
    "DFT": "DEEP FOREST",
    "DEEP FOREST": "DEEP FOREST",
    "DFM": "DEEP FOREST MELANGE",
    "DEEP FOREST MELANGE": "DEEP FOREST MELANGE",
    "DNM": "DENIM",
    "DENIM": "DENIM",
    "DMM": "DENIM MELANGE",
    "DENIM MELANGE": "DENIM MELANGE",
    "DCN": "DARK CARBON",
    "DARK CARBON": "DARK CARBON",
    "DCM": "DARK CARBON MELANGE",
    "DARK CARBON MELANGE": "DARK CARBON MELANGE",
    "FWN": "FAWN",
    "FAWN": "FAWN",
    "FPE": "FUSHIA PURPLE",
    "FUSHIA PURPLE": "FUSHIA PURPLE",
    "HTG": "HUNTER GREEN",
    "HUNTER GREEN": "HUNTER GREEN",
    "HTM": "HUNTER GREEN MELANGE",
    "HUNTER GREEN MELANGE": "HUNTER GREEN MELANGE",
    "IGO": "INDIGO",
    "INDIGO": "INDIGO",
    "IRG": "IRON GRAY",
    "IRON GREY": "IRON GRAY",
    "IRON GRAY": "IRON GRAY",
    "LGN": "LAKE GREEN",
    "LAKE GREEN": "LAKE GREEN",
    "LLC": "LILAC",
    "LILAC": "LILAC",
    "LTC": "LIGHT CARBON",
    "LIGHT CARBON": "LIGHT CARBON",
    "LCM": "LIGHT CARBON MELANGE",
    "LIGHT CARBON MELANGE": "LIGHT CARBON MELANGE",
    "LRN": "LIGHT GREEN",
    "LIGHT GREEN": "LIGHT GREEN",
    "LTG": "LIGHT GRAY",
    "LIGHT GREY": "LIGHT GRAY",
    "LIGHT GRAY": "LIGHT GRAY",
    "LGM": "LIGHT GRAY MELANGE",
    "LIGHT GREY MELANGE": "LIGHT GRAY MELANGE",
    "LIGHT GRAY MELANGE": "LIGHT GRAY MELANGE",
    "LTL": "LIGHT LIME",
    "LIGHT LIME": "LIGHT LIME",
    "LTN": "LIGHT NAVY",
    "LIGHT NAVY": "LIGHT NAVY",
    "LNM": "LIGHT NAVY MELANGE",
    "LIGHT NAVY MELANGE": "LIGHT NAVY MELANGE",
    "LTO": "LIGHT OLIVE",
    "LIGHT OLIVE": "LIGHT OLIVE",
    "MGO": "MANGO",
    "MANGO": "MANGO",
    "MRN": "MAROON",
    "MAROON": "MAROON",
    "MNB": "MIDNIGHT BLUE",
    "MIDNIGHT BLUE": "MIDNIGHT BLUE",
    "MBM": "MIDNIGHT BLUE MELANGE",
    "MIDNIGHT BLUE MELANGE": "MIDNIGHT BLUE MELANGE",
    "MTG": "MINT GREEN",
    "MINT GREEN": "MINT GREEN",
    "MBN": "MODERN BROWN",
    "MODERN BROWN": "MODERN BROWN",
    "MON": "MOON",
    "MOON": "MOON",
    "MGN": "MOSS GREEN",
    "MOSS GREEN": "MOSS GREEN",
    "NVY": "NAVY",
    "NAVY": "NAVY",
    "NVM": "NAVY MELANGE",
    "NAVY MELANGE": "NAVY MELANGE",
    "OVE": "OLIVE",
    "OLIVE": "OLIVE",
    "OVM": "OLIVE MELANGE",
    "OLIVE MELANGE": "OLIVE MELANGE",
    "OIN": "ONION",
    "ONION": "ONION",
    "ORG": "ORANGE",
    "ORANGE": "ORANGE",
    "PCH": "PEACH",
    "PEACH": "PEACH",
    "PNK": "PINK",
    "PINK": "PINK",
    "PLB": "POOL BLUE",
    "POOL BLUE": "POOL BLUE",
    "RED": "RED",
    "REM": "RED MELANGE",
    "RED MELANGE": "RED MELANGE",
    "RPK": "ROSE PINK",
    "ROSE PINK": "ROSE PINK",
    "RYB": "ROYAL BLUE",
    "ROYAL BLUE": "ROYAL BLUE",
    "RST": "RUST",
    "RUST": "RUST",
    "SCB": "SCAL BLUE",
    "SCAL BLUE": "SCAL BLUE",
    "SLB": "SCHOOL BLUE",
    "SCHOOL BLUE": "SCHOOL BLUE",
    "SAN": "SEA GREEN",
    "SEA GREEN": "SEA GREEN",
    "SSE": "SHADED SPRUCE",
    "SHADED SPRUCE": "SHADED SPRUCE",
    "SWP": "SHADOW PURPLE",
    "SHADOW PURPLE": "SHADOW PURPLE",
    "SGY": "SILVER GRAY",
    "SILVER GREY": "SILVER GRAY",
    "SILVER GRAY": "SILVER GRAY",
    "SKY": "SKY BLUE",
    "SKY BLUE": "SKY BLUE",
    "SGN": "SNARKEL GREEN",
    "SNARKEL GREEN": "SNARKEL GREEN",
    "SPN": "SPACE NAVY",
    "SPACE NAVY": "SPACE NAVY",
    "SPM": "SPACE NAVY MEANGE",
    "SPACE NAVY MEANGE": "SPACE NAVY MEANGE",
    "SPG": "SPRING GREEN",
    "SPRING GREEN": "SPRING GREEN",
    "SEG": "STONE GRAY",
    "STONE GREY": "STONE GRAY",
    "STONE GRAY": "STONE GRAY",
    "SFR": "SUN FLOWER",
    "SUN FLOWER": "SUN FLOWER",
    "TEL": "TEAL",
    "TEAL": "TEAL",
    "TLM": "TEAL MELANGE",
    "TEAL MELANGE": "TEAL MELANGE",
    "TKM": "TURKISH MELANGE",
    "TURKISH MELANGE": "TURKISH MELANGE",
    "WAM": "WATER AQUA MELANGE",
    "WATER AQUA MELANGE": "WATER AQUA MELANGE",
    "WHT": "WHITE",
    "WHITE": "WHITE",
    "WGN": "WINTER GREEN",
    "WINTER GREEN": "WINTER GREEN",
    "YLW": "YELLOW",
    "YELLOW": "YELLOW",
    "OFW": "OFF WHITE",
    "OFF WHITE": "OFF WHITE",
    "CRB": "CAPRI BLUE",
    "CAPRI BLUE": "CAPRI BLUE",
    "CEM": "CAPRI BLUE MELANGE",
    "CAPRI BLUE MELANGE": "CAPRI BLUE MELANGE",
    "SAD": "SAND",
    "SAND": "SAND",
    "OML": "OAT MEAL",
    "OAT MEAL": "OAT MEAL",
    "PAE": "PINE APPLE",
    "PINE APPLE": "PINE APPLE",
    "CBE": "COBALT BLUE",
    "COBALT BLUE": "COBALT BLUE",
    "EBE": "EVENING BLUE",
    "EVENING BLUE": "EVENING BLUE",
    "BRD": "BERRY RED",
    "BERRY RED": "BERRY RED",
    "SBE": "SMOKE BLUE",
    "SMOKE BLUE": "SMOKE BLUE",
    "DRE": "DUSTY ROSE",
    "DUSTY ROSE": "DUSTY ROSE",
    "PGN": "PINE GREEN",
    "PINE GREEN": "PINE GREEN",
    "TLE": "TURBULENCE",
    "TURBULENCE": "TURBULENCE",
    "DSK": "DUSK",
    "DUSK": "DUSK",
    "RAB": "RAIN BLUE",
    "RAIN BLUE": "RAIN BLUE",
    "LRG": "LUNAR GRAY",
    "LUNAR GREY": "LUNAR GRAY",
    "LUNAR GRAY": "LUNAR GRAY",
    "DSB": "DRESS BLUE",
    "DRESS BLUE": "DRESS BLUE",
    "DSN": "DAMSON",
    "DAMSON": "DAMSON",
    "BEL": "BLUE TEAL",
    "BLUE TEAL": "BLUE TEAL",
    "PSA": "PISTA",
    "PISTA": "PISTA",
    "CTA": "CAT AQUA",
    "CAT AQUA": "CAT AQUA",
    "CTT": "CAT LIGHT TURKISH",
    "CAT LIGHT TURKISH": "CAT LIGHT TURKISH",
    "CTR": "CAT RED",
    "CAT RED": "CAT RED",
    "CTG": "CAT LIGHT GRAY",
    "CAT LIGHT GREY": "CAT LIGHT GRAY",
    "CAT LIGHT GRAY": "CAT LIGHT GRAY",
    "CTN": "CAT NAVY",
    "CAT NAVY": "CAT NAVY",
    "CTO": "CAT ONION",
    "CAT ONION": "CAT ONION",
    "CTL": "CAT LIGHT LIME",
    "CAT LIGHT LIME": "CAT LIGHT LIME",
    "CTV": "CAT BLUEBERRY VIOLET",
    "CAT BLUEBERRY VIOLET": "CAT BLUEBERRY VIOLET",
    "BTL": "BRIGHT TEAL",
    "BRIGHT TEAL": "BRIGHT TEAL",
    "AAE": "ALLURING AZURE",
    "ALLURING AZURE": "ALLURING AZURE",
    "BBE": "BRUSHED BLUE",
    "BRUSHED BLUE": "BRUSHED BLUE",
    "LLA": "LAVISH LILAC",
    "LAVISH LILAC": "LAVISH LILAC",
    "RSS": "REGAL STREAKS",
    "REGAL STREAKS": "REGAL STREAKS",
    "RRF": "ROSY REEF",
    "ROSY REEF": "ROSY REEF",
    "SYS": "SHADY SPUME",
    "SHADY SPUME": "SHADY SPUME",
    "SSR": "SOOTY SMEAR",
    "SOOTY SMEAR": "SOOTY SMEAR",
    "SNT": "STARRY NIGHT",
    "STARRY NIGHT": "STARRY NIGHT",
    "BTM": "BRIGHT TEAL MELANGE",
    "BRIGHT TEAL MELANGE": "BRIGHT TEAL MELANGE",
    "MLG": "MED LIGHT GRAY",
    "MED LIGHT GREY": "MED LIGHT GRAY",
    "MED LIGHT GRAY": "MED LIGHT GRAY",
    "EEG": "EASTER EGG",
    "EASTER EGG": "EASTER EGG",
    "DCH": "DAISY CHAIN",
    "DAISY CHAIN": "DAISY CHAIN",
    "BLT": "BLUE LIGHT",
    "BLUE LIGHT": "BLUE LIGHT",
    "WYL": "WAX YELLOW",
    "WAX YELLOW": "WAX YELLOW",
    "CSN": "CRIMSON",
    "CRIMSON": "CRIMSON",
    "GJD": "GRAYED JADE",
    "GRAYED JADE": "GRAYED JADE",
    "AGN": "ASSURE GREEN",
    "ASSURE GREEN": "ASSURE GREEN",
    "DKE": "DARK EARTH",
    "DARK EARTH": "DARK EARTH",
    "GHE": "GOLDEN HAZE",
    "GOLDEN HAZE": "GOLDEN HAZE",
    "SOM": "STORM",
    "STORM": "STORM",
    "DRD": "DUSTY RED",
    "DUSTY RED": "DUSTY RED",
    "RON": "RHODOENDRAN",
    "RHODOENDRAN": "RHODOENDRAN",
    "DGY": "DARK GRAY",
    "DARK GREY": "DARK GRAY",
    "DARK GRAY": "DARK GRAY",
    "NGN": "NEON GREEN",
    "NEON GREEN": "NEON GREEN",
    "RBD": "RUBY RED",
    "RUBY RED": "RUBY RED",
    "LPE": "LIGHT PURPLE",
    "LIGHT PURPLE": "LIGHT PURPLE",
    "MLN": "MED LIGHT NAVY",
    "MED LIGHT NAVY": "MED LIGHT NAVY",
    "LBE": "LICHEN BLUE",
    "LICHEN BLUE": "LICHEN BLUE",
    "ESY": "ENDLESS SKY",
    "ENDLESS SKY": "ENDLESS SKY",
    "SSL": "SET SAIL",
    "SET SAIL": "SET SAIL",
    "HMS": "HUMUS",
    "HUMUS": "HUMUS",
    "SRT": "SKY ROCKET",
    "SKY ROCKET": "SKY ROCKET",
    "ADP": "ATLANTIC DEEP",
    "ATLANTIC DEEP": "ATLANTIC DEEP",
    "RRN": "RUM RAISIN",
    "RUM RAISIN": "RUM RAISIN",
    "MUE": "MAUVE",
    "MAUVE": "MAUVE",
    "UGY": "ULTIMATE GRAY",
    "ULTIMATE GREY": "ULTIMATE GRAY",
    "ULTIMATE GRAY": "ULTIMATE GRAY",
    "CGE": "COGNAE",
    "COGNAE": "COGNAE",
    "SSD": "SUMMER SAND",
    "SUMMER SAND": "SUMMER SAND",
    "RBM": "RASBERRY RADIENCE MIL",
    "RASBERRY RADIENCE MIL": "RASBERRY RADIENCE MIL",
    "CIM": "CHILLI OIL MIL",
    "CHILLI OIL MIL": "CHILLI OIL MIL",
    "CCM": "COCOA CREAM",
    "COCOA CREAM": "COCOA CREAM",
    "ARA": "AURORA",
    "AURORA": "AURORA",
    "RSR": "RASPBERRY RADIANCE",
    "RASPBERRY RADIANCE": "RASPBERRY RADIANCE",
    "COL": "CHILLI OIL",
    "CHILLI OIL": "CHILLI OIL",
    "GRY": "GRAY",
    "GREY": "GRAY",
    "GRAY": "GRAY",
    "MDM": "MEDIUM DENIM",
    "MEDIUM DENIM": "MEDIUM DENIM",
    "RLK": "ROYAL PINK",
    "ROYAL PINK": "ROYAL PINK",
    "GYM": "GRAY MELANGE",
    "GREY MELANGE": "GRAY MELANGE",
    "GRAY MELANGE": "GRAY MELANGE",
    "AQS": "AQUA SEA",
    "AQUA SEA": "AQUA SEA",
    "IAO": "INDIA ORANGE",
    "INDIA ORANGE": "INDIA ORANGE",
    "BRN": "BROWN",
    "BROWN": "BROWN",
    "TMC": "TARMAC",
    "TARMAC": "TARMAC",
    "SNM": "SPACE NAVY MELANGE",
    "SPACE NAVY MELANGE": "SPACE NAVY MELANGE",
    "RSM": "RUST MELANGE",
    "RUST MELANGE": "RUST MELANGE",
    "DRM": "DUSTY ROSE MELANGE",
    "DUSTY ROSE MELANGE": "DUSTY ROSE MELANGE",
    "SBM": "SCAL BLUE MELANGE",
    "SCAL BLUE MELANGE": "SCAL BLUE MELANGE",
    "CLM": "CAT LIGHT GRAY MELANGE",
    "CAT LIGHT GREY MELANGE": "CAT LIGHT GRAY MELANGE",
    "CAT LIGHT GRAY MELANGE": "CAT LIGHT GRAY MELANGE",
    "OLT": "OVERLAND TREK",
    "OVERLAND TREK": "OVERLAND TREK",
    "DKR": "DARK RED",
    "DARK RED": "DARK RED",
    "SLE": "SALUTE",
    "SALUTE": "SALUTE",
    "RVA": "RIVERA",
    "RIVERA": "RIVERA",
    "SPK": "SILVER PINK",
    "SILVER PINK": "SILVER PINK",
    "WWD": "WILD WIND",
    "WILD WIND": "WILD WIND",
    "LTB": "LIGHT BLUE",
    "LIGHT BLUE": "LIGHT BLUE",
    "WAQ": "WATER AQUA",
    "WATER AQUA": "WATER AQUA",
    "CLY": "CLAY",
    "CLAY": "CLAY",
    "LDN": "LINDEN",
    "LINDEN": "LINDEN",
    "CWA": "CAT WATER AQUA",
    "CAT WATER AQUA": "CAT WATER AQUA",
    "BTE": "BEETLE",
    "BEETLE": "BEETLE",
    "CFB": "COFFEE BEAN",
    "COFFEE BEAN": "COFFEE BEAN",
    "CCY": "CAT CLAY",
    "CAT CLAY": "CAT CLAY",
    "SBK": "SILVER BRICK",
    "SILVER BRICK": "SILVER BRICK",
    "DTE": "DEEP TAUPE",
    "DEEP TAUPE": "DEEP TAUPE",
    "GUL": "GULL",
    "GULL": "GULL",
    "KUI": "KUI",
    "BGE": "BLUE GRANITE",
    "BLUE GRANITE": "BLUE GRANITE",
    "CBK": "CAT BLACK",
    "CAT BLACK": "CAT BLACK",
    "PCL": "PEACH CARAMEL",
    "PEACH CARAMEL": "PEACH CARAMEL",
    "ABT": "AMERICAN BEAUTY",
    "AMERICAN BEAUTY": "AMERICAN BEAUTY",
    "PCA": "PLUM CASPA",
    "PLUM CASPA": "PLUM CASPA",
    "MDG": "MALLARD GREEN",
    "MALLARD GREEN": "MALLARD GREEN",
    "HPK": "HOT PINK",
    "HOT PINK": "HOT PINK",
    "BRM": "BERRY RED MELANGE",
    "BERRY RED MELANGE": "BERRY RED MELANGE",
    "CYM": "CLAY MELANGE",
    "CLAY MELANGE": "CLAY MELANGE",
    "ASM": "AQUA SEA MELANGE",
    "AQUA SEA MELANGE": "AQUA SEA MELANGE",
    "SGM": "STONE GRAY MELANGE",
    "STONE GREY MELANGE": "STONE GRAY MELANGE",
    "STONE GRAY MELANGE": "STONE GRAY MELANGE",
    "SLM": "SALUTE MELANGE",
    "SALUTE MELANGE": "SALUTE MELANGE",
    "CNG": "CARBON GRAY",
    "CARBON GREY": "CARBON GRAY",
    "CARBON GRAY": "CARBON GRAY",
    "PGM": "PINE GREEN MELANGE",
    "PINE GREEN MELANGE": "PINE GREEN MELANGE",
    "CGM": "CARBON GRAY MELANGE",
    "CARBON GREY MELANGE": "CARBON GRAY MELANGE",
    "CARBON GRAY MELANGE": "CARBON GRAY MELANGE",
    "EBM": "EVENING BLUE MELANGE",
    "EVENING BLUE MELANGE": "EVENING BLUE MELANGE",
    "BLL": "BLACK LUNAR GRAY",
    "BLACK LUNAR GREY": "BLACK LUNAR GRAY",
    "BLACK LUNAR GRAY": "BLACK LUNAR GRAY",
    "BLS": "BLACK STONE GRAY",
    "BLACK STONE GREY": "BLACK STONE GRAY",
    "BLACK STONE GRAY": "BLACK STONE GRAY",
    "BLB": "BLACK BEETLE",
    "BLACK BEETLE": "BLACK BEETLE",
    "BLD": "BLACK DR CARBON",
    "BLACK DR CARBON": "BLACK DR CARBON",
    "BLI": "BLACK LINDEN",
    "BLACK LINDEN": "BLACK LINDEN",
    "BLGR": "BLACK GARP",
    "BLACK GARP": "BLACK GARP",
    "BLEX": "BLACK EXPLORE",
    "BLACK EXPLORE": "BLACK EXPLORE",
    "WGM": "WINTER GREEN MELANGE",
    "WINTER GREEN MELANGE": "WINTER GREEN MELANGE",
    "BCG": "BLACK CARBON GRAY",
    "BLACK CARBON GREY": "BLACK CARBON GRAY",
    "BLACK CARBON GRAY": "BLACK CARBON GRAY",
    "GVE": "GRAPE VINE",
    "GRAPE VINE": "GRAPE VINE",
    "GVM": "GRAPE VINE MELANGE",
    "GRAPE VINE MELANGE": "GRAPE VINE MELANGE",
    "FIG": "FIG",
    "BTD": "BLUE TIDE",
    "BLUE TIDE": "BLUE TIDE",
    "IBE": "INK BLUE",
    "INK BLUE": "INK BLUE",
    "CMT": "COOL MINT",
    "COOL MINT": "COOL MINT",
    "NPY": "NAVY PEONY",
    "NAVY PEONY": "NAVY PEONY",
    "ASH": "ASH",
    "FGV": "FOX GLOVE",
    "FOX GLOVE": "FOX GLOVE",
    "BEM": "BLUE TIDE MELANGE",
    "BLUE TIDE MELANGE": "BLUE TIDE MELANGE",
    "BIG": "BLACK IRON GRAY",
    "BLACK IRON GREY": "BLACK IRON GRAY",
    "BLACK IRON GRAY": "BLACK IRON GRAY",
    "BCY": "BLACK CITY GRAY",
    "BLACK CITY GREY": "BLACK CITY GRAY",
    "BLACK CITY GRAY": "BLACK CITY GRAY",
    "TAS": "TRINAGLES",
    "TRINAGLES": "TRINAGLES",
    "CLB": "COOL BLUE",
    "COOL BLUE": "COOL BLUE",
    "SYN": "STARRY NIGHT",
    "FNS": "FLATTEN STONE",
    "FLATTEN STONE": "FLATTEN STONE",
    "NTM": "NIGHTMARE",
    "NIGHTMARE": "NIGHTMARE",
    "BKM": "BLACK MAGIC",
    "BLACK MAGIC": "BLACK MAGIC",
    "CNM": "COFFEE BEAN MELANGE",
    "COFFEE BEAN MELANGE": "COFFEE BEAN MELANGE",
    "FGM": "FIG MELANGE",
    "FIG MELANGE": "FIG MELANGE",
    "COM": "CHILLI OIL MELANGE",
    "CHILLI OIL MELANGE": "CHILLI OIL MELANGE",
    "LYM": "LUNAR GRAY MELANGE",
    "LUNAR GREY MELANGE": "LUNAR GRAY MELANGE",
    "LUNAR GRAY MELANGE": "LUNAR GRAY MELANGE",
    "CLS": "CLAWS",
    "CLAWS": "CLAWS",
    "FSL": "FREE STYLE LITE",
    "FREE STYLE LITE": "FREE STYLE LITE",
    "TIL": "TRI LOG",
    "TRI LOG": "TRI LOG",
    "MSS": "MELANGEANGE STRIPES",
    "MELANGEANGE STRIPES": "MELANGEANGE STRIPES",
    "BOE": "BLACK OLIVE",
    "BLACK OLIVE": "BLACK OLIVE",
    "RBN": "RUSTIC BROWN",
    "RUSTIC BROWN": "RUSTIC BROWN",
    "DBE": "DARK BLUE",
    "DARK BLUE": "DARK BLUE",
    "PBM": "POOL BLUE MELANGE",
    "POOL BLUE MELANGE": "POOL BLUE MELANGE",
    "BNM": "BROWN MELANGE",
    "BROWN MELANGE": "BROWN MELANGE",
    "UGM": "ULTIMATE GRAY MELANGE",
    "ULTIMATE GREY MELANGE": "ULTIMATE GRAY MELANGE",
    "ULTIMATE GRAY MELANGE": "ULTIMATE GRAY MELANGE",
    "AQA": "AQUA",
    "AQUA": "AQUA",
    "BTH": "BLUE TURKISH",
    "BLUE TURKISH": "BLUE TURKISH",
    "LLG": "LAUREL GREEN",
    "LAUREL GREEN": "LAUREL GREEN",
    "LNR": "LAVENDER",
    "LAVENDER": "LAVENDER",
    "LCB": "LILAC BREEZE",
    "LILAC BREEZE": "LILAC BREEZE",
    "LSE": "LIMESTONE",
    "LIMESTONE": "LIMESTONE",
    "MBE": "MAJOLICA BLUE",
    "MAJOLICA BLUE": "MAJOLICA BLUE",
    "NPK": "NEON PINK",
    "NEON PINK": "NEON PINK",
    "PPA": "PAPAYA",
    "PAPAYA": "PAPAYA",
    "PPE": "POTENT PURPLE",
    "POTENT PURPLE": "POTENT PURPLE",
    "AME": "AURORA MELANGE",
    "AURORA MELANGE": "AURORA MELANGE",
    "MBG": "MODERN BROWN MELANGE",
    "MODERN BROWN MELANGE": "MODERN BROWN MELANGE",
    "RVM": "RIVERA MELANGE",
    "RIVERA MELANGE": "RIVERA MELANGE",
    "SGE": "SNARKEL GREEN MELANGE",
    "SNARKEL GREEN MELANGE": "SNARKEL GREEN MELANGE",
    "WWM": "WILD WIND MELANGE",
    "WILD WIND MELANGE": "WILD WIND MELANGE",
    "PKN": "PUMPKIN",
    "PUMPKIN": "PUMPKIN",
    "TGN": "TANGERINE",
    "TANGERINE": "TANGERINE",
    "BTQ": "BLUE TURQUOISE",
    "BLUE TURQUOISE": "BLUE TURQUOISE",
    "BLG": "BLACK-LT GRAY",
    "BLACK-LT GREY": "BLACK-LT GRAY",
    "BLACK-LT GRAY": "BLACK-LT GRAY",
    "OME": "ONION MELANGE",
    "ONION MELANGE": "ONION MELANGE",
    "DYM": "DUSTY RED MELANGE",
    "DUSTY RED MELANGE": "DUSTY RED MELANGE",
    "PDB": "POWDER BLUE",
    "POWDER BLUE": "POWDER BLUE",
    "TSP": "TROPOSPHERE",
    "TROPOSPHERE": "TROPOSPHERE",
    "BME": "BEETLE MELANGE",
    "BEETLE MELANGE": "BEETLE MELANGE",
    "KHI": "KHAKI",
    "KHAKI": "KHAKI",
    "SDL": "SHADOW LIME",
    "SHADOW LIME": "SHADOW LIME",
    "OLD": "OMPHALODES",
    "OMPHALODES": "OMPHALODES",
    "GFT": "GULF COAST",
    "GULF COAST": "GULF COAST",
    "SWS": "SMOKE WAVES",
    "SMOKE WAVES": "SMOKE WAVES",
    "CSD": "CROSSROAD",
    "CROSSROAD": "CROSSROAD",
    "RFT": "RAIN FOREST",
    "RAIN FOREST": "RAIN FOREST",
    "BSS": "BLUE STROKES",
    "BLUE STROKES": "BLUE STROKES",
    "PLS": "PURPLE LINES",
    "PURPLE LINES": "PURPLE LINES",
    "NSY": "NIGHT SKY",
    "NIGHT SKY": "NIGHT SKY",
    "CBT": "CLOUD BURST",
    "CLOUD BURST": "CLOUD BURST",
    "GYS": "GRAY SMUDGE",
    "GREY SMUDGE": "GRAY SMUDGE",
    "GRAY SMUDGE": "GRAY SMUDGE",
    "WCR": "WATER COLOUR",
    "WATER COLOUR": "WATER COLOUR",
    "TSM": "THUNDER STORM",
    "THUNDER STORM": "THUNDER STORM",
    "PET": "PURPLE EFFECT",
    "PURPLE EFFECT": "PURPLE EFFECT",
    "RSE": "RED SMUDGE",
    "RED SMUDGE": "RED SMUDGE",
    "TEM": "TURBULENCE MELANGE",
    "TURBULENCE MELANGE": "TURBULENCE MELANGE",
    "LBM": "LICHEN BLUE MELANGE",
    "LICHEN BLUE MELANGE": "LICHEN BLUE MELANGE",
    "LLM": "LIGHT LIME MELANGE",
    "LIGHT LIME MELANGE": "LIGHT LIME MELANGE",
    "BWS": "BLUE WAVES",
    "BLUE WAVES": "BLUE WAVES",
    "GME": "GRAY MARBLE",
    "GREY MARBLE": "GRAY MARBLE",
    "GRAY MARBLE": "GRAY MARBLE",
    "GNB": "GREEN BOG",
    "GREEN BOG": "GREEN BOG",
    "BCL": "BLUE CORAL",
    "BLUE CORAL": "BLUE CORAL",
    "BSE": "BLACK STROKE",
    "BLACK STROKE": "BLACK STROKE",
    "BSM": "BLACK STORM",
    "BLACK STORM": "BLACK STORM",
    "GCM": "GULF COAST MELANGE",
    "GULF COAST MELANGE": "GULF COAST MELANGE",
    "PEM": "POWDER BLUE MELANGE",
    "POWDER BLUE MELANGE": "POWDER BLUE MELANGE",
    "DTM": "DEEP TAUPE MELANGE",
    "DEEP TAUPE MELANGE": "DEEP TAUPE MELANGE",
    "PSY": "PURPLE SPRAY",
    "PURPLE SPRAY": "PURPLE SPRAY",
    "CSH": "CAMO SPLASH",
    "CAMO SPLASH": "CAMO SPLASH",
    "GBH": "GREEN BRUSH",
    "GREEN BRUSH": "GREEN BRUSH",
    "BBH": "BROWN BRUSH",
    "BROWN BRUSH": "BROWN BRUSH",
    "GMT": "GLACIER MELT",
    "GLACIER MELT": "GLACIER MELT",
    "CCD": "CAMO CLOUD",
    "CAMO CLOUD": "CAMO CLOUD",
    "PSS": "PURPLE SHARDS",
    "PURPLE SHARDS": "PURPLE SHARDS",
    "BES": "BLUE STRIPES",
    "BLUE STRIPES": "BLUE STRIPES",
    "OSE": "OLIVE SMUDGE",
    "OLIVE SMUDGE": "OLIVE SMUDGE",
    "BCO": "BLACK CAMO",
    "BLACK CAMO": "BLACK CAMO",
    "SES": "SHAPE SHIFTER",
    "SHAPE SHIFTER": "SHAPE SHIFTER",
    "BKG": "BLACK GRUNGE",
    "BLACK GRUNGE": "BLACK GRUNGE",
    "NTC": "NORTH ATLANTIC",
    "NORTH ATLANTIC": "NORTH ATLANTIC",
    "AGY": "AQUA GRAY",
    "AQUA GREY": "AQUA GRAY",
    "AQUA GRAY": "AQUA GRAY",
    "CTF": "CHOCOLATE TRUFFLE",
    "CHOCOLATE TRUFFLE": "CHOCOLATE TRUFFLE",
    "TRN": "TREKKING GREEN",
    "TREKKING GREEN": "TREKKING GREEN",
    "BGY": "BLUE GRAY",
    "BLUE GREY": "BLUE GRAY",
    "BLUE GRAY": "BLUE GRAY",
    "DTY": "DARK TEAL",
    "DARK TEAL": "DARK TEAL",
    "HTR": "HEATHER",
    "HEATHER": "HEATHER",
    "FCT": "FOOTBALL COURT",
    "FOOTBALL COURT": "FOOTBALL COURT",
    "CST": "CYCLE SPORT",
    "CYCLE SPORT": "CYCLE SPORT",
    "BMN": "BADMINTON",
    "BADMINTON": "BADMINTON",
    "TTS": "TABLE TENNIS",
    "TABLE TENNIS": "TABLE TENNIS",
    "TNS": "TENNIS",
    "TENNIS": "TENNIS",
    "FST": "FOOTBALL SPORT",
    "FOOTBALL SPORT": "FOOTBALL SPORT",
    "ACO": "AQUA CAMO",
    "AQUA CAMO": "AQUA CAMO",
    "GLS": "GRASSLANDS",
    "GRASSLANDS": "GRASSLANDS",
    "FBE": "FUMY BLUE",
    "FUMY BLUE": "FUMY BLUE",
    "WDK": "WILD DUSK",
    "WILD DUSK": "WILD DUSK",
    "GTH": "GLITCH",
    "GLITCH": "GLITCH",
    "DMR": "DARK MATTER",
    "DARK MATTER": "DARK MATTER",
    "DPK": "DUSTY PINK",
    "DUSTY PINK": "DUSTY PINK",
    "SWR": "STORMY WEATHER",
    "STORMY WEATHER": "STORMY WEATHER",
    "BHG": "BLACK HUNTER GREEN",
    "BLACK HUNTER GREEN": "BLACK HUNTER GREEN",
    "BKR": "BLACK RED",
    "BLACK RED": "BLACK RED",
    "BBT": "BLACK BLUE TIDE",
    "BLACK BLUE TIDE": "BLACK BLUE TIDE",
    "LNT": "LIGHT NAVY TROPOSPHERE",
    "LIGHT NAVY TROPOSPHERE": "LIGHT NAVY TROPOSPHERE",
    "LNS": "LIGHT NAVY SHADOW LIME",
    "LIGHT NAVY SHADOW LIME": "LIGHT NAVY SHADOW LIME",
    "LNW": "LIGHT NAVY WHITE",
    "LIGHT NAVY WHITE": "LIGHT NAVY WHITE",
    "OGY": "OYSTER GRAY",
    "OYSTER GRAY": "OYSTER GRAY",
    "HBE": "HORIZON BLUE",
    "HORIZON BLUE": "HORIZON BLUE",
    "TCA": "TERRACOTTA",
    "TERRACOTTA": "TERRACOTTA",
    "NAM": "NORTH ATLANTIC MELANGE",
    "NORTH ATLANTIC MELANGE": "NORTH ATLANTIC MELANGE",
    "IGM": "IRON GRAY MELANGE",
    "IRON GREY MELANGE": "IRON GRAY MELANGE",
    "IRON GRAY MELANGE": "IRON GRAY MELANGE",
    "HBM": "HORIZON BLUE MELANGE",
    "HORIZON BLUE MELANGE": "HORIZON BLUE MELANGE",
    "TCM": "TERRACOTTA MELANGE",
    "TERRACOTTA MELANGE": "TERRACOTTA MELANGE",
    "CTM": "CHOCOLATE TRUFFLE MELANGE",
    "CHOCOLATE TRUFFLE MELANGE": "CHOCOLATE TRUFFLE MELANGE",
    "CWB": "COB WEB",
    "COB WEB": "COB WEB",
    "PWD": "PLYWOOD",
    "PLYWOOD": "PLYWOOD",
    "WPR": "WALL PAPER",
    "WALL PAPER": "WALL PAPER",
    "WWH": "WATER WASH",
    "WATER WASH": "WATER WASH",
    "GNS": "GREEN SMUDGE",
    "GREEN SMUDGE": "GREEN SMUDGE",
    "GCO": "GRAY CAMO",
    "GREY CAMO": "GRAY CAMO",
    "GRAY CAMO": "GRAY CAMO",
    "CCO": "CARBON CAMO",
    "CARBON CAMO": "CARBON CAMO",
    "DCC": "DARK CARBON CAMO",
    "DARK CARBON CAMO": "DARK CARBON CAMO",
    "GSE": "GREEN SMOKE",
    "GREEN SMOKE": "GREEN SMOKE",
    "RDS": "RED SMOKE",
    "RED SMOKE": "RED SMOKE",
    "TLT": "TEAL TRACKS",
    "TEAL TRACKS": "TEAL TRACKS",
    "GPE": "GRAY PLAGUE",
    "GREY PLAGUE": "GRAY PLAGUE",
    "GRAY PLAGUE": "GRAY PLAGUE",
    "TME": "TROPOSPHERE MELANGE",
    "TROPOSPHERE MELANGE": "TROPOSPHERE MELANGE",
    "BVM": "BLUE BERRY VIOLET MELANGE",
    "BLUE BERRY VIOLET MELANGE": "BLUE BERRY VIOLET MELANGE",
    "LEM": "LAKE GREEN MELANGE",
    "LAKE GREEN MELANGE": "LAKE GREEN MELANGE",
    "ONW": "OCEAN WAVES",
    "OCEAN WAVES": "OCEAN WAVES",
    "FCS": "FOREST CUBES",
    "FOREST CUBES": "FOREST CUBES",
    "VWL": "VINTAGE WALL",
    "VINTAGE WALL": "VINTAGE WALL",
    "JVS": "JUNGLE VINES",
    "JUNGLE VINES": "JUNGLE VINES",
    "ISN": "INK STAIN",
    "INK STAIN": "INK STAIN",
    "MCS": "MELTING CRAYONS",
    "MELTING CRAYONS": "MELTING CRAYONS",
    "OBR": "OX BLOOD RED",
    "OX BLOOD RED": "OX BLOOD RED",
    "GYE": "GRAY STROKE",
    "GREY STROKE": "GRAY STROKE",
    "GRAY STROKE": "GRAY STROKE",
    "PSE": "PURPLE SMOKE",
    "PURPLE SMOKE": "PURPLE SMOKE",
    "BEC": "BLUE CAMO",
    "BLUE CAMO": "BLUE CAMO",
    "TGE": "TEAL GRUNGE",
    "TEAL GRUNGE": "TEAL GRUNGE",
    "BNS": "BROWN STRIPES",
    "BROWN STRIPES": "BROWN STRIPES",
    "GSS": "GREEN STRIPES",
    "GREEN STRIPES": "GREEN STRIPES",
    "CSE": "CLOUDSCAPE",
    "CLOUDSCAPE": "CLOUDSCAPE",
    "RGS": "RAIN GLASS",
    "RAIN GLASS": "RAIN GLASS",
    "GGH": "GREEN GLITCH",
    "GREEN GLITCH": "GREEN GLITCH",
    "OMS": "OMINOUS",
    "OMINOUS": "OMINOUS",
    "PME": "PUMPKIN MELANGE",
    "PUMPKIN MELANGE": "PUMPKIN MELANGE",
    "BEG": "BLUE GENIE",
    "BLUE GENIE": "BLUE GENIE",
    "BVA": "BOUGAINVILLEA",
    "BOUGAINVILLEA": "BOUGAINVILLEA",
    "BMD": "BREAD MOULD",
    "BREAD MOULD": "BREAD MOULD",
    "RRM": "RIPPLE REALM",
    "RIPPLE REALM": "RIPPLE REALM",
    "TPN": "TYPHOON",
    "TYPHOON": "TYPHOON",
    "GWE": "GRAY WAVE",
    "GREY WAVE": "GRAY WAVE",
    "GRAY WAVE": "GRAY WAVE",
    "LPD": "LEOPARD",
    "LEOPARD": "LEOPARD",
    "PEP": "PURPLE PUDDLE",
    "PURPLE PUDDLE": "PURPLE PUDDLE",
    "CEB": "CODE BLUE",
    "CODE BLUE": "CODE BLUE",
    "VTX": "VORTEX",
    "VORTEX": "VORTEX",
    "TTT": "TIC TAC TOE",
    "TIC TAC TOE": "TIC TAC TOE",
    "ATY": "ALLOTROPHY",
    "ALLOTROPHY": "ALLOTROPHY",
    "SDE": "SAND DUNE",
    "SAND DUNE": "SAND DUNE",
    "KPN": "KRYPTON",
    "KRYPTON": "KRYPTON",
    "MMM": "MACRO MESO MICRO",
    "MACRO MESO MICRO": "MACRO MESO MICRO",
    "COT": "COBALT",
    "COBALT": "COBALT",
    "LLD": "LEGO LAND",
    "LEGO LAND": "LEGO LAND",
    "MSM": "MOLECULAR SPECTRUM",
    "MOLECULAR SPECTRUM": "MOLECULAR SPECTRUM",
    "PFT": "PINE FOREST",
    "PINE FOREST": "PINE FOREST",
    "PDS": "PIXELATED STEPS",
    "PIXELATED STEPS": "PIXELATED STEPS",
    "ABN": "ANT BATTALION",
    "ANT BATTALION": "ANT BATTALION",
    "JSC": "JIGSAW CAMO",
    "JIGSAW CAMO": "JIGSAW CAMO",
    "PLP": "PARALLEL PULSE",
    "PARALLEL PULSE": "PARALLEL PULSE",
    "ASL": "AQUA SWIRL",
    "AQUA SWIRL": "AQUA SWIRL",
    "LFX": "LINEAR FLUX",
    "LINEAR FLUX": "LINEAR FLUX",
    "PHN": "PAINTED HORIZON",
    "PAINTED HORIZON": "PAINTED HORIZON",
    "WGW": "WAVEGLOW",
    "WAVEGLOW": "WAVEGLOW",
    "ASK": "ABSTRACT STREAK",
    "ABSTRACT STREAK": "ABSTRACT STREAK",
    "TGM": "TREKKING GREEN MELANGE",
    "TREKKING GREEN MELANGE": "TREKKING GREEN MELANGE",
    "AGM": "AQUA GRAY MELANGE",
    "AQUA GREY MELANGE": "AQUA GRAY MELANGE",
    "AQUA GRAY MELANGE": "AQUA GRAY MELANGE",
    "PKP": "PINK PLUMAGE",
    "PINK PLUMAGE": "PINK PLUMAGE",
    "BLP": "BLUE LEOPARD",
    "BLUE LEOPARD": "BLUE LEOPARD",
    "BHS": "BRUSH STROKES",
    "BRUSH STROKES": "BRUSH STROKES",
    "BKD": "BLACK DIAMOND",
    "BLACK DIAMOND": "BLACK DIAMOND",
    "LCC": "LILAC CAMO",
    "LILAC CAMO": "LILAC CAMO",
    "OBM": "OX BLOOD RED MELANGE",
    "OX BLOOD RED MELANGE": "OX BLOOD RED MELANGE",
    "OGM": "OYSTER GRAY MELANGE",
    "OYSTER GRAY MELANGE": "OYSTER GRAY MELANGE",
    "MSA": "MARSALA",
    "MARSALA": "MARSALA",
    "FXM": "FOX GLOVE MELANGE",
    "FOX GLOVE MELANGE": "FOX GLOVE MELANGE",
    "SME": "SCHOOL BLUE MELANGE",
    "SCHOOL BLUE MELANGE": "SCHOOL BLUE MELANGE",
    "TAM": "TARMAC MELANGE",
    "TARMAC MELANGE": "TARMAC MELANGE",
    "SKO": "SMOKEY OLIVE",
    "SMOKEY OLIVE": "SMOKEY OLIVE",
    "WLG": "WOODLAND GRAY",
    "WOODLAND GRAY": "WOODLAND GRAY",
    "CFK": "CHILI FLAKES",
    "CHILI FLAKES": "CHILI FLAKES",
    "AHM": "ASH MELANGE",
    "ASH MELANGE": "ASH MELANGE",
    "BOM": "BLACK OLIVE MELANGE",
    "BLACK OLIVE MELANGE": "BLACK OLIVE MELANGE",
    "BST": "BLUESTONE",
    "BLUESTONE": "BLUESTONE",
    "GYR": "GRAY RIDGE",
    "GRAY RIDGE": "GRAY RIDGE",
    "PGI": "PIXELATED GRAFFITI",
    "PIXELATED GRAFFITI": "PIXELATED GRAFFITI",
    "DDC": "DUSTED CAMO",
    "DUSTED CAMO": "DUSTED CAMO",
    "LHS": "LAYERED HEIGHTS",
    "LAYERED HEIGHTS": "LAYERED HEIGHTS",
    "WFE": "WILD FIRE",
    "WILD FIRE": "WILD FIRE",
    "OCT": "OVERCAST",
    "OVERCAST": "OVERCAST",
    "MYE": "MYRTLE",
    "MYRTLE": "MYRTLE",
    "VIO": "VINTAGE INDIGO",
    "VINTAGE INDIGO": "VINTAGE INDIGO",
    "GBE": "GRAY BOSCAGE",
    "GREY BOSCAGE": "GRAY BOSCAGE",
    "GRAY BOSCAGE": "GRAY BOSCAGE",
    "DMT": "DARK MIST",
    "DARK MIST": "DARK MIST",
    "GNM": "GREEN MIST",
    "GREEN MIST": "GREEN MIST",
    "BEB": "BLUE BOSCAGE",
    "BLUE BOSCAGE": "BLUE BOSCAGE",
    "GMI": "GRAY MIST",
    "GREY MIST": "GRAY MIST",
    "GRAY MIST": "GRAY MIST",
    "SCT": "SHAPECRAFT",
    "SHAPECRAFT": "SHAPECRAFT",
    "HRT": "HORIZONTAL RIFT",
    "HORIZONTAL RIFT": "HORIZONTAL RIFT",
    "BLH": "BLUE LABYRINTH",
    "BLUE LABYRINTH": "BLUE LABYRINTH",
    "DLH": "DARK LABYRINTH",
    "DARK LABYRINTH": "DARK LABYRINTH",
    "BAS": "BLUE AMORPHOUS",
    "BLUE AMORPHOUS": "BLUE AMORPHOUS",
    "ESE": "EBONY SMUDGE",
    "EBONY SMUDGE": "EBONY SMUDGE",
    "DCO": "DESERT CAMO",
    "DESERT CAMO": "DESERT CAMO",
    "GMR": "GRAY MATTER",
    "GREY MATTER": "GRAY MATTER",
    "GRAY MATTER": "GRAY MATTER",
    "GST": "GREEN STROKES",
    "GREEN STROKES": "GREEN STROKES",
    "SPE": "STICK PUZZLE",
    "STICK PUZZLE": "STICK PUZZLE",
    "LME": "LINE MAZE",
    "LINE MAZE": "LINE MAZE",
    "SLD": "SEAMLESS DIAMOND",
    "SEAMLESS DIAMOND": "SEAMLESS DIAMOND",
    "SLG": "SEAMLESS GEOMETRICAL",
    "SEAMLESS GEOMETRICAL": "SEAMLESS GEOMETRICAL",
    "SLC": "SEAMLESS CROSS",
    "SEAMLESS CROSS": "SEAMLESS CROSS",
    "DCS": "DOUBLE CROSS",
    "DOUBLE CROSS": "DOUBLE CROSS",
    "GYB": "GRAY BLUE",
    "GREY BLUE": "GRAY BLUE",
    "GRAY BLUE": "GRAY BLUE",
    "RRE": "RUM RAISIN MELANGE",
    "RUM RAISIN MELANGE": "RUM RAISIN MELANGE",
    "BLE": "BLUE STONE",
    "BLUE STONE": "BLUE STONE",
    "SRG": "SILVER GRAY MELANGE",
    "SILVER GREY MELANGE": "SILVER GRAY MELANGE",
    "SILVER GRAY MELANGE": "SILVER GRAY MELANGE",
    "MMX": "MIRAGE MATRIX",
    "MIRAGE MATRIX": "MIRAGE MATRIX",
    "ZZH": "ZIGZAG ZENITH",
    "ZIGZAG ZENITH": "ZIGZAG ZENITH",
    "CHS": "CHRONOS SQUARE",
    "CHRONOS SQUARE": "CHRONOS SQUARE",
    "GCE": "GRUNGE CUBE",
    "GRUNGE CUBE": "GRUNGE CUBE",
    "RDT": "RHOMBUS DRIFT",
    "RHOMBUS DRIFT": "RHOMBUS DRIFT",
    "EET": "ECHO ELEMENT",
    "ECHO ELEMENT": "ECHO ELEMENT",
    "NYB": "NAVY B",
    "NAVY B": "NAVY B",
    "KME": "KHAKI MELANGE",
    "KHAKI MELANGE": "KHAKI MELANGE",
    "AGW": "AZURE GLOW",
    "AZURE GLOW": "AZURE GLOW",
    "GLB": "GLACIAL BLUE",
    "GLACIAL BLUE": "GLACIAL BLUE",
    "IBM": "INDIGO BLOOM",
    "INDIGO BLOOM": "INDIGO BLOOM",
    "CBH": "CERULEAN BLUSH",
    "CERULEAN BLUSH": "CERULEAN BLUSH",
    "DHE": "DUAL HUE",
    "DUAL HUE": "DUAL HUE",
    "FCY": "FIERY CANOPY",
    "FIERY CANOPY": "FIERY CANOPY",
    "MME": "MARSALA MELANGE",
    "MARSALA MELANGE": "MARSALA MELANGE",
    "BET": "BLUE LIGHT MELANGE",
    "BLUE LIGHT MELANGE": "BLUE LIGHT MELANGE",
    "CFM": "CHILI FLAKES MELANGE",
    "CHILI FLAKES MELANGE": "CHILI FLAKES MELANGE",
    "NBM": "NAVY B MELANGE",
    "NAVY B MELANGE": "NAVY B MELANGE",
    "MEM": "MYRTLE MELANGE",
    "MYRTLE MELANGE": "MYRTLE MELANGE",
    "SYM": "SMOKEY OLIVE MELANGE",
    "SMOKEY OLIVE MELANGE": "SMOKEY OLIVE MELANGE",
    "WDM": "WOODLAND GRAY MELANGE",
    "WOODLAND GRAY MELANGE": "WOODLAND GRAY MELANGE",
    "GRM": "GRAY RIDGE MELANGE",
    "GRAY RIDGE MELANGE": "GRAY RIDGE MELANGE",
    "BMM": "BLUE STONE MELANGE",
    "BLUE STONE MELANGE": "BLUE STONE MELANGE",
    "STE": "SKY TRACE",
    "SKY TRACE": "SKY TRACE",
    "MTE": "MOON TRACE",
    "MOON TRACE": "MOON TRACE",
    "GTE": "GREEN TRACE",
    "GREEN TRACE": "GREEN TRACE",
    "REH": "RIDGE HUSH",
    "RIDGE HUSH": "RIDGE HUSH",
    "DKH": "DARK HUSH",
    "DARK HUSH": "DARK HUSH",
    "NYH": "NAVY HUSH",
    "NAVY HUSH": "NAVY HUSH",
    "GSL": "GRAY SWIRL",
    "GREY SWIRL": "GRAY SWIRL",
    "GRAY SWIRL": "GRAY SWIRL",
    "MGR": "MUTED GREEN",
    "MUTED GREEN": "MUTED GREEN",
    "BCS": "BLACK CANVAS",
    "BLACK CANVAS": "BLACK CANVAS",
    "BSL": "BLUE SWIRL",
    "BLUE SWIRL": "BLUE SWIRL",
    "MYI": "MISTY INK",
    "MISTY INK": "MISTY INK",
    "MYD": "MISTY DEPTHS",
    "MISTY DEPTHS": "MISTY DEPTHS",
    "VIM": "VINTAGE INDIGO MELANGE",
    "VINTAGE INDIGO MELANGE": "VINTAGE INDIGO MELANGE",
    "NCN": "NEW CARBON",
    "NEW CARBON": "NEW CARBON",
    "WFS": "WIND FORTRESS",
    "WIND FORTRESS": "WIND FORTRESS",
    "DFS": "DARK FORTRESS",
    "DARK FORTRESS": "DARK FORTRESS",
    "CTS": "COBALT SQUADRON",
    "COBALT SQUADRON": "COBALT SQUADRON",
    "SYP": "SMOKY PLATOON",
    "SMOKY PLATOON": "SMOKY PLATOON",
    "SSN": "SHADOW SQUADRON",
    "SHADOW SQUADRON": "SHADOW SQUADRON",
    "GPN": "GREEN PLATOON",
    "GREEN PLATOON": "GREEN PLATOON",
    "BTG": "BLACK_TREKKING GREEN",
    "BLACK_TREKKING GREEN": "BLACK_TREKKING GREEN",
    "DPM": "DUSTY PINK MELANGE",
    "DUSTY PINK MELANGE": "DUSTY PINK MELANGE",
    "OLM": "OMPHALODES MELANGE",
    "OMPHALODES MELANGE": "OMPHALODES MELANGE",
    "BDA": "BLUE DELTA",
    "BLUE DELTA": "BLUE DELTA",
    "GDA": "GREEN DELTA",
    "GREEN DELTA": "GREEN DELTA",
    "PMT": "PURPLE MIST",
    "PURPLE MIST": "PURPLE MIST",
    "BRE": "BLUE MIRAGE",
    "BLUE MIRAGE": "BLUE MIRAGE",
    "BKE": "BLACK MIRAGE",
    "BLACK MIRAGE": "BLACK MIRAGE",
    "PSL": "PINKY SWIRL",
    "PINKY SWIRL": "PINKY SWIRL",
    "BRR": "BLUEY RADIANT RIPPLE",
    "BLUEY RADIANT RIPPLE": "BLUEY RADIANT RIPPLE",
    "GRR": "GREENY RADIANT RIPPLE",
    "GREENY RADIANT RIPPLE": "GREENY RADIANT RIPPLE",
    "BLW": "BLUEY LUMIWAVE",
    "BLUEY LUMIWAVE": "BLUEY LUMIWAVE",
    "BKW": "BLACKY LUMIWAVE",
    "BLACKY LUMIWAVE": "BLACKY LUMIWAVE",
    "BLA": "BLACK ASTRONAUT",
    "BLACK ASTRONAUT": "BLACK ASTRONAUT",
    "BBL": "BLACK BASKET BALL",
    "BLACK BASKET BALL": "BLACK BASKET BALL",
    "BKS": "BLACK SHUTTLE",
    "BLACK SHUTTLE": "BLACK SHUTTLE",
    "BKF": "BLACK FORWARD",
    "BLACK FORWARD": "BLACK FORWARD",
    "BKB": "BLACK BELIEVE",
    "BLACK BELIEVE": "BLACK BELIEVE",
    "CHA": "CHOCOLATE A",
    "CHOCOLATE A": "CHOCOLATE A",
    "PLA": "PURPLE A",
    "PURPLE A": "PURPLE A",
    "BSU": "BLACK SUMMIT",
    "BLACK SUMMIT": "BLACK SUMMIT",
    "BKA": "BLACK ASCENT",
    "BLACK ASCENT": "BLACK ASCENT",
    "CNA": "CARBON GRAY ASH",
    "CARBON GREY ASH": "CARBON GRAY ASH",
    "CARBON GRAY ASH": "CARBON GRAY ASH",
    "CNR": "CARBON GRAY GRAY RIDGE",
    "CARBON GREY GRAY RIDGE": "CARBON GRAY GRAY RIDGE",
    "CARBON GRAY GRAY RIDGE": "CARBON GRAY GRAY RIDGE",
    "NBO": "NAVY B OMPHALODES",
    "NAVY B OMPHALODES": "NAVY B OMPHALODES",
    "NBV": "NAVY B VINTAGE INDIGO",
    "NAVY B VINTAGE INDIGO": "NAVY B VINTAGE INDIGO",
    "BAM": "BLACK MYRTLE",
    "BLACK MYRTLE": "BLACK MYRTLE",
    "BMA": "BLACK MARSALA",
    "BLACK MARSALA": "BLACK MARSALA",
    "BSO": "BLACK_SMOKEY OLIVE",
    "BLACK_SMOKEY OLIVE": "BLACK_SMOKEY OLIVE",
    "CME": "CHOCOLATE MELANGE",
    "CHOCOLATE MELANGE": "CHOCOLATE MELANGE",
    "MGM": "MINT GREEN MELANGE",
    "MINT GREEN MELANGE": "MINT GREEN MELANGE",
    "LAM": "LILAC MELANGE",
    "LILAC MELANGE": "LILAC MELANGE",
    "YLM": "YELLOW MELANGE",
    "YELLOW MELANGE": "YELLOW MELANGE",
    "PHM": "PEACH MELANGE",
    "PEACH MELANGE": "PEACH MELANGE",
    "RLM": "ROYAL BLUE MELANGE",
    "ROYAL BLUE MELANGE": "ROYAL BLUE MELANGE",
    "PLM": "PURPLE A MELANGE",
    "PURPLE A MELANGE": "PURPLE A MELANGE",
    "SKM": "SKY BLUE MELANGE",
    "SKY BLUE MELANGE": "SKY BLUE MELANGE",
    "SDM": "SHADOW LIME B MELANGE",
    "SHADOW LIME B MELANGE": "SHADOW LIME B MELANGE",
    "MVO": "MIDNIGHT BLUE_VINTAGE INDIGO",
    "MIDNIGHT BLUE_VINTAGE INDIGO": "MIDNIGHT BLUE_VINTAGE INDIGO",
    "AGC": "AQUA GRAY_GULF COAST B",
    "AQUA GREY_GULF COAST B": "AQUA GRAY_GULF COAST B",
    "AQUA GRAY_GULF COAST B": "AQUA GRAY_GULF COAST B",
    "HRM": "HEATHER MELANGE",
    "HEATHER MELANGE": "HEATHER MELANGE",
    "LWG": "LIGHT GRAY_WOODLAND GRAY",
    "LIGHT GREY_WOODLAND GRAY": "LIGHT GRAY_WOODLAND GRAY",
    "LIGHT GRAY_WOODLAND GRAY": "LIGHT GRAY_WOODLAND GRAY",
    "VNB": "VINTAGE INDIGO-NAVY B",
    "VINTAGE INDIGO-NAVY B": "VINTAGE INDIGO-NAVY B",
    "AWG": "ASH_WOODLAND GRAY",
    "ASH_WOODLAND GRAY": "ASH_WOODLAND GRAY",
    "WCB": "WHITE_CAPRI BLUE",
    "WHITE_CAPRI BLUE": "WHITE_CAPRI BLUE",
    "BBR": "BLACK_BERRY RED",
    "BLACK_BERRY RED": "BLACK_BERRY RED",
    "CBL": "CARBON_BLACK",
    "CARBON_BLACK": "CARBON_BLACK",
    "BWG": "BLACK_WOODLAND GRAY",
    "BLACK_WOODLAND GRAY": "BLACK_WOODLAND GRAY",
    "BGT": "BLACK_GULF COAST B",
    "BLACK_GULF COAST B": "BLACK_GULF COAST B",
    "KBL": "KHAKI_BLACK",
    "KHAKI_BLACK": "KHAKI_BLACK",
    "LNB": "LIGHT GRAY B_NAVY B",
    "LIGHT GREY B_NAVY B": "LIGHT GRAY B_NAVY B",
    "LIGHT GRAY B_NAVY B": "LIGHT GRAY B_NAVY B",
    "NLB": "NAVY B_LT GRAY B",
    "NAVY B_LT GREY B": "NAVY B_LT GRAY B",
    "NAVY B_LT GRAY B": "NAVY B_LT GRAY B"
}



def normalize_color(x):
    if not x:
        return ""
    c = str(x).strip().upper().replace("GREY", "GRAY").replace("_", " ").replace("-", " ")
    descriptors = [" PRO", " NEO", " PLUS", " PREMIUM", " LITE", " MAX", " ULTRA", " SPORT", " ACTIVE", " EDITION", " SERIES", " FIT", " CLASSIC", " FLEX", " PRIME", " STUDIO", " COLLECTION", " LINE", " AIR", " TECH", " DRY"]
    for suffix in descriptors:
        if c.endswith(suffix):
            c = c[:-len(suffix)].strip()
    res = color_map.get(c, c)
    res_str = str(res).strip().upper().replace("GREY", "GRAY").replace("_", " ").replace("-", " ")
    for var_suffix in [" A", " B", " C", " D"]:
        if res_str.endswith(var_suffix):
            res_str = res_str[:-len(var_suffix)].strip()
    return res_str


def normalize_number(x):
    if x is None or x == "":
        return None
    s = str(x).replace(",", "").replace("₹", "").strip()
    m = re.search(r"[\d,]+\.?\d*", s)
    if m:
        try:
            return round(float(m.group().replace(",", "")), 2)
        except ValueError:
            pass
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


def clean_style_for_gsheet(style):
    if not style:
        return ""
    s = str(style).strip().upper()
    if s.startswith("OR"):
        s = s[1:]
    elif s.startswith("SOR"):
        s = s[2:]
    return s


def find_mrp_by_style_digits(style_clean, df, col_name, tag_type="Standard Garment / Dress Tags"):
    s_clean = str(style_clean).strip().upper()
    if "/" in s_clean:
        s_clean = s_clean.split("/")[0].strip()
    s_digits = "".join([c for c in s_clean if c.isdigit()])
    if not s_digits:
        return None

    # 1. Exact style match first
    match = df[df[col_name].astype(str).str.strip().str.upper() == style_clean]
    if not match.empty:
        mrp_val = match.iloc[0].get("MRP")
        if pd.notna(mrp_val):
            return mrp_val

    # 2. Base style match (e.g. split by /)
    match = df[df[col_name].astype(str).str.strip().str.upper() == s_clean]
    if not match.empty:
        mrp_val = match.iloc[0].get("MRP")
        if pd.notna(mrp_val):
            return mrp_val

    # 3. Digit-based lookup fallback (ONLY for B2B Box Stickers!)
    if tag_type == "B2B Box Sticker tag file":
        for _, row in df.iterrows():
            gs_style = str(row.get(col_name, "")).strip().upper()
            if "/" in gs_style:
                gs_style = gs_style.split("/")[0].strip()
            gs_digits = "".join([c for c in gs_style if c.isdigit()])
            if gs_digits == s_digits:
                mrp_val = row.get("MRP")
                if pd.notna(mrp_val):
                    return mrp_val
    return None


def get_updated_mrp(pdf_style, pdf_sku, gsheet_dfs, tag_type="Standard Garment / Dress Tags"):
    if not gsheet_dfs:
        return None

    style_clean = str(pdf_style).strip().upper() if pdf_style else ""
    sku_clean = str(pdf_sku).strip().upper() if pdf_sku else ""

    if not style_clean and sku_clean:
        extracted_style, _, _ = extract_sku_details(sku_clean)
        style_clean = extracted_style if extracted_style else ""

    style_clean = clean_style_for_gsheet(style_clean)

    # 1. Search in DT FINAL MRP (matching against Column H (8th column, index 7))
    df_dt = gsheet_dfs.get("DT FINAL MRP")
    if df_dt is not None and len(df_dt.columns) > 7:
        col_h = df_dt.columns[7]
        mrp_val = find_mrp_by_style_digits(style_clean, df_dt, col_h, tag_type)
        if mrp_val is not None:
            return mrp_val

    # 2. Search in New MRP 26-27 (matching against Column I (9th column, index 8))
    df_new = gsheet_dfs.get("New MRP 26-27")
    if df_new is not None and len(df_new.columns) > 8:
        col_i = df_new.columns[8]
        mrp_val = find_mrp_by_style_digits(style_clean, df_new, col_i, tag_type)
        if mrp_val is not None:
            return mrp_val

    return None


def get_updated_description(pdf_style, pdf_sku, gsheet_dfs, tag_type="Standard Garment / Dress Tags"):
    if not gsheet_dfs:
        return None

    style_clean = str(pdf_style).strip().upper() if pdf_style else ""
    sku_clean = str(pdf_sku).strip().upper() if pdf_sku else ""

    if not style_clean and sku_clean:
        extracted_style, _, _ = extract_sku_details(sku_clean)
        style_clean = extracted_style if extracted_style else ""

    style_clean = clean_style_for_gsheet(style_clean)

    # 1. Search in DT FINAL MRP
    df_dt = gsheet_dfs.get("DT FINAL MRP")
    if df_dt is not None and len(df_dt.columns) > 7:
        col_h = df_dt.columns[7]
        s_clean = style_clean.split("/")[0].strip() if "/" in style_clean else style_clean
        s_digits = "".join([c for c in s_clean if c.isdigit()])
        
        # Exact/Base Match first
        match = df_dt[df_dt[col_h].astype(str).str.strip().str.upper() == style_clean]
        if not match.empty:
            desc = match.iloc[0].get("DESCRIPTION")
            if pd.notna(desc):
                return str(desc).strip()
        match = df_dt[df_dt[col_h].astype(str).str.strip().str.upper() == s_clean]
        if not match.empty:
            desc = match.iloc[0].get("DESCRIPTION")
            if pd.notna(desc):
                return str(desc).strip()
                
        # Digit match (B2B only)
        if tag_type == "B2B Box Sticker tag file" and s_digits:
            for _, row in df_dt.iterrows():
                gs_style = str(row.get(col_h, "")).strip().upper()
                if "/" in gs_style:
                    gs_style = gs_style.split("/")[0].strip()
                gs_digits = "".join([c for c in gs_style if c.isdigit()])
                if gs_digits == s_digits:
                    desc = row.get("DESCRIPTION")
                    if pd.notna(desc):
                        return str(desc).strip()

    # 2. Search in New MRP 26-27
    df_new = gsheet_dfs.get("New MRP 26-27")
    if df_new is not None and len(df_new.columns) > 8:
        col_i = df_new.columns[8]
        s_clean = style_clean.split("/")[0].strip() if "/" in style_clean else style_clean
        s_digits = "".join([c for c in s_clean if c.isdigit()])
        
        # Exact/Base Match first
        match = df_new[df_new[col_i].astype(str).str.strip().str.upper() == style_clean]
        if not match.empty:
            desc = match.iloc[0].get("DESCRIPTION")
            if pd.notna(desc):
                return str(desc).strip()
        match = df_new[df_new[col_i].astype(str).str.strip().str.upper() == s_clean]
        if not match.empty:
            desc = match.iloc[0].get("DESCRIPTION")
            if pd.notna(desc):
                return str(desc).strip()
                
        # Digit match (B2B only)
        if tag_type == "B2B Box Sticker tag file" and s_digits:
            for _, row in df_new.iterrows():
                gs_style = str(row.get(col_i, "")).strip().upper()
                if "/" in gs_style:
                    gs_style = gs_style.split("/")[0].strip()
                gs_digits = "".join([c for c in gs_style if c.isdigit()])
                if gs_digits == s_digits:
                    desc = row.get("DESCRIPTION")
                    if pd.notna(desc):
                        return str(desc).strip()

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
        elif sku[-6:-3] in apparel_sizes:
            style = sku[2:-9]
            if len(style) >= 3 and style[1:3] == "OR":
                style = style[1:]
            color = sku[-9:-6]
            size = sku[-6:-3]
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


def parse_product_name_info(prod_name_str):
    if not prod_name_str:
        return "", "", "", None
    parts = str(prod_name_str).strip().split()
    lot_no = parts[0] if parts else ""
    base_style = lot_no.split("/")[0] if "/" in lot_no else lot_no

    pcs_match = re.search(r"(\d+)\s*PCS", str(prod_name_str), re.IGNORECASE)
    pcs_qty = int(pcs_match.group(1)) if pcs_match else None

    desc = " ".join(parts[1:])
    if pcs_match:
        desc = desc.split(pcs_match.group(0))[0].strip()
    desc = re.sub(r"\s+(ASSORTED|BLACK|NAVY|GREY|GRAY|WHITE)$", "", desc, flags=re.IGNORECASE).strip()

    return lot_no, base_style, desc, pcs_qty


def extract_style_and_size_from_sku(sku_str):
    style, color, size = extract_sku_details(sku_str)
    return style, size





def compare(pdf_df: pd.DataFrame, excel_df: pd.DataFrame, gsheet_dfs: dict, tag_type: str = "D2C Dress tag file") -> pd.DataFrame:
    desc_col = find_col(excel_df, "PRODUCT NAME", "DESCRIPTION", "PRODUCT")
    lot_col = find_col(excel_df, "LOT NO", "LOT", "STYLE", "STYLE CODE")
    sku_col = find_col(excel_df, "SKU CODE", "SKU")
    barcode_col = find_col(excel_df, "BARCODE", "BAR CODE", "EAN", "GTIN")
    mrp_col = find_col(excel_df, "MRP")
    total_mrp_col = find_col(excel_df, "TOTAL MRP")
    size_col = find_col(excel_df, "SIZE")
    color_col = find_col(excel_df, "COLOUR", "COLOR")
    qty_col = find_col(excel_df, "PACK QTY", "NET QTY", "QTY", "TAG QTY")

    if sku_col is None:
        raise ValueError("Could not find an SKU column in the Excel sheet.")

    excel_idx = {normalize_sku(row[sku_col]): row for _, row in excel_df.iterrows()}

    if tag_type == "B2B Box Sticker tag file":
        field_map = [
            ("Description", desc_col, normalize_text),
            ("Lot No", lot_col, normalize_text),
            ("Qty", qty_col, normalize_number),
            ("Total MRP", total_mrp_col, normalize_number),
            ("SKU", sku_col, normalize_sku),
            ("EAN", barcode_col, normalize_text),
            ("Size", size_col, normalize_size),
        ]
    else:
        field_map = [
            ("Description", desc_col, normalize_text),
            ("Lot No", lot_col, normalize_text),
            ("Qty", qty_col, normalize_number),
            ("MRP", mrp_col, normalize_number),
            ("Total MRP", total_mrp_col, normalize_number),
            ("SKU", sku_col, normalize_sku),
            ("EAN", barcode_col, normalize_text),
            ("Size", size_col, normalize_size),
            ("Color", color_col, normalize_color),
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

        prod_name_val = excel_row.get(desc_col) if desc_col else None
        lot_info, base_style_info, desc_info, pack_qty_info = parse_product_name_info(prod_name_val)

        for field_name, excel_col, norm_fn in field_map:
            pdf_val = tag.get(field_name)
            excel_val = None

            if field_name == "Description":
                pdf_val = tag.get("Description") or tag.get("Product")
                if not pdf_val:
                    pdf_val = desc_info
                excel_val = desc_info if desc_info else (excel_row.get(excel_col) if excel_col else None)
            elif field_name == "Lot No":
                pdf_val = tag.get("Lot No") or tag.get("Style")
                excel_val = lot_info if lot_info else (excel_row.get(excel_col) if excel_col else None)
            elif field_name == "Qty":
                pdf_val = tag.get("Net Quantity") or tag.get("Qty")
                excel_val = pack_qty_info if pack_qty_info else (excel_row.get(excel_col) if excel_col else None)
            elif field_name == "MRP":
                excel_val = get_updated_mrp(tag.get("Style") or base_style_info, tag.get("SKU"), gsheet_dfs, tag_type=tag_type)
                if excel_val is None:
                    excel_val = excel_row.get(excel_col) if excel_col else None
            elif field_name == "Total MRP":
                pdf_val = tag.get("Total MRP")
                if pdf_val is None:
                    continue
                single_mrp = get_updated_mrp(tag.get("Style") or base_style_info, tag.get("SKU"), gsheet_dfs, tag_type=tag_type)
                if single_mrp is None and mrp_col:
                    single_mrp = excel_row.get(mrp_col)
                p_qty = norm_fn(tag.get("Net Quantity") or tag.get("Qty")) or pack_qty_info or 1
                if single_mrp and p_qty:
                    try:
                        excel_val = float(single_mrp) * float(p_qty)
                    except (ValueError, TypeError):
                        excel_val = None
                else:
                    excel_val = excel_row.get(excel_col) if excel_col else None
            elif field_name == "Size":
                excel_sku = excel_row.get(sku_col)
                _, _, extracted_size = extract_sku_details(excel_sku)
                excel_val = format_size_as_tag(extracted_size) if extracted_size else None
                if not pdf_val:
                    pdf_val = excel_val
            elif field_name == "Color":
                if excel_col and pd.notna(excel_row.get(excel_col)):
                    excel_val = excel_row.get(excel_col)
                else:
                    excel_sku = excel_row.get(sku_col)
                    _, extracted_color, _ = extract_sku_details(excel_sku)
                    excel_val = color_map.get(extracted_color, extracted_color) if extracted_color else None
            elif field_name == "EAN":
                pdf_val = tag.get("EAN") or tag.get("Barcode")
                excel_val = excel_row.get(barcode_col) if barcode_col else None
            else:
                if excel_col is None:
                    continue
                excel_val = excel_row.get(excel_col)

            pdf_norm = norm_fn(pdf_val)
            excel_norm = norm_fn(excel_val)

            if field_name == "Description":
                p_words = set(re.findall(r"\w+", str(pdf_norm).upper()))
                e_words = set(re.findall(r"\w+", str(excel_norm).upper()))
                common_words = p_words.intersection(e_words)
                is_match = (
                    pdf_norm == excel_norm
                    or len(common_words) >= 2
                    or (bool(pdf_norm) and bool(excel_norm) and (
                        pdf_norm in excel_norm
                        or excel_norm in pdf_norm
                    ))
                )
                if not is_match:
                    g_desc = get_updated_description(tag.get("Style") or base_style_info, tag.get("SKU"), gsheet_dfs, tag_type=tag_type)
                    if g_desc:
                        g_norm = norm_fn(g_desc)
                        g_words = set(re.findall(r"\w+", str(g_norm).upper()))
                        common_g_words = p_words.intersection(g_words)
                        if (
                            pdf_norm == g_norm
                            or len(common_g_words) >= 2
                            or (bool(pdf_norm) and bool(g_norm) and (
                                pdf_norm in g_norm
                                or g_norm in pdf_norm
                            ))
                        ):
                            is_match = True
                            excel_val = g_desc
                            excel_norm = g_norm
                status = "✅ Match" if is_match else "❌ Mismatch"
            elif field_name == "Lot No":
                p_b = str(pdf_norm).split("/")[0] if "/" in str(pdf_norm) else str(pdf_norm)
                e_b = str(excel_norm).split("/")[0] if "/" in str(excel_norm) else str(excel_norm)
                is_match = (pdf_norm == excel_norm or p_b == e_b)
                status = "✅ Match" if is_match else "❌ Mismatch"
            elif field_name == "Color":
                is_match = (
                    pdf_norm == excel_norm
                    or (bool(pdf_norm) and bool(excel_norm) and (
                        pdf_norm.startswith(excel_norm)
                        or excel_norm.startswith(pdf_norm)
                        or pdf_norm in excel_norm
                        or excel_norm in pdf_norm
                    ))
                )
                status = "✅ Match" if is_match else "❌ Mismatch"
            else:
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

    tag_type = "B2B Box Sticker tag file" if "b2b" in os.path.basename(pdf_path).lower() or "box" in os.path.basename(pdf_path).lower() else "D2C Dress tag file"
    report_df = compare(pdf_df, excel_df, gsheet_dfs, tag_type=tag_type)

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
