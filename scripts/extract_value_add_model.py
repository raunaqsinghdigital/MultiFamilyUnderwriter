#!/usr/bin/env python3
"""Extract Value Add workbook model from XLS file.

Reads data/09_MFM_Property_Analyzer_VALUE_ADD_Rev1_-_11212_124_ST.xls and
outputs two files with the same schema used by the Buy & Hold model:
  - data/value_add_model.json
  - data/value_add_formulas_expanded.json

The XLS is encrypted with the default VelvetSweatshop password (Excel worksheet
protection, not file encryption). This script handles decryption automatically.

Usage:
    python3 scripts/extract_value_add_model.py
    python3 scripts/extract_value_add_model.py --xls path/to/file.xls
    python3 scripts/extract_value_add_model.py --debug
"""

from __future__ import annotations

import argparse
import io
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT_DIR / "data"

DEFAULT_XLS = DATA_DIR / "09_MFM_Property_Analyzer_VALUE_ADD_Rev1_-_11212_124_ST.xls"
OUTPUT_MODEL = DATA_DIR / "value_add_model.json"
OUTPUT_FORMULAS = DATA_DIR / "value_add_formulas_expanded.json"

# Sheets visible to analysts (in tab order)
ANALYST_SHEETS = ["Rent Roll", "Valuation", "Refinance", "Return"]
# Sheets hidden from analysts (admin/internal only)
HIDDEN_SHEETS = {"Mortgage", "1.1", "IO Calculator"}

# Input cells in Value Add have this RGB fill colour (light cream yellow).
# This matches the "ENTER DATA IN SHADED CELLS" convention in the template.
INPUT_FILL_RGB = (255, 255, 204)

# ─── Cells to exclude from input detection (labels / auto-filled from rent roll) ──
EXCLUDE_INPUT_KEYS: set[str] = {
    "Valuation!K6",   # Label: "ENTER DATA IN SHADED CELLS"
    "Refinance!K6",   # Same label
    "Valuation!G10",  # Monthly rent total — auto-populated from Rent Roll state
    "Refinance!G10",  # Projected rent total — auto-populated from Rent Roll state
}

# ─── Explicit Value Add formulas ──────────────────────────────────────────────
# xlrd cannot extract formula strings from .xls files — it returns evaluated values.
# We define the key formulas manually here, matching the Excel logic exactly.
# Percentages in Value Add are stored as decimals (e.g. 0.06 for 6%).
#
# Sheet structure (Value Add differs from Buy & Hold in row numbers):
#   Rent Roll  : B6:C12  = unit / regular rent  |  H6:H12 = projected rent
#   Valuation  : D6=units, D7=price, D15=vacancy, E19=taxes, F20-F24=OpEx/unit
#                D25=mgmt%, D26=other%, E30=cap%, E33=rate, E34=amort, E35=LTV, E36=DSCR
#   Return     : Financing + cash flow + ROI projections

EXPLICIT_FORMULAS: dict[str, str] = {

    # ── Valuation — Rental Revenue ────────────────────────────────────────────
    # G10 = monthly rent sum from Rent Roll (pre-populated, driven by rent roll state)
    # E10 = annual rent = G10 × 12
    "Valuation!E10": "Valuation!G10 * 12",
    # F10 = per unit per year
    "Valuation!F10": "IF(Valuation!D6 > 0, Valuation!E10 / Valuation!D6, 0)",
    # E11 = laundry (user input, left as-is)
    "Valuation!F11": "IF(Valuation!D6 > 0, Valuation!E11 / Valuation!D6, 0)",
    "Valuation!G11": "IF(Valuation!D6 > 0, Valuation!E11 / Valuation!D6 / 12, 0)",
    # E12 = parking (none in this model, formula produces 0)
    # E13 = other (none in this model)
    # E14 = PGI = sum of all revenue
    "Valuation!E14": "SUM(Valuation!E10,Valuation!E11,Valuation!E12,Valuation!E13)",
    "Valuation!F14": "IF(Valuation!D6 > 0, Valuation!E14 / Valuation!D6, 0)",
    "Valuation!G14": "Valuation!E14 / 12",

    # ── Valuation — Vacancy ───────────────────────────────────────────────────
    # D15 = vacancy rate (decimal), E15 = vacancy loss
    "Valuation!E15": "Valuation!E10 * Valuation!D15",
    "Valuation!F15": "IF(Valuation!D6 > 0, Valuation!E15 / Valuation!D6, 0)",
    "Valuation!G15": "Valuation!E15 / 12",

    # ── Valuation — Effective Gross Income (EGI) ──────────────────────────────
    "Valuation!E16": "Valuation!E14 - Valuation!E15",
    "Valuation!F16": "IF(Valuation!D6 > 0, Valuation!E16 / Valuation!D6, 0)",
    "Valuation!G16": "Valuation!E16 / 12",

    # ── Valuation — Operating Expenses ───────────────────────────────────────
    # Property Taxes (E19) = user input; F/G derived
    "Valuation!F19": "IF(Valuation!D6 > 0, Valuation!E19 / Valuation!D6, 0)",
    "Valuation!G19": "Valuation!E19 / 12",
    "Valuation!H19": "IF(Valuation!E16 > 0, Valuation!E19 / Valuation!E16, 0)",
    # Insurance: F20 = $/unit/year (input), E20 = total, G20 = monthly
    "Valuation!E20": "Valuation!F20 * Valuation!D6",
    "Valuation!G20": "Valuation!E20 / 12",
    "Valuation!H20": "IF(Valuation!E16 > 0, Valuation!E20 / Valuation!E16, 0)",
    # Utilities
    "Valuation!E21": "Valuation!F21 * Valuation!D6",
    "Valuation!G21": "Valuation!E21 / 12",
    "Valuation!H21": "IF(Valuation!E16 > 0, Valuation!E21 / Valuation!E16, 0)",
    # Repairs & Maintenance
    "Valuation!E22": "Valuation!F22 * Valuation!D6",
    "Valuation!G22": "Valuation!E22 / 12",
    "Valuation!H22": "IF(Valuation!E16 > 0, Valuation!E22 / Valuation!E16, 0)",
    # Appliances
    "Valuation!E23": "Valuation!F23 * Valuation!D6",
    "Valuation!G23": "Valuation!E23 / 12",
    "Valuation!H23": "IF(Valuation!E16 > 0, Valuation!E23 / Valuation!E16, 0)",
    # Wages
    "Valuation!E24": "Valuation!F24 * Valuation!D6",
    "Valuation!G24": "Valuation!E24 / 12",
    "Valuation!H24": "IF(Valuation!E16 > 0, Valuation!E24 / Valuation!E16, 0)",
    # Management fee: D25 = % of EGI (decimal), E25 = total
    "Valuation!E25": "Valuation!D25 * Valuation!E16",
    "Valuation!F25": "IF(Valuation!D6 > 0, Valuation!E25 / Valuation!D6, 0)",
    "Valuation!G25": "Valuation!E25 / 12",
    "Valuation!H25": "Valuation!D25",
    # Other/Advertising: D26 = % of EGI (decimal)
    "Valuation!E26": "Valuation!D26 * Valuation!E16",
    "Valuation!F26": "IF(Valuation!D6 > 0, Valuation!E26 / Valuation!D6, 0)",
    "Valuation!G26": "Valuation!E26 / 12",
    "Valuation!H26": "Valuation!D26",
    # Total Operating Expense
    "Valuation!E27": "SUM(Valuation!E19,Valuation!E20,Valuation!E21,Valuation!E22,Valuation!E23,Valuation!E24,Valuation!E25,Valuation!E26)",
    "Valuation!F27": "IF(Valuation!D6 > 0, Valuation!E27 / Valuation!D6, 0)",
    "Valuation!G27": "Valuation!E27 / 12",
    "Valuation!H27": "IF(Valuation!E16 > 0, Valuation!E27 / Valuation!E16, 0)",

    # ── Valuation — Net Operating Income (NOI) ────────────────────────────────
    "Valuation!E29": "Valuation!E16 - Valuation!E27",
    "Valuation!F29": "IF(Valuation!D6 > 0, Valuation!E29 / Valuation!D6, 0)",
    "Valuation!G29": "Valuation!E29 / 12",

    # ── Valuation — Market Valuation ─────────────────────────────────────────
    # E30 = market cap rate (decimal input), E31 = value based on cap rate
    "Valuation!E31": "IF(Valuation!E30 > 0, Valuation!E29 / Valuation!E30, 0)",
    "Valuation!F31": "IF(Valuation!D6 > 0, Valuation!E31 / Valuation!D6, 0)",

    # ── Valuation — Debt Service ──────────────────────────────────────────────
    # Canadian mortgage: semi-annual compounding (Interest Act)
    # Monthly rate = (1 + annual_rate/2)^(1/6) - 1
    # Annual DS (IO) = loan × 12 × monthly_rate
    "Valuation!E37": "IF(Valuation!E45 > 0, Valuation!E45 * 12 * (POWER((1 + Valuation!E33 / 2), (1 / 6)) - 1), 0)",
    # Annual DS (P&I) — full amortization formula
    "Valuation!E38": "IF(Valuation!E45 > 0, Valuation!E45 * (POWER((1 + Valuation!E33 / 2), (1 / 6)) - 1) / (1 - POWER((1 + (POWER((1 + Valuation!E33 / 2), (1 / 6)) - 1)), -(Valuation!E34 * 12))) * 12, 0)",
    # DSCR using P&I
    "Valuation!E39": "IF(Valuation!E38 > 0, Valuation!E29 / Valuation!E38, 0)",
    # Available for debt service = NOI / required DSCR
    "Valuation!E40": "IF(Valuation!E36 > 0, Valuation!E29 / Valuation!E36, 0)",

    # ── Valuation — Three Maximum Loan Tests ──────────────────────────────────
    # Max loan by LTV × cap rate value
    "Valuation!E41": "Valuation!E35 * Valuation!E31",
    # Max loan by LTV × purchase price
    "Valuation!E42": "Valuation!E35 * Valuation!D7",
    # Max loan by payment capacity (PV of available DS at given rate/amort)
    "Valuation!E43": "IF(Valuation!E33 > 0, Valuation!E40 / 12 * (1 - POWER((1 + (POWER((1 + Valuation!E33 / 2), (1 / 6)) - 1)), -(Valuation!E34 * 12))) / (POWER((1 + Valuation!E33 / 2), (1 / 6)) - 1), 0)",
    # Lesser of three
    "Valuation!E45": "MIN(Valuation!E41, Valuation!E42, Valuation!E43)",
    # Actual LTC = loan / purchase price
    "Valuation!E46": "IF(Valuation!D7 > 0, Valuation!E45 / Valuation!D7, 0)",
    # Actual DS (P&I) on the approved loan amount
    "Valuation!E47": "IF(Valuation!E45 > 0, Valuation!E45 * (POWER((1 + Valuation!E33 / 2), (1 / 6)) - 1) / (1 - POWER((1 + (POWER((1 + Valuation!E33 / 2), (1 / 6)) - 1)), -(Valuation!E34 * 12))) * 12, 0)",
    # Actual DSCR
    "Valuation!E48": "IF(Valuation!E47 > 0, Valuation!E29 / Valuation!E47, 0)",
    # Property cap rate = NOI / purchase price
    "Valuation!E49": "IF(Valuation!D7 > 0, Valuation!E29 / Valuation!D7, 0)",

    # ── Valuation — Total Funds Needed ────────────────────────────────────────
    "Valuation!E51": "Valuation!D7 - Valuation!E45",   # Downpayment
    # E52 = closing costs (summed from Return sheet or static)
    # E53 = CapEx budget (summed from budget lines below)
    "Valuation!E54": "SUM(Valuation!E51, Valuation!E52, Valuation!E53)",

    # ── Rent Roll — Totals ────────────────────────────────────────────────────
    "Rent Roll!G15": "SUM(Rent Roll!G6,Rent Roll!G7,Rent Roll!G8,Rent Roll!G9,Rent Roll!G10,Rent Roll!G11,Rent Roll!G12)",
    "Rent Roll!H15": "SUM(Rent Roll!H6,Rent Roll!H7,Rent Roll!H8,Rent Roll!H9,Rent Roll!H10,Rent Roll!H11,Rent Roll!H12)",

    # ── Valuation — G10 auto-filled from Rent Roll!G15 (current rents) ────────
    "Valuation!G10": "Rent Roll!G15",

    # ── Refinance — parallel to Valuation but for stabilized (post-reno) rents ──
    # G10 auto-filled from Rent Roll!H15 (projected rents)
    "Refinance!G10": "Rent Roll!H15",
    # Annual rent = projected monthly × 12
    "Refinance!E10": "Refinance!G10 * 12",
    "Refinance!F10": "IF(Refinance!D6 > 0, Refinance!E10 / Refinance!D6, 0)",
    # Laundry (E11 if present)
    "Refinance!F11": "IF(Refinance!D6 > 0, Refinance!E11 / Refinance!D6, 0)",
    "Refinance!G11": "IF(Refinance!D6 > 0, Refinance!E11 / Refinance!D6 / 12, 0)",
    # Other income (E12)
    "Refinance!F12": "IF(Refinance!D6 > 0, Refinance!E12 / Refinance!D6, 0)",
    "Refinance!G12": "Refinance!E12 / 12",
    # PGI
    "Refinance!E14": "SUM(Refinance!E10,Refinance!E11,Refinance!E12,Refinance!E13)",
    "Refinance!F14": "IF(Refinance!D6 > 0, Refinance!E14 / Refinance!D6, 0)",
    "Refinance!G14": "Refinance!E14 / 12",
    # Vacancy
    "Refinance!E15": "Refinance!E10 * Refinance!D15",
    "Refinance!F15": "IF(Refinance!D6 > 0, Refinance!E15 / Refinance!D6, 0)",
    "Refinance!G15": "Refinance!E15 / 12",
    # EGI
    "Refinance!E16": "Refinance!E14 - Refinance!E15",
    "Refinance!F16": "IF(Refinance!D6 > 0, Refinance!E16 / Refinance!D6, 0)",
    "Refinance!G16": "Refinance!E16 / 12",
    # Operating Expenses (same structure as Valuation)
    "Refinance!F19": "IF(Refinance!D6 > 0, Refinance!E19 / Refinance!D6, 0)",
    "Refinance!G19": "Refinance!E19 / 12",
    "Refinance!H19": "IF(Refinance!E16 > 0, Refinance!E19 / Refinance!E16, 0)",
    "Refinance!E20": "Refinance!F20 * Refinance!D6",
    "Refinance!G20": "Refinance!E20 / 12",
    "Refinance!H20": "IF(Refinance!E16 > 0, Refinance!E20 / Refinance!E16, 0)",
    "Refinance!E21": "Refinance!F21 * Refinance!D6",
    "Refinance!G21": "Refinance!E21 / 12",
    "Refinance!H21": "IF(Refinance!E16 > 0, Refinance!E21 / Refinance!E16, 0)",
    "Refinance!E22": "Refinance!F22 * Refinance!D6",
    "Refinance!G22": "Refinance!E22 / 12",
    "Refinance!H22": "IF(Refinance!E16 > 0, Refinance!E22 / Refinance!E16, 0)",
    "Refinance!E23": "Refinance!F23 * Refinance!D6",
    "Refinance!G23": "Refinance!E23 / 12",
    "Refinance!H23": "IF(Refinance!E16 > 0, Refinance!E23 / Refinance!E16, 0)",
    "Refinance!E24": "Refinance!F24 * Refinance!D6",
    "Refinance!G24": "Refinance!E24 / 12",
    "Refinance!H24": "IF(Refinance!E16 > 0, Refinance!E24 / Refinance!E16, 0)",
    "Refinance!E25": "Refinance!D25 * Refinance!E16",
    "Refinance!F25": "IF(Refinance!D6 > 0, Refinance!E25 / Refinance!D6, 0)",
    "Refinance!G25": "Refinance!E25 / 12",
    "Refinance!H25": "Refinance!D25",
    "Refinance!E26": "Refinance!D26 * Refinance!E16",
    "Refinance!F26": "IF(Refinance!D6 > 0, Refinance!E26 / Refinance!D6, 0)",
    "Refinance!G26": "Refinance!E26 / 12",
    "Refinance!H26": "Refinance!D26",
    "Refinance!E27": "SUM(Refinance!E19,Refinance!E20,Refinance!E21,Refinance!E22,Refinance!E23,Refinance!E24,Refinance!E25,Refinance!E26)",
    "Refinance!F27": "IF(Refinance!D6 > 0, Refinance!E27 / Refinance!D6, 0)",
    "Refinance!G27": "Refinance!E27 / 12",
    "Refinance!H27": "IF(Refinance!E16 > 0, Refinance!E27 / Refinance!E16, 0)",
    # NOI
    "Refinance!E29": "Refinance!E16 - Refinance!E27",
    "Refinance!F29": "IF(Refinance!D6 > 0, Refinance!E29 / Refinance!D6, 0)",
    "Refinance!G29": "Refinance!E29 / 12",
    # Cap rate value
    "Refinance!E31": "IF(Refinance!E30 > 0, Refinance!E29 / Refinance!E30, 0)",
    "Refinance!F31": "IF(Refinance!D6 > 0, Refinance!E31 / Refinance!D6, 0)",
    # Debt service
    "Refinance!E37": "IF(Refinance!E45 > 0, Refinance!E45 * 12 * (POWER((1 + Refinance!E33 / 2), (1 / 6)) - 1), 0)",
    "Refinance!E38": "IF(Refinance!E45 > 0, Refinance!E45 * (POWER((1 + Refinance!E33 / 2), (1 / 6)) - 1) / (1 - POWER((1 + (POWER((1 + Refinance!E33 / 2), (1 / 6)) - 1)), -(Refinance!E34 * 12))) * 12, 0)",
    "Refinance!E39": "IF(Refinance!E38 > 0, Refinance!E29 / Refinance!E38, 0)",
    "Refinance!E40": "IF(Refinance!E36 > 0, Refinance!E29 / Refinance!E36, 0)",
    # Three max loan tests
    "Refinance!E41": "Refinance!E35 * Refinance!E31",
    "Refinance!E42": "Refinance!E35 * Refinance!D7",
    "Refinance!E43": "IF(Refinance!E33 > 0, Refinance!E40 / 12 * (1 - POWER((1 + (POWER((1 + Refinance!E33 / 2), (1 / 6)) - 1)), -(Refinance!E34 * 12))) / (POWER((1 + Refinance!E33 / 2), (1 / 6)) - 1), 0)",
    "Refinance!E45": "MIN(Refinance!E41, Refinance!E42, Refinance!E43)",
    "Refinance!E46": "IF(Refinance!D7 > 0, Refinance!E45 / Refinance!D7, 0)",
    "Refinance!E47": "IF(Refinance!E45 > 0, Refinance!E45 * (POWER((1 + Refinance!E33 / 2), (1 / 6)) - 1) / (1 - POWER((1 + (POWER((1 + Refinance!E33 / 2), (1 / 6)) - 1)), -(Refinance!E34 * 12))) * 12, 0)",
    "Refinance!E48": "IF(Refinance!E47 > 0, Refinance!E29 / Refinance!E47, 0)",
    "Refinance!E49": "IF(Refinance!D7 > 0, Refinance!E29 / Refinance!D7, 0)",
    # Capital out of refinance = new loan - original loan - capex
    "Refinance!E51": "Refinance!D7 - Refinance!E45",
}


# ─── Helpers ──────────────────────────────────────────────────────────────────

def col_letter(n: int) -> str:
    """Convert 1-based column index to Excel letter (1=A, 26=Z, 27=AA…)."""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def cell_address(row: int, col: int) -> str:
    """1-based row, 1-based col → Excel address like D6."""
    return f"{col_letter(col)}{row}"


# ─── Decryption ───────────────────────────────────────────────────────────────

def decrypt_xls(path: Path) -> bytes:
    """Decrypt an XLS file protected with the default VelvetSweatshop password."""
    try:
        import msoffcrypto  # type: ignore
    except ImportError:
        print("ERROR: msoffcrypto-tool not installed. Run:")
        print("  pip3 install msoffcrypto-tool --break-system-packages")
        sys.exit(1)

    with open(path, "rb") as f:
        office = msoffcrypto.OfficeFile(f)
        office.load_key(password="VelvetSweatshop")
        buf = io.BytesIO()
        office.decrypt(buf)
        return buf.getvalue()


def open_workbook(path: Path, debug: bool = False):
    """Open XLS workbook, decrypting if necessary. Returns xlrd workbook."""
    try:
        import xlrd  # type: ignore
    except ImportError:
        print("ERROR: xlrd not installed. Run:")
        print("  pip3 install xlrd==1.2.0 --break-system-packages")
        sys.exit(1)

    raw = path.read_bytes()
    try:
        wb = xlrd.open_workbook(file_contents=raw, formatting_info=True)
        if debug:
            print("  Opened without decryption.")
        return wb
    except xlrd.XLRDError as exc:
        if "encrypted" not in str(exc).lower() and "password" not in str(exc).lower():
            print(f"ERROR reading XLS: {exc}")
            sys.exit(1)

    # Try default password
    if debug:
        print("  File is encrypted — trying VelvetSweatshop password…")
    decrypted = decrypt_xls(path)
    try:
        wb = xlrd.open_workbook(file_contents=decrypted, formatting_info=True)
        print("  Decrypted with default password (VelvetSweatshop).")
        return wb
    except xlrd.XLRDError as exc2:
        print(f"ERROR: Could not open XLS after decryption attempt: {exc2}")
        print("The file may use a custom password. Open in Excel/LibreOffice,")
        print("remove the password, and re-run this script.")
        sys.exit(1)


# ─── Fill colour detection ─────────────────────────────────────────────────────

def build_input_fill_ids(wb, debug: bool = False) -> set[int]:
    """Return XF (style) indices whose fill colour matches INPUT_FILL_RGB."""
    colour_map = wb.colour_map
    xf_list = wb.xf_list
    input_fill_ids: set[int] = set()

    for xf_idx, xf in enumerate(xf_list):
        bg = xf.background
        for ci in (bg.pattern_colour_index, bg.background_colour_index):
            rgb = colour_map.get(ci)
            if rgb == INPUT_FILL_RGB:
                input_fill_ids.add(xf_idx)
                if debug:
                    print(f"    fill_id {xf_idx} → input (rgb={rgb})")
                break

    if debug:
        print(f"  Input fill IDs detected: {sorted(input_fill_ids)}")
    return input_fill_ids


# ─── Cell extraction ──────────────────────────────────────────────────────────

def xlrd_value(sheet, row_0: int, col_0: int, wb) -> tuple[Any, str]:
    """Return (value, value_type) for a cell using xlrd (0-based indices)."""
    import xlrd  # type: ignore

    ctype = sheet.cell_type(row_0, col_0)
    raw = sheet.cell_value(row_0, col_0)

    if ctype in (0, 6):  # XL_CELL_EMPTY, XL_CELL_BLANK
        return None, "text"
    if ctype == 1:  # XL_CELL_TEXT
        stripped = str(raw).strip()
        return stripped if stripped else None, "text"
    if ctype == 2:  # XL_CELL_NUMBER
        if isinstance(raw, float) and raw == int(raw) and abs(raw) < 1e9:
            return int(raw), "number"
        return raw, "number"
    if ctype == 4:  # XL_CELL_BOOLEAN
        return bool(raw), "text"
    if ctype == 3:  # XL_CELL_DATE
        try:
            dt = xlrd.xldate_as_datetime(raw, wb.datemode)
            return dt.strftime("%Y-%m-%d"), "text"
        except Exception:
            return str(raw), "text"
    if ctype == 5:  # XL_CELL_ERROR
        codes = {0: "#NULL!", 7: "#DIV/0!", 15: "#VALUE!", 23: "#REF!",
                 29: "#NAME?", 36: "#NUM!", 42: "#N/A"}
        return codes.get(int(raw), "#ERR!"), "text"
    return raw, "number"


def extract_sheet_cells(sheet, wb, input_fill_ids: set[int],
                        sheet_name: str, debug: bool) -> tuple[list[dict], int, int, int, int]:
    """Extract all non-empty cells from a sheet."""
    cells: list[dict] = []
    min_row = min_col = 999999
    max_row = max_col = 0

    for r in range(sheet.nrows):
        for c in range(sheet.ncols):
            value, vtype = xlrd_value(sheet, r, c, wb)
            if value is None:
                continue

            row_1 = r + 1
            col_1 = c + 1
            addr = cell_address(row_1, col_1)
            key = f"{sheet_name}!{addr}"

            try:
                fill_id = sheet.cell_xf_index(r, c)
            except Exception:
                fill_id = 0

            is_input = (fill_id in input_fill_ids) and (key not in EXCLUDE_INPUT_KEYS)

            cell: dict = {
                "key": key,
                "address": addr,
                "row": row_1,
                "col": col_1,
                "fill_id": fill_id,
                "style_idx": fill_id,
                "has_formula": False,
                "formula": "",
                "is_input": is_input,
                "value_type": vtype,
                "value": value,
            }

            if debug and is_input:
                print(f"    INPUT {key} = {repr(value)}")

            min_row = min(min_row, row_1)
            min_col = min(min_col, col_1)
            max_row = max(max_row, row_1)
            max_col = max(max_col, col_1)
            cells.append(cell)

    return cells, min_row, min_col, max_row, max_col


def annotate_formulas(cells: list[dict], formula_map: dict[str, str]) -> dict[str, str]:
    """Mark cells that have formulas and return a flat formula map for the model."""
    derived: dict[str, str] = {}
    for cell in cells:
        key = cell["key"]
        formula = formula_map.get(key, "")
        if formula and not cell.get("is_input"):
            cell["has_formula"] = True
            cell["formula"] = formula
            derived[key] = formula
    return derived


# ─── Model assembly ───────────────────────────────────────────────────────────

def build_input_cells_list(sheets_data: list[dict], hidden: set[str]) -> list[dict]:
    result: list[dict] = []
    for sheet in sheets_data:
        name = sheet["name"]
        for cell in sheet["cells"]:
            if cell.get("is_input"):
                result.append({
                    "key": cell["key"],
                    "sheet": name,
                    "address": cell["address"],
                    "value_type": cell["value_type"],
                    "default_value": cell.get("value", 0),
                })
    return result


# ─── Main ─────────────────────────────────────────────────────────────────────

def extract(xls_path: Path, debug: bool = False) -> tuple[dict, dict[str, str]]:
    wb = open_workbook(xls_path, debug=debug)
    print(f"  Sheets: {wb.sheet_names()}")

    input_fill_ids = build_input_fill_ids(wb, debug=debug)

    sheets_out: list[dict] = []
    all_formula_map: dict[str, str] = {}

    for sheet_name in wb.sheet_names():
        sheet = wb.sheet_by_name(sheet_name)
        if debug:
            print(f"\n  === {sheet_name} ({sheet.nrows}r × {sheet.ncols}c) ===")

        cells, min_row, min_col, max_row, max_col = extract_sheet_cells(
            sheet, wb, input_fill_ids, sheet_name, debug
        )
        if not cells:
            if debug:
                print(f"    (empty sheet, skipping)")
            continue

        # Annotate with explicit formulas
        derived = annotate_formulas(cells, EXPLICIT_FORMULAS)
        all_formula_map.update(derived)

        n_inputs = sum(1 for c in cells if c["is_input"])
        n_formulas = sum(1 for c in cells if c["has_formula"])
        print(f"  [{sheet_name}] {len(cells)} cells | {n_inputs} inputs | "
              f"{n_formulas} formulas")

        sheets_out.append({
            "name": sheet_name,
            "min_row": min_row,
            "min_col": min_col,
            "max_row": max_row,
            "max_col": max_col,
            "cells": cells,
        })

    input_cells = build_input_cells_list(sheets_out, HIDDEN_SHEETS)

    model: dict = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "template_xlsx": str(xls_path),
        "input_fill_ids": sorted(input_fill_ids),
        "sheets": sheets_out,
        "input_cells": input_cells,
        "hidden_sheets": sorted(HIDDEN_SHEETS),
    }

    return model, all_formula_map


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Extract Value Add XLS → value_add_model.json")
    p.add_argument("--xls", default=str(DEFAULT_XLS))
    p.add_argument("--out-model", default=str(OUTPUT_MODEL))
    p.add_argument("--out-formulas", default=str(OUTPUT_FORMULAS))
    p.add_argument("--debug", action="store_true")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    xls_path = Path(args.xls).expanduser()

    if not xls_path.exists():
        print(f"ERROR: XLS file not found: {xls_path}")
        sys.exit(1)

    print(f"Extracting Value Add model from: {xls_path.name}")
    model, formula_map = extract(xls_path, debug=args.debug)

    out_model = Path(args.out_model)
    out_formulas = Path(args.out_formulas)

    with open(out_model, "w", encoding="utf-8") as f:
        json.dump(model, f, indent=2, ensure_ascii=True)
    print(f"\nWrote model     → {out_model} ({out_model.stat().st_size // 1024} KB)")

    with open(out_formulas, "w", encoding="utf-8") as f:
        json.dump(formula_map, f, indent=2, ensure_ascii=True)
    print(f"Wrote formulas  → {out_formulas} ({len(formula_map)} entries)")

    total = sum(len(s["cells"]) for s in model["sheets"])
    inputs = len(model["input_cells"])
    print(f"\nSummary: {len(model['sheets'])} sheets | {total} cells | "
          f"{inputs} inputs | {len(formula_map)} formulas")
    print(f"Hidden sheets: {model['hidden_sheets']}")
    print("\nDone. Start server with:  bash run_local.sh")


if __name__ == "__main__":
    main()
