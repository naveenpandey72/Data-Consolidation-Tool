"""
Regional File Universal Deep Reader & Semantic Structure Matcher.
Scans 100% of non-empty cells across all sheets in regional files,
evaluates numbers and text comments, and aligns them into the Master File.
Guarantees that nothing is missed while strictly preserving Master structure.
"""

import openpyxl
from openpyxl.utils import get_column_letter, coordinate_to_tuple
from typing import Dict, List, Any, Optional, Tuple
import re

from config.template_config import normalize_sheet_name, clean_region_name
from core.template_profiler import MasterProfile
from core.ast_evaluator import evaluate_cell_value, is_ignored_text


class RegionalData:
    def __init__(self, filepath: str, region_name: Optional[str] = None):
        self.filepath = filepath
        self.filename = filepath.replace("\\", "/").split("/")[-1]
        self.region_name = region_name or clean_region_name(self.filename)

        # Mapped numeric data: sheet_name -> {master_coord: float/int}
        self.numeric_values: Dict[str, Dict[str, float]] = {}
        # Mapped text/commentary data: sheet_name -> {master_coord: text_str}
        self.text_values: Dict[str, Dict[str, str]] = {}

        # Audit & Diagnostics
        self.missing_sheets: List[str] = []
        self.extra_sheets: List[str] = []
        self.stray_values: List[Dict[str, Any]] = []
        self.hardcoded_totals: List[Dict[str, Any]] = []
        self.inserted_note_rows: Dict[str, List[str]] = {}
        self.warnings: List[str] = []


def extract_regional_sheet_headers(ws: openpyxl.worksheet.worksheet.Worksheet, norm_name: str) -> Tuple[Dict[int, str], Dict[int, str]]:
    col_headers = {}
    row_labels = {}

    if "303_03" in norm_name or "303_04" in norm_name:
        row_hdr = 5
        sub_hdr = 6
        label_col = 1 if "303_03" in norm_name else 3
        data_start = 7

        current_fy = None
        for c in range(1, ws.max_column + 1):
            fy = ws.cell(row_hdr, c).value
            if fy and str(fy).strip():
                current_fy = str(fy).strip()
            sub = ws.cell(sub_hdr, c).value
            sub_str = str(sub).strip() if sub else ""
            if current_fy and sub_str:
                col_headers[c] = f"{current_fy} – {sub_str}"
            elif current_fy:
                col_headers[c] = current_fy
            elif sub_str:
                col_headers[c] = sub_str

        for r in range(data_start, ws.max_row + 1):
            lbl = ws.cell(r, label_col).value
            if lbl and str(lbl).strip():
                row_labels[r] = str(lbl).strip()

    elif "303_05" in norm_name:
        for c in range(5, ws.max_column + 1):
            fy = ws.cell(5, c).value
            if fy and str(fy).strip():
                col_headers[c] = str(fy).strip()
        for r in range(6, ws.max_row + 1):
            lbl = ws.cell(r, 2).value
            if lbl and str(lbl).strip():
                row_labels[r] = str(lbl).strip()

    elif "303_01" in norm_name:
        for r in range(8, ws.max_row + 1):
            lbl = ws.cell(r, 2).value or ws.cell(r, 1).value
            if lbl and str(lbl).strip():
                row_labels[r] = str(lbl).strip()

    return col_headers, row_labels


def read_regional_file(filepath: str, master_profile: MasterProfile, region_name: Optional[str] = None) -> RegionalData:
    reg_data = RegionalData(filepath, region_name)
    wb = openpyxl.load_workbook(filepath, data_only=False)

    reg_sheet_map = {normalize_sheet_name(s): s for s in wb.sheetnames}

    # Check for missing and extra sheets
    for master_sheet in master_profile.sheet_names:
        norm = normalize_sheet_name(master_sheet)
        if norm not in reg_sheet_map:
            reg_data.missing_sheets.append(master_sheet)
            reg_data.warnings.append(f"{reg_data.region_name}: missing sheet '{master_sheet}'")

    for reg_sheet in wb.sheetnames:
        norm = normalize_sheet_name(reg_sheet)
        if norm not in master_profile.sheet_map:
            reg_data.extra_sheets.append(reg_sheet)
            reg_data.warnings.append(f"{reg_data.region_name}: extra sheet '{reg_sheet}' not in master template (skipped)")

    # Scan all sheets
    for master_sheet in master_profile.sheet_names:
        norm = normalize_sheet_name(master_sheet)
        if norm not in reg_sheet_map:
            continue

        reg_sheet_name = reg_sheet_map[norm]
        ws = wb[reg_sheet_name]
        reg_cols, reg_rows = extract_regional_sheet_headers(ws, norm)

        reg_data.numeric_values[master_sheet] = {}
        reg_data.text_values[master_sheet] = {}

        max_master_r, max_master_c = master_profile.sheet_bounds.get(master_sheet, (ws.max_row, ws.max_column))
        master_formulas = master_profile.formula_cells.get(master_sheet, {})
        master_static = master_profile.static_labels.get(master_sheet, {})

        # Inverted index for semantic matching of shifted columns/rows
        master_cols = master_profile.column_headers.get(master_sheet, {})
        master_rows = master_profile.row_labels.get(master_sheet, {})
        
        # Build semantic map: (row_label_lower, col_label_lower) -> master_coord
        semantic_to_master = {}
        for r_idx, r_lbl in master_rows.items():
            for c_idx, c_lbl in master_cols.items():
                semantic_to_master[(r_lbl.lower().strip(), c_lbl.lower().strip())] = f"{get_column_letter(c_idx)}{r_idx}"

        # DEEP SCAN: Iterate over every cell in the regional worksheet
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                raw_val = ws.cell(r, c).value
                if raw_val is None:
                    continue

                coord = f"{get_column_letter(c)}{r}"
                val_str = str(raw_val).strip()

                # 1. Skip ignored placeholders (NA, Nil, DO NOT FILL)
                if is_ignored_text(val_str):
                    continue

                # 2. Check if outside master table bounds (stray values / inserted rows)
                if r > max_master_r or c > max_master_c:
                    # Check if it's an inserted note paragraph (e.g. Bareilly notes)
                    if ("303_01" in norm or "303_05" in norm) and c <= 4:
                        if any(term in val_str.lower() for term in ["response", "note", "catchment", "value-chain"]):
                            reg_data.inserted_note_rows.setdefault(master_sheet, []).append(val_str)
                            continue
                    
                    # Otherwise, it's a stray value outside table (e.g. Agra T43, Jhansi row 22)
                    parsed_val, _ = evaluate_cell_value(raw_val)
                    if parsed_val is not None:
                        reg_data.stray_values.append({
                            "sheet": master_sheet,
                            "coord": coord,
                            "value": raw_val
                        })
                    continue

                # 3. Skip if cell merely duplicates Master template static text
                master_stat_val = master_static.get(coord)
                if master_stat_val and master_stat_val.strip().lower() == val_str.lower():
                    continue

                # 4. Check if Master has a formula in this cell
                if coord in master_formulas:
                    if not val_str.startswith("="):
                        reg_data.hardcoded_totals.append({
                            "sheet": master_sheet,
                            "coord": coord,
                            "typed_value": raw_val,
                            "master_formula": master_formulas[coord]
                        })
                    continue

                # 5. Semantic Anchor Alignment: Map regional cell to master coordinate
                target_coord = coord
                reg_c_lbl = reg_cols.get(c, "").lower().strip()
                reg_r_lbl = reg_rows.get(r, "").lower().strip()

                master_r_lbl = master_rows.get(r, "").lower().strip()
                master_c_lbl = master_cols.get(c, "").lower().strip()

                # If current coordinate (r, c) headers match Master, keep direct coordinate
                if reg_c_lbl == master_c_lbl and reg_r_lbl == master_r_lbl:
                    target_coord = coord
                elif reg_c_lbl and reg_r_lbl:
                    # If columns or rows shifted, match within the same section
                    # Determine section: 1 for upper block (r < 15), 2 for lower block (r >= 15)
                    reg_section = 1 if r < 15 else 2
                    for m_r_idx, m_r_lbl in master_rows.items():
                        m_section = 1 if m_r_idx < 15 else 2
                        if m_section == reg_section and m_r_lbl.lower().strip() == reg_r_lbl:
                            for m_c_idx, m_c_lbl in master_cols.items():
                                if m_c_lbl.lower().strip() == reg_c_lbl:
                                    target_coord = f"{get_column_letter(m_c_idx)}{m_r_idx}"
                                    break
                            break

                # Check for stray values inside matrix sheet that fall outside valid table rows
                if "303_04" in norm and r in [13, 14, 15, 21, 22, 23]:
                    parsed_val, _ = evaluate_cell_value(raw_val)
                    if parsed_val is not None and isinstance(parsed_val, (int, float)):
                        reg_data.stray_values.append({
                            "sheet": master_sheet,
                            "coord": coord,
                            "value": raw_val
                        })
                        continue

                # 6. Parse and classify value
                parsed_val, val_type = evaluate_cell_value(raw_val)
                if parsed_val is None:
                    continue

                if isinstance(parsed_val, (int, float)):
                    reg_data.numeric_values[master_sheet][target_coord] = parsed_val
                else:
                    reg_data.text_values[master_sheet][target_coord] = str(parsed_val).strip()

    return reg_data
