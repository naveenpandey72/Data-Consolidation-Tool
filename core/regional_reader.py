"""
Regional File Reader & Semantic Structure Matcher.
Reads regional files with arbitrary names and potentially varying/shifted columns/rows,
and maps extracted data safely into the Master File's canonical schema.
"""

import openpyxl
from openpyxl.utils import get_column_letter
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
        
        # Mapped data: sheet_name -> {master_coord: parsed_value}
        self.numeric_values: Dict[str, Dict[str, float]] = {}
        self.text_values: Dict[str, Dict[str, str]] = {}
        
        # Diagnostics & Audit
        self.missing_sheets: List[str] = []
        self.extra_sheets: List[str] = []
        self.stray_values: List[Dict[str, Any]] = []
        self.hardcoded_totals: List[Dict[str, Any]] = []
        self.inserted_note_rows: Dict[str, List[str]] = {} # sheet -> list of text paragraphs
        self.warnings: List[str] = []


def extract_regional_sheet_headers(ws: openpyxl.worksheet.worksheet.Worksheet, norm_name: str) -> Tuple[Dict[int, str], Dict[int, str]]:
    """Extracts column headers and row labels from a regional sheet."""
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
    """
    Reads a regional file, aligning its data to the master template.
    Preserves master column structure: any regional column not present in the master
    is logged and ignored, never distorting the master schema.
    """
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

    # Read mapped sheets
    for master_sheet, input_cells_dict in master_profile.input_cells.items():
        norm = normalize_sheet_name(master_sheet)
        if norm not in reg_sheet_map:
            continue

        reg_sheet_name = reg_sheet_map[norm]
        ws = wb[reg_sheet_name]
        reg_cols, reg_rows = extract_regional_sheet_headers(ws, norm)

        reg_data.numeric_values[master_sheet] = {}
        reg_data.text_values[master_sheet] = {}

        # Inverted index for shifted columns/rows
        reg_col_lookup = {v.lower().strip(): c for c, v in reg_cols.items()}
        reg_row_lookup = {v.lower().strip(): r for r, v in reg_rows.items()}

        # Match each Master input cell
        for master_coord, cell_meta in input_cells_dict.items():
            target_r = cell_meta["row"]
            target_c = cell_meta["col"]
            row_lbl = cell_meta["row_label"]
            col_lbl = cell_meta["col_label"]
            is_text_cell = cell_meta.get("is_text", False)

            # 1. Coordinate check: does cell (target_r, target_c) have matching headers?
            matched_r = target_r
            matched_c = target_c

            reg_c_lbl = reg_cols.get(target_c, "").lower().strip()
            reg_r_lbl = reg_rows.get(target_r, "").lower().strip()
            expected_c = col_lbl.lower().strip()
            expected_r = row_lbl.lower().strip()

            # If layout shifted in regional file, find via semantic anchor
            if reg_c_lbl != expected_c or (reg_r_lbl and expected_r and not reg_r_lbl.startswith(expected_r[:15])):
                if expected_c in reg_col_lookup:
                    matched_c = reg_col_lookup[expected_c]
                if expected_r in reg_row_lookup:
                    matched_r = reg_row_lookup[expected_r]

            # Read regional cell
            if matched_r <= ws.max_row and matched_c <= ws.max_column:
                raw_val = ws.cell(matched_r, matched_c).value
                parsed_val, val_type = evaluate_cell_value(raw_val)

                if parsed_val is not None:
                    if is_text_cell:
                        reg_data.text_values[master_sheet][master_coord] = str(parsed_val).strip()
                    else:
                        if isinstance(parsed_val, (int, float)):
                            reg_data.numeric_values[master_sheet][master_coord] = parsed_val
                        else:
                            # Text entered in numeric cell
                            reg_data.warnings.append(
                                f"{reg_data.region_name}: text '{parsed_val}' entered in numeric cell {master_coord} of {master_sheet}"
                            )

        # Check for hardcoded totals typed over formulas in regional files
        for master_coord, formula_str in master_profile.formula_cells.get(master_sheet, {}).items():
            coord_cell = ws[master_coord] if master_coord in ws else None
            if coord_cell and coord_cell.value is not None:
                cell_str = str(coord_cell.value).strip()
                if not cell_str.startswith("=") and not is_ignored_text(cell_str):
                    reg_data.hardcoded_totals.append({
                        "sheet": master_sheet,
                        "coord": master_coord,
                        "typed_value": coord_cell.value,
                        "master_formula": formula_str,
                    })

        # Check for stray values outside master bounds (FR-8)
        max_master_r, max_master_c = master_profile.sheet_bounds.get(master_sheet, (ws.max_row, ws.max_column))
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                # Outside table range if row > max_master_r or col > max_master_c
                if r > max_master_r or c > max_master_c:
                    val = ws.cell(r, c).value
                    if val is not None and not is_ignored_text(val):
                        coord = f"{get_column_letter(c)}{r}"
                        reg_data.stray_values.append({
                            "sheet": master_sheet,
                            "coord": coord,
                            "value": val,
                        })

        # Detect inserted notes (e.g. Bareilly notes inserted after row 16)
        if "303_01" in norm or "303_05" in norm:
            for r in range(16, ws.max_row + 1):
                val = ws.cell(r, 2).value or ws.cell(r, 4).value
                if val and ("note" in str(val).lower() or "catchment" in str(val).lower() or "value-chain" in str(val).lower()):
                    reg_data.inserted_note_rows.setdefault(master_sheet, []).append(str(val).strip())

    return reg_data
