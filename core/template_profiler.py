"""
Master Template Deep Profiler.
Scans the uploaded master file completely:
- Profiles all sheets, dimensions, headers, labels, and formulas
- Records static template text to distinguish boilerplate from user data
- Maps semantic anchors (Row Label + Column / FY Period) for robust matching
- Ensures the Master File defines the canonical ground truth structure
"""

import openpyxl
from openpyxl.utils import get_column_letter
from typing import Dict, List, Tuple, Any, Optional
import re
from config.template_config import normalize_sheet_name


class MasterProfile:
    def __init__(self, filepath: str):
        self.filepath = filepath
        self.wb = openpyxl.load_workbook(filepath, data_only=False)
        self.sheet_names = [s for s in self.wb.sheetnames if s not in ["Region_Data", "Merge_Notes"]]
        self.sheet_map = {normalize_sheet_name(s): s for s in self.sheet_names}

        # Formula cells: sheet -> {coord: formula_str}
        self.formula_cells: Dict[str, Dict[str, str]] = {}
        # Static template values: sheet -> {coord: text_val}
        self.static_labels: Dict[str, Dict[str, str]] = {}
        # Sheet bounds: sheet -> (max_row, max_column)
        self.sheet_bounds: Dict[str, Tuple[int, int]] = {}
        # Semantic column headers: sheet -> {col_idx: header_label}
        self.column_headers: Dict[str, Dict[int, str]] = {}
        # Semantic row labels: sheet -> {row_idx: row_label}
        self.row_labels: Dict[str, Dict[int, str]] = {}

        self._profile()

    def _profile(self):
        for sname in self.sheet_names:
            ws = self.wb[sname]
            self.sheet_bounds[sname] = (ws.max_row, ws.max_column)
            self.formula_cells[sname] = {}
            self.static_labels[sname] = {}
            self.column_headers[sname] = {}
            self.row_labels[sname] = {}

            norm_name = normalize_sheet_name(sname)

            # 1. Extract semantic column headers and row labels
            self._extract_semantic_anchors(sname, ws, norm_name)

            # 2. Record formulas and static labels
            for r in range(1, ws.max_row + 1):
                for c in range(1, ws.max_column + 1):
                    val = ws.cell(r, c).value
                    coord = f"{get_column_letter(c)}{r}"
                    if val is not None:
                        val_str = str(val).strip()
                        if val_str.startswith("="):
                            self.formula_cells[sname][coord] = val_str
                        else:
                            self.static_labels[sname][coord] = val_str

    def _extract_semantic_anchors(self, sname: str, ws: openpyxl.worksheet.worksheet.Worksheet, norm_name: str):
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
                    self.column_headers[sname][c] = f"{current_fy} – {sub_str}"
                elif current_fy:
                    self.column_headers[sname][c] = current_fy
                elif sub_str:
                    self.column_headers[sname][c] = sub_str

            for r in range(data_start, ws.max_row + 1):
                lbl = ws.cell(r, label_col).value
                if lbl and str(lbl).strip():
                    self.row_labels[sname][r] = str(lbl).strip()

        elif "303_05" in norm_name:
            for c in range(5, ws.max_column + 1):
                fy = ws.cell(5, c).value
                if fy and str(fy).strip():
                    self.column_headers[sname][c] = str(fy).strip()
            for r in range(6, ws.max_row + 1):
                lbl = ws.cell(r, 2).value
                if lbl and str(lbl).strip():
                    self.row_labels[sname][r] = str(lbl).strip()

        elif "303_01" in norm_name:
            for r in range(8, ws.max_row + 1):
                lbl = ws.cell(r, 2).value or ws.cell(r, 1).value
                if lbl and str(lbl).strip():
                    self.row_labels[sname][r] = str(lbl).strip()

    def get_semantic_labels(self, sname: str, r: int, c: int) -> Tuple[str, str]:
        """Returns (row_label, col_label) for a given coordinate."""
        row_lbl = self.row_labels.get(sname, {}).get(r, "")
        col_lbl = self.column_headers.get(sname, {}).get(c, "")
        return (row_lbl, col_lbl)
