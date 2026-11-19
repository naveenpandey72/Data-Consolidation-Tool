"""
Master Template Profiler.
Analyzes the uploaded master file dynamically:
- Identifies all sheets, columns, and rows.
- Distinguishes labels, formulas, and input cells.
- Maps semantic anchors (Row Header + Column Header / FY Period) for every cell.
- Guarantees that the master file defines the ground truth data structure.
"""

import openpyxl
from openpyxl.utils import get_column_letter, coordinate_to_tuple
from typing import Dict, List, Tuple, Any, Optional
import re
from config.template_config import normalize_sheet_name


class MasterProfile:
    def __init__(self, filepath: str):
        self.filepath = filepath
        self.wb = openpyxl.load_workbook(filepath, data_only=False)
        self.sheet_names = [s for s in self.wb.sheetnames if s not in ["Region_Data", "Merge_Notes"]]
        self.sheet_map = {normalize_sheet_name(s): s for s in self.sheet_names}
        
        # Structure information
        self.input_cells: Dict[str, Dict[str, Dict[str, Any]]] = {}
        self.formula_cells: Dict[str, Dict[str, str]] = {}
        self.sheet_bounds: Dict[str, Tuple[int, int]] = {} # sheet -> (max_r, max_c)
        self.column_headers: Dict[str, Dict[int, str]] = {} # sheet -> {col_idx: header_label}
        self.row_labels: Dict[str, Dict[int, str]] = {} # sheet -> {row_idx: row_label}

        self._profile()

    def _profile(self):
        for sname in self.sheet_names:
            ws = self.wb[sname]
            self.sheet_bounds[sname] = (ws.max_row, ws.max_column)
            self.input_cells[sname] = {}
            self.formula_cells[sname] = {}
            self.column_headers[sname] = {}
            self.row_labels[sname] = {}

            norm_name = normalize_sheet_name(sname)
            
            # 1. Profile 303-01 sheets (Qualitative disclosure)
            if "303_01" in norm_name:
                # Text answers are typically in column D (NHAI Response for FY 25-26)
                for r in range(8, ws.max_row + 1):
                    row_lbl = ws.cell(r, 2).value or ws.cell(r, 1).value
                    if row_lbl:
                        self.row_labels[sname][r] = str(row_lbl).strip()
                    
                    val = ws.cell(r, 4).value
                    val_str = str(val or "")
                    if not val_str.startswith("="):
                        self.input_cells[sname][f"D{r}"] = {
                            "coord": f"D{r}",
                            "row": r,
                            "col": 4,
                            "row_label": str(row_lbl).strip() if row_lbl else f"Requirement {ws.cell(r, 1).value}",
                            "col_label": "Response for FY 2025-26",
                            "is_text": True,
                        }

            # 2. Profile 303-03 sheets (Withdrawal)
            elif "303_03" in norm_name:
                self._profile_matrix_sheet(sname, ws, row_header_row=5, sub_header_row=6, label_col=1, data_start_row=7)

            # 3. Profile 303-04 sheets (Discharge)
            elif "303_04" in norm_name:
                self._profile_matrix_sheet(sname, ws, row_header_row=5, sub_header_row=6, label_col=3, data_start_row=7)

            # 4. Profile 303-05 sheets (Consumption & storage)
            elif "303_05" in norm_name:
                self._profile_303_05_sheet(sname, ws)

            # Record formulas
            for r in range(1, ws.max_row + 1):
                for c in range(1, ws.max_column + 1):
                    v = ws.cell(r, c).value
                    if v and str(v).startswith("="):
                        coord = f"{get_column_letter(c)}{r}"
                        self.formula_cells[sname][coord] = str(v)

    def _profile_matrix_sheet(self, sname: str, ws: openpyxl.worksheet.worksheet.Worksheet,
                              row_header_row: int, sub_header_row: int, label_col: int, data_start_row: int):
        # Extract column headers by carrying forward merged cells
        current_fy = None
        for c in range(1, ws.max_column + 1):
            fy_val = ws.cell(row_header_row, c).value
            if fy_val and str(fy_val).strip():
                current_fy = str(fy_val).strip()
            
            sub_val = ws.cell(sub_header_row, c).value
            sub_str = str(sub_val).strip() if sub_val else ""
            
            if current_fy and sub_str:
                self.column_headers[sname][c] = f"{current_fy} – {sub_str}"
            elif current_fy:
                self.column_headers[sname][c] = current_fy
            elif sub_str:
                self.column_headers[sname][c] = sub_str

        # Extract row labels
        for r in range(data_start_row, ws.max_row + 1):
            lbl_val = ws.cell(r, label_col).value
            if lbl_val and str(lbl_val).strip():
                self.row_labels[sname][r] = str(lbl_val).strip()

        # Identify input cells (cells that are not formulas and not "DO NOT FILL")
        for r in range(data_start_row, ws.max_row + 1):
            row_lbl = self.row_labels[sname].get(r, "")
            # Skip summary rows like Total volume, Total water withdrawal
            if any(term in row_lbl.lower() for term in ["total volume", "total water"]):
                continue

            for c in range(1, ws.max_column + 1):
                col_lbl = self.column_headers[sname].get(c, "")
                if not col_lbl or not row_lbl:
                    continue

                coord = f"{get_column_letter(c)}{r}"
                cell_val = ws.cell(r, c).value
                val_str = str(cell_val or "").strip()

                if val_str.upper() == "DO NOT FILL":
                    continue
                if val_str.startswith("=") and "Region_Data" not in val_str:
                    # In a blank master, total formulas like =SUM(...) should not be treated as inputs
                    continue

                self.input_cells[sname][coord] = {
                    "coord": coord,
                    "row": r,
                    "col": c,
                    "row_label": row_lbl,
                    "col_label": col_lbl,
                    "is_text": False,
                }

    def _profile_303_05_sheet(self, sname: str, ws: openpyxl.worksheet.worksheet.Worksheet):
        # Row 5 has headers: Unit | FY 2019-20 | FY 2020-21 ...
        for c in range(5, ws.max_column + 1):
            fy = ws.cell(5, c).value
            if fy and str(fy).strip():
                self.column_headers[sname][c] = str(fy).strip()

        for r in range(6, ws.max_row + 1):
            lbl = ws.cell(r, 2).value
            if lbl and str(lbl).strip():
                self.row_labels[sname][r] = str(lbl).strip()

        # Row 8 is change in storage (input cells in FY columns)
        # Row 9 is qualitative narrative in FY columns (or text)
        for c in range(5, ws.max_column + 1):
            col_lbl = self.column_headers[sname].get(c, "")
            if not col_lbl:
                continue

            # Row 8: Change in storage (numeric)
            coord_r8 = f"{get_column_letter(c)}8"
            self.input_cells[sname][coord_r8] = {
                "coord": coord_r8,
                "row": 8,
                "col": c,
                "row_label": self.row_labels[sname].get(8, "Change in water storage"),
                "col_label": col_lbl,
                "is_text": False,
            }

            # Row 9: Water-related impact narrative (text)
            coord_r9 = f"{get_column_letter(c)}9"
            self.input_cells[sname][coord_r9] = {
                "coord": coord_r9,
                "row": 9,
                "col": c,
                "row_label": self.row_labels[sname].get(9, "water-related impact"),
                "col_label": col_lbl,
                "is_text": True,
            }
