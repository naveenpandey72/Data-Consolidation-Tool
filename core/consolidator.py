"""
GRI 303 Water Data Consolidator.
Core engine that executes the consolidation:
- Preserves 100% of the uploaded Master file's column layout and formulas
- Builds the Region_Data sheet with dynamic Excel formulas
- Links Master cells to Region_Data totals
- Deduplicates and groups text responses
- Preserves merged cells and formatting
- Generates the Merge_Notes audit sheet
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from typing import List, Dict, Any, Tuple
import datetime
import os

from core.template_profiler import MasterProfile
from core.regional_reader import RegionalData
from core.validator import run_validations, ValidationReport


def consolidate_workbooks(master_profile: MasterProfile, regional_datasets: List[RegionalData], output_path: str) -> Tuple[str, ValidationReport]:
    """
    Main consolidation function.
    Returns (output_filepath, validation_report).
    """
    wb = openpyxl.load_workbook(master_profile.filepath, data_only=False)

    # Remove pre-existing Region_Data or Merge_Notes sheets if re-running
    for sheet_to_remove in ["Region_Data", "Merge_Notes"]:
        if sheet_to_remove in wb.sheetnames:
            del wb[sheet_to_remove]

    # Run validation checks
    val_report = run_validations(regional_datasets)

    # 1. Identify all unique numeric cells that received data from at least one region
    region_names = [r.region_name for r in regional_datasets]
    region_col_letters = [get_column_letter(5 + i) for i in range(len(region_names))]
    total_col_letter = get_column_letter(5 + len(region_names))

    # Collect all numeric input entries
    active_numeric_cells = [] # list of (sheet, coord, meta)
    for sname in master_profile.sheet_names:
        input_meta = master_profile.input_cells.get(sname, {})
        for coord, meta in input_meta.items():
            if meta.get("is_text"):
                continue
            # Check if any region reported a numeric value for this cell
            has_value = any(coord in reg.numeric_values.get(sname, {}) for reg in regional_datasets)
            if has_value:
                active_numeric_cells.append((sname, coord, meta))

    # 2. Build Region_Data Sheet
    ws_reg = wb.create_sheet(title="Region_Data")
    ws_reg.views.sheetView[0].showGridLines = True

    # Styling definitions
    title_font = Font(name="Calibri", size=11, bold=True, color="1F497D")
    sub_font = Font(name="Calibri", size=9, italic=True, color="595959")
    hdr_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    hdr_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    data_font = Font(name="Calibri", size=10)
    num_font = Font(name="Calibri", size=10, color="002060")
    total_font = Font(name="Calibri", size=10, bold=True)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9")
    )

    ws_reg["A1"] = "Region-wise numeric data (source: individual region files). Master cells link to the Total column."
    ws_reg["A1"].font = title_font
    ws_reg["A2"] = "Blank = region reported NA/Nil/nothing for that cell. Values in blue are copied from region files (formula entries in source files were evaluated)."
    ws_reg["A2"].font = sub_font

    # Headers at row 4
    headers = ["Master sheet", "Cell", "Row label", "Period / column"] + region_names + ["Total (feeds master)"]
    for c_idx, h_text in enumerate(headers, start=1):
        cell = ws_reg.cell(row=4, column=c_idx, value=h_text)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal="center" if c_idx >= 5 else "left", vertical="center")

    # Populate rows
    current_row = 5
    for sname, coord, meta in active_numeric_cells:
        ws_reg.cell(row=current_row, column=1, value=sname).font = data_font
        ws_reg.cell(row=current_row, column=2, value=coord).font = data_font
        ws_reg.cell(row=current_row, column=3, value=meta["row_label"]).font = data_font
        ws_reg.cell(row=current_row, column=4, value=meta["col_label"]).font = data_font

        for i, reg in enumerate(regional_datasets):
            reg_val = reg.numeric_values.get(sname, {}).get(coord)
            cell = ws_reg.cell(row=current_row, column=5 + i, value=reg_val)
            cell.font = num_font
            cell.border = thin_border
            if reg_val is not None:
                cell.number_format = "#,##0"

        # Total formula =SUM(E5:K5)
        sum_formula = f"=SUM({region_col_letters[0]}{current_row}:{region_col_letters[-1]}{current_row})"
        total_cell = ws_reg.cell(row=current_row, column=5 + len(region_names), value=sum_formula)
        total_cell.font = total_font
        total_cell.border = thin_border
        total_cell.number_format = "#,##0"

        # Link master cell to this total cell: =Region_Data!L5
        master_ws = wb[sname]
        master_ws[coord].value = f"=Region_Data!{total_col_letter}{current_row}"
        master_ws[coord].font = data_font
        master_ws[coord].number_format = "#,##0"

        current_row += 1

    # Auto-adjust column widths for Region_Data
    col_widths = {1: 22, 2: 8, 3: 40, 4: 25}
    for c_idx, w in col_widths.items():
        ws_reg.column_dimensions[get_column_letter(c_idx)].width = w
    for c_idx in range(5, 5 + len(region_names) + 1):
        ws_reg.column_dimensions[get_column_letter(c_idx)].width = 15

    # 3. Consolidate Text Responses in Master Sheets
    for sname, input_meta in master_profile.input_cells.items():
        if sname not in wb.sheetnames:
            continue
        master_ws = wb[sname]

        for coord, meta in input_meta.items():
            if not meta.get("is_text"):
                continue

            # Gather text responses for this cell across regions
            # Map response text -> list of region names
            response_groups: Dict[str, List[str]] = {}
            for reg in regional_datasets:
                ans = reg.text_values.get(sname, {}).get(coord)
                if ans and ans.strip():
                    norm_ans = ans.strip()
                    response_groups.setdefault(norm_ans, []).append(reg.region_name)

            if response_groups:
                paragraphs = []
                for ans_text, r_names in response_groups.items():
                    tag = f"[{', '.join(sorted(r_names))}]"
                    paragraphs.append(f"{tag} {ans_text}")
                combined_text = "\n\n".join(paragraphs)

                cell = master_ws[coord]
                cell.value = combined_text
                cell.alignment = Alignment(wrap_text=True, vertical="top")

                # Dynamic row height
                line_count = combined_text.count("\n") + len(combined_text) // 80 + 1
                row_idx = meta["row"]
                master_ws.row_dimensions[row_idx].height = max(24, min(450, line_count * 15))

        # Check for inserted note rows (e.g. Bareilly)
        all_inserted_notes = []
        for reg in regional_datasets:
            notes = reg.inserted_note_rows.get(sname, [])
            for note in notes:
                all_inserted_notes.append((reg.region_name, note))

        if all_inserted_notes:
            note_row = master_ws.max_row + 1
            for reg_name, note_txt in all_inserted_notes:
                cell = master_ws.cell(row=note_row, column=2, value=f"[{reg_name}] {note_txt}")
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                master_ws.row_dimensions[note_row].height = 35
                note_row += 1

    # 4. Build Merge_Notes Sheet
    ws_notes = wb.create_sheet(title="Merge_Notes")
    ws_notes.views.sheetView[0].showGridLines = True
    ws_notes.column_dimensions["A"].width = 110

    sec_hdr_font = Font(name="Calibri", size=11, bold=True, color="1F497D")
    body_font = Font(name="Calibri", size=10)
    bullet_font = Font(name="Calibri", size=10)

    n_row = 1
    ws_notes.cell(row=n_row, column=1, value="How this master was consolidated").font = sec_hdr_font
    n_row += 1

    notes_intro = [
        "Numeric cells (303-03, 303-04, 303-05 row c, non-compliance counts): sum across all regions. Region-wise values are saved on sheet Region_Data; master cells link to its Total column.",
        "Text responses: all regions combined in the same cell, each tagged with [Region]. Identical answers are grouped under one tag.",
        "'NA', 'Nil', 'N/A' and 'DO NOT FILL' entries were excluded from sums and treated as blanks.",
        "Existing master formulas (totals and the 303-05 consumption rows) were preserved and calculate dynamically from merged data.",
        "Formulas inside source input cells (e.g. =16000000*2+4567) were evaluated to numeric values before summing.",
        f"Consolidated at: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} with {len(regional_datasets)} regional files.",
    ]
    for text in notes_intro:
        ws_notes.cell(row=n_row, column=1, value=text).font = body_font
        n_row += 1

    n_row += 1
    ws_notes.cell(row=n_row, column=1, value="Items to check with the regions (Warnings & Audit Log)").font = sec_hdr_font
    n_row += 1

    if val_report.warnings:
        for w in val_report.warnings:
            ws_notes.cell(row=n_row, column=1, value=f"• {w}").font = bullet_font
            n_row += 1
    else:
        ws_notes.cell(row=n_row, column=1, value="• No anomalies or suspicious values detected. All checks passed cleanly.").font = bullet_font
        n_row += 1

    n_row += 1
    ws_notes.cell(row=n_row, column=1, value="Source Files & Submission Log").font = sec_hdr_font
    n_row += 1
    for reg in regional_datasets:
        ws_notes.cell(row=n_row, column=1, value=f"• Region: {reg.region_name} | File: {reg.filename}").font = bullet_font
        n_row += 1

    # Save consolidated file
    wb.save(output_path)
    return output_path, val_report
