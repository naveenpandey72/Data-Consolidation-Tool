# Data Consolidation Tool

Production consolidation engine built to specifications for Environment ESG reporting.

---

## Key Capabilities & Architectural Guarantees

### 1. Invariance to File Names (FR-1, FR-2, FR-3)
- **Master file can be named anything** (e.g., `Master_final_v3.xlsx`, `Blank_Template.xlsx`, `NHAI_303_Water_Master.xlsx`).
- **Regional files can be named anything** (e.g., `303_Water_Agra.xlsx`, `PIU_Kanpur_Water.xlsx`, `Jhansi.xlsx`).
- Regional names are auto-inferred from file names with an editable override in the UI/CLI.

### 2. Column Adaptation & Strict Master Structure Preservation (FR-6, FR-7, FR-11)
- **The Uploaded Master File is the Ground Truth**: Whatever sheets, columns, rows, formulas, headers, and merged cells exist in the uploaded master file are preserved 100%.
- **Handling Shifted or Wider Regional Structures**:
  - If a regional office added extra columns (e.g., Kanpur's wider sheet), moved columns, or inserted rows (e.g., Bareilly notes), the engine **maps values using semantic anchors** (`Row Label` + `Column / FY Header`) rather than blind coordinate lookups.
  - Extra columns in regional files that do not exist in the master are safely omitted from master sheets and logged in `Merge_Notes`.
  - Stray cells outside the master table boundary (e.g., Agra `T43`, Jhansi row 22) are never copied into the master table; they are captured in the audit log.

### 3. Dynamic Audit Traceability (`Region_Data` & `Merge_Notes`)
- **`Region_Data` Sheet**: Contains row-by-row breakdown across all regions with an Excel formula `=SUM(E5:K5)`.
- **Master Links**: Master numeric input cells link dynamically to `Region_Data` (e.g., `=Region_Data!L5`).
- **Formula Recalculation**: All master summary totals (`=SUM(...)`) and `303-05` consumption rows (`=Withdrawal - Discharge`) remain live native formulas and compute automatically.
- **Text Grouping**: Text responses across regions are grouped and tagged under brackets (e.g., `[Jhansi, Lucknow] Ground Water`).

---

## How to Run

### Option A: Web User Interface
Run the self-contained local web application:
```bash
python app.py
```
Open your browser at `http://127.0.0.1:8080`.
- Upload any Master template file.
- Upload any number of Regional files.
- Click **"Run Data Consolidation"** and download the consolidated workbook.

### Option B: Command-Line Interface (CLI)
Consolidate from terminal or automated pipelines:
```bash
python cli.py --master "sample_files/Master_Blank_Template.xlsx" --regions "sample_files/303_Water_*.xlsx" --output "Consolidated_Water_Report.xlsx"
```
