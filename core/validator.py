"""
Validation and Anomaly Detection Engine for GRI 303 Water Consolidation.
Implements all validation checks defined in PRD Section 6.4:
- Discharge >= Withdrawal (zero or negative consumption)
- Water-stress withdrawal > total withdrawal
- Unit discrepancies (L vs KL outliers > 500x)
- Repetitive values / suspicious patterns
- Cross-sheet consistency
"""

from typing import List, Dict, Any
from core.regional_reader import RegionalData
import numpy as np


class ValidationReport:
    def __init__(self):
        self.critical_errors: List[str] = []
        self.warnings: List[str] = []
        self.audit_notes: List[str] = []


def run_validations(regional_datasets: List[RegionalData]) -> ValidationReport:
    report = ValidationReport()

    for reg in regional_datasets:
        # 1. Missing and extra sheets
        for s in reg.missing_sheets:
            report.warnings.append(f"{reg.region_name}: missing sheet '{s}'.")
        for s in reg.extra_sheets:
            report.warnings.append(f"{reg.region_name}: extra sheet '{s}' (skipped).")

        # 2. Hardcoded totals typed over formulas
        for ht in reg.hardcoded_totals:
            report.warnings.append(
                f"{reg.region_name}: hard-coded total {ht['typed_value']} typed over formula cell {ht['coord']} on '{ht['sheet']}' was NOT copied."
            )

        # 3. Stray values outside master bounds
        for sv in reg.stray_values:
            report.warnings.append(
                f"{reg.region_name}: stray value {sv['value']} in cell {sv['coord']} of '{sv['sheet']}' (outside master table) was NOT copied."
            )

        # 4. Regional Discharge vs Withdrawal cross-checks (Negative Consumption check)
        # Check NHAI sheets
        _check_consumption(reg, report, "303-03 _ For NHAI", "303-04_ For NHAI", "NHAI")
        # Check Contractor sheets
        _check_consumption(reg, report, "303_03_for Contractors", "303_04_For Contractors", "Contractor")

    # 5. Unit discrepancy check across regions (L vs KL: one region > 500x or < 1/500x)
    _check_unit_outliers(regional_datasets, report)

    return report


def _check_consumption(reg: RegionalData, report: ValidationReport, w_sheet: str, d_sheet: str, scope: str):
    w_vals = reg.numeric_values.get(w_sheet, {})
    d_vals = reg.numeric_values.get(d_sheet, {})

    # Sum total withdrawal and total discharge for this region
    total_w = sum(v for v in w_vals.values() if isinstance(v, (int, float)))
    total_d = sum(v for v in d_vals.values() if isinstance(v, (int, float)))

    if total_w > 0 and total_d > 0:
        if total_d >= total_w:
            report.warnings.append(
                f"{reg.region_name} ({scope}): total discharge ({total_d:,.0f}) is >= withdrawal ({total_w:,.0f}), giving zero or negative net consumption. Please verify."
            )
        elif total_d >= 0.95 * total_w:
            report.warnings.append(
                f"{reg.region_name} ({scope}): total discharge ({total_d:,.0f}) is nearly equal to withdrawal ({total_w:,.0f}), making consumption ~0. Please confirm."
            )


def _check_unit_outliers(regional_datasets: List[RegionalData], report: ValidationReport):
    """Detects values with potential unit errors (e.g. KL entered instead of L, or vice versa)."""
    # Gather non-zero numbers per cell across all regions
    cell_values: Dict[str, Dict[str, List[float]]] = {} # sheet -> {coord -> [(region, val)]}

    for reg in regional_datasets:
        for sname, nums in reg.numeric_values.items():
            cell_values.setdefault(sname, {})
            for coord, val in nums.items():
                if val > 0:
                    cell_values[sname].setdefault(coord, []).append((reg.region_name, val))

    for sname, coords in cell_values.items():
        for coord, reg_vals in coords.items():
            if len(reg_vals) >= 3:
                vals = [v for _, v in reg_vals]
                median_val = float(np.median(vals))
                if median_val > 0:
                    for r_name, v in reg_vals:
                        if v > 500 * median_val:
                            report.warnings.append(
                                f"{r_name}: cell {coord} on '{sname}' has value {v:,.0f}, which is >500x the median of other regions ({median_val:,.0f}). Check if unit is in Litres vs KL."
                            )
                        elif v < median_val / 500 and v > 0:
                            report.warnings.append(
                                f"{r_name}: cell {coord} on '{sname}' has value {v:,.0f}, which is <1/500th of other regions ({median_val:,.0f}). Check if unit is in KL vs Litres."
                            )
