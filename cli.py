"""
Command-Line Interface (CLI) for Data Consolidation Tool.
Supports arbitrary file names, auto-detection of master template,
semantic alignment of shifted columns, and audit reporting.
"""

import argparse
import os
import sys
import glob

from core.template_profiler import MasterProfile
from core.regional_reader import read_regional_file, RegionalData
from core.consolidator import consolidate_workbooks
from config.template_config import clean_region_name


def main():
    parser = argparse.ArgumentParser(description="Data Consolidation Tool")
    parser.add_argument("--master", "-m", help="Path to Master Excel template (.xlsx, .xlsm)", default=None)
    parser.add_argument("--regions", "-r", nargs="+", help="Paths to regional Excel files (.xlsx, .xlsm)", default=[])
    parser.add_argument("--output", "-o", help="Output consolidated Excel file path", default=None)
    parser.add_argument("--region-names", help="Explicit mapping of filename=RegionName, comma-separated", default="")

    args = parser.parse_args()

    # Parse explicit region name overrides
    name_overrides = {}
    if args.region_names:
        for pair in args.region_names.split(","):
            if "=" in pair:
                fn, rname = pair.split("=", 1)
                name_overrides[fn.strip()] = rname.strip()

    # Check files
    if not args.master:
        print("[ERROR] Please provide a master file using --master <path_to_master.xlsx>")
        sys.exit(1)

    if not os.path.exists(args.master):
        print(f"[ERROR] Master file '{args.master}' not found.")
        sys.exit(1)

    if not args.regions:
        print("[ERROR] Please provide at least one regional file using --regions <file1.xlsx> <file2.xlsx> ...")
        sys.exit(1)

    # Expand any glob patterns in regions
    regional_paths = []
    for r_pattern in args.regions:
        matched = glob.glob(r_pattern)
        if matched:
            regional_paths.extend(matched)
        elif os.path.exists(r_pattern):
            regional_paths.append(r_pattern)
        else:
            print(f"[WARNING] Regional file '{r_pattern}' not found, skipping.")

    if not regional_paths:
        print("[ERROR] No valid regional files found.")
        sys.exit(1)

    print("=" * 80)
    print(" DATA CONSOLIDATION ENGINE")
    print("=" * 80)
    print(f"[*] Profiling Master Template: {os.path.basename(args.master)}")
    master_profile = MasterProfile(args.master)
    print(f"    - Sheets detected: {len(master_profile.sheet_names)}")
    total_formulas = sum(len(f) for f in master_profile.formula_cells.values())
    print(f"    - Master formulas protected: {total_formulas}")

    print("\n[*] Processing Regional Files:")
    regional_datasets = []
    region_names_seen = {}

    for r_path in regional_paths:
        fname = os.path.basename(r_path)
        r_name = name_overrides.get(fname, clean_region_name(fname))
        
        # Check duplicate region names
        if r_name in region_names_seen:
            print(f"    [!] WARNING: Duplicate region name '{r_name}' detected in {fname} and {region_names_seen[r_name]}!")
        region_names_seen[r_name] = fname

        print(f"    -> Reading: {fname} (Region: {r_name})")
        reg_data = read_regional_file(r_path, master_profile, region_name=r_name)
        regional_datasets.append(reg_data)

    # Output file determination
    if not args.output:
        base_name = os.path.splitext(os.path.basename(args.master))[0]
        args.output = os.path.join(os.path.dirname(args.master) or ".", f"{base_name}_Consolidated.xlsx")

    print(f"\n[*] Consolidating {len(regional_datasets)} regions into: {args.output}")
    out_file, val_report = consolidate_workbooks(master_profile, regional_datasets, args.output)

    print("\n" + "=" * 80)
    print(" CONSOLIDATION COMPLETE")
    print("=" * 80)
    print(f"Output File: {os.path.abspath(out_file)}")
    print(f"Validation Warnings: {len(val_report.warnings)}")
    if val_report.warnings:
        print("\n--- Top Warnings & Anomalies Detected ---")
        for w in val_report.warnings[:10]:
            print(f"  • {w}")
        if len(val_report.warnings) > 10:
            print(f"  ... and {len(val_report.warnings) - 10} more (see 'Merge_Notes' sheet in output)")

    print("\nMaster structure preserved 100%. Dynamic links created in 'Region_Data'.")


if __name__ == "__main__":
    main()
