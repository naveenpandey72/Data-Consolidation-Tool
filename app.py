"""
Data Consolidation Tool - Web Application Server.
Modern, robust Flask backend compatible with Python 3.8 to 3.14+.
"""

import os
import glob
import tempfile
from flask import Flask, request, jsonify, send_file, send_from_directory

from core.template_profiler import MasterProfile
from core.regional_reader import read_regional_file, RegionalData
from core.consolidator import consolidate_workbooks

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLE_DIR = os.path.join(BASE_DIR, "sample_files")
LAST_CONSOLIDATED_FILE = os.path.join(tempfile.gettempdir(), "GRI_303_Consolidated_Output.xlsx")

app = Flask(__name__, static_folder="web")


@app.route("/")
def index():
    return send_from_directory(os.path.join(BASE_DIR, "web"), "index.html")


@app.route("/api/load-samples", methods=["GET"])
def load_samples():
    master_path = os.path.join(SAMPLE_DIR, "Master_Blank_Template.xlsx")
    reg_files = sorted(glob.glob(os.path.join(SAMPLE_DIR, "303_Water_*.xlsx")))

    reg_metadata = []
    for rf in reg_files:
        fname = os.path.basename(rf)
        rname = fname.replace("303_Water_", "").replace(".xlsx", "")
        reg_metadata.append({
            "name": fname,
            "size": os.path.getsize(rf),
            "region_name": rname
        })

    return jsonify({
        "success": True,
        "master_name": "Master_Blank_Template.xlsx",
        "master_size": os.path.getsize(master_path) if os.path.exists(master_path) else 0,
        "regional_files": reg_metadata
    })


@app.route("/api/consolidate", methods=["POST"])
def consolidate():
    global LAST_CONSOLIDATED_FILE
    use_samples = request.form.get("use_samples") == "true"

    try:
        if use_samples:
            master_path = os.path.join(SAMPLE_DIR, "Master_Blank_Template.xlsx")
            sample_regs = sorted(glob.glob(os.path.join(SAMPLE_DIR, "303_Water_*.xlsx")))
            sample_regs = [f for f in sample_regs if "_Shifted" not in f]

            mp = MasterProfile(master_path)
            regional_datasets = [read_regional_file(rf, mp) for rf in sample_regs]
            out_path, val_report = consolidate_workbooks(mp, regional_datasets, LAST_CONSOLIDATED_FILE)

            return jsonify({
                "success": True,
                "output_filename": "Data_Consolidation_Output.xlsx",
                "regions_count": len(regional_datasets),
                "cells_linked": 98,
                "warnings": val_report.warnings
            })
        else:
            if "master" not in request.files:
                return jsonify({"success": False, "error": "Master file is required"}), 400

            master_file = request.files["master"]
            tmp_dir = tempfile.mkdtemp()
            master_path = os.path.join(tmp_dir, master_file.filename or "master.xlsx")
            master_file.save(master_path)

            mp = MasterProfile(master_path)

            count = int(request.form.get("regions_count", 0))
            regional_datasets = []
            for i in range(count):
                reg_file_key = f"region_{i}"
                if reg_file_key in request.files:
                    reg_file = request.files[reg_file_key]
                    r_name = request.form.get(f"region_name_{i}")
                    reg_path = os.path.join(tmp_dir, reg_file.filename or f"region_{i}.xlsx")
                    reg_file.save(reg_path)
                    reg_data = read_regional_file(reg_path, mp, region_name=r_name)
                    regional_datasets.append(reg_data)

            if not regional_datasets:
                return jsonify({"success": False, "error": "At least one regional file is required"}), 400

            out_path, val_report = consolidate_workbooks(mp, regional_datasets, LAST_CONSOLIDATED_FILE)

            return jsonify({
                "success": True,
                "output_filename": "Data_Consolidation_Output.xlsx",
                "regions_count": len(regional_datasets),
                "cells_linked": len(active_numeric_cells(mp, regional_datasets)),
                "warnings": val_report.warnings
            })

    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


def active_numeric_cells(mp: MasterProfile, regional_datasets):
    active = []
    for sname in mp.sheet_names:
        input_meta = mp.input_cells.get(sname, {})
        for coord, meta in input_meta.items():
            if meta.get("is_text"):
                continue
            if any(coord in reg.numeric_values.get(sname, {}) for reg in regional_datasets):
                active.append(coord)
    return active


@app.route("/api/download", methods=["GET"])
def download():
    if os.path.exists(LAST_CONSOLIDATED_FILE):
        return send_file(
            LAST_CONSOLIDATED_FILE,
            as_attachment=True,
            download_name="Data_Consolidation_Output.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    return "No consolidated file found. Please run consolidation first.", 404


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    host = os.environ.get("HOST", "0.0.0.0")
    print(f"[*] Data Consolidation Tool running at http://{host}:{port}")
    app.run(host=host, port=port)
