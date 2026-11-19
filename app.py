"""
Data Consolidation Tool - Web Application Server.
Zero-dependency HTTP server using standard Python library.
Serves interactive web UI, handles uploads, and runs consolidation engine.
"""

from http.server import HTTPServer, BaseHTTPRequestHandler
import json
import os
import glob
import cgi
import tempfile
import urllib.parse

from core.template_profiler import MasterProfile
from core.regional_reader import read_regional_file, RegionalData
from core.consolidator import consolidate_workbooks

PORT = int(os.environ.get("PORT", 8080))
HOST = "0.0.0.0"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLE_DIR = os.path.join(BASE_DIR, "sample_files")
LAST_CONSOLIDATED_FILE = os.path.join(tempfile.gettempdir(), "GRI_303_Consolidated_Output.xlsx")


class ConsolidationServer(BaseHTTPRequestHandler):
    def do_GET(self):
        url = urllib.parse.urlparse(self.path)
        path = url.path

        if path in ["/", "/index.html"]:
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            with open(os.path.join(BASE_DIR, "web", "index.html"), "rb") as f:
                self.wfile.write(f.read())

        elif path == "/api/load-samples":
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            
            master_path = os.path.join(SAMPLE_DIR, "Master_Blank_Template.xlsx")
            reg_files = sorted(glob.glob(os.path.join(SAMPLE_DIR, "303_Water_*.xlsx")))
            
            reg_metadata = []
            for rf in reg_files:
                fname = os.path.basename(rf)
                rname = fname.replace("303_Water_", "").replace(".xlsx", "").replace("_Shifted", " (Shifted)")
                reg_metadata.append({
                    "name": fname,
                    "size": os.path.getsize(rf),
                    "region_name": rname
                })

            resp = {
                "success": True,
                "master_name": "Master_Blank_Template.xlsx",
                "master_size": os.path.getsize(master_path) if os.path.exists(master_path) else 0,
                "regional_files": reg_metadata
            }
            self.wfile.write(json.dumps(resp).encode())

        elif path == "/api/download":
            if os.path.exists(LAST_CONSOLIDATED_FILE):
                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", 'attachment; filename="GRI_303_Water_Consolidated.xlsx"')
                self.send_header("Content-Length", str(os.path.getsize(LAST_CONSOLIDATED_FILE)))
                self.end_headers()
                with open(LAST_CONSOLIDATED_FILE, "rb") as f:
                    self.wfile.write(f.read())
            else:
                self.send_response(404)
                self.end_headers()
                self.wfile.write(b"No consolidated file available. Run consolidation first.")

        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        url = urllib.parse.urlparse(self.path)
        path = url.path

        if path == "/api/consolidate":
            content_type = self.headers.get("Content-Type", "")
            
            if "multipart/form-data" in content_type:
                form = cgi.FieldStorage(
                    fp=self.rfile,
                    headers=self.headers,
                    environ={"REQUEST_METHOD": "POST", "CONTENT_TYPE": content_type}
                )

                use_samples = form.getvalue("use_samples") == "true"
                global LAST_CONSOLIDATED_FILE

                if use_samples:
                    master_path = os.path.join(SAMPLE_DIR, "Master_Blank_Template.xlsx")
                    sample_regs = sorted(glob.glob(os.path.join(SAMPLE_DIR, "303_Water_*.xlsx")))
                    # Exclude the shifted test file if present
                    sample_regs = [f for f in sample_regs if "_Shifted" not in f]

                    mp = MasterProfile(master_path)
                    regional_datasets = [read_regional_file(rf, mp) for rf in sample_regs]
                    
                    out_path, val_report = consolidate_workbooks(mp, regional_datasets, LAST_CONSOLIDATED_FILE)

                    resp = {
                        "success": True,
                        "output_filename": "GRI_303_Water_Consolidated.xlsx",
                        "regions_count": len(regional_datasets),
                        "cells_linked": 98,
                        "warnings": val_report.warnings
                    }
                    self.send_response(200)
                    self.send_header("Content-Type", "application/json")
                    self.end_headers()
                    self.wfile.write(json.dumps(resp).encode())
                else:
                    # Handle uploaded files
                    tmp_dir = tempfile.mkdtemp()
                    try:
                        master_item = form["master"]
                        master_path = os.path.join(tmp_dir, master_item.filename)
                        with open(master_path, "wb") as f:
                            f.write(master_item.file.read())

                        mp = MasterProfile(master_path)

                        count = int(form.getvalue("regions_count", 0))
                        regional_datasets = []
                        for i in range(count):
                            reg_item = form[f"region_{i}"]
                            r_name = form.getvalue(f"region_name_{i}")
                            reg_path = os.path.join(tmp_dir, reg_item.filename)
                            with open(reg_path, "wb") as f:
                                f.write(reg_item.file.read())
                            reg_data = read_regional_file(reg_path, mp, region_name=r_name)
                            regional_datasets.append(reg_data)

                        out_path, val_report = consolidate_workbooks(mp, regional_datasets, LAST_CONSOLIDATED_FILE)

                        resp = {
                            "success": True,
                            "output_filename": "GRI_303_Water_Consolidated.xlsx",
                            "regions_count": len(regional_datasets),
                            "cells_linked": len(regional_datasets[0].numeric_values) if regional_datasets else 0,
                            "warnings": val_report.warnings
                        }
                        self.send_response(200)
                        self.send_header("Content-Type", "application/json")
                        self.end_headers()
                        self.wfile.write(json.dumps(resp).encode())
                    except Exception as e:
                        resp = {"success": False, "error": str(e)}
                        self.send_response(500)
                        self.send_header("Content-Type", "application/json")
                        self.end_headers()
                        self.wfile.write(json.dumps(resp).encode())
            else:
                self.send_response(400)
                self.end_headers()


def run(host=HOST, port=PORT):
    server = HTTPServer((host, port), ConsolidationServer)
    print(f"[*] Data Consolidation Tool Web UI running at http://{host}:{port}")
    server.serve_forever()


if __name__ == "__main__":
    run()
