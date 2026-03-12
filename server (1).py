"""
HBS QA Dashboard — Backend Server
Deployed on Render.com | Jira + n8n edition

Data priority:
  1. data/jira_data.json   — written by n8n POST to /api/sync
  2. data/HBS_QA_Data.xlsx — manual fallback if JSON missing
"""

import os, json
from datetime import datetime, date
from flask import Flask, jsonify, send_from_directory, request
from flask_cors import CORS
from openpyxl import load_workbook

app = Flask(__name__, static_folder=".")
CORS(app)

DATA_DIR  = os.path.join(os.path.dirname(__file__), "data")
JSON_FILE = os.path.join(DATA_DIR, "jira_data.json")
XLSX_FILE = os.path.join(DATA_DIR, "HBS_QA_Data.xlsx")

os.makedirs(DATA_DIR, exist_ok=True)

def safe_str(v):
    if v is None: return ""
    if isinstance(v, (datetime, date)): return v.strftime("%Y-%m-%d")
    return str(v).strip()

def safe_int(v, default=0):
    try:    return int(v) if v is not None else default
    except: return default

def safe_bool(v):
    if isinstance(v, bool): return v
    if isinstance(v, str):  return v.strip().upper() in ("TRUE","YES","1")
    return bool(v) if v else False

def xlsx_rows(ws, start=5):
    for row in ws.iter_rows(min_row=start, values_only=True):
        if any(c is not None and str(c).strip() != "" for c in row):
            yield row

def load_from_json():
    with open(JSON_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    if "programmes" not in data:
        raise ValueError("jira_data.json missing 'programmes' key")
    data["source"] = "jira"
    data["file_modified"] = datetime.fromtimestamp(os.path.getmtime(JSON_FILE)).isoformat()
    return data

def load_data():
    if os.path.exists(JSON_FILE):
        try:
            return load_from_json()
        except Exception as e:
            print(f"JSON load failed ({e})")
    raise FileNotFoundError("No data source found — trigger n8n sync first")

# ── Routes ────────────────────────────────────────────────

@app.route("/api/data")
def api_data():
    try:
        return jsonify(load_data())
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 404
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/api/sync", methods=["POST"])
def api_sync():
    """n8n cloud POSTs transformed Jira data here"""
    try:
        payload = request.get_json(force=True)
        if not payload or "programmes" not in payload:
            return jsonify({"ok": False, "error": "Invalid payload"}), 400

        payload["synced_at"] = datetime.now().isoformat()
        payload["source"]    = "jira"

        with open(JSON_FILE, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2, ensure_ascii=False)

        prog_count  = len(payload.get("programmes", {}))
        issue_count = sum(
            len(p.get("ipd",[])) + len(p.get("prd",[])) +
            len(p.get("br",[])) + len(p.get("tc",[]))
            for p in payload.get("programmes", {}).values()
        )
        print(f"Sync — {prog_count} programmes, {issue_count} issues @ {payload['synced_at']}")
        return jsonify({"ok": True, "synced_at": payload["synced_at"], "programmes": prog_count})

    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/api/status")
def api_status():
    j = os.path.exists(JSON_FILE)
    active = JSON_FILE if j else None
    return jsonify({
        "source":    "jira" if j else "none",
        "jira_json": j,
        "modified":  datetime.fromtimestamp(os.path.getmtime(active)).isoformat() if active else None,
    })

@app.route("/")
def serve_dashboard():
    return send_from_directory(".", "dashboard.html")

@app.route("/<path:filename>")
def serve_static(filename):
    return send_from_directory(".", filename)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print("=" * 54)
    print("  HBS QA Dashboard — Render.com edition")
    print(f"  Running on port {port}")
    print("=" * 54)
    app.run(host="0.0.0.0", port=port, debug=False)
