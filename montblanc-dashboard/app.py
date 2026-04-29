import json
import os
from datetime import datetime

from flask import (Flask, flash, jsonify, redirect, render_template,
                   request, send_file, session, url_for)
from werkzeug.utils import secure_filename

from utils.excel_parser import ExcelParser
from utils.pptx_exporter import PPTXExporter

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "montblanc-dashboard-dev-key-2024")

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
CONFIG_FILE = os.path.join(os.path.dirname(__file__), "dashboard_config.json")
ALLOWED_EXTENSIONS = {"xlsx", "xls"}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def load_config() -> dict:
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE) as f:
            return json.load(f)
    return _default_config()


def _default_config() -> dict:
    return {
        "kpis": [],
        "colors": {
            "primary": "#5B2D8E",
            "secondary": "#C9A227",
            "tertiary": "#6DBF8B",
            "quaternary": "#E8735A",
        },
    }


# ── routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    config = load_config()
    return render_template("index.html", has_config=bool(config["kpis"]))


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        flash("No file selected.", "error")
        return redirect(url_for("index"))

    file = request.files["file"]
    if not file or file.filename == "":
        flash("No file selected.", "error")
        return redirect(url_for("index"))

    if not allowed_file(file.filename):
        flash("Please upload an Excel file (.xlsx or .xls).", "error")
        return redirect(url_for("index"))

    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    session["current_file"] = filepath
    session["period"] = request.form.get("period", datetime.now().strftime("%B %Y"))
    session["comments"] = request.form.get("comments", "")
    session["key_events"] = request.form.get("key_events", "")

    parser = ExcelParser(filepath)
    session["excel_structure"] = json.dumps(parser.get_structure())

    config = load_config()
    if not config["kpis"]:
        return redirect(url_for("configure"))
    return redirect(url_for("dashboard"))


@app.route("/configure")
def configure():
    structure_raw = session.get("excel_structure", "{}")
    structure = json.loads(structure_raw)
    config = load_config()
    return render_template("configure.html", structure=structure, config=config)


@app.route("/api/save_config", methods=["POST"])
def save_config():
    data = request.get_json(force=True)
    with open(CONFIG_FILE, "w") as f:
        json.dump(data, f, indent=2)
    return jsonify({"status": "ok"})


@app.route("/dashboard")
def dashboard():
    filepath = session.get("current_file")
    if not filepath or not os.path.exists(filepath):
        return redirect(url_for("index"))

    config = load_config()
    parser = ExcelParser(filepath)
    kpi_data = parser.extract_kpi_data(config["kpis"])

    return render_template(
        "dashboard.html",
        data=kpi_data,
        config=config,
        period=session.get("period", ""),
        comments=session.get("comments", ""),
        key_events=session.get("key_events", ""),
    )


@app.route("/export/pptx")
def export_pptx():
    filepath = session.get("current_file")
    if not filepath or not os.path.exists(filepath):
        return redirect(url_for("index"))

    config = load_config()
    parser = ExcelParser(filepath)
    kpi_data = parser.extract_kpi_data(config["kpis"])

    exporter = PPTXExporter(config)
    out_path = exporter.generate(
        kpi_data,
        period=session.get("period", ""),
        comments=session.get("comments", ""),
        key_events=session.get("key_events", ""),
    )

    period_slug = session.get("period", "dashboard").replace(" ", "_")
    return send_file(
        out_path,
        as_attachment=True,
        download_name=f"dashboard_{period_slug}.pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


@app.route("/reset_config", methods=["POST"])
def reset_config():
    if os.path.exists(CONFIG_FILE):
        os.remove(CONFIG_FILE)
    return redirect(url_for("configure"))


if __name__ == "__main__":
    app.run(debug=True, port=5000)
