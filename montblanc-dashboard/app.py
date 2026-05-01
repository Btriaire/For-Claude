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
CONFIG_FILE   = os.path.join(os.path.dirname(__file__), "dashboard_config.json")
ALLOWED_EXTENSIONS = {"xlsx", "xls"}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def _config_matches(config: dict, parser) -> bool:
    for kpi in config.get("kpis", []):
        sheet = kpi.get("sheet")
        if not sheet or sheet not in parser.sheets:
            return False
        cols = parser.sheets[sheet].columns.tolist()
        if kpi.get("value_column") and kpi["value_column"] not in cols:
            return False
        if kpi.get("category_column") and kpi["category_column"] not in cols:
            return False
    return True


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


def _get_parser() -> ExcelParser | None:
    filepath = session.get("current_file")
    if filepath and os.path.exists(filepath):
        return ExcelParser(filepath)
    return None


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

    # Only store small values in session (avoid cookie overflow)
    session["current_file"]  = filepath
    session["period"]        = request.form.get("period", datetime.now().strftime("%B %Y"))
    session["report_title"]  = request.form.get("report_title", "")
    session["key_events"]    = request.form.get("key_events", "")
    session["comments"]      = request.form.get("comments", "")
    session["outcome"]       = request.form.get("outcome", "")
    session["major_points"]  = request.form.get("major_points", "")
    session["help_needed"]   = request.form.get("help_needed", "")
    session["bottom_banner"] = request.form.get("bottom_banner", "")

    parser = ExcelParser(filepath)
    config = load_config()

    # Auto-apply suggestions if no config or config doesn't match this Excel
    if not config["kpis"] or not _config_matches(config, parser):
        config["kpis"] = parser.get_suggestions()
        with open(CONFIG_FILE, "w") as f:
            json.dump(config, f, indent=2)

    return redirect(url_for("dashboard"))


@app.route("/configure")
def configure():
    parser = _get_parser()
    structure   = parser.get_structure()   if parser else {}
    suggestions = parser.get_suggestions() if parser else []
    config      = load_config()
    return render_template("configure.html",
                           structure=structure,
                           suggestions=suggestions,
                           config=config)


@app.route("/api/save_config", methods=["POST"])
def save_config():
    data = request.get_json(force=True)
    with open(CONFIG_FILE, "w") as f:
        json.dump(data, f, indent=2)
    return jsonify({"status": "ok"})


@app.route("/api/hide_kpi/<kpi_id>", methods=["POST"])
def hide_kpi(kpi_id):
    config = load_config()
    for kpi in config["kpis"]:
        if kpi["id"] == kpi_id:
            kpi["visible"] = False
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=2)
    return jsonify({"status": "ok"})


@app.route("/save_config_form", methods=["POST"])
def save_config_form():
    raw = request.form.get("config_json", "{}")
    try:
        data = json.loads(raw)
        with open(CONFIG_FILE, "w") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass
    return redirect(url_for("dashboard"))


@app.route("/dashboard")
def dashboard():
    parser = _get_parser()
    if not parser:
        return redirect(url_for("index"))

    config   = load_config()
    kpi_data = parser.extract_kpi_data(config["kpis"])

    return render_template(
        "dashboard.html",
        data=kpi_data,
        config=config,
        period=session.get("period", ""),
        report_title=session.get("report_title", ""),
        key_events=session.get("key_events", ""),
        comments=session.get("comments", ""),
        outcome=session.get("outcome", ""),
        major_points=session.get("major_points", ""),
        help_needed=session.get("help_needed", ""),
        bottom_banner=session.get("bottom_banner", ""),
    )


@app.route("/export/pptx")
def export_pptx():
    parser = _get_parser()
    if not parser:
        return redirect(url_for("index"))

    config   = load_config()
    kpi_data = parser.extract_kpi_data(config["kpis"])

    exporter = PPTXExporter(config)
    out_path = exporter.generate(
        kpi_data,
        period=session.get("period", ""),
        report_title=session.get("report_title", ""),
        key_events=session.get("key_events", ""),
        comments=session.get("comments", ""),
        outcome=session.get("outcome", ""),
        major_points=session.get("major_points", ""),
        help_needed=session.get("help_needed", ""),
        bottom_banner=session.get("bottom_banner", ""),
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
    port = int(os.environ.get("PORT", 5000))
    from waitress import serve
    print(f"Starting Montblanc Dashboard on http://0.0.0.0:{port}")
    serve(app, host="0.0.0.0", port=port)
