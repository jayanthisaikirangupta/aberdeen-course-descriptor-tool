"""
Local web app for the Aberdeen Course Descriptor Builder.

Run it:
    python app.py
then open http://localhost:5000 in your browser.

Everything happens locally on your machine — paste course codes, set the level,
year and cover-page details, click Generate, and download the .docx.
"""

import os
import io
import json
from flask import Flask, request, render_template, send_file, jsonify

import descriptor_builder as core

HERE = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__)


def load_config():
    with open(os.path.join(HERE, "config.json"), "r", encoding="utf-8") as f:
        return json.load(f)


@app.route("/")
def index():
    cfg = load_config()
    return render_template(
        "index.html",
        cfg=cfg,
        courses_text="\n".join(cfg.get("courses", [])),
        prefix_text="\n".join(f"{k} = {v}" for k, v in cfg.get("prefix_map", {}).items()),
    )


def _parse_form(form):
    courses = [c.strip() for c in form.get("courses", "").replace(",", "\n").splitlines() if c.strip()]
    prefix_map = {}
    for line in form.get("prefix_map", "").splitlines():
        if "=" in line:
            k, v = line.split("=", 1)
            if k.strip() and v.strip():
                prefix_map[k.strip().upper()] = v.strip()
    manual_urls = {}
    for key in form:
        if key.startswith("manual_"):
            code = key[len("manual_"):]
            if form.get(key, "").strip():
                manual_urls[code] = form.get(key).strip()
    config = {
        "level": form.get("level", "undergraduate"),
        "year": form.get("year", "2024-2025"),
        "courses": courses,
        "prefix_map": prefix_map,
        "manual_urls": manual_urls,
        "cover": {
            "student": form.get("student", ""),
            "degree": form.get("degree", ""),
            "compiled_by": form.get("compiled_by", ""),
            "position": form.get("position", ""),
            "date": form.get("date", ""),
        },
    }
    return config


@app.route("/preview", methods=["POST"])
def preview():
    """Fetch + parse only; return per-course status as JSON (no document)."""
    config = _parse_form(request.form)
    courses, errors = core.build_all(config)
    return jsonify({
        "courses": [{"id": c["id"], "title": c["title"], "cp": c["cp"],
                     "ects": c["ects"], "url": c["url"]} for c in courses],
        "errors": errors,
    })


@app.route("/generate", methods=["POST"])
def generate():
    """Fetch, parse, build the .docx and stream it back as a download."""
    config = _parse_form(request.form)
    courses, errors = core.build_all(config)
    if not courses:
        return jsonify({"ok": False, "errors": errors,
                        "message": "No courses could be fetched."}), 422
    doc = core.build_document(config["cover"], courses, config["year"])
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    fname = "Module_Descriptors_" + core.safe_name(config["cover"].get("student")) + ".docx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=fname,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


if __name__ == "__main__":
    print("\n  Aberdeen Course Descriptor Builder")
    print("  Open  ->  http://localhost:5000\n")
    app.run(debug=False, port=5000)
