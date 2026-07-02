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


def _normalise_year_groups(cfg):
    """Accept legacy {year, courses} config and return a list of year groups."""
    groups = cfg.get("years")
    if groups:
        return [
            {
                "year": g.get("year", "").strip(),
                "courses": [c.strip() for c in g.get("courses", []) if c and c.strip()],
            }
            for g in groups
            if g.get("year")
        ]
    if cfg.get("year"):
        return [{
            "year": cfg["year"],
            "courses": [c.strip() for c in cfg.get("courses", []) if c and c.strip()],
        }]
    return []


@app.route("/")
def index():
    cfg = load_config()
    year_groups = _normalise_year_groups(cfg)
    if not year_groups:
        year_groups = [{"year": "2024-2025", "courses": []}]
    return render_template(
        "index.html",
        cfg=cfg,
        year_groups=year_groups,
        year_options=["2021-2022", "2022-2023", "2023-2024", "2024-2025", "2025-2026"],
        prefix_text="\n".join(f"{k} = {v}" for k, v in cfg.get("prefix_map", {}).items()),
    )


def _parse_form(form):
    years_raw = form.getlist("years[]")
    courses_raw = form.getlist("courses[]")
    year_groups = []
    for i, y in enumerate(years_raw):
        y = (y or "").strip()
        if not y:
            continue
        raw = courses_raw[i] if i < len(courses_raw) else ""
        codes = [c.strip() for c in raw.replace(",", "\n").splitlines() if c.strip()]
        if codes:
            year_groups.append({"year": y, "courses": codes})
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
        "years": year_groups,
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
                     "ects": c["ects"], "url": c["url"],
                     "year": c.get("year", "")} for c in courses],
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
    doc = core.build_document(config["cover"], courses, config["years"])
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
