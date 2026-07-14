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
import re
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
        year_options=[f"{y}-{y+1}" for y in range(2025, 2002, -1)],
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


_YEAR_RE = re.compile(r"^\d{4}-\d{4}$")
_CODE_RE = re.compile(r"^[A-Za-z]{1,5}\d{3,5}$")
_PREFIX_LINE_RE = re.compile(r"^[A-Za-z]+\s*=\s*\S.*$")


def _validate_config(config):
    """Return a list of human-readable validation errors (empty list = OK)."""
    errors = []

    cover = config.get("cover", {}) or {}
    for key, label in (("student", "Student name"), ("degree", "Degree title"),
                        ("compiled_by", "Compiled by"), ("date", "Date"),
                        ("position", "Position")):
        if not str(cover.get(key, "") or "").strip():
            errors.append(f"{label} is required.")

    year_groups = config.get("years") or []
    if not year_groups:
        errors.append("Add at least one academic year with course codes.")
    for i, g in enumerate(year_groups, start=1):
        label = f"Year {i}"
        yr = (g.get("year") or "").strip()
        if not yr:
            errors.append(f"{label}: pick an academic year.")
        elif not _YEAR_RE.match(yr):
            errors.append(f"{label}: academic year must be in YYYY-YYYY format (got '{yr}').")
        codes = [c for c in (g.get("courses") or []) if str(c).strip()]
        if not codes:
            errors.append(f"{label}: enter at least one course code.")
        bad = [c for c in codes if not _CODE_RE.match(c)]
        if bad:
            errors.append(f"{label}: invalid course code(s): {', '.join(bad)} (expected letters+digits, e.g. PS2517).")

    prefix_map = config.get("prefix_map") or {}
    if not prefix_map:
        errors.append("Prefix → subject mapping is empty. Add at least one PREFIX = subject line.")
    else:
        map_keys = {str(k).strip().upper() for k in prefix_map.keys() if str(k).strip()}
        missing = set()
        for g in year_groups:
            for c in (g.get("courses") or []):
                m = re.match(r"^([A-Za-z]+)", str(c))
                if m and m.group(1).upper() not in map_keys:
                    missing.add(m.group(1).upper())
        if missing:
            example = sorted(missing)[0]
            errors.append(
                f"No prefix mapping for: {', '.join(sorted(missing))}. "
                f"Add a line like '{example} = subject_path'."
            )

    return errors


def _validation_response(errors):
    return jsonify({
        "ok": False,
        "validation_errors": errors,
        "message": "Please fix the form fields before continuing.",
    }), 422


@app.route("/save_prefix_map", methods=["POST"])
def save_prefix_map():
    """Persist the Prefix → subject mapping to config.json."""
    text = request.form.get("prefix_map", "")

    prefix_map, bad_lines = {}, []
    for idx, raw in enumerate(text.splitlines(), start=1):
        line = raw.strip()
        if not line:
            continue
        if not _PREFIX_LINE_RE.match(line):
            bad_lines.append(idx)
            continue
        k, v = line.split("=", 1)
        prefix_map[k.strip().upper()] = v.strip()

    errors = []
    if not prefix_map and not bad_lines:
        errors.append("Prefix → subject mapping is empty. Add at least one PREFIX = subject line.")
    if bad_lines:
        errors.append(
            f"Line{'s' if len(bad_lines) > 1 else ''} {', '.join(str(n) for n in bad_lines)}: "
            f"use format PREFIX = subject_path."
        )
    if errors:
        return _validation_response(errors)

    path = os.path.join(HERE, "config.json")
    try:
        with open(path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
    except (OSError, json.JSONDecodeError):
        cfg = {}
    cfg["prefix_map"] = prefix_map
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)
        f.write("\n")
    return jsonify({"ok": True, "prefix_map": prefix_map, "count": len(prefix_map)})


@app.route("/preview", methods=["POST"])
def preview():
    """Fetch + parse only; return per-course status as JSON (no document)."""
    config = _parse_form(request.form)
    vld = _validate_config(config)
    if vld:
        return _validation_response(vld)
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
    vld = _validate_config(config)
    if vld:
        return _validation_response(vld)
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
