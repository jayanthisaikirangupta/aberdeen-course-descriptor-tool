# Aberdeen Course Descriptor Builder

A small program that **fetches University of Aberdeen "Catalogue of Courses" pages
live, parses them, and compiles a Word `.docx` "Module Descriptors" document** in
the registry template format.

Everything lives in this one folder and runs on your own machine. Nothing is sent
anywhere except the requests to `abdn.ac.uk` to read the course pages.

---

## What you need (once)

1. **Python 3.9 or newer** — https://www.python.org/downloads/
   (On Windows, tick *"Add Python to PATH"* in the installer.)
2. The Python packages listed in `requirements.txt` (installed automatically by
   the start scripts, or manually with `pip install -r requirements.txt`).

---

## How to run it

### Option A — Web app (easiest, recommended for colleagues)

- **Windows:** double-click **`start-windows.bat`**
- **macOS:** double-click **`start-mac.command`**
  (first time: right-click → Open, to get past Gatekeeper)
- **Any system, from a terminal:**
  ```
  pip install -r requirements.txt
  python app.py
  ```

Then open **http://localhost:5000** in your browser. Paste the course codes, set
the level / year, fill the cover-page details, and click **Generate & download**.
The Word document downloads to your computer.

Use **Check URLs & fetch** first if you just want to see which courses resolve
before building the document. Any course that fails shows a box where you can
**paste the correct catalogue URL** and re-run just by clicking again.

### Option B — Command line (for batch / scripted use)

1. Edit **`config.json`** — set `courses`, `level`, `year`, the `cover` details,
   and the `prefix_map`.
2. Run:
   ```
   python descriptor_builder.py
   ```
3. The document is written to **`output/Module_Descriptors_<student>.docx`**.

---

## How the URL is built

```
https://www.abdn.ac.uk/registry/courses/<level>/<year>/<subject>/<code>
```

- `<level>`  — `undergraduate` or `postgraduate`
- `<year>`   — e.g. `2024-2025`
- `<subject>`— looked up from `prefix_map` using the **letters** of the code
               (e.g. `AN3009` → prefix `AN` → `Anatomy`)
- `<code>`   — the course code

**Casing matters.** The catalogue is case-sensitive, so:
- the **subject** is used exactly as written in `prefix_map` (`Anatomy`, not `anatomy`);
- the **code** is tried as you typed it *and* upper-cased as a fallback.

If a page can't be found, supply a **manual URL** — in the web app's error box, or
in `config.json` under `manual_urls`, e.g.:

```json
"manual_urls": {
  "AN3009": "https://www.abdn.ac.uk/registry/courses/undergraduate/2024-2025/Anatomy/AN3009"
}
```

---

## What gets extracted from each page

| Document field            | Source on the catalogue page                        |
| ------------------------- | --------------------------------------------------- |
| Code & Title              | the page `<h1>` (e.g. *AN3009: Architecture Of Life*) |
| Credit Points / ECTS      | the *Course Details* table ("Credit Points" row)    |
| Course Coordinator(s)     | the *Course Details* table ("Co-ordinators" row)    |
| Course Overview           | the *Course Overview* section                       |
| Course Description        | the *Course Description* section                    |
| Summative / Resit         | assessment names + weightings under each heading    |
| Formative Assessment      | the *Formative Assessment* section                  |

---

## Files

```
course_descriptor_tool/
├─ app.py                 # local web app (Flask)
├─ descriptor_builder.py  # core: fetch + parse + Word generation (+ CLI)
├─ config.json            # your inputs for command-line use
├─ requirements.txt       # Python dependencies
├─ templates/index.html   # web app page
├─ assets/                # University of Aberdeen logo (used in the document header)
├─ output/                # generated .docx files appear here (CLI)
├─ start-windows.bat
└─ start-mac.command
```

---

## Notes & limitations

- The parser targets the catalogue page layout as of 2025. If the University
  changes the page structure, the section-matching in `descriptor_builder.py`
  (functions `_section_lines`, `_details_table`, `_assessment_items`) may need a
  small update — they locate content by heading text, which is fairly robust.
- Please scrape responsibly: fetch only the courses you need.
- The document uses **Calibri** to match the original Word template.
