"""
Citation Renumberer -- local + deployable web app.

Detect in-text citations, link them to the reference list, then reorder /
insert / delete references and have everything renumber automatically, with a
fresh exported .docx.

Runs locally (double-click the .bat) or on a shared host such as Render. There
are no secrets and nothing is stored long-term: uploaded files get a random
token, live only on disk for the session, and the original is never modified.
"""

import os
import re
import threading
import uuid
import webbrowser

from flask import Flask, jsonify, render_template, request, send_file

import citations

HERE = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(HERE, "uploads")
EXPORT_DIR = os.path.join(HERE, "exports")
SAMPLE = os.path.join(HERE, "samples", "sample-article.docx")
PORT = int(os.environ.get("PORT", "5057"))

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(EXPORT_DIR, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 40 * 1024 * 1024  # 40 MB is plenty for a .docx

_TOKEN_RE = re.compile(r"^[a-f0-9]{32}$")


def _resolve_source(token):
    """Map a token to a .docx path. 'sample' -> the demo; else uploads/<token>."""
    if token == "sample":
        return SAMPLE if os.path.exists(SAMPLE) else None
    if token and _TOKEN_RE.match(token):
        path = os.path.join(UPLOAD_DIR, token + ".docx")
        return path if os.path.exists(path) else None
    return None


def _safe_base(filename):
    name = os.path.basename(filename or "document.docx")
    name = re.sub(r"\s*\(demo\)\s*", "", name)
    name = re.sub(r"[^A-Za-z0-9._ -]", "_", name)
    return os.path.splitext(name)[0] or "document"


def to_dict(a, token):
    return {
        "token": token,
        "filename": a.filename,
        "blocks": a.blocks,
        "occurrences": a.occurrences,
        "references": a.references,
        "issues": a.issues,
        "counts": {
            "references": len(a.references),
            "citations": len(a.occurrences),
            "body": sum(1 for o in a.occurrences if o["location"] == "body"),
            "table": sum(1 for o in a.occurrences if o["location"] == "table"),
            "legend": sum(1 for o in a.occurrences if o["location"] == "legend"),
            "footnote": sum(1 for o in a.occurrences if o["location"] == "footnote"),
        },
    }


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/analysis")
def get_analysis():
    """First page load -> the demo article, so there's always something to see."""
    if not os.path.exists(SAMPLE):
        return jsonify({"empty": True})
    a = citations.analyze(SAMPLE, "sample-article.docx (demo)")
    return jsonify(to_dict(a, "sample"))


@app.route("/api/upload", methods=["POST"])
def upload():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"error": "No file received."}), 400
    if not f.filename.lower().endswith(".docx"):
        return jsonify({"error": "Please choose a .docx file."}), 400
    token = uuid.uuid4().hex
    dest = os.path.join(UPLOAD_DIR, token + ".docx")
    f.save(dest)
    try:
        a = citations.analyze(dest, f.filename)
    except Exception as exc:
        return jsonify({"error": "Could not read that document: %s" % exc}), 400
    return jsonify(to_dict(a, token))


@app.route("/api/renumber", methods=["POST"])
def renumber():
    payload = request.get_json(force=True, silent=True) or {}
    source = _resolve_source(payload.get("token"))
    order = payload.get("order")
    if source is None:
        return jsonify({"error": "Session expired -- please re-import your document."}), 400
    if not order:
        return jsonify({"error": "No reference order supplied."}), 400

    out_name = _safe_base(payload.get("filename")) + "__renumbered.docx"
    out_path = os.path.join(EXPORT_DIR, out_name)
    try:
        report = citations.apply_renumber(source, order, out_path)
    except Exception as exc:
        return jsonify({"error": "Renumber failed: %s" % exc}), 400
    return jsonify({
        "ok": True,
        "download": "/download/" + out_name,
        "change_log": report["change_log"],
        "inserted": report["inserted"],
        "deleted": report["deleted"],
        "warnings": report["warnings"],
    })


@app.route("/download/<path:name>")
def download(name):
    safe = os.path.basename(name)
    path = os.path.join(EXPORT_DIR, safe)
    if not os.path.exists(path):
        return "Not found", 404
    return send_file(path, as_attachment=True, download_name=safe)


def _open_browser():
    webbrowser.open("http://127.0.0.1:%d/" % PORT)


if __name__ == "__main__":
    print("\n  Citation Renumberer running at  http://127.0.0.1:%d/\n" % PORT)
    print("  Close this window to stop the tool.\n")
    threading.Timer(1.0, _open_browser).start()
    app.run(host="127.0.0.1", port=PORT, debug=False)
