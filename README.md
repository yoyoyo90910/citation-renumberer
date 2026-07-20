# Citation Renumberer

A local tool for editorial staff to take the manual work out of renumbering
references when clients or authors add, remove, or reorder citations.

## How to run

Double-click **`Start Citation Renumberer.bat`**. The first run sets itself up
(creates a `.venv`, installs requirements, builds the demo article); after that
it opens in your browser at `http://127.0.0.1:5057/`. Close the black window to
stop it.

Everything runs on your own machine. Nothing is uploaded, and your original
`.docx` is never modified.

## What it does

1. **Import** a `.docx` article (client or non-client — same engine).
2. **Detect** in-text citations. Citations are **superscript numbers**, and are
   found wherever they appear — body text, tables, figure legends, and
   footnotes. Ranges (`14–16`) and lists (`12,13`) are handled.
3. **Parse** the numbered reference list.
4. **Link** the two: click a citation to spotlight its reference, or a
   reference to spotlight everywhere it is cited.
5. **Flag** problems: citations with no matching reference, and references that
   are never cited.
6. **Re-reference** — reorder references (drag the handle or use ↑↓), delete
   one, or add a new one. Both the in-text citations *and* the reference list
   renumber automatically and update live in the preview. Ranges recompress
   (`14,15,16` → `14–16`).
   - **Auto-renumber by text order** applies true Vancouver numbering (by order
     of first appearance in the text) in one click.
7. **Export** a fresh renumbered `.docx` to the `exports/` folder, with a
   change log (`13 → 14`, …). The original file is never modified.

## Coming next

- EMJ/AMJ polish and testing against real production articles.
- Confirming how table/figure parenthetical citations should rank in
  first-appearance order for auto mode.

## Deploying (Render)

The app is stateless and needs no secrets, so it can be hosted for the team to
share:

1. Push this folder to GitHub.
2. On [render.com](https://render.com), create a **Web Service** from the repo.
   `render.yaml` is already set up (`gunicorn -w 2 --timeout 120 app:app`).
3. Open the Render URL.

Uploaded files are given a random token, live only on disk for the session, and
are never persisted or shared. Exported files are transient. Because there are
no secrets, the repository is safe to be public.

## Files

| File | Purpose |
|------|---------|
| `app.py` | Local web server (Flask) |
| `citations.py` | Detection + linking engine (deterministic, no AI guessing) |
| `make_sample.py` | Builds the demo article in `samples/` |
| `templates/index.html` | The browser interface |
| `Start Citation Renumberer.bat` | Double-click launcher |

## A note on accuracy

Citation detection is deliberately rule-based, not model-guesswork — a wrong
reference number is a real editorial error. Citations are identified purely by
the superscript formatting flag in the Word file, so ordinary numbers on the
baseline (years, doses, sample sizes) are never mistaken for citations. Always
review the change log before sending anything to a client.
