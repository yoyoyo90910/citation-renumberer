# Citation Tools

A local tool for editorial staff with **two modes you flip between** in the
header:

- **Renumber** — take the manual work out of renumbering references when
  clients or authors add, remove, or reorder citations (numbered styles:
  superscript `¹²`, brackets `[12]`, parentheses `(12)`).
- **Convert author-date** — turn `(Smith, 2020)` style citations into numbered
  superscripts, with a review-before-you-commit step so you approve every match.

Import a document once; flipping modes re-reads the same file for that mode.

**Mendeley documents** are handled automatically: if you import a file written
with the Mendeley Cite add-in (even with an ungenerated bibliography), the tool
reads the citation data straight from the Mendeley fields, flattens every
citation to a numbered superscript, writes the numbered reference list, and
opens in Renumber mode — no need to flatten in Word first.

## How to run

Double-click **`Start Citation Renumberer.bat`**. The first run sets itself up
(creates a `.venv`, installs requirements, builds the demo article); after that
it opens in your browser at `http://127.0.0.1:5057/`. Close the black window to
stop it.

Everything runs on your own machine. Nothing is uploaded, and your original
`.docx` is never modified.

## What it does

1. **Import** a `.docx` article (client or non-client — same engine).
2. **Detect** in-text citations. Numbered citations are supported in any of
   three styles — superscript `¹²`, brackets `[12]`, or parentheses `(12)` —
   found wherever they appear: body text, tables, figure legends, and
   footnotes. Ranges (`14–16`) and lists (`12,13`) are handled. The style is
   **auto-detected**, with a dropdown to correct it if needed. (Author–date
   styles like Harvard/APA `(Smith, 2023)` aren't numbered, so renumbering
   doesn't apply.)
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
