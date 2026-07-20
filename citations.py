"""
Core engine for the Citation Renumberer.

  * analyze()          -- read a .docx, detect citations, parse the reference
                          list, link the two (read-only).
  * apply_renumber()   -- given a desired final reference order, renumber every
                          in-text citation AND the reference list, then save a
                          fresh .docx. The original is never modified.

Detection rule (agreed with the editorial team):
  * Citations are ALWAYS superscript numbers. They can appear anywhere --
    body text, tables, figure legends, and footnotes. There are no
    parenthetical "(n)" citations.
  * Ranges (14-16) and lists (12,13) inside a superscript are handled.

Everything is deterministic format matching -- no model guessing -- because a
wrong reference number is a real editorial error.
"""

import html
import os
import re
import shutil
import zipfile

import docx
from docx.oxml.ns import qn

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# --------------------------------------------------------------------------
# Number parsing / formatting
# --------------------------------------------------------------------------

_DASHES = r"\-‑‒–—−"
_RANGE_RE = re.compile(r"^(\d+)\s*[" + _DASHES + r"]\s*(\d+)$")


def expand_numbers(text):
    """'14-16, 19' -> [14, 15, 16, 19]. Ranges expanded; non-numbers dropped."""
    cleaned = re.sub(r"[^0-9,;" + _DASHES + r"\s]", " ", text)
    out = []
    for part in re.split(r"[,;]", cleaned):
        part = part.strip()
        if not part:
            continue
        m = _RANGE_RE.match(part)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if a <= b and (b - a) < 500:
                out.extend(range(a, b + 1))
            else:
                out.extend([a, b])
        else:
            out.extend(int(n) for n in re.findall(r"\d+", part))
    return out


def format_citation(numbers, comma_sep=",", dash="–"):
    """[13,14,15,19] -> '13–15,19'. Contiguous runs of 3+ become ranges."""
    nums = sorted(set(numbers))
    if not nums:
        return ""
    groups, i = [], 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
            j += 1
        if j - i >= 2:
            groups.append("%d%s%d" % (nums[i], dash, nums[j]))
        else:
            groups.extend(str(n) for n in nums[i:j + 1])
        i = j + 1
    return comma_sep.join(groups)


def _seps_from_raw(raw):
    comma = ", " if ", " in raw else ","
    dash = "–"
    for ch in raw:
        if ch in "-‑‒–—−":
            dash = ch
            break
    return comma, dash


# --------------------------------------------------------------------------
# Run helpers
# --------------------------------------------------------------------------

def _is_superscript(run):
    if run.font.superscript:
        return True
    rpr = run._element.find(qn("w:rPr"))
    if rpr is not None:
        va = rpr.find(qn("w:vertAlign"))
        if va is not None and va.get(qn("w:val")) == "superscript":
            return True
    return False


def _esc(s):
    return html.escape(s, quote=False)


# --------------------------------------------------------------------------
# Analysis container
# --------------------------------------------------------------------------

class Analysis:
    def __init__(self, filename, source_path):
        self.filename = filename
        self.source_path = source_path
        self.document = None
        self.blocks = []
        self.occurrences = []
        self.references = []
        self.issues = {}
        self._occ_seq = 0
        self._sup_runs = {}     # occ_id -> [run, ...]   (rewrite targets)
        self._ref_paras = {}    # reference number -> paragraph
        self._ref_heading_el = None

    def _add_occurrence(self, numbers, location, context, raw):
        self._occ_seq += 1
        occ = {"id": self._occ_seq, "numbers": numbers, "style": "superscript",
               "location": location, "context": context, "raw": raw}
        self.occurrences.append(occ)
        return occ

    def _cite_html(self, raw, numbers, occ_id):
        nums = ",".join(str(n) for n in numbers)
        return ('<span class="cite" data-occ="%d" data-nums="%s" data-style="superscript">'
                '<sup>%s</sup></span>' % (occ_id, nums, _esc(raw)))


# --------------------------------------------------------------------------
# Detection: superscript in any paragraph (body / table cell / legend)
# --------------------------------------------------------------------------

def _process_paragraph(para, analysis, location):
    runs = list(para.runs)
    plain = "".join(r.text for r in runs)
    parts, i, n = [], 0, len(runs)
    while i < n:
        run = runs[i]
        if _is_superscript(run) and any(c.isdigit() for c in run.text):
            j, group, buf = i, [], ""
            while j < n and _is_superscript(runs[j]):
                group.append(runs[j])
                buf += runs[j].text
                j += 1
            numbers = expand_numbers(buf)
            if numbers:
                occ = analysis._add_occurrence(numbers, location, plain.strip()[:160], buf)
                analysis._sup_runs[occ["id"]] = group
                parts.append(analysis._cite_html(buf, numbers, occ["id"]))
            else:
                parts.append("<sup>%s</sup>" % _esc(buf))
            i = j
        else:
            parts.append(("<sup>%s</sup>" % _esc(run.text)) if _is_superscript(run)
                         else _esc(run.text))
            i += 1
    return "".join(parts)


# --------------------------------------------------------------------------
# Reference-list parsing
# --------------------------------------------------------------------------

_REF_HEADING_RE = re.compile(r"^\s*(references|reference list|bibliography)\s*$", re.I)
_REF_LEADING_NUM_RE = re.compile(r"^\s*(\d+)[.)\]]?\s+(.*\S.*)$", re.S)


def _parse_references(paragraphs, analysis):
    start = None
    for idx, p in enumerate(paragraphs):
        if _REF_HEADING_RE.match(p.text or ""):
            start = idx + 1
            analysis._ref_heading_el = p._p
            break
    refs = []
    if start is None:
        return refs
    for p in paragraphs[start:]:
        text = (p.text or "").strip()
        if not text:
            continue
        if p.style and p.style.name and p.style.name.lower().startswith("heading") \
                and not _REF_LEADING_NUM_RE.match(text):
            break
        m = _REF_LEADING_NUM_RE.match(text)
        if m:
            num = int(m.group(1))
            refs.append({"number": num, "text": m.group(2).strip(), "cited_by": []})
            analysis._ref_paras[num] = p
        elif refs:
            refs[-1]["text"] += " " + text
    return refs


# --------------------------------------------------------------------------
# Footnotes (separate part; detected for display, rewritten via zip surgery)
# --------------------------------------------------------------------------

def _xml_is_super(r):
    rpr = r.find(W + "rPr")
    if rpr is None:
        return False
    va = rpr.find(W + "vertAlign")
    return va is not None and va.get(W + "val") == "superscript"


def _run_text(r):
    return "".join(t.text or "" for t in r.findall(W + "t"))


def _footnote_html(fn):
    """Build display HTML for a footnote and yield (numbers, raw) citations."""
    out, cites = [], []
    for p in fn.findall(W + "p"):
        runs = p.findall(W + "r")
        i, n = 0, len(runs)
        while i < n:
            r = runs[i]
            if _xml_is_super(r) and any(c.isdigit() for c in _run_text(r)):
                j, buf = i, ""
                while j < n and _xml_is_super(runs[j]):
                    buf += _run_text(runs[j])
                    j += 1
                numbers = expand_numbers(buf)
                if numbers:
                    cites.append((numbers, buf))
                    out.append("\x00%d\x01" % (len(cites) - 1))  # placeholder
                else:
                    out.append("<sup>%s</sup>" % _esc(buf))
                i = j
            else:
                out.append(("<sup>%s</sup>" % _esc(_run_text(r))) if _xml_is_super(r)
                           else _esc(_run_text(r)))
                i += 1
        out.append(" ")
    return "".join(out), cites


def _parse_footnotes(path, analysis):
    try:
        with zipfile.ZipFile(path) as z:
            if "word/footnotes.xml" not in z.namelist():
                return
            xml = z.read("word/footnotes.xml")
    except Exception:
        return
    from lxml import etree
    root = etree.fromstring(xml)
    for fn in root.findall(W + "footnote"):
        if fn.get(W + "type") in ("separator", "continuationSeparator"):
            continue
        try:
            if int(fn.get(W + "id", "0")) < 1:
                continue
        except ValueError:
            continue
        raw_html, cites = _footnote_html(fn)
        if not "".join(t.text or "" for t in fn.iter(W + "t")).strip():
            continue
        # replace placeholders with real citation spans (registering occurrences)
        def sub(m):
            numbers, buf = cites[int(m.group(1))]
            occ = analysis._add_occurrence(numbers, "footnote", "", buf)
            return analysis._cite_html(buf, numbers, occ["id"])
        html_body = re.sub("\x00(\\d+)\x01", sub, raw_html)
        analysis.blocks.append({"type": "footnote", "html": html_body})


# --------------------------------------------------------------------------
# analyze()
# --------------------------------------------------------------------------

def analyze(path, filename=None):
    document = docx.Document(path)
    analysis = Analysis(filename or os.path.basename(path), path)
    analysis.document = document

    analysis.references = _parse_references(list(document.paragraphs), analysis)
    ref_numbers = {r["number"] for r in analysis.references}

    body = document.element.body
    para_map = {p._element: p for p in document.paragraphs}
    table_map = {t._element: t for t in document.tables}
    in_refs = False

    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            para = para_map.get(child)
            if para is None:
                continue
            text = (para.text or "").strip()
            if _REF_HEADING_RE.match(text):
                in_refs = True
            if in_refs or not text:
                continue
            analysis.blocks.append({"type": "para",
                                    "html": _process_paragraph(para, analysis, "body")})
        elif child.tag == qn("w:tbl"):
            table = table_map.get(child)
            if table is None or in_refs:
                continue
            rows_html = []
            for row in table.rows:
                cells = []
                for cell in row.cells:
                    cells.append(" ".join(_process_paragraph(p, analysis, "table")
                                          for p in cell.paragraphs))
                rows_html.append(cells)
            analysis.blocks.append({"type": "table", "rows": rows_html})

    _parse_footnotes(path, analysis)

    for occ in analysis.occurrences:
        for num in occ["numbers"]:
            for ref in analysis.references:
                if ref["number"] == num:
                    ref["cited_by"].append(occ["id"])

    all_cited = set()
    for occ in analysis.occurrences:
        all_cited.update(occ["numbers"])
    analysis.issues = {
        "orphan_citations": sorted(n for n in all_cited if n not in ref_numbers),
        "uncited_references": [r["number"] for r in analysis.references if not r["cited_by"]],
    }
    return analysis


# --------------------------------------------------------------------------
# Renumbering + export
# --------------------------------------------------------------------------

def compute_auto_order(analysis):
    seen, order = set(), []
    for occ in analysis.occurrences:
        for num in occ["numbers"]:
            if num not in seen and num in analysis._ref_paras:
                seen.add(num)
                order.append(num)
    for r in analysis.references:
        if r["number"] not in seen:
            order.append(r["number"])
    return order


def _set_ref_number(paragraph, new_num):
    m = re.match(r"\s*(\d+)", paragraph.text)
    if not m:
        return
    start, end = m.start(1), m.end(1)
    bounds, pos = [], 0
    for r in paragraph.runs:
        bounds.append((r, pos, pos + len(r.text)))
        pos += len(r.text)
    first = True
    for r, rs, re_ in bounds:
        if re_ <= start or rs >= end:
            continue
        ls, le = max(start, rs) - rs, min(end, re_) - rs
        r.text = r.text[:ls] + (str(new_num) if first else "") + r.text[le:]
        first = False


def apply_renumber(source_path, order, out_path):
    """Renumber a fresh copy of the document to the given final reference order.

    `order`: list of {"orig": <int>} (existing reference kept) or
    {"orig": None, "text": "..."} (newly inserted). Missing originals = deleted.
    """
    analysis = analyze(source_path)
    document = analysis.document

    original_numbers = [r["number"] for r in analysis.references]
    present = [it["orig"] for it in order if it.get("orig") is not None]
    deleted = sorted(set(original_numbers) - set(present))

    mapping = {}
    for pos, it in enumerate(order, start=1):
        if it.get("orig") is not None:
            mapping[it["orig"]] = pos

    warnings = []

    # 1. Rewrite superscript citations (body + tables + legends).
    for occ in analysis.occurrences:
        if occ["id"] not in analysis._sup_runs:
            continue
        dropped = [n for n in occ["numbers"] if n in deleted]
        kept = [mapping[n] for n in occ["numbers"] if n in mapping]
        comma, dash = _seps_from_raw(occ["raw"])
        runs = analysis._sup_runs[occ["id"]]
        runs[0].text = format_citation(kept, comma, dash)
        for r in runs[1:]:
            r.text = ""
        if dropped:
            warnings.append("A citation referenced deleted reference(s) %s -- "
                            "removed from the text." % ", ".join(map(str, dropped)))

    # 2. Rewrite + reorder + insert/delete the reference list.
    anchor = analysis._ref_heading_el
    new_els = []
    for pos, it in enumerate(order, start=1):
        if it.get("orig") is not None:
            para = analysis._ref_paras[it["orig"]]
            _set_ref_number(para, pos)
            new_els.append(para._p)
        else:
            para = document.add_paragraph()
            run = para.add_run("%d. " % pos)
            run.bold = True
            para.add_run(it.get("text", "").strip())
            new_els.append(para._p)

    for r in analysis.references:
        el = analysis._ref_paras[r["number"]]._p
        if el.getparent() is not None:
            el.getparent().remove(el)
    for el in new_els:
        if el.getparent() is not None:
            el.getparent().remove(el)
    cursor = anchor
    for el in new_els:
        cursor.addnext(el)
        cursor = el

    document.save(out_path)

    # 3. Footnotes live outside the main part -> patch the saved zip.
    _rewrite_footnotes_zip(out_path, mapping, deleted, warnings)

    change_log = sorted(((o, n) for o, n in mapping.items() if o != n), key=lambda x: x[0])
    inserted = [pos for pos, it in enumerate(order, start=1) if it.get("orig") is None]
    return {"mapping": mapping, "change_log": change_log, "inserted": inserted,
            "deleted": deleted, "warnings": warnings, "out_path": out_path}


def _set_xml_run_text(r, new_text):
    ts = r.findall(W + "t")
    if not ts:
        return
    ts[0].text = new_text
    if new_text:
        ts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for extra in ts[1:]:
        extra.text = ""


def _rewrite_footnotes_zip(path, mapping, deleted, warnings):
    """Renumber superscript citations inside word/footnotes.xml."""
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            if "word/footnotes.xml" not in names:
                return
            data = {n: z.read(n) for n in names}
    except Exception:
        return

    from lxml import etree
    root = etree.fromstring(data["word/footnotes.xml"])
    for fn in root.findall(W + "footnote"):
        if fn.get(W + "type") in ("separator", "continuationSeparator"):
            continue
        for p in fn.findall(W + "p"):
            runs = p.findall(W + "r")
            i, n = 0, len(runs)
            while i < n:
                r = runs[i]
                if _xml_is_super(r) and any(c.isdigit() for c in _run_text(r)):
                    j, buf, group = i, "", []
                    while j < n and _xml_is_super(runs[j]):
                        group.append(runs[j])
                        buf += _run_text(runs[j])
                        j += 1
                    numbers = expand_numbers(buf)
                    if numbers and any(x in mapping or x in deleted for x in numbers):
                        comma, dash = _seps_from_raw(buf)
                        kept = [mapping[x] for x in numbers if x in mapping]
                        _set_xml_run_text(group[0], format_citation(kept, comma, dash))
                        for g in group[1:]:
                            _set_xml_run_text(g, "")
                        if any(x in deleted for x in numbers):
                            warnings.append("A footnote citation referenced a deleted "
                                            "reference -- removed.")
                    i = j
                else:
                    i += 1

    data["word/footnotes.xml"] = etree.tostring(root, xml_declaration=True,
                                                 encoding="UTF-8", standalone=True)
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for nm, b in data.items():
            zout.writestr(nm, b)
    shutil.move(tmp, path)
