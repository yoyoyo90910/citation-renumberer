"""
Core engine for the Citation Renumberer.

  * analyze()          -- read a .docx, detect citations, parse the reference
                          list, link the two (read-only).
  * apply_renumber()   -- given a desired final reference order, renumber every
                          in-text citation AND the reference list, then save a
                          fresh .docx. The original is never modified.

Citations are NUMBERED references shown in one of three styles; the tool
auto-detects which the document uses (and the caller can force one):
  * superscript ... raised numbers          e.g.  finding.^12
  * bracket ....... square brackets         e.g.  finding [12]
  * paren ......... round brackets          e.g.  finding (12)

They can appear anywhere -- body text, tables, figure legends, footnotes.
Ranges (14-16) and lists (12,13) are handled. Front matter (title / authors /
affiliations / metadata) is skipped so affiliation markers aren't mistaken for
citations. Everything is deterministic pattern matching -- no model guessing.

Author-date styles (Harvard/APA "Smith, 2023") are not numbered and therefore
out of scope for a renumbering tool.
"""

import copy
import html
import os
import re
import shutil
import zipfile

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

STYLES = ("superscript", "bracket", "paren")

# --------------------------------------------------------------------------
# Number parsing / formatting
# --------------------------------------------------------------------------

_DASHES = r"\-‑‒–—−"
_RANGE_RE = re.compile(r"^(\d+)\s*[" + _DASHES + r"]\s*(\d+)$")
_GROUP = (r"(\d+(?:\s*[" + _DASHES + r"]\s*\d+)?"
          r"(?:\s*[,;]\s*\d+(?:\s*[" + _DASHES + r"]\s*\d+)?)*)")
BRACKET_RE = re.compile(r"\[\s*" + _GROUP + r"\s*\]")
PAREN_RE = re.compile(r"\(\s*" + _GROUP + r"\s*\)")


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


def _wrap(style, inner):
    """Wrap a formatted number string in the style's delimiters (plain text)."""
    if style == "bracket":
        return "[" + inner + "]"
    if style == "paren":
        return "(" + inner + ")"
    return inner


# --------------------------------------------------------------------------
# Run / paragraph helpers
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


def _replace_span(paragraph, start, end, new_text):
    """Replace characters [start, end) of a paragraph's text, keeping runs."""
    bounds, pos = [], 0
    for r in paragraph.runs:
        bounds.append((r, pos, pos + len(r.text)))
        pos += len(r.text)
    first = True
    for r, rs, re_ in bounds:
        if re_ <= start or rs >= end:
            continue
        ls, le = max(start, rs) - rs, min(end, re_) - rs
        if first:
            r.text = r.text[:ls] + new_text + r.text[le:]
            first = False
        else:
            r.text = r.text[:ls] + r.text[le:]


def _in_range(numbers, max_ref):
    return bool(numbers) and (max_ref == 0 or all(1 <= x <= max_ref for x in numbers))


# --------------------------------------------------------------------------
# Analysis container
# --------------------------------------------------------------------------

class Analysis:
    def __init__(self, filename, source_path):
        self.filename = filename
        self.source_path = source_path
        self.document = None
        self.style = None
        self.style_counts = {s: 0 for s in STYLES}
        self.blocks = []
        self.occurrences = []
        self.references = []
        self.issues = {}
        self._occ_seq = 0
        self._max_ref = 0
        self._sup_runs = {}       # occ_id -> [run, ...]        (superscript)
        self._span_targets = {}   # occ_id -> (paragraph, s, e) (bracket/paren)
        self._ref_paras = {}
        self._ref_heading_el = None
        self._auto_exemplar = None

    def _add_occurrence(self, numbers, location, context, raw, style):
        self._occ_seq += 1
        occ = {"id": self._occ_seq, "numbers": numbers, "style": style,
               "location": location, "context": context, "raw": raw}
        self.occurrences.append(occ)
        return occ

    def _cite_html(self, raw, numbers, occ_id, style):
        nums = ",".join(str(n) for n in numbers)
        inner = ("<sup>%s</sup>" % _esc(raw)) if style == "superscript" else _esc(raw)
        return ('<span class="cite" data-occ="%d" data-nums="%s" data-style="%s">'
                "%s</span>" % (occ_id, nums, style, inner))


# --------------------------------------------------------------------------
# Per-paragraph detection
# --------------------------------------------------------------------------

def _proc_superscript(para, analysis, location):
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
                occ = analysis._add_occurrence(numbers, location, plain.strip()[:160],
                                               buf, "superscript")
                analysis._sup_runs[occ["id"]] = group
                parts.append(analysis._cite_html(buf, numbers, occ["id"], "superscript"))
            else:
                parts.append("<sup>%s</sup>" % _esc(buf))
            i = j
        else:
            parts.append(("<sup>%s</sup>" % _esc(run.text)) if _is_superscript(run)
                         else _esc(run.text))
            i += 1
    return "".join(parts)


def _proc_delimited(para, analysis, location, style):
    """Bracket / paren detection over the paragraph text (records span targets)."""
    text = para.text
    regex = BRACKET_RE if style == "bracket" else PAREN_RE
    out, last = [], 0
    for m in regex.finditer(text):
        numbers = expand_numbers(m.group(1))
        out.append(_esc(text[last:m.start()]))
        # brackets are unambiguous; parens must be within the reference range so
        # a year (2024) or sample size (45) is never treated as a citation.
        ok = numbers and (style == "bracket" or _in_range(numbers, analysis._max_ref))
        if ok:
            occ = analysis._add_occurrence(numbers, location, text.strip()[:160],
                                           m.group(0), style)
            analysis._span_targets[occ["id"]] = (para, m.start(), m.end())
            out.append(analysis._cite_html(m.group(0), numbers, occ["id"], style))
        else:
            out.append('<span class="cite-maybe" title="Out of reference range '
                       '-- not treated as a citation">%s</span>' % _esc(m.group(0)))
        last = m.end()
    out.append(_esc(text[last:]))
    return "".join(out)


def _proc(para, analysis, location):
    if analysis.style == "superscript":
        return _proc_superscript(para, analysis, location)
    return _proc_delimited(para, analysis, location, analysis.style)


# --------------------------------------------------------------------------
# Auto-detection
# --------------------------------------------------------------------------

def _count_superscript(para):
    runs = list(para.runs)
    i, n, cnt = 0, len(runs), 0
    while i < n:
        if _is_superscript(runs[i]) and any(c.isdigit() for c in runs[i].text):
            j, buf = i, ""
            while j < n and _is_superscript(runs[j]):
                buf += runs[j].text
                j += 1
            if expand_numbers(buf):
                cnt += 1
            i = j
        else:
            i += 1
    return cnt


def _unit_paragraphs(units):
    for kind, obj in units:
        if kind == "para":
            yield obj
        else:  # table
            for row in obj.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p


def _autodetect(units, max_ref):
    counts = {s: 0 for s in STYLES}
    for para in _unit_paragraphs(units):
        counts["superscript"] += _count_superscript(para)
        text = para.text
        counts["bracket"] += sum(1 for m in BRACKET_RE.finditer(text)
                                 if expand_numbers(m.group(1)))
        counts["paren"] += sum(1 for m in PAREN_RE.finditer(text)
                               if _in_range(expand_numbers(m.group(1)), max_ref))
    best = max(STYLES, key=lambda s: counts[s])
    return (best if counts[best] > 0 else None), counts


# --------------------------------------------------------------------------
# Reference-list parsing
# --------------------------------------------------------------------------

_REF_HEADING_RE = re.compile(
    r"^\s*(?:\d+\.?\s*)?(references|reference list|bibliography|works cited)\s*:?\s*$", re.I)
_REF_LEADING_NUM_RE = re.compile(r"^\s*(\d+)[.)\]]\s+(.*\S.*)$", re.S)


def _is_list_item(para):
    ppr = para._p.find(qn("w:pPr"))
    return ppr is not None and ppr.find(qn("w:numPr")) is not None


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
    auto_seq = 0
    for p in paragraphs[start:]:
        text = (p.text or "").strip()
        is_list = _is_list_item(p)
        m = _REF_LEADING_NUM_RE.match(text)
        if is_list:
            auto_seq += 1
            body = re.sub(r"^\s*\d+[.)\]]\s+", "", text)
            refs.append({"number": auto_seq, "text": body, "cited_by": [], "auto": True})
            analysis._ref_paras[auto_seq] = p
            if analysis._auto_exemplar is None:
                analysis._auto_exemplar = p
        elif m:
            num = int(m.group(1))
            auto_seq = num
            refs.append({"number": num, "text": m.group(2).strip(),
                         "cited_by": [], "auto": False})
            analysis._ref_paras[num] = p
        elif not text:
            continue
        elif refs and not (p.style and p.style.name
                           and p.style.name.lower().startswith("heading")):
            refs[-1]["text"] += " " + text
        else:
            break
    analysis._max_ref = max((r["number"] for r in refs), default=0)
    return refs


# --------------------------------------------------------------------------
# Body-start detection (skip front matter / author affiliations)
# --------------------------------------------------------------------------

_BODY_START_RE = re.compile(
    r"^(?:\d+\.?\s*)?(abstract|summary|meeting summary|synopsis|introduction|"
    r"background|main text|main article|main body)\b", re.I)


def _find_body_start(paragraphs):
    for p in paragraphs:
        t = (p.text or "").strip()
        if t and len(t) <= 40 and _BODY_START_RE.match(t):
            return p._p
    return None


# --------------------------------------------------------------------------
# Footnotes
# --------------------------------------------------------------------------

def _xml_is_super(r):
    rpr = r.find(W + "rPr")
    if rpr is None:
        return False
    va = rpr.find(W + "vertAlign")
    return va is not None and va.get(W + "val") == "superscript"


def _run_text(r):
    return "".join(t.text or "" for t in r.findall(W + "t"))


def _read_footnotes(path):
    """Return the real footnote elements (skipping separators), or []."""
    try:
        with zipfile.ZipFile(path) as z:
            if "word/footnotes.xml" not in z.namelist():
                return []
            xml = z.read("word/footnotes.xml")
    except Exception:
        return []
    from lxml import etree
    root = etree.fromstring(xml)
    out = []
    for fn in root.findall(W + "footnote"):
        if fn.get(W + "type") in ("separator", "continuationSeparator"):
            continue
        try:
            if int(fn.get(W + "id", "0")) < 1:
                continue
        except ValueError:
            continue
        out.append(fn)
    return out


def _footnote_block(fn, analysis):
    style = analysis.style
    text = "".join(t.text or "" for t in fn.iter(W + "t")).strip()
    if not text:
        return None
    if style == "superscript":
        out = []
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
                        occ = analysis._add_occurrence(numbers, "footnote", text[:160],
                                                       buf, "superscript")
                        out.append(analysis._cite_html(buf, numbers, occ["id"], "superscript"))
                    else:
                        out.append("<sup>%s</sup>" % _esc(buf))
                    i = j
                else:
                    out.append(("<sup>%s</sup>" % _esc(_run_text(r))) if _xml_is_super(r)
                               else _esc(_run_text(r)))
                    i += 1
            out.append(" ")
        return {"type": "footnote", "html": "".join(out)}
    # bracket / paren -- regex over the footnote's plain text
    regex = BRACKET_RE if style == "bracket" else PAREN_RE
    out, last = [], 0
    for m in regex.finditer(text):
        numbers = expand_numbers(m.group(1))
        out.append(_esc(text[last:m.start()]))
        if numbers and (style == "bracket" or _in_range(numbers, analysis._max_ref)):
            occ = analysis._add_occurrence(numbers, "footnote", text[:160],
                                           m.group(0), style)
            out.append(analysis._cite_html(m.group(0), numbers, occ["id"], style))
        else:
            out.append(_esc(m.group(0)))
        last = m.end()
    out.append(_esc(text[last:]))
    return {"type": "footnote", "html": "".join(out)}


# --------------------------------------------------------------------------
# analyze()
# --------------------------------------------------------------------------

def analyze(path, filename=None, style=None):
    document = docx.Document(path)
    analysis = Analysis(filename or os.path.basename(path), path)
    analysis.document = document

    paragraphs = list(document.paragraphs)
    analysis.references = _parse_references(paragraphs, analysis)
    ref_numbers = {r["number"] for r in analysis.references}

    body = document.element.body
    para_map = {p._element: p for p in document.paragraphs}
    table_map = {t._element: t for t in document.tables}
    body_start_el = _find_body_start(paragraphs)
    started = body_start_el is None
    in_refs = False

    units = []
    for child in body.iterchildren():
        if not started:
            if child is body_start_el:
                started = True
            continue
        if child.tag == qn("w:p"):
            para = para_map.get(child)
            if para is None:
                continue
            text = (para.text or "").strip()
            if _REF_HEADING_RE.match(text):
                in_refs = True
            if in_refs or not text:
                continue
            units.append(("para", para))
        elif child.tag == qn("w:tbl"):
            table = table_map.get(child)
            if table is None or in_refs:
                continue
            units.append(("table", table))

    footnotes = _read_footnotes(path)

    # Decide the citation style.
    detected, counts = _autodetect(units, analysis._max_ref)
    analysis.style_counts = counts
    analysis.style = style or detected

    if analysis.style:
        for kind, obj in units:
            if kind == "para":
                analysis.blocks.append({"type": "para",
                                        "html": _proc(obj, analysis, "body")})
            else:
                rows = []
                for row in obj.rows:
                    rows.append([" ".join(_proc(p, analysis, "table")
                                          for p in cell.paragraphs)
                                 for cell in row.cells])
                analysis.blocks.append({"type": "table", "rows": rows})
        for fn in footnotes:
            block = _footnote_block(fn, analysis)
            if block:
                analysis.blocks.append(block)

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


def _clone_auto_ref(exemplar, text):
    new_p = copy.deepcopy(exemplar._p)
    for child in list(new_p):
        if child.tag in (qn("w:r"), qn("w:hyperlink")):
            new_p.remove(child)
    run = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    run.append(t)
    new_p.append(run)
    return new_p


def _set_ref_number(paragraph, new_num):
    m = re.match(r"\s*(\d+)", paragraph.text)
    if m:
        _replace_span(paragraph, m.start(1), m.end(1), str(new_num))


def apply_renumber(source_path, order, out_path, style=None):
    analysis = analyze(source_path, style=style)
    style = analysis.style
    document = analysis.document

    original_numbers = [r["number"] for r in analysis.references]
    present = [it["orig"] for it in order if it.get("orig") is not None]
    deleted = sorted(set(original_numbers) - set(present))

    mapping = {}
    for pos, it in enumerate(order, start=1):
        if it.get("orig") is not None:
            mapping[it["orig"]] = pos

    warnings = []

    # 1. Superscript citations -> rewrite the runs.
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
            warnings.append("A citation referenced deleted reference(s) %s -- removed."
                            % ", ".join(map(str, dropped)))

    # 2. Bracket / paren citations -> span replace, right-to-left per paragraph.
    by_para = {}
    for occ in analysis.occurrences:
        if occ["id"] not in analysis._span_targets:
            continue
        para, s, e = analysis._span_targets[occ["id"]]
        by_para.setdefault(id(para), (para, []))[1].append((s, e, occ))
    for para, spans in by_para.values():
        for s, e, occ in sorted(spans, key=lambda x: x[0], reverse=True):
            dropped = [n for n in occ["numbers"] if n in deleted]
            kept = [mapping[n] for n in occ["numbers"] if n in mapping]
            comma, dash = _seps_from_raw(occ["raw"])
            new_raw = _wrap(style, format_citation(kept, comma, dash)) if kept else ""
            _replace_span(para, s, e, new_raw)
            if dropped:
                warnings.append("A citation referenced deleted reference(s) %s -- removed."
                                % ", ".join(map(str, dropped)))

    # 3. Reference list -- reorder / renumber / insert / delete.
    auto_map = {r["number"]: r.get("auto") for r in analysis.references}
    anchor = analysis._ref_heading_el
    new_els = []
    for pos, it in enumerate(order, start=1):
        if it.get("orig") is not None:
            para = analysis._ref_paras[it["orig"]]
            if not auto_map.get(it["orig"]):
                _set_ref_number(para, pos)
            new_els.append(para._p)
        elif analysis._auto_exemplar is not None:
            new_els.append(_clone_auto_ref(analysis._auto_exemplar, it.get("text", "").strip()))
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

    _rewrite_footnotes_zip(out_path, mapping, deleted, warnings, style)

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


def _rewrite_footnotes_zip(path, mapping, deleted, warnings, style):
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            if "word/footnotes.xml" not in names:
                return
            data = {n: z.read(n) for n in names}
    except Exception:
        return

    from lxml import etree

    if style == "superscript":
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
                        i = j
                    else:
                        i += 1
        data["word/footnotes.xml"] = etree.tostring(root, xml_declaration=True,
                                                     encoding="UTF-8", standalone=True)
    else:
        xml = data["word/footnotes.xml"].decode("utf-8")
        regex = BRACKET_RE if style == "bracket" else PAREN_RE

        def repl(m):
            numbers = expand_numbers(m.group(1))
            if numbers and all(n in mapping for n in numbers):
                comma, dash = _seps_from_raw(m.group(0))
                return _wrap(style, format_citation([mapping[n] for n in numbers], comma, dash))
            return m.group(0)

        data["word/footnotes.xml"] = regex.sub(repl, xml).encode("utf-8")

    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for nm, b in data.items():
            zout.writestr(nm, b)
    shutil.move(tmp, path)
