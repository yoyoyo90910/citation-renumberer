"""
Author-date -> numbered citation converter (engine).

Detects author-date in-text citations, proposes a match to each reference, and
-- only once the editor has approved the mapping -- converts them to numbered
superscripts and renumbers the reference list into citation order.

The review step is deliberate: converting "(Smith, 2020)" to a bare number
destroys the author/year cue, so a wrong match would be invisible afterwards.
Nothing is converted without an explicit, checkable mapping.

Matching key = (first-author surname, year). A reference matches a citation
when the surnames agree and the citation's year appears in the reference.
"""

import copy
import html
import re

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def _esc(s):
    return html.escape(s or "", quote=True)

_YEAR = r"(?:19|20)\d{2}"
CITE_RE = re.compile(r"\(([^()]*\b" + _YEAR + r"[a-z]?[^()]*)\)")
PIECE_RE = re.compile(r"^\s*(?P<auth>.*?)[,\s]*\b(?P<year>" + _YEAR + r")(?P<suf>[a-z]?)\b"
                      r"(?:\s*,?\s*p{1,2}\.?\s*\d+)?\s*$")
NARR_AUTHOR_RE = re.compile(
    r"([A-Z][A-Za-z'’\-]+(?:\s+et\s+al\.?)?"
    r"(?:\s*(?:&|and)\s*[A-Z][A-Za-z'’\-]+)?)\s*$")

_REF_HEADING_RE = re.compile(
    r"^\s*(references|reference list|bibliography|works cited)\s*:?\s*$", re.I)
_BODY_START_RE = re.compile(
    r"^(?:\d+\.?\s*)?(abstract|summary|meeting summary|synopsis|introduction|"
    r"background|main text|main article|main body)\b", re.I)


def _surname_key(text):
    """First real name-word, lowercased -- the matching key. Handles hyphens,
    apostrophes and a leading number ('1. Smith ...')."""
    t = re.sub(r"^\s*\[?\d+\]?[.)]?\s*", "", text or "")
    m = re.search(r"[A-Za-z][A-Za-z'’\-]+", t)
    return m.group(0).lower().replace("’", "'") if m else ""


def _years(text):
    return set(re.findall(r"\b(" + _YEAR + r")\b", text or ""))


def token_key(surname, year, suf):
    return "%s|%s|%s" % (surname, year, suf)


def format_citation(numbers, comma=",", dash="–"):
    nums = sorted(set(numbers))
    if not nums:
        return ""
    groups, i = [], 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
            j += 1
        groups.append("%d%s%d" % (nums[i], dash, nums[j]) if j - i >= 2
                      else ",".join(str(n) for n in nums[i:j + 1]))
        i = j + 1
    return comma.join(groups)


# --------------------------------------------------------------------------
# Parsing
# --------------------------------------------------------------------------

_REF_LEAD_NUM = re.compile(r"^\s*\d+[.)\]]\s+")


def parse_references(paragraphs):
    start = None
    for i, p in enumerate(paragraphs):
        if _REF_HEADING_RE.match(p.text or ""):
            start = i + 1
            break
    refs = []
    if start is None:
        return refs
    # If the list is numbered ("1. ...", "2) ..."), a new entry begins only on a
    # leading number and unnumbered lines are continuations. Otherwise treat each
    # non-empty paragraph as its own entry.
    numbered = None
    for p in paragraphs[start:]:
        text = (p.text or "").strip()
        if not text:
            continue
        lead = bool(_REF_LEAD_NUM.match(text))
        if numbered is None:
            numbered = lead
        if refs and numbered and not lead:
            refs[-1]["text"] += " " + text
            refs[-1]["years"] |= _years(text)
            continue
        refs.append({
            "id": len(refs),
            "surname": _surname_key(text),
            "years": _years(text),
            "text": text,
            "paragraph": p,
        })
    return refs


def _find_body_start(paragraphs):
    for p in paragraphs:
        t = (p.text or "").strip()
        if t and len(t) <= 40 and _BODY_START_RE.match(t):
            return p._p
    return None


def _pieces_for(group_text, preceding_text):
    pieces = []
    for chunk in group_text.split(";"):
        pm = PIECE_RE.match(chunk)
        if not pm:
            continue
        surname = _surname_key(pm.group("auth"))
        if not surname:                       # narrative form: author before "(year)"
            am = NARR_AUTHOR_RE.search(preceding_text.rstrip())
            if am:
                surname = _surname_key(am.group(1))
        pieces.append({"surname": surname, "year": pm.group("year"), "suf": pm.group("suf"),
                       "key": token_key(surname, pm.group("year"), pm.group("suf")),
                       "raw": chunk.strip()})
    return pieces


def scan(document):
    """Return (occurrences, blocks). blocks render the article body with every
    citation wrapped in a span carrying its piece keys -- for the live preview."""
    paragraphs = list(document.paragraphs)
    body_start = _find_body_start(paragraphs)
    started = body_start is None
    occ, blocks = [], []
    for p in paragraphs:
        if not started:
            if p._p is body_start:
                started = True
            continue
        if _REF_HEADING_RE.match((p.text or "").strip()):
            break
        text = p.text or ""
        matches = []
        for m in CITE_RE.finditer(text):
            pieces = _pieces_for(m.group(1), text[:m.start()])
            if pieces:
                o = {"paragraph": p, "start": m.start(), "end": m.end(),
                     "raw": m.group(0), "pieces": pieces}
                occ.append(o)
                matches.append(o)
        if not text.strip():
            continue
        parts, last = [], 0
        for o in matches:
            parts.append(_esc(text[last:o["start"]]))
            keys = ",".join(pc["key"] for pc in o["pieces"])
            parts.append('<span class="cite" data-keys="%s" data-raw="%s">%s</span>'
                         % (_esc(keys), _esc(o["raw"]), _esc(o["raw"])))
            last = o["end"]
        parts.append(_esc(text[last:]))
        blocks.append({"type": "para", "html": "".join(parts)})
    return occ, blocks


def detect_citations(document):
    return scan(document)[0]


# --------------------------------------------------------------------------
# Proposing matches (for the review step)
# --------------------------------------------------------------------------

def _candidates(references, surname, year):
    return [r["id"] for r in references if r["surname"] == surname and year in r["years"]]


def propose(references, occurrences):
    """One row per distinct in-text citation token, with a default match and a
    status the UI can flag: matched | review | unmatched."""
    tokens = {}
    order = []
    for occ in occurrences:
        for pc in occ["pieces"]:
            k = pc["key"]
            if k not in tokens:
                cands = _candidates(references, pc["surname"], pc["year"])
                if len(cands) == 1:
                    default, status = cands[0], "matched"
                elif len(cands) > 1:
                    pos = (ord(pc["suf"]) - ord("a")) if pc["suf"] else 0
                    default = cands[pos] if 0 <= pos < len(cands) else cands[0]
                    status = "review"       # same surname+year -> a/b ambiguity
                else:
                    default, status = None, "unmatched"
                tokens[k] = {"key": k, "display": pc["raw"], "surname": pc["surname"],
                             "year": pc["year"], "suf": pc["suf"], "count": 0,
                             "candidates": cands, "default": default, "status": status}
                order.append(k)
            tokens[k]["count"] += 1
    return [tokens[k] for k in order]


# --------------------------------------------------------------------------
# Applying an approved mapping
# --------------------------------------------------------------------------

def _mk_super_run(text):
    r = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    va = OxmlElement("w:vertAlign")
    va.set(qn("w:val"), "superscript")
    rpr.append(va)
    r.append(rpr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def _set_elem_text(elem, text):
    ts = elem.findall(qn("w:t"))
    if ts:
        ts[0].text = text
        ts[0].set(qn("xml:space"), "preserve")
        for extra in ts[1:]:
            extra.text = ""


def _replace_with_super(paragraph, start, end, number_text):
    bounds, pos = [], 0
    for r in paragraph.runs:
        bounds.append((r, pos, pos + len(r.text)))
        pos += len(r.text)
    cover = [(r, rs, re_) for r, rs, re_ in bounds if not (re_ <= start or rs >= end)]
    if not cover:
        return False
    first_r, frs, _ = cover[0]
    last_r, lrs, _ = cover[-1]
    prefix = first_r.text[:start - frs]
    suffix = last_r.text[end - lrs:]
    sup = _mk_super_run(number_text)
    if last_r is first_r:
        first_r.text = prefix
        first_r._element.addnext(sup)
        if suffix:
            suf = copy.deepcopy(first_r._element)
            _set_elem_text(suf, suffix)
            sup.addnext(suf)
    else:
        first_r.text = prefix
        first_r._element.addnext(sup)
        for r, _, _ in cover[1:-1]:
            r.text = ""
        last_r.text = suffix
    return True


def apply_with_mapping(document, references, occurrences, mapping):
    """mapping: {token_key: reference_id or None}. Only occurrences whose every
    piece is mapped get converted; the rest are left untouched and reported."""
    number_of, order, warnings = {}, [], []

    def assign(ref_id):
        if ref_id not in number_of:
            number_of[ref_id] = len(order) + 1
            order.append(ref_id)
        return number_of[ref_id]

    # which occurrences can be converted (all pieces mapped)?
    for occ in occurrences:
        occ["_conv"] = all(mapping.get(pc["key"]) is not None for pc in occ["pieces"])
        if not occ["_conv"]:
            warnings.append("Left unchanged (unresolved match): %s" % occ["raw"])

    # assign numbers by order of first appearance, over convertible occurrences
    for occ in occurrences:
        if not occ["_conv"]:
            continue
        for pc in occ["pieces"]:
            assign(mapping[pc["key"]])

    # replace in-text, right-to-left within each paragraph
    by_para = {}
    for occ in occurrences:
        if not occ["_conv"]:
            continue
        nums = [number_of[mapping[pc["key"]]] for pc in occ["pieces"]]
        by_para.setdefault(id(occ["paragraph"]), (occ["paragraph"], []))[1].append((occ, nums))
    for para, items in by_para.values():
        for occ, nums in sorted(items, key=lambda x: x[0]["start"], reverse=True):
            _replace_with_super(para, occ["start"], occ["end"], format_citation(nums))

    # uncited references keep their place at the end
    for r in references:
        if r["id"] not in number_of:
            assign(r["id"])
            warnings.append("Reference never cited: %s" % r["text"][:60])

    # reorder + number the reference list
    if references:
        ref_by_id = {r["id"]: r for r in references}
        anchor = references[0]["paragraph"]._p.getprevious()
        for r in references:
            el = r["paragraph"]._p
            if el.getparent() is not None:
                el.getparent().remove(el)
        cursor = anchor
        for num, ref_id in enumerate(order, start=1):
            para = ref_by_id[ref_id]["paragraph"]
            run = OxmlElement("w:r")
            t = OxmlElement("w:t")
            t.set(qn("xml:space"), "preserve")
            t.text = "%d. " % num
            run.append(t)
            ppr = para._p.find(qn("w:pPr"))
            para._p.insert(1 if ppr is not None else 0, run)
            if cursor is not None:
                cursor.addnext(para._p)
                cursor = para._p

    change = [(number_of[r["id"]], r["text"][:70]) for r in references if r["id"] in number_of]
    change.sort()
    converted = sum(1 for o in occurrences if o["_conv"])
    return {"number_of": number_of, "order": order, "warnings": warnings,
            "numbered_refs": change, "converted": converted,
            "left": len(occurrences) - converted}
