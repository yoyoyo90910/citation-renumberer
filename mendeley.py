"""
Mendeley support.

Documents written with the Mendeley Cite add-in store each in-text citation as
a Word content control (<w:sdt>) whose tag holds base64-encoded JSON: the
rendered text (e.g. "(Riccardi et al., 2019)") AND the full reference metadata
(authors, title, journal, year, DOI). The bibliography is often an ungenerated
placeholder, so the reference list looks "empty" to a plain-text reader.

Because Mendeley records exactly which reference each citation points to, we can
convert precisely -- no surname/year guessing. flatten_to_numbered() rewrites a
copy of the document so that:
  * every Mendeley citation becomes a superscript number (Vancouver order), and
  * a numbered reference list is written under the References heading.
The result is an ordinary .docx the rest of the tool reads normally.
"""

import base64
import json
import re
import shutil
import zipfile

from lxml import etree

import citations

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
def qn(tag):
    return "{%s}%s" % (W, tag)

CIT_PREFIX = "MENDELEY_CITATION_v3_"


def is_mendeley(path):
    try:
        with zipfile.ZipFile(path) as z:
            return CIT_PREFIX.encode() in z.read("word/document.xml")
    except Exception:
        return False


def _decode_tag(val):
    b64 = val[len(CIT_PREFIX):]
    b64 += "=" * (-len(b64) % 4)
    return json.loads(base64.b64decode(b64).decode("utf-8", "ignore"))


def _sdt_tag(sdt):
    pr = sdt.find(qn("sdtPr"))
    if pr is None:
        return None
    tag = pr.find(qn("tag"))
    return tag.get(qn("val")) if tag is not None else None


def _initials(given):
    return "".join(w[0] for w in re.split(r"[\s.\-]+", given or "") if w[:1].isalpha()).upper()


def format_reference(d):
    """Build a readable Vancouver-ish reference string from CSL item data."""
    names = []
    for a in (d.get("author") or []):
        fam = (a.get("family") or "").strip()
        nm = (fam + " " + _initials(a.get("given") or "")).strip() if fam else (a.get("literal") or "").strip()
        if nm:
            names.append(nm)
    author = (", ".join(names[:6]) + ", et al") if len(names) > 6 else ", ".join(names)
    try:
        yr = str(d.get("issued", {}).get("date-parts", [[None]])[0][0] or "")
    except Exception:
        yr = ""
    title = (d.get("title") or "").strip()
    jour = (d.get("container-title") or d.get("container-title-short") or "").strip()
    vol, iss, pg, doi = d.get("volume"), d.get("issue"), d.get("page"), d.get("DOI")
    seg = []
    if author:
        seg.append(author.rstrip(".") + ".")
    if title:
        seg.append(title.rstrip(".") + ".")
    if jour:
        seg.append(jour.rstrip(".") + ".")
    tail = ""
    if yr:
        tail += yr
    if vol:
        tail += ";" + str(vol)
    if iss:
        tail += "(" + str(iss) + ")"
    if pg:
        tail += ":" + str(pg)
    if tail:
        seg.append(tail + ".")
    if doi:
        seg.append("doi:" + str(doi))
    return " ".join(seg).strip() or "[reference details unavailable]"


def _super_run(text):
    r = etree.SubElement(etree.Element(qn("wrapper")), qn("r"))
    rpr = etree.SubElement(r, qn("rPr"))
    va = etree.SubElement(rpr, qn("vertAlign"))
    va.set(qn("val"), "superscript")
    t = etree.SubElement(r, qn("t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = text
    return r


def _ref_paragraph(num, text):
    p = etree.Element(qn("p"))
    r = etree.SubElement(p, qn("r"))
    b = etree.SubElement(r, qn("rPr"))
    etree.SubElement(b, qn("b"))
    t = etree.SubElement(r, qn("t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = "%d. " % num
    r2 = etree.SubElement(p, qn("r"))
    t2 = etree.SubElement(r2, qn("t"))
    t2.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t2.text = text
    return p


def _para_text(p):
    return "".join(t.text or "" for t in p.iter(qn("t"))).strip()


def flatten_to_numbered(src, out):
    with zipfile.ZipFile(src) as z:
        data = {n: z.read(n) for n in z.namelist()}
    root = etree.fromstring(data["word/document.xml"])
    body = root.find(qn("body"))

    # 1. Assign numbers by first appearance across citations (document order).
    number_of, order, ref_data = {}, [], {}
    cit_sdts = []
    for sdt in root.iter(qn("sdt")):
        val = _sdt_tag(sdt)
        if not val or not val.startswith(CIT_PREFIX):
            continue
        try:
            c = _decode_tag(val)
        except Exception:
            continue
        nums = []
        for it in c.get("citationItems", []):
            d = it.get("itemData", {})
            rid = d.get("id") or it.get("id")
            if not rid:
                continue
            if rid not in number_of:
                number_of[rid] = len(order) + 1
                order.append(rid)
                ref_data[rid] = d
            nums.append(number_of[rid])
        cit_sdts.append((sdt, nums))

    # 2. Replace each citation control with a superscript number run.
    for sdt, nums in cit_sdts:
        parent = sdt.getparent()
        if parent is None:
            continue
        run = _super_run(citations.format_citation(nums) if nums else "")
        if parent.tag == qn("body"):        # block-level -> wrap the run in a paragraph
            wrap = etree.Element(qn("p"))
            wrap.append(run)
            parent.replace(sdt, wrap)
        else:
            parent.replace(sdt, run)

    # 3. Remove any Mendeley bibliography placeholder control.
    for sdt in list(root.iter(qn("sdt"))):
        val = _sdt_tag(sdt) or ""
        if "MENDELEY_BIBLIOGRAPHY" in val and sdt.getparent() is not None:
            sdt.getparent().remove(sdt)

    # 4. Write the numbered reference list under the References heading.
    ref_ps = [_ref_paragraph(i + 1, format_reference(ref_data[rid])) for i, rid in enumerate(order)]
    anchor = None
    for p in body.findall(qn("p")):
        if re.fullmatch(r"references|reference list|bibliography", _para_text(p), re.I):
            anchor = p
            break
    if anchor is None:                       # no heading -> add one at the end
        sect = body.find(qn("sectPr"))
        heading = etree.Element(qn("p"))
        hr = etree.SubElement(heading, qn("r"))
        hpr = etree.SubElement(hr, qn("rPr"))
        etree.SubElement(hpr, qn("b"))
        ht = etree.SubElement(hr, qn("t"))
        ht.text = "References"
        if sect is not None:
            sect.addprevious(heading)
        else:
            body.append(heading)
        anchor = heading
    for p in reversed(ref_ps):
        anchor.addnext(p)

    data["word/document.xml"] = etree.tostring(root, xml_declaration=True,
                                               encoding="UTF-8", standalone=True)
    tmp = out + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, b in data.items():
            zout.writestr(n, b)
    shutil.move(tmp, out)
    return {"citations": len(cit_sdts), "references": len(order)}
