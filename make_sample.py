"""Generate a realistic sample article for testing the Citation Renumberer.

Produces samples/sample-article.docx containing:
  * body paragraphs with superscript citations (including a range and a list)
  * a data table with parenthetical (n) citations
  * a figure legend with a parenthetical citation
  * a genuine Word footnote that itself carries a citation

Run:  python make_sample.py
"""

import os
import re
import shutil
import zipfile

import docx
from docx.shared import Pt

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(HERE, "samples")
OUT = os.path.join(OUT_DIR, "sample-article.docx")

REFERENCES = [
    "Packer M, et al. The effect of carvedilol on morbidity and mortality in chronic heart failure. N Engl J Med. 1996;334(21):1349-55.",
    "CIBIS-II Investigators. The Cardiac Insufficiency Bisoprolol Study II. Lancet. 1999;353(9146):9-13.",
    "MERIT-HF Study Group. Effect of metoprolol CR/XL in chronic heart failure. Lancet. 1999;353(9169):2001-7.",
    "Flather MD, et al. Randomised trial to determine the effect of nebivolol. Eur Heart J. 2005;26(3):215-25.",
    "Yancy CW, et al. 2013 ACCF/AHA guideline for the management of heart failure. Circulation. 2013;128(16):e240-327.",
    "Ponikowski P, et al. 2016 ESC guidelines for heart failure. Eur Heart J. 2016;37(27):2129-200.",
    "McMurray JJV, et al. Angiotensin-neprilysin inhibition versus enalapril. N Engl J Med. 2014;371(11):993-1004.",
    "Solomon SD, et al. Angiotensin-neprilysin inhibition in HFpEF. N Engl J Med. 2019;381(17):1609-20.",
    "Zannad F, et al. Eplerenone in patients with systolic heart failure. N Engl J Med. 2011;364(1):11-21.",
    "Pitt B, et al. The effect of spironolactone on morbidity and mortality. N Engl J Med. 1999;341(10):709-17.",
    "Swedberg K, et al. Ivabradine and outcomes in chronic heart failure. Lancet. 2010;376(9744):875-85.",
    "Digitalis Investigation Group. The effect of digoxin on mortality. N Engl J Med. 1997;336(8):525-33.",
    "McMurray JJV, et al. Dapagliflozin in patients with HFrEF. N Engl J Med. 2019;381(21):1995-2008.",
    "Anker SD, et al. Empagliflozin in heart failure with preserved EF. N Engl J Med. 2021;385(16):1451-61.",
    "Packer M, et al. Cardiovascular outcomes with empagliflozin in HFrEF. N Engl J Med. 2020;383(15):1413-24.",
    "Bhatt DL, et al. Sotagliflozin in patients with diabetes and heart failure. N Engl J Med. 2021;384(2):117-28.",
    "Cleland JGF, et al. Cardiac resynchronization therapy in heart failure. N Engl J Med. 2005;352(15):1539-49.",
    "Moss AJ, et al. Cardiac-resynchronization therapy for prevention of events. N Engl J Med. 2009;361(14):1329-38.",
    "Kadish A, et al. Prophylactic defibrillator implantation in cardiomyopathy. N Engl J Med. 2004;350(21):2151-8.",
    "Bardy GH, et al. Amiodarone or an implantable defibrillator. N Engl J Med. 2005;352(3):225-37.",
]


def add_super(paragraph, text):
    """Append a superscript run (a citation) to a paragraph."""
    run = paragraph.add_run(text)
    run.font.superscript = True
    return run


def build():
    os.makedirs(OUT_DIR, exist_ok=True)
    doc = docx.Document()

    doc.add_heading("Contemporary pharmacological management of chronic heart failure", level=0)
    doc.add_paragraph("A narrative review", style="Subtitle")

    doc.add_heading("Introduction", level=1)
    p = doc.add_paragraph("Beta-blockers remain a cornerstone of therapy for heart "
                          "failure with reduced ejection fraction")
    add_super(p, "1")
    p.add_run(" and have consistently reduced mortality across landmark trials")
    add_super(p, "2,3")
    p.add_run(". Guideline bodies continue to recommend early initiation")
    add_super(p, "4-6")
    p.add_run(".")

    p = doc.add_paragraph("Renin-angiotensin blockade has evolved from ACE inhibition "
                          "toward angiotensin-neprilysin inhibition")
    add_super(p, "7,8")
    p.add_run(", while mineralocorticoid antagonists add further benefit in selected "
              "patients")
    add_super(p, "9,10")
    p.add_run(".")

    doc.add_heading("Emerging therapies", level=1)
    p = doc.add_paragraph("Rate control with ivabradine")
    add_super(p, "11")
    p.add_run(" and the historical role of digoxin")
    add_super(p, "12")
    p.add_run(" illustrate the breadth of options. Most strikingly, SGLT2 inhibitors")
    add_super(p, "13")
    p.add_run(" have extended benefit to preserved ejection fraction")
    add_super(p, "14-16")
    p.add_run(".")

    p = doc.add_paragraph("Device therapy, including resynchronisation")
    add_super(p, "17,18")
    p.add_run(" and defibrillator implantation")
    add_super(p, "19,20")
    p.add_run(", remains important in appropriately selected patients.")

    # ---- a data table with SUPERSCRIPT citations inside cells ----
    doc.add_heading("Summary of pivotal trials", level=1)
    table = doc.add_table(rows=4, cols=3)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    hdr[0].text = "Agent class"
    hdr[1].text = "Landmark trial"
    hdr[2].text = "Mortality effect"

    def cell_with_super(cell, text, cite):
        cell.text = text
        add_super(cell.paragraphs[0], cite)

    body_rows = [
        ("Beta-blocker", "CIBIS-II", "2", "Reduced", "3"),
        ("ARNI", "PARADIGM-HF", "7", "Reduced", "8"),
        ("SGLT2 inhibitor", "DAPA-HF", "13", "Reduced", "14-16"),
    ]
    for i, (agent, trial, tc, effect, ec) in enumerate(body_rows, start=1):
        table.rows[i].cells[0].text = agent
        cell_with_super(table.rows[i].cells[1], trial, tc)
        cell_with_super(table.rows[i].cells[2], effect, ec)

    # ---- a figure legend with a SUPERSCRIPT citation ----
    leg = doc.add_paragraph(
        "Figure 1. Kaplan-Meier estimates of survival by treatment group, "
        "adapted from the pivotal SGLT2 inhibitor trials.")
    add_super(leg, "13,14")

    # ---- references ----
    doc.add_heading("References", level=1)
    for i, ref in enumerate(REFERENCES, start=1):
        rp = doc.add_paragraph()
        rp.add_run("%d. " % i).bold = True
        rp.add_run(ref)
        rp.paragraph_format.space_after = Pt(4)

    doc.save(OUT)
    _inject_footnote(OUT)
    print("Wrote", OUT)
    print("References:", len(REFERENCES))


# --------------------------------------------------------------------------
# Footnote injection (python-docx cannot create footnotes directly)
# --------------------------------------------------------------------------

FOOTNOTES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:id="2"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:t xml:space="preserve">Estimates adjusted for age, sex and baseline ejection fraction.</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>14</w:t></w:r></w:p></w:footnote>
</w:footnotes>"""


def _inject_footnote(path):
    """Add a real footnote (id=2) plus a reference marker in the last table cell."""
    tmp = path + ".tmp"
    with zipfile.ZipFile(path, "r") as zin:
        names = zin.namelist()
        data = {n: zin.read(n) for n in names}

    # 1. content-types override
    ct = data["[Content_Types].xml"].decode("utf-8")
    if "footnotes+xml" not in ct:
        ct = ct.replace(
            "</Types>",
            '<Override PartName="/word/footnotes.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.'
            'wordprocessingml.footnotes+xml"/></Types>')
        data["[Content_Types].xml"] = ct.encode("utf-8")

    # 2. relationship from the document to the footnotes part
    rels_name = "word/_rels/document.xml.rels"
    rels = data[rels_name].decode("utf-8")
    if "footnotes.xml" not in rels:
        rels = rels.replace(
            "</Relationships>",
            '<Relationship Id="rIdFootnotes" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/footnotes" Target="footnotes.xml"/></Relationships>')
        data[rels_name] = rels.encode("utf-8")

    # 3. the footnotes part itself
    data["word/footnotes.xml"] = FOOTNOTES_XML.encode("utf-8")

    # 4. drop a footnote reference marker into the document body so Word shows it
    docxml = data["word/document.xml"].decode("utf-8")
    marker = ('<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
              '<w:footnoteReference w:id="2"/></w:r>')
    # place it just before the first sentence-ending period we can find in the body
    docxml = re.sub(r"(</w:p>)", marker + r"\1", docxml, count=1)
    data["word/document.xml"] = docxml.encode("utf-8")

    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, b in data.items():
            zout.writestr(n, b)
    shutil.move(tmp, path)


if __name__ == "__main__":
    build()
