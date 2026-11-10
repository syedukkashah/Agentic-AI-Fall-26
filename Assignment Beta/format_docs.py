"""Plain house style for every Assignment Beta document.

Black text on white, no shading or highlights, headings underlined.
Sizes: title and level-1 headings 20 pt, level-2/3 headings 18 pt, body paragraphs 16 pt.
Table text keeps its size (the wide comparison tables cannot fit at 16 pt).

    python format_docs.py <file.docx> [...]      # formats the files in place
The Q1 report builder imports normalize() so a rebuild keeps this style.
"""
import re
import sys

from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

BLACK = RGBColor(0, 0, 0)
TITLE_PT, H1_PT, H2_PT, BODY_PT = 20, 20, 18, 16
HEADING_SIZES = {"Title": TITLE_PT, "Heading 1": H1_PT, "Heading 2": H2_PT, "Heading 3": H2_PT}
BORDER_TAGS = {qn(f"w:{t}") for t in ("top", "left", "bottom", "right", "insideH", "insideV", "start", "end")}


def _plain_xml(root):
    """Drop shading, highlights and paragraph rules; force every colour to black."""
    for el in list(root.iter(qn("w:shd"), qn("w:highlight"), qn("w:pBdr"))):
        el.getparent().remove(el)
    for el in root.iter(qn("w:color")):
        el.set(qn("w:val"), "000000")
        for a in ("w:themeColor", "w:themeShade", "w:themeTint"):
            el.attrib.pop(qn(a), None)
    for el in root.iter():
        if el.tag in BORDER_TAGS and el.get(qn("w:val")) not in (None, "nil", "none"):
            el.set(qn("w:color"), "000000")
            el.attrib.pop(qn("w:themeColor"), None)


def _in_table(par):
    node = par._p.getparent()
    while node is not None:
        if node.tag == qn("w:tc"):
            return True
        node = node.getparent()
    return False


def style_heading(par, size):
    for r in par.runs:
        r.font.size = Pt(size)
        r.font.bold = True
        r.font.underline = True
        r.font.color.rgb = BLACK


def normalize(doc):
    _plain_xml(doc.styles.element)
    _plain_xml(doc.element.body)
    for sec in doc.sections:
        for part in (sec.header, sec.footer, sec.first_page_header, sec.first_page_footer):
            if not part.is_linked_to_previous:
                _plain_xml(part._element)

    for s in doc.styles:
        if not hasattr(s, "font"):
            continue
        s.font.color.rgb = BLACK
        if s.name in HEADING_SIZES:
            s.font.size = Pt(HEADING_SIZES[s.name])
            s.font.bold = True
            s.font.underline = True
            s.font.italic = False
    doc.styles["Normal"].font.size = Pt(BODY_PT)

    for par in doc.paragraphs:  # body paragraphs only; table cells keep their sizes
        if _in_table(par):
            continue
        size = HEADING_SIZES.get(par.style.name)
        if size:
            style_heading(par, size)
        else:
            for r in par.runs:
                if r.font.size is not None:
                    r.font.size = Pt(BODY_PT)
    return doc


def set_text(par, text, size=None, bold=None, underline=None, italic=None):
    """Replace a paragraph's text (and any line breaks) with one run."""
    for r in list(par.runs):
        r._r.getparent().remove(r._r)
    run = par.add_run(text)
    if size: run.font.size = Pt(size)
    if bold is not None: run.font.bold = bold
    if underline is not None: run.font.underline = underline
    if italic is not None: run.font.italic = italic
    run.font.color.rgb = BLACK
    return run


def main(paths):
    for path in paths:
        doc = normalize(Document(path))
        doc.save(path)
        print(f"formatted {path}")


if __name__ == "__main__":
    main(sys.argv[1:])
