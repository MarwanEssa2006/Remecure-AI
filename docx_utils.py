"""python-docx helpers shared by the COC and FPS generators. The official templates are filled, never rebuilt."""
import copy
import os
import re
import shlex
import shutil
import subprocess
import tempfile

from lxml import etree
from docx.oxml import OxmlElement
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.table import _Row
from docx.text.paragraph import Paragraph

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}


def xp(el, query):
    return etree.ElementBase.xpath(el, query, namespaces=NS)


def set_text(paragraph, text):
    """Write text in the first run (keeps its formatting) and blank the other runs. '\\n' becomes a line break."""
    if not paragraph.runs:
        paragraph.add_run("")
    paragraph.runs[0].text = text
    for r in paragraph.runs[1:]:
        if not xp(r._r, ".//w:drawing | .//mc:AlternateContent"):
            r.text = ""


def par_text(paragraph):
    return "".join(r.text for r in paragraph.runs)


def replace_in_par(paragraph, old, new):
    full = par_text(paragraph)
    if old in full:
        set_text(paragraph, full.replace(old, new))


def fill_placeholders(paragraph, value):
    """Replace every [...] placeholder in the paragraph with value ('' leaves just the label)."""
    full = par_text(paragraph)
    set_text(paragraph, re.sub(r"\s*\[[^\]]*\]", (" " + value) if value else "", full).rstrip())


def remove_paragraph(paragraph):
    paragraph._p.getparent().remove(paragraph._p)


def add_after(paragraph):
    """Duplicate a paragraph right after itself, return the copy."""
    new_p = copy.deepcopy(paragraph._p)
    paragraph._p.addnext(new_p)
    return Paragraph(new_p, paragraph._parent)


def paragraph_copy_without_drawings(paragraph):
    new_p = copy.deepcopy(paragraph._p)
    for r in xp(new_p, ".//w:r[.//w:drawing or .//mc:AlternateContent]"):
        r.getparent().remove(r)
    return Paragraph(new_p, paragraph._parent)


def _set_trpr_flag(tr, tag):
    trpr = tr.find(qn("w:trPr"))
    if trpr is None:
        trpr = OxmlElement("w:trPr")
        tr.insert(1 if tr.find(qn("w:tblPrEx")) is not None else 0, trpr)
    if trpr.find(qn(tag)) is None:
        trpr.append(OxmlElement(tag))


def no_split_rows(table, header_rows=1):
    """Rows never break across pages; the header row repeats on a new page."""
    for i, row in enumerate(table.rows):
        _set_trpr_flag(row._tr, "w:cantSplit")
        if i < header_rows:
            _set_trpr_flag(row._tr, "w:tblHeader")


def clone_rows(table, n, proto_idx=-1, keep=1):
    """Use one template row as prototype: drop rows after `keep`, return n fresh copies of the prototype."""
    proto = copy.deepcopy(table.rows[proto_idx]._tr)
    for row in list(table.rows)[keep:]:
        table._tbl.remove(row._tr)
    new = []
    for _ in range(n):
        tr = copy.deepcopy(proto)
        table._tbl.append(tr)
        new.append(_Row(tr, table))
    return new


# ----------------------------------------------------------------------------
# Signature block: keep it glued to the name, wherever the tables end
# ----------------------------------------------------------------------------


def _anchors_in(run_el):
    return xp(run_el, ".//wp:anchor")


SIG_OFFSET_EMU = -1092200  # signature picture top, relative to the signatory name paragraph (calibrated on the templates)


def fix_signature(doc, name_marker="Andersen"):
    """In the templates the signature picture floats relative to a paragraph that sits BEFORE the last table, so it drifts
    when the tables grow. Anchor it to the signatory name paragraph instead and keep the whole block on one page."""
    name_p = next((p for p in doc.paragraphs if name_marker in p.text.replace("\xa0", " ")), None)
    if name_p is None:
        return
    for run in xp(doc.element.body, ".//w:r[.//wp:anchor]"):
        names = xp(run, ".//wp:docPr/@name")
        anchors = _anchors_in(run)
        if not names or not names[0].startswith("Picture") or not anchors or anchors[0].get("behindDoc") == "1":
            continue
        if run.getparent() is not name_p._p:
            name_p._p.append(run)  # move the whole run (with its drawing)
        for a in anchors:
            pv = a.find(qn("wp:positionV"))
            pv.set("relativeFrom", "paragraph")
            for child in list(pv):
                pv.remove(child)
            off = OxmlElement("wp:posOffset")
            off.text = str(SIG_OFFSET_EMU)
            pv.append(off)
    block = [name_p._p.getprevious(), name_p._p, name_p._p.getnext()]
    for el in block:
        if el is not None and el.tag == qn("w:p"):
            p = Paragraph(el, name_p._parent)
            p.paragraph_format.keep_with_next = True
            p.paragraph_format.keep_together = True


def _is_blank_p(el):
    return (el.tag == qn("w:p") and not "".join(xp(el, ".//w:t/text()")).strip()
            and not xp(el, ".//w:drawing | .//mc:AlternateContent | .//w:pict"))


def _only_floating_drawings(el):
    """Empty paragraph that only holds floating (anchored) pictures: its own height is wasted space."""
    return (el.tag == qn("w:p") and not "".join(xp(el, ".//w:t/text()")).strip()
            and xp(el, ".//wp:anchor") and not xp(el, ".//wp:inline"))


def compact_layout(doc, spacer_pt=6):
    """Stop wasting pages: no forced page breaks, no stacks of empty paragraphs, no empty paragraphs at the very end
    (they can create a blank last page). One small spacer stays between tables so Word does not merge them."""
    body = doc.element.body
    for el in xp(body, ".//w:pageBreakBefore"):
        el.getparent().remove(el)
    for br in xp(body, './/w:br[@w:type="page"]'):
        run = br.getparent()
        run.remove(br)
        if len(run) == 0 or (len(run) == 1 and run[0].tag == qn("w:rPr")):
            run.getparent().remove(run)
    kids = [c for c in body if c.tag != qn("w:sectPr")]
    while kids and _is_blank_p(kids[-1]):  # trailing empty paragraphs
        body.remove(kids.pop())
    prev_blank = False
    for el in list(kids):
        if _only_floating_drawings(el):
            pf = Paragraph(el, None).paragraph_format
            pf.space_before, pf.space_after, pf.line_spacing = Pt(0), Pt(0), Pt(2)
            prev_blank = False
            continue
        if not _is_blank_p(el):
            prev_blank = False
            continue
        if prev_blank:  # collapse stacks of empty paragraphs
            body.remove(el)
            continue
        pf = Paragraph(el, None).paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing = Pt(spacer_pt)
        prev_blank = True


def shrink_tall_rows(table, max_twips=700):
    """Template rows were sized for placeholder text (e.g. 11 lines in 'Primary packaging'). After filling, drop minimum
    heights above ~0.5 inch so each row only takes the space its real content needs."""
    for row in table.rows:
        trpr = row._tr.find(qn("w:trPr"))
        if trpr is None:
            continue
        for h in trpr.findall(qn("w:trHeight")):
            try:
                if int(h.get(qn("w:val"))) > max_twips:
                    trpr.remove(h)
            except (TypeError, ValueError):
                pass


def keep_last_row_with_next(table):
    """Ask Word to keep the last table row on the same page as what follows (the signature block)."""
    for cell in table.rows[-1].cells:
        for p in cell.paragraphs:
            p.paragraph_format.keep_with_next = True


def keep_table_together(table):
    """Small table that sits right above the signature: keep every row with the next one so the table and the
    signature move to the next page together instead of leaving a single orphan row behind."""
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.keep_with_next = True


def docx_to_pdf(docx_bytes, timeout=180):
    """DOCX -> PDF with LibreOffice (headless). In Colab: !apt-get -qq install -y libreoffice-writer
    SOFFICE_CMD can point to a custom command."""
    cmd = shlex.split(os.getenv("SOFFICE_CMD", "")) or [shutil.which("soffice") or shutil.which("libreoffice") or ""]
    if not cmd[0]:
        raise RuntimeError("LibreOffice is not installed, so no PDF can be made. In Colab run: "
                           "!apt-get -qq install -y libreoffice-writer")
    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, "document.docx")
        with open(src, "wb") as f:
            f.write(docx_bytes)
        subprocess.run(cmd + ["--headless", f"-env:UserInstallation=file://{tmp}/lo", "--convert-to", "pdf", "--outdir", tmp, src],
                       check=True, timeout=timeout, capture_output=True)
        out = os.path.join(tmp, "document.pdf")
        if not os.path.exists(out):
            raise RuntimeError("LibreOffice did not produce a PDF.")
        with open(out, "rb") as f:
            return f.read()
