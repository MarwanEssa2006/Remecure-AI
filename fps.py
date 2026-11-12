"""Finished Product Specification: fills the official FPS template. Anything the TDS does not state stays blank."""
import copy
import datetime
import io

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from common import clean, fmt_amount, fmt_num, form_plural_lower, form_plural_title, heading_text, norm, to_float, weight_sentence
from docx_utils import (compact_layout, fill_placeholders, fix_signature, keep_last_row_with_next, no_split_rows, remove_paragraph, set_text,
                        shrink_tall_rows)

# (section, label) -> product field. Label = first word of the template line ("Dimensions / capacity" -> "Dimensions").
PRIMARY = {("container", "Type"): "pkg_type", ("container", "Material"): "pkg_material", ("container", "Color"): "pkg_color",
           ("container", "Dimensions"): "pkg_dimensions", ("closure", "Type"): "closure_type",
           ("closure", "Material"): "closure_material", ("closure", "Color"): "closure_color",
           ("protective", ""): "protective_filling"}
SECONDARY = {("label", "Material"): "label_material", ("label", "Dimensions"): "label_dimensions",
             ("carton", "Material"): "carton_material", ("carton", "Dimensions"): "carton_dimensions",
             ("carton", "Design"): "carton_design", ("leaflet", ""): "leaflet"}
REMOVE_IF_EMPTY = {"closure", "protective", "label", "leaflet"}
LEAF_LABELS = ("Type", "Material", "Color", "Dimensions", "Design")


def _fill_sections(cell, mapping, product):
    """Fill 'Label: [ ]' lines section by section; drop optional sections that end up empty."""
    sections, current = {}, None
    for p in list(cell.paragraphs):
        t = p.text.strip()
        first = t.split(":")[0].split(" ")[0].strip()
        if t.endswith(":") and first not in LEAF_LABELS:
            current = first.lower().rstrip("/")
            sections[current] = {"header": p, "lines": []}
        elif current:
            sections[current]["lines"].append((p, first if first in LEAF_LABELS else ""))
    for sec, info in sections.items():
        any_value = False
        for p, label in info["lines"]:
            value = clean(product.get(mapping.get((sec, label), ""), ""))
            any_value |= bool(value)
            if label:
                fill_placeholders(p, value)
            else:  # free line (protective filling / leaflet)
                set_text(p, value)
        if not any_value and sec in REMOVE_IF_EMPTY:
            remove_paragraph(info["header"])
            for p, _ in info["lines"]:
                remove_paragraph(p)


def _restore_borders(tc, top=False, bottom=False):
    """Template cells inside the block carry explicit 'nil' top/bottom borders (one visual block).
    The first / last row of the block gets the normal table border back by dropping that 'nil'."""
    tcpr = tc.find(qn("w:tcPr"))
    b = tcpr.find(qn("w:tcBorders")) if tcpr is not None else None
    if b is None:
        return
    for side, on in (("w:top", top), ("w:bottom", bottom)):
        el = b.find(qn(side))
        if on and el is not None:
            b.remove(el)


def _nil_bottom(tc):
    """Left-column cells of the block: no line under them (the block is read as one cell)."""
    tcpr = tc.find(qn("w:tcPr"))
    b = tcpr.find(qn("w:tcBorders"))
    if b is None:
        b = OxmlElement("w:tcBorders")
        tcpr.append(b)
    if b.find(qn("w:bottom")) is None:
        el = OxmlElement("w:bottom")
        el.set(qn("w:val"), "nil")
        top = b.find(qn("w:top"))
        (top.addnext(el) if top is not None else b.insert(0, el))


def _rebuild_assay(table, unit_lower, actives, subs):
    """Template: one merged 'Assay' cell whose lines are aligned by blank paragraphs with a separate 'Between...' column.
    That cannot grow safely, so each ingredient / sub-component becomes its own row (same fonts, same columns)."""
    rows = list(table.rows)
    r1, r2 = rows[1], rows[2]
    left_paras = r1._tr.findall(qn("w:tc"))[0].findall(qn("w:p"))
    p_assay, p_each, p_ing, p_sub = (copy.deepcopy(p) for p in left_paras[:4])
    p_spec = copy.deepcopy(r2._tr.findall(qn("w:tc"))[1].findall(qn("w:p"))[0])
    for p in (p_ing, p_sub, p_spec):  # tight and identical spacing on both sides so every spec value stays level with its line
        sp = p.find(qn("w:pPr")).find(qn("w:spacing"))
        if sp is not None:
            sp.set(qn("w:before"), "0")
    proto_tr = copy.deepcopy(r2._tr)
    for r in rows[1:7]:
        table._tbl.remove(r._tr)

    def make_row(left_ps, left_texts, spec_text, first=False, last=False):
        tr = copy.deepcopy(proto_tr)
        tcs = tr.findall(qn("w:tc"))
        for tc in tcs:
            for p in tc.findall(qn("w:p")):
                tc.remove(p)
        vm = tcs[0].find(qn("w:tcPr")).find(qn("w:vMerge"))
        if vm is not None:
            vm.getparent().remove(vm)
        for src, txt in zip(left_ps, left_texts):
            p = copy.deepcopy(src)
            tcs[0].append(p)
            set_text(Paragraph(p, None), txt)
        p = copy.deepcopy(p_spec)
        tcs[1].append(p)
        set_text(Paragraph(p, None), spec_text)
        trpr = tr.find(qn("w:trPr"))
        if trpr is not None:
            for h in trpr.findall(qn("w:trHeight")):
                trpr.remove(h)
        for tc in tcs:
            _restore_borders(tc, top=first, bottom=last)
        if not last:
            _nil_bottom(tcs[0])
        return tr

    def spec(lo, hi, unit):
        lo_f, hi_f = to_float(lo), to_float(hi)
        if lo_f is None or hi_f is None:
            return ""
        return f"Between {fmt_num(lo_f)} and {fmt_num(hi_f)} {clean(unit)}".strip()

    new_rows = [make_row([p_assay, p_each], ["Assay:", f"** Each 1 {unit_lower} contains:"], "", first=True)]
    for _, a in actives.iterrows():
        line = f"{clean(a['name'])} {fmt_amount(a['amount'], a['unit'])}".strip()
        new_rows.append(make_row([p_ing], [line], spec(a["min_amount"], a["max_amount"], a["unit"])))
        for _, s in subs[subs["parent"].map(norm) == norm(a["name"])].iterrows():
            label = clean(s["name"]) + (f" ({clean(s['standardization'])})" if clean(s["standardization"]) else "")
            new_rows.append(make_row([p_sub], [f"- {label} {fmt_amount(s['amount'], s['unit'])}".strip()],
                                     spec(s["min_amount"], s["max_amount"], s["unit"])))
    for tc in new_rows[-1].findall(qn("w:tc")):  # the block's last row closes with the normal bottom border
        _restore_borders(tc, bottom=True)
    prev = table.rows[0]._tr
    for tr in new_rows:
        prev.addnext(tr)
        prev = tr


def build_fps(template_bytes, product, actives, subs, micro):
    doc = Document(io.BytesIO(template_bytes))
    form = clean(product["dosage_form"])
    unit_lower = form.lower()
    t_info, t_assay, t_org, t_add = doc.tables[0], doc.tables[1], doc.tables[2], doc.tables[3]

    def paragraph(marker):
        return next(p for p in doc.paragraphs if marker in p.text)

    set_text(paragraph("[Product Name]"), heading_text(product))
    set_text(paragraph("Date:"), "Date: " + datetime.date.today().strftime("%d/%m/%Y"))

    # --- general information table
    cell = lambda r: t_info.rows[r].cells[1]
    set_text(cell(0).paragraphs[0], clean(product["denomination"]))
    set_text(cell(1).paragraphs[0], clean(product["format"]))
    set_text(cell(2).paragraphs[0], weight_sentence(product, spaced_percent=True))
    q = to_float(product["quantity"])
    set_text(cell(3).paragraphs[0], f"{fmt_num(q)} {form_plural_title(form)}" if q is not None else "")
    set_text(cell(4).paragraphs[0], clean(product["net_content"]))
    set_text(cell(4).paragraphs[1], clean(product["net_content_labeling"]))
    _fill_sections(cell(5), PRIMARY, product)
    _fill_sections(cell(6), SECONDARY, product)
    shelf = clean(product["shelf_life"])
    set_text(cell(7).paragraphs[0], f"{shelf} from manufacturing date" if shelf else "")
    set_text(cell(8).paragraphs[0], clean(product["storage"]))
    dd = to_float(product["daily_dose"])
    unit_word = form.lower() if dd == 1 else form_plural_lower(form)
    set_text(cell(9).paragraphs[0], f"{fmt_num(dd)} {unit_word} per day" if dd is not None else "")

    # --- assay / heavy metals / microbial
    rows = list(t_assay.rows)
    hm_rows, micro_proto_tr = rows[7], copy.deepcopy(rows[10]._tr)
    micro_head = rows[9]._tr
    for r in rows[10:]:
        t_assay._tbl.remove(r._tr)
    hm_vals = [clean(product[k]) for k in ("lead", "cadmium", "mercury")]
    hm_par = hm_rows.cells[1].paragraphs[-1]
    set_text(hm_par, "\n".join(hm_vals) if any(hm_vals) else "")
    prev = micro_head
    for _, m in micro.iterrows():
        tr = copy.deepcopy(micro_proto_tr)
        tcs = tr.findall(qn("w:tc"))
        set_text(Paragraph(tcs[0].findall(qn("w:p"))[0], None), clean(m["name"]))
        set_text(Paragraph(tcs[1].findall(qn("w:p"))[0], None), clean(m["limit"]))
        prev.addnext(tr)
        prev = tr
    _rebuild_assay(t_assay, unit_lower, actives, subs)

    # --- organoleptic
    for r, value in ((1, product["org_description"]), (2, product["org_size"]),
                     (3, weight_sentence(product, spaced_percent=True)), (4, product["org_color"]),
                     (5, product["org_odour"]), (6, product["org_taste"])):
        set_text(t_org.rows[r].cells[1].paragraphs[0], clean(value))

    # --- additional information: fixed text stays; the placeholder row is removed
    last = t_add.rows[-1]
    if "[Additional information" in last.cells[0].text:
        t_add._tbl.remove(last._tr)

    no_split_rows(t_assay)
    for t in (t_info, t_org, t_add):
        shrink_tall_rows(t)
    fix_signature(doc)
    compact_layout(doc)
    keep_last_row_with_next(t_add)
    out = io.BytesIO()
    doc.save(out)
    return out.getvalue(), []
