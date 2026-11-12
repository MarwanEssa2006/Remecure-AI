"""Certificate of Composition: fills the official COC template."""
import datetime
import io

from docx import Document

from common import clean, fmt_amount, fmt_num, heading_text, norm, nrv_per_unit, to_float
from docx_utils import (add_after, compact_layout, fix_signature, keep_table_together, no_split_rows, paragraph_copy_without_drawings,
                        remove_paragraph, replace_in_par, set_text, clone_rows)

FIXED_SENTENCES = ["Provisional composition", "European origin"]  # live in the template, written by the template owner


def _find(doc, marker):
    return next((p for p in doc.paragraphs if marker in p.text), None)


def build_coc(template_bytes, product, actives, subs, excs, reg_footnote=None):
    """Returns (docx_bytes, notes). notes = things the user should know about (template / TDS gaps)."""
    doc = Document(io.BytesIO(template_bytes))
    notes = []
    form = clean(product["dosage_form"])
    form_title = form[:1].upper() + form[1:]
    unit_lower = form.lower()

    set_text(_find(doc, "[Product Name]"), heading_text(product))
    set_text(_find(doc, "Date:"), "Date: " + datetime.date.today().strftime("%d/%m/%Y"))
    tol = to_float(product["tolerance"])
    set_text(_find(doc, "Weight:"), f"{form_title} Weight: {fmt_num(to_float(product['unit_weight_value']))} "
                                    f"{clean(product['unit_weight_unit'])} ± {fmt_num(tol)}%")

    # Footnotes: the '**Regulation...' sentence comes from the TDS (verbatim); the other two are fixed template text.
    all_text = "\n".join(p.text for p in doc.paragraphs)
    nrv_p = _find(doc, "NRV:")
    if "Guidance document for competent authorities" not in all_text:
        if reg_footnote and nrv_p is not None:
            new_p = paragraph_copy_without_drawings(nrv_p)
            nrv_p._p.addnext(new_p._p)
            set_text(new_p, reg_footnote)
        else:
            notes.append("The '**Regulation (EU) 1169/2011 ...' footnote was not found in the TDS, so it was not added.")
    for marker in FIXED_SENTENCES:
        if marker not in all_text:
            notes.append(f"The template has no fixed sentence containing '{marker}' - add it to the template.")

    t_act, t_exc = doc.tables[0], doc.tables[1]
    replace_in_par(t_act.rows[0].cells[2].paragraphs[0], "[unit]", unit_lower)
    replace_in_par(t_exc.rows[0].cells[2].paragraphs[0], "[unit]", unit_lower)

    # --- active ingredients: the LAST template row (no decorative shapes) is cloned once per ingredient
    for row, (_, a) in zip(clone_rows(t_act, len(actives)), actives.iterrows()):
        mine = subs[subs["parent"].map(norm) == norm(a["name"])]
        c0, c1, c2, c3 = row.cells
        p_name, p_bullet = c0.paragraphs[0], c0.paragraphs[1]
        a_total, a_sub = c2.paragraphs[0], c2.paragraphs[1]
        set_text(p_name, clean(a["name"]))
        set_text(c1.paragraphs[0], clean(a["function"]))
        set_text(a_total, fmt_amount(a["amount"], a["unit"]))
        nrv = nrv_per_unit(a["nrv_percent"], product.get("nrv_basis"), product.get("daily_dose"))  # per ONE unit
        set_text(c3.paragraphs[0], "" if nrv is None else f"{fmt_num(nrv)}%")
        if mine.empty:
            remove_paragraph(p_bullet)
            remove_paragraph(a_sub)
            continue
        prev_b, prev_a = p_bullet, a_sub
        for k, (_, s) in enumerate(mine.iterrows()):
            if k > 0:
                prev_b, prev_a = add_after(prev_b), add_after(prev_a)
            label = clean(s["name"])
            if clean(s["standardization"]):
                label += f" ({clean(s['standardization'])})"
            set_text(prev_b, label)
            set_text(prev_a, fmt_amount(s["amount"], s["unit"]))

    # --- excipients
    for row, (_, e) in zip(clone_rows(t_exc, len(excs)), excs.iterrows()):
        c0, c1, c2 = row.cells
        set_text(c0.paragraphs[0], clean(e["name"]))
        set_text(c1.paragraphs[0], clean(e["function"]))
        unit, amt = clean(e["unit"]), to_float(e["amount"])
        set_text(c2.paragraphs[0], "" if amt is None else (fmt_num(amt) if unit.lower() == "mg" else f"{fmt_num(amt)} {unit}"))

    no_split_rows(t_act)
    no_split_rows(t_exc)
    fix_signature(doc)
    compact_layout(doc)
    keep_table_together(t_exc)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue(), notes
