"""
Remecure AI - COC + FPS Generator (MVP)
=======================================
TDS PDF -> Gemini extraction -> validation -> human review -> COC and FPS (.docx) from the official templates.

Files that must sit together:  app.py common.py extract.py validate.py docx_utils.py coc.py fps.py
PDF download needs LibreOffice (Colab: !apt-get -qq install -y libreoffice-writer)
Templates (or upload them in the sidebar):  COC_Template.docx  FPS_Template.docx
API key:  GEMINI_API_KEY in .env or in Colab Secrets
Run:      streamlit run app.py
"""
import datetime
import json
import os
import re
import time

import pandas as pd
import streamlit as st

from coc import build_coc
from common import clean, nrv_per_unit, same
from docx_utils import docx_to_pdf
from extract import (FALLBACK_MODELS, MAX_TRIES, MODEL, PRODUCT_GROUPS, PRODUCT_KEYS, READONLY, cached_extract, find_reg_footnote, get_api_key,
                     pages_to_prompt, read_pdf, select_pages, to_frames)
from validate import diff_frames, validate
from fps import build_fps

BASIS_OPTIONS = ["per_unit", "per_daily_dose", "per_package", "unknown"]
TEMPLATES = {
    "coc": ("COC template (.docx)", ["COC_Template.docx", "coc_template.docx"]),
    "fps": ("FPS template (.docx)", ["FPS_Template.docx", "fps_template.docx"]),
}


def load_templates(sidebar):
    ss = st.session_state
    for key, (label, paths) in TEMPLATES.items():
        up = sidebar.file_uploader(label, type="docx", key=f"up_{key}")
        if up:
            ss[f"tpl_{key}"] = up.getvalue()
        elif f"tpl_{key}" not in ss:
            for path in paths:
                if os.path.exists(path):
                    with open(path, "rb") as f:
                        ss[f"tpl_{key}"] = f.read()
                    break
        if f"tpl_{key}" in ss:
            sidebar.success(f"{key.upper()} template loaded")
        else:
            sidebar.error(f"{key.upper()} template not found. Upload it here or save it as '{paths[0]}'.")


def pdf_pages(pdf_bytes):
    import io

    import pdfplumber

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        return len(pdf.pages)


def run_extraction(pdf, send_all, use_cache, strip_formula):
    ss = st.session_state
    t0 = time.time()
    with st.status("Extracting...", expanded=True) as status:
        st.write("Reading the PDF...")
        pages = read_pdf(pdf.getvalue())
        if not any(p["text"].strip() for p in pages):
            status.update(label="No text found in this PDF (scanned?). OCR is not supported.", state="error")
            return
        chosen = select_pages(pages, send_all)
        prompt = pages_to_prompt(chosen, pages)
        st.write(f"Sending pages {[p['page'] for p in chosen]} of {len(pages)} ({len(prompt):,} characters) to {MODEL}")
        box = st.empty()
        try:
            data, raw, from_cache, used = cached_extract(prompt, use_cache, lambda n: box.write(f"Receiving answer... {n:,} characters"),
                                                         notify=lambda msg: st.write(msg))
        except Exception as e:  # show the real error, never fake success
            status.update(label=f"Extraction failed: {e}", state="error")
            return
        product, frames, conflicts = to_frames(data, pages, strip_formula)
        ss["ex"] = {"id": ss.get("ex", {}).get("id", 0) + 1, "pages": pages, "sent": [p["page"] for p in chosen], "raw": raw,
                    "product": product, "frames": frames, "conflicts": conflicts, "file": pdf.name,
                    "footnote": find_reg_footnote(pages)}
        for k in ("coc_docx", "fps_docx", "coc_notes", "fps_notes", "coc_pdf", "fps_pdf", "coc_pdf_err", "fps_pdf_err"):
            ss.pop(k, None)
        status.update(label=f"Done in {time.time() - t0:.1f}s" + (" (from cache)" if from_cache else f" - answered by {used}"),
                      state="complete")


def main():
    st.set_page_config(page_title="Remecure AI - COC + FPS Generator", layout="wide")
    st.title("Remecure AI - COC + FPS Generator")
    st.caption("One TDS -> one extraction -> review -> Certificate of Composition and Finished Product Specification.")
    ss = st.session_state

    with st.sidebar:
        st.header("Setup")
        load_templates(st.sidebar)
        if get_api_key():
            st.success("GEMINI_API_KEY found")
        else:
            st.error("GEMINI_API_KEY not found (.env or Colab Secrets).")
        st.caption(f"Model: {MODEL}")
        st.caption(f"If it is busy: {MAX_TRIES} attempts per model, then fallback: {', '.join(FALLBACK_MODELS) or 'none'}")
        st.header("Options")
        send_all = st.checkbox("Send ALL pages to the model (slower)", value=False)
        use_cache = st.checkbox("Reuse previous answer for the same TDS", value=True)
        strip_formula = st.checkbox("Remove trailing 'FORMULA' from the product name", value=True)

    pdf = st.file_uploader("1. Upload the TDS (PDF)", type="pdf")
    if pdf and st.button("2. Extract data", type="primary"):
        if not get_api_key():
            st.error("Cannot extract: no API key configured.")
        else:
            run_extraction(pdf, send_all, use_cache, strip_formula)

    ex = ss.get("ex")
    if not ex:
        st.info("Upload a TDS and click Extract to begin.")
        return
    xid = ex["id"]
    st.subheader(f"3. Review - {ex['file']}")
    st.caption(f"Pages sent to the model: {ex['sent']} (of {len(ex['pages'])})")

    # ---- product fields (grouped)
    product = {}
    for gi, (group, fields) in enumerate(PRODUCT_GROUPS.items()):
        with st.expander(group, expanded=(gi == 0)):
            cols = st.columns(3)
            for i, (key, label) in enumerate(fields):
                product[key] = cols[i % 3].text_input(label, value=ex["product"][key], key=f"p_{key}_{xid}")

    # ---- ingredient tables
    cfg_num = st.column_config.NumberColumn
    st.markdown("**Active ingredients** (top-level ingredient as formulated, per unit)")
    actives = st.data_editor(
        ex["frames"]["actives"], num_rows="dynamic", width="stretch", key=f"act_{xid}", disabled=READONLY,
        column_config={"amount": cfg_num("amount"), "nrv_percent": cfg_num("nrv_percent"), "min_amount": cfg_num("min_amount"),
                       "max_amount": cfg_num("max_amount"), "amount_basis": st.column_config.SelectboxColumn(options=BASIS_OPTIONS)})
    st.markdown("**Standardized compounds / sub-components** (`parent` must match an active ingredient name)")
    subs = st.data_editor(
        ex["frames"]["subs"], num_rows="dynamic", width="stretch", key=f"sub_{xid}",
        disabled=["row_id", "page", "source_text"],
        column_config={"amount": cfg_num("amount"), "min_amount": cfg_num("min_amount"), "max_amount": cfg_num("max_amount")})
    st.markdown("**Excipients**")
    excs = st.data_editor(
        ex["frames"]["excipients"], num_rows="dynamic", width="stretch", key=f"exc_{xid}", disabled=READONLY,
        column_config={"amount": cfg_num("amount"), "amount_basis": st.column_config.SelectboxColumn(options=BASIS_OPTIONS)})
    st.markdown("**Microbiological specification (FPS)**")
    micro = st.data_editor(ex["frames"]["micro"], num_rows="dynamic", width="stretch", key=f"mic_{xid}",
                           disabled=["row_id"])

    if any(clean(x) for x in actives["nrv_percent"]):
        with st.expander("%NRV: value in the TDS vs. value written in the COC (per 1 unit)", expanded=True):
            prev = actives[["name", "nrv_percent"]].rename(columns={"nrv_percent": "NRV in TDS"}).copy()
            prev["NRV per 1 unit (COC)"] = [nrv_per_unit(v, product["nrv_basis"], product["daily_dose"]) for v in actives["nrv_percent"]]
            st.caption(f"NRV basis: {clean(product['nrv_basis']) or 'unknown'} | units per day: {clean(product['daily_dose']) or '?'} "
                       "| per_daily_dose -> value / units per day; per_unit -> unchanged")
            st.dataframe(prev[prev["NRV in TDS"].map(lambda v: clean(v) != "")], width="stretch", hide_index=True)

    # ---- validation
    flags = validate(product, actives, subs, excs, ex["pages"], ex["conflicts"])
    crit = [f for f in flags if f["severity"] == "CRITICAL"]
    warn = [f for f in flags if f["severity"] == "WARNING"]
    info = [f for f in flags if f["severity"] == "INFO"]
    st.subheader("4. Validation")
    if crit:
        st.error(f"{len(crit)} critical issue(s) - nothing can be generated until they are fixed.")
        st.dataframe(pd.DataFrame(crit), width="stretch", hide_index=True)
    else:
        st.success("No critical issues.")
    if warn:
        st.warning(f"{len(warn)} warning(s) to review.")
        st.dataframe(pd.DataFrame(warn), width="stretch", hide_index=True)
    for f in info:
        st.info(f"{f['where']}: {f['issue']}")

    # ---- original vs corrected
    changes = [f"Product [{k}]: '{ex['product'][k]}' -> '{product[k]}'" for k in PRODUCT_KEYS if not same(ex["product"][k], product[k])]
    changes += diff_frames("Active", ex["frames"]["actives"], actives)
    changes += diff_frames("Sub-component", ex["frames"]["subs"], subs)
    changes += diff_frames("Excipient", ex["frames"]["excipients"], excs)
    changes += diff_frames("Micro", ex["frames"]["micro"], micro)
    with st.expander(f"Manual corrections vs. original extraction ({len(changes)})"):
        st.write("\n".join(f"- {c}" for c in changes) if changes else "No manual changes.")
    with st.expander("Raw LLM answer"):
        st.code(ex["raw"], language="json")
    report = {"original": {"product": ex["product"], **{k: v.to_dict("records") for k, v in ex["frames"].items()}},
              "corrected": {"product": product, "actives": actives.to_dict("records"), "subs": subs.to_dict("records"),
                            "excipients": excs.to_dict("records"), "micro": micro.to_dict("records")},
              "manual_changes": changes, "validation": flags}
    st.download_button("Download extraction report (JSON)", json.dumps(report, ensure_ascii=False, indent=2, default=str),
                       file_name="extraction_report.json", mime="application/json")

    # ---- generate
    st.subheader("5. Generate documents")
    reviewed = st.checkbox("I reviewed all warnings and the data matches the TDS", key=f"rev_{xid}")
    safe = re.sub(r"[^A-Za-z0-9]+", "_", clean(product["display_name"]) or "product")
    stamp = f"{datetime.date.today():%Y%m%d}"
    docx_mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    c1, c2 = st.columns(2)
    for col, kind, builder in ((c1, "coc", build_coc), (c2, "fps", build_fps)):
        tpl_ok = f"tpl_{kind}" in ss
        if not tpl_ok:
            col.error(f"{kind.upper()} template missing - cannot generate.")
        if col.button(f"Generate {kind.upper()}", disabled=bool(crit) or not reviewed or not tpl_ok, key=f"gen_{kind}"):
            try:
                if kind == "coc":
                    data, notes = build_coc(ss["tpl_coc"], product, actives, subs, excs, ex["footnote"])
                else:
                    data, notes = build_fps(ss["tpl_fps"], product, actives, subs, micro)
                ss[f"{kind}_docx"], ss[f"{kind}_notes"] = data, notes
                ss.pop(f"{kind}_pdf", None)
                ss.pop(f"{kind}_pdf_err", None)
                try:
                    with st.spinner("Making the PDF..."):
                        ss[f"{kind}_pdf"] = docx_to_pdf(data)
                except Exception as pdf_err:  # the DOCX is still delivered
                    ss[f"{kind}_pdf_err"] = str(pdf_err)
            except Exception as e:
                col.error(f"Could not generate the {kind.upper()}: {e}")
        if f"{kind}_docx" in ss:
            b1, b2 = col.columns(2)
            b1.download_button(f"Download {kind.upper()} (.docx)", ss[f"{kind}_docx"], mime=docx_mime,
                               file_name=f"{kind.upper()}_{safe}_{stamp}.docx", key=f"dl_{kind}")
            if f"{kind}_pdf" in ss:
                b2.download_button(f"Download {kind.upper()} (.pdf)", ss[f"{kind}_pdf"], mime="application/pdf",
                                   file_name=f"{kind.upper()}_{safe}_{stamp}.pdf", key=f"dlpdf_{kind}")
                col.caption(f"PDF: {pdf_pages(ss[f'{kind}_pdf'])} page(s)")
            elif f"{kind}_pdf_err" in ss:
                col.info(f"PDF not available: {ss[f'{kind}_pdf_err']}")
            for n in ss.get(f"{kind}_notes", []):
                col.warning(n)
    st.caption("If you edit anything after generating, click Generate again.")


if __name__ == "__main__":
    main()
