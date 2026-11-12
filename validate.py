"""Validation rules + comparison between the original extraction and manual corrections."""
import re

from common import UNIT_TO_MG, clean, fmt_num, in_numbers, norm, numbers_in, same, to_float
from extract import PRODUCT_GROUPS, READONLY

# FPS fields: if the TDS does not give them they are simply left blank (never a blocker).
FPS_OPTIONAL = [k for g, items in PRODUCT_GROUPS.items() if g.startswith("FPS") for k, _ in items]


def page_corpus(p):
    rows = [" ".join(r) for t in p["tables"] for r in t]
    return p["text"] + "\n" + "\n".join(rows)


def validate(product, actives, subs, excs, pages, conflicts):
    flags = []

    def add(sev, where, issue):
        flags.append({"severity": sev, "where": where, "issue": issue})

    nums = set()
    for p in pages:
        nums |= numbers_in(page_corpus(p))

    def check_amount(where, value, label="Amount"):
        if value is not None and not in_numbers(value, nums):
            add("WARNING", where, f"{label} {fmt_num(value)} does not appear anywhere in the TDS - verify")

    def check_source(where, source_text, amount):
        """The name must be found in the PDF, and (for table rows) the amount must be on that same row."""
        if not clean(source_text):
            add("WARNING", where, "Name could not be located in the PDF text - verify manually")
        elif " | " in source_text and amount is not None and not in_numbers(amount, numbers_in(source_text)):
            add("WARNING", where, f"Amount {fmt_num(amount)} is not on the same PDF row as the name - possible column mix-up")

    # --- product (core fields block generation)
    core = [("name", "Product name"), ("display_name", "Name used in documents"), ("quantity", "Quantity per package"),
            ("dosage_form", "Dosage form"), ("unit_weight_value", "Dosage form weight"),
            ("unit_weight_unit", "Weight unit"), ("tolerance", "Weight tolerance")]
    for key, label in core:
        v = clean(product.get(key))
        if not v:
            add("CRITICAL", "Product", f"{label} is missing")
        elif key in ("quantity", "unit_weight_value", "tolerance"):
            f = to_float(v)
            if f is None or f <= 0:
                add("CRITICAL", "Product", f"{label} must be a positive number")
            else:
                check_amount("Product", f, label)

    if re.search(r"(?<!\S)0+(?!\S)", clean(product.get("display_name"))):
        add("CRITICAL", "Product", "The name used in documents contains a stray '0' - remove it")

    # --- actives
    if actives.empty:
        add("CRITICAL", "Actives", "No active ingredients")
    top = [norm(n) for n in actives["name"]]
    if len(set(top)) != len(top):
        add("WARNING", "Actives", "Duplicate ingredient names")
    for _, r in actives.iterrows():
        where = "Active: " + (clean(r["name"]) or clean(r["row_id"]) or "(new row)")
        if not clean(r["name"]):
            add("CRITICAL", where, "Name is missing")
        amt = to_float(r["amount"])
        if amt is None or amt <= 0:
            add("CRITICAL", where, "Amount is missing or not a positive number")
        else:
            check_amount(where, amt)
        if not clean(r["unit"]):
            add("CRITICAL", where, "Unit is missing")
        if clean(r["amount_basis"]) != "per_unit":
            add("CRITICAL", where, f"Amount basis is '{clean(r['amount_basis']) or 'unknown'}' - must be confirmed as per_unit")
        if not clean(r["function"]):
            add("WARNING", where, "Function not stated in the TDS (left blank in the COC unless you fill it)")
        if clean(r["nrv_percent"]):
            nrv = to_float(r["nrv_percent"])
            if nrv is None:
                add("CRITICAL", where, "%NRV is not a number")
            else:
                check_amount(where, nrv, "%NRV")
        if clean(r["review_reason"]):
            add("WARNING", where, "Model asked for review: " + clean(r["review_reason"]))
        check_source(where, r["source_text"], amt)
        for col in ("min_amount", "max_amount"):
            if clean(r[col]) and to_float(r[col]) is None:
                add("CRITICAL", where, f"{col} is not a number")

    # --- %NRV: the TDS value is usually for the whole daily dose, the COC needs it per ONE unit
    if any(clean(x) for x in actives["nrv_percent"]):
        basis, dd = clean(product.get("nrv_basis")).lower(), to_float(product.get("daily_dose"))
        if basis not in ("per_daily_dose", "per_unit"):
            add("CRITICAL", "NRV", "%NRV values exist but the TDS does not say what they are calculated for. Set 'NRV basis' to "
                                   "per_daily_dose or per_unit after checking the TDS")
        elif basis == "per_daily_dose" and (dd is None or dd <= 0):
            add("CRITICAL", "NRV", "%NRV is per daily dose but the daily dose (number of units per day) is missing")
        elif basis == "per_daily_dose":
            add("INFO", "NRV", f"%NRV in the TDS is for the daily dose: every value is divided by {fmt_num(dd)} to get "
                               "the per-unit %NRV written in the COC")

    # --- sub-components
    for _, r in subs.iterrows():
        where = f"Sub-component: {clean(r['name']) or clean(r['row_id'])}"
        parent = norm(r["parent"])
        if parent not in top:
            add("CRITICAL", where, f"Parent '{clean(r['parent'])}' does not match any active ingredient")
        if norm(r["name"]) in top and norm(r["name"]) != parent:
            add("CRITICAL", where, "Same name as a different top-level ingredient - possible double counting")
        amt = to_float(r["amount"])
        if amt is None and not clean(r["standardization"]):
            add("WARNING", where, "Neither amount nor standardization given")
        if amt is not None:
            check_amount(where, amt)
            if not clean(r["unit"]):
                add("CRITICAL", where, "Unit is missing")
        if clean(r["row_id"]):
            check_source(where, r["source_text"], amt)

    # --- excipients
    for _, r in excs.iterrows():
        where = "Excipient: " + (clean(r["name"]) or clean(r["row_id"]) or "(new row)")
        if not clean(r["name"]):
            add("CRITICAL", where, "Name is missing")
        amt = to_float(r["amount"])
        if amt is None or amt <= 0:
            add("CRITICAL", where, "Amount is missing or not a positive number")
        else:
            check_amount(where, amt)
        if not clean(r["unit"]):
            add("CRITICAL", where, "Unit is missing")
        if clean(r["amount_basis"]) != "per_unit":
            add("CRITICAL", where, f"Amount basis is '{clean(r['amount_basis']) or 'unknown'}' - must be confirmed as per_unit")
        if not clean(r["function"]):
            add("WARNING", where, "Function not stated in the TDS")
        if clean(r["review_reason"]):
            add("WARNING", where, "Model asked for review: " + clean(r["review_reason"]))
        check_source(where, r["source_text"], amt)

    # --- do the listed components add up to the unit weight?
    w = to_float(product.get("unit_weight_value"))
    wf = UNIT_TO_MG.get(clean(product.get("unit_weight_unit")).lower())
    if w and wf:
        total, skipped = 0.0, 0
        for df in (actives, excs):
            for _, r in df.iterrows():
                a, f = to_float(r["amount"]), UNIT_TO_MG.get(clean(r["unit"]).lower())
                if a is None or f is None:
                    skipped += 1
                else:
                    total += a * f
        target = w * wf
        if skipped == 0 and abs(total - target) > 0.02 * target:
            add("WARNING", "Consistency", f"Listed ingredients sum to {fmt_num(total)} mg but dosage form weight is "
                                          f"{fmt_num(target)} mg")

    for c in conflicts:
        add("WARNING", "Conflict reported by model", f"{c.get('field')}: {c.get('description')}")

    # --- FPS fields that will stay blank (information only)
    blank = [k for k in FPS_OPTIONAL if not clean(product.get(k))]
    if blank:
        add("INFO", "FPS", "Left blank (not in the TDS): " + ", ".join(blank))
    return flags


def diff_frames(label, orig, edited):
    changes = []
    editable = [c for c in orig.columns if c not in READONLY and c != "row_id"]
    o = orig.set_index("row_id")
    seen = set()
    for _, r in edited.iterrows():
        rid = clean(r.get("row_id"))
        if not rid or rid not in o.index:
            changes.append(f"{label}: added row '{clean(r.get('name'))}'")
            continue
        seen.add(rid)
        for c in editable:
            if not same(o.loc[rid, c], r[c]):
                changes.append(f"{label} {rid} [{c}]: '{clean(o.loc[rid, c])}' -> '{clean(r[c])}'")
    for rid in o.index:
        if rid not in seen:
            changes.append(f"{label}: deleted row {rid} ('{clean(o.loc[rid].get('name'))}')")
    return changes
