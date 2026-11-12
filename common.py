"""Small helpers shared by every module."""
import math
import re

import pandas as pd

UNIT_TO_MG = {"mg": 1.0, "g": 1000.0, "kg": 1e6, "mcg": 0.001, "µg": 0.001, "μg": 0.001, "ug": 0.001}


def clean(v):
    """Any value -> stripped string ('' for None/NaN)."""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return str(v).strip()


def to_float(v):
    """'1.530,61' / '7,5' / 12.0 -> float, or None if not a number."""
    s = clean(v).replace(" ", "")
    if not s:
        return None
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def fmt_num(x):
    return f"{x:.4f}".rstrip("0").rstrip(".")


def fmt_amount(amount, unit):
    a = to_float(amount)
    return "" if a is None else f"{fmt_num(a)} {clean(unit)}".strip()


def norm(s):
    return re.sub(r"[^0-9a-z\u00c0-\u024f]+", "", clean(s).lower())


def same(a, b):
    fa, fb = to_float(a), to_float(b)
    if fa is not None and fb is not None:
        return math.isclose(fa, fb, rel_tol=1e-9, abs_tol=1e-12)
    return clean(a) == clean(b)


def form_plural_title(form):
    """'capsule' -> 'Capsules'."""
    f = clean(form)
    if not f:
        return ""
    pl = f if f.lower().endswith("s") else f + "s"
    return pl[0].upper() + pl[1:]


def form_plural_lower(form):
    return form_plural_title(form).lower()


def heading_text(product):
    """'CANNON 30 Capsules'"""
    q = to_float(product.get("quantity"))
    parts = [clean(product.get("display_name")) or clean(product.get("name")), fmt_num(q) if q is not None else "",
             form_plural_title(product.get("dosage_form"))]
    return " ".join(p for p in parts if p)


def weight_range(product):
    """(value, min, max) in the TDS unit, computed ONLY from the TDS weight and tolerance."""
    v, t = to_float(product.get("unit_weight_value")), to_float(product.get("tolerance"))
    if v is None or t is None:
        return v, None, None
    return v, v * (1 - t / 100), v * (1 + t / 100)


def weight_sentence(product, spaced_percent=False):
    """'650 mg ± 7.5% (Between 601.25 mg and 698.75 mg)'"""
    v, lo, hi = weight_range(product)
    if v is None:
        return ""
    u, t = clean(product.get("unit_weight_unit")), to_float(product.get("tolerance"))
    pct = f" ± {fmt_num(t)}{' ' if spaced_percent else ''}%" if t is not None else ""
    rng = f" (Between {fmt_num(lo)} {u} and {fmt_num(hi)} {u})" if lo is not None else ""
    return f"{fmt_num(v)} {u}{pct}{rng}".strip()


def numbers_in(text):
    """All numbers in a text, trying both '1.500' = 1500 and 1.5 readings."""
    out = set()
    for tok in re.findall(r"\d[\d.,]*\d|\d", text):
        variants = []
        if "," in tok and "." in tok:
            variants.append(tok.replace(".", "").replace(",", ".") if tok.rfind(",") > tok.rfind(".") else tok.replace(",", ""))
        elif "," in tok:
            variants += [tok.replace(",", "."), tok.replace(",", "")]
        elif "." in tok:
            variants.append(tok)
            parts = tok.split(".")
            if all(len(x) == 3 for x in parts[1:]):
                variants.append(tok.replace(".", ""))
        else:
            variants.append(tok)
        for v in variants:
            try:
                out.add(float(v))
            except ValueError:
                pass
    return out


def in_numbers(value, nums):
    return any(math.isclose(value, n, rel_tol=1e-6, abs_tol=1e-9) for n in nums)


def nrv_per_unit(nrv, basis, daily_dose):
    """%NRV printed in the TDS -> %NRV for ONE unit (capsule / sachet).
    The TDS value is usually for the whole daily dose: 3 capsules a day -> divide by 3; 1 a day -> unchanged.
    Returns None when it cannot be converted safely (unknown basis, or daily dose missing)."""
    v = to_float(nrv)
    if v is None:
        return None
    b = clean(basis).lower()
    if b == "per_unit":
        return v
    if b == "per_daily_dose":
        d = to_float(daily_dose)
        return round(v / d, 2) if d and d > 0 else None
    return None
