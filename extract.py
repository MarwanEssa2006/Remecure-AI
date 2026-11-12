"""PDF reading -> compact prompt -> Gemini -> editable tables (+ evidence found by Python)."""
import difflib
import hashlib
import io
import json
import os
import random
import re
import time
import unicodedata
from collections import Counter

import pandas as pd

from common import clean

try:  # .env is optional
    from dotenv import load_dotenv

    load_dotenv()
except ImportError:
    pass

MODEL = os.getenv("GEMINI_MODEL", "gemini-3.8-flash")
# Tried in order when MODEL stays busy / unavailable. Use stable (non-preview) model names from Google AI Studio.
FALLBACK_MODELS = [m.strip() for m in os.getenv("GEMINI_FALLBACK_MODELS", "gemini-2.5-flash").split(",") if m.strip()]
MAX_TRIES = int(os.getenv("GEMINI_MAX_TRIES", "3"))  # attempts per model
TIMEOUT_MS = int(float(os.getenv("GEMINI_TIMEOUT_SEC", "90")) * 1000)
BACKOFF = [3, 8, 15]  # seconds to wait before attempt 2, 3, 4
PROMPT_VERSION = "4"
CACHE_DIR = os.getenv("COC_CACHE_DIR", ".cache")


def get_api_key():
    key = os.getenv("GEMINI_API_KEY")
    if key:
        return key
    try:  # Google Colab Secrets
        from google.colab import userdata

        return userdata.get("GEMINI_API_KEY")
    except Exception:
        return None


# ----------------------------------------------------------------------------
# 1. PDF -> pages
# ----------------------------------------------------------------------------


def read_pdf(pdf_bytes):
    import pdfplumber

    pages = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for i, pg in enumerate(pdf.pages, start=1):
            tables = []
            for t in pg.extract_tables() or []:
                tables.append([[clean(c).replace("\n", " ") for c in row] for row in t])
            pages.append({"page": i, "text": pg.extract_text() or "", "tables": tables})
    return pages


def norm_text(s):
    return re.sub(r"[^0-9a-z]+", "", unicodedata.normalize("NFKD", clean(s).lower()))


# Pages without any of these words are regulatory boilerplate (certifications, disclaimers...).
MARKERS = ["units per package", "uds. por envase", "net weight", "formula's version", "mg/ut", "mg act", "mg/sachet", "mg/caps",
           "mg/tab", "microbiological", "physical-chemical", "organoleptic", "daily dose", "storage & shipping", "heavy metal"]


def select_pages(pages, send_all=False):
    if send_all or len(pages) <= 2:
        return pages
    keep = []
    for p in pages:
        blob = (p["text"] + " " + " ".join(" ".join(r) for t in p["tables"] for r in t)).lower()
        if p["page"] == 1 or any(m in blob for m in MARKERS):
            keep.append(p)
    return keep


def pages_to_prompt(pages, all_pages):
    """Compact text: repeated headers/footers removed, and text already present in a table is not repeated."""
    cnt = Counter()
    for p in all_pages:
        cnt.update({ln.strip() for ln in p["text"].splitlines() if ln.strip()})
    n = len(all_pages)
    repeated = {ln for ln, k in cnt.items() if n >= 3 and k >= 0.6 * n}
    out = []
    for p in pages:
        blob = norm_text(" ".join(c for t in p["tables"] for r in t for c in r))
        lines = [ln.strip() for ln in p["text"].splitlines()
                 if ln.strip() and ln.strip() not in repeated and not (len(norm_text(ln)) > 6 and norm_text(ln) in blob)]
        out.append(f"=== PAGE {p['page']} ===")
        out.extend(lines)
        for k, t in enumerate(p["tables"], start=1):
            out.append(f"[TABLE {k}]")
            out.extend(" | ".join(c for c in r if c) for r in t if any(r))
    return "\n".join(out)


# ----------------------------------------------------------------------------
# 2. Gemini
# ----------------------------------------------------------------------------

SYSTEM_PROMPT = """You extract data from a supplement Technical Data Sheet (TDS) to build a Certificate of Composition (COC) and a Finished Product Specification (FPS).
The TDS layout differs between products (capsules, sachets, tablets...): find fields by MEANING, not by position.
Answer with ONE compact JSON object only (no markdown, no comments).

ABSOLUTE RULES
- Never invent. If a value is not stated in the TDS use null. Never fill from general knowledge.
- Keep ingredient names spelled EXACTLY as printed in the TDS (do not correct spelling).
- Write all text values in English (TDS lines are bilingual "ES // EN": use the English half).
- "function" only when the TDS itself states it (e.g. "(sweetener)", "(anti-caking agent)", "(bulking agent)"). Otherwise null.
- "nrv" only if printed in the TDS. Never calculate it.
- Numbers are JSON numbers. Convert European formats: "1.530,61" -> 1530.61 ; "7,5" -> 7.5.
- All ingredient amounts are PER SINGLE UNIT (one capsule / sachet / tablet). If a value is per daily dose or per package, still report it with basis "per_daily_dose" / "per_package". Use "unknown" if unclear.
- Standardized / active compounds (e.g. withanolides, fulvic acids, or the "Active" column next to a raw material) are "subs" of their parent ingredient. NEVER list them again as top-level ingredients.
- Top-level "amount" = the ingredient as formulated per unit (raw material / extract amount in the formula table). "subs" carry the declared active amount or standardization ("std", e.g. "2.5%"). If the actives column repeats the ingredient itself with another amount, still report it as a sub.
- Active = nutrients, botanicals, functional substances. Excipient = fillers, carriers, flavours, sweeteners, anti-caking agents, capsule shell, etc. Never include packaging. If classification is not clearly supported, put the reason in "review".
- "min"/"max": only the minimum/maximum amounts printed in the TDS composition table for that same line (active, or sub). Otherwise null.
- "page": the PAGE number where the ingredient's formula row appears.
- "nrv_basis": read the footnote / column header that says what the %NRV column is calculated for. "for the daily dose" -> "per_daily_dose"; "per capsule / per sachet / per unit" -> "per_unit"; nothing stated -> "unknown". "nrv_note": that phrase, short and verbatim. Report each "nrv" exactly as printed (do NOT convert it).
- "daily_dose_units": the number of units per day from the TDS daily dose row (e.g. "Capsulas por día 2" -> 2).
- "name": the product name only. Ignore stray cells next to it (such as a lone "0").
- Report contradictions between pages in "conflicts".
- "micro": microbiological specification lines. Write limits in English like "Not more than 10^6 cfu/g" (use the superscript digit, e.g. 10⁶) or "Should be absent in 25 g". Unit "ufc" means cfu.
- "heavy_metals": only if the TDS prints limits (e.g. "Not more than 3 ppm"); otherwise null.

JSON SHAPE (use null for anything missing)
{"product":{"name":str,"quantity_per_package":number,"dosage_form":str(singular, e.g. "Capsule","Sachet"),"format":str(e.g. "Capsule 00","Powder"),
 "denomination":str(e.g. "Food supplement"),"unit_weight_value":number,"unit_weight_unit":str,"weight_tolerance_percent":number,
 "nrv_basis":"per_daily_dose"|"per_unit"|"unknown","nrv_note":str,"net_content":str(e.g. "19.5 g"),"net_content_labeling":str(e.g. "18 g"),"shelf_life":str(e.g. "3 years"),"storage":str,"daily_dose_units":number,
 "primary":{"type":str,"material":str,"color":str,"dimensions":str,"closure_type":str,"closure_material":str,"closure_color":str,"protective_filling":str},
 "secondary":{"label_material":str,"label_dimensions":str,"carton_material":str,"carton_dimensions":str,"carton_design":str,"leaflet":str(description if given, else "Yes" if a leaflet is included)},
 "organoleptic":{"description":str,"size":str,"color":str,"odour":str,"taste":str},
 "heavy_metals":{"lead":str,"cadmium":str,"mercury":str}},
"actives":[{"name":str,"function":str,"amount":number,"unit":str,"basis":"per_unit"|"per_daily_dose"|"per_package"|"unknown","nrv":number,"min":number,"max":number,"page":int,
  "subs":[{"name":str,"amount":number,"unit":str,"std":str,"min":number,"max":number}],"review":str}],
"excipients":[{"name":str,"function":str,"amount":number,"unit":str,"basis":str,"page":int,"review":str}],
"micro":[{"name":str,"limit":str}],
"conflicts":[{"field":str,"note":str}]}"""


def _gemini_stream(client, model, config, prompt, progress):
    chunks, size = [], 0
    for ch in client.models.generate_content_stream(model=model, contents=prompt, config=config):
        if ch.text:
            chunks.append(ch.text)
            size += len(ch.text)
            if progress:
                progress(size)
    return "".join(chunks)


def _generate(client, types, model, prompt, progress):
    """One request to one model. Tries thinking OFF first (fastest), falls back if the model rejects that setting."""
    base = dict(system_instruction=SYSTEM_PROMPT, response_mime_type="application/json", max_output_tokens=16000, temperature=0)
    last = None
    for think in (types.ThinkingConfig(thinking_budget=0), types.ThinkingConfig(thinking_level="low"), None):
        try:
            cfg = types.GenerateContentConfig(**base, **({"thinking_config": think} if think else {}))
            raw = _gemini_stream(client, model, cfg, prompt, progress)
            if not raw.strip():
                raise RuntimeError("Gemini temporarily returned an empty response (unavailable)")
            return raw
        except Exception as ex:
            last = ex
            if "think" not in str(ex).lower():
                raise
    raise last


def _err_code(ex):
    c = getattr(ex, "code", None)
    if isinstance(c, int):
        return c
    m = re.search(r"\b([45]\d\d)\b", str(ex))
    return int(m.group(1)) if m else None


def _is_auth_error(ex):
    return _err_code(ex) in (401, 403) or "api key" in str(ex).lower()


def _is_missing_model(ex):
    s = str(ex).lower()
    return _err_code(ex) == 404 or ("model" in s and "not found" in s)


def _is_retryable(ex):
    s = str(ex).lower()
    return _err_code(ex) in (408, 429, 500, 502, 503, 504) or any(
        w in s for w in ("high demand", "overloaded", "unavailable", "resource_exhausted", "timed out", "timeout", "deadline",
                         "connection", "temporarily"))


def call_llm(prompt_text, progress=None, notify=None):
    """Retries with waiting when the model is busy (503 'high demand' / 429 / timeouts), then tries fallback models.
    notify(msg) is called for every retry / switch so the UI can show what is happening.
    Returns (data, raw_text, model_that_answered)."""
    key = get_api_key()
    if not key:
        raise RuntimeError("GEMINI_API_KEY not found. Add it to .env or Colab Secrets.")
    from google import genai
    from google.genai import types

    say = notify or (lambda m: None)
    client = genai.Client(
    api_key=key,
    http_options=types.HttpOptions(
        timeout=TIMEOUT_MS,
        api_version="v1beta"
    )
    )
    prompt = "TDS CONTENT:\n" + prompt_text
    models = [MODEL] + [m for m in FALLBACK_MODELS if m != MODEL]
    last = None
    for model in models:
        for attempt in range(1, MAX_TRIES + 1):
            try:
                raw = _generate(client, types, model, prompt, progress)
                return parse_json(raw), raw, model
            except Exception as ex:
                last = ex
                if _is_auth_error(ex):
                    raise RuntimeError(f"Gemini rejected the API key ({ex}). Check GEMINI_API_KEY.") from ex
                if _is_missing_model(ex):
                    say(f"Model '{model}' was not found - trying the next one.")
                    break
                if not _is_retryable(ex):
                    raise
                if attempt == MAX_TRIES:
                    say(f"'{model}' is still busy after {MAX_TRIES} attempts.")
                    break
                wait = BACKOFF[min(attempt - 1, len(BACKOFF) - 1)] + random.random()
                say(f"'{model}' is busy ({_err_code(ex) or 'temporary error'}). Retrying in {wait:.0f}s "
                    f"(attempt {attempt + 1} of {MAX_TRIES})...")
                time.sleep(wait)
    raise RuntimeError(f"All Gemini models are busy or unavailable right now. Last error: {last}")


def parse_json(raw):
    a, b = raw.find("{"), raw.rfind("}")
    if a == -1 or b == -1:
        raise ValueError("The model did not return JSON.")
    return json.loads(raw[a : b + 1])


def cached_extract(prompt_text, use_cache=True, progress=None, notify=None):
    """Same TDS text + same model + same prompt -> reuse the previous answer (instant).
    Returns (data, raw, from_cache, model_used)."""
    key = hashlib.sha256(f"{MODEL}|{PROMPT_VERSION}|{prompt_text}".encode()).hexdigest()[:32]
    path = os.path.join(CACHE_DIR, key + ".json")
    if use_cache and os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            d = json.load(f)
        return d["data"], d["raw"], True, d.get("model", MODEL)
    data, raw, used = call_llm(prompt_text, progress, notify)
    os.makedirs(CACHE_DIR, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"data": data, "raw": raw, "model": used}, f, ensure_ascii=False)
    return data, raw, False, used


# ----------------------------------------------------------------------------
# 3. Evidence: found by Python, never trusted from the LLM
# ----------------------------------------------------------------------------


def _tokens(s):
    s = unicodedata.normalize("NFKD", clean(s).lower())
    s = "".join(c for c in s if not unicodedata.combining(c))
    return set(re.findall(r"[a-z0-9]{3,}", s))


def build_rows(pages):
    rows = []
    for p in pages:
        for ln in p["text"].splitlines():
            if ln.strip():
                rows.append((p["page"], ln.strip(), False))
        for t in p["tables"]:
            for r in t:
                txt = " | ".join(c for c in r if c)
                if txt:
                    rows.append((p["page"], txt, True))
    return rows


def _overlap(target, toks):
    hits = len(target & toks)
    if hits < len(target):  # tolerate small spelling differences (e.g. withanolides / whitanolides)
        for t in target - toks:
            if difflib.get_close_matches(t, toks, n=1, cutoff=0.85):
                hits += 1
    return hits / len(target)


def find_evidence(name, hint_page, rows, row_tokens):
    """Best matching PDF row for an ingredient name -> (page, text) or (None, None)."""
    target = _tokens(name)
    if not target:
        return None, None
    best, best_score = None, 0.0
    for i, (pg, _, is_tbl) in enumerate(rows):
        if not (target & row_tokens[i]) and not any(t[:2] == r[:2] for t in target for r in row_tokens[i]):
            continue
        score = _overlap(target, row_tokens[i])
        if score <= 0:
            continue
        score += (0.05 if hint_page and pg == hint_page else 0) + (0.02 if is_tbl else 0)
        if score > best_score:
            best, best_score = i, score
    if best is None or best_score < 0.6:
        return None, None
    return rows[best][0], rows[best][1][:700]


# ----------------------------------------------------------------------------
# 4. LLM JSON -> editable tables
# ----------------------------------------------------------------------------

ACT_COLS = ["row_id", "name", "function", "amount", "unit", "amount_basis", "nrv_percent", "min_amount", "max_amount",
            "page", "source_text", "needs_review", "review_reason"]
SUB_COLS = ["row_id", "parent", "name", "amount", "unit", "standardization", "min_amount", "max_amount", "page", "source_text"]
EXC_COLS = ["row_id", "name", "function", "amount", "unit", "amount_basis", "page", "source_text", "needs_review", "review_reason"]
MICRO_COLS = ["row_id", "name", "limit"]
READONLY = ["row_id", "page", "source_text", "needs_review", "review_reason"]

PRODUCT_GROUPS = {
    "Core (COC + FPS)": [("name", "Product name (as in TDS)"), ("display_name", "Name used in documents"),
                         ("quantity", "Quantity per package"), ("dosage_form", "Dosage form (singular)"),
                         ("unit_weight_value", "Dosage form weight"), ("unit_weight_unit", "Weight unit"),
                         ("tolerance", "Weight tolerance (%)")],
    "NRV (COC)": [("nrv_basis", "NRV basis: per_daily_dose or per_unit"), ("nrv_note", "NRV note found in the TDS")],
    "FPS - general": [("denomination", "Denomination"), ("format", "Format"), ("net_content", "Net content"),
                      ("net_content_labeling", "Net content for labeling"), ("shelf_life", "Shelf life (e.g. 3 years)"),
                      ("storage", "Storage information"), ("daily_dose", "Daily dose (number of units)")],
    "FPS - primary packaging": [("pkg_type", "Container type"), ("pkg_material", "Container material"),
                                ("pkg_color", "Container color"), ("pkg_dimensions", "Dimensions / capacity"),
                                ("closure_type", "Closure type"), ("closure_material", "Closure material"),
                                ("closure_color", "Closure color"), ("protective_filling", "Protective filling")],
    "FPS - secondary packaging": [("label_material", "Label material"), ("label_dimensions", "Label dimensions"),
                                  ("carton_material", "Carton material"), ("carton_dimensions", "Carton dimensions"),
                                  ("carton_design", "Carton design"), ("leaflet", "Leaflet (description or 'Yes')")],
    "FPS - organoleptic": [("org_description", "Product description"), ("org_size", "Product size"),
                           ("org_color", "Final product color"), ("org_odour", "Odour"), ("org_taste", "Taste")],
    "FPS - heavy metals (only if printed in TDS)": [("lead", "Lead"), ("cadmium", "Cadmium"), ("mercury", "Mercury")],
}
PRODUCT_KEYS = [k for g in PRODUCT_GROUPS.values() for k, _ in g]


def clean_product_name(raw, strip_formula=True):
    """'CANNON FORMULA 0' -> 'CANNON'; 'NOVARIUS_sachets' -> 'NOVARIUS' (stray '0' cells and packaging words removed)."""
    s = clean(raw).replace("_", " ")
    for _ in range(4):
        before = s
        s = " ".join(t for t in s.split() if not re.fullmatch(r"0+", t))  # a lone 0 is never part of a name
        s = re.sub(r"\s+(sachets?|capsules?|caps|tablets?|softgels?)$", "", s, flags=re.I)
        if strip_formula:
            s = re.sub(r"\s+formula$", "", s, flags=re.I)
        s = s.strip()
        if s == before:
            break
    return s


def _g(d, key):
    return (d or {}).get(key)


def to_frames(data, all_pages, strip_formula=True):
    p = data.get("product") or {}
    pr, sec, org, hm = p.get("primary"), p.get("secondary"), p.get("organoleptic"), p.get("heavy_metals")
    product = {
        "name": p.get("name"), "display_name": clean_product_name(p.get("name"), strip_formula),
        "quantity": p.get("quantity_per_package"), "dosage_form": p.get("dosage_form"),
        "unit_weight_value": p.get("unit_weight_value"), "unit_weight_unit": p.get("unit_weight_unit"),
        "tolerance": p.get("weight_tolerance_percent"),
        "denomination": p.get("denomination"), "format": p.get("format"), "net_content": p.get("net_content"),
        "net_content_labeling": p.get("net_content_labeling"), "shelf_life": p.get("shelf_life"),
        "storage": p.get("storage"), "daily_dose": p.get("daily_dose_units"),
        "nrv_basis": p.get("nrv_basis"), "nrv_note": p.get("nrv_note"),
        "pkg_type": _g(pr, "type"), "pkg_material": _g(pr, "material"), "pkg_color": _g(pr, "color"),
        "pkg_dimensions": _g(pr, "dimensions"), "closure_type": _g(pr, "closure_type"),
        "closure_material": _g(pr, "closure_material"), "closure_color": _g(pr, "closure_color"),
        "protective_filling": _g(pr, "protective_filling"),
        "label_material": _g(sec, "label_material"), "label_dimensions": _g(sec, "label_dimensions"),
        "carton_material": _g(sec, "carton_material"), "carton_dimensions": _g(sec, "carton_dimensions"),
        "carton_design": _g(sec, "carton_design"), "leaflet": _g(sec, "leaflet"),
        "org_description": _g(org, "description"), "org_size": _g(org, "size"), "org_color": _g(org, "color"),
        "org_odour": _g(org, "odour"), "org_taste": _g(org, "taste"),
        "lead": _g(hm, "lead"), "cadmium": _g(hm, "cadmium"), "mercury": _g(hm, "mercury"),
    }
    product = {k: clean(product.get(k)) for k in PRODUCT_KEYS}

    rows = build_rows(all_pages)
    row_tokens = [_tokens(r[1]) for r in rows]

    def ev(name, hint):
        return find_evidence(name, hint, rows, row_tokens)

    acts, subs, excs, micro = [], [], [], []
    for i, a in enumerate(data.get("actives") or [], start=1):
        pg, txt = ev(a.get("name"), a.get("page"))
        acts.append({"row_id": f"A{i}", "name": a.get("name"), "function": a.get("function"), "amount": a.get("amount"),
                     "unit": a.get("unit"), "amount_basis": a.get("basis") or "unknown", "nrv_percent": a.get("nrv"),
                     "min_amount": a.get("min"), "max_amount": a.get("max"), "page": pg, "source_text": txt,
                     "needs_review": bool(a.get("review")), "review_reason": a.get("review")})
        for j, s in enumerate(a.get("subs") or [], start=1):
            spg, stxt = ev(s.get("name"), pg or a.get("page"))
            subs.append({"row_id": f"S{i}.{j}", "parent": a.get("name"), "name": s.get("name"), "amount": s.get("amount"),
                         "unit": s.get("unit"), "standardization": s.get("std"), "min_amount": s.get("min"),
                         "max_amount": s.get("max"), "page": spg, "source_text": stxt})
    for i, e in enumerate(data.get("excipients") or [], start=1):
        pg, txt = ev(e.get("name"), e.get("page"))
        excs.append({"row_id": f"E{i}", "name": e.get("name"), "function": e.get("function"), "amount": e.get("amount"),
                     "unit": e.get("unit"), "amount_basis": e.get("basis") or "unknown", "page": pg, "source_text": txt,
                     "needs_review": bool(e.get("review")), "review_reason": e.get("review")})
    for i, m in enumerate(data.get("micro") or [], start=1):
        micro.append({"row_id": f"M{i}", "name": m.get("name"), "limit": m.get("limit")})

    frames = {"actives": pd.DataFrame(acts, columns=ACT_COLS), "subs": pd.DataFrame(subs, columns=SUB_COLS),
              "excipients": pd.DataFrame(excs, columns=EXC_COLS), "micro": pd.DataFrame(micro, columns=MICRO_COLS)}
    conflicts = [{"field": c.get("field"), "description": c.get("note")} for c in (data.get("conflicts") or [])]
    return product, frames, conflicts


def find_reg_footnote(pages):
    """The '**Regulation (EU) 1169/2011 ... Minerals 80-145%).' sentence, copied verbatim from the TDS (English part)."""
    full = " ".join(" ".join(p["text"].split()) for p in pages)
    m = re.search(r"\*\*\s*Regulation \(EU\) 1169/2011.*?Minerals\s*80-145%\)\.?", full)
    return m.group(0).strip() if m else None
