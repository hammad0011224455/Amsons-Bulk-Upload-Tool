#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Shopify Bulk Upload — Post-Processor (FIXED + TEMPLATE MAKER)
- Fixes accidental main() override that caused no outputs to be written.
- Writes Shopify CSV + validation report + input-with-SKUs + image report + title match report.
- NEW: --make-template writes an Excel/CSV input template that INCLUDES Barcode + Weight columns.
"""

import os
import re, csv, sys, argparse
from pathlib import Path
from typing import List, Dict, Tuple
import itertools
import pandas as pd

def build_barcode_pipe_for_sizes(row, o1vals):
    import pandas as pd, re
    base = row.get("Variant Barcode (EAN/UPC)", "")
    # If user provided a pipe string (or single barcode), use it and normalise to match size count.
    if not pd.isna(base) and str(base).strip():
        parts = [p.strip() for p in str(base).split("|")]
        if o1vals:
            n = len(o1vals)
            if len(parts) < n:
                parts += [""] * (n - len(parts))
            elif len(parts) > n:
                parts = parts[:n]
        return "|".join(parts)
    # If no sizes or no base barcode, fall back to per-size columns.
    if not o1vals:
        return ""
    per_size = {}
    for col in row.index:
        col_str = str(col).strip()
        col_lower = col_str.lower()
        if not col_lower.startswith("barcode"):
            continue
        if col_str in ("Variant Barcode (EAN/UPC)", "Variant Barcode"):
            continue
        val = row.get(col)
        if pd.isna(val) or str(val).strip() == "":
            continue
        parts = re.split(r"[ _-]+", col_str, 1)
        if len(parts) != 2:
            continue
        suffix = parts[1].strip()
        if not suffix:
            continue
        per_size[suffix.lower()] = str(val).strip()
    if not per_size:
        return ""
    barcodes = []
    for size in o1vals:
        key = str(size).strip().lower()
        barcodes.append(per_size.get(key, ""))
    return "|".join(barcodes)


def build_weight_pipe_for_sizes(row, o1vals):
    import pandas as pd, re
    base = row.get("Variant Weight", "")
    # If user provided a pipe string (or single weight), use it and normalise to match size count.
    if not pd.isna(base) and str(base).strip():
        parts = [p.strip() for p in str(base).split("|")]
        if o1vals:
            n = len(o1vals)
            if len(parts) < n:
                parts += [""] * (n - len(parts))
            elif len(parts) > n:
                parts = parts[:n]
        return "|".join(parts)
    # If no sizes or no base weight, fall back to per-size columns like "Weight 50".
    if not o1vals:
        return ""
    per_size = {}
    for col in row.index:
        col_str = str(col).strip()
        col_lower = col_str.lower()
        if not col_lower.startswith("weight"):
            continue
        if col_str in ("Variant Weight", "Variant Weight Unit (g,kg,lb,oz)"):
            continue
        val = row.get(col)
        if pd.isna(val) or str(val).strip() == "":
            continue
        parts = re.split(r"[ _-]+", col_str, 1)
        if len(parts) != 2:
            continue
        suffix = parts[1].strip()
        if not suffix:
            continue
        per_size[suffix.lower()] = str(val).strip()
    if not per_size:
        return ""
    weights = []
    for size in o1vals:
        key = str(size).strip().lower()
        weights.append(per_size.get(key, ""))
    return "|".join(weights)


def build_inventory_pipe_for_sizes(row, o1vals):
    """Build Variant Inventory pipe list from either:
    - 'Variant Inventory' (single value or pipe list)
    - per-size columns like 'Inventory 50', 'Inventory_52', etc.
    Values are expected to be 0/1 or a quantity.
    """
    import pandas as pd, re
    base = row.get("Variant Inventory", "")
    if not pd.isna(base) and str(base).strip():
        parts = [p.strip() for p in str(base).split("|")]
        if o1vals:
            n = len(o1vals)
            if len(parts) < n:
                parts += [""] * (n - len(parts))
            elif len(parts) > n:
                parts = parts[:n]
        return "|".join(parts)

    if not o1vals:
        return ""

    per_size = {}
    for col in row.index:
        col_str = str(col).strip()
        col_lower = col_str.lower()
        if not col_lower.startswith("inventory"):
            continue
        if col_str in ("Variant Inventory",):
            continue
        val = row.get(col)
        if pd.isna(val) or str(val).strip() == "":
            continue
        parts = re.split(r"[ _-]+", col_str, 1)
        if len(parts) != 2:
            continue
        suffix = parts[1].strip()
        if not suffix:
            continue
        per_size[suffix.lower()] = str(val).strip()

    if not per_size:
        return ""
    invs = []
    for size in o1vals:
        key = str(size).strip().lower()
        invs.append(per_size.get(key, ""))
    return "|".join(invs)


def build_grams_pipe_for_sizes(row, o1vals):
    """Build Variant Grams pipe list from either:
    - 'Variant Grams' (single value or pipe list)
    - per-size columns like 'Grams 50', 'Grams_52', etc.
    """
    import pandas as pd, re
    base = row.get("Variant Grams", "")
    if not pd.isna(base) and str(base).strip():
        parts = [p.strip() for p in str(base).split("|")]
        if o1vals:
            n = len(o1vals)
            if len(parts) < n:
                parts += [""] * (n - len(parts))
            elif len(parts) > n:
                parts = parts[:n]
        return "|".join(parts)

    if not o1vals:
        return ""

    per_size = {}
    for col in row.index:
        col_str = str(col).strip()
        col_lower = col_str.lower()
        if not col_lower.startswith("grams"):
            continue
        if col_str in ("Variant Grams",):
            continue
        val = row.get(col)
        if pd.isna(val) or str(val).strip() == "":
            continue
        parts = re.split(r"[ _-]+", col_str, 1)
        if len(parts) != 2:
            continue
        suffix = parts[1].strip()
        if not suffix:
            continue
        per_size[suffix.lower()] = str(val).strip()

    if not per_size:
        return ""
    grams = []
    for size in o1vals:
        key = str(size).strip().lower()
        grams.append(per_size.get(key, ""))
    return "|".join(grams)



# ---------------- Optional dependencies ----------------
try:
    import requests  # for real image checks
except Exception:
    requests = None

# Correct slugify import (function), with safe fallback
try:
    from slugify import slugify as _slugify
except Exception:
    _slugify = None

# ---- Console UTF-8 safety on Windows ----
def _force_utf8_stdio():
    """
    Ensure prints don't crash with UnicodeEncodeError on Windows consoles.
    """
    try:
        # Python 3.7+ supports reconfigure
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        # Fallback: wrap buffers
        import io
        if hasattr(sys.stdout, "buffer"):
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        if hasattr(sys.stderr, "buffer"):
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')



def slugify_str(s: str) -> str:
    if _slugify:
        return _slugify(s)[:255]
    import re as _re
    s = s.lower()
    s = _re.sub(r"[^a-z0-9]+", "-", s)
    s = _re.sub(r"-+", "-", s).strip("-")
    return s[:255]

# ---------------- Shopify CSV columns ------------------
SHOPIFY_HEADERS = [
    "Handle","Title","Body (HTML)","Vendor","Type","Tags","Published",
    "Option1 Name","Option1 Value","Option2 Name","Option2 Value","Option3 Name","Option3 Value",
    "Variant SKU","Variant Grams","Variant Inventory Tracker","Variant Inventory Qty","Variant Inventory Policy",
    "Variant Fulfillment Service","Variant Price","Variant Compare At Price","Variant Requires Shipping","Variant Taxable","Variant Barcode",
    "Image Src","Image Position","Image Alt Text",
    "Gift Card","SEO Title","SEO Description","Status","Variant Weight Unit"
]

# Inventory export (Shopify) column order (matches Shopify "Inventory export" CSV)
INVENTORY_EXPORT_HEADERS = [
    "Handle","Title",
    "Option1 Name","Option1 Value","Option2 Name","Option2 Value","Option3 Name","Option3 Value",
    "SKU","HS Code","COO","Location","Bin name",
    "Incoming (not editable)","Unavailable (not editable)","Committed (not editable)",
    "Available (not editable)","On hand (current)","On hand (new)"
]

# Default locations used for Amsons inventory export output (first location is treated as the primary stocking location).
DEFAULT_INVENTORY_LOCATIONS = [
    "Amsons Birmingham - Small Heath",
    "Amsons Bradford",
    "Amsons Birmingham - Alum Rock",
]

def _to_int_safe(x, default=0):
    try:
        if x is None: return default
        s = str(x).strip()
        if s == "": return default
        # Handle "1.0" from Excel
        return int(float(s))
    except Exception:
        return default

def build_shopify_inventory_export_rows(shopify_rows: list, locations=None, in_stock_qty: int = 1000) -> list:
    """
    Build a Shopify Inventory Export-style sheet from the generated Shopify import rows.

    Rules (per user request):
    - Uses the same column sequence as Shopify inventory export.
    - Reads "Variant Inventory Qty" (0/1 or quantity) from the Shopify import row.
    * If qty > 0 -> sets primary location On hand (new) = in_stock_qty (default 100). Available + On hand(current) are forced to 0 for fresh import.
      * If qty <= 0 -> sets primary location Available/On hand(current) = 0.
    - Additional locations are output as "not stocked" across inventory columns (matches Shopify export).
    """
    locs = locations or DEFAULT_INVENTORY_LOCATIONS
    if not locs:
        locs = ["Default"]
    out = []
    for r in shopify_rows:
        sku = str(r.get("Variant SKU","") or "").strip()
        # Skip non-variant rows (e.g., image-only lines).
        if not sku:
            continue

        qty_raw = r.get("Variant Inventory Qty", "")
        qty_in = _to_int_safe(qty_raw, default=0)
        qty_primary = in_stock_qty if qty_in > 0 else 0

        base = {
            "Handle": r.get("Handle",""),
            "Title": r.get("Title",""),
            "Option1 Name": r.get("Option1 Name",""),
            "Option1 Value": r.get("Option1 Value",""),
            "Option2 Name": r.get("Option2 Name",""),
            "Option2 Value": r.get("Option2 Value",""),
            "Option3 Name": r.get("Option3 Name",""),
            "Option3 Value": r.get("Option3 Value",""),
            "SKU": sku,
            "HS Code": r.get("Variant HS Code","") or r.get("HS Code","") or "",
            "COO": r.get("Variant Country of Origin","") or r.get("COO","") or "",
            "Bin name": "",
            "On hand (new)": "",
        }

        # Primary stocking location row
        primary_row = dict(base)
        primary_row.update({
            "Location": locs[0],
            "Incoming (not editable)": 0,
            "Unavailable (not editable)": 0,
            "Committed (not editable)": 0,
            "Available (not editable)": 0,
            "On hand (current)": 0,
            "On hand (new)": qty_primary,
        })
        out.append(primary_row)

        # Other locations as not stocked
        for loc in locs[1:]:
            nr = dict(base)
            nr.update({
                "Location": loc,
                "Incoming (not editable)": "not stocked",
                "Unavailable (not editable)": "not stocked",
                "Committed (not editable)": "not stocked",
                "Available (not editable)": 0,
                "On hand (current)": 0,
                "On hand (new)": "not stocked",
            })
            out.append(nr)
    return out

STATUS_VALUES = {"active","draft","archived"}
WEIGHT_UNITS = {"g","kg","lb","oz"}

# ---------------- Regex & parsing helpers ----------------
STRICT_SKU_RE = re.compile(r"^(?P<base>\d{6})(?:-(?P<idx>\d{2}))?$")

def _clean_sku_text(s: str) -> str:
    """
    Strip Excel-style leading apostrophes and whitespace.
    Examples:
      '110374  -> 110374
      ’110374  -> 110374
    """
    return str(s or "").strip().lstrip("'").lstrip("’").strip()

def extract_base_6(s: str):
    """Return 6-digit base int if SKU matches ^\\d{6}(-\\d{2})?$, else None."""
    s2 = _clean_sku_text(s)
    m = STRICT_SKU_RE.match(s2)
    if not m:
        return None
    return int(m.group("base"))

def split_pipe(val) -> List[str]:
    """Split a cell on |; return [] for blank/NaN."""
    if pd.isna(val) or str(val).strip() == "":
        return []
    s = str(val)
    parts = [p.strip() for p in s.split("|")]
    return [p for p in parts if p != ""]

def coerce_bool_token(x):
    if pd.isna(x): return ""
    s = str(x).strip().lower()
    if s in {"true","t","yes","y","1"}: return "TRUE"
    if s in {"false","f","no","n","0"}: return "FALSE"
    return ""

def grams_from_weight(val, unit):
    try:
        w = float(val)
    except Exception:
        return ""
    u = (unit or "").strip().lower()
    if u == "kg": return int(round(w * 1000))
    if u == "g":  return int(round(w))
    if u == "lb": return int(round(w * 453.59237))
    if u == "oz": return int(round(w * 28.3495231))
    return ""

def is_url(s: str) -> bool:
    return bool(re.match(r"^https?://", str(s or "").strip(), re.I))

def looks_like_image_url(s: str) -> bool:
    return bool(re.search(r"\.(jpg|jpeg|png|gif|webp|tiff?)($|\?)", str(s or "").lower()))

def check_image_url(url: str, timeout=10):
    if not url or not is_url(url):
        return False, "Not a URL"
    if requests is None:
        return (looks_like_image_url(url), "requests not installed; only extension check")
    try:
        resp = requests.head(url, allow_redirects=True, timeout=timeout)
        if resp.status_code == 405:
            resp = requests.get(url, stream=True, timeout=timeout)
        if resp.status_code != 200:
            return False, f"HTTP {resp.status_code}"
        ctype = (resp.headers.get("Content-Type") or "").lower()
        if "image" not in ctype:
            return False, f"Content-Type '{ctype}' not image"
        return True, "OK"
    except Exception as e:
        return False, f"Error: {e}"

def uniqueness_suffix(existing: set, base: str) -> str:
    if base not in existing:
        existing.add(base); return base
    i = 1
    while True:
        cand = f"{base}-{i}"
        if cand not in existing:
            existing.add(cand); return cand
        i += 1

# ---------------- Broadcasting helpers for per-variant fields ----------
def broadcast_values(value_str, n1, n2, n3, field, rownum, issues: List[dict]) -> List[str]:
    """
    Broadcast a pipe list to the full variant grid (n1 x n2 x n3) in the order:
       product(o1vals, o2vals, o3vals)  -> o1 major, then o2, then o3.
    Supports lengths: 0, 1, n3, n2, n1, n1*n2, n1*n3, n2*n3, n1*n2*n3.
    Otherwise warns and repeats/truncates to fit.
    """
    vals = split_pipe(value_str)
    total = n1 * n2 * n3
    if len(vals) == 0:
        return [""] * total
    if len(vals) == 1:
        return vals * total
    if len(vals) == total:
        return vals

    out = []
    if len(vals) == n1:
        for i in range(n1):
            for j in range(n2):
                for k in range(n3):
                    out.append(vals[i])
        return out
    if len(vals) == n2:
        for i in range(n1):
            for j in range(n2):
                for k in range(n3):
                    out.append(vals[j])
        return out
    if len(vals) == n3:
        for i in range(n1):
            for j in range(n2):
                for k in range(n3):
                    out.append(vals[k])
        return out
    if len(vals) == n1 * n2 and n3 > 1:
        idx = 0
        for i in range(n1):
            for j in range(n2):
                v = vals[idx]; idx += 1
                for k in range(n3):
                    out.append(v)
        return out
    if len(vals) == n1 * n3 and n2 > 1:
        out2 = []
        for i in range(n1):
            for j in range(n2):
                for k in range(n3):
                    out2.append(vals[i*n3 + k])
        return out2
    if len(vals) == n2 * n3 and n1 > 1:
        out = []
        for i in range(n1):
            out.extend(vals)
        return out

    issues.append({"level":"warning","row":rownum,"field":field,
                   "message":f"Count mismatch for broadcasting: have {len(vals)}, expected 1, {n1}, {n2}, {n3}, {n1*n2}, {n1*n3}, {n2*n3}, or {total}. Repeating/truncating."})
    if len(vals) < total:
        return vals + [vals[-1]] * (total - len(vals))
    else:
        return vals[:total]

# ---------------- Previous export (row 2 priority) ----------------
def gather_used_bases(series) -> set:
    used = set()
    if series is None:
        return used
    try:
        import pandas as _pd
        if isinstance(series, _pd.Series):
            iterable = series.astype(str).tolist()
        else:
            iterable = list(series)
    except Exception:
        iterable = list(series)
    for sku in iterable:
        b = extract_base_6(sku)  # extract_base_6 cleans apostrophes internally
        if b is not None:
            used.add(b)
    return used

def load_prev_highest_base(prev_path: Path) -> int:
    """
    Return highest base from the previous export:
      1) Prefer the first data cell (Excel row 2) in 'Variant SKU' column (cleaned).
      2) Else fallback to max base across the column.
    """
    if not prev_path or not prev_path.exists():
        return 0
    if prev_path.suffix.lower() in {".xlsx", ".xls"}:
        pdf = pd.read_excel(prev_path, dtype=str)
    else:
        pdf = pd.read_csv(prev_path, dtype=str)
    pdf = pdf.fillna("")
    if "Variant SKU" not in pdf.columns:
        return 0
    col = pdf["Variant SKU"].astype(str).str.strip()

    # Row-2 priority (first non-empty cell)
    top_cell = None
    for val in col.tolist():
        val = _clean_sku_text(val)  # clean apostrophes/spaces
        if val and val.lower() != "nan":
            top_cell = val
            break
    if top_cell:
        b = extract_base_6(top_cell)
        if b is not None:
            return b

    # Fallback: scan entire column, cleaned
    bases_prev = set()
    for v in col.tolist():
        v2 = _clean_sku_text(v)
        b = extract_base_6(v2)
        if b is not None:
            bases_prev.add(b)
    return max(bases_prev) if bases_prev else 0

# ---------------- Core build ---------------------------
def build_shopify_rows(df: pd.DataFrame, highest_prev_base: int, respect_existing: bool):
    """
    One row per product in df. Generates variants by exploding pipe lists.
    Returns rows, issues, highest_before, highest_after, image_results, df_with_skus
    """
    issues, rows, image_results = [], [], []
    df = df.fillna("")
    used_handles = set()

    # Allow override via env var set by the GUI (optional)
    env_base = os.getenv("AMS_START_BASE")
    if env_base and str(env_base).isdigit():
        highest_prev_base = int(env_base)

    highest_before = highest_prev_base
    # Start AT the previous highest (so first new product -> 110374-01, next product -> 110375-01, etc.)
    next_base = (highest_prev_base + 1) if highest_prev_base else 100001  # start AFTER the highest found

    if "Variant SKU" not in df.columns:
        df["Variant SKU"] = ""

    for i, r in df.iterrows():
        excel_row = i + 2
        title = str(r.get("Title*","")).strip()
        vendor = str(r.get("Vendor*","")).strip()
        if not title:
            issues.append({"level":"error","row":excel_row,"field":"Title*","message":"Empty title"})
        if not vendor:
            issues.append({"level":"error","row":excel_row,"field":"Vendor*","message":"Empty vendor"})

        handle_src = str(r.get("Handle (optional)","")).strip() or title or f"product-{excel_row}"
        handle = uniqueness_suffix(used_handles, slugify_str(handle_src))

        body = r.get("Body (HTML)","")
        ptype = r.get("Type (Product Type)","")
        tags = r.get("Tags (comma-separated)","")
        published = coerce_bool_token(r.get("Published (TRUE/FALSE)","")) or "TRUE"
        status = (str(r.get("Status (active/draft/archived)","")).strip().lower() or "active")
        if status not in STATUS_VALUES:
            issues.append({"level":"warning","row":excel_row,"field":"Status","message":f"Unknown status '{status}', defaulting to active"})
            status = "active"
        seo_title = r.get("SEO Title","")
        seo_desc = r.get("SEO Description","")

        o1n = (str(r.get("Option1 Name","")).strip() or "Title")
        o1vals = split_pipe(r.get("Option1 Values",""))
        if o1n.lower() == "title" and not o1vals:
            o1vals = ["Default Title"]
        if not o1vals:
            o1vals = ["Default Title"]

        o2n = str(r.get("Option2 Name","")).strip()
        o2vals = split_pipe(r.get("Option2 Values",""))
        o3n = str(r.get("Option3 Name","")).strip()
        o3vals = split_pipe(r.get("Option3 Values",""))

        n1 = len(o1vals)
        o2l = o2vals if (o2n and o2vals) else [""]
        o3l = o3vals if (o3n and o3vals) else [""]
        n2, n3 = len(o2l), len(o3l)
        combos = list(itertools.product(o1vals, o2l, o3l))
        nvars = len(combos)
        if nvars > 300:
            issues.append({"level":"error","row":excel_row,"field":"Options","message":f"Too many variants ({nvars}). Please reduce combinations."})
            continue

        vprice_list = broadcast_values(r.get("Variant Price*",""), n1, n2, n3, "Variant Price*", excel_row, issues)
        vcmp_list   = broadcast_values(r.get("Variant Compare At Price",""), n1, n2, n3, "Variant Compare At Price", excel_row, issues)

        # Inventory input:
        # - If no variants, user enters a single value (0 or 1; 0=out of stock, 1=in stock)
        # - If variants, user enters pipe list (e.g. 0|1|1|0)
        inv_value_str = build_inventory_pipe_for_sizes(r, o1vals) or str(r.get("Variant Inventory", "")).strip() or "1"
        vinv_list_raw = broadcast_values(inv_value_str, n1, n2, n3, "Variant Inventory", excel_row, issues)

        def _inv_to_qty(tok: str) -> str:
            s = str(tok or "").strip().lower()
            if s in {"", "1", "in", "instock", "in stock", "true", "yes"}:
                return "1000"  # in stock
            if s in {"0", "out", "oos", "outofstock", "out of stock", "false", "no"}:
                return "0"     # out of stock
            # Allow explicit quantities if user ever provides them
            if s.isdigit():
                return s
            return "1000"

        vqty_list = [_inv_to_qty(x) for x in vinv_list_raw]
        barcode_value_str = build_barcode_pipe_for_sizes(r, o1vals)
        vbar_list   = broadcast_values(barcode_value_str, n1, n2, n3, "Variant Barcode", excel_row, issues)
        # Grams input (preferred). If blank, we will compute grams from Weight + Unit later.
        grams_value_str = build_grams_pipe_for_sizes(r, o1vals)
        vgrams_list = broadcast_values(grams_value_str, n1, n2, n3, "Variant Grams", excel_row, issues) if grams_value_str else ["" for _ in range(nvars)]

        weight_value_str = build_weight_pipe_for_sizes(r, o1vals)
        vwt_list    = broadcast_values(weight_value_str, n1, n2, n3, "Variant Weight", excel_row, issues)
        vunit_list  = broadcast_values(r.get("Variant Weight Unit (g,kg,lb,oz)",""), n1, n2, n3, "Variant Weight Unit", excel_row, issues)
        vship_list  = [coerce_bool_token(x) or "TRUE" for x in broadcast_values(r.get("Variant Requires Shipping (TRUE/FALSE)",""), n1, n2, n3, "Variant Requires Shipping", excel_row, issues)]
        vtax_list   = [coerce_bool_token(x) or "TRUE" for x in broadcast_values(r.get("Variant Taxable (TRUE/FALSE)",""), n1, n2, n3, "Variant Taxable", excel_row, issues)]

        image_pairs = []
        for n in range(1, 9):
            u = r.get(f"Image URL {n}","")
            a = r.get(f"Image Alt {n}","")
            if u:
                image_pairs.append((u, a))

        existing_skus = split_pipe(r.get("Variant SKU",""))
        assigned_skus: List[str] = []
        if respect_existing and any(existing_skus):
            ex = broadcast_values(r.get("Variant SKU",""), n1, n2, n3, "Variant SKU", excel_row, issues)
            assigned_skus = ex[:]
            need_fill = [idx for idx, s in enumerate(assigned_skus) if not s]
            if need_fill:
                base_str = f"{next_base:06d}"
                if nvars == 1:
                    assigned_skus[0] = base_str
                else:
                    counter = 1
                    for idxv in need_fill:
                        assigned_skus[idxv] = f"{base_str}-{counter:02d}"
                        counter += 1
                next_base += 1
        else:
            base_str = f"{next_base:06d}"
            if nvars == 1:
                assigned_skus = [base_str]
            else:
                assigned_skus = [f"{base_str}-{j:02d}" for j in range(1, nvars+1)]
            next_base += 1

        df.at[i, "Variant SKU"] = "|".join(assigned_skus)

        for idx_img, (u, a) in enumerate(image_pairs, start=1):
            ok, note = check_image_url(u)
            image_results.append({
                "handle": handle,
                "position": idx_img,
                "url": u,
                "ok": bool(ok),
                "note": note
            })

        is_first_row_for_product = True
        extra_image_rows = []
        if image_pairs:
            pos = 2
            for (uu, aa) in image_pairs[1:]:
                rimg = {k:"" for k in SHOPIFY_HEADERS}
                rimg["Handle"] = handle
                rimg["Image Src"] = uu
                rimg["Image Position"] = str(pos)
                rimg["Image Alt Text"] = aa
                extra_image_rows.append(rimg)
                pos += 1

        for idxv, (opt1, opt2, opt3) in enumerate(combos):
            vprice = vprice_list[idxv]
            if not vprice:
                issues.append({"level":"error","row":excel_row,"field":"Variant Price*","message":f"Empty price for variant #{idxv+1}"})
            vcmp   = vcmp_list[idxv]
            vqty   = vqty_list[idxv]
            vbar   = vbar_list[idxv]
            vwt    = vwt_list[idxv]
            vunit  = (vunit_list[idxv] or "").lower()
            vship  = vship_list[idxv]
            vtax   = vtax_list[idxv]
            # Use explicit grams if provided, otherwise compute from Weight + Unit
            vgrams_in = str(vgrams_list[idxv] or "").strip()
            if vgrams_in:
                vgrams = vgrams_in
            else:
                vgrams = grams_from_weight(vwt, vunit) if vwt and vunit in WEIGHT_UNITS else ""
            vsku   = assigned_skus[idxv]

            base_row = {
                "Handle": handle,
                "Title": title,
                "Body (HTML)": body,
                "Vendor": vendor,
                "Type": ptype,
                "Tags": tags,
                "Published": published,
                "Option1 Name": o1n,
                "Option1 Value": opt1,
                "Option2 Name": o2n,
                "Option2 Value": opt2,
                "Option3 Name": o3n,
                "Option3 Value": opt3,
                "Variant SKU": vsku,
                "Variant Grams": vgrams,
                "Variant Inventory Tracker": "shopify",
                "Variant Inventory Qty": vqty,
                "Variant Inventory Policy": "deny",
                "Variant Fulfillment Service": "manual",
                "Variant Price": vprice,
                "Variant Compare At Price": vcmp,
                "Variant Requires Shipping": vship,
                "Variant Taxable": vtax,
                "Variant Barcode": vbar,
                "Gift Card": "FALSE",
                "SEO Title": seo_title,
                "SEO Description": seo_desc,
                "Status": status,
                "Variant Weight Unit": vunit if vunit else "",
            }

            if is_first_row_for_product and image_pairs:
                base_row["Image Src"] = image_pairs[0][0]
                base_row["Image Position"] = "1"
                base_row["Image Alt Text"] = image_pairs[0][1]
                is_first_row_for_product = False
                rows.append(base_row)
                rows.extend(extra_image_rows)
            else:
                rows.append(base_row)

    highest_after = next_base - 1 if next_base > 0 else highest_before
    return rows, issues, highest_before, highest_after, image_results, df

# ---------------- Excel writer helper ----------------
def get_excel_engine():
    """Return a usable Excel writer engine or None if neither is available."""
    try:
        import xlsxwriter  # noqa: F401
        return "xlsxwriter"
    except Exception:
        try:
            import openpyxl  # noqa: F401
            return "openpyxl"
        except Exception:
            return None

# ---------------- NEW: Image report writer (Excel) ----------------
def write_image_report_xlsx(image_results: List[dict], rows: List[dict], outdir: Path, engine: str) -> Path:
    """
    Create an Excel file 'image_report.xlsx' with columns:
      Title | Handle | Image Position | Image URL | Working | Note
    Title is derived from the product rows via Handle.
    Falls back to CSV if no Excel engine is available.
    """
    handle_to_title = {}
    for r in rows:
        try:
            h = r.get("Handle", "")
            t = r.get("Title", "")
        except Exception:
            h, t = "", ""
        if h and t and h not in handle_to_title:
            handle_to_title[h] = t

    records = []
    for im in image_results:
        title = handle_to_title.get(im.get("handle", ""), "")
        records.append({
            "Title": title,
            "Handle": im.get("handle", ""),
            "Image Position": im.get("position", ""),
            "Image URL": im.get("url", ""),
            "Working": "Working" if im.get("ok", False) else "Not Working",
            "Note": im.get("note", "")
        })

    df_img = pd.DataFrame(records, columns=["Title","Handle","Image Position","Image URL","Working","Note"])

    if engine:
        path = outdir / "image_report.xlsx"
        with pd.ExcelWriter(path, engine=engine) as writer:
            df_img.to_excel(writer, sheet_name="Images", index=False)
        return path
    else:
        path = outdir / "image_report.csv"
        df_img.to_csv(path, index=False, encoding="utf-8-sig")
        return path

# ---------------- NEW: Title match report ----------------
def _norm_title(s) -> str:
    return str(s).strip().lower() if pd.notna(s) else ""

def build_title_matches(prev_path: Path, df_input: pd.DataFrame) -> pd.DataFrame:
    """
    Compare titles between previous export (prev_path) and current input df.
    Returns a DataFrame with columns:
      Title | In Previous Count | In Input Count
    Only rows where Title appears in BOTH are included.
    """
    if not prev_path or not prev_path.exists():
        return pd.DataFrame(columns=["Title","In Previous Count","In Input Count"])

    # Load previous export
    if prev_path.suffix.lower() in {".xlsx", ".xls"}:
        pdf = pd.read_excel(prev_path, dtype=str)
    else:
        pdf = pd.read_csv(prev_path, dtype=str)
    pdf = pdf.fillna("")

    if "Title" not in pdf.columns:
        return pd.DataFrame(columns=["Title","In Previous Count","In Input Count"])

    prev_titles = pdf["Title"].astype(str)
    prev_norm = prev_titles.apply(_norm_title)
    prev_counts = prev_norm.value_counts()

    # Current input titles
    if "Title*" not in df_input.columns:
        return pd.DataFrame(columns=["Title","In Previous Count","In Input Count"])
    inp_titles = df_input["Title*"].astype(str)
    inp_norm = inp_titles.apply(_norm_title)
    inp_counts = inp_norm.value_counts()

    # Intersection
    common = set(prev_counts.index) & set(inp_counts.index)
    if not common:
        return pd.DataFrame(columns=["Title","In Previous Count","In Input Count"])

    # Use the first appearance in input df for nice casing of the Title
    norm_to_pretty = {}
    for t in df_input["Title*"].astype(str):
        n = _norm_title(t)
        if n and n not in norm_to_pretty:
            norm_to_pretty[n] = t

    rows = []
    for n in sorted(common):
        rows.append({
            "Title": norm_to_pretty.get(n, n),
            "In Previous Count": int(prev_counts.get(n, 0)),
            "In Input Count": int(inp_counts.get(n, 0)),
        })
    return pd.DataFrame(rows, columns=["Title","In Previous Count","In Input Count"])

def write_title_matches_xlsx(matches_df: pd.DataFrame, outdir: Path, engine: str) -> Path:
    """
    Write the matches DataFrame to 'title_matches.xlsx' (or CSV fallback).
    """
    if matches_df is None or matches_df.empty:
        # Still write an empty file to make it obvious
        if engine:
            path = outdir / "title_matches.xlsx"
            with pd.ExcelWriter(path, engine=engine) as writer:
                matches_df.to_excel(writer, sheet_name="Matches", index=False)
            return path
        else:
            path = outdir / "title_matches.csv"
            matches_df.to_csv(path, index=False, encoding="utf-8-sig")
            return path

    if engine:
        path = outdir / "title_matches.xlsx"
        with pd.ExcelWriter(path, engine=engine) as writer:
            matches_df.to_excel(writer, sheet_name="Matches", index=False)
        return path
    else:
        path = outdir / "title_matches.csv"
        matches_df.to_csv(path, index=False, encoding="utf-8-sig")
        return path

# ---------------- NEW: Template maker ----------------
TEMPLATE_COLUMNS = [
    "Handle (optional)",
    "Title*",
    "Body (HTML)",
    "Vendor*",
    "Type (Product Type)",
    "Tags (comma-separated)",
    "Published (TRUE/FALSE)",
    "Status (active/draft/archived)",
    "SEO Title",
    "SEO Description",

    # Options
    "Option1 Name",
    "Option1 Values",
    "Option2 Name",
    "Option2 Values",
    "Option3 Name",
    "Option3 Values",

    # Pricing & tax/shipping
    "Variant Price*",
    "Variant Compare At Price",
    "Variant Requires Shipping (TRUE/FALSE)",
    "Variant Taxable (TRUE/FALSE)",

    # Barcode, Grams & Inventory inputs in the template (read by the builder)
    "Variant Barcode (EAN/UPC)",
    "Variant Grams",
    "Variant Inventory",

    # Backwards-compat (older templates)
    "Variant Weight",
    "Variant Weight Unit (g,kg,lb,oz)",

    # Optional prefill SKUs (pipe list). If blank, script assigns.
    "Variant SKU",

    # Images (first is primary)
    "Image URL 1","Image Alt 1",
    "Image URL 2","Image Alt 2",
    "Image URL 3","Image Alt 3",
    "Image URL 4","Image Alt 4",
    "Image URL 5","Image Alt 5",
    "Image URL 6","Image Alt 6",
    "Image URL 7","Image Alt 7",
    "Image URL 8","Image Alt 8",
]

README_LINES = [
    "HOW TO USE:",
    "• Fill one row per product. Use '|' to separate variant values (e.g., Option1 Values: 50|52|54).",
    "• Columns marked * are required. Leave 'Variant SKU' blank to auto-assign.",
    "• Variant Barcode/Grams/Inventory support pipe-lists and auto-broadcasting across combinations.",
    "• 'Variant Inventory' uses 0=out of stock, 1=in stock (or you can enter quantities).",
    "• If 'Variant Grams' is blank, you may still use the older 'Variant Weight' + 'Variant Weight Unit' (g,kg,lb,oz) and the script will compute grams.",
    "• Image URL 1 is the main image; add up to 8 images per product.",
    "• Published/Taxable/Requires Shipping accept TRUE/FALSE (case-insensitive).",
]

def write_template_file(path: Path):
    df = pd.DataFrame(columns=TEMPLATE_COLUMNS)
    engine = get_excel_engine()

    # Always try to add a README sheet for Excel templates
    if engine and str(path).lower().endswith((".xlsx", ".xls")):
        with pd.ExcelWriter(path, engine=engine) as writer:
            df.to_excel(writer, sheet_name="Products", index=False)
            readme = pd.DataFrame({"Notes": README_LINES})
            readme.to_excel(writer, sheet_name="README", index=False)
        return path
    else:
        # CSV fallback (no README sheet possible)
        if path.suffix.lower() != ".csv":
            path = path.with_suffix(".csv")
        df.to_csv(path, index=False, encoding="utf-8-sig")
        return path

# ---------------- CLI -------------------------------
def main():
    _force_utf8_stdio()

    ap = argparse.ArgumentParser()
    # Hardcoded defaults (Windows paths)
    default_input = r"C:\Users\Ahmed Amsons\OneDrive - Amsons\Desktop\testt\Shopify_Easy_Template.xlsx"
    default_outdir = r"C:\Users\Ahmed Amsons\OneDrive - Amsons\Desktop\testt\out"
    default_prev = r"C:\Users\Ahmed Amsons\OneDrive - Amsons\Desktop\testt\lasttest_export.csv"

    ap.add_argument("--input",  default=default_input,  help="Path to the easy template Excel")
    ap.add_argument("--sheet",  default="Products",     help="Worksheet name")
    ap.add_argument("--outdir", default=default_outdir, help="Output directory")
    ap.add_argument("--prev",   default=default_prev,   help="Previous export (CSV/XLSX); uses row 2 of 'Variant SKU' as highest if valid")
    ap.add_argument("--respect-existing-skus", action="store_true",
                    help="Keep any existing SKUs (pipe-list) and only fill blanks; default overwrites all SKUs per product")

    # NEW: make a blank input template (with Barcode & Weight columns)
    ap.add_argument("--make-template", metavar="PATH", help="Write a fresh input template to PATH (.xlsx preferred). Then exit.")

    args = ap.parse_args()

    # --- NEW: Template creation mode ---
    if args.make_template:
        outp = Path(args.make_template)
        outp.parent.mkdir(parents=True, exist_ok=True)
        template_path = write_template_file(outp)
        print(f"Template written: {template_path}")
        return

    inp = Path(args.input)
    outdir = Path(args.outdir); outdir.mkdir(parents=True, exist_ok=True)
    if not inp.exists():
        print(f"ERROR: Input not found: {inp}", file=sys.stderr); sys.exit(2)

    # Highest base from previous export
    highest_prev_base = 0
    prev_path = Path(args.prev) if args.prev else None
    if prev_path:
        if prev_path.exists():
            highest_prev_base = load_prev_highest_base(prev_path)
        else:
            print(f"WARNING: --prev file not found: {prev_path}", file=sys.stderr)

    # Load input
    try:
        df = pd.read_excel(inp, sheet_name=args.sheet, dtype=str)
        # --- Normalise Shopify-style template columns ---
        cols = {c.strip(): c for c in df.columns}

        # If Shopify's Variant Barcode column exists, map it to our template name
        if "Variant Barcode (EAN/UPC)" not in df.columns and "Variant Barcode" in cols:
            df["Variant Barcode (EAN/UPC)"] = df[cols["Variant Barcode"]]

        # If Shopify's Variant Grams exists, convert into Variant Weight + Unit "g"
        if "Variant Weight" not in df.columns and "Variant Grams" in cols:
            grams_col = cols["Variant Grams"]
            df["Variant Weight"] = df[grams_col]
            df["Variant Weight Unit (g,kg,lb,oz)"] = "g"

    except Exception as e:
        print(f"ERROR: Could not read input Excel: {e}", file=sys.stderr)
        sys.exit(3)

    rows, issues, highest_before, highest_after, image_results, df_with_skus = build_shopify_rows(
        df, highest_prev_base, respect_existing=args.respect_existing_skus
    )

    # Shopify CSV
    out_csv = outdir / "shopify_import.csv"
    with out_csv.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=SHOPIFY_HEADERS, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k,"") for k in SHOPIFY_HEADERS})


    # Shopify Inventory Export-style CSV (for locations/stock sync)
    inv_rows = build_shopify_inventory_export_rows(rows)
    out_inv = outdir / "shopify_inventory_export.csv"
    with out_inv.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=INVENTORY_EXPORT_HEADERS, extrasaction="ignore")
        w.writeheader()
        for r in inv_rows:
            w.writerow({k: r.get(k,"") for k in INVENTORY_EXPORT_HEADERS})

    # Validation report
    rep_csv = outdir / "validation_report.csv"
    with rep_csv.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["level","row","field","message"])
        w.writeheader()
        for it in issues:
            w.writerow(it)

    # Input-with-SKUs copy (xlsx if possible; else CSV)
    engine = get_excel_engine()
    if engine:
        xlsx_out = outdir / "input_with_skus.xlsx"
        with pd.ExcelWriter(xlsx_out, engine=engine) as writer:
            df_with_skus.to_excel(writer, sheet_name=args.sheet, index=False)
        back_out = xlsx_out
    else:
        csv_out = outdir / "input_with_skus.csv"
        df_with_skus.to_csv(csv_out, index=False, encoding="utf-8-sig")
        back_out = csv_out
        print("NOTE: Neither 'xlsxwriter' nor 'openpyxl' is installed; wrote CSV instead of Excel.")

    # -------- Terminal output --------
    print("===== SKU SUMMARY =====")
    print(f"Highest 6-digit base (from row 2 under 'Variant SKU'): {highest_before or 'none'}")
    print(f"Highest 6-digit base AFTER assignment:               {highest_after or 'none'}")

    print("\n===== IMAGE CHECKS =====")
    if requests is None:
        print("NOTE: 'requests' not installed; image checks limited to URL/extension pattern.\n")

    total = len(image_results)
    working = sum(1 for x in image_results if x["ok"])
    broken = total - working
    for r in image_results:
        status = "OK" if r["ok"] else "BROKEN"
        print(f"[{status}] Handle='{r['handle']}' Pos={r['position']} URL={r['url']} ({r['note']})")

    print("\n----- IMAGE SUMMARY -----")
    print(f"Total images: {total}")
    print(f"Working:      {working}")
    print(f"Broken:       {broken}")
    if broken:
        print("\nBroken image URLs:")
        for r in image_results:
            if not r["ok"]:
                print(f"- Handle='{r['handle']}' Pos={r['position']} URL={r['url']}  Reason: {r['note']}")

    # -------- NEW: Write the Excel image report --------
    img_report_path = write_image_report_xlsx(image_results, rows, outdir, engine)
    print(f"\n- Image report     : {img_report_path}")

    # -------- NEW: Title matches report (prev vs input) --------
    matches_df = build_title_matches(prev_path if prev_path and prev_path.exists() else None, df)
    match_report_path = write_title_matches_xlsx(matches_df, outdir, engine)
    if matches_df.empty:
        print(f"- Title matches    : {match_report_path} (no matches found)")
    else:
        print(f"- Title matches    : {match_report_path} ({len(matches_df)} matched title(s))")

    print("\n===== OUTPUT FILES =====")
    print(f"- Shopify CSV       : {out_csv}")
    print(f"- Shopify inventory : {out_inv}")
    print(f"- Validation report : {rep_csv}")
    print(f"- Input with SKUs   : {back_out}")
    print(f"- Image report      : {img_report_path}")
    print(f"- Title matches     : {match_report_path}")
    print(f"Writing to: {outdir.resolve()}")

if __name__ == "__main__":
    _force_utf8_stdio()
    main()

# UPDATE:
# On hand (current) is forced to 0
# On hand (new) carries inventory (0 or 100)