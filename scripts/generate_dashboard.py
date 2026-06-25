#!/usr/bin/env python3
"""
generate_dashboard.py
Reads data/CIM_Master_Data.xlsx and writes index.html (GitHub Pages entry point).

Run locally:   python scripts/generate_dashboard.py
Triggered by:  .github/workflows/generate.yml on every push that touches data/*.xlsx
"""

import json, os
from datetime import datetime
import openpyxl

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX_PATH = os.path.join(REPO_ROOT, "data", "CIM_Master_Data.xlsx")
TEMPLATE  = os.path.join(REPO_ROOT, "template", "dashboard.html")
OUTPUT    = os.path.join(REPO_ROOT, "index.html")

REGION_NORMALIZE = {
    "WI / IL":   "WI/IL",
    "Southeast": "Carolinas",
}

REGION_CONFIG = {
    "Colorado":   {"label": "Colorado",               "askNote": "regional ask",         "regions": ["Colorado"],            "key": "Colorado"},
    "Minnesota":  {"label": "Minnesota",              "askNote": "regional ask",         "regions": ["Minnesota"],           "key": "Minnesota"},
    "Florida":    {"label": "Florida",                "askNote": "regional ask",         "regions": ["Florida"],             "key": "Florida"},
    "WI/IL":      {"label": "Wisconsin / Illinois",   "askNote": "regional ask",         "regions": ["WI/IL"],               "key": "WI/IL"},
    "Midwest":    {"label": "Midwest (IN / KY / OH)", "askNote": "regional ask",         "regions": ["Midwest"],             "key": "Midwest"},
    "Arizona":    {"label": "Arizona",                "askNote": "5.5x EBITDA",          "regions": ["Arizona"],             "key": "Arizona",   "isAz": True},
    "Texas":      {"label": "Texas",                  "askNote": "5.5x EBITDA",          "regions": ["Texas"],               "key": "Texas",     "isTx": True},
    "Carolinas":  {"label": "Carolinas (NC)",         "askNote": "new / ramping",        "regions": ["Carolinas"],           "key": "Carolinas", "isNc": True},
    "California": {"label": "California",             "askNote": "pricing under review", "regions": ["California"],          "key": "California","isCa": True},
}

REGION_ORDER = ["Colorado","Minnesota","Florida","WI/IL","Midwest","Arizona","Texas","Carolinas","California"]

wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)

ws_loc = wb["Locations"]
# Strip trailing spaces from header names
loc_headers = {(cell.value.strip() if cell.value else None): cell.column - 1
               for cell in ws_loc[2] if cell.value}

def col(row, name):
    idx = loc_headers.get(name.strip())
    return row[idx] if idx is not None else None

def rnd(v, decimals=1):
    if v is None: return None
    try: return round(float(v), decimals)
    except: return None

def to_year(v):
    """Convert Excel date serial or datetime to 4-digit year string."""
    if v is None: return None
    if isinstance(v, datetime): return str(v.year)
    try:
        # Excel serial date
        n = int(float(v))
        if n > 59: n -= 1  # Excel leap year bug
        d = datetime(1900, 1, 1).toordinal() + n - 2
        return str(datetime.fromordinal(d).year)
    except: return None

def normalize_region(raw_region, state):
    if raw_region == "MN / FL":
        return "Minnesota" if state == "MN" else "Florida"
    return REGION_NORMALIZE.get(raw_region, raw_region)

locations = []
for row in ws_loc.iter_rows(min_row=3, max_row=ws_loc.max_row, values_only=True):
    uid = col(row, "Unit ID")
    if not uid:
        continue

    name         = col(row, "Location Name") or ""
    full_name    = col(row, "Name") or name
    state        = col(row, "State") or ""
    vintage      = col(row, "Vintage")
    loc_type     = col(row, "Type") or ""
    raw_rgn      = col(row, "Region / Package") or ""
    region       = normalize_region(raw_rgn, state)
    suites       = col(row, "Suites")
    sqft         = col(row, "Sq Ft")

    # Current (P3 2026)
    p3_occ       = col(row, "P3 2026 Paid Occupancy")
    p3_rev       = col(row, "P3 2026 LTM Rev ($K)")
    p3_ebitda    = col(row, "P3 2026 LTM EBITDA ($K)")
    awr          = col(row, "Most Recent Weekly Average Rate")

    # P12 2026
    p12_occ      = col(row, "P12 2026 Paid Occupancy")
    p12_rev      = col(row, "LTM P12 2026 Rev ($K)")
    p12_ebitda   = col(row, "LTM P12 2026 EBITDA ($K)")

    # Mature
    mature_occ   = col(row, "Mature Paid Occupancy")
    mature_rev   = col(row, "Mature Rev ($K)")
    mature_eb    = col(row, "Mature EBITDA ($K)")
    mature_awr   = col(row, "Mature Weekly Average Rate")

    # Valuation
    val_p12      = col(row, "Valuation P12 2026")
    val_mature   = col(row, "Valuation Mature")
    ask_ovr      = col(row, "Ask Price ($K)")
    ask_price    = ask_ovr if ask_ovr else (val_p12 if val_p12 and float(val_p12) > 0 else None)

    # Lease
    lease_exp    = to_year(col(row, "Current Lease Expiration"))
    options      = col(row, "Options") or ""
    next_opt     = to_year(col(row, "Next option Expiration"))
    last_opt     = to_year(col(row, "Last Option Expiration"))

    # Demographics
    salons_3mi   = col(row, "3 mile salon counts")
    pop_3mi      = col(row, "3 mile population")

    # Notes
    notes_txt    = col(row, "Notes") or ""
    exp_notes    = col(row, "Expansion Notes") or ""
    short_note   = notes_txt.split(".")[0] if notes_txt else ""
    if len(short_note) > 90: short_note = short_note[:87] + "..."

    loc_id = (state.lower() + "_" + "".join(
        c for c in name.lower() if c.isalnum()
    ))[:30]

    locations.append({
        "id":         loc_id,
        "uid":        uid,
        "name":       name,
        "fullName":   full_name,
        "vintage":    int(vintage) if vintage else None,
        "locType":    loc_type,
        "state":      state,
        "region":     region,
        "suites":     int(suites) if suites else None,
        "sqft":       int(sqft) if sqft else None,
        # Current P3
        "p3Occ":      rnd(p3_occ),
        "p3Rev":      rnd(p3_rev),
        "p3EBITDA":   rnd(p3_ebitda),
        "awr":        rnd(awr),
        # P12
        "p12Occ":     rnd(p12_occ),
        "p12Rev":     rnd(p12_rev),
        "p12EBITDA":  rnd(p12_ebitda),
        # Mature
        "matureOcc":  rnd(mature_occ),
        "matureRev":  rnd(mature_rev),
        "matureEB":   rnd(mature_eb),
        "matureAWR":  rnd(mature_awr),
        # Ask
        "askPrice":   rnd(ask_price),
        "valP12":     rnd(val_p12),
        "valMature":  rnd(val_mature),
        # Lease
        "leaseExp":   lease_exp,
        "options":    options,
        "nextOpt":    next_opt,
        "lastOpt":    last_opt,
        # Demographics
        "salons3mi":  int(salons_3mi) if salons_3mi else None,
        "pop3mi":     int(pop_3mi) if pop_3mi else None,
        # Notes
        "notes":      short_note,
        "fullNotes":  notes_txt,
        "expansionNotes": exp_notes,
        # Legacy aliases (sidebar totals still use these)
        "occ":        rnd(p3_occ),
        "ltmRev":     rnd(p3_rev),
        "ltmEBITDA":  rnd(p3_ebitda),
        "pfEBITDA":   rnd(mature_eb),
    })

# Packages tab
ws_pkg = wb["Packages"]
pkg_headers = {(cell.value.strip() if cell.value else None): cell.column - 1
               for cell in ws_pkg[2] if cell.value}

def pcol(row, name):
    idx = pkg_headers.get(name.strip())
    return row[idx] if idx is not None else None

pkg_asks = {}
for row in ws_pkg.iter_rows(min_row=3, max_row=ws_pkg.max_row, values_only=True):
    pkg_name = pcol(row, "Package Name")
    if not pkg_name: continue
    total_ask_k = pcol(row, "Total Ask Override ($K)")
    if pkg_name == "MN / FL":
        norm = "Minnesota"
    else:
        norm = REGION_NORMALIZE.get(pkg_name, pkg_name)
    if total_ask_k and float(total_ask_k) > 0:
        pkg_asks[norm] = total_ask_k

region_groups = []
for key in REGION_ORDER:
    cfg = REGION_CONFIG.get(key)
    if not cfg: continue
    total_ask_k = pkg_asks.get(key)
    if not total_ask_k:
        total_ask_k = sum(
            (loc["askPrice"] or 0)
            for loc in locations
            if loc["region"] in cfg["regions"] and loc["askPrice"]
        ) or None
    ask_str = ("$" + "{:.2f}".format(float(total_ask_k) / 1000) + "M") if total_ask_k and float(total_ask_k) > 0 else "TBD"
    grp = {"key": cfg["key"], "label": cfg["label"], "ask": ask_str, "askNote": cfg["askNote"], "regions": cfg["regions"]}
    for flag in ("isAz","isTx","isCa","isNc"):
        if cfg.get(flag): grp[flag] = True
    region_groups.append(grp)

# JS rendering
def js_val(v):
    if v is None:           return "null"
    if isinstance(v, bool): return "true" if v else "false"
    if isinstance(v, str):  return json.dumps(v, ensure_ascii=False)
    return str(v)

def render_loc(loc):
    keys = ["id","uid","name","fullName","state","region","vintage","locType","suites","sqft",
            "p3Occ","p3Rev","p3EBITDA","awr",
            "p12Occ","p12Rev","p12EBITDA",
            "matureOcc","matureRev","matureEB","matureAWR",
            "askPrice","valP12","valMature","leaseExp","options","nextOpt","lastOpt",
            "salons3mi","pop3mi","notes","fullNotes",
            "occ","ltmRev","ltmEBITDA","pfEBITDA","expansionNotes"]
    return "  { " + ", ".join(f"{k}:{js_val(loc[k])}" for k in keys) + " }"

def render_grp(grp):
    keys  = ["key","label","ask","askNote","regions"]
    flags = ["isAz","isTx","isCa","isNc"]
    parts = [f"{k}:{js_val(grp[k])}" for k in keys if k in grp]
    parts += [f"{k}:{js_val(grp[k])}" for k in flags if k in grp]
    return "  { " + ", ".join(parts) + " }"

data_block = (
    "const locations = [\n"
    + ",\n".join(render_loc(l) for l in locations)
    + "\n];\n\n"
    + "const REGION_GROUPS = [\n"
    + ",\n".join(render_grp(g) for g in region_groups)
    + "\n];"
)

with open(TEMPLATE, "r", encoding="utf-8") as f:
    template = f.read()

with open(OUTPUT, "w", encoding="utf-8") as f:
    f.write(template.replace("{{DATA_BLOCK}}", data_block))

print(f"index.html written  ({len(locations)} locations, {len(region_groups)} region groups)")
from collections import Counter
counts = Counter(l["region"] for l in locations)
for r in REGION_ORDER:
    print(f"  {r}: {counts.get(r,0)} locations")
