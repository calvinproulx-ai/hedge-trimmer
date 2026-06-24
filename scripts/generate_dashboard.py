#!/usr/bin/env python3
"""
generate_dashboard.py
Reads data/CIM_Master_Data.xlsx and writes index.html (GitHub Pages entry point).

Run locally:   python scripts/generate_dashboard.py
Triggered by:  .github/workflows/generate.yml on every push that touches data/*.xlsx
"""

import json
import os
import openpyxl

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX_PATH = os.path.join(REPO_ROOT, "data", "CIM_Master_Data.xlsx")
TEMPLATE  = os.path.join(REPO_ROOT, "template", "dashboard.html")
OUTPUT    = os.path.join(REPO_ROOT, "index.html")

# Normalize region strings from xlsx so they match filter button values exactly.
# Keys are what appears in xlsx "Region / Package" column.
# Values are what the filter buttons pass to setFilter().
REGION_NORMALIZE = {
    "WI / IL":   "WI/IL",
    "Southeast": "Carolinas",
    # MN / FL is handled per-location by state (MN->Minnesota, FL->Florida)
}

# Display config keyed by NORMALIZED region name
REGION_CONFIG = {
    "Colorado":   {"label": "Colorado",               "askNote": "regional ask",
                   "regions": ["Colorado"],            "key": "Colorado"},
    "Minnesota":  {"label": "Minnesota + Florida",    "askNote": "regional ask",
                   "regions": ["Minnesota","Florida"], "key": "MN+FL"},
    "WI/IL":      {"label": "Wisconsin / Illinois",   "askNote": "regional ask",
                   "regions": ["WI/IL"],               "key": "WI/IL"},
    "Midwest":    {"label": "Midwest (IN / KY / OH)", "askNote": "regional ask",
                   "regions": ["Midwest"],             "key": "Midwest"},
    "Arizona":    {"label": "Arizona",                "askNote": "5.5x EBITDA",
                   "regions": ["Arizona"],             "key": "Arizona",    "isAz": True},
    "Texas":      {"label": "Texas",                  "askNote": "5.5x EBITDA",
                   "regions": ["Texas"],               "key": "Texas",      "isTx": True},
    "Carolinas":  {"label": "Carolinas (NC)",         "askNote": "new / ramping",
                   "regions": ["Carolinas"],           "key": "Carolinas",  "isNc": True},
    "California": {"label": "California",             "askNote": "pricing under review",
                   "regions": ["California"],          "key": "California", "isCa": True},
}

# Desired display order
REGION_ORDER = ["Colorado","Minnesota","WI/IL","Midwest","Arizona","Texas","Carolinas","California"]

wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)

ws_loc = wb["Locations"]
loc_headers = {cell.value: cell.column - 1 for cell in ws_loc[2] if cell.value}

def col(row, name):
    idx = loc_headers.get(name)
    return row[idx] if idx is not None else None

def rnd(v):
    if v is None: return None
    try: return round(float(v), 1)
    except: return None

def normalize_region(raw_region, state):
    """Map xlsx region string to filter-compatible value."""
    if raw_region == "MN / FL":
        return "Minnesota" if state == "MN" else "Florida"
    return REGION_NORMALIZE.get(raw_region, raw_region)

locations = []
for row in ws_loc.iter_rows(min_row=3, max_row=ws_loc.max_row, values_only=True):
    uid = col(row, "Unit ID")
    if not uid:
        continue
    name      = col(row, "Location Name") or ""
    state     = col(row, "State") or ""
    raw_rgn   = col(row, "Region / Package") or ""
    region    = normalize_region(raw_rgn, state)
    suites    = col(row, "Suites")
    sqft      = col(row, "Sq Ft")
    paid_occ  = col(row, "Paid Occ% (5/5)")
    ltm_rev   = col(row, "LTM Rev ($K)")
    ltm_eb    = col(row, "LTM EBITDA ($K)")
    mature_eb = col(row, "Mature EBITDA ($K)")
    val55     = col(row, "Val @ 5.5x ($K)")
    ask_ovr   = col(row, "Ask Price ($K)")
    notes_txt = col(row, "Notes") or ""
    awr       = col(row, "Wkly Rev/Suite")

    ask_price = ask_ovr if ask_ovr else (val55 if val55 and val55 > 0 else None)
    short_note = notes_txt.split(".")[0] if notes_txt else ""
    if len(short_note) > 80:
        short_note = short_note[:77] + "..."

    loc_id = (state.lower() + "_" + "".join(
        c for c in name.lower() if c.isalnum()
    ))[:30]

    locations.append({
        "id":        loc_id,
        "uid":       uid,
        "name":      name,
        "state":     state,
        "region":    region,
        "suites":    int(suites) if suites else None,
        "sqft":      int(sqft) if sqft else None,
        "occ":       rnd(paid_occ),
        "ltmRev":    rnd(ltm_rev),
        "ltmEBITDA": rnd(ltm_eb),
        "pfEBITDA":  rnd(mature_eb),
        "askPrice":  rnd(ask_price),
        "awr":       rnd(awr),
        "notes":     short_note,
        "fullNotes": notes_txt,
    })

# Build REGION_GROUPS from the packages tab for ask prices,
# then use REGION_CONFIG for display + REGION_ORDER for sequence.
ws_pkg = wb["Packages"]
pkg_headers = {cell.value: cell.column - 1 for cell in ws_pkg[2] if cell.value}

def pcol(row, name):
    idx = pkg_headers.get(name)
    return row[idx] if idx is not None else None

# Collect ask overrides from Packages tab keyed by normalized region
pkg_asks = {}
for row in ws_pkg.iter_rows(min_row=3, max_row=ws_pkg.max_row, values_only=True):
    pkg_name = pcol(row, "Package Name")
    if not pkg_name:
        continue
    total_ask_k = pcol(row, "Total Ask Override ($K)")
    # Normalize the package name too
    if pkg_name == "MN / FL":
        norm = "Minnesota"   # MN+FL package maps to Minnesota key
    else:
        norm = REGION_NORMALIZE.get(pkg_name, pkg_name)
    if total_ask_k and total_ask_k > 0:
        pkg_asks[norm] = total_ask_k

region_groups = []
for key in REGION_ORDER:
    cfg = REGION_CONFIG.get(key)
    if not cfg:
        continue

    total_ask_k = pkg_asks.get(key)
    if not total_ask_k:
        # Sum from locations
        total_ask_k = sum(
            (loc["askPrice"] or 0)
            for loc in locations
            if loc["region"] in cfg["regions"] and loc["askPrice"]
        ) or None

    ask_str = ("$" + "{:.2f}".format(total_ask_k / 1000) + "M") if total_ask_k and total_ask_k > 0 else "TBD"

    grp = {
        "key":     cfg["key"],
        "label":   cfg["label"],
        "ask":     ask_str,
        "askNote": cfg["askNote"],
        "regions": cfg["regions"],
    }
    for flag in ("isAz", "isTx", "isCa", "isNc"):
        if cfg.get(flag):
            grp[flag] = True

    region_groups.append(grp)

# Render JS
def js_val(v):
    if v is None:           return "null"
    if isinstance(v, bool): return "true" if v else "false"
    if isinstance(v, str):  return json.dumps(v, ensure_ascii=False)
    return str(v)

def render_loc(loc):
    keys = ["id","uid","name","state","region","suites","sqft",
            "occ","ltmRev","ltmEBITDA","pfEBITDA","askPrice","awr","notes","fullNotes"]
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
# Show region breakdown
from collections import Counter
counts = Counter(l["region"] for l in locations)
for r in REGION_ORDER:
    print(f"  {r}: {counts.get(r,0)} locations")
