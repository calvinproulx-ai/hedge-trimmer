#!/usr/bin/env python3
"""
generate_dashboard.py
Auto-detects xlsx format and generates index.html.

Format A (26-col legacy):  max_col < 35
Format B (42-col current): 35 <= max_col < 45
Format C (47-col current): max_col >= 45  ← v21+ data
"""
import json, os, io, re
from datetime import datetime
import openpyxl

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX_PATH = os.path.join(REPO_ROOT, "data", "CIM_Master_Data.xlsx")
TEMPLATE  = os.path.join(REPO_ROOT, "template", "dashboard.html")
OUTPUT    = os.path.join(REPO_ROOT, "index.html")

REGION_NORMALIZE = {
    "WI / IL": "WI/IL",
    "Southeast": "Carolinas",
}
REGION_CONFIG = {
    "Colorado":   {"label": "Colorado",               "askNote": "regional ask",         "regions": ["Colorado"]},
    "Minnesota":  {"label": "Minnesota",               "askNote": "regional ask",         "regions": ["Minnesota"]},
    "Florida":    {"label": "Florida",                 "askNote": "regional ask",         "regions": ["Florida"]},
    "WI/IL":      {"label": "Wisconsin / Illinois",    "askNote": "regional ask",         "regions": ["WI/IL"]},
    "Midwest":    {"label": "Midwest (IN / KY / OH)",  "askNote": "regional ask",         "regions": ["Midwest"]},
    "Arizona":    {"label": "Arizona",                 "askNote": "5.5x EBITDA",          "regions": ["Arizona"]},
    "Texas":      {"label": "Texas",                   "askNote": "regional ask",         "regions": ["Texas"]},
    "Houston":    {"label": "Houston",                 "askNote": "regional ask",         "regions": ["Houston"]},
    "Carolinas":  {"label": "Carolinas (NC)",          "askNote": "new / ramping",        "regions": ["Carolinas"]},
    "California": {"label": "California",              "askNote": "pricing under review", "regions": ["California"]},
}
REGION_ORDER = ["Colorado","Minnesota","Florida","WI/IL","Midwest","Arizona","Texas","Houston","Carolinas","California"]


def normalize_region(raw, state):
    raw = str(raw or '').strip()
    if raw in ("MN / FL", "MN/FL"):
        return "Minnesota" if str(state).strip() == "MN" else "Florida"
    return REGION_NORMALIZE.get(raw, raw)

def sf(v):
    try:
        f = float(v)
        return None if f != f else f
    except: return None

def si(v):
    f = sf(v)
    return int(round(f)) if f is not None else None

def sy(v):
    if v is None: return None
    if isinstance(v, datetime): return str(v.year)
    m = re.search(r'\b(20\d{2})\b', str(v))
    return m.group(1) if m else (str(v).strip() or None)

def ss(v):
    return str(v).strip() if v else None

def rnd(v):
    return round(float(v), 4) if v is not None and sf(v) is not None else None

def mk_id(state, name):
    return f"{str(state).lower()}_{re.sub(r'[^a-z0-9]','',str(name).lower())}"


# ── Format A: 26-column legacy ───────────────────────────────────────────────
def parse_26(ws):
    locs = []
    for r in range(3, ws.max_row + 1):
        def g(c): return ws.cell(r, c).value
        uid = g(1); name = g(2); state = g(4)
        if not uid or not name or not state: continue
        region = normalize_region(g(5), state)
        if region not in REGION_CONFIG: continue
        state = str(state).strip()
        p3Rev = sf(g(13)); p3EB = sf(g(14)); mEB = sf(g(16))
        ask   = sf(g(19)); notes = ss(g(25)) or ''
        locs.append({
            "id": mk_id(state, name), "uid": uid,
            "name": ss(name), "fullName": f"{state} | {ss(name)}",
            "state": state, "region": region,
            "vintage": si(g(8)), "locType": ss(g(7)) or '',
            "suites": si(g(9)), "sqft": si(g(10)),
            "p6Occ": sf(g(11)), "p3Occ": sf(g(11)), "p3Rev": rnd(p3Rev), "p3EBITDA": rnd(p3EB),
            "awr": rnd(sf(g(26))),
            "p12Occ": None, "p12Rev": None, "p12EBITDA": None,
            "matureOcc": None, "matureRev": None, "matureEB": rnd(mEB), "matureAWR": None,
            "askPrice": rnd(ask),
            "valP12": rnd(sf(g(17))), "valMature": rnd(sf(g(18))),
            "leaseExp": sy(g(22)), "options": None, "nextOpt": None, "lastOpt": None,
            "salons3mi": None, "pop3mi": None,
            "notes": notes[:80], "fullNotes": notes, "expansionNotes": "",
            "occ": sf(g(11)), "ltmRev": rnd(p3Rev), "ltmEBITDA": rnd(p3EB), "pfEBITDA": rnd(mEB),
        })
    return locs


# ── Format B: 42-column (max_col 35-44) ──────────────────────────────────────
def parse_42(ws):
    locs = []
    for r in range(3, ws.max_row + 1):
        def g(c): return ws.cell(r, c).value
        uid = g(1); state = g(5)
        name = ss(g(3)) or ss(g(2))
        if not uid or not name or not state: continue
        region = normalize_region(g(6), state)
        if region not in REGION_CONFIG: continue
        state = str(state).strip()
        p3Rev = sf(g(16)); p3EB = sf(g(17))
        p12Rev= sf(g(21)); p12EB= sf(g(22))
        mRev  = sf(g(26)); mEB  = sf(g(27))
        ask   = sf(g(31))
        notes = ss(g(42)) or ''; expn = ss(g(43)) or ''
        locs.append({
            "id": mk_id(state, name), "uid": uid,
            "name": name, "fullName": ss(g(2)) or f"{state} | {name}",
            "state": state, "region": region,
            "vintage": si(g(9)), "locType": ss(g(8)) or '',
            "suites": si(g(10)), "sqft": si(g(11)),
            "p6Occ": sf(g(14)), "p3Occ": sf(g(14)), "p3Rev": rnd(p3Rev), "p3EBITDA": rnd(p3EB),
            "awr": rnd(sf(g(19))),
            "p12Occ": sf(g(20)), "p12Rev": rnd(p12Rev), "p12EBITDA": rnd(p12EB),
            "matureOcc": sf(g(25)), "matureRev": rnd(mRev),
            "matureEB": rnd(mEB), "matureAWR": rnd(sf(g(24))),
            "askPrice": rnd(ask),
            "valP12": rnd(sf(g(29))), "valMature": rnd(sf(g(30))),
            "leaseExp": sy(g(34)), "options": ss(g(35)),
            "nextOpt": sy(g(36)), "lastOpt": sy(g(37)),
            "salons3mi": si(g(38)), "pop3mi": si(g(39)),
            "notes": notes[:80], "fullNotes": notes, "expansionNotes": expn,
            "occ": sf(g(14)), "ltmRev": rnd(p3Rev), "ltmEBITDA": rnd(p3EB), "pfEBITDA": rnd(mEB),
        })
    return locs


# ── Format C: 47-column (max_col >= 45) — v21+ ───────────────────────────────
# Col 1=UID  2=FullName  3=ShortName  4=Address  5=City  6=State  7=Region
# 8=SubRegion  9=Type  10=Vintage  11=Suites  12=SqFt
# 13=P3PaidOcc  14=P3FilledOcc  15=P3Rev($K)  16=P3EBITDA($K)  17=P3Margin
# 18=P6PaidOcc  19=P6FilledOcc  20=AWR
# 21=P12PaidOcc  22=P12Rev($K)  23=P12EBITDA($K)  24=P12Margin
# 25=MatureAWR  26=MatureOcc  27=MatureRev($K)  28=MatureEBITDA($K)  29=MatureMargin
# 30=ValP12($K)  31=ValMature($K)  32=AskPrice($K)
# 33=OpenDate  34=AcqDate  35=LeaseExp  36=Options  37=NextOpt  38=LastOpt
# 39=Salons3mi  40=Pop3mi  41=InPkg  42=PkgName
# 43=Notes  44=ExpNotes  45=Entity  46=MarketLabel  47=CalloutNote

def parse_47(ws):
    locs = []
    for r in range(3, ws.max_row + 1):
        def g(c): return ws.cell(r, c).value
        uid = g(1); state = g(6)
        name = ss(g(3)) or ss(g(2))
        if not uid or not name or not state: continue
        region = normalize_region(g(7), state)
        if region not in REGION_CONFIG: continue
        state = str(state).strip()

        # All monetary values in xlsx are $K → store as-is (JS divides by 1000 for M display)
        p3Rev  = sf(g(15)); p3EB  = sf(g(16))
        p12Rev = sf(g(22)); p12EB = sf(g(23))
        mRev   = sf(g(27)); mEB   = sf(g(28))
        ask    = sf(g(32))

        # Skip locations with no financial data
        if ask is None and p3Rev is None:
            continue

        p6occ  = sf(g(18))  # P6 2026 Paid Occ — primary display metric
        p3occ  = sf(g(13))  # P3 2026 Paid Occ
        p12occ = sf(g(21))  # P12 2026 Paid Occ
        mawocc = sf(g(26))  # Mature Paid Occ

        notes = ss(g(43)) or ''; expn = ss(g(44)) or ''
        leaseRaw = g(35)
        leaseExp = None
        if isinstance(leaseRaw, datetime):
            leaseExp = leaseRaw.strftime("%b %Y")
        elif leaseRaw:
            leaseExp = str(leaseRaw).strip()[:10]

        locs.append({
            "id": mk_id(state, name), "uid": uid,
            "name": name, "fullName": ss(g(2)) or f"{state} | {name}",
            "state": state, "region": region,
            "vintage": si(g(10)), "locType": ss(g(9)) or '',
            "suites": si(g(11)), "sqft": si(g(12)),
            "p6Occ": p6occ,
            "p3Occ": p3occ, "p3Rev": rnd(p3Rev), "p3EBITDA": rnd(p3EB),
            "awr": rnd(sf(g(20))),
            "p12Occ": p12occ, "p12Rev": rnd(p12Rev), "p12EBITDA": rnd(p12EB),
            "matureOcc": mawocc, "matureRev": rnd(mRev),
            "matureEB": rnd(mEB), "matureAWR": rnd(sf(g(25))),
            "askPrice": rnd(ask),
            "valP12": rnd(sf(g(30))), "valMature": rnd(sf(g(31))),
            "leaseExp": leaseExp, "options": ss(g(36)),
            "nextOpt": sy(g(37)), "lastOpt": sy(g(38)),
            "salons3mi": si(g(39)), "pop3mi": si(g(40)),
            "notes": notes[:80], "fullNotes": notes, "expansionNotes": expn,
            # Legacy aliases used by dashboard JS
            "occ": p6occ,
            "ltmRev": rnd(p3Rev), "ltmEBITDA": rnd(p3EB), "pfEBITDA": rnd(mEB),
        })
    return locs


def load_and_parse():
    with open(XLSX_PATH, 'rb') as f:
        raw = f.read()
    wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
    ws = wb['Locations'] if 'Locations' in wb.sheetnames else wb.active
    mc = ws.max_column
    if mc >= 45:
        fmt = "47-column v21+"
        locs = parse_47(ws)
    elif mc >= 35:
        fmt = "42-column current"
        locs = parse_42(ws)
    else:
        fmt = "26-column legacy"
        locs = parse_26(ws)
    print(f"Detected: {fmt} format  (max_col={mc})")
    return locs


def build_groups(locs):
    totals = {k: 0.0 for k in REGION_CONFIG}
    for l in locs:
        if l['region'] in totals and l['askPrice']:
            totals[l['region']] += l['askPrice']
    groups = []
    for key in REGION_ORDER:
        cfg = REGION_CONFIG[key]
        if not any(l['region'] == key for l in locs):
            continue
        t = totals[key]
        ask_str = f"${t/1000:.2f}M" if t > 0 else "TBD"
        groups.append({"key": key, "label": cfg["label"], "ask": ask_str,
                       "askNote": cfg["askNote"], "regions": cfg["regions"]})
    return groups


def main():
    locs   = load_and_parse()
    groups = build_groups(locs)

    real = [l for l in locs if l.get('id') != '_']
    with_data = [l for l in real if l['p3Rev'] is not None]
    total_suites = sum(l['suites'] or 0 for l in real)
    total_rev    = sum(l['p3Rev']  or 0 for l in with_data)
    total_eb     = sum(l['p3EBITDA'] or 0 for l in with_data)
    total_ask    = sum(l['askPrice'] or 0 for l in real if l['askPrice'])
    has_tbd      = any(not l['askPrice'] for l in real)

    stat_locs   = len(real)
    stat_suites = f"{total_suites:,}"
    stat_rev    = f"${total_rev/1000:.1f}M"
    stat_eb     = f"${total_eb/1000:.1f}M"
    stat_ask    = f"${total_ask/1000:.1f}M" + ("+" if has_tbd else "")

    data_block = (f"const locations = {json.dumps(locs, indent=2)};\n\n"
                  f"const REGION_GROUPS = {json.dumps(groups, indent=2)};\n")

    with open(TEMPLATE, 'r', encoding='utf-8') as f:
        tmpl = f.read()

    out = tmpl.replace('{{DATA_BLOCK}}', data_block)
    out = re.sub(r'id="statLocs">[^<]*',   f'id="statLocs">{stat_locs}',   out)
    out = re.sub(r'id="statSuites">[^<]*', f'id="statSuites">{stat_suites}',out)
    out = re.sub(r'id="statRev">[^<]*',    f'id="statRev">{stat_rev}',      out)
    out = re.sub(r'id="statEBITDA">[^<]*', f'id="statEBITDA">{stat_eb}',    out)
    out = re.sub(r'id="statAsk">[^<]*',    f'id="statAsk">{stat_ask}',      out)

    with open(OUTPUT, 'w', encoding='utf-8') as f:
        f.write(out)

    rc = {}
    for l in real: rc[l['region']] = rc.get(l['region'], 0) + 1
    print(f"index.html written  ({stat_locs} locations, {len(groups)} region groups)")
    for key in REGION_ORDER:
        if key in rc: print(f"  {key}: {rc[key]} locations")

if __name__ == '__main__':
    main()
