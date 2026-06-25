#!/usr/bin/env python3
"""
generate_dashboard.py
Auto-detects xlsx format and generates index.html.

Format A (26-col legacy):  max_col < 35, headers row 2, data row 3+
Format B (42-col current): max_col >= 35, headers row 2, data row 3+
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
    "Colorado":   {"label": "Colorado",               "askNote": "regional ask",         "regions": ["Colorado"],   "key": "Colorado"},
    "Minnesota":  {"label": "Minnesota",               "askNote": "regional ask",         "regions": ["Minnesota"],  "key": "Minnesota"},
    "Florida":    {"label": "Florida",                 "askNote": "regional ask",         "regions": ["Florida"],    "key": "Florida"},
    "WI/IL":      {"label": "Wisconsin / Illinois",    "askNote": "regional ask",         "regions": ["WI/IL"],      "key": "WI/IL"},
    "Midwest":    {"label": "Midwest (IN / KY / OH)",  "askNote": "regional ask",         "regions": ["Midwest"],    "key": "Midwest"},
    "Arizona":    {"label": "Arizona",                 "askNote": "5.5x EBITDA",          "regions": ["Arizona"],    "key": "Arizona",    "isAz": True},
    "Texas":      {"label": "Texas",                   "askNote": "5.5x EBITDA",          "regions": ["Texas"],      "key": "Texas",      "isTx": True},
    "Carolinas":  {"label": "Carolinas (NC)",          "askNote": "new / ramping",        "regions": ["Carolinas"],  "key": "Carolinas",  "isNc": True},
    "California": {"label": "California",              "askNote": "pricing under review", "regions": ["California"], "key": "California", "isCa": True},
}
REGION_ORDER = ["Colorado","Minnesota","Florida","WI/IL","Midwest","Arizona","Texas","Carolinas","California"]


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
    return round(float(v), 1) if v is not None and sf(v) is not None else None

def mk_id(state, name):
    return f"{str(state).lower()}_{re.sub(r'[^a-z0-9]','',str(name).lower())}"


# ── Format A: 26-column legacy ───────────────────────────────────────────────
# A=UID  B=Name  C=City  D=State  E=Region  F=SubRegion  G=Type  H=Vintage
# I=Suites  J=SqFt  K=PaidOcc  L=FilledOcc  M=LTMRev  N=LTMEBITDA
# O=EBITDAMargin  P=MatureEBITDA  Q=Val5.5x  R=Val6.0x  S=AskPrice
# T=InitOpen  U=AcqDate  V=LeaseExp  W=InPkg  X=PkgName  Y=Notes  Z=AWR

def parse_26(ws):
    locs = []
    for r in range(3, ws.max_row + 1):
        def g(c): return ws.cell(r, c).value
        uid   = g(1); name = g(2); state = g(4)
        if not uid or not name or not state: continue
        region = normalize_region(g(5), state)
        if region not in REGION_CONFIG: continue
        state = str(state).strip()
        p3Rev = sf(g(13)); p3EB = sf(g(14)); mEB = sf(g(16))
        ask   = sf(g(19)); lexp = sy(g(22)); awr = sf(g(26))
        notes = ss(g(25)) or ''
        locs.append({
            "id": mk_id(state, name), "uid": uid,
            "name": ss(name), "fullName": f"{state} | {ss(name)}",
            "state": state, "region": region,
            "vintage": si(g(8)), "locType": ss(g(7)) or '',
            "suites": si(g(9)), "sqft": si(g(10)),
            "p3Occ": sf(g(11)), "p3Rev": rnd(p3Rev), "p3EBITDA": rnd(p3EB),
            "awr": rnd(awr),
            "p12Occ": None, "p12Rev": None, "p12EBITDA": None,
            "matureOcc": None, "matureRev": None, "matureEB": rnd(mEB), "matureAWR": None,
            "askPrice": rnd(ask),
            "valP12": rnd(sf(g(17))), "valMature": rnd(sf(g(18))),
            "leaseExp": lexp, "options": None, "nextOpt": None, "lastOpt": None,
            "salons3mi": None, "pop3mi": None,
            "notes": notes[:80], "fullNotes": notes, "expansionNotes": "",
            "occ": sf(g(11)), "ltmRev": rnd(p3Rev), "ltmEBITDA": rnd(p3EB), "pfEBITDA": rnd(mEB),
        })
    return locs


# ── Format B: 42-column current ──────────────────────────────────────────────
# Row 2 = headers, row 3+ = data
# A=UID  B=FullName  C=ShortName  D=City  E=State  F=Region  G=SubRegion
# H=Type  I=Vintage  J=Suites  K=SqFt
# L=PaidOcc  M=FilledOcc  N=P3PaidOcc  O=P3FilledOcc
# P=P3Rev  Q=P3EBITDA  R=P3Margin  S=AWR
# T=P12Occ  U=P12Rev  V=P12EBITDA  W=P12Margin
# X=MatureAWR  Y=MatureOcc  Z=MatureRev  AA=MatureEBITDA  AB=MatureMargin
# AC=ValP12  AD=ValMature  AE=AskPrice
# AF=InitOpen  AG=AcqDate  AH=LeaseExp  AI=Options  AJ=NextOpt  AK=LastOpt
# AL=Salons3mi  AM=Pop3mi  AN=InPkg  AO=PkgName  AP=Notes  AQ=ExpNotes

def parse_42(ws):
    locs = []
    for r in range(3, ws.max_row + 1):
        def g(c): return ws.cell(r, c).value
        uid   = g(1); state = g(5)
        name  = ss(g(3)) or ss(g(2))  # prefer short name (col C)
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
            "p3Occ": sf(g(14)), "p3Rev": rnd(p3Rev), "p3EBITDA": rnd(p3EB),
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


def load_and_parse():
    with open(XLSX_PATH, 'rb') as f:
        raw = f.read()
    wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
    ws = wb['Locations'] if 'Locations' in wb.sheetnames else wb.active
    fmt = "42-column current" if ws.max_column >= 35 else "26-column legacy"
    print(f"Detected: {fmt} format  (max_col={ws.max_column})")
    return parse_42(ws) if ws.max_column >= 35 else parse_26(ws)


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
        g = {"key": key, "label": cfg["label"], "ask": ask_str,
             "askNote": cfg["askNote"], "regions": cfg["regions"]}
        for flag in ("isAz","isTx","isNc","isCa"):
            if cfg.get(flag): g[flag] = True
        groups.append(g)
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
