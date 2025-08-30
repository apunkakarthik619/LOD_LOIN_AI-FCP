# Merges params_export.csv, geom_export.csv, and labels_Lxxx.csv into merged_Lxxx.csv for the chosen stage. CSV writing is Excel-friendly.

import os, csv, re

BASE      = r"C:\LOD_LOIN_AI"
DATA_DIR  = os.path.join(BASE, "data")
OUT_DIR   = os.path.join(BASE, "outputs")

PARAMS_IN = os.path.join(DATA_DIR, "params_export.csv")
GEOM_IN   = os.path.join(DATA_DIR, "geom_export.csv")          # optional
RULES_IN  = os.path.join(DATA_DIR, "LOIN_rules_by_LOD.csv")

STAGES    = ["LOD200","LOD300","LOD350","LOD400"]
KEEP_DUPLICATES = True   # keep identical (ElementId, Category) rows

NBSP = u"\u00A0"
def _s(x):
    try: return str(x).replace("\ufeff","").replace(NBSP," ").strip()
    except: return x

def read_csv(path):
    if not path or not os.path.exists(path): return [], []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        rdr = csv.DictReader(f)
        hdr = [_s(h) for h in (rdr.fieldnames or []) if _s(h)!=""]
        rows = []
        for r in rdr:
            clean = {}
            for k,v in (r or {}).items():
                if k is None: continue
                kk = _s(k)
                if kk == "": continue
                clean[kk] = _s(v)
            rows.append(clean)
        return hdr, rows

def write_csv(path, rows, headers):
    d = os.path.dirname(path)
    if d and not os.path.exists(d): os.makedirs(d)
    headers = [h for h in headers if isinstance(h, str) and h.strip()!=""]
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers); w.writeheader()
        for r in rows:
            rr = {}
            for h in headers: rr[h] = r.get(h, "")
            w.writerow(rr)

def key(r): return (_s(r.get("ElementId","")), _s(r.get("Category","")))
def dedupe(rows):
    seen, out = set(), []
    for r in rows:
        k = key(r)
        if k in seen: continue
        seen.add(k); out.append(r)
    return out

def load_rules_stage(path, stage):
    _, rr = read_csv(path)
    need = []
    for r in rr:
        try:
            if _s(r.get(stage,"0")) == "1":
                pname = _s(r.get("ParamName",""))
                if pname and pname not in need:
                    need.append(pname)
        except: pass
    return need

def derive_labels_from_loin(stage):
    suf = stage.replace("LOD","L")
    loin_csv = os.path.join(OUT_DIR, "loin_{}.csv".format(suf))
    _, rows = read_csv(loin_csv)
    out = []
    for r in rows:
        eid = _s(r.get("ElementId","")); cat = _s(r.get("Category",""))
        if eid=="" and cat=="": continue
        lp = _s(r.get("loin_pass",""))
        label = "Pass" if lp == "1" else "Fail"
        out.append({
            "ElementId": eid,
            "Category":  cat,
            "ApprovedLabel": label,
            "MissingList": _s(r.get("missing_list",""))
        })
    return out

def try_read_labels(stage):
    suf = stage.replace("LOD","L")
    lbl = os.path.join(OUT_DIR, "labels_{}.csv".format(suf))
    if os.path.exists(lbl):
        _, rows = read_csv(lbl)
        if rows: return rows
    return derive_labels_from_loin(stage)

def build_header_map(headers):
    m = {}
    for h in headers:
        m[re.sub(r"\s+"," ",_s(h)).lower()] = h
    return m
def find_header(headers, desired):
    m = build_header_map(headers)
    key = re.sub(r"\s+"," ",_s(desired)).lower()
    return m.get(key, desired)

def merge_stage(stage, params_rows, geom_rows, rules_rows, out_path):
    mandatory = load_rules_stage(RULES_IN, stage)
    g_ix = { key(r): r for r in geom_rows } if geom_rows else {}

    base = params_rows[:] if KEEP_DUPLICATES else dedupe(params_rows)

    merged = []
    for pr in base:
        m = dict(pr)
        gr = g_ix.get(key(pr))
        if gr:
            for h,v in gr.items():
                hs = _s(h)
                if hs and (hs not in m or _s(m.get(hs,""))==""):
                    m[hs] = _s(v)
        merged.append(m)

    # Flexible header matching: tolerate minor header variations (e.g., extra spaces/case) when joining files.
    # Collect the actual headers from each input CSV so merges don’t fail if a column is missing.
    hdrs = []
    for r in merged:
        for c in r.keys():
            if c not in hdrs: hdrs.append(c)

    resolved = []
    for pname in mandatory:
        actual = find_header(hdrs, pname)
        if actual not in resolved: resolved.append(actual)

    for m in merged:
        for actual in resolved:
            flag = "is_missing_" + actual
            val  = _s(m.get(actual,""))
            m[flag] = "TRUE" if (val=="" or val.lower()=="nan") else "FALSE"

    # If labels_Lxxx.csv exists, left-join to add label columns (otherwise continue without labels).
    labels = try_read_labels(stage)
    if labels:
        lab_ix = { key(r): r for r in labels }
        for m in merged:
            lr = lab_ix.get(key(m))
            if lr:
                m["ApprovedLabel"] = _s(lr.get("ApprovedLabel",""))
                m["MissingList"]  = _s(lr.get("MissingList",""))

    # Output columns in a stable order: metadata first, then parameters, then geometry, then label/flags.
    first = ["ElementId","Category","Type Name","Level","ApprovedLabel","MissingList"]
    rest, seen = [], set(first)
    for r in merged:
        for c in r.keys():
            if c not in seen:
                rest.append(c); seen.add(c)
    nonflags = sorted([c for c in rest if not c.startswith("is_missing_")])
    flags    = sorted([c for c in rest if c.startswith("is_missing_")])
    headers  = first + nonflags + flags

    write_csv(out_path, merged, headers)
    print("Stage {} → wrote {} (rows: {})".format(stage, out_path, len(merged)))

# MAIN: validate input paths, read CSVs, join on ElementId+Category (and FileName if present), write merged_Lxxx.csv.
if not os.path.exists(PARAMS_IN): raise RuntimeError("Missing: "+PARAMS_IN)
if not os.path.exists(RULES_IN):  raise RuntimeError("Missing: "+RULES_IN)

p_hdr, p_rows = read_csv(PARAMS_IN)
g_hdr, g_rows = read_csv(GEOM_IN) if os.path.exists(GEOM_IN) else ([], [])

# Remove rows where key fields (ElementId/Category) are empty to avoid junk records.
p_rows = [r for r in p_rows if not (_s(r.get("ElementId",""))=="" and _s(r.get("Category",""))=="")]

for st in STAGES:
    out_csv = os.path.join(OUT_DIR, "merged_{}.csv".format(st.replace("LOD","L")))
    merge_stage(st, p_rows, g_rows, RULES_IN, out_csv)

print("Done.")
