# Reads loin_Lxxx.csv and writes labels_Lxxx.csv with simple review labels (e.g., "Compliant"/"Non-Compliant") for each element. CSV is Excel-friendly.

import os, csv
BASE    = r"C:\LOD_LOIN_AI"
IN_DIR  = os.path.join(BASE, "outputs")
OUT_DIR = os.path.join(BASE, "outputs")
STAGES  = ["L200","L300","L350","L400"]

def _s(x):
    try: return str(x).replace("\ufeff","").strip()
    except: return x

def read_csv(path):
    if not os.path.exists(path): return [], []
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
            row = {h: r.get(h,"") for h in headers}
            w.writerow(row)

for suf in STAGES:
    loin_path   = os.path.join(IN_DIR, "loin_{}.csv".format(suf))
    labels_path = os.path.join(OUT_DIR, "labels_{}.csv".format(suf))
    hdr, rows = read_csv(loin_path)
    if not rows:
        print("Missing or empty:", loin_path); 
        continue

    out = []
    for r in rows:
        eid = _s(r.get("ElementId","")); cat = _s(r.get("Category",""))
        if eid=="" and cat=="":    # skip fully blank keys
            continue
        lp = _s(r.get("loin_pass",""))
        label = "Pass" if lp == "1" else "Fail"
        out.append({
            "ElementId": eid,
            "Category":  cat,
            "ApprovedLabel": label,
            "MissingList": _s(r.get("missing_list",""))
        })

    write_csv(labels_path, out, ["ElementId","Category","ApprovedLabel","MissingList"])
    print("Wrote {} (rows: {})".format(labels_path, len(out)))

print("Done.")
