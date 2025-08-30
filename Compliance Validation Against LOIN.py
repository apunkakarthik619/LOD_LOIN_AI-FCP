# Validates exported parameters against LOIN rules with unit handling (e.g., mm↔m). CSV output is Excel-safe (handles commas/quotes).

import os, csv, re
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory as BIC

BASE_DIR   = r"C:\LOD_LOIN_AI"
RULES_CSV  = os.path.join(BASE_DIR, r"data\LOIN_rules_by_LOD.csv")
PARAMS_OUT = os.path.join(BASE_DIR, r"data\params_export.csv")   # assumes you've run Step 2 already
OUT_DIR    = os.path.join(BASE_DIR, r"outputs")
STAGES     = ["LOD200","LOD300","LOD350","LOD400"]

NBSP = u"\u00A0"
def _s(x):
    try: return str(x).replace("\ufeff","").replace(NBSP," ").strip()
    except: return x
def clean_row(row):
    out = {}
    for k,v in (row or {}).items():
        if k is None: continue
        ks = _s(k)
        if ks == "": continue
        out[ks] = _s(v)
    return out

def load_rules(path):
    rules = []
    with open(path, "r", encoding="utf-8-sig") as f:
        rdr = csv.DictReader(f)
        for row in rdr:
            row = clean_row(row)
            flags = {}
            for s in STAGES:
                try: flags[s] = int(_s(row.get(s,"0")) or 0)
                except: flags[s] = 0
            rules.append({
                "Category": _s(row.get("Category","")),
                "ParamName": _s(row.get("ParamName","")),
                "Type": _s(row.get("Type","text")).lower(),
                "AllowedValues": _s(row.get("AllowedValues","")),
                "Regex": _s(row.get("Regex","")),
                "Min": _s(row.get("Min","")),
                "Max": _s(row.get("Max","")),
                "Notes": _s(row.get("Notes","")),
                "Stages": flags
            })
    return rules

def present(v):
    s = _s(v)
    return s != "" and s.lower() != "nan"

def get_unit_hint(rule):
    if "units: mm" in _s(rule.get("Notes","")).lower():
        return "mm"
    return None

def parse_number_unit_aware(value, unit_hint):
    s = _s(value).lower().replace(",","")
    try: return (unit_hint if unit_hint else None, float(s))
    except: pass
    if s.endswith(" mm"):
        return ("mm", float(s[:-3].strip()))
    if s.endswith(" m"):
        return ("m", float(s[:-2].strip()))
    tail = re.sub(r"[^\d\.\-\+eE]", "", s)
    if tail not in ("",".","-","+"):
        return (None, float(tail))
    raise ValueError("not_number")

def normalize_to_mm(num, unit):
    if unit == "mm": return num
    if unit == "m":  return num * 1000.0
    return num

def validate(rule, value):
    pname = rule["ParamName"]; typ = rule["Type"]
    allowed = rule["AllowedValues"]; regex = rule["Regex"]
    minv = rule["Min"]; maxv = rule["Max"]

    if not present(value):
        return False, "%s:missing" % pname

    if typ == "number":
        unit_hint = get_unit_hint(rule)
        try:
            unit, num = parse_number_unit_aware(value, unit_hint)
            n_mm = normalize_to_mm(num, unit or unit_hint)
        except:
            return False, "%s:not_number" % pname

        def to_mm(v):
            if not v: return None
            f = float(_s(v).replace(",",""))
            if unit_hint == "mm": return f
            if unit_hint == "m":  return f*1000.0
            return f

        mn = to_mm(minv); mx = to_mm(maxv)
        if mn is not None and n_mm < mn: return False, "%s:lt_min" % pname
        if mx is not None and n_mm > mx: return False, "%s:gt_max" % pname

    if allowed:
        if _s(value) not in [_s(a) for a in allowed.split("|")]:
            return False, "%s:not_allowed" % pname

    if regex:
        try:
            if not re.match(regex, _s(value)):
                return False, "%s:regex_fail" % pname
        except:
            return False, "%s:regex_error" % pname

    return True, ""

def check_stage(params_csv, rules, stage, out_csv):
    with open(params_csv, "r", encoding="utf-8-sig") as f:
        rows = [clean_row(r) for r in csv.DictReader(f)]

    needed = [r for r in rules if r["Stages"].get(stage,0)==1]
    by_cat = {}
    for r in needed: by_cat.setdefault(r["Category"], []).append(r)

    out = []
    for r in rows:
        cat = _s(r.get("Category",""))
        if cat == "" and _s(r.get("ElementId","")) == "":  # skip fully blank keys
            continue
        errs = []
        if cat in by_cat:
            for rule in by_cat[cat]:
                col = rule["ParamName"]
                val = r.get(col, None)
                ok, msg = validate(rule, val)
                if not ok: errs.append(msg)
        out.append({
            "ElementId": _s(r.get("ElementId","")),
            "Category": cat,
            "LOD_Stage": stage,
            "loin_pass": 0 if errs else 1,
            "missing_list": ";".join(errs)
        })

    d = os.path.dirname(out_csv)
    if d and not os.path.exists(d): os.makedirs(d)
    with open(out_csv, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["ElementId","Category","LOD_Stage","loin_pass","missing_list"])
        w.writeheader()
        for r in out: w.writerow(r)
    print("Stage {} → wrote {}".format(stage, out_csv))

# MAIN: 1) Load rules from data/LOIN_rules_by_LOD.csv 2) Load params_export.csv 3) Apply regex/enum/range checks 4) Write loin_Lxxx.csv with Pass/Fail and Missing columns.
if not os.path.exists(RULES_CSV):
    raise RuntimeError("Rules CSV not found: {}".format(RULES_CSV))
if not os.path.exists(PARAMS_OUT):
    raise RuntimeError("Run Step 2 first to create params_export.csv")

rules = load_rules(RULES_CSV)
for s in STAGES:
    out_csv = os.path.join(OUT_DIR, "loin_{}.csv".format(s.replace("LOD","L")))
    check_stage(PARAMS_OUT, rules, s, out_csv)

print("Done.")
