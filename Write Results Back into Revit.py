# RPS_Write_Results_To_Revit_DEBUG_CLEAN_fixed.py
import os, csv, re
import Autodesk.Revit.DB as DB
import Autodesk.Revit.UI as UI

CSV_PATH = r"C:\LOD_LOIN_AI\outputs\results_for_revit.csv"

uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document

def parse_yes(v):
    s = str(v).strip().lower()
    return 1 if s in ("1","true","yes","y","pass","compliant") else 0

def parse_float(v, default=0.0):
    try: return float(str(v).strip())
    except: return default

def set_param(el, name, value):
    p = el.LookupParameter(name)
    if not p: return False, "param-not-found"
    st = p.StorageType
    try:
        if st == DB.StorageType.String:
            p.Set("" if value is None else str(value)); return True, None
        elif st == DB.StorageType.Integer:  # Yes/No
            p.Set(parse_yes(value)); return True, None
        elif st == DB.StorageType.Double:   # Number
            p.Set(parse_float(value, 0.0)); return True, None
        else:
            return False, "unsupported-storage"
    except Exception as ex:
        return False, "exception: {}".format(ex)

# Clean up: remove empty rows, normalize headers, and validate required columns before writing back to Revit.
BOMs = u"\ufeff"
ZWS  = u"\u200b\u200c\u200d\u2060"
NBSP = u"\u00a0"

def clean_text(s):
    if s is None: return ""
    s = str(s)
    for ch in (BOMs, ZWS, NBSP):
        s = s.replace(ch, "")
    return s.strip()

def clean_elid(s):
    s = clean_text(s)
    s = re.sub(r"[^\d]", "", s)
    return int(s) if s else None

# Open outputs/results_for_revit.csv. This must include ElementId, Category, and the target result columns.
if not os.path.exists(CSV_PATH):
    UI.TaskDialog.Show("Push Results", "CSV not found:\n{}".format(CSV_PATH)); raise SystemExit

with open(CSV_PATH, "r", encoding="utf-8-sig", newline="") as f:
    rdr = csv.reader(f)
    rows = list(rdr)

if not rows or len(rows) <= 1:
    UI.TaskDialog.Show("Push Results", "CSV empty or header-only:\n{}".format(CSV_PATH)); raise SystemExit

header = [clean_text(h) for h in rows[0]]
data   = rows[1:]

def idx(name_variants):
    nv = [n.lower().replace(" ","").replace("_","") for n in name_variants]
    for i, h in enumerate(header):
        hh = h.lower().replace(" ","").replace("_","")
        if hh in nv: return i
    return None

i_el  = idx(["ElementId","Element ID","elementid"])
i_cat = idx(["Category","categoryname"])
i_stg = idx(["LOD_Stage","LODStage","loin_stage"])
i_pss = idx(["loin_pass","loinpass"])
i_mis = idx(["missing_list","missing","loin_missing"])
i_scr = idx(["lod_score","lodscore"])
i_sta = idx(["final_status","aifinalstatus","status"])

if i_el is None or i_cat is None:
    UI.TaskDialog.Show("Push Results", "Required columns missing.\nHeader:\n{}".format(header)); raise SystemExit

# Use a Revit Transaction to write parameter values safely. If any row fails, the transaction rolls back or logs the error without corrupting the model.
if getattr(doc, "IsReadOnly", False):
    UI.TaskDialog.Show("Push Results", "Doc is read-only. Obtain write access and re-run."); raise SystemExit

ok=fail=miss=skipped=0; fail_by_param={}
started_here=False
tx=DB.Transaction(doc,"Write LOIN/AI results (clean)")

try:
    if not doc.IsModifiable:
        tx.Start(); started_here=True

    for r in data:
        elid = clean_elid(r[i_el] if i_el is not None and i_el < len(r) else "")
        if not elid:
            skipped += 1; continue
        el = doc.GetElement(DB.ElementId(elid))
        if not el:
            miss += 1; continue

        def val(idx, default=""):
            if idx is None or idx >= len(r): return default
            return clean_text(r[idx])

        writes = [
            ("LOIN_Stage",     val(i_stg, "")),
            ("LOIN_Pass",      val(i_pss, "")),
            ("LOIN_Missing",   val(i_mis, "")),
            ("LOD_Score",      val(i_scr, "")),
            ("AI_FinalStatus", val(i_sta, "")),
        ]
        for pname, v in writes:
            ok1, err = set_param(el, pname, v)
            if ok1: ok += 1
            else:
                fail += 1; fail_by_param[pname] = fail_by_param.get(pname, 0) + 1

    if started_here: tx.Commit()
except Exception as ex:
    if started_here: tx.RollBack()
    UI.TaskDialog.Show("Push Results", "Error during write:\n{}".format(ex)); raise

msg = (
    "Rows read: {}\n"
    "Skipped (bad ElementId): {}\n"
    "Missing elements: {}\n"
    "OK writes: {}\n"
    "Failed writes: {}\n"
).format(len(data), skipped, miss, ok, fail)
if fail_by_param:
    msg += "\nFailures by parameter:\n" + "\n".join("  {}: {}".format(k,v) for k,v in fail_by_param.items())
UI.TaskDialog.Show("Push Results", msg); print(msg)
