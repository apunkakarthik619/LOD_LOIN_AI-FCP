# RPS_Merge_Results_For_Revit.py  (run inside RevitPythonShell)
import os, csv

BASE        = r"C:\LOD_LOIN_AI"
OUTPUTS     = os.path.join(BASE, "outputs")

# Change if your file names differ
LOIN_CSV    = os.path.join(OUTPUTS, "loin_L300.csv")        # your rule-based output
AI_CSV      = os.path.join(OUTPUTS, "results_ai.csv")       # Step 7 output
OUT_CSV     = os.path.join(OUTPUTS, "results_for_revit.csv")

def key(row):
    return (row.get("ElementId","").strip(), row.get("Category","").strip())

if not (os.path.exists(LOIN_CSV) and os.path.exists(AI_CSV)):
    raise SystemExit("Missing input CSVs. Check paths:\n{}\n{}".format(LOIN_CSV, AI_CSV))

# Load loin_Lxxx.csv (rule results) keyed by ElementId (and FileName if present).
loin_map = {}
with open(LOIN_CSV, "r", encoding="utf-8-sig") as f:
    r = csv.DictReader(f)
    for row in r:
        loin_map[key(row)] = row

# After prediction, results_ai.csv can be merged with LOIN output to prepare a single file for writing back to Revit.
merged_rows = []
with open(AI_CSV, "r", encoding="utf-8-sig") as f:
    r = csv.DictReader(f)
    for ai in r:
        k = key(ai)
        base = loin_map.get(k, {})  # may be empty if not in LOIN file
        # For overlapping fields, prefer explicit LOIN values; use AI values only when LOIN value is missing.
        row = {}
        row["ElementId"]  = ai.get("ElementId","").strip() or base.get("ElementId","")
        row["Category"]   = ai.get("Category","").strip()  or base.get("Category","")
        row["LOD_Stage"]  = base.get("LOD_Stage","") or ai.get("LOD_Stage","")
        # For "LOIN_Pass": prefer the deterministic LOIN Pass/Fail result. Do not override with AI.
        row["loin_pass"]  = (ai.get("loin_pass","") or base.get("loin_pass","") or "1").strip()
        # For "LOIN_Missing": take from LOIN (authoritative list of missing params).
        row["missing_list"] = base.get("missing_list","") or ai.get("missing_list","")
        # Keep AI-only fields (probability, recommended action) as additional context columns.
        row["lod_score"]    = ai.get("lod_score","")
        row["final_status"] = ai.get("final_status","")
        row["checked_on"]   = ai.get("checked_on","")
        merged_rows.append(row)

# Output order expected by the writer script: metadata → LOIN_* fields → AI_* fields.
cols = ["ElementId","Category","LOD_Stage","loin_pass","missing_list","lod_score","final_status","checked_on"]

with open(OUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=cols)
    w.writeheader()
    for row in merged_rows:
        w.writerow({c: row.get(c,"") for c in cols})

print("Wrote:", OUT_CSV, "rows:", len(merged_rows))
