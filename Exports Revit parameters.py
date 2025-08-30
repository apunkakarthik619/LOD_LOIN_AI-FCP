# Exports Revit parameters listed in data/LOIN_rules_by_LOD.csv. Tested in Revit 2026. Use pyRevit/RevitPythonShell (IronPython 2.7 or IronPython 3 if available). External helper scripts use CPython 3.10+.
# Exports only the parameters declared in data\LOIN_rules_by_LOD.csv, plus ElementId, Category, Type Name, Level.
# Walls: if Level is missing, read "Base Constraint" as a fallback. Read "Type Name" from the element's type. For parameters missing on the instance, try the type.

import os, csv, re
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory as BIC

# ---------- Paths ----------
BASE_DIR   = r"C:\LOD_LOIN_AI"
RULES_CSV  = os.path.join(BASE_DIR, r"data\LOIN_rules_by_LOD.csv")
OUT_CSV    = os.path.join(BASE_DIR, r"data\params_export.csv")

# ---------- Helpers ----------
NBSP = u"\u00A0"

def _s(x):
    """Normalize strings: strip BOM/NBSP/whitespace; keep non-strings intact."""
    try:
        return str(x).replace("\ufeff", "").replace(NBSP, " ").strip()
    except:
        return x

def clean_row(d):
    """Clean dict keys/values; drop None/blank headers."""
    out = {}
    for k, v in (d or {}).items():
        if k is None:
            continue
        ks = _s(k)
        if ks == "":
            continue
        out[ks] = _s(v)
    return out

def id_to_str(eid):
    if not eid:
        return ""
    for attr in ("Value", "IntegerValue"):
        try:
            return _s(getattr(eid, attr))
        except:
            pass
    return _s(eid)

def param_as_text(p):
    """Return a readable value for any parameter. Prefer AsValueString() for display units, then fallbacks."""
    if not p:
        return ""
    # Prefer AsValueString() for user-facing values (respects project units, e.g., Width "200 mm"). Use AsDouble() only when you will convert units yourself.
    try:
        vs = p.AsValueString()
        if vs is not None:
            return _s(vs)
    except:
        pass
    # Try instance parameter first (element.LookupParameter), then type parameter if not found.
    for g in (p.AsString, p.AsInteger, p.AsDouble):
        try:
            v = g()
            if v is not None:
                return _s(v)
        except:
            pass
    # When a parameter is an ElementId, convert it to its readable string (e.g., referenced type name or the numeric id).
    try:
        return id_to_str(p.AsElementId())
    except:
        return ""

def lookup_param_value(el, p_name):
    """Instance-first, then type-level lookup for parameter p_name."""
    pn = _s(p_name)
    if pn == "":
        return ""
    # Instance
    try:
        p = el.LookupParameter(pn)
        if p:
            return param_as_text(p)
    except:
        pass
    # Type
    try:
        t = el.Document.GetElement(el.GetTypeId())
        if t:
            p = t.LookupParameter(pn)
            if p:
                return param_as_text(p)
    except:
        pass
    return ""

def get_type_name(el):
    """Always fetch the type's Name."""
    try:
        t = el.Document.GetElement(el.GetTypeId())
        if t:
            return _s(getattr(t, "Name", ""))
    except:
        pass
    # If no Type Name is available, fall back to element.Name (often blank for system families like walls).
    return _s(getattr(el, "Name", ""))

def get_level_value(el, category_label):
    """
    Robust Level detection:
    Try Level → Reference Level → Schedule Level → (Walls) Base Constraint.
    """
    for name in ("Level", "Reference Level", "Schedule Level"):
        val = lookup_param_value(el, name)
        if _s(val) != "":
            return _s(val)
    # Walls only: if Level can't be read via the usual property, use "Base Constraint" to capture the level.
    if _s(category_label).lower() == "walls":
        bc = lookup_param_value(el, "Base Constraint")
        if _s(bc) != "":
            return _s(bc)
    return ""

def get_bic(category_label):
    """Map CSV category → BuiltInCategory."""
    label = _s(category_label).lower()
    mapping = {
        "walls": "OST_Walls",
        "floors": "OST_Floors",
        "ducts": "OST_DuctCurves",
        "pipes": "OST_PipeCurves",
        "cable trays": "OST_CableTray",
    }
    name = mapping.get(label)
    if name and hasattr(BIC, name):
        return getattr(BIC, name)
    return None

# ---------- Load rules & build parameter list per category ----------
if not os.path.exists(RULES_CSV):
    raise RuntimeError("Rules CSV not found at: {}".format(RULES_CSV))

cat_params = {}  # { "Walls": set(["Type Mark", "Model", ...]), ... }
with open(RULES_CSV, "r", encoding="utf-8-sig") as f:
    rdr = csv.DictReader(f)
    for row in rdr:
        row = clean_row(row)
        cat = row.get("Category", "")
        pn  = row.get("ParamName", "")
        if cat and pn:
            cat_params.setdefault(cat, set()).add(pn)

# Always include metadata columns: ElementId, Category, Type Name, Level (first four columns).
for c in list(cat_params.keys()):
    cat_params[c].update(["Type Mark", "Level"])  # safe to ask even if not in rules

# ---------- Collect & export ----------
doc = __revit__.ActiveUIDocument.Document
rows = []
counts = {}  # category → element count

for cat in sorted(cat_params.keys()):
    bic = get_bic(cat)
    if not bic:
        print("Skipping category (no BIC mapping):", cat)
        continue

    elems = (FilteredElementCollector(doc)
             .OfCategory(bic)
             .WhereElementIsNotElementType()
             .ToElements())
    counts[cat] = len(elems)

    # Build the final export list: fixed metadata columns first, then the rule-driven parameter names in the order they appear in the CSV.
    # fixed order: ElementId, Category, Type Name, Level, then all rule params (stable)
    rule_param_list = sorted(list(cat_params[cat]))  # alphabetical for stability

    for el in elems:
        rec = {
            "ElementId": id_to_str(el.Id),
            "Category":  _s(cat),
            "Type Name": get_type_name(el),              # from TYPE (important for Walls)
            "Level":     get_level_value(el, cat),       # Walls → fall back to Base Constraint
        }
        # Fill requested parameters
        for p in rule_param_list:
            rec[p] = lookup_param_value(el, p)
        rows.append(clean_row(rec))

# Write headers: keep the four metadata columns first; then add the rule parameters in the same order as the rules file (do NOT sort alphabetically).
fixed_first = ["ElementId", "Category", "Type Name", "Level"]
param_headers = set()
for r in rows:
    for k in r.keys():
        if k not in fixed_first:
            param_headers.add(k)
param_headers = sorted(list(param_headers))
headers = fixed_first + param_headers

# Create outputs/ if it doesn’t exist to avoid file write errors.
out_dir = os.path.dirname(OUT_CSV)
if out_dir and not os.path.exists(out_dir):
    os.makedirs(out_dir)

# Write CSV (Excel-friendly; quotes where needed)
with open(OUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=headers, quoting=csv.QUOTE_MINIMAL)
    w.writeheader()
    for r in rows:
        # guarantee all headers exist
        for h in headers:
            r.setdefault(h, "")
        w.writerow(r)

print("Exported params:", OUT_CSV)
for k in sorted(counts.keys()):
    print("{:<12} {}".format(_s(k) + ":", counts[k]))
