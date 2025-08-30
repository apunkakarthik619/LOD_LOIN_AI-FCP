# Exports geometry features (solids, faces, area, volume) and MEP connectors. Tested in Revit 2026. Handles commas/quotes safely in CSV.

import os, csv
from Autodesk.Revit.DB import (
    FilteredElementCollector, Options, GeometryInstance, Solid,
    LocationCurve, Line, UnitUtils, UnitTypeId, StorageType, MEPCurve,
    BuiltInCategory as BIC
)

NBSP = u"\u00A0"
def _s(x):
    try: return str(x).replace("\ufeff","").replace(NBSP," ").strip()
    except: return x

OUT_CSV = r"C:\LOD_LOIN_AI\data\geom_export.csv"

def id_to_str(eid):
    if not eid: return ""
    for attr in ("Value","IntegerValue"):
        try: return _s(getattr(eid, attr))
        except: pass
    return _s(eid)

def ft_to_m(x):  return UnitUtils.ConvertFromInternalUnits(x, UnitTypeId.Meters)
def sf_to_m2(x): return UnitUtils.ConvertFromInternalUnits(x, UnitTypeId.SquareMeters)
def cf_to_m3(x): return UnitUtils.ConvertFromInternalUnits(x, UnitTypeId.CubicMeters)

def get_bics():
    names = {
        "Walls": "OST_Walls",
        "Floors": "OST_Floors",
        "Ducts": "OST_DuctCurves",
        "Pipes": "OST_PipeCurves",
        "Cable Trays": "OST_CableTray",
    }
    out = {}
    for label, n in names.items():
        if hasattr(BIC, n): out[_s(label)] = getattr(BIC, n)
    return out

def get_solids(el):
    opts = Options(); opts.IncludeNonVisibleObjects = False
    g = el.get_Geometry(opts)
    if not g: return []
    solids = []
    def walk(it):
        for o in it:
            if isinstance(o, Solid) and o.Faces and o.Volume > 1e-9:
                solids.append(o)
            elif isinstance(o, GeometryInstance):
                walk(o.GetInstanceGeometry())
    walk(g); return solids

def get_length_m(el):
    p = el.LookupParameter("Length")
    if p and p.StorageType == StorageType.Double:
        try: return ft_to_m(p.AsDouble())
        except: pass
    loc = el.Location
    if isinstance(loc, LocationCurve):
        try: return ft_to_m(loc.Curve.Length)
        except: pass
    return ""

def get_level_name(el):
    for name in ("Level","Reference Level","Schedule Level","Base Constraint"):
        p = el.LookupParameter(name)
        if not p: continue
        try:
            vs = p.AsValueString()
            if vs: return _s(vs)
        except: pass
        try:
            s = p.AsString()
            if s: return _s(s)
        except: pass
    return ""

doc  = __revit__.ActiveUIDocument.Document
bics = get_bics()
rows = []

for label, bic in bics.items():
    elems = (FilteredElementCollector(doc).OfCategory(bic)
             .WhereElementIsNotElementType().ToElements())
    for el in elems:
        solids = get_solids(el)
        area_m2 = sum(sf_to_m2(s.SurfaceArea) for s in solids) if solids else ""
        vol_m3  = sum(cf_to_m3(s.Volume) for s in solids)      if solids else ""

        sx=sy=sz=ex=ey=ez=""
        loc = el.Location
        if isinstance(loc, LocationCurve) and isinstance(loc.Curve, Line):
            sp = loc.Curve.GetEndPoint(0); ep = loc.Curve.GetEndPoint(1)
            sx, sy, sz = ft_to_m(sp.X), ft_to_m(sp.Y), ft_to_m(sp.Z)
            ex, ey, ez = ft_to_m(ep.X), ft_to_m(ep.Y), ft_to_m(ep.Z)

        # For MEP: export connector count and basic connector data (direction, system type if available). Skip if the category has no connectors.
        try:
            if isinstance(el, MEPCurve): cm = el.ConnectorManager
            else:
                try: cm = el.MEPModel.ConnectorManager
                except: cm = None
            conns = list(cm.Connectors) if cm else []
            ccount = len(conns)
            csizes = []
            for c in conns:
                try:
                    if hasattr(c,"Radius") and c.Radius:
                        csizes.append("D={:.4f}".format(2.0*ft_to_m(c.Radius)))
                    else:
                        w = getattr(c,"Width",0.0) or 0.0
                        h = getattr(c,"Height",0.0) or 0.0
                        if w>0 or h>0:
                            csizes.append("{:.4f}x{:.4f}".format(ft_to_m(w), ft_to_m(h)))
                except: pass
            csizes = ";".join(csizes)
        except:
            ccount, csizes = 0, ""

        rows.append({
            "ElementId": id_to_str(el.Id),
            "Category": _s(label),
            "Type Name": _s(getattr(el.Document.GetElement(el.GetTypeId()), "Name", getattr(el,"Name",""))),
            "Level":     get_level_name(el),
            "Length_m":  get_length_m(el),
            "SurfaceArea_m2": area_m2,
            "Volume_m3": vol_m3,
            "StartX_m": sx, "StartY_m": sy, "StartZ_m": sz,
            "EndX_m":   ex, "EndY_m":   ey, "EndZ_m":   ez,
            "ConnCount": ccount, "ConnSizes_m": _s(csizes)
        })

headers = ["ElementId","Category","Type Name","Level","Length_m","SurfaceArea_m2","Volume_m3",
           "StartX_m","StartY_m","StartZ_m","EndX_m","EndY_m","EndZ_m","ConnCount","ConnSizes_m"]
os.makedirs(os.path.dirname(OUT_CSV), exist_ok=True)
with open(OUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=headers); w.writeheader()
    for r in rows: w.writerow(r)

print("Exported geom:", OUT_CSV)
