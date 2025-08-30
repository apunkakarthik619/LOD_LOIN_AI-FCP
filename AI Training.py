# Launches external CPython (3.10+) to train the AI model on merged_Lxxx.csv and saves model.joblib. Run from RevitPythonShell/pyRevit.
# Trains your LOD/LOIN acceptance model by invoking your system Python directly (no 'py -3' needed).
# - Auto-detects python.exe; or set PY_CMD_OVERRIDE manually.
# - Installs pandas, scikit-learn, joblib (user scope) if needed.
# - Writes train_lod_ai.py to C:\LOD_LOIN_AI\scripts (if missing).
# - Runs training and shows AUC + success popup.

import os, sys, subprocess, codecs, re

# ---------- Revit popup helper ----------
try:
    from Autodesk.Revit.UI import TaskDialog
    def popup(title, msg): TaskDialog.Show(title, msg)
except:
    def popup(title, msg):
        print("=== {} ===\n{}".format(title, msg))

# -------------------- CONFIG --------------------
BASE     = r"C:\LOD_LOIN_AI"
SCRIPTS  = os.path.join(BASE, "scripts")
OUTPUTS  = os.path.join(BASE, "outputs")
TRAIN_PY = os.path.join(SCRIPTS, "train_lod_ai.py")

# Change if your file names differ.
MERGED_CSV = os.path.join(OUTPUTS, "merged_L300.csv")
MODEL_OUT  = os.path.join(OUTPUTS, "lod_model.joblib")

# If you already know your python.exe, set it here and skip auto-detect, e.g.:
# PY_CMD_OVERRIDE = r"C:\Users\YOU\AppData\Local\Programs\Python\Python312\python.exe"
PY_CMD_OVERRIDE = None
# -----------------------------------------------

def run(cmd):
    """Run a shell command, capture stdout+stderr."""
    p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, shell=True)
    out = p.communicate()[0]
    try:
        out = out.decode("utf-8", "ignore")
    except:
        pass
    return p.returncode, out

def looks_like_python_exe(path):
    return bool(path) and path.lower().endswith("python.exe") and os.path.exists(path)

def autodetect_python():
    # 1) Respect manual override
    if PY_CMD_OVERRIDE and looks_like_python_exe(PY_CMD_OVERRIDE):
        return PY_CMD_OVERRIDE

    # 2) Try PATH via `where python`
    try:
        rc, out = run("where python")
        if rc == 0 and out:
            for line in out.splitlines():
                line = line.strip()
                if looks_like_python_exe(line):
                    return line
    except:
        pass

    # 3) Common install locations (newest first)
    guesses = []
    # User-local installs
    user_base = os.path.expandvars(r"%LOCALAPPDATA%\Programs\Python")
    if os.path.isdir(user_base):
        for name in sorted(os.listdir(user_base), reverse=True):
            guesses.append(os.path.join(user_base, name, "python.exe"))
    # System-wide
    guesses += [
        r"C:\Python312\python.exe",
        r"C:\Python311\python.exe",
        r"C:\Python310\python.exe",
        r"C:\Program Files\Python312\python.exe",
        r"C:\Program Files\Python311\python.exe",
        r"C:\Program Files\Python310\python.exe",
        r"C:\Program Files (x86)\Python312\python.exe",
        r"C:\Program Files (x86)\Python311\python.exe",
        r"C:\Program Files (x86)\Python310\python.exe",
    ]
    for g in guesses:
        if looks_like_python_exe(g):
            return g

    return None

# 0) Ensure folders
for d in (BASE, SCRIPTS, OUTPUTS):
    if not os.path.exists(d):
        os.makedirs(d)

# 1) Write trainer file (exactly your provided script)
TRAIN_CODE = u'''import os, re, joblib, argparse
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import OneHotEncoder
from sklearn.pipeline import Pipeline
from sklearn.ensemble import GradientBoostingClassifier
from sklearn.metrics import roc_auc_score

NBSP = u"\\u00A0"

def _s(x):
    try: return str(x).replace("\\ufeff","").replace(NBSP, " ").strip()
    except: return x

def read_csv(path):
    return pd.read_csv(path, dtype=str, keep_default_na=False, na_values=[], encoding="utf-8-sig")

def to_float_mm(v):
    s = _s(v).lower().replace(",", "")
    if s == "" or s == "nan": return None
    try:
        return float(s)  # assume already mm-like if no unit
    except: pass
    if s.endswith(" mm"):
        try: return float(s[:-3].strip())
        except: return None
    if s.endswith(" m"):
        try: return float(s[:-2].strip()) * 1000.0
        except: return None
    tail = re.sub(r"[^\d\\.\\-\\+eE]", "", s)
    try: return float(tail)
    except: return None

def to_float(v):
    s = _s(v).lower().replace(",", "")
    if s == "" or s == "nan": return None
    try: return float(s)
    except:
        tail = re.sub(r"[^\d\\.\\-\\+eE]", "", s)
        try: return float(tail)
        except: return None

def build_features(df):
    num_map_generic = {
        "Length_m": to_float,
        "SurfaceArea_m2": to_float,
        "Volume_m3": to_float,
        "ConnCount": to_float,
    }
    flag_cols = [c for c in df.columns if isinstance(c, str) and c.startswith("is_missing_")]

    mm_pref = []
    for name in ["Width","Default Thickness","Insulation Thickness","Lining Thickness"]:
        if name in df.columns: mm_pref.append(name)

    num_data = {}
    for col, fn in num_map_generic.items():
        if col in df.columns:
            num_data[col] = df[col].map(fn).fillna(0.0)

    for col in mm_pref:
        num_data[col + "_mm"] = df[col].map(to_float_mm).fillna(0.0)

    for col in flag_cols:
        num_data[col] = df[col].astype(str).str.upper().isin(["TRUE","1","YES"]).astype(int)

    import pandas as pd
    num_df = pd.DataFrame(num_data) if num_data else pd.DataFrame(index=df.index)

    cat_candidates = [
        "Category","System Type","Material","Assembly Code","Type Mark","Level",
        "Service Type","Model","Structural","Fire Rating","Description"
    ]
    cat_cols = [c for c in cat_candidates if c in df.columns]
    cat_df = df[cat_cols].applymap(_s) if cat_cols else pd.DataFrame(index=df.index)

    X = pd.concat([cat_df, num_df], axis=1)
    return X, cat_cols, list(num_data.keys())

def normalize_labels(df):
    lab = df["ApprovedLabel"].astype(str).str.lower().map(_s)
    y = lab.isin(["pass","approved","yes","true","1"]).astype(int)
    return y

def train(merged_csv, model_out):
    if not os.path.exists(merged_csv):
        raise FileNotFoundError("merged_csv not found: {}".format(merged_csv))
    df = read_csv(merged_csv)
    if "ApprovedLabel" not in df.columns:
        raise RuntimeError("ApprovedLabel column missing in {}".format(merged_csv))
    df = df[df["ApprovedLabel"].astype(str).str.strip()!=""].copy()
    if df.empty:
        raise RuntimeError("No labeled rows found (ApprovedLabel empty).")
    y = normalize_labels(df)
    X, cat_cols, num_cols = build_features(df)

    pre = ColumnTransformer(
        transformers=[("onehot", OneHotEncoder(handle_unknown="ignore"), cat_cols)],
        remainder="passthrough"
    )
    pipe = Pipeline([
        ("pre", pre),
        ("clf", GradientBoostingClassifier(random_state=0))
    ])

    Xtr, Xval, ytr, yval = train_test_split(X, y, test_size=0.30, random_state=0, stratify=y)
    pipe.fit(Xtr, ytr)
    proba = pipe.predict_proba(Xval)[:,1]
    auc = roc_auc_score(yval, proba)
    print("Validation AUC:", round(auc, 3))

    meta = { "cat_cols": cat_cols, "num_cols": num_cols }
    joblib.dump({"model": pipe, "meta": meta}, model_out)
    print("Saved model to", model_out)

if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--merged_csv", required=True)
    ap.add_argument("--model_out", required=True)
    args = ap.parse_args()
    train(args.merged_csv, args.model_out)
'''

if not os.path.exists(TRAIN_PY):
    with codecs.open(TRAIN_PY, "w", "utf-8") as f:
        f.write(TRAIN_CODE)
    print("Wrote trainer to:", TRAIN_PY)
else:
    print("Trainer already present at:", TRAIN_PY)

# 2) Find python.exe
PY_CMD = autodetect_python()
if not PY_CMD:
    popup("Train LOD AI — Python not found",
          "Could not find python.exe.\n\n"
          "Fix: Install Python 3.x from python.org and/or set PY_CMD_OVERRIDE at the top of this script.")
    raise SystemExit

print("Using Python:", PY_CMD)

# 3) Install/validate packages
print("\nInstalling/validating Python packages (user scope)...\n")
for pkg in ("pandas", "scikit-learn", "joblib"):
    rc, out = run(r'"{}" -m pip install --user {}'.format(PY_CMD, pkg))
    print(out)

# 4) Sanity checks
if not os.path.exists(MERGED_CSV):
    popup("Train LOD AI", "Missing merged CSV:\n{}".format(MERGED_CSV))
    raise SystemExit

# 5) Train
cmd = r'"{}" "{}" --merged_csv "{}" --model_out "{}"'.format(PY_CMD, TRAIN_PY, MERGED_CSV, MODEL_OUT)
print("Running:\n" + cmd + "\n")
rc, out = run(cmd)
print(out)

# 6) Notify
if rc == 0 and os.path.exists(MODEL_OUT):
    popup("Train LOD AI", "Training finished.\nModel saved:\n{}".format(MODEL_OUT))
else:
    popup("Train LOD AI — Error", out or "Unknown error")
