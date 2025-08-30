# Launches external CPython (3.10+) to score merged_Lxxx.csv with model.joblib and writes results_ai.csv. (run inside RevitPythonShell)
# Scores a merged_Lxxx.csv using the trained model, outputs results_ai.csv.
# - Auto-detects python.exe (or set PY_CMD_OVERRIDE).
# - Installs pandas, scikit-learn, joblib if needed.
# - Writes predict_lod_ai.py (if missing), then runs it.

import os, sys, subprocess, codecs

# ---------- Revit popup helper ----------
try:
    from Autodesk.Revit.UI import TaskDialog
    def popup(title, msg): TaskDialog.Show(title, msg)
except:
    def popup(title, msg):
        print("=== {} ===\n{}".format(title, msg))

# -------------------- CONFIG --------------------
BASE       = r"C:\LOD_LOIN_AI"
SCRIPTS    = os.path.join(BASE, "scripts")
OUTPUTS    = os.path.join(BASE, "outputs")
PREDICT_PY = os.path.join(SCRIPTS, "predict_lod_ai.py")

# Change these if your filenames differ:
MERGED_CSV = os.path.join(OUTPUTS, "merged_L300.csv")       # the package you want to score
MODEL_PATH = os.path.join(OUTPUTS, "lod_model.joblib")      # trained model from Step 6
OUT_CSV    = os.path.join(OUTPUTS, "results_ai.csv")        # output file
THRESHOLD  = 0.75                                            # adjust if you want stricter/looser triage

# If you want to hard-set python.exe, put the full path here; otherwise leave None for auto-detect.
# Example: r"C:\Users\YOU\AppData\Local\Programs\Python\Python312\python.exe"
PY_CMD_OVERRIDE = None
# -----------------------------------------------

def run(cmd):
    p = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, shell=True)
    out = p.communicate()[0]
    try: out = out.decode("utf-8", "ignore")
    except: pass
    return p.returncode, out

def looks_like_python_exe(path):
    return bool(path) and path.lower().endswith("python.exe") and os.path.exists(path)

def autodetect_python():
    if PY_CMD_OVERRIDE and looks_like_python_exe(PY_CMD_OVERRIDE):
        return PY_CMD_OVERRIDE
    # Try PATH
    try:
        rc, out = run("where python")
        if rc == 0 and out:
            for line in out.splitlines():
                line = line.strip()
                if looks_like_python_exe(line): return line
    except: pass
    # Common locations
    guesses = []
    user_base = os.path.expandvars(r"%LOCALAPPDATA%\Programs\Python")
    if os.path.isdir(user_base):
        for name in sorted(os.listdir(user_base), reverse=True):
            guesses.append(os.path.join(user_base, name, "python.exe"))
    guesses += [
        r"C:\Python312\python.exe", r"C:\Python311\python.exe", r"C:\Python310\python.exe",
        r"C:\Program Files\Python312\python.exe", r"C:\Program Files\Python311\python.exe",
        r"C:\Program Files\Python310\python.exe",
        r"C:\Program Files (x86)\Python312\python.exe", r"C:\Program Files (x86)\Python311\python.exe",
        r"C:\Program Files (x86)\Python310\python.exe",
    ]
    for g in guesses:
        if looks_like_python_exe(g): return g
    return None

# 0) Ensure folders
for d in (BASE, SCRIPTS, OUTPUTS):
    if not os.path.exists(d): os.makedirs(d)

# 1) Write predict_lod_ai.py (your provided scorer)
PREDICT_CODE = u'''import os, re, joblib, argparse
import pandas as pd
from datetime import datetime

NBSP = u"\\u00A0"

def _s(x):
    try: return str(x).replace("\\ufeff","").replace(NBSP, " ").strip()
    except: return x

def read_csv(path):
    return pd.read_csv(path, dtype=str, keep_default_na=False, na_values=[], encoding="utf-8-sig")

def to_float_mm(v):
    s = _s(v).lower().replace(",", "")
    if s == "" or s == "nan": return None
    try: return float(s)
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

def build_features(df, meta):
    cat_cols = meta.get("cat_cols", [])
    num_cols = meta.get("num_cols", [])

    parts = []

    # categorical
    if cat_cols:
        if set(cat_cols).issubset(df.columns):
            cat_df = df[cat_cols].applymap(_s)
        else:
            cat_df = pd.DataFrame({c: df.get(c, "") for c in cat_cols}).applymap(_s)
        parts.append(cat_df)

    # numeric (rebuild from raw if needed)
    num_df = pd.DataFrame(index=df.index)
    for col in num_cols:
        if col in df.columns:
            if col.endswith("_mm"):
                base = col[:-3]
                num_df[col] = df[col] if df[col].dtype != object else df[col].map(to_float).fillna(0.0)
            else:
                num_df[col] = df[col].map(to_float).fillna(0.0)
        else:
            if col.endswith("_mm"):
                base = col[:-3]
                if base in df.columns:
                    num_df[col] = df[base].map(to_float_mm).fillna(0.0)
                else:
                    num_df[col] = 0.0
            else:
                num_df[col] = 0.0
    if not num_df.empty: parts.append(num_df)

    X = pd.concat(parts, axis=1) if parts else pd.DataFrame(index=df.index)
    return X

def predict(merged_csv, model_path, out_csv, threshold=0.75):
    if not os.path.exists(merged_csv): raise FileNotFoundError(merged_csv)
    if not os.path.exists(model_path): raise FileNotFoundError(model_path)

    pack  = joblib.load(model_path)
    model = pack["model"]
    meta  = pack.get("meta", {})

    df = read_csv(merged_csv)
    X  = build_features(df, meta)
    if X.empty:
        raise RuntimeError("No features available to score; check merged CSV and training meta.")

    proba = model.predict_proba(X)[:,1]
    df["lod_score"] = proba

    if "loin_pass" not in df.columns:
        df["loin_pass"] = 1

    def decide(row):
        try:
            lp = int(str(row.get("loin_pass","1")))
        except:
            lp = 1
        if lp == 0:
            return "Non-compliant"
        return "Compliant" if row["lod_score"] >= threshold else "Needs Review"

    df["final_status"] = df.apply(decide, axis=1)
    df["checked_on"]  = datetime.now().strftime("%Y-%m-%d %H:%M")

    cols = ["ElementId","Category","loin_pass","lod_score","final_status","checked_on"]
    if "missing_list" in df.columns: cols.insert(3, "missing_list")
    if "LOD_Stage" in df.columns:    cols.insert(2, "LOD_Stage")

    out = df[cols]
    out.to_csv(out_csv, index=False, encoding="utf-8-sig")
    print("Wrote", out_csv, "rows:", len(out))

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--merged_csv", required=True)
    ap.add_argument("--model_path", required=True)
    ap.add_argument("--out_csv", required=True)
    ap.add_argument("--threshold", type=float, default=0.75)
    args = ap.parse_args()
    predict(args.merged_csv, args.model_path, args.out_csv, args.threshold)
'''

if not os.path.exists(PREDICT_PY):
    with codecs.open(PREDICT_PY, "w", "utf-8") as f:
        f.write(PREDICT_CODE)
    print("Wrote scorer to:", PREDICT_PY)
else:
    print("Scorer already present at:", PREDICT_PY)

# 2) Find python.exe
PY_CMD = autodetect_python()
if not PY_CMD:
    popup("Predict LOD AI — Python not found",
          "Could not find python.exe.\nSet PY_CMD_OVERRIDE at the top of the script to your Python path.")
    raise SystemExit
print("Using Python:", PY_CMD)

# 3) Ensure packages (scikit-learn needed to load the pipeline)
print("\nInstalling/validating Python packages (user scope)...\n")
for pkg in ("pandas", "scikit-learn", "joblib"):
    rc, out = run(r'"{}" -m pip install --user {}'.format(PY_CMD, pkg))
    print(out)

# 4) Sanity checks
missing = []
if not os.path.exists(MERGED_CSV): missing.append(MERGED_CSV)
if not os.path.exists(MODEL_PATH): missing.append(MODEL_PATH)
if missing:
    popup("Predict LOD AI — Missing files", "Not found:\n" + "\n".join(missing))
    raise SystemExit

# 5) Run predictor
cmd = r'"{}" "{}" --merged_csv "{}" --model_path "{}" --out_csv "{}" --threshold {}'.format(
    PY_CMD, PREDICT_PY, MERGED_CSV, MODEL_PATH, OUT_CSV, THRESHOLD)
print("Running:\n" + cmd + "\n")
rc, out = run(cmd)
print(out)

# 6) Notify
if rc == 0 and os.path.exists(OUT_CSV):
    try:
        # Count rows (optional)
        import csv
        with open(OUT_CSV, "r", encoding="utf-8-sig") as f:
            n = sum(1 for _ in csv.reader(f)) - 1
        popup("Predict LOD AI", "Done.\nSaved: {}\nRows: {}".format(OUT_CSV, max(n,0)))
    except:
        popup("Predict LOD AI", "Done.\nSaved: {}".format(OUT_CSV))
else:
    popup("Predict LOD AI — Error", out or "Unknown error")
