#!/usr/bin/env python3
"""Convert raw battery data to cleaned CSVs and prepare them for benchmarking.

Raw data locations (relative to this file):
  ../data/nasa/       — NASA .mat files (38 files in subfolders)
  ../data/calce/      — CALCE .zip files (CS2_33, CS2_34, CS2_36)
  ../data/oxford/     — Oxford .mat file (or .aa/.ab chunks)

Output:
  ../data/nasa_clean_filtered.csv
  ../data/calce_clean.csv
  ../data/oxford_clean.csv
"""

import os, sys, subprocess, glob, importlib.util
import numpy as np
import pandas as pd

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(ROOT, "..", "data")

FEATURES = ["cycle", "avg_voltage", "min_voltage", "avg_current", "avg_temp", "duration", "SOH"]
REQUIRED_COLS = ["cycle", "SOH", "cell", "RUL"]

def clean_df(df):
    cols_available = [c for c in FEATURES if c in df.columns]
    df = df.replace([np.inf, -np.inf], np.nan)
    df = df.dropna(subset=REQUIRED_COLS).copy()
    df = df[(df["SOH"] > 0) & (df["SOH"] < 1.2)].copy()
    df = df[df["RUL"] >= 0].copy()
    df = df.sort_values(["cell", "cycle"]).copy()
    df[cols_available] = df[cols_available].fillna(0)
    return df

def run_loader(script, output_csv, description, filter_fn=False):
    out_path = os.path.join(DATA, output_csv)
    if os.path.exists(out_path):
        size = os.path.getsize(out_path)
        print(f"[SKIP] {description} — {output_csv} exists ({size:,} bytes)")
        return True

    loader_out = output_csv.replace("_filtered", "")
    loader_out_path = os.path.join(DATA, loader_out)

    print(f"[RUN]  {description} → {loader_out}")
    result = subprocess.run(
        [sys.executable, script],
        cwd=ROOT, capture_output=True, text=True
    )
    if result.returncode != 0:
        print(f"[FAIL] {description}")
        for line in result.stderr.splitlines()[-5:]:
            print(f"       {line}")
        return False

    if filter_fn and os.path.exists(loader_out_path):
        print(f"[FILT] Applying clean_df() to {loader_out} → {output_csv}")
        df = pd.read_csv(loader_out_path)
        df = clean_df(df)
        df.to_csv(out_path, index=False)
        print(f"       {len(df)} rows, {df['cell'].nunique()} cells")
    print(f"[OK]   {description}")
    return True

def ensure_oxford_mat():
    mat_path = os.path.join(DATA, "oxford", "Oxford_Battery_Degradation_Dataset_1.mat")
    if os.path.exists(mat_path) and os.path.getsize(mat_path) > 1e8:
        return True

    chunks = sorted(glob.glob(os.path.join(DATA, "oxford", "*.mat.*")))
    if not chunks:
        print("[WARN] Oxford: no full .mat file and no split chunks found.")
        print("       Download from: https://ora.ox.ac.uk/objects/uuid:03ba4b01-cfed-46b3-9ab8-b3534433d6b8")
        print(f"       Save as: {mat_path}")
        print("       Or split into <95MB chunks named Oxford_Battery_Degradation_Dataset_1.mat.aa,.ab")
        return False

    print("[INFO] Recombining Oxford .mat from split chunks...")
    try:
        with open(mat_path, "wb") as out:
            for chunk in chunks:
                with open(chunk, "rb") as f:
                    out.write(f.read())
        print(f"[OK]   {os.path.getsize(mat_path):,} bytes")
        return True
    except Exception as e:
        print(f"[FAIL] {e}")
        return False

def check_calce():
    calce_dir = os.path.join(DATA, "calce")
    missing = []
    for name in ["CS2_33.zip", "CS2_34.zip", "CS2_35.zip", "CS2_36.zip",
                  "CX2_36.zip", "CX2_37.zip", "CX2_38.zip"]:
        path = os.path.join(calce_dir, name)
        if not os.path.exists(path) or os.path.getsize(path) == 0:
            missing.append(name)
    if missing:
        print(f"[WARN] CALCE: {len(missing)} file(s) missing/corrupted: {', '.join(missing[:3])}...")
        print("       Cells with missing raw data will be skipped. Existing calce_clean.csv has all 7.")
    return len(missing) == 0

def main():
    os.makedirs(DATA, exist_ok=True)

    check_calce()
    ensure_oxford_mat()

    pipelines = [
        ("loader.py",        "nasa_clean_filtered.csv", "NASA",   True),
        ("loader_calce.py",  "calce_clean.csv",         "CALCE",  False),
        ("loader_oxford.py", "oxford_clean.csv",        "Oxford", False),
    ]
    successes = 0
    for script, csv_name, desc, filter_fn in pipelines:
        if run_loader(script, csv_name, desc, filter_fn=filter_fn):
            successes += 1

    print(f"\n{'='*50}")
    print(f"Done: {successes}/{len(pipelines)} datasets ready.")
    if successes < len(pipelines):
        print("Some datasets could not be reprocessed. Existing cleaned CSVs in data/ may still be usable.")
    print(f"Run 'python3 benchmark_cv.py' next to generate results.")

if __name__ == "__main__":
    main()
