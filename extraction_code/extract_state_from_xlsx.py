import os
import pandas as pd
import numpy as np

# ---- settings ----
SRC_XLSX = "short_term_arrivals_state_of_stay.xlsx"   # 
OUT_CSV  = "outputs/visiting_visas_by_state.csv"

os.makedirs("outputs", exist_ok=True)

# 1) Open workbook and list sheets
xls = pd.ExcelFile(SRC_XLSX, engine="openpyxl")

def looks_like_data(df: pd.DataFrame) -> bool:
    """Heuristic: first col mostly dates, and >=6 numeric columns."""
    if df.empty:
        return False
    # 
    dt = pd.to_datetime(df.iloc[:, 0], errors="coerce")
    date_ratio = dt.notna().mean()
    # 
    num_cols = sum(pd.api.types.is_numeric_dtype(df[c]) for c in df.columns[1:])
    return (date_ratio > 0.6) and (num_cols >= 6)

candidate = None
candidate_name = None

# 2) Try each sheet with a few header guesses
for sh in xls.sheet_names:
    # 
    tmp = pd.read_excel(xls, sheet_name=sh, header=None)
    # 
    # 
    found = False
    for hdr_row in range(0, min(25, len(tmp))):
        df = pd.read_excel(xls, sheet_name=sh, header=hdr_row)
        # drop
        df = df.loc[:, ~df.columns.astype(str).str.contains("^Unnamed")]
        if looks_like_data(df):
            candidate = df
            candidate_name = sh
            found = True
            break
    if found:
        break

if candidate is None:
    raise RuntimeError("Could not find a data-like sheet. Open the xlsx manually and tell me its sheet name and header row.")

# 3) Rename columns: first column is date, others are states (wide format)
wide = candidate.copy()
# 
first_col = wide.columns[0]
# 
keep_cols = [first_col] + [c for c in wide.columns[1:] if pd.api.types.is_numeric_dtype(wide[c])]
wide = wide[keep_cols]

# 
wide[first_col] = pd.to_datetime(wide[first_col], errors="coerce")
wide = wide.dropna(subset=[first_col])

# 4) Melt to long format: date, state, arrivals
long_df = wide.melt(id_vars=[first_col], var_name="state", value_name="arrivals")
long_df = long_df.dropna(subset=["arrivals"])
long_df = long_df[long_df["arrivals"] >= 0]

# 5) Standardize date to YYYY-MM and clean state names
long_df["year_month"] = long_df[first_col].dt.to_period("M").astype(str)
state_map = {
    "New South Wales": "NSW", "NSW": "NSW",
    "Victoria": "VIC", "VIC": "VIC",
    "Queensland": "QLD", "QLD": "QLD",
    "South Australia": "SA", "SA": "SA",
    "Western Australia": "WA", "WA": "WA",
    "Tasmania": "TAS", "TAS": "TAS",
    "Northern Territory": "NT", "NT": "NT",
    "Australian Capital Territory": "ACT", "ACT": "ACT",
    "Other Territories": "OT", "Australia": "TOTAL",
}
long_df["state"] = long_df["state"].astype(str).str.strip().replace(state_map)

# 
long_df = long_df[~long_df["state"].isin(["TOTAL", "OT"])]

# 6) Keep tidy columns and export
final = long_df[["year_month", "state", "arrivals"]].sort_values(["year_month", "state"])

# 
final = long_df[["year_month", "state", "arrivals"]].sort_values(["year_month", "state"])


series_to_state = {
    "A85247916X": "NSW",
    "A85247923W": "VIC",
    "A85247917A": "QLD",
    "A85247924X": "SA",
    "A85247921T": "WA",
    "A85247918C": "TAS",
    "A85247920R": "NT",
    "A85247919F": "ACT",
    "A85247922V": "OT",
    "A85247925A": "TOTAL"
}

final["state"] = final["state"].replace(series_to_state)


final = final[~final["state"].isin(["OT", "TOTAL"])]

# CSV
out_path = "outputs/visiting_visas_by_state.csv"
final.to_csv(out_path, index=False, encoding="utf-8-sig")


final.to_csv(OUT_CSV, index=False, encoding="utf-8-sig")

print(f"✅ Parsed sheet '{candidate_name}' and saved: {OUT_CSV}")
print(final.head(12))
