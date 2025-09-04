import pandas as pd
import os

os.makedirs("outputs", exist_ok=True)

# 1) Load the cleaned dataset
df = pd.read_csv("cleaned_data/short_term_arrivals_state_of_stay_clean.csv")

# 2) Keep only rows that mention "Short-term Visitors arriving"
mask = df["unnamed_0"].astype(str).str.contains("Short-term Visitors arriving", na=False)
visitor_df = df.loc[mask].copy()

# 3) Split the descriptor column
parts = visitor_df["unnamed_0"].str.split(";", n=3, expand=True)
visitor_df["metric"]   = parts[0].str.strip()
visitor_df["category"] = parts[1].str.strip()
visitor_df["state"]    = parts[2].str.strip().str.replace(";", "", regex=False)

# 4) Manually pick date & value columns (更稳)
date_col = "unnamed_6"    # 你的文件里是 Series End / 日期
value_col = "unnamed_7"   # 最后的列通常是 arrivals 数字

# 5) Build final tidy table
final_df = visitor_df[[date_col, "state", "category", value_col]].copy()
final_df = final_df.rename(columns={date_col: "date", value_col: "arrivals"})

# 6) Clean state names
state_mapping = {
    "NSW": "NSW", "Vic": "VIC", "Qld": "QLD", "SA": "SA", "WA": "WA",
    "Tas": "TAS", "NT": "NT", "ACT": "ACT",
    "Other Territories": "OT", "Total (State of residence/stay)": "TOTAL"
}
final_df["state"] = final_df["state"].replace(state_mapping)

# 7) Parse date & clean numbers
final_df["date"] = pd.to_datetime(final_df["date"], errors="coerce")
final_df["year_month"] = final_df["date"].dt.to_period("M").astype(str)
final_df["arrivals"] = pd.to_numeric(final_df["arrivals"], errors="coerce")

# 8) Drop invalid rows and TOTAL if not needed
final_df = final_df.dropna(subset=["arrivals"])
final_df = final_df[final_df["arrivals"] >= 0]
final_df = final_df[final_df["state"] != "TOTAL"]

# 9) Export
out_path = "outputs/visiting_visas_by_state.csv"


# === Diagnose which column has the actual counts ===
import numpy as np

# visitor_df 是已经筛出 "Short-term Visitors arriving" 的子表
candidates = []
for c in visitor_df.columns:
    s = pd.to_numeric(visitor_df[c], errors="coerce")
    if s.notna().sum() > 0:
        candidates.append((c, int(s.notna().sum()), float(np.nanmin(s)), float(np.nanmax(s))))
diag = pd.DataFrame(candidates, columns=["col","n_nonnull","min","max"]).sort_values("max", ascending=False)
print("\nTop numeric columns by max value:")
print(diag.head(8))





final_df.to_csv(out_path, index=False, encoding="utf-8-sig")

print(f"✅ Visiting visa data saved to {out_path}")
print(final_df.head(10))
