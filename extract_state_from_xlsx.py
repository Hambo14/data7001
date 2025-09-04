import os
import pandas as pd
import numpy as np

# ---- settings ----
SRC_XLSX = "short_term_arrivals_state_of_stay.xlsx"   # 原始ABS工作簿
OUT_CSV  = "outputs/visiting_visas_by_state.csv"

os.makedirs("outputs", exist_ok=True)

# 1) Open workbook and list sheets
xls = pd.ExcelFile(SRC_XLSX, engine="openpyxl")

def looks_like_data(df: pd.DataFrame) -> bool:
    """Heuristic: first col mostly dates, and >=6 numeric columns."""
    if df.empty:
        return False
    # 尝试把第一列转成日期
    dt = pd.to_datetime(df.iloc[:, 0], errors="coerce")
    date_ratio = dt.notna().mean()
    # 统计数值列个数
    num_cols = sum(pd.api.types.is_numeric_dtype(df[c]) for c in df.columns[1:])
    return (date_ratio > 0.6) and (num_cols >= 6)

candidate = None
candidate_name = None

# 2) Try each sheet with a few header guesses
for sh in xls.sheet_names:
    # 直接读（可能没有表头）
    tmp = pd.read_excel(xls, sheet_name=sh, header=None)
    # 尝试定位“数据起始行”（包含类似 Month 的行）
    # 粗暴办法：从前 25 行里，找到一行让其下方像“日期+数字”的
    found = False
    for hdr_row in range(0, min(25, len(tmp))):
        df = pd.read_excel(xls, sheet_name=sh, header=hdr_row)
        # 丢弃全空列
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
# 只保留“第一列+数值列”
first_col = wide.columns[0]
# 去掉非数值的后续列
keep_cols = [first_col] + [c for c in wide.columns[1:] if pd.api.types.is_numeric_dtype(wide[c])]
wide = wide[keep_cols]

# 丢弃第一列不是日期的行
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

# 只保留各州；如需包含全国合计，把这行注释掉
long_df = long_df[~long_df["state"].isin(["TOTAL", "OT"])]

# 6) Keep tidy columns and export
final = long_df[["year_month", "state", "arrivals"]].sort_values(["year_month", "state"])

# 生成 long_df 后，state 里现在是 Series ID
final = long_df[["year_month", "state", "arrivals"]].sort_values(["year_month", "state"])

# 🔹 在这里加映射
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

# 如果不想要 OT 和 TOTAL，可以过滤掉：
final = final[~final["state"].isin(["OT", "TOTAL"])]

# 最后再保存 CSV
out_path = "outputs/visiting_visas_by_state.csv"
final.to_csv(out_path, index=False, encoding="utf-8-sig")


final.to_csv(OUT_CSV, index=False, encoding="utf-8-sig")

print(f"✅ Parsed sheet '{candidate_name}' and saved: {OUT_CSV}")
print(final.head(12))
