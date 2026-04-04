import csv
from collections import Counter
import pandas as pd

# =========================
# 0. 入出力ファイル
# =========================
input_file = "20260316.csv"
output_file = "quantified_results.xlsx"

# =========================
# 1. QC-2 基準強度（cps/μA）
# =========================
reference_qc2 = {
    "Mid-Z_K-Kα": 4.2841,
    "Mid-Z_Ca-Kβ1": 0.1575,
    "Mid-Z_Mn-Kα": 0.9233,
    "Mid-Z_Fe-Kβ1": 3.2176,
    "High-Z_Zn-Kα": 0.0108,
    "High-Z_Rb-Kα": 0.1363,
    "High-Z_Sr-Kα": 0.0451,
    "High-Z_Y-Kα": 0.0600,
    "High-Z_Zr-Kα": 0.1032,
    "High-Z_Nb-Kα": 0.0657,
    "High-Z_Ag-Kα": 1.3920,
}

target_cols = list(reference_qc2.keys())

# =========================
# 2. 検量線定数
# =========================
B = {
    "K": 9278.19,
    "Ca": 30241.3,
    "Mn": 742.281,
    "Fe": 3996.0,
    "Zn": 11115.8,
    "Rb": 2021.13,
    "Sr": 1747.57,
    "Y": 1291.92,
    "Zr": 1226.15,
    "Nb": 1120.36,
}

C = {
    "K": -1334.6,
    "Ca": -275.371,
    "Mn": -328.074,
    "Fe": -5437.0,
    "Zn": -44.1916,
    "Rb": -8.69875,
    "Sr": -12.4792,
    "Y": -8.57,
    "Zr": -16.6585,
    "Nb": -29.34,
}

# 重なり補正係数
BG14 = -0.12     # Y に対する Rb
BG15 = -0.12412  # Zr に対する Sr
BG16 = -0.18     # Nb に対する Y

# =========================
# 3. 生CSV読み込み
# =========================
with open(input_file, "r", encoding="cp932", errors="replace", newline="") as f:
    rows = list(csv.reader(f))

# =========================
# 4. cps/μA ブロック位置を探す
# =========================
cps_indices = [
    i for i, row in enumerate(rows)
    if "cps/μa" in [str(x).strip().lower() for x in row]
]

if not cps_indices:
    raise ValueError("cps/μA 行が見つかりません")

# =========================
# 5. cps/μA データ抽出
# =========================
all_data = []
columns = None

for idx in cps_indices:
    if idx < 2:
        continue

    header1 = rows[idx - 2]
    header2 = rows[idx - 1]

    # 最初のブロックで列名作成
    if columns is None:
        new_cols = []
        seen = {}

        for h1, h2 in zip(header1[3:], header2[3:]):
            h1 = str(h1).strip()
            h2 = str(h2).strip()

            if h1 and h2:
                name = f"{h1}_{h2}"
            elif h2:
                name = h2
            else:
                name = h1

            # 重複列名があれば _2, _3 を付ける
            if name in seen:
                seen[name] += 1
                name = f"{name}_{seen[name]}"
            else:
                seen[name] = 1

            new_cols.append(name)

        columns = ["Sample", "Method", "Date"] + new_cols

    # cps/μA 行の下から空行までをデータ取得
    for row in rows[idx + 1:]:
        if len(row) == 0 or all(str(x).strip() == "" for x in row):
            break

        # 列数を合わせる
        if len(row) < len(columns):
            row = row + [""] * (len(columns) - len(row))
        elif len(row) > len(columns):
            row = row[:len(columns)]

        all_data.append(row)

if not all_data:
    raise ValueError("cps/μA データが見つかりません")

# =========================
# 6. DataFrame化
# =========================
df = pd.DataFrame(all_data, columns=columns)

# 数値変換（列番号で行う：重複列名対策）
for i in range(3, len(columns)):
    df.iloc[:, i] = pd.to_numeric(df.iloc[:, i], errors="coerce")

# 日付変換（失敗してもそのまま）
df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

# =========================
# 7. Method順に並べ替え
# =========================
order = {
    "obart_prec2_air": 0,
    "obart_stand2_air": 1,
    "obart_quick2_air": 2
}

df["sort_key"] = df["Method"].map(order).fillna(999)
df = df.sort_values(["sort_key", "Sample"], kind="stable").drop(columns="sort_key")

# =========================
# 8. QC-2 / QC2 を認識
# =========================
sample_norm = (
    df["Sample"]
    .astype(str)
    .str.strip()
    .str.upper()
    .str.replace("-", "", regex=False)
    .str.replace(" ", "", regex=False)
)

qc2_rows = df[sample_norm == "QC2"]

if qc2_rows.empty:
    raise ValueError("QC-2 または QC2 の行が見つかりません")

qc2_measured = qc2_rows.iloc[0]

# =========================
# 9. ドリフト補正係数算出
# =========================
drift_factors = {}

for col in target_cols:
    if col not in df.columns:
        raise ValueError(f"必要な列が見つかりません: {col}")

    measured_val = qc2_measured[col]
    ref_val = reference_qc2[col]

    if pd.isna(measured_val) or measured_val == 0:
        drift_factors[col] = None
    else:
        drift_factors[col] = ref_val / measured_val

# =========================
# 10. ドリフト補正
# =========================
df_corrected = df.copy()

for col in target_cols:
    factor = drift_factors[col]
    if factor is not None:
        df_corrected[col] = df_corrected[col] * factor

# =========================
# 11. 定量計算
# =========================
ag_col = "High-Z_Ag-Kα"

if ag_col not in df_corrected.columns:
    raise ValueError(f"Ag 内標準列が見つかりません: {ag_col}")

# K〜Zn（強度そのまま）
df_corrected["K"]  = B["K"]  * df_corrected["Mid-Z_K-Kα"]    + C["K"]
df_corrected["Ca"] = B["Ca"] * df_corrected["Mid-Z_Ca-Kβ1"]  + C["Ca"]
df_corrected["Mn"] = B["Mn"] * df_corrected["Mid-Z_Mn-Kα"]   + C["Mn"]
df_corrected["Fe"] = B["Fe"] * df_corrected["Mid-Z_Fe-Kβ1"]  + C["Fe"]
df_corrected["Zn"] = B["Zn"] * df_corrected["High-Z_Zn-Kα"]  + C["Zn"]

# Rb, Sr（Ag 内標準）
df_corrected["Rb"] = B["Rb"] * (df_corrected["High-Z_Rb-Kα"] / df_corrected[ag_col]) + C["Rb"]
df_corrected["Sr"] = B["Sr"] * (df_corrected["High-Z_Sr-Kα"] / df_corrected[ag_col]) + C["Sr"]

# Y, Zr, Nb（重なり補正込み）
df_corrected["Y"]  = B["Y"]  * (df_corrected["High-Z_Y-Kα"]  / df_corrected[ag_col]) + C["Y"]  + df_corrected["Rb"] * BG14
df_corrected["Zr"] = B["Zr"] * (df_corrected["High-Z_Zr-Kα"] / df_corrected[ag_col]) + C["Zr"] + df_corrected["Sr"] * BG15
df_corrected["Nb"] = B["Nb"] * (df_corrected["High-Z_Nb-Kα"] / df_corrected[ag_col]) + C["Nb"] + df_corrected["Y"]  * BG16

# =========================
# 12. 最終結果だけ残す
# =========================
result_cols = [
    "Sample", "Method", "Date",
    "K", "Ca", "Mn", "Fe", "Zn",
    "Rb", "Sr", "Y", "Zr", "Nb"
]

df_result = df_corrected[result_cols].copy()

# 列名を ppm 付きに変更
df_result = df_result.rename(columns={
    "K": "K ppm",
    "Ca": "Ca ppm",
    "Mn": "Mn ppm",
    "Fe": "Fe ppm",
    "Zn": "Zn ppm",
    "Rb": "Rb ppm",
    "Sr": "Sr ppm",
    "Y": "Y ppm",
    "Zr": "Zr ppm",
    "Nb": "Nb ppm",
})

# =========================
# 13. Excel 出力
# =========================
df_result.to_excel(output_file, index=False)

print(f"保存完了: {output_file}")
print(df_result.head())
