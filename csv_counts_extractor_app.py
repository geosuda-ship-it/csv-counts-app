import csv
from io import StringIO, BytesIO
import pandas as pd
import streamlit as st

st.set_page_config(page_title="XRF quantification app", page_icon="🧪", layout="wide")

st.title("XRF quantification app")
st.write("cps/μA 形式のCSVをアップロードすると，強度抽出 → ドリフト補正 → 定量値算出を行い，Excelで出力します。")

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
BG14 = -0.12
BG15 = -0.12412
BG16 = -0.18


def read_uploaded_csv(uploaded_file):
    raw = uploaded_file.getvalue()
    encodings = ["cp932", "utf-8-sig", "utf-8"]
    last_error = None

    for enc in encodings:
        try:
            text = raw.decode(enc, errors="replace")
            return list(csv.reader(StringIO(text)))
        except Exception as e:
            last_error = e

    raise ValueError(f"CSVを読み込めませんでした: {last_error}")


def extract_cps_blocks(rows):
    cps_indices = [
        i for i, row in enumerate(rows)
        if "cps/μa" in [str(x).strip().lower() for x in row]
    ]

    if not cps_indices:
        raise ValueError("cps/μA 行が見つかりません")

    all_data = []
    columns = None

    for idx in cps_indices:
        if idx < 2:
            continue

        header1 = rows[idx - 2]
        header2 = rows[idx - 1]

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

                if name in seen:
                    seen[name] += 1
                    name = f"{name}_{seen[name]}"
                else:
                    seen[name] = 1

                new_cols.append(name)

            columns = ["Sample", "Method", "Date"] + new_cols

        for row in rows[idx + 1:]:
            if len(row) == 0 or all(str(x).strip() == "" for x in row):
                break

            if len(row) < len(columns):
                row = row + [""] * (len(columns) - len(row))
            elif len(row) > len(columns):
                row = row[:len(columns)]

            all_data.append(row)

    if not all_data:
        raise ValueError("cps/μA データが見つかりません")

    df = pd.DataFrame(all_data, columns=columns)

    def clean_numeric(x):
        if pd.isna(x):
            return None

        if isinstance(x, (int, float)):
            return x

        s = str(x).strip()
        if s == "":
            return None

        parts = s.split()

        # "0 3.477209" のような場合は最後の値を使う
        if len(parts) >= 2:
            try:
                return float(parts[-1])
            except Exception:
                pass

        try:
            return float(s)
        except Exception:
            return None

    numeric_df = df.iloc[:, 3:].copy()
    for col in numeric_df.columns:
        numeric_df[col] = numeric_df[col].map(clean_numeric)
        numeric_df[col] = pd.to_numeric(numeric_df[col], errors="coerce")

    df = pd.concat([df.iloc[:, :3].copy(), numeric_df], axis=1)

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    order = {
        "obart_prec2_air": 0,
        "obart_stand2_air": 1,
        "obart_quick2_air": 2,
    }

    df["sort_key"] = df["Method"].map(order).fillna(999)
    df = df.sort_values(["sort_key", "Sample"], kind="stable").drop(columns="sort_key")

    return df


def apply_drift_and_quantification(df):
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

    df_corrected = df.copy()
    for col in target_cols:
        factor = drift_factors[col]
        if factor is not None:
            df_corrected[col] = df_corrected[col] * factor

    ag_col = "High-Z_Ag-Kα"
    if ag_col not in df_corrected.columns:
        raise ValueError(f"Ag 内標準列が見つかりません: {ag_col}")

    # K〜Zn
    df_corrected["K"] = B["K"] * df_corrected["Mid-Z_K-Kα"] + C["K"]
    df_corrected["Ca"] = B["Ca"] * df_corrected["Mid-Z_Ca-Kβ1"] + C["Ca"]
    df_corrected["Mn"] = B["Mn"] * df_corrected["Mid-Z_Mn-Kα"] + C["Mn"]
    df_corrected["Fe"] = B["Fe"] * df_corrected["Mid-Z_Fe-Kβ1"] + C["Fe"]
    df_corrected["Zn"] = B["Zn"] * df_corrected["High-Z_Zn-Kα"] + C["Zn"]

    # Rb, Sr
    df_corrected["Rb"] = B["Rb"] * (df_corrected["High-Z_Rb-Kα"] / df_corrected[ag_col]) + C["Rb"]
    df_corrected["Sr"] = B["Sr"] * (df_corrected["High-Z_Sr-Kα"] / df_corrected[ag_col]) + C["Sr"]

    # Y, Zr, Nb
    df_corrected["Y"] = (
        B["Y"] * (df_corrected["High-Z_Y-Kα"] / df_corrected[ag_col])
        + C["Y"]
        + df_corrected["Rb"] * BG14
    )
    df_corrected["Zr"] = (
        B["Zr"] * (df_corrected["High-Z_Zr-Kα"] / df_corrected[ag_col])
        + C["Zr"]
        + df_corrected["Sr"] * BG15
    )
    df_corrected["Nb"] = (
        B["Nb"] * (df_corrected["High-Z_Nb-Kα"] / df_corrected[ag_col])
        + C["Nb"]
        + df_corrected["Y"] * BG16
    )

    result_cols = [
        "Sample", "Method", "Date",
        "K", "Ca", "Mn", "Fe", "Zn",
        "Rb", "Sr", "Y", "Zr", "Nb"
    ]

    df_result = df_corrected[result_cols].copy()

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

    return df_result


def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="quantified_results")
