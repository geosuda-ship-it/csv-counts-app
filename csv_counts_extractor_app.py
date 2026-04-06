import csv
from io import StringIO, BytesIO

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="NEX DE 強度→定量値変換アプリ（長崎大ver1）",
    page_icon="🧪",
    layout="wide"
)

st.title("NEX DE 強度→定量値変換アプリ（長崎大ver1）")
st.write(
    "NEX DEから出力したCSVファイルをアップロードすると，"
    "強度データの抽出，ドリフト補正，定量値の算出を自動で行い，"
    "ボタンを押すことで結果をExcelファイルとしてダウンロードできます。"
)
st.write("注意）ドリフト補正のため「QC-2」のデータはCSVファイルに必ず含めてください。")

# =========================
# 1. QC-2 基準強度（cps/μA）
# =========================
reference_qc2 = {
    "Mid-Z_K-Kα": 4.27826,
    "Mid-Z_Ca-Kβ1": 0.15422,
    "Mid-Z_Mn-Kα": 0.90407,
    "Mid-Z_Fe-Kβ1": 3.15459,
    "High-Z_Zn-Kα": 0.01778,
    "High-Z_Rb-Kα": 0.22959,
    "High-Z_Sr-Kα": 0.07396,
    "High-Z_Y-Kα": 0.09792,
    "High-Z_Zr-Kα": 0.17312,
    "High-Z_Nb-Kα": 0.10719,
    "High-Z_Ag-Kα": 1.53429,
}
target_cols = list(reference_qc2.keys())

# =========================
# 2. 検量線定数
# =========================
B = {
    "K": 9312.59,
    "Ca": 30145.8,
    "Mn": 746.088,
    "Fe": 4061.51,
    "Zn": 7048.49,
    "Rb": 1360.26,
    "Sr": 1112.8,
    "Y": 898.725,
    "Zr": 810.974,
    "Nb": 737.066,
}

C = {
    "K": -1378.82,
    "Ca": -204.571,
    "Mn": -315.439,
    "Fe": -5513.26,
    "Zn": -44.4817,
    "Rb": -16.0436,
    "Sr": -11.9096,
    "Y": -12.6423,
    "Zr": -18.6844,
    "Nb": -31.1933,
}

# =========================
# 3. 重なり補正係数
# =========================
BG14 = -0.113036
BG15 = -0.121299
BG16 = -0.161034


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


def clean_numeric(x):
    if pd.isna(x):
        return None
    if isinstance(x, (int, float)):
        return x

    s = str(x).strip()
    if s == "":
        return None

    parts = s.split()
    if len(parts) >= 2:
        try:
            return float(parts[-1])
        except Exception:
            pass

    try:
        return float(s)
    except Exception:
        return None


def extract_cps_blocks(rows):
    cps_indices = []
    for i, row in enumerate(rows):
        normalized = [str(x).strip().lower() for x in row]
        if "cps/μa" in normalized:
            cps_indices.append(i)

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

            row = list(row)

            if len(row) < len(columns):
                row = row + [""] * (len(columns) - len(row))
            elif len(row) > len(columns):
                row = row[:len(columns)]

            all_data.append(row)

    if not all_data:
        raise ValueError("cps/μA データが見つかりません")

    df = pd.DataFrame(all_data, columns=columns)

    numeric_df = df.iloc[:, 3:].copy()
    for col in numeric_df.columns:
        numeric_df[col] = numeric_df[col].map(clean_numeric)
        numeric_df[col] = pd.to_numeric(numeric_df[col], errors="coerce")

    df = pd.concat([df.iloc[:, :3].copy(), numeric_df], axis=1)

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    return df


def calculate_drift_factors(df):
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

    return drift_factors


def apply_drift_and_quantification(df):
    drift_factors = calculate_drift_factors(df)

    df_corrected = df.copy()
    for col in target_cols:
        factor = drift_factors[col]
        if factor is not None:
            df_corrected[col] = df_corrected[col] * factor

    ag_col = "High-Z_Ag-Kα"
    if ag_col not in df_corrected.columns:
        raise ValueError(f"Ag 内標準列が見つかりません: {ag_col}")

    # K〜Fe
    df_corrected["K"] = B["K"] * df_corrected["Mid-Z_K-Kα"] + C["K"]
    df_corrected["Ca"] = B["Ca"] * df_corrected["Mid-Z_Ca-Kβ1"] + C["Ca"]
    df_corrected["Mn"] = B["Mn"] * df_corrected["Mid-Z_Mn-Kα"] + C["Mn"]
    df_corrected["Fe"] = B["Fe"] * df_corrected["Mid-Z_Fe-Kβ1"] + C["Fe"]

    # Zn も Ag コンプトン内標準化
    df_corrected["Zn"] = B["Zn"] * (df_corrected["High-Z_Zn-Kα"] / df_corrected[ag_col]) + C["Zn"]

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
        "Rb", "Sr", "Y", "Zr", "Nb",
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

    df_result["Date"] = pd.to_datetime(df_result["Date"], errors="coerce")

    df_result = df_result.sort_values(
        ["Date", "Sample", "Method"],
        ascending=[False, True, True],
        na_position="last"
    ).reset_index(drop=True)

    return df_result, drift_factors


def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="quantified_results")
    output.seek(0)
    return output.getvalue()


# =========================
# 4. 対話UI
# =========================
uploaded = st.file_uploader("NEX DEから出力したCSVファイルをアップロード", type=["csv"])

if "messages" not in st.session_state:
    st.session_state.messages = []

if "df_result" not in st.session_state:
    st.session_state.df_result = None

if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None

if "drift_factors" not in st.session_state:
    st.session_state.drift_factors = None

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.write(msg["content"])

prompt = st.chat_input(
    "例：定量計算して / 補正係数を見せて / 検量線定数を出して / 重なり補正係数を出して"
)

if prompt:
    st.session_state.messages.append({"role": "user", "content": prompt})

    with st.chat_message("user"):
        st.write(prompt)

    lower_prompt = prompt.lower().strip()

    if uploaded is None and (
        ("定量" in prompt) or
        ("計算" in prompt) or
        ("結果" in prompt) or
        ("表" in prompt) or
        ("補正係数" in prompt) or
        ("drift" in lower_prompt)
    ):
        reply = "まず，NEX DEから出力したCSVファイルをアップロードしてください。"
        with st.chat_message("assistant"):
            st.write(reply)
        st.session_state.messages.append({"role": "assistant", "content": reply})

    elif ("定量" in prompt) or ("計算" in prompt):
        try:
            rows = read_uploaded_csv(uploaded)
            df_counts = extract_cps_blocks(rows)
            df_result, drift_factors = apply_drift_and_quantification(df_counts)
            excel_bytes = to_excel_bytes(df_result)

            st.session_state.df_result = df_result
            st.session_state.excel_bytes = excel_bytes
            st.session_state.drift_factors = drift_factors

            reply = (
                f"定量計算が完了しました。{len(df_result)} 行の結果を作成しました。"
                "下に結果表を表示し，Excelもダウンロードできます。"
            )

            with st.chat_message("assistant"):
                st.write(reply)

            st.session_state.messages.append({"role": "assistant", "content": reply})

        except Exception as e:
            reply = f"処理中にエラーが出ました: {e}"
            with st.chat_message("assistant"):
                st.write(reply)
            st.session_state.messages.append({"role": "assistant", "content": reply})

    elif ("補正係数" in prompt) or ("drift" in lower_prompt):
        if st.session_state.drift_factors is None:
            reply = "まだ補正係数がありません。先に「定量計算して」と入力してください。"
            with st.chat_message("assistant"):
                st.write(reply)
            st.session_state.messages.append({"role": "assistant", "content": reply})
        else:
            with st.chat_message("assistant"):
                st.write("ドリフト補正係数を表示します。")
                drift_df = pd.DataFrame({
                    "Line": list(st.session_state.drift_factors.keys()),
                    "Drift factor": list(st.session_state.drift_factors.values()),
                })
                st.dataframe(drift_df, use_container_width=True)

            st.session_state.messages.append(
                {"role": "assistant", "content": "ドリフト補正係数を表示しました。"}
            )

    elif ("検量線定数" in prompt) or ("bとc" in lower_prompt) or ("b c" in lower_prompt):
        with st.chat_message("assistant"):
            st.write("検量線定数（B, C）を表示します。")
            calib_df = pd.DataFrame({
                "Element": list(B.keys()),
                "B": [B[k] for k in B.keys()],
                "C": [C[k] for k in B.keys()],
            })
            st.dataframe(calib_df, use_container_width=True)

        st.session_state.messages.append(
            {"role": "assistant", "content": "検量線定数（B, C）を表示しました。"}
        )

    elif ("重なり補正係数" in prompt) or ("bg14" in lower_prompt) or ("bg15" in lower_prompt) or ("bg16" in lower_prompt):
        with st.chat_message("assistant"):
            st.write("重なり補正係数を表示します。")
            overlap_df = pd.DataFrame({
                "Coefficient": ["BG14", "BG15", "BG16"],
                "Value": [BG14, BG15, BG16],
                "Meaning": [
                    "Y 補正に使う Rb 項の係数",
                    "Zr 補正に使う Sr 項の係数",
                    "Nb 補正に使う Y 項の係数",
                ],
            })
            st.dataframe(overlap_df, use_container_width=True)

        st.session_state.messages.append(
            {"role": "assistant", "content": "重なり補正係数を表示しました。"}
        )

    elif ("qc-2基準強度" in lower_prompt) or ("qc2基準強度" in lower_prompt) or ("基準強度" in prompt):
        with st.chat_message("assistant"):
            st.write("QC-2 基準強度を表示します。")
            qc2_df = pd.DataFrame({
                "Line": list(reference_qc2.keys()),
                "Reference intensity (cps/μA)": list(reference_qc2.values()),
            })
            st.dataframe(qc2_df, use_container_width=True)

        st.session_state.messages.append(
            {"role": "assistant", "content": "QC-2 基準強度を表示しました。"}
        )

    elif ("結果" in prompt) or ("表" in prompt):
        if st.session_state.df_result is None:
            reply = "まだ結果がありません。先に「定量計算して」と入力してください。"
            with st.chat_message("assistant"):
                st.write(reply)
            st.session_state.messages.append({"role": "assistant", "content": reply})
        else:
            with st.chat_message("assistant"):
                st.write("現在の定量結果を表示します。")
                st.dataframe(st.session_state.df_result, use_container_width=True)

            st.session_state.messages.append(
                {"role": "assistant", "content": "定量結果を表示しました。"}
            )

    else:
        reply = (
            "現在対応している指示は，"
            "「定量計算して」「補正係数を見せて」「結果を見せて」"
            "「検量線定数を出して」「重なり補正係数を出して」"
            "「QC-2基準強度を出して」です。"
        )
        with st.chat_message("assistant"):
            st.write(reply)
        st.session_state.messages.append({"role": "assistant", "content": reply})

if st.session_state.df_result is not None:
    st.subheader("定量結果")
    st.dataframe(st.session_state.df_result, use_container_width=True)

    if uploaded is not None:
        base_name = uploaded.name.rsplit(".", 1)[0]
    else:
        base_name = "result"

    st.download_button(
        label="Excelをダウンロード",
        data=st.session_state.excel_bytes,
        file_name=f"{base_name}_quantified_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
