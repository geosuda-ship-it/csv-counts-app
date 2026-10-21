
import hashlib
import re
from pathlib import Path

import streamlit as st

st.set_page_config(
    page_title="NEX DE 定量計算",
    page_icon="🔬",
    layout="centered",
)

st.subheader(":red[Obsidian Provenance Analysis Tools Using Quantitative EDXRF Data]")

st.subheader("STEP 1: XRF Intensity-to-Concentration Conversion  \nNEX DE 強度データの定量計算")

st.write(
    "ドリフト補正用試料（QC-2）の基準強度と測定強度に基づいてドリフト補正係数を求め、検量線法により定量値を算出します。"
    "本プログラムでは、QC-2を複数回測定した場合、その測定回数と同じ数の定量値を、1つの資料のXRF強度データから算出します。"
    "NEX DEから出力したCSV形式の測定結果ファイル（例：20260728.csv）と、機関ごとの定量計算パラメーターファイル（例：長崎大学_定量計算パラメーターver1.xlsx）をアップロードすると、"
    "定量値の計算結果をExcelファイルとしてダウンロードできます。"
)

def make_output_name(csv_files):
    names = []
    for file in csv_files:
        stem = Path(file.name).stem
        stem = re.sub(r"\s*\(\d+\)$", "", stem)
        names.append(stem)

    return "_".join(sorted(names)) + "_定量値.xlsx"


def make_signature(csv_files, parameter_file):
    digest = hashlib.sha256()

    for file in sorted(csv_files, key=lambda f: f.name):
        digest.update(file.name.encode())
        digest.update(file.getvalue())

    if parameter_file is not None:
        digest.update(parameter_file.name.encode())
        digest.update(parameter_file.getvalue())

    return digest.hexdigest()


st.subheader("### 1. Upload NEX DE Intensity CSV File")

csv_files = st.file_uploader(
    "NED DEから出力したCSVファイルを選択してください（複数選択可）",
    type=["csv"],
    accept_multiple_files=True,
)

if csv_files:
    st.write(f"選択されたファイル：{len(csv_files)}個")
    for file in csv_files:
        st.write(f"・{file.name}")


st.subheader("### 2. Upload Quantification Paramers")

parameter_file = st.file_uploader(
    "機関ごとの定量計算パラメーターファイルをアップロードしてください",
    type=["xlsx", "xlsm"],
)


st.subheader("### 3．Quantification")

ready = bool(csv_files) and parameter_file is not None

names = [file.name for file in csv_files]
duplicate_names = len(names) != len(set(names))

if duplicate_names:
    st.error("同じファイル名のCSVがあります。名前を変更してください。")

if ready:
    output_name = make_output_name(csv_files)
    st.info(f"出力ファイル名：{output_name}")

signature = make_signature(csv_files, parameter_file)

if st.button(
    "定量計算を実行",
    type="primary",
    disabled=not ready or duplicate_names,
):
    try:
        from parameterized_quantitative import (
            create_multi_quantitative_excel,
        )

        csv_data = {
            file.name: file.getvalue()
            for file in csv_files
        }

        with st.spinner("定量計算を実行しています..."):
            result = create_multi_quantitative_excel(
                csv_data,
                parameter_file.getvalue(),
                parameter_file.name,
            )

        st.session_state["quant_result"] = {
            "signature": signature,
            "content": result["content"],
            "filename": output_name,
            "file_count": result["file_count"],
            "row_count": result["output_row_count"],
            "warnings": result["warnings"],
        }

    except Exception as error:
        st.session_state.pop("quant_result", None)
        st.error("定量計算中にエラーが発生しました。")
        st.exception(error)


saved = st.session_state.get("quant_result")

if saved and saved["signature"] == signature:
    st.success("定量計算が完了しました。")

    st.write(f"入力CSV：{saved['file_count']}ファイル")
    st.write(f"出力行数：{saved['row_count']}行")

    for name, day, message in saved["warnings"]:
        st.warning(f"{name} / {day}：{message}")

    st.download_button(
        label="定量値Excelをダウンロード",
        data=saved["content"],
        file_name=saved["filename"],
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
    )
