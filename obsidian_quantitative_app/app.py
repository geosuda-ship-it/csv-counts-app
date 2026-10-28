import hashlib
import re
from pathlib import Path

import streamlit as st


st.set_page_config(
    page_title="NEX DE 定量計算",
    page_icon="🔬",
    layout="centered",
)

st.markdown(
    """
    <h3 style="text-align: center; font-size: 32px;">
        Obsidian Provenance Analysis Tools<br>
        Using Quantitative EDXRF Data
    </h3>
    """,
    unsafe_allow_html=True,
)

st.subheader(
    "STEP 1: XRF Intensity-to-Concentration Conversion  \n"
    "NEX DE 強度データの定量計算"
)

st.write(
    "ドリフト補正用試料（QC-2）の基準強度と測定強度に基づいて"
    "ドリフト補正係数を求め、検量線法により定量値を算出します。 "
    "QC-2を複数回測定した場合は、その測定回数と同じ数の定量値を、"
    "1つの資料のXRF強度データから算出します。NEX DEから出力した"
    "強度CSVファイルと、機関ごとの定量計算パラメーターExcelを"
    "アップロードすると、定量値と中央値をExcelファイルとして"
    "ダウンロードできます。"
)


def make_output_name(csv_files):
    """CSV名を昇順に連結し、STEP 1の出力名を作る。"""
    names = []

    for file in csv_files:
        stem = Path(file.name).stem
        # Colabやブラウザによる重複名末尾「 (1)」などを除く。
        stem = re.sub(r"\s*\(\d+\)$", "", stem).strip()
        names.append(stem)

    return "_".join(sorted(names)) + "_step1_定量値.xlsx"


def make_signature(csv_files, parameter_file):
    """入力が変更されたとき、古い計算結果を表示しないための識別子。"""
    digest = hashlib.sha256()

    for file in sorted(csv_files, key=lambda item: item.name):
        digest.update(file.name.encode())
        digest.update(file.getvalue())

    if parameter_file is not None:
        digest.update(parameter_file.name.encode())
        digest.update(parameter_file.getvalue())

    return digest.hexdigest()


st.subheader("1. Upload XRF Intensity CSV Files")

st.write(
    "NEX DEから出力した強度CSVファイルをアップロードしてください。"
    "複数ファイルを同時に選択できます。"
)

csv_files = st.file_uploader(
    "強度CSVファイルを選択してください",
    type=["csv"],
    accept_multiple_files=True,
    key="intensity_csv",
)

st.caption("CSV形式・複数ファイル選択可")

if csv_files:
    st.write(f"選択されたCSV：{len(csv_files)}ファイル")
    for file in csv_files:
        st.write(f"・{file.name}")


st.subheader("2. Upload Quantification Parameters File")

st.write(
    "機関ごとの定量計算パラメーターExcelファイルを"
    "アップロードしてください。"
)

parameter_file = st.file_uploader(
    "機関ごとの定量計算パラメーターExcelを選択してください",
    type=["xlsx", "xlsm"],
    key="quantification_parameters",
)

st.caption("「定量計算パラメーター」シートを含むExcelファイル")


st.subheader("3. Run Quantification")

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
        from parameterized_quantitative import create_multi_quantitative_excel

        csv_data = {
            file.name: file.getvalue()
            for file in csv_files
        }

        with st.spinner("定量計算と中央値の集計を実行しています..."):
            result = create_multi_quantitative_excel(
                csv_data,
                parameter_file.getvalue(),
                parameter_file.name,
            )

        st.session_state["quant_result"] = {
            "signature": signature,
            "content": result["content"],
            "filename": output_name,
            "warnings": result["warnings"],
        }

    except Exception as error:
        st.session_state.pop("quant_result", None)
        st.error("定量計算中にエラーが発生しました。")
        st.exception(error)


saved = st.session_state.get("quant_result")

if saved and saved["signature"] == signature:
    st.success("STEP 1の定量計算が完了しました。")

    for name, day, message in saved["warnings"]:
        day_text = day.isoformat() if hasattr(day, "isoformat") else str(day)
        st.warning(f"{name} / {day_text}：{message}")

    st.download_button(
        label="STEP 1 定量値Excelをダウンロード",
        data=saved["content"],
        file_name=saved["filename"],
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
    )
