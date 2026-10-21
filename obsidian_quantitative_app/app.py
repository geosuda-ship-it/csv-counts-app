
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

st.subheader("STEP 3: Cluster-based Chemical Group Assignment (Clustering Method)  \nクラスター単位の化学組成グループ判定（クラスター法）")

st.write(
    "各資料から得られた複数の定量値の中央値を算出し、Rb分率、Sr分率、Zr分率に基づいてクラスター区分を行います。"
    "次に、各クラスターの分布楕円を作成し、既知の化学組成グループの判別楕円との重なりを比較することで、クラスターごとに化学組成グループを判定します。"
    "STEP 2で出力されたExcelファイル（日付_定量値_ellipse_results.xlsx）をアップロードして解析を実行すると、各資料の化学組成グループの判定結果をExcelファイルとしてダウンロードできます。"
    "出力ファイルには、STEP 2による判定結果も含まれています。"
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


st.subheader("1．強度CSVファイル")

csv_files = st.file_uploader(
    "CSVファイルを選択してください（複数選択可）",
    type=["csv"],
    accept_multiple_files=True,
)

if csv_files:
    st.write(f"選択されたファイル：{len(csv_files)}個")
    for file in csv_files:
        st.write(f"・{file.name}")


st.subheader("2．定量計算パラメーター")

parameter_file = st.file_uploader(
    "定量計算パラメーターExcelを選択してください",
    type=["xlsx", "xlsm"],
)


st.subheader("3．定量計算")

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
