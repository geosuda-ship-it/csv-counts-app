
import hashlib
from pathlib import Path

import streamlit as st

from ellipse_analysis import (
    read_measurements,
    read_ellipse_parameters,
    run_ellipse_analysis,
    export_results,
    public_table,
)


st.set_page_config(
    page_title="黒曜石の産地判別",
    page_icon="🔬",
    layout="centered",
)

st.subheader(":red[Obsidian Provenance Analysis Tools Using Quantitative EDXRF Data]")

st.subheader("STEP 2: Sample-based Chemical Group Assignment (Ellipse Method)  \n資料単位の化学組成グループ判定（判別楕円法）")

st.write(
    "各資料から得られた複数の定量分析値について、それぞれ既知の化学組成グループの判別楕円との適合性を評価し、その結果を集約して資料単位の化学組成グループを判定します。"
    "STEP 1で作成した定量値の計算結果ファイル（日付_定量値.xlsx）と、"
    "共通の判別楕円パラメータファイル（例：判別楕円パラメータ_ver24.xlsx）をアップロードすると、"
    "各資料の化学組成グループの判定結果をExcelファイルとしてダウンロードできます。"
)


# --------------------------------------------------
# 入力ファイル
# --------------------------------------------------

st.subheader("1．測定データ")

measurement_file = st.file_uploader(
    "元素定量値のExcelファイルを選択してください",
    type=["xlsx"],
    key="measurement",
)

st.caption("「定量値」シートを含むExcelファイル")


st.subheader("2．判別楕円パラメーター")

parameter_file = st.file_uploader(
    "判別楕円パラメーターのExcelファイルを選択してください",
    type=["xlsx"],
    key="parameters",
)

st.caption(
    "AS-1以降の各化学組成グループのシートから，"
    "D1〜D6の判別楕円パラメーターを読み込みます。"
)


# --------------------------------------------------
# 解析処理
# --------------------------------------------------

st.subheader("3．産地判別")

ready = (
    measurement_file is not None
    and parameter_file is not None
)

signature = None

if ready:
    digest = hashlib.sha256()

    for uploaded in (measurement_file, parameter_file):
        digest.update(uploaded.name.encode("utf-8"))
        digest.update(uploaded.getvalue())

    signature = digest.hexdigest()


if st.button(
    "産地判別を実行",
    type="primary",
    disabled=not ready,
):

    try:
        with st.spinner("判別楕円法による解析を実行しています..."):

            # 測定データの読み込み
            inputs = read_measurements(
                measurement_file.getvalue()
            )

            # 判別楕円パラメーターの読み込み
            parameters = read_ellipse_parameters(
                parameter_file.getvalue()
            )

            # 判別・得点・資料別集計・中央値
            results = run_ellipse_analysis(
                inputs,
                parameters,
            )

            # 結果Excel
            excel_bytes = export_results(
                results,
                inputs,
            )

            output_name = (
                Path(measurement_file.name).stem
                + "_ellipse_results.xlsx"
            )

            st.session_state["ellipse_result"] = {
                "signature": signature,
                "excel": excel_bytes,
                "filename": output_name,
                "sample_count": len(inputs.measurement_counts),
                "measurement_count": len(inputs.measurements),
                "group_count": parameters["Group"].nunique(),
                "summaries": public_table(results.summaries),
                "medians": results.medians,
                "data": public_table(results.data),
            }

    except Exception as error:
        st.session_state.pop("ellipse_result", None)
        st.error("産地判別の処理中にエラーが発生しました。")
        st.exception(error)


# --------------------------------------------------
# 結果表示・ダウンロード
# --------------------------------------------------

saved = st.session_state.get("ellipse_result")

if saved and saved["signature"] == signature:

    st.success("産地判別が完了しました。")

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "資料数",
        saved["sample_count"],
    )

    col2.metric(
        "測定データ数",
        saved["measurement_count"],
    )

    col3.metric(
        "化学組成グループ数",
        saved["group_count"],
    )

    st.subheader("4．資料ごとの最終判別結果")

    st.dataframe(
        saved["summaries"],
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("5．元素別中央値")

    st.dataframe(
        saved["medians"],
        use_container_width=True,
        hide_index=True,
    )

    with st.expander("測定ごとの判別結果を表示"):
        st.dataframe(
            saved["data"],
            use_container_width=True,
            hide_index=True,
        )

    st.subheader("6．Excelファイルのダウンロード")

    st.download_button(
        label="判別結果Excelをダウンロード",
        data=saved["excel"],
        file_name=saved["filename"],
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
    )
