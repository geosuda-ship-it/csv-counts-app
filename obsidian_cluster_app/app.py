
import hashlib
import tempfile
from pathlib import Path

import streamlit as st

from cluster_analysis import (
    run_analysis,
    export_results_xlsx,
    make_output_name,
)


st.set_page_config(
    page_title="黒曜石のクラスター解析",
    page_icon="🔬",
    layout="wide",
)

st.subheader("Obsidian Provenance Analysis Tools Using Quantitative EDXRF Data")

st.subheader("STEP 3  \nクラスター単位の化学組成グループ判定（クラスター法）")

st.write(
    "各資料から得られた複数の定量値の中央値を算出し、Rb分率、Sr分率、Zr分率に基づいてクラスター区分を行います。"
    "次に、各クラスターの分布楕円を作成し、既知の化学組成グループの判別楕円との重なりを比較することで、クラスターごとに化学組成グループを判定します。"
    "STEP 2で出力されたExcelファイル（日付_定量値_ellipse_results.xlsx）をアップロードして解析を実行すると、各資料の化学組成グループの判定結果をExcelファイルとしてダウンロードできます。"
    "出力ファイルには、STEP 2による判定結果も含まれています。"
)

# ------------------------------------------
# 入力ファイル
# ------------------------------------------

st.subheader("1．解析データ")

uploaded_file = st.file_uploader(
    "STEP 2で出力されたExcelファイルをアップロードしてください",
    type=["xlsx"],
    key="cluster_input",
)

st.caption(
    "Medianシートと，各化学組成グループの"
    "判別楕円パラメーターを含むExcelファイル"
)


# ------------------------------------------
# 解析実行
# ------------------------------------------

st.subheader("1.解析用ファイルのアップロード")

signature = None

if uploaded_file is not None:
    digest = hashlib.sha256()
    digest.update(uploaded_file.name.encode("utf-8"))
    digest.update(uploaded_file.getvalue())
    signature = digest.hexdigest()

if st.button(
    "クラスター解析を実行",
    type="primary",
    disabled=uploaded_file is None,
):
    try:
        with st.spinner("クラスター解析を実行しています..."):

            with tempfile.TemporaryDirectory() as temp_dir:
                temp_dir = Path(temp_dir)

                input_name = Path(uploaded_file.name).name
                input_path = temp_dir / input_name
                input_path.write_bytes(uploaded_file.getvalue())

                results = run_analysis(
                    input_path,
                    min_cluster_size=3,
                    min_samples=3,
                    ellipse_confidence=0.95,
                )


                output_name = make_output_name(input_name)
                output_path = temp_dir / output_name

                export_results_xlsx(results, output_path, input_path)

                excel_bytes = output_path.read_bytes()

            st.session_state["cluster_result"] = {
                "signature": signature,
                "excel": excel_bytes,
                "filename": output_name,
                "summary": results.summary_metrics,
                "counts": results.cluster_counts,
                "top3": results.cluster_top3,
                "samples": results.sample_results,
                "ellipses": results.ellipse_parameters,
                "comparisons": results.all_group_comparisons,
            }

    except Exception as error:
        st.session_state.pop("cluster_result", None)
        st.error("クラスター解析中にエラーが発生しました。")
        st.exception(error)


# ------------------------------------------
# 結果表示
# ------------------------------------------

saved = st.session_state.get("cluster_result")

if saved and saved["signature"] == signature:

    st.success("クラスター解析が完了しました。")

    counts = saved["counts"]

    cluster_count = (
        counts["Cluster"] != "未所属"
    ).sum()

    assigned = counts.loc[
        counts["Cluster"] != "未所属", "資料数"
    ].sum()

    unassigned = counts.loc[
        counts["Cluster"] == "未所属", "資料数"
    ].sum()

    col1, col2, col3 = st.columns(3)

    col1.metric("検出クラスター数", int(cluster_count))
    col2.metric("所属資料数", int(assigned))
    col3.metric("未所属資料数", int(unassigned))

    st.subheader("3．クラスター別資料数")

    st.dataframe(
        counts,
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("4．クラスターごとの上位3候補")

    st.dataframe(
        saved["top3"],
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("5．資料ごとのクラスター判定結果")

    st.dataframe(
        saved["samples"],
        use_container_width=True,
        hide_index=True,
    )

    with st.expander("クラスター楕円パラメーター"):
        st.dataframe(
            saved["ellipses"],
            use_container_width=True,
            hide_index=True,
        )

    with st.expander("全化学組成グループとの比較結果"):
        st.dataframe(
            saved["comparisons"],
            use_container_width=True,
            hide_index=True,
        )

    with st.expander("解析条件・概要"):
        st.dataframe(
            saved["summary"],
            use_container_width=True,
            hide_index=True,
        )

    st.subheader("6．解析結果のダウンロード")

    st.download_button(
        label="クラスター解析結果Excelをダウンロード",
        data=saved["excel"],
        file_name=saved["filename"],
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
    )
