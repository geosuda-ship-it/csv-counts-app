
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

st.subheader("STEP 3: Cluster-based Chemical Group Assignment (Clustering Method) \nクラスター単位の化学組成グループ判定（クラスター法）")

st.write(
    "各資料から得られた複数の定量値の中央値を算出し、Rb分率、Sr分率、Zr分率に基づいてクラスター区分を行います。"
    "次に、各クラスターの分布楕円を作成し、既知の化学組成グループの判別楕円との重なりを比較することで、クラスターごとに化学組成グループを判定します。"
    "STEP 2で出力されたExcelファイル（日付_定量値_ellipse_results.xlsx）をアップロードして解析を実行すると、各資料の化学組成グループの判定結果をExcelファイルとしてダウンロードできます。"
    "出力ファイルには、STEP 2による判定結果も含まれています。"
)

# ------------------------------------------
# 入力ファイル
# ------------------------------------------

st.subheader("解析データのアップロード")

uploaded_file = st.file_uploader(
    "STEP 2で出力されたExcelファイル（日付_定量値_ellipse_results.xlsx）をアップロードしてください",
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

st.subheader("解析の実行")

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
