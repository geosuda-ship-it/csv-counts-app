import hashlib
import tempfile
from pathlib import Path

import streamlit as st

from step3A_cluster_judgment import analyze


st.set_page_config(
    page_title="Manual Clustering Method",
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
    "STEP 3: Cluster-based Chemical Group Assignment "
    "(Manual Clustering Method)  \n"
    "クラスター単位の化学組成グループ判定（手動クラスター法）"
)

st.write(
    "手動によるクラスター解析結果のExcelファイルと、既知の化学組成グループの"
    "判別楕円パラメータファイルをアップロードします。"
    "6種類の散布図（D1〜D6）において、各クラスターグループの分布楕円と"
    "既知の化学組成グループの判別楕円との重なりを比較することで、"
    "クラスターごと、および各試料の化学組成グループ候補を判定します。"
    "解析結果はExcelファイルとしてダウンロードできます。"
    "出力ファイルには、「候補一覧」「試料別判定」「重なり結果」および"
    "入力ファイルの「Median」シートが含まれます。"
)


# ------------------------------------------
# 入力ファイル
# ------------------------------------------

st.subheader("1. Upload Analysis Data")

cluster_file = st.file_uploader(
    "手動によるクラスター区分結果のExcelファイル（例：遺跡名_step3_クラスター区分_ver1.xlsx）をアップロードしてください",
    type=["xlsx"],
    key="step3a_cluster_input",
)

st.caption(
    "「散布図」「DATA」「Median」シートを含むExcelファイル"
)

group_file = st.file_uploader(
    "判別楕円パラメータファイル（例：長崎大_判別楕円パラメーター_ver24.xlsx）をアップロードしてください",
    type=["xlsx"],
    key="step3a_ellipse_input",
)

st.caption(
    "「AS-1」以降の各化学組成グループの"
    "判別楕円パラメータを含むExcelファイル"
)


# ------------------------------------------
# 解析実行
# ------------------------------------------

st.subheader("2. Run Analysis")

signature = None

if cluster_file is not None and group_file is not None:
    digest = hashlib.sha256()

    digest.update(cluster_file.name.encode("utf-8"))
    digest.update(cluster_file.getvalue())

    digest.update(group_file.name.encode("utf-8"))
    digest.update(group_file.getvalue())

    signature = digest.hexdigest()

if st.button(
    "クラスター判定を実行",
    type="primary",
    disabled=cluster_file is None or group_file is None,
):
    try:
        with st.spinner("クラスター判定を実行しています..."):

            with tempfile.TemporaryDirectory() as temp_dir:
                temp_dir = Path(temp_dir)

                cluster_name = Path(cluster_file.name).name
                group_name = Path(group_file.name).name

                cluster_path = temp_dir / cluster_name
                group_path = temp_dir / group_name

                cluster_path.write_bytes(cluster_file.getvalue())
                group_path.write_bytes(group_file.getvalue())

                output_path = analyze(
                    cluster_path=cluster_path,
                    group_path=group_path,
                    output_dir=temp_dir,
                    quadrature_points=512,
                    tolerance_pct=99.5,
                )

                excel_bytes = output_path.read_bytes()
                output_name = output_path.name

            st.session_state["step3a_result"] = {
                "signature": signature,
                "excel": excel_bytes,
                "filename": output_name,
            }

    except Exception as error:
        st.session_state.pop("step3a_result", None)
        st.error("クラスター判定中にエラーが発生しました。")
        st.exception(error)


# ------------------------------------------
# 結果表示
# ------------------------------------------

saved = st.session_state.get("step3a_result")

if saved and saved["signature"] == signature:

    st.success("クラスター判定が完了しました。")

    st.download_button(
        label="クラスター判定結果Excelをダウンロード",
        data=saved["excel"],
        file_name=saved["filename"],
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
    )
