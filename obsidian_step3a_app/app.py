from pathlib import Path
from tempfile import TemporaryDirectory

import streamlit as st

from step3A_cluster_judgment import analyze


st.set_page_config(
    page_title="STEP 3A: Chemical Group Assignment",
    page_icon="📊",
    layout="centered",
)


# --------------------------------------------------
# 表示設定
# --------------------------------------------------

st.markdown(
    """
    <style>
    .block-container {
        max-width: 760px;
        padding-top: 3rem;
        padding-bottom: 4rem;
    }

    .main-title {
        text-align: center;
        font-size: 2rem;
        font-weight: 700;
        line-height: 1.2;
        margin-bottom: 2.2rem;
        color: #30313d;
    }

    .step-title {
        font-size: 1.75rem;
        font-weight: 700;
        line-height: 1.25;
        margin-top: 1rem;
        margin-bottom: 1rem;
        color: #30313d;
    }

    .japanese-subtitle {
        font-size: 1rem;
        margin-bottom: 1.25rem;
        color: #30313d;
    }

    .description {
        font-size: 1rem;
        line-height: 1.8;
        color: #30313d;
        margin-bottom: 2rem;
    }

    .section-title {
        font-size: 1.55rem;
        font-weight: 700;
        margin-top: 2rem;
        margin-bottom: 0.75rem;
        color: #30313d;
    }

    .upload-description {
        font-size: 0.92rem;
        line-height: 1.6;
        margin-bottom: 0.3rem;
        color: #30313d;
    }

    .small-note {
        font-size: 0.85rem;
        color: #8a8d98;
        margin-top: 0.3rem;
        margin-bottom: 1rem;
    }

    div[data-testid="stFileUploader"] {
        margin-bottom: 0.3rem;
    }

    div[data-testid="stFileUploaderDropzone"] {
        background-color: #f0f2f6;
        border: none;
        border-radius: 10px;
        padding: 0.8rem;
    }

    div.stButton > button {
        min-height: 2.8rem;
        font-size: 1rem;
    }

    div.stDownloadButton > button {
        min-height: 2.8rem;
        font-size: 1rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# --------------------------------------------------
# タイトルと説明
# --------------------------------------------------

st.markdown(
    """
    <div class="main-title">
        Obsidian Provenance Analysis Tools<br>
        Using Quantitative EDXRF Data
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="step-title">
        STEP 3A: Cluster-based Chemical Group Assignment<br>
        (Manual Clustering Method)
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="japanese-subtitle">
        クラスター単位の化学組成グループ判定（手動クラスター法）
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="description">
        手動によるクラスター区分結果と、既知の化学組成グループの判別楕円を
        比較します。6種類の散布図（D1〜D6）における楕円の重なりを解析し、
        各クラスターおよび各試料に対する化学組成グループの候補を判定します。
        <br><br>
        解析結果は、「候補一覧」「試料別判定」「重なり結果」「Median」の
        4シートを含むExcelファイルとしてダウンロードできます。
    </div>
    """,
    unsafe_allow_html=True,
)


# --------------------------------------------------
# 1. クラスター区分ファイル
# --------------------------------------------------

st.markdown(
    '<div class="section-title">1. Upload Cluster Analysis Data</div>',
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="upload-description">
        手動で作成したクラスター区分Excelファイルをアップロードしてください。
    </div>
    """,
    unsafe_allow_html=True,
)

cluster_file = st.file_uploader(
    "クラスター区分Excel",
    type=["xlsx"],
    key="cluster_file",
    label_visibility="collapsed",
)

st.markdown(
    """
    <div class="small-note">
        「散布図」「DATA」「Median」シートを含むExcelファイル
    </div>
    """,
    unsafe_allow_html=True,
)


# --------------------------------------------------
# 2. 判別楕円パラメータファイル
# --------------------------------------------------

st.markdown(
    '<div class="section-title">2. Upload Ellipse Parameters</div>',
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="upload-description">
        化学組成グループの判別楕円パラメータファイルをアップロードしてください。
    </div>
    """,
    unsafe_allow_html=True,
)

group_file = st.file_uploader(
    "判別楕円パラメータExcel",
    type=["xlsx"],
    key="group_file",
    label_visibility="collapsed",
)

st.markdown(
    """
    <div class="small-note">
        「AS-1」以降の判別楕円パラメータシートを含むExcelファイル
    </div>
    """,
    unsafe_allow_html=True,
)


# --------------------------------------------------
# 3. 解析
# --------------------------------------------------

st.markdown(
    '<div class="section-title">3. Run Analysis</div>',
    unsafe_allow_html=True,
)

ready = cluster_file is not None and group_file is not None

if st.button(
    "クラスター判定を実行",
    type="primary",
    disabled=not ready,
):
    try:
        with st.spinner("楕円の重なりを解析しています…"):
            with TemporaryDirectory() as temporary_directory:
                work_dir = Path(temporary_directory)

                cluster_path = work_dir / cluster_file.name
                group_path = work_dir / group_file.name

                cluster_path.write_bytes(cluster_file.getvalue())
                group_path.write_bytes(group_file.getvalue())

                output_path = analyze(
                    cluster_path=cluster_path,
                    group_path=group_path,
                    output_dir=work_dir,
                    quadrature_points=512,
                    tolerance_pct=99.5,
                )

                st.session_state["result_bytes"] = output_path.read_bytes()
                st.session_state["result_name"] = output_path.name

        st.success("解析が完了しました。")

    except ValueError as error:
        st.session_state.pop("result_bytes", None)
        st.session_state.pop("result_name", None)
        st.error(str(error))

    except Exception as error:
        st.session_state.pop("result_bytes", None)
        st.session_state.pop("result_name", None)
        st.error("予期しないエラーが発生しました。")

        with st.expander("エラーの詳細"):
            st.exception(error)


# --------------------------------------------------
# 結果のダウンロード
# --------------------------------------------------

if "result_bytes" in st.session_state:
    st.download_button(
        label="判定結果Excelをダウンロード",
        data=st.session_state["result_bytes"],
        file_name=st.session_state["result_name"],
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
    )
