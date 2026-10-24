from pathlib import Path
from tempfile import TemporaryDirectory

import streamlit as st

from step3A_cluster_judgment import analyze


st.set_page_config(
    page_title="STEP 3A クラスター判定",
    page_icon="📊",
    layout="centered",
)

st.title("STEP 3A クラスター判定")

st.write(
    "クラスター区分Excelと判別楕円パラメータExcelを比較し、"
    "クラスターおよび試料ごとの化学組成グループ判定結果を作成します。"
)

st.info(
    "アップロードしたファイルは解析処理にのみ使用されます。"
    "クラスター区分Excelには「散布図」「DATA」「Median」シートが必要です。"
)

cluster_file = st.file_uploader(
    "1. クラスター区分Excel",
    type=["xlsx"],
    help="「散布図」「DATA」「Median」シートを含むExcelファイル",
)

group_file = st.file_uploader(
    "2. 判別楕円パラメータExcel",
    type=["xlsx"],
    help="「AS-1」シートを含む判別楕円パラメータファイル",
)

ready = cluster_file is not None and group_file is not None

if st.button(
    "解析を実行",
    type="primary",
    disabled=not ready,
    use_container_width=True,
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

                result_bytes = output_path.read_bytes()
                result_name = output_path.name

        st.success("解析が完了しました。")

        st.download_button(
            label="判定結果Excelをダウンロード",
            data=result_bytes,
            file_name=result_name,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary",
            use_container_width=True,
        )

    except ValueError as error:
        st.error(str(error))

    except Exception as error:
        st.error(
            "予期しないエラーが発生しました。"
            "アップロードしたExcelの内容を確認してください。"
        )
        with st.expander("エラーの詳細"):
            st.exception(error)
