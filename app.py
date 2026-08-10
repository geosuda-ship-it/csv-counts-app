import streamlit as st

from final_judgment import create_final_judgment


st.set_page_config(
    page_title="各サンプルの最終判定",
    page_icon="🪨",
    layout="centered",
)

st.title("各サンプルの最終判定")

st.write(
    "各定量値のグループ判定結果をサンプルごとに集約し，"
    "最終判定結果と集計済みシートを作成します。"
)

st.info(
    "入力：日付_判定結果.xlsx ＋ 最終判定集計シート.xlsx\n\n"
    "出力：日付_最終判定結果.xlsx ＋ "
    "日付_最終判定集計シート.xlsx"
)


# ============================================================
# 1．判定結果ファイル
# ============================================================

st.write(
    "1. STEP 2で作成した判定結果ファイルを"
    "アップロードしてください。"
)

st.caption("例：20260728_判定結果.xlsx")

judgment_file = st.file_uploader(
    "判定結果ファイル",
    type=["xlsx", "xlsm"],
    key="judgment",
    label_visibility="collapsed",
)


# ============================================================
# 2．最終判定集計シート
# ============================================================

st.write(
    "2. 共通の「最終判定集計シート.xlsx」を"
    "アップロードしてください。"
)

st.caption("例：最終判定集計シート_ver1.xlsx")

template_file = st.file_uploader(
    "最終判定集計シート",
    type=["xlsx", "xlsm"],
    key="template",
    label_visibility="collapsed",
)


# ============================================================
# 3．実行
# ============================================================

st.write(
    "3. 2つのファイルを選択した後，"
    "「最終判定を実行」ボタンを押してください。"
)

run_button = st.button(
    "最終判定を実行",
    use_container_width=True,
    type="primary",
)


# ============================================================
# 最終判定処理
# ============================================================

if run_button:

    if judgment_file is None:
        st.error(
            "STEP 2で作成した判定結果ファイルを"
            "アップロードしてください。"
        )
        st.stop()

    if template_file is None:
        st.error(
            "共通の「最終判定集計シート.xlsx」を"
            "アップロードしてください。"
        )
        st.stop()

    try:
        with st.spinner(
            "サンプルごとの最終判定と集計を行っています..."
        ):
            output = create_final_judgment(
                judgment_file.getvalue(),
                judgment_file.name,
                template_file.getvalue(),
                template_file.name,
            )

        st.success("最終判定と集計が完了しました。")

        metric_columns = st.columns(3)
        metric_columns[0].metric(
            "サンプル数",
            output["sample_count"],
        )
        metric_columns[1].metric(
            "判定数",
            output["judged_count"],
        )
        metric_columns[2].metric(
            "判定不能",
            output["unclassifiable_count"],
        )

        st.download_button(
            label=(
                f'{output["judgment_file_name"]}をダウンロード'
            ),
            data=output["judgment_content"],
            file_name=output["judgment_file_name"],
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
            type="primary",
        )

        st.download_button(
            label=(
                f'{output["summary_file_name"]}をダウンロード'
            ),
            data=output["summary_content"],
            file_name=output["summary_file_name"],
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
            type="primary",
        )

    except Exception as error:
        st.error("最終判定処理中にエラーが発生しました。")
        st.exception(error)
