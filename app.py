import streamlit as st

from group_judgment import create_group_judgment


# ============================================================
# ページ設定
# ============================================================

st.set_page_config(
    page_title="各定量値の化学組成グループ判定",
    page_icon="🪨",
    layout="centered",
)


# ============================================================
# タイトル・説明
# ============================================================

st.title("各定量値の化学組成グループ判定")

st.write(
    "判別楕円パラメータと定量値ファイルを使用して，"
    "各定量分析値の化学組成グループを判定します。"
)

st.info(
    "入力：判別楕円パラメータ.xlsx ＋ 日付_定量値.xlsx\n\n"
    "出力：日付_判定結果.xlsx"
)


# ============================================================
# 1．判別楕円パラメータ
# ============================================================

st.write(
    '1. 「判別楕円パラメータ.xlsx」をアップロードしてください。'
)

st.caption(
    "例：長崎大学_判別楕円パラメータ_ver1.xlsx"
)

parameter_file = st.file_uploader(
    "判別楕円パラメータ",
    type=["xlsx", "xlsm"],
    key="parameter",
    label_visibility="collapsed",
)


# ============================================================
# 2．定量値ファイル
# ============================================================

st.write(
    "2. 強度から定量値への変換ツールで作成した"
    "定量値ファイルをアップロードしてください。"
)

st.caption(
    "例：20260803_定量値.xlsx"
)

quantitative_file = st.file_uploader(
    "定量値ファイル",
    type=["xlsx", "xlsm"],
    key="quantitative",
    label_visibility="collapsed",
)


# ============================================================
# 3．実行
# ============================================================

st.write(
    "3. 2つのファイルを選択した後，"
    "「グループ判定を実行」ボタンを押してください。"
)

run_button = st.button(
    "グループ判定を実行",
    use_container_width=True,
    type="primary",
)


# ============================================================
# 判定処理
# ============================================================

if run_button:

    if parameter_file is None:
        st.error(
            "「判別楕円パラメータ.xlsx」をアップロードしてください。"
        )
        st.stop()

    if quantitative_file is None:
        st.error(
            "強度から定量値への変換ツールで作成した"
            "定量値ファイルをアップロードしてください。"
        )
        st.stop()

    try:

        with st.spinner(
            "化学組成グループを判定しています..."
        ):

            output = create_group_judgment(
                parameter_file,
                quantitative_file,
                quantitative_file.name,
            )

        st.success(
            "化学組成グループの判定が完了しました。"
        )

        st.download_button(
            label="判定結果をダウンロード",
            data=output,
            file_name=output.name,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
            type="primary",
        )

    except Exception as error:

        st.error(
            "判定処理中にエラーが発生しました。"
        )

        st.exception(error)
