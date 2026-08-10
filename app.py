import streamlit as st

from group_judgment import create_group_judgment


st.set_page_config(
    page_title="各定量値の化学組成グループ判定",
    page_icon="🪨",
    layout="centered",
)

st.title("各定量値の化学組成グループ判定")

st.write(
    "判別楕円パラメータと定量値ファイルを使用して，各定量分析値の"
    "化学組成グループを判定します。"
)

st.info(
    "入力：判別楕円パラメータ.xlsx ＋ 日付_定量値.xlsx\n\n"
    "出力：日付_判定結果.xlsx"
)

# ------------------------------------------------------
# 判別楕円パラメータ
# ------------------------------------------------------

st.markdown(
    """
### 1. 「判別楕円パラメータ.xlsx」をアップロードしてください。

例：長崎大学_判別楕円パラメータ_ver1.xlsx
"""
)

parameter_file = st.file_uploader(
    "",
    type=["xlsx", "xlsm"],
    key="parameter",
)

# ------------------------------------------------------
# 定量値
# ------------------------------------------------------

st.markdown(
    """
### 2. 強度から定量値への変換ツールで作成した定量値ファイルをアップロードしてください。

例：20260803_定量値.xlsx
"""
)

quantitative_file = st.file_uploader(
    "",
    type=["xlsx", "xlsm"],
    key="quantitative",
)

# ------------------------------------------------------
# 実行
# ------------------------------------------------------

st.markdown(
    """
### 3. 2つのファイルを選択した後，「グループ判定を実行」ボタンを押してください。
"""
)

if st.button(
    "グループ判定を実行",
    use_container_width=True,
):

    if parameter_file is None:
        st.error("判別楕円パラメータを選択してください。")
        st.stop()

    if quantitative_file is None:
        st.error("定量値ファイルを選択してください。")
        st.stop()

    with st.spinner("判定中です..."):

        output = create_group_judgment(
            parameter_file,
            quantitative_file,
        )

    st.success("判定が完了しました。")

    st.download_button(
        label="判定結果をダウンロード",
        data=output,
        file_name=output.name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
