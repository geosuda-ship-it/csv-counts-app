import streamlit as st

from group_judgment import create_group_judgment


st.set_page_config(
    page_title="各定量値の化学組成グループ判定",
    page_icon="🔬",
    layout="centered",
)

st.title("各定量値の化学組成グループ判定")
st.write(
    "判別楕円パラメーターと定量値ファイルを使用して、"
    "各定量分析値の化学組成グループを判定します。"
)

st.info(
    "入力：判別楕円パラメーター.xlsx ＋ 日付_定量値.xlsx\n\n"
    "出力：日付_判定結果.xlsx"
)

parameter_file = st.file_uploader(
    "1．「判別楕円パラメーター.xlsx」をアップロードしてください。",
    type=["xlsx", "xlsm"],
    key="parameter_file",
)

quantitative_file = st.file_uploader(
    "2．「日付_定量値.xlsx」をアップロードしてください。",
    type=["xlsx", "xlsm"],
    key="quantitative_file",
)

run_button = st.button(
    "グループ判定を実行",
    type="primary",
    use_container_width=True,
)

if run_button:
    st.session_state.pop("group_judgment_result", None)

    if parameter_file is None or quantitative_file is None:
        st.warning("2つのExcelファイルを両方アップロードしてください。")
    else:
        try:
            with st.spinner("化学組成グループを判定しています…"):
                result = create_group_judgment(
                    parameter_content=parameter_file.getvalue(),
                    quantitative_content=quantitative_file.getvalue(),
                    quantitative_file_name=quantitative_file.name,
                )
            st.session_state["group_judgment_result"] = result
        except Exception as error:
            st.error(f"処理できませんでした。\n\n{error}")

result = st.session_state.get("group_judgment_result")
if result is not None:
    st.success("判定が完了しました。")

    column1, column2, column3 = st.columns(3)
    column1.metric("分析値数", result["analysis_count"])
    column2.metric("有効", result["valid_count"])
    column3.metric("無効", result["invalid_count"])

    st.download_button(
        label=f"{result['file_name']}をダウンロード",
        data=result["content"],
        file_name=result["file_name"],
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=True,
    )

