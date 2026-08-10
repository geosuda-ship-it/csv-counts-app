import streamlit as st

from final_judgment import create_final_judgment


EXCEL_MIME = (
    "application/vnd.openxmlformats-officedocument."
    "spreadsheetml.sheet"
)

st.set_page_config(
    page_title="各サンプルの最終判定",
    page_icon="🔬",
    layout="centered",
)

st.title("各サンプルの最終判定")
st.write(
    "STEP 2で作成した判定結果ファイルと、"
    "共通のグループ集計シートを使用して、"
    "各サンプルの最終判定と化学組成グループの集計を行います。"
)

st.info(
    "入力：日付_判定結果.xlsx ＋ グループ集計シート.xlsx\n\n"
    "出力：日付_最終判定結果.xlsx ＋ "
    "日付_グループ集計結果.xlsx"
)

st.write(
    "1. STEP 2で作成した判定結果ファイルを"
    "アップロードしてください。"
)
st.caption("例：20260728_判定結果.xlsx")
judgment_file = st.file_uploader(
    "判定結果ファイル",
    type=["xlsx", "xlsm"],
    key="final_judgment_input",
    label_visibility="collapsed",
)

st.write(
    "2. 共通の「グループ集計シート.xlsx」を"
    "アップロードしてください。"
)
st.caption("例：グループ集計シート_ver1.xlsx")
template_file = st.file_uploader(
    "グループ集計シート",
    type=["xlsx", "xlsm"],
    key="final_summary_template",
    label_visibility="collapsed",
)

st.write(
    "3. 2つのファイルを選択した後、"
    "「最終判定を実行」ボタンを押してください。"
)

run_button = st.button(
    "最終判定を実行",
    type="primary",
    use_container_width=True,
)

if run_button:
    st.session_state.pop("final_judgment_result", None)

    if judgment_file is None:
        st.warning(
            "STEP 2で作成した判定結果ファイルを"
            "アップロードしてください。"
        )
        st.stop()

    if template_file is None:
        st.warning(
            "共通の「グループ集計シート.xlsx」を"
            "アップロードしてください。"
        )
        st.stop()

    try:
        with st.spinner("サンプルごとの最終判定を計算しています…"):
            result = create_final_judgment(
                judgment_content=judgment_file.getvalue(),
                judgment_file_name=judgment_file.name,
                template_content=template_file.getvalue(),
                template_file_name=template_file.name,
            )
        st.session_state["final_judgment_result"] = result
    except Exception as error:
        st.error(f"処理できませんでした。\n\n{error}")

result = st.session_state.get("final_judgment_result")
if result is not None:
    st.success("最終判定と集計が完了しました。")

    column1, column2, column3 = st.columns(3)
    column1.metric("サンプル数", result["sample_count"])
    column2.metric("判定数", result["judged_count"])
    column3.metric("判定不能", result["unclassifiable_count"])

    st.download_button(
        label=f"{result['judgment_file_name']}をダウンロード",
        data=result["judgment_content"],
        file_name=result["judgment_file_name"],
        mime=EXCEL_MIME,
        type="primary",
        use_container_width=True,
    )

    st.download_button(
        label=f"{result['summary_file_name']}をダウンロード",
        data=result["summary_content"],
        file_name=result["summary_file_name"],
        mime=EXCEL_MIME,
        type="primary",
        use_container_width=True,
    )
