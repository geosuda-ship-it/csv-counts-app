import streamlit as st

from parameterized_quantitative import create_quantitative_excel


EXCEL_MIME = (
    "application/vnd.openxmlformats-officedocument."
    "spreadsheetml.sheet"
)

st.set_page_config(
    page_title="STEP 1　強度データから定量値を作成する",
    page_icon="🧪",
    layout="centered",
)

st.title("STEP 1　強度データから定量値を作成する")
st.write(
    "NEX DEから出力したCSV形式の測定結果ファイルと、"
    "機関ごとの定量計算パラメーターファイルをアップロードすると、"
    "定量値の計算結果をExcelファイルとしてダウンロードできます。"
)

st.info(
    "入力：日付.csv ＋ 機関名_定量計算パラメーター.xlsx\n\n"
    "出力：日付_定量値.xlsx"
)

st.markdown(
    "1. NEX DEから出力したCSV形式の測定結果ファイルを"
    "アップロードしてください。"
)
st.caption("例：20260728.csv")
csv_file = st.file_uploader(
    "NEX DEの強度データファイル",
    type=["csv"],
    key="quantitative_csv_file",
    label_visibility="collapsed",
)

st.markdown(
    "2. 機関ごとの定量計算パラメーターファイルを"
    "アップロードしてください。"
)
st.caption("例：長崎大学_定量計算パラメーター_ver1.xlsx")
parameter_file = st.file_uploader(
    "定量計算パラメーターファイル",
    type=["xlsx", "xlsm"],
    key="quantitative_parameter_file",
    label_visibility="collapsed",
)

st.markdown(
    "3. 2つのファイルを選択した後、"
    "「定量値を計算」ボタンを押してください。"
)
run_button = st.button(
    "定量値を計算",
    type="primary",
    use_container_width=True,
)

if run_button:
    st.session_state.pop("quantitative_result", None)

    if csv_file is None:
        st.warning(
            "NEX DEから出力したCSV形式の測定結果ファイルを"
            "アップロードしてください。"
        )
        st.stop()

    if parameter_file is None:
        st.warning(
            "機関ごとの定量計算パラメーターファイルを"
            "アップロードしてください。"
        )
        st.stop()

    try:
        with st.spinner("ドリフト補正と定量計算を実行しています…"):
            result = create_quantitative_excel(
                parameter_content=parameter_file.getvalue(),
                csv_content=csv_file.getvalue(),
                csv_file_name=csv_file.name,
            )
        st.session_state["quantitative_result"] = result
    except Exception as error:
        st.error(f"処理できませんでした。\n\n{error}")

result = st.session_state.get("quantitative_result")
if result is not None:
    st.success("定量値の計算が完了しました。")

    column1, column2, column3 = st.columns(3)
    column1.metric("機関", result["institution"] or "未記載")
    column2.metric("バージョン", result["version"] or "未記載")
    column3.metric("出力行数", result["output_row_count"])

    st.caption(
        f"QC試料：{result['qc_sample_name']} ／ "
        f"測定モード：{result['qc_mode']}"
    )

    if result["missing_dates"]:
        missing_text = "、".join(
            date_value.isoformat()
            for date_value in result["missing_dates"]
        )
        st.warning(
            "同日のQCがなく、出力から除外した測定日："
            + missing_text
        )

    st.download_button(
        label=f"{result['file_name']}をダウンロード",
        data=result["content"],
        file_name=result["file_name"],
        mime=EXCEL_MIME,
        type="primary",
        use_container_width=True,
    )
