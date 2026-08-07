import streamlit as st

from parameterized_quantitative import create_quantitative_excel


EXCEL_MIME = (
    "application/vnd.openxmlformats-officedocument."
    "spreadsheetml.sheet"
)

st.set_page_config(
    page_title="強度から定量値への変換",
    page_icon="🧪",
    layout="centered",
)

st.title("強度から定量値への変換")
st.write(
    "機関ごとの定量計算パラメーターを使用して、"
    "EDXの強度データから定量値を計算します。"
)

st.info(
    "入力：定量計算パラメーター.xlsx ＋ 日付.csv\n\n"
    "出力：日付_定量値.xlsx"
)

parameter_file = st.file_uploader(
    "1．「定量計算パラメーター.xlsx」をアップロードしてください。",
    type=["xlsx", "xlsm"],
    key="quantitative_parameter_file",
)

csv_file = st.file_uploader(
    "2．「日付.csv」をアップロードしてください。",
    type=["csv"],
    key="quantitative_csv_file",
)

run_button = st.button(
    "定量値を計算",
    type="primary",
    use_container_width=True,
)

if run_button:
    st.session_state.pop("quantitative_result", None)

    if parameter_file is None or csv_file is None:
        st.warning("パラメーターExcelとCSVを両方アップロードしてください。")
    else:
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

