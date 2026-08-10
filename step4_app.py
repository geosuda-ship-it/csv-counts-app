"""Streamlit用：STEP 4 判定結果を判別図上で確認する。"""

import streamlit as st

from discrimination_diagram import create_discrimination_diagram


PNG_MIME = "image/png"

st.set_page_config(
    page_title="判定結果を判別図上で確認する",
    page_icon="📊",
    layout="centered",
)

st.title("STEP 4　判定結果を判別図上で確認する")

st.markdown(
    "STEP 1で作成した定量値の計算結果ファイル"
    "（例：20260728_定量値.xlsx）と、共通の判別楕円"
    "パラメーターファイル（例："
    "[長崎大学_判別楕円パラメータ_ver1.xlsx]"
    "(https://docs.google.com/spreadsheets/d/"
    "1JEbuQlqNbXQJIX-iEROrPR8X3t0d2y3xqJG_mu0-jac/"
    "edit?usp=sharing)）をアップロードすると、"
    "Mandatoryの判別楕円と各定量分析値を重ねた判別図が表示されます。"
    "ValidとInvalidは異なるシンボルで表示され、"
    "作成した判別図をPNG画像としてダウンロードできます。"
)

st.info(
    "Upload：日付_定量値.xlsx ＋ 判別楕円パラメータ.xlsx\n\n"
    "Display：Mandatoryの判別楕円と各定量分析値を重ねた判別図\n\n"
    "Download：日付_判別図.png"
)

st.write(
    "1. STEP 1で作成した定量値の計算結果ファイルを"
    "アップロードしてください。"
)
st.caption("例：20260728_定量値.xlsx")
quantitative_file = st.file_uploader(
    "定量値ファイル",
    type=["xlsx", "xlsm"],
    key="step4_quantitative_file",
    label_visibility="collapsed",
)

st.write(
    "2. 共通の判別楕円パラメーターファイルを"
    "アップロードしてください。"
)
st.caption("例：長崎大学_判別楕円パラメータ_ver1.xlsx")
parameter_file = st.file_uploader(
    "判別楕円パラメーターファイル",
    type=["xlsx", "xlsm"],
    key="step4_parameter_file",
    label_visibility="collapsed",
)

st.write(
    "3. 2つのファイルを選択した後、"
    "「判別図を作成」ボタンを押してください。"
)
run_button = st.button(
    "判別図を作成",
    type="primary",
    use_container_width=True,
)

if run_button:
    st.session_state.pop("step4_result", None)

    if quantitative_file is None or parameter_file is None:
        st.warning("2つのExcelファイルを両方アップロードしてください。")
    else:
        try:
            with st.spinner(
                "判定を実行し、Mandatoryの判別図を作成しています…"
            ):
                result = create_discrimination_diagram(
                    parameter_content=parameter_file.getvalue(),
                    quantitative_content=quantitative_file.getvalue(),
                    quantitative_file_name=quantitative_file.name,
                )
            st.session_state["step4_result"] = result
        except Exception as error:
            st.error("判別図を作成できませんでした。")
            st.exception(error)

result = st.session_state.get("step4_result")
if result is not None:
    st.success("判別図の作成が完了しました。")

    column1, column2, column3 = st.columns(3)
    column1.metric("定量分析値数", result["analysis_count"])
    column2.metric("Valid", result["valid_count"])
    column3.metric("Invalid", result["invalid_count"])

    st.caption(
        f"化学組成グループ数：{result['group_count']} ／ "
        "表示対象：Mandatory（D1～D3）"
    )

    st.image(
        result["content"],
        caption="Mandatoryの判別楕円と各定量分析値を重ねた判別図",
        use_container_width=True,
    )

    st.download_button(
        label=f"{result['file_name']}をダウンロード",
        data=result["content"],
        file_name=result["file_name"],
        mime=PNG_MIME,
        type="primary",
        use_container_width=True,
    )
