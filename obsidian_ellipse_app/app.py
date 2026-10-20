from io import BytesIO

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Obsidian EDXRF Ellipse Method",
    layout="wide",
)

st.title("判別楕円法による黒曜石の化学組成グループ候補の評価")

uploaded = st.file_uploader(
    "① 測定データのExcelファイルをアップロードしてください。",
    type=["xlsx"],
)

if uploaded is not None:
    try:
        data = pd.read_excel(
            BytesIO(uploaded.getvalue()),
            sheet_name="定量値",
            dtype={"sample": "string"},
            engine="openpyxl",
        ).dropna(how="all").copy()

        if data.empty:
            raise ValueError("定量値シートに測定データがありません。")

        if "sample" not in data.columns:
            raise ValueError("定量値シートにsample列がありません。")

        # 元の資料IDを残す
        data["sample_元表記"] = data["sample"].copy()

        # アポストロフィを無視して資料IDを統一
        data["sample"] = (
            data["sample"]
            .str.replace("'", "", regex=False)
            .str.strip()
        )

        if data["sample"].isna().any() or data["sample"].eq("").any():
            raise ValueError("資料IDが空白の測定データがあります。")

    except Exception as exc:
        st.error(f"読み込めませんでした：{exc}")
        st.stop()

    sample_counts = (
        data.groupby("sample", sort=False)
        .size()
        .reset_index(name="測定回数")
    )

    st.success(
        f"{len(sample_counts)}資料・{len(data)}測定を読み込みました。"
    )

    st.subheader("資料別の測定回数")
    st.dataframe(sample_counts, hide_index=True, width="stretch")

    with st.expander("測定データを確認する"):
        st.dataframe(data, hide_index=True, width="stretch")
