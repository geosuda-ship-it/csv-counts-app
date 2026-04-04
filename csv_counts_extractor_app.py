import csv
from io import BytesIO, StringIO

import pandas as pd
import streamlit as st

st.set_page_config(page_title="CSV counts extractor", page_icon="🧪", layout="wide")


def read_uploaded_csv(uploaded_file):
    raw = uploaded_file.getvalue()

    encodings = ["cp932", "utf-8-sig", "utf-8"]
    for enc in encodings:
        try:
            text = raw.decode(enc)
            return list(csv.reader(StringIO(text)))
        except Exception:
            pass
    raise ValueError("CSVを読み込めませんでした")


def find_counts_section(rows):
    for i, row in enumerate(rows):
        count_hits = sum(1 for cell in row if str(cell).strip().lower() == "counts")
        if count_hits >= 3:
            return i - 2, i - 1, i, i + 1
    raise ValueError("counts行が見つかりません")


def build_counts_dataframe(rows):
    zone_i, line_i, unit_i, data_i = find_counts_section(rows)

    line_row = rows[line_i]
    unit_row = rows[unit_i]

    cols = []
    headers = []

    for i in range(3, len(unit_row)):
        if str(unit_row[i]).strip().lower() == "counts":
            cols.append(i)
            headers.append(line_row[i].strip() if i < len(line_row) else f"counts_{i}")

    records = []

    for row in rows[data_i:]:
        if not any(str(x).strip() for x in row):
            continue

        rec = {
            "試料番号": row[0].lstrip("'") if len(row) > 0 else "",
            "測定日": row[2] if len(row) > 2 else "",
        }

        for c, h in zip(cols, headers):
            try:
                rec[h] = float(row[c])
            except Exception:
                rec[h] = None

        records.append(rec)

    df = pd.DataFrame(records)
    if "測定日" in df.columns:
        df["測定日"] = pd.to_datetime(df["測定日"], errors="coerce")

    return df


def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output


st.title("CSV counts extractor")

uploaded = st.file_uploader("CSVをドロップ", type=["csv"])

if uploaded:
    rows = read_uploaded_csv(uploaded)
    df = build_counts_dataframe(rows)

    st.dataframe(df, use_container_width=True)

    excel = to_excel(df)

    st.download_button(
        label="Excelを保存",
        data=excel,
        file_name="counts_only.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )