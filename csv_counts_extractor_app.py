import csv
from io import StringIO, BytesIO

import pandas as pd
import streamlit as st

st.set_page_config(page_title="CSV counts extractor", page_icon="🧪", layout="wide")

st.title("CSV counts extractor")
st.write("CSVをアップロードすると，counts ブロックを整理して表にします。")


def read_uploaded_csv(uploaded_file):
    raw = uploaded_file.getvalue()

    encodings = ["cp932", "utf-8-sig", "utf-8"]
    last_error = None

    for enc in encodings:
        try:
            text = raw.decode(enc, errors="replace")
            return list(csv.reader(StringIO(text)))
        except Exception as e:
            last_error = e

    raise ValueError(f"CSVを読み込めませんでした: {last_error}")


def build_dataframe_from_counts_blocks(rows):
    # -----------------------------
    # counts 行をすべて取得
    # -----------------------------
    counts_indices = [
        i for i, row in enumerate(rows)
        if "counts" in [str(x).strip().lower() for x in row]
    ]

    if not counts_indices:
        raise ValueError("counts 行が見つかりません")

    # -----------------------------
    # 全ブロックを格納
    # -----------------------------
    all_data = []
    columns = None

    for idx in counts_indices:
        if idx < 2:
            continue

        # ヘッダー（2段）
        header1 = rows[idx - 2]
        header2 = rows[idx - 1]

        # 列名作成（最初だけ使う）
        if columns is None:
            new_cols = []
            for h1, h2 in zip(header1[3:], header2[3:]):
                h1 = str(h1).strip()
                h2 = str(h2).strip()
                if h1 and h2:
                    new_cols.append(f"{h1}_{h2}")
                elif h2:
                    new_cols.append(h2)
                else:
                    new_cols.append(h1)

            columns = ["Sample", "Method", "Date"] + new_cols

        # データ取得（countsの下から次の空行まで）
        for row in rows[idx + 1:]:
            if len(row) == 0 or all(str(x).strip() == "" for x in row):
                break

            # 列数をそろえる
            if len(row) < len(columns):
                row = row + [""] * (len(columns) - len(row))
            elif len(row) > len(columns):
                row = row[:len(columns)]

            all_data.append(row)

    if not all_data:
        raise ValueError("counts データが見つかりませんでした")

    # -----------------------------
    # DataFrame化
    # -----------------------------
    df = pd.DataFrame(all_data, columns=columns)

    # 数値変換
    for col in columns[3:]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # 日付変換
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

    # 並び順を定義
    order = {
        "obart_prec2_air": 0,
        "obart_stand2_air": 1,
        "obart_quick2_air": 2
        
    }

    # 並び替え
    df["sort_key"] = df["Method"].map(order).fillna(999)
    df = df.sort_values(["sort_key", "Sample"], kind="stable").drop(columns="sort_key")

    return df


def dataframe_to_csv_bytes(df):
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def dataframe_to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="cleaned_counts")
    output.seek(0)
    return output.getvalue()


uploaded = st.file_uploader("CSVをドロップ", type=["csv"])

if uploaded is not None:
    try:
        rows = read_uploaded_csv(uploaded)
        df = build_dataframe_from_counts_blocks(rows)

        st.success(f"処理成功: {len(df)} 行 × {len(df.columns)} 列")
        st.dataframe(df, use_container_width=True)

        base_name = uploaded.name.rsplit(".", 1)[0]

        csv_bytes = dataframe_to_csv_bytes(df)
        excel_bytes = dataframe_to_excel_bytes(df)

        col1, col2 = st.columns(2)

        with col1:
            st.download_button(
                label="CSVを保存",
                data=csv_bytes,
                file_name=f"{base_name}_cleaned_counts.csv",
                mime="text/csv",
            )

        with col2:
            st.download_button(
                label="Excelを保存",
                data=excel_bytes,
                file_name=f"{base_name}_cleaned_counts.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    except Exception as e:
        st.error(f"処理中にエラーが出ました: {e}")
