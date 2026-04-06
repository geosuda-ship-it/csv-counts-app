import csv
from io import StringIO, BytesIO
import pandas as pd
import streamlit as st

st.set_page_config(page_title="NEX DE app", layout="wide")
st.title("NEX DE 強度→定量値変換アプリ")

# =========================
# QC-2基準
# =========================
reference_qc2 = {
    "Mid-Z_K-Kα": 4.27826,
    "Mid-Z_Ca-Kβ1": 0.15422,
    "Mid-Z_Mn-Kα": 0.90407,
    "Mid-Z_Fe-Kβ1": 3.15459,
    "High-Z_Zn-Kα": 0.01778,
    "High-Z_Rb-Kα": 0.22959,
    "High-Z_Sr-Kα": 0.07396,
    "High-Z_Y-Kα": 0.09792,
    "High-Z_Zr-Kα": 0.17312,
    "High-Z_Nb-Kα": 0.10719,
    "High-Z_Ag-Kα": 1.53429,
}

target_cols = list(reference_qc2.keys())

# =========================
# 検量線
# =========================
B = {"K":9312.59,"Ca":30145.8,"Mn":746.088,"Fe":4061.51,"Zn":7048.49,
     "Rb":1360.26,"Sr":1112.8,"Y":898.725,"Zr":810.974,"Nb":737.066}

C = {"K":-1378.82,"Ca":-204.571,"Mn":-315.439,"Fe":-5513.26,"Zn":-44.4817,
     "Rb":-16.0436,"Sr":-11.9096,"Y":-12.6423,"Zr":-18.6844,"Nb":-31.1933}

# =========================
# 重なり補正
# =========================
BG14 = -0.113036
BG15 = -0.121299
BG16 = -0.161034

# =========================
# CSV読み込み
# =========================
def read_csv(uploaded):
    text = uploaded.getvalue().decode("cp932", errors="replace")
    return list(csv.reader(StringIO(text)))

# =========================
# cps抽出
# =========================
def extract(rows):
    idxs = [i for i,r in enumerate(rows) if "cps/μa" in [str(x).lower() for x in r]]
    data = []
    cols = None

    for idx in idxs:
        h1, h2 = rows[idx-2], rows[idx-1]

        if cols is None:
            cols = ["Sample","Method","Date"] + [
                f"{a}_{b}" if a and b else a or b
                for a,b in zip(h1[3:],h2[3:])
            ]

        for r in rows[idx+1:]:
            if not any(r): break
            r = r + [""]*(len(cols)-len(r))
            data.append(r[:len(cols)])

    df = pd.DataFrame(data, columns=cols)

    for c in df.columns[3:]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    return df

# =========================
# ドリフト補正
# =========================
def drift(df):
    qc = df[df["Sample"].str.replace("-","").str.upper()=="QC2"].iloc[0]
    f = {}
    for c in target_cols:
        f[c] = reference_qc2[c]/qc[c] if qc[c]!=0 else None
    return f

# =========================
# 定量
# =========================
def quantify(df):

    f = drift(df)
    d = df.copy()

    for c in target_cols:
        if f[c]:
            d[c] *= f[c]

    ag = d["High-Z_Ag-Kα"]

    # 軽元素
    d["K"]  = B["K"]  * d["Mid-Z_K-Kα"] + C["K"]
    d["Ca"] = B["Ca"] * d["Mid-Z_Ca-Kβ1"] + C["Ca"]
    d["Mn"] = B["Mn"] * d["Mid-Z_Mn-Kα"] + C["Mn"]
    d["Fe"] = B["Fe"] * d["Mid-Z_Fe-Kβ1"] + C["Fe"]

    # ZnもAg補正
    d["Zn"] = B["Zn"]*(d["High-Z_Zn-Kα"]/ag)+C["Zn"]

    d["Rb"] = B["Rb"]*(d["High-Z_Rb-Kα"]/ag)+C["Rb"]
    d["Sr"] = B["Sr"]*(d["High-Z_Sr-Kα"]/ag)+C["Sr"]

    d["Y"]  = B["Y"]*(d["High-Z_Y-Kα"]/ag)+C["Y"] + d["Rb"]*BG14
    d["Zr"] = B["Zr"]*(d["High-Z_Zr-Kα"]/ag)+C["Zr"] + d["Sr"]*BG15
    d["Nb"] = B["Nb"]*(d["High-Z_Nb-Kα"]/ag)+C["Nb"] + d["Y"]*BG16

    cols = ["Sample","Method","Date","K","Ca","Mn","Fe","Zn","Rb","Sr","Y","Zr","Nb"]
    r = d[cols].copy()

    r = r.rename(columns={c:f"{c} ppm" for c in cols if c not in ["Sample","Method","Date"]})

    # ===== 並び（古い→新しい） =====
    r = r.sort_values("Date", ascending=True).reset_index(drop=True)

    # ===== 小数点2桁 =====
    ppm_cols = [c for c in r.columns if "ppm" in c]
    r[ppm_cols] = r[ppm_cols].round(2)

    return r

# =========================
# Excel出力
# =========================
def to_excel(df):
    df = df.copy()
    df["Date"] = df["Date"].dt.strftime("%Y-%m-%d %H:%M:%S")

    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    return bio.getvalue()

# =========================
# UI
# =========================
file = st.file_uploader("CSVアップロード", type="csv")

if file:
    rows = read_csv(file)
    df = extract(rows)
    result = quantify(df)

    st.subheader("定量結果")
    st.dataframe(result, use_container_width=True)

    st.download_button(
        "Excelダウンロード",
        to_excel(result),
        file_name="result.xlsx"
    )
