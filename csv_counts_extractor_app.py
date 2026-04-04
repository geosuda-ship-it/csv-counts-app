import streamlit as st

st.title("test app")

uploaded = st.file_uploader("CSVアップロード", type=["csv"])

if uploaded is not None:
    st.write("アップロード成功")
