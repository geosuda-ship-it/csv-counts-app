"""Streamlit entry point: streamlit run app.py."""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory

import numpy as np
import plotly.graph_objects as go
from plotly.colors import qualitative
import streamlit as st

from analysis_core import DIAGRAMS, run_analysis, export_results_xlsx


def analyze_upload(content):
    # Every session gets its own temporary files; remove them after exporting.
    with TemporaryDirectory(prefix="obsidian_") as folder:
        source = Path(folder) / "input.xlsx"
        source.write_bytes(content)
        results = run_analysis(source)
        target = export_results_xlsx(results, Path(folder) / "results.xlsx")
        return results, target.read_bytes()


def ellipse_xy(row):
    angle = np.linspace(0, 2 * np.pi, 181)
    theta = np.deg2rad(float(row["theta"]))
    u, v = row["a"] * np.cos(angle), row["b"] * np.sin(angle)
    return (row["x0"] + u * np.cos(theta) - v * np.sin(theta),
            row["y0"] + u * np.sin(theta) + v * np.cos(theta))


def make_diagram(results, diagram, selected_groups):
    x_name, y_name = DIAGRAMS[diagram]
    fig = go.Figure()
    for i, cluster in enumerate(results.cluster_counts["Cluster"]):
        color = "#999999" if cluster == "未所属" else qualitative.Plotly[i % 10]
        part = results.median_with_stars.query("Cluster == @cluster")
        fig.add_trace(go.Scatter(
            x=part[x_name], y=part[y_name], mode="markers", name=cluster,
            text=part["Sample"], marker=dict(color=color, size=8),
            hovertemplate="%{text}<br>" + x_name + ": %{x:.3f}<br>" + y_name + ": %{y:.3f}<extra>%{fullData.name}</extra>",
        ))
        ellipse = results.ellipse_parameters
        ellipse = ellipse[(ellipse["Cluster"] == cluster) & (ellipse["Diagram"] == diagram)]
        for _, row in ellipse.iterrows():
            x, y = ellipse_xy(row)
            fig.add_trace(go.Scatter(x=x, y=y, mode="lines", name=cluster,
                                     line=dict(color=color), showlegend=False))
    params = results.parameter_table
    for _, row in params[(params["Diagram"] == diagram) & params["Group"].isin(selected_groups)].iterrows():
        x, y = ellipse_xy(row)
        fig.add_trace(go.Scatter(x=x, y=y, mode="lines", name=row["Group"],
                                 line=dict(color="#555555", dash="dash")))
    fig.update_layout(title=f"D{diagram}: {x_name} vs {y_name}",
                      xaxis_title=x_name, yaxis_title=y_name, height=520,
                      template="plotly_white", margin=dict(l=30, r=20, t=60, b=30))
    fig.update_yaxes(scaleanchor="x", scaleratio=1)
    return fig


def show_table(frame):
    display = frame.copy()
    for column in display.columns:
        if "内包率" in column:
            display[column] = display[column].map(lambda v: "" if np.isnan(v) else f"{v:.1%}")
    st.dataframe(display, hide_index=True, width="stretch")


def main():
    st.set_page_config(page_title="黒曜石EDXRF クラスター解析", page_icon="🔬", layout="wide")
    st.title("黒曜石EDXRF産地判別・クラスター解析プログラム")
    st.write("判定済みExcelをアップロードして、資料のクラスターと化学組成グループとの整合性を確認します。")
    st.caption("Excelで個別判定・再計算を実行し、保存した .xlsx ファイルを使用してください。")
    uploaded = st.file_uploader("Excelファイル", type=["xlsx"])
    with st.expander("解析方法・入力ファイルについて"):
        st.write("Medianシートの資料ごとの代表値と、AS-1以降の各シートのL4:T7にあるD1〜D3の判別楕円パラメータを読み込みます。")
        st.write("Rb*・Sr*・Zr*は、Rb+Sr+Y+Zrを分母として計算します。3変数を標準化し、HDBSCAN（最小クラスター資料数3、min_samples=3）で検出します。")
        st.write("クラスターの楕円は共分散に基づく95%分布楕円です。中心の信頼区間ではありません。")
        st.write("内包率はクラスター内の資料が判別楕円に入る割合です。3図共通内包率は、同じ資料がD1〜D3すべての楕円に入る割合です。")
        st.write("候補はクラスター全体に対する比較結果です。各資料の個別判定や産地の確定を意味しません。代表値の選択方法は元のMedianシートに依存します。")
    if uploaded is None:
        st.session_state.pop("analysis", None)
        st.info("Excelを選択すると、解析を開始できます。")
        return
    content = uploaded.getvalue()
    signature = sha256(content).hexdigest()
    saved = st.session_state.get("analysis")
    if saved and saved["signature"] != signature:
        st.session_state.pop("analysis", None)
    if st.button("解析を実行", type="primary"):
        st.session_state.pop("analysis", None)
        try:
            with st.spinner("クラスターを検出し、化学組成グループと比較しています…"):
                results, output = analyze_upload(content)
            st.session_state["analysis"] = dict(signature=signature, results=results, output=output)
        except Exception as exc:
            st.error(f"解析を完了できませんでした：{exc}")
    saved = st.session_state.get("analysis")
    if not saved:
        return
    results = saved["results"]
    st.success("解析が完了しました。結果Excelをダウンロードできます。")
    st.download_button("結果Excelをダウンロード", data=saved["output"],
                       file_name=f"{Path(uploaded.name).stem}_cluster_analysis_results.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    overview, samples, plots, parameters, comparisons, top3 = st.tabs([
        "検出結果の概要", "資料ごとの結果", "D1〜D3の散布図", "クラスター楕円パラメータ", "全化学組成グループとの比較", "クラスターごとの上位3件"])
    with overview:
        st.dataframe(results.summary_metrics.astype(str), hide_index=True)
        st.dataframe(results.cluster_counts, hide_index=True)
        if not len(results.ellipse_parameters):
            st.info("今回はクラスターが検出されませんでした。全資料が未所属です。")
    with samples:
        show_table(results.sample_results)
        st.caption("化学組成グループと内包率は、所属クラスターの上位候補です。未所属の資料には候補を付与しません。")
    with plots:
        groups = list(results.parameter_table["Group"].unique())
        defaults = [g for g in results.cluster_top3["化学組成グループ"].unique() if g in groups]
        selected = st.multiselect("重ねて表示する化学組成グループの判別楕円（破線）", groups, default=defaults)
        st.caption("点：資料、実線：クラスターの95%分布楕円、破線：選択した化学組成グループの判別楕円")
        for diagram in DIAGRAMS:
            st.plotly_chart(make_diagram(results, diagram, selected), width="stretch")
    with parameters:
        st.dataframe(results.ellipse_parameters, hide_index=True, width="stretch")
        st.caption("a、bは半軸長、thetaはX軸から反時計回りの角度（度）です。")
    with comparisons:
        show_table(results.all_group_comparisons)
    with top3:
        show_table(results.cluster_top3)
        st.caption("3図共通内包率が0%を超えるグループから最大3件。共通内包率が同じ場合は各図の内包率合計、さらに同じ場合は元のシート順で選びます。")


if __name__ == "__main__":
    main()
