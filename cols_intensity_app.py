from pathlib import Path

import pandas as pd
import streamlit as st

from cols_edx_full_pipeline import (
    QC_MODE,
    calculate_all,
    convert_csv,
    extract_measurements,
    is_qc,
)


st.set_page_config(
    page_title="Rigaku NEX DE 強度・定量変換ツール（COLS版）",
    page_icon="🧪",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container {
        max-width: 1180px;
        padding-top: 3.2rem;
        padding-bottom: 3rem;
    }
    .app-title {
        color: #1f2937;
        font-size: clamp(1.35rem, 2.4vw, 1.8rem);
        font-weight: 700;
        line-height: 1.5;
        padding: 0.25rem 0 0.35rem 0;
        margin-bottom: 0.3rem;
        white-space: normal;
        overflow-wrap: anywhere;
        overflow: visible;
    }
    .app-subtitle {
        color: #475569;
        font-size: 1.05rem;
        margin-bottom: 1.5rem;
    }
    .step-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 0.75rem;
        padding: 1rem 1.2rem;
        margin: 0.5rem 0 1.2rem 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="app-title">
    Rigaku NEX DE用 強度から定量値への変換ツール
    （明治大学黒耀石研究センター版）
    </div>
    <div class="app-subtitle">
    NEX DEのCSVから、ドリフト補正・Ag内標準補正・定量・重なり補正を
    一括実行します。
    </div>
    """,
    unsafe_allow_html=True,
)

with st.expander("計算内容", expanded=False):
    st.markdown(
        """
        - `QC-2`または`QC2`、測定モード`obart_prec3_air`をQCとして認識
        - QCと試料を測定日ごとに分離し、日をまたいで補正しない
        - 同日のQC測定回数分だけ、各試料の補正結果と定量値を算出
        - K、Ca、Mn、Fe、Zn：ドリフト補正後強度を使用
        - Rb、Sr、Y、Zr、Nb：ドリフト補正後Ag強度で内標準化
        - YはRb、ZrはSr、NbはYによる重なり補正を適用
        - 基準強度・検量線・重なり補正にはCOLS版の定数を使用
        """
    )

st.markdown(
    """
    <div class="step-card">
    <strong>1. CSVを選択</strong><br>
    NEX DEから出力したCSVを選択してください。QC測定を含む必要があります。
    </div>
    """,
    unsafe_allow_html=True,
)

uploaded = st.file_uploader(
    "NEX DEのCSVファイル",
    type=["csv"],
    help="1回の処理につき1つのCSVを選択します。",
)


def clear_previous_result():
    st.session_state.pop("cols_result_bytes", None)
    st.session_state.pop("cols_result_name", None)


if uploaded is None:
    clear_previous_result()
    st.info("CSVファイルを選択すると、測定内容を確認できます。")
    st.stop()

file_bytes = uploaded.getvalue()

if st.session_state.get("cols_uploaded_name") != uploaded.name:
    clear_previous_result()
    st.session_state.cols_uploaded_name = uploaded.name

try:
    measurements = extract_measurements(file_bytes)
    qc_measurements = [item for item in measurements if is_qc(item)]
    sample_measurements = [item for item in measurements if not is_qc(item)]

    qc_by_date = {}
    sample_by_date = {}
    for item in qc_measurements:
        measurement_date = item["measured_at"].date()
        qc_by_date[measurement_date] = qc_by_date.get(measurement_date, 0) + 1
    for item in sample_measurements:
        measurement_date = item["measured_at"].date()
        sample_by_date[measurement_date] = (
            sample_by_date.get(measurement_date, 0) + 1
        )

    summary_dates = sorted(set(qc_by_date) | set(sample_by_date))
    summary = pd.DataFrame(
        [
            {
                "測定日": measurement_date.isoformat(),
                "QC回数": qc_by_date.get(measurement_date, 0),
                "試料測定数": sample_by_date.get(measurement_date, 0),
                "出力予定行数": (
                    qc_by_date.get(measurement_date, 0)
                    * sample_by_date.get(measurement_date, 0)
                ),
            }
            for measurement_date in summary_dates
        ]
    )
except Exception as exc:
    clear_previous_result()
    st.error(f"CSVを確認できませんでした：{exc}")
    st.stop()

st.success(f"`{uploaded.name}`を読み込みました。")

metric1, metric2, metric3, metric4 = st.columns(4)
metric1.metric("測定日数", len(summary_dates))
metric2.metric("QC測定数", len(qc_measurements))
metric3.metric("試料測定数", len(sample_measurements))
metric4.metric(
    "出力予定行数",
    int(summary["出力予定行数"].sum()) if not summary.empty else 0,
)
st.dataframe(summary, use_container_width=True, hide_index=True)

missing_dates = [
    measurement_date
    for measurement_date in summary_dates
    if sample_by_date.get(measurement_date, 0) > 0
    and qc_by_date.get(measurement_date, 0) == 0
]
if missing_dates:
    st.warning(
        "次の測定日には同日のQCがないため、試料を計算できません："
        + "、".join(item.isoformat() for item in missing_dates)
    )

if not qc_measurements:
    st.error(
        f"QC-2またはQC2、測定モード`{QC_MODE}`のQC測定がありません。"
    )
    st.stop()

st.markdown(
    """
    <div class="step-card">
    <strong>2. 一括計算</strong><br>
    内容を確認して、下のボタンを押してください。
    </div>
    """,
    unsafe_allow_html=True,
)

if st.button("計算を実行", type="primary", use_container_width=True):
    try:
        with st.spinner("COLS版の定数で計算しています…"):
            _, corrected_rows, skipped_dates = calculate_all(measurements)
            result_bytes = convert_csv(file_bytes)

        output_name = (
            f"{Path(uploaded.name).stem}_COLS_補正後強度_定量値.xlsx"
        )
        st.session_state.cols_result_bytes = result_bytes
        st.session_state.cols_result_name = output_name
        st.session_state.cols_result_count = len(corrected_rows)
        st.session_state.cols_skipped_dates = skipped_dates
    except Exception as exc:
        clear_previous_result()
        st.error(f"計算中にエラーが発生しました：{exc}")

if st.session_state.get("cols_result_bytes") is not None:
    st.success(
        f"計算が完了しました。"
        f"{st.session_state.get('cols_result_count', 0)}行を出力します。"
    )

    skipped_dates = st.session_state.get("cols_skipped_dates", [])
    if skipped_dates:
        st.warning(
            "同日のQCがなく除外した測定日："
            + "、".join(item.isoformat() for item in skipped_dates)
        )

    st.markdown(
        """
        <div class="step-card">
        <strong>3. Excelを保存</strong><br>
        QC補正係数、補正後強度、定量値、計算定数の4シートを含みます。
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.download_button(
        "最終Excelをダウンロード",
        data=st.session_state.cols_result_bytes,
        file_name=st.session_state.cols_result_name,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=True,
    )
