"""STEP 4：Mandatoryの判別図をPNGとして作成する。"""

from __future__ import annotations

import colorsys
import io
import re
from pathlib import Path

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.lines import Line2D

from group_judgment import (
    EXPECTED_DIAGRAM_VARIABLES,
    INPUT_SHEET_NAME,
    judge_one_analysis,
    prepare_parameters,
    prepare_quantitative_data,
)


MANDATORY_DIAGRAMS = (1, 2, 3)
FIGURE_SIZE = (11, 29)
VALID_POINT_SIZE = 25
INVALID_POINT_SIZE = 20
ELLIPSE_LINE_WIDTH = 1.7
OUTPUT_DPI = 300

AXIS_LABELS = {
    "Rb*": "Rb* = 100 × Rb / (Rb + Sr + Y + Zr)",
    "Sr*": "Sr* = 100 × Sr / (Rb + Sr + Y + Zr)",
    "Y*": "Y* = 100 × Y / (Rb + Sr + Y + Zr)",
    "Zr*": "Zr* = 100 × Zr / (Rb + Sr + Y + Zr)",
    "ln(100*K/Ca)": "ln(100 × K / Ca)",
}


def make_output_image_name(input_name: str) -> str:
    cleaned_stem = re.sub(
        r"\s*\(\d+\)$", "", Path(input_name).stem
    ).strip()
    date_match = re.match(r"^(\d{8})", cleaned_stem)
    if date_match is None:
        raise ValueError(
            "定量値ファイル名の先頭から8桁の日付を読み取れません。\n"
            "例：20260728_定量値.xlsx"
        )
    return f"{date_match.group(1)}_判別図.png"


def read_inputs(
    parameter_content: bytes,
    quantitative_content: bytes,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    parameter_excel = pd.ExcelFile(io.BytesIO(parameter_content))
    raw_parameter_df = pd.read_excel(
        io.BytesIO(parameter_content),
        sheet_name=parameter_excel.sheet_names[0],
    )
    parameter_df = prepare_parameters(raw_parameter_df)

    quantitative_excel = pd.ExcelFile(io.BytesIO(quantitative_content))
    if INPUT_SHEET_NAME not in quantitative_excel.sheet_names:
        raise ValueError(
            f"定量値ファイルに「{INPUT_SHEET_NAME}」シートがありません。\n"
            f"実際のシート：{quantitative_excel.sheet_names}"
        )
    raw_quantitative_df = pd.read_excel(
        io.BytesIO(quantitative_content),
        sheet_name=INPUT_SHEET_NAME,
    )
    quantitative_df, _ = prepare_quantitative_data(raw_quantitative_df)
    if quantitative_df.empty:
        raise ValueError("定量値シートに表示対象となるデータがありません。")
    return parameter_df, quantitative_df


def add_judgments(
    parameter_df: pd.DataFrame,
    quantitative_df: pd.DataFrame,
) -> pd.DataFrame:
    grouped_parameters = list(
        parameter_df.groupby(
            "Group", sort=False, observed=True
        )
    )
    judgments = [
        judge_one_analysis(row, grouped_parameters)
        for _, row in quantitative_df.iterrows()
    ]
    result = quantitative_df.copy()
    result["判定結果"] = [item["判定結果"] for item in judgments]
    result["得点"] = [item["得点"] for item in judgments]
    result["有効/無効"] = [item["有効/無効"] for item in judgments]
    return result


def ellipse_coordinates(parameter: pd.Series, points: int = 361):
    angle = np.linspace(0, 2 * np.pi, points)
    theta = np.deg2rad(float(parameter["theta"]))
    local_x = float(parameter["a"]) * np.cos(angle)
    local_y = float(parameter["b"]) * np.sin(angle)
    x = (
        float(parameter["x0"])
        + local_x * np.cos(theta)
        - local_y * np.sin(theta)
    )
    y = (
        float(parameter["y0"])
        + local_x * np.sin(theta)
        + local_y * np.cos(theta)
    )
    return x, y


def make_group_colors(group_names: list[str]) -> dict[str, tuple]:
    count = max(len(group_names), 1)
    colors = []
    for index in range(count):
        hue = (index / count + 0.03) % 1.0
        red, green, blue = colorsys.hls_to_rgb(hue, 0.36, 0.78)
        colors.append((red, green, blue, 1.0))
    return {
        group: colors[index]
        for index, group in enumerate(group_names)
    }


def add_group_label(ax, parameter: pd.Series, color, index: int) -> None:
    offsets = [
        (0, 13), (13, 8), (13, -8), (0, -13),
        (-13, -8), (-13, 8), (20, 0), (-20, 0),
    ]
    ax.annotate(
        str(parameter["Group"]),
        xy=(float(parameter["x0"]), float(parameter["y0"])),
        xytext=offsets[index % len(offsets)],
        textcoords="offset points",
        ha="center",
        va="center",
        fontsize=7,
        fontweight="bold",
        color=color,
        bbox={
            "boxstyle": "round,pad=0.15",
            "facecolor": "white",
            "edgecolor": "none",
            "alpha": 0.86,
        },
        arrowprops={
            "arrowstyle": "-",
            "color": color,
            "linewidth": 0.55,
            "alpha": 0.8,
        },
        zorder=6,
    )


def draw_diagram(
    ax,
    diagram: int,
    parameter_df: pd.DataFrame,
    data_df: pd.DataFrame,
    group_colors: dict,
) -> None:
    x_name, y_name = EXPECTED_DIAGRAM_VARIABLES[diagram]
    diagram_parameters = parameter_df[
        (parameter_df["Diagram"] == diagram)
        & (parameter_df["Required"] == 1)
    ]

    for index, (_, parameter) in enumerate(diagram_parameters.iterrows()):
        group = parameter["Group"]
        ellipse_x, ellipse_y = ellipse_coordinates(parameter)
        ax.plot(
            ellipse_x,
            ellipse_y,
            color=group_colors[group],
            linewidth=ELLIPSE_LINE_WIDTH,
            alpha=0.9,
            zorder=2,
        )
        add_group_label(ax, parameter, group_colors[group], index)

    valid_df = data_df[data_df["有効/無効"] == "有効"]
    invalid_df = data_df[data_df["有効/無効"] == "無効"]

    ax.scatter(
        valid_df[x_name],
        valid_df[y_name],
        s=VALID_POINT_SIZE,
        marker="o",
        facecolor="#111111",
        edgecolor="white",
        linewidth=0.45,
        alpha=0.78,
        zorder=4,
    )
    ax.scatter(
        invalid_df[x_name],
        invalid_df[y_name],
        s=INVALID_POINT_SIZE,
        marker="x",
        color="#666666",
        linewidth=0.9,
        alpha=0.65,
        zorder=3,
    )

    ax.set_xlabel(AXIS_LABELS[x_name], fontsize=11)
    ax.set_ylabel(AXIS_LABELS[y_name], fontsize=11)
    ax.grid(True, linestyle=":", linewidth=0.7, alpha=0.5)
    ax.set_axisbelow(True)
    ax.margins(x=0.06, y=0.08)
    ax.set_box_aspect(1)
    ax.text(
        0.012,
        0.985,
        {1: "(a)", 2: "(b)", 3: "(c)"}[diagram],
        transform=ax.transAxes,
        ha="left",
        va="top",
        fontsize=12,
        fontweight="bold",
    )


def create_figure(
    parameter_df: pd.DataFrame,
    data_df: pd.DataFrame,
) -> plt.Figure:
    group_names = list(dict.fromkeys(parameter_df["Group"].tolist()))
    group_colors = make_group_colors(group_names)
    fig, axes = plt.subplots(3, 1, figsize=FIGURE_SIZE)
    for ax, diagram in zip(axes, MANDATORY_DIAGRAMS):
        draw_diagram(
            ax,
            diagram,
            parameter_df,
            data_df,
            group_colors,
        )

    valid_count = int((data_df["有効/無効"] == "有効").sum())
    invalid_count = int((data_df["有効/無効"] == "無効").sum())
    axes[0].legend(
        handles=[
            Line2D(
                [0], [0], marker="o", linestyle="None", markersize=7,
                markerfacecolor="#111111", markeredgecolor="white",
                label=f"Valid ({valid_count})",
            ),
            Line2D(
                [0], [0], marker="x", linestyle="None", markersize=7,
                markeredgewidth=0.9, color="#666666",
                label=f"Invalid ({invalid_count})",
            ),
        ],
        loc="upper right",
        frameon=True,
        facecolor="white",
        edgecolor="#B8B8B8",
        framealpha=0.92,
        fontsize=9,
    )
    fig.subplots_adjust(
        top=0.985,
        bottom=0.045,
        left=0.11,
        right=0.98,
        hspace=0.30,
    )
    return fig


def create_discrimination_diagram(
    parameter_content: bytes,
    quantitative_content: bytes,
    quantitative_file_name: str,
) -> dict:
    parameter_df, quantitative_df = read_inputs(
        parameter_content,
        quantitative_content,
    )
    data_df = add_judgments(parameter_df, quantitative_df)
    figure = create_figure(parameter_df, data_df)
    output = io.BytesIO()
    figure.savefig(
        output,
        format="png",
        dpi=OUTPUT_DPI,
        bbox_inches="tight",
        facecolor="white",
    )
    plt.close(figure)
    output.seek(0)

    return {
        "content": output.getvalue(),
        "file_name": make_output_image_name(quantitative_file_name),
        "analysis_count": len(data_df),
        "valid_count": int((data_df["有効/無効"] == "有効").sum()),
        "invalid_count": int((data_df["有効/無効"] == "無効").sum()),
        "group_count": int(parameter_df["Group"].nunique()),
    }
