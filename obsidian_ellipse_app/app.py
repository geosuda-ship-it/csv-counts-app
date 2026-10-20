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
if uploaded is not None:
    import openpyxl

    parameter_file = st.file_uploader(
        "② 判別楕円パラメーターのExcelファイルをアップロードしてください。",
        type=["xlsx"],
    )

    if parameter_file is not None:
        parameter_columns = [
            "Diagram", "Group", "X", "Y",
            "x0", "y0", "a", "b", "theta",
        ]

        parameter_rows = []
        workbook = None

        try:
            workbook = openpyxl.load_workbook(
                BytesIO(parameter_file.getvalue()),
                read_only=True,
                data_only=True,
            )

            if "AS-1" not in workbook.sheetnames:
                raise ValueError("AS-1シートがありません。")

            start = workbook.sheetnames.index("AS-1")

            for sheet_name in workbook.sheetnames[start:]:
                sheet = workbook[sheet_name]

                block = list(sheet.iter_rows(
                    min_row=4,
                    max_row=10,
                    min_col=12,
                    max_col=20,
                    values_only=True,
                ))

                headers = [
                    str(value).strip() if value is not None else ""
                    for value in block[0]
                ]

                if headers != parameter_columns:
                    raise ValueError(
                        f"{sheet_name}シートのL4:T4の見出しが一致しません。"
                    )

                group_data = pd.DataFrame(
                    block[1:],
                    columns=parameter_columns,
                )

                if sorted(group_data["Diagram"].tolist()) != [1, 2, 3, 4, 5, 6]:
                    raise ValueError(
                        f"{sheet_name}シートにD1〜D6が揃っていません。"
                    )

                if not group_data["Group"].eq(sheet_name).all():
                    raise ValueError(
                        f"{sheet_name}シートのGroupとシート名が一致しません。"
                    )

                parameter_rows.extend(
                    group_data.to_dict("records")
                )

        except Exception as exc:
            st.error(f"パラメーターを読み込めませんでした：{exc}")
            st.stop()

        finally:
            if workbook is not None:
                workbook.close()

        ellipse_parameters = pd.DataFrame(
            parameter_rows,
            columns=parameter_columns,
        )

        st.success(
            f"{ellipse_parameters['Group'].nunique()}化学組成グループ・"
            f"{len(ellipse_parameters)}楕円を読み込みました。"
        )

        with st.expander("判別楕円パラメーターを確認する"):
            st.dataframe(
                ellipse_parameters,
                hide_index=True,
                width="stretch",
            )
if uploaded is not None and parameter_file is not None:
    import numpy as np

    calculation_data = data.copy()

    elements = [
        "K", "Ca", "Mn", "Fe", "Zn",
        "Rb", "Sr", "Y", "Zr", "Nb",
    ]

    try:
        # 元素値を数値として確認する
        for element in elements:
            column = f"{element}_ppm"

            if column not in calculation_data.columns:
                raise ValueError(f"{column}列がありません。")

            calculation_data[column] = pd.to_numeric(
                calculation_data[column],
                errors="raise",
            )

        element_columns = [f"{e}_ppm" for e in elements]

        if not np.isfinite(
            calculation_data[element_columns].to_numpy(dtype=float)
        ).all():
            raise ValueError("元素値に空白や計算不能な値があります。")

        denominator = calculation_data[
            ["Rb_ppm", "Sr_ppm", "Y_ppm", "Zr_ppm"]
        ].sum(axis=1)

        if (denominator <= 0).any():
            raise ValueError("Rb＋Sr＋Y＋Zrが0以下の測定があります。")

        if calculation_data[
            ["K_ppm", "Ca_ppm", "Fe_ppm"]
        ].eq(0).any().any():
            raise ValueError("K・Ca・Feが0のため、比を計算できません。")

        with np.errstate(over="ignore", invalid="ignore"):
            for element in ["Rb", "Sr", "Zr"]:
                calculation_data[f"{element}*"] = (
                    100 * calculation_data[f"{element}_ppm"]
                    / denominator
                )

            calculation_data["100Fe/K"] = (
                100 * calculation_data["Fe_ppm"]
                / calculation_data["K_ppm"]
            )

            calculation_data["K/Ca"] = (
                calculation_data["K_ppm"]
                / calculation_data["Ca_ppm"]
            )

            calculation_data["100Mn/Fe"] = (
                100 * calculation_data["Mn_ppm"]
                / calculation_data["Fe_ppm"]
            )

        indicator_columns = [
            "Rb*", "Sr*", "Zr*",
            "100Fe/K", "K/Ca", "100Mn/Fe",
        ]

        if not np.isfinite(
            calculation_data[indicator_columns].to_numpy(dtype=float)
        ).all():
            raise ValueError("判別指標に計算不能な値があります。")

    except Exception as exc:
        st.error(f"指標を計算できませんでした：{exc}")
        st.stop()

    st.success("判別に使う6種類の指標を計算しました。")

    with st.expander("判別指標を確認する"):
        st.dataframe(
            calculation_data[["sample"] + indicator_columns],
            hide_index=True,
            width="stretch",
        )
if uploaded is not None and parameter_file is not None:
    import numpy as np

    diagram_axes = {
        1: ("Rb*", "Sr*"),
        2: ("Rb*", "Zr*"),
        3: ("Sr*", "Zr*"),
        4: ("Sr*", "100Fe/K"),
        5: ("Sr*", "K/Ca"),
        6: ("Rb*", "100Mn/Fe"),
    }

    measurement_data = calculation_data.reset_index(drop=True).copy()

    try:
        parameters = ellipse_parameters.copy()
        numeric_columns = ["Diagram", "x0", "y0", "a", "b", "theta"]
        parameters[numeric_columns] = parameters[numeric_columns].apply(
            pd.to_numeric, errors="raise"
        )

        if not np.isfinite(parameters[numeric_columns].to_numpy(dtype=float)).all():
            raise ValueError("楕円パラメーターに空白や計算不能な値があります。")
        if parameters[["a", "b"]].le(0).any().any():
            raise ValueError("楕円の半軸長a・bは正の値にしてください。")

        detail_frames = []

        for group, group_parameters in parameters.groupby("Group", sort=False):
            p = group_parameters.set_index("Diagram")
            distances = []

            for diagram, (x_name, y_name) in diagram_axes.items():
                row = p.loc[diagram]
                if (row["X"], row["Y"]) != (x_name, y_name):
                    raise ValueError(f"{group}のD{diagram}のX・Yが一致しません。")

                theta = np.deg2rad(row["theta"])
                dx = measurement_data[x_name].to_numpy() - row["x0"]
                dy = measurement_data[y_name].to_numpy() - row["y0"]
                with np.errstate(over="ignore", invalid="ignore"):
                    distance = np.hypot(
                        (dx * np.cos(theta) + dy * np.sin(theta)) / row["a"],
                        (-dx * np.sin(theta) + dy * np.cos(theta)) / row["b"],
                    )
                distances.append(distance)

            distances = np.column_stack(distances)
            inside = distances <= 1
            valid = (
                np.isfinite(distances).all(axis=1)
                & inside[:, :3].all(axis=1)
                & (distances[:, 3:] <= [3, 3, 5]).all(axis=1)
            )

            frame = pd.DataFrame({
                "測定No": np.arange(1, len(measurement_data) + 1),
                "sample": measurement_data["sample"].to_numpy(),
                "化学組成グループ": group,
            })
            for diagram in range(1, 7):
                distance = distances[:, diagram - 1]
                frame[f"D{diagram}規格化距離"] = np.where(
                    np.isfinite(distance), distance, np.nan
                )
                frame[f"D{diagram}判定"] = np.where(
                    ~np.isfinite(distance), "計算不能",
                    np.where(inside[:, diagram - 1], "楕円内", "楕円外"),
                )
            frame["有効性"] = np.where(valid, "有効", "無効")
            frame["得点"] = np.where(
                valid, 7 + inside[:, 3:].sum(axis=1), np.nan
            )
            detail_frames.append(frame)

        if not detail_frames:
            raise ValueError("判別に使用できるパラメーターがありません。")

        detail_results = pd.concat(detail_frames, ignore_index=True)
        detail_results = detail_results.sort_values(
            "測定No", kind="stable"
        ).reset_index(drop=True)

        measurement_results = measurement_data[
            ["sample", "測定モード", "試料測定日時"]
        ].copy()
        measurement_results.insert(
            0, "測定No", np.arange(1, len(measurement_data) + 1)
        )
        measurement_results["判別結果"] = "判定不能"
        measurement_results["得点"] = np.nan
        measurement_results["有効性"] = "判定不能"
        measurement_results["_Groups"] = [()] * len(measurement_results)

        valid_results = detail_results[detail_results["有効性"].eq("有効")]
        for number, part in valid_results.groupby("測定No", sort=False):
            score = int(part["得点"].max())
            groups = tuple(
                part.loc[part["得点"].eq(score), "化学組成グループ"]
            )
            index = number - 1
            measurement_results.at[index, "判別結果"] = "・".join(groups)
            measurement_results.at[index, "得点"] = score
            measurement_results.at[index, "有効性"] = "有効"
            measurement_results.at[index, "_Groups"] = groups

    except Exception as exc:
        st.error(f"判別できませんでした：{exc}")
        st.stop()

    st.success("測定ごとの判別が完了しました。")
    st.dataframe(
        measurement_results.drop(columns="_Groups"),
        hide_index=True,
        width="stretch",
    )

    with st.expander("化学組成グループ別の判別詳細を確認する"):
        st.dataframe(detail_results, hide_index=True, width="stretch")
if uploaded is not None and parameter_file is not None:
    from collections import Counter

    group_order = list(ellipse_parameters["Group"].unique())
    summary_rows = []

    for sample, part in measurement_results.groupby("sample", sort=False):
        valid = part[part["有効性"].eq("有効")]
        final_groups = ()
        final_score = np.nan

        if not valid.empty:
            final_score = int(valid["得点"].max())
            highest = valid[valid["得点"].eq(final_score)]

            # 最高得点の測定結果に含まれるグループの出現回数
            highest_counts = Counter(
                group
                for groups in highest["_Groups"]
                for group in groups
            )
            maximum = max(highest_counts.values())
            candidates = {
                group for group, count in highest_counts.items()
                if count == maximum
            }

            # 同数の場合は、残った候補だけを全有効測定で再比較
            if len(candidates) > 1:
                all_counts = Counter(
                    group
                    for groups in valid["_Groups"]
                    for group in groups
                )
                maximum = max(all_counts[group] for group in candidates)
                candidates = {
                    group for group in candidates
                    if all_counts[group] == maximum
                }

            # 複合判定の表記順は、パラメーターのシート順
            final_groups = tuple(
                group for group in group_order if group in candidates
            )

        summary_rows.append({
            "Sample": sample,
            "最終判別結果": "・".join(final_groups) if final_groups else "判定不能",
            "得点": final_score,
            "測定データ数": len(part),
            "有効測定数": len(valid),
            "_Groups": final_groups,
        })

    sample_summary = pd.DataFrame(summary_rows)

    st.success("資料ごとの最終判別結果を集計しました。")
    st.subheader("資料ごとの最終判別結果")
    st.dataframe(
        sample_summary.drop(columns="_Groups"),
        hide_index=True,
        width="stretch",
    )
if uploaded is not None and parameter_file is not None:
    median_elements = [
        "K", "Ca", "Mn", "Fe", "Zn",
        "Rb", "Sr", "Y", "Zr", "Nb",
    ]

    try:
        median_rows = []
        median_counts = []
        selected = pd.Series(False, index=measurement_results.index)

        for _, summary in sample_summary.iterrows():
            # 資料IDが一致する測定を対象にする
            mask = measurement_results["sample"].eq(summary["Sample"])

            if summary["最終判別結果"] != "判定不能":
                final_groups = set(summary["_Groups"])

                # 有効・最終得点と一致・最終グループのいずれかを含む
                mask &= measurement_results["有効性"].eq("有効")
                mask &= measurement_results["得点"].eq(summary["得点"])
                mask &= measurement_results["_Groups"].map(
                    lambda groups: bool(final_groups.intersection(groups))
                )

            # 判定不能の場合は、資料IDが一致する全測定を使用する
            chosen = measurement_data.loc[mask]
            if chosen.empty:
                raise ValueError(
                    f"{summary['Sample']}の中央値に使用できる測定がありません。"
                )

            selected |= mask
            median_counts.append(len(chosen))

            row = {"Sample": summary["Sample"]}
            for element in median_elements:
                row[element] = chosen[f"{element}_ppm"].median()

            row["主要判定結果"] = summary["最終判別結果"]
            row["得点"] = summary["得点"]
            row["中央値使用数"] = len(chosen)
            median_rows.append(row)

        median_results = pd.DataFrame(median_rows)
        sample_summary["中央値使用数"] = median_counts
        measurement_results["中央値使用"] = np.where(
            selected, "使用", "除外"
        )

    except Exception as exc:
        st.error(f"中央値を計算できませんでした：{exc}")
        st.stop()

    st.success("資料ごとの中央値を計算しました。")
    st.subheader("資料ごとの中央値（ppm）")
    st.dataframe(median_results, hide_index=True, width="stretch")

    with st.expander("中央値への使用・除外を確認する"):
        st.dataframe(
            measurement_results.drop(columns="_Groups"),
            hide_index=True,
            width="stretch",
        )
