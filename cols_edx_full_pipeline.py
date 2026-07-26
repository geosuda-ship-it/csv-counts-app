"""
明治大学黒耀石研究センター版の計算設定。

共通のCSV解析・日別QC処理・Excel出力は
colab_edx_full_pipeline.pyを利用し、明大専用定数だけを切り替える。
"""

from contextlib import contextmanager

import colab_edx_full_pipeline as core


REFERENCE_INTENSITY = {
    "K": 4.16538,
    "Ca": 0.14892,
    "Mn": 0.90216,
    "Fe": 3.16091,
    "Zn": 0.08522,
    "Rb": 1.16831,
    "Sr": 0.37845,
    "Y": 0.50381,
    "Zr": 0.88033,
    "Nb": 0.53323,
    "Ag": 10.46824,
}

CALIBRATION_B = {
    "K": 9538.77,
    "Ca": 31518.00,
    "Mn": 746.09,
    "Fe": 4041.97,
    "Zn": 10015.90,
    "Rb": 1831.54,
    "Sr": 1507.65,
    "Y": 1215.25,
    "Zr": 1048.10,
    "Nb": 1057.98,
}

CALIBRATION_C = {
    "K": -1177.73,
    "Ca": -267.40,
    "Mn": -318.16,
    "Fe": -5373.50,
    "Zn": -41.0826,
    "Rb": -11.6155,
    "Sr": -9.5885,
    "Y": -9.9429,
    "Zr": -11.8799,
    "Nb": -33.9491,
}

OVERLAP_CORRECTION = {
    "Y": ("Rb", -0.118793),
    "Zr": ("Sr", -0.121227),
    "Nb": ("Y", -0.002482),
}

QC_MODE = "cols_prec1_air"
extract_measurements = core.extract_measurements


def is_qc(measurement: dict) -> bool:
    return (
        core.normalize_qc_name(measurement["sample"]) == "QC2"
        and measurement["mode"].strip().lower() == QC_MODE.lower()
    )


@contextmanager
def cols_configuration():
    original_reference = core.REFERENCE_INTENSITY
    original_b = core.CALIBRATION_B
    original_c = core.CALIBRATION_C
    original_overlap = core.OVERLAP_CORRECTION
    original_qc_mode = core.QC_MODE

    core.REFERENCE_INTENSITY = REFERENCE_INTENSITY
    core.CALIBRATION_B = CALIBRATION_B
    core.CALIBRATION_C = CALIBRATION_C
    core.OVERLAP_CORRECTION = OVERLAP_CORRECTION
    core.QC_MODE = QC_MODE
    try:
        yield
    finally:
        core.REFERENCE_INTENSITY = original_reference
        core.CALIBRATION_B = original_b
        core.CALIBRATION_C = original_c
        core.OVERLAP_CORRECTION = original_overlap
        core.QC_MODE = original_qc_mode


def calculate_all(measurements):
    with cols_configuration():
        return core.calculate_all(measurements)


def convert_csv(csv_bytes: bytes) -> bytes:
    with cols_configuration():
        return core.convert_csv(csv_bytes)
