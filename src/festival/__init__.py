"""Festival attendance, capacity and growth models for the WDPR Fall 2026 IE case.

Every module is a plain function of its inputs so the same numbers can be reproduced
in the Excel workbook (src/festival/workbook.py writes the formulas) and checked by tests.
"""
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
DATA = ROOT / "data"
OUTPUTS = ROOT / "outputs"


def load_attendance() -> pd.DataFrame:
    return pd.read_csv(DATA / "attendance_2021_2025.csv")


def load_case_parameters() -> dict:
    df = pd.read_csv(DATA / "case_parameters.csv")
    return dict(zip(df["parameter"], df["value"]))


def load_anchors() -> pd.DataFrame:
    return pd.read_csv(DATA / "research_anchors.csv").set_index("key")
