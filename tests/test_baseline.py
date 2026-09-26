import numpy as np
from festival import baseline as bl

YEARS = [2021, 2022, 2023, 2024, 2025]
ATT = [49000, 48500, 50000, 53500, 56000]


def test_ols_matches_numpy_polyfit():
    x = np.arange(1, 6); f = bl.ols_linear(x, ATT, 6)
    slope, intercept = np.polyfit(x, ATT, 1)
    assert abs(f.forecast - (slope * 6 + intercept)) < 1e-6
    assert abs(f.slope - 1900) < 1e-6          # 1,900 guests per year on the case data


def test_log_linear_recovers_exact_growth():
    y = [100 * 1.05 ** t for t in range(5)]
    f = bl.log_linear(np.arange(5), y, 5)
    assert abs(f.slope - 0.05) < 1e-9 and abs(f.forecast - 100 * 1.05 ** 5) < 1e-6


def test_holt_reproduces_linear_series():
    y = [10, 12, 14, 16, 18]
    assert abs(bl.holt(y, 0.5, 0.5) - 20) < 1e-9


def test_holdout_is_negative_for_every_method():
    x = np.arange(1, 6)
    h = bl.holdout(x, ATT, 4)
    assert all(v < 0 for v in h.values()), "series has momentum: every method should under-call 2025"


def test_combined_baseline_in_expected_band():
    r = bl.baseline(YEARS, ATT)
    assert 58000 < r["combined"] < 59500
    lo, hi = r["pi80_linear"]; assert lo < r["fits"]["A_ols_all"] < hi
    assert abs(sum(r["weights"].values()) - 1) < 1e-9
