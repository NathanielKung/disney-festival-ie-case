"""Organic baseline: five trend methods, holdout backtests, forecast combination, prediction interval.

Methods
-------
A. OLS linear trend on all years            (constant headcount growth)
B. Log-linear trend on all years            (constant percentage growth, i.e. CAGR)
C. OLS linear trend on the last three years (post-pandemic regime only)
D. Log-linear trend on the last three years
E. Holt double exponential smoothing        (level + trend, recent years weighted more)

The holdout test hides the last actual year, forecasts it from the earlier years, and
reports the percentage miss. Negative misses across every method mean the series has
momentum the fitted trend under-calls.
"""
from dataclasses import dataclass
import numpy as np
from scipy import stats


@dataclass
class TrendFit:
    name: str
    forecast: float
    slope: float          # per-year change (linear) or growth rate (log-linear)
    params: dict


def ols_linear(x, y, xf):
    x = np.asarray(x, float); y = np.asarray(y, float)
    slope, intercept = np.polyfit(x, y, 1)
    return TrendFit("ols_linear", slope * xf + intercept, slope, {"intercept": intercept})


def log_linear(x, y, xf):
    x = np.asarray(x, float); y = np.asarray(y, float)
    b, a = np.polyfit(x, np.log(y), 1)
    return TrendFit("log_linear", float(np.exp(b * xf + a)), float(np.exp(b) - 1), {"log_intercept": a})


def holt(y, alpha, beta, h=1):
    """Holt's linear-trend smoothing. Initial level = y[0], initial trend = y[1]-y[0]."""
    y = np.asarray(y, float)
    level, trend = y[0], y[1] - y[0]
    for t in range(1, len(y)):
        new_level = alpha * y[t] + (1 - alpha) * (level + trend)
        trend = beta * (new_level - level) + (1 - beta) * trend
        level = new_level
    return level + h * trend


def holt_sse(y, alpha, beta):
    y = np.asarray(y, float)
    level, trend, sse = y[0], y[1] - y[0], 0.0
    for t in range(1, len(y)):
        sse += (y[t] - (level + trend)) ** 2
        new_level = alpha * y[t] + (1 - alpha) * (level + trend)
        trend = beta * (new_level - level) + (1 - beta) * trend
        level = new_level
    return sse


def holt_optimize(y, grid=np.linspace(0.1, 0.9, 17)):
    best = None
    for a in grid:
        for b in grid:
            s = holt_sse(y, a, b)
            if best is None or s < best[0]:
                best = (s, float(a), float(b))
    return best[1], best[2]


def prediction_interval(x, y, xf, level=0.80):
    """Classic OLS prediction interval for a new observation at xf."""
    x = np.asarray(x, float); y = np.asarray(y, float)
    n = len(x)
    slope, intercept = np.polyfit(x, y, 1)
    resid = y - (slope * x + intercept)
    s = np.sqrt((resid ** 2).sum() / (n - 2))
    se = s * np.sqrt(1 + 1 / n + (xf - x.mean()) ** 2 / ((x - x.mean()) ** 2).sum())
    t = stats.t.ppf(0.5 + level / 2, n - 2)
    f = slope * xf + intercept
    return f - t * se, f + t * se, se


def holdout(x, y, target_idx):
    """Forecast y[target_idx] using only observations before it. Returns % error by method."""
    x = np.asarray(x, float); y = np.asarray(y, float)
    xt, yt, xf, actual = x[:target_idx], y[:target_idx], x[target_idx], y[target_idx]
    out = {
        "A_ols_all": ols_linear(xt, yt, xf).forecast,
        "B_loglin_all": log_linear(xt, yt, xf).forecast,
        "E_holt": holt(yt, 0.5, 0.5),
    }
    if len(xt) >= 3:
        out["C_ols_recent3"] = ols_linear(xt[-3:], yt[-3:], xf).forecast
        out["D_loglin_recent3"] = log_linear(xt[-3:], yt[-3:], xf).forecast
    return {k: (v - actual) / actual for k, v in out.items()}


DEFAULT_WEIGHTS = {"A_ols_all": 0.15, "B_loglin_all": 0.15, "C_ols_recent3": 0.25,
                   "D_loglin_recent3": 0.20, "E_holt": 0.25}


def baseline(years, attendance, forecast_year=2026, weights=DEFAULT_WEIGHTS):
    years = np.asarray(years); att = np.asarray(attendance, float)
    x = years - years[0] + 1
    xf = forecast_year - years[0] + 1
    A = ols_linear(x, att, xf)
    B = log_linear(x, att, xf)
    C = ols_linear(x[-3:], att[-3:], xf)
    D = log_linear(x[-3:], att[-3:], xf)
    alpha, beta = holt_optimize(att)
    E = holt(att, alpha, beta)
    fits = {"A_ols_all": A.forecast, "B_loglin_all": B.forecast, "C_ols_recent3": C.forecast,
            "D_loglin_recent3": D.forecast, "E_holt": E}
    combined = sum(weights[k] * fits[k] for k in fits)
    lo, hi, se = prediction_interval(x, att, xf, 0.80)
    return {
        "fits": fits,
        "slopes": {"A_per_year": A.slope, "B_cagr": B.slope, "C_per_year": C.slope, "D_cagr": D.slope},
        "holt_params": {"alpha": alpha, "beta": beta},
        "weights": weights,
        "combined": combined,
        "pi80_linear": (lo, hi),
        "pi80_se": se,
        "holdout_2024": holdout(x, att, len(att) - 2),
        "holdout_2025": holdout(x, att, len(att) - 1),
    }
