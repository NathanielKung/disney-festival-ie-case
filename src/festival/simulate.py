"""Driver bridge, Monte Carlo scenarios, tornado sensitivity, weekend split and design day."""
import numpy as np

DRIVERS = {
    # name: (kind, low, mode, high)   kind: "mult" applied as (1+x), "add" as visits, "weather" as multiplier
    "baseline": ("level", 56800, 58500, 59800),
    "marketing": ("mult", -0.02, 0.005, 0.024),
    "anniversary": ("mult", 0.01, 0.035, 0.07),
    "macro": ("mult", -0.02, -0.008, 0.015),
    "reach": ("add", 1000, 2000, 4500),
    "concert": ("add", 400, 800, 2000),
    "weather": ("weather", 0.90, 1.00, 1.02),
}
CONCERT_SHIFT = 100
WEEKEND_SHARES = np.array([0.20, 0.25, 0.30, 0.25])
SHIFT_FROM = np.array([0.2, 0.3, 0.5, 0.0])


def total_from(d):
    return (d["baseline"] * (1 + d["marketing"]) * (1 + d["anniversary"]) * (1 + d["macro"])
            + d["reach"] + d["concert"]) * d["weather"]


def bridge(drivers=DRIVERS):
    m = {k: v[2] for k, v in drivers.items()}
    steps = [("2025 actual", 56000)]
    x = m["baseline"]; steps.append(("+ trend", x))
    x *= 1 + m["marketing"]; steps.append(("+ marketing", x))
    x *= 1 + m["anniversary"]; steps.append(("+ anniversary", x))
    x *= 1 + m["macro"]; steps.append(("+ macro", x))
    x += m["reach"]; steps.append(("+ reach", x))
    x += m["concert"]; steps.append(("+ concert", x))
    return steps, total_from(m)


def _draw(kind, lo, mode, hi, n, rng):
    if kind == "add":   # PERT for headcount drivers (smoother, more weight near mode)
        lam = 4
        a = 1 + lam * (mode - lo) / (hi - lo); b = 1 + lam * (hi - mode) / (hi - lo)
        return lo + (hi - lo) * rng.beta(a, b, n)
    return rng.triangular(lo, mode, hi, n)


def simulate(n=100_000, seed=21, drivers=DRIVERS, saturday_share=0.55):
    rng = np.random.default_rng(seed)
    d = {k: _draw(v[0], v[1], v[2], v[3], n, rng) for k, v in drivers.items()}
    nonconcert = (d["baseline"] * (1 + d["marketing"]) * (1 + d["anniversary"]) * (1 + d["macro"]) + d["reach"]) * d["weather"]
    concert = d["concert"] * d["weather"]
    total = nonconcert + concert
    wk = np.outer(nonconcert, WEEKEND_SHARES)
    wk[:, 3] += concert + CONCERT_SHIFT
    wk -= np.outer(np.full(n, CONCERT_SHIFT), SHIFT_FROM)
    sat3 = wk[:, 2] * saturday_share
    sat4 = (wk[:, 3] - concert - CONCERT_SHIFT) * saturday_share + concert + CONCERT_SHIFT
    pct = lambda a: [float(np.percentile(a, p)) for p in (10, 50, 90)]
    return {
        "total_p10_p50_p90": pct(total),
        "p_over_60k": float((total > 60000).mean()),
        "p_over_65k": float((total > 65000).mean()),
        "weekends_p10_p50_p90": [pct(wk[:, i]) for i in range(4)],
        "wk3_sat_p10_p50_p90": pct(sat3),
        "wk4_sat_p10_p50_p90": pct(sat4),
        "p_wk4_over_wk3": float((wk[:, 3] > wk[:, 2]).mean()),
        "p_wk4sat_over_wk3sat": float((sat4 > sat3).mean()),
        "design_day": float(np.median(sat3)),
    }


def tornado(drivers=DRIVERS):
    """Swing in total when each driver moves to its 10th and 90th percentile, others at mode.
    Percentiles of a triangular/PERT are taken from a large sample for consistency with simulate()."""
    rng = np.random.default_rng(7)
    m = {k: v[2] for k, v in drivers.items()}
    base = total_from(m)
    out = {}
    for k, v in drivers.items():
        s = _draw(v[0], v[1], v[2], v[3], 200_000, rng)
        lo, hi = np.percentile(s, 10), np.percentile(s, 90)
        d = dict(m); d[k] = lo; a = total_from(d); d[k] = hi; b = total_from(d)
        out[k] = {"low_delta": a - base, "high_delta": b - base, "swing": abs(b - a)}
    return base, out


def weekend_split_deterministic(total_nonconcert, concert_incremental, shares=WEEKEND_SHARES, saturday_share=0.55):
    wk = total_nonconcert * shares
    wk[3] += concert_incremental + CONCERT_SHIFT
    wk -= CONCERT_SHIFT * SHIFT_FROM
    sat = wk * saturday_share
    sat[3] = (wk[3] - concert_incremental - CONCERT_SHIFT) * saturday_share + concert_incremental + CONCERT_SHIFT
    return wk, sat
