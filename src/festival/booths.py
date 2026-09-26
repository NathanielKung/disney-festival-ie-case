"""Part 2: booth mix. Design-day demand -> Erlang C queue floors -> vendor economics -> mix search.

Attendance fixes demand; booths are capacity. Extra booths raise capture (share of intending
buyers who transact instead of balking) with diminishing returns, and dilute each vendor's sales.
"""
import math

CATEGORIES = {
    "F&B":   dict(spend=21.5, part=0.775, tx=1.6, units=1, fee=2000, share=0.12, cost=900, floor=18000,
                  svc_points=2, svc_min=2.5, peak=0.18, wait_target=10, min_var=0),
    "Merch": dict(spend=16.0, part=0.325, tx=1.1, units=1, fee=900, share=0.10, cost=500, floor=5000,
                  svc_points=1, svc_min=3.0, peak=0.15, wait_target=5, min_var=20),
    "A&E":   dict(spend=8.5, part=0.40, tx=1.3, units=2, fee=700, share=0.08, cost=600, floor=4000,
                  svc_points=None, svc_min=None, peak=0.15, wait_target=15, min_var=12, ohrc=60),
}


def service_rate(cat):
    c = CATEGORIES[cat]
    if c["svc_points"] is None:
        return c["ohrc"]
    return (60 / c["svc_min"]) * c["svc_points"]


def erlang_c_wait(lam, mu, c):
    """Mean wait in queue (minutes) for an M/M/c system. lam, mu per hour."""
    a = lam / mu
    if c <= a:
        return math.inf
    s = sum(a ** k / math.factorial(k) for k in range(c))
    last = a ** c / math.factorial(c) * c / (c - a)
    p_wait = last / (s + last)
    return p_wait / (c * mu - lam) * 60


def peak_hour_demand(design_day, cat):
    c = CATEGORIES[cat]
    return design_day * c["part"] * c["tx"] * c["peak"]


def queue_floor(design_day, cat):
    lam, mu = peak_hour_demand(design_day, cat), service_rate(cat)
    n = 1
    while erlang_c_wait(lam, mu, n) > CATEGORIES[cat]["wait_target"]:
        n += 1
    return n


def season_sales_potential(season, cat, percap_multiplier=1.0):
    c = CATEGORIES[cat]
    return season * c["part"] * c["tx"] * c["spend"] * percap_multiplier


def capture(n, floor_n, at_floor=0.85):
    """Saturating capture curve: 1 - exp(-n/n0), with n0 chosen so capture = at_floor at the queue floor."""
    n0 = floor_n / (-math.log(1 - at_floor))
    return 1 - math.exp(-n / n0)


def evaluate_mix(season, design_day, mix, floors, percap_multiplier=1.0):
    rows, total = {}, 0.0
    for cat, n in mix.items():
        c = CATEGORIES[cat]
        cap = capture(n, floors[cat])
        sales = season_sales_potential(season, cat, percap_multiplier) * cap
        contrib = c["fee"] * n + c["share"] * sales - c["cost"] * n
        lam, mu = peak_hour_demand(design_day, cat), service_rate(cat)
        rows[cat] = dict(booths=n, capture=cap, sales=sales, sales_per_booth=sales / n,
                         wait_min=erlang_c_wait(lam, mu, n), contribution=contrib,
                         above_floor=sales / n >= c["floor"])
        total += contrib
    units = sum(CATEGORIES[k]["units"] * n for k, n in mix.items())
    return {"rows": rows, "total_contribution": total, "units": units}


def search_best_mix(season, design_day, floors, budget=120, percap_multiplier=1.0):
    best = None
    for F in range(floors["F&B"], budget + 1):
        for M in range(floors["Merch"], budget + 1):
            for A in range(floors["A&E"], budget // 2 + 1):
                if F + M + 2 * A > budget:
                    break
                r = evaluate_mix(season, design_day, {"F&B": F, "Merch": M, "A&E": A}, floors, percap_multiplier)
                if all(v["above_floor"] for v in r["rows"].values()) and (best is None or r["total_contribution"] > best[0]):
                    best = (r["total_contribution"], F, M, A)
    return best
