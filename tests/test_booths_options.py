import math
from festival import booths as bo, options as op, concert as cc, simulate as sm


def test_erlang_c_reduces_to_mm1():
    lam, mu = 5.0, 8.0
    mm1_wq_min = lam / (mu * (mu - lam)) * 60
    assert abs(bo.erlang_c_wait(lam, mu, 1) - mm1_wq_min) < 1e-9


def test_erlang_c_wait_falls_with_servers_and_infinite_when_unstable():
    assert bo.erlang_c_wait(100, 10, 12) > bo.erlang_c_wait(100, 10, 15) > bo.erlang_c_wait(100, 10, 20)
    assert math.isinf(bo.erlang_c_wait(100, 10, 10))


def test_queue_floor_meets_wait_target():
    for cat in bo.CATEGORIES:
        n = bo.queue_floor(10100, cat)
        lam, mu = bo.peak_hour_demand(10100, cat), bo.service_rate(cat)
        assert bo.erlang_c_wait(lam, mu, n) <= bo.CATEGORIES[cat]["wait_target"]
        assert bo.erlang_c_wait(lam, mu, n - 1) > bo.CATEGORIES[cat]["wait_target"]


def test_capture_curve_anchored_at_floor():
    assert abs(bo.capture(48, 48) - 0.85) < 1e-9 and bo.capture(96, 48) > 0.97


def test_mix_uses_all_units_and_clears_floors():
    floors = {k: bo.queue_floor(10100, k) for k in bo.CATEGORIES}
    best = bo.search_best_mix(62100, 10100, floors)
    assert best[1] + best[2] + 2 * best[3] == 120
    r = bo.evaluate_mix(62100, 10100, {"F&B": best[1], "Merch": best[2], "A&E": best[3]}, floors)
    assert all(v["above_floor"] for v in r["rows"].values())


def test_options_ranking_and_capacity_factor():
    rpg, o = op.options(62100)
    assert o["B"]["net"] > o["C"]["net"] > o["A"]["net"]
    assert all(0 < v["capacity_factor"] < 1 for v in o.values())
    assert 15 < rpg < 20


def test_concert_linear_in_listeners_and_capped():
    a, b = cc.concert(5e6), cc.concert(10e6)
    assert abs(b["incremental"] - 2 * a["incremental"]) < 1e-6
    big = cc.concert(200e6); assert big["overflow"] > 0
    assert abs(cc.listeners_for_incremental(cc.concert(10e6)["incremental"]) - 10e6) < 1


def test_simulation_percentiles_ordered_and_bridge_consistent():
    s = sm.simulate(n=20000, seed=1)
    p10, p50, p90 = s["total_p10_p50_p90"]; assert p10 < p50 < p90
    steps, fair = sm.bridge(); assert abs(steps[-1][1] - fair) < 1e-6
    assert p50 < fair, "downside-skewed weather pulls the median below the fair-weather bridge"
