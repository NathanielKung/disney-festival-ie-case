"""Run every model in order and write outputs/results.json.

    python scripts/run_all.py

Part 1 (forecast) -> design day -> Part 2 (booths) -> Part 3 (options).
"""
import json, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from festival import load_attendance, load_case_parameters, load_anchors, OUTPUTS
from festival import baseline as bl, capture as cp, reach as rc, concert as cc, anniversary as an, macro as mc
from festival import simulate as sm, booths as bo, options as op


def main():
    att = load_attendance(); P = load_case_parameters(); A = load_anchors()
    years, y = att["year"].tolist(), att["attendance"].tolist()
    res = {}

    # ---- Part 1a. Organic baseline
    res["baseline"] = bl.baseline(years, y)

    # ---- Part 1b. Capture rates and reach
    outer_share = A.loc["outer_share_of_visits", "base"]
    rings = cp.rings(56000, P["local_population"], P["outer_population"], outer_share)
    cal = cp.calibrated(56000, P["local_population"], P["outer_population"], outer_share)
    huff = {e: cp.huff_distance_factor(12, 45, e) for e in (1.0, 1.5, 2.0)}
    res["capture"] = {"rings": rings, "calibrated": cal, "huff_factor_12mi_vs_45mi": huff,
                      "diffusion_path": cp.awareness_diffusion_path(0.30, 0.10, 0.38, range(2026, 2031),
                                                                    cal["conv_aware_local"], cal["conv_unaware_local"],
                                                                    cal["implied_distance_factor"], P["outer_population"])}
    el = A.loc["ad_elasticity_short"]
    funnel = rc.awareness_funnel(P["outer_population"], A.loc["campaign_awareness", "base"], A.loc["aware_conversion", "base"])
    bass = rc.bass_year_one(P["outer_population"], 56000 / P["local_population"], 0.25, A.loc["bass_p", "base"])
    bass_hi = rc.bass_year_one(P["outer_population"], 56000 / P["local_population"], 0.25, A.loc["bass_p", "high"])
    res["reach"] = {
        "elasticity_low_high": [rc.elasticity_lift(56000, P["marketing_spend_increase"], el["base"]),
                                rc.elasticity_lift(56000, P["marketing_spend_increase"], el["high"])],
        "local_readings": rc.local_readings(56000, el["base"], P["marketing_spend_increase"], 2.4),
        "funnel": funnel, "bass_year_one_low_high": [bass["year_one"], bass_hi["year_one"]],
        "bass_market_potential": bass["market_potential"], "bass_path_p03": rc.bass_path(bass["market_potential"], 0.03, 0.38),
        "capture_model": cal["incremental_visits"], "carried": 2000, "range": [1000, 4500],
    }

    # ---- Part 1c. Anniversary, concert, macro
    res["anniversary"] = {"top_down": an.top_down(), "bottom_up": an.bottom_up(56000, P["local_population"]), "carried": 0.035}
    tiers = {f"{t:g}M": cc.concert(t * 1e6) for t in (1, 5, 10, 20, 40)}
    res["concert"] = {"tiers": tiers, "regional_1M": cc.concert(1e6, affinity=4.0, engagement=1.5),
                      "playlist_10M": cc.concert(10e6, affinity=0.5, engagement=0.6),
                      "cml_for_2000": cc.listeners_for_incremental(2000), "carried_tier": "10M",
                      "passes_to_issue": cc.passes_to_issue(P["concert_capacity"], A.loc["free_event_no_show", "base"])}
    res["macro"] = mc.macro_effects(56000)

    # ---- Part 1d. Bridge, simulation, tornado
    steps, fair = sm.bridge(); res["bridge"] = {"steps": steps, "fair_weather_total": fair}
    sim = sm.simulate(); res["simulation"] = sim
    tb, tor = sm.tornado(); res["tornado"] = {"base": tb, "drivers": tor}
    season = round(sim["total_p10_p50_p90"][1], -2); design = round(sim["design_day"], -2)
    res["season"], res["design_day"] = season, design

    # ---- Part 2
    percap = 1 + A.loc["macro_percap", "base"]
    floors = {k: bo.queue_floor(design, k) for k in bo.CATEGORIES}
    best = bo.search_best_mix(season, design, floors, percap_multiplier=percap)
    mixes = {"profit_max": {"F&B": best[1], "Merch": best[2], "A&E": best[3]},
             "balanced": {"F&B": 60, "Merch": 28, "A&E": 16}, "variety": {"F&B": 56, "Merch": 32, "A&E": 16}}
    res["part2"] = {"floors": floors, "peak_hour_demand": {k: bo.peak_hour_demand(design, k) for k in bo.CATEGORIES},
                    "service_rate": {k: bo.service_rate(k) for k in bo.CATEGORIES}, "percap_multiplier": percap,
                    "mixes": {n: bo.evaluate_mix(season, design, m, floors, percap) for n, m in mixes.items()},
                    "low_case_check": bo.evaluate_mix(round(sim["total_p10_p50_p90"][0], -2), round(sim["wk3_sat_p10_p50_p90"][0], -2), mixes["balanced"], floors, percap)}

    # ---- Part 3
    rpg, opts = op.options(season, percap_multiplier=percap)
    res["part3"] = {"revenue_per_guest": rpg, "options": opts, "scores": op.scorecard(),
                    "design_day_2029": (season + 3 * 1500 + 5700) * 0.30 * 0.55}

    OUTPUTS.mkdir(exist_ok=True)
    (OUTPUTS / "results.json").write_text(json.dumps(res, indent=1, default=float))
    print(json.dumps({"season_p50": season, "design_day": design, "p10_p90": sim["total_p10_p50_p90"],
                      "best_mix": best, "options_net": {k: v["net"] for k, v in opts.items()}}, indent=1, default=float))


if __name__ == "__main__":
    main()
