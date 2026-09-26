from festival import capture as cp, reach as rc


def test_capture_rates_reconcile_to_attendance():
    r = cp.rings(56000, 500e3, 700e3, 0.25)
    assert abs(r["local_visits"] + r["outer_visits"] + r["beyond_visits"] - 56000) < 1e-6
    assert abs(r["local_capture"] - 0.0784) < 1e-4 and abs(r["outer_capture"] - 0.02) < 1e-9


def test_calibration_reproduces_observed_captures_before_campaign():
    c = cp.calibrated(56000, 500e3, 700e3, 0.25)
    local = 0.75 * c["conv_aware_local"] + 0.25 * c["conv_unaware_local"]
    assert abs(local - 0.0784) < 1e-6
    assert abs(c["outer_capture_before"] - 0.02) < 1e-9
    assert c["incremental_visits"] > 0 and c["outer_ceiling_visits"] > 14000


def test_huff_factor_decreases_with_exponent():
    assert cp.huff_distance_factor(12, 45, 1.0) > cp.huff_distance_factor(12, 45, 1.5) > cp.huff_distance_factor(12, 45, 2.0)


def test_reach_methods_overlap():
    el = rc.elasticity_lift(56000, 0.20, 0.12)
    funnel = rc.awareness_funnel(700e3, 0.25, 0.017)
    bass = rc.bass_year_one(700e3, 0.112, 0.25, 0.03)["year_one"]
    assert 1000 < el < 3000 and 2500 < funnel < 3500 and 400 < bass < 1000


def test_bass_path_sums_below_potential():
    path = rc.bass_path(19600, 0.03, 0.38, years=10)
    assert all(a > 0 for a in path) and sum(path) < 19600
