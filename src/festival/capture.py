"""Capture-rate model for the local ring (500k, historically advertised to) and the outer ring (700k).

capture rate = visits from a ring / population of the ring

The festival was only marketed to the 500k, but roughly a quarter of past visits still came from
the neighboring metros (verified origin surveys put "local" at about 70 percent of festival crowds).
The calibrated model backs an aware-vs-unaware conversion out of the local ring and a distance
factor out of the outer ring, then asks what a campaign that raises outer-ring awareness does.
"""
import numpy as np


def rings(attendance, local_pop, outer_pop, outer_share, beyond_share=0.05):
    local_share = 1 - outer_share - beyond_share
    return {
        "local_visits": attendance * local_share,
        "outer_visits": attendance * outer_share,
        "beyond_visits": attendance * beyond_share,
        "local_capture": attendance * local_share / local_pop,
        "outer_capture": attendance * outer_share / outer_pop,
        "ratio_outer_to_local": (outer_share / outer_pop) / (local_share / local_pop),
    }


def huff_distance_factor(d_local, d_outer, exponent):
    """Propensity of an outer-ring resident relative to a local one under a power distance decay."""
    return (d_outer / d_local) ** (-exponent)


def calibrated(attendance, local_pop, outer_pop, outer_share=0.25, aware_local=0.75,
               aware_outer=0.30, awareness_gain=0.15, lift=2.0, beyond_share=0.05):
    """Aware residents convert at `lift` times the unaware rate (Wyoming: 14% vs 7%).

    Returns the implied conversions, the implied outer-ring distance factor, the outer-ring
    capture after the campaign lifts awareness, and the incremental visits.
    """
    r = rings(attendance, local_pop, outer_pop, outer_share, beyond_share)
    c_u = r["local_capture"] / (aware_local * lift + (1 - aware_local))
    c_a = lift * c_u
    d = r["outer_capture"] / (aware_outer * c_a + (1 - aware_outer) * c_u)
    new_aware = min(1.0, aware_outer + awareness_gain)
    co_new = (new_aware * c_a + (1 - new_aware) * c_u) * d
    return {
        "conv_aware_local": c_a,
        "conv_unaware_local": c_u,
        "implied_distance_factor": d,
        "outer_capture_before": r["outer_capture"],
        "outer_capture_after": co_new,
        "incremental_visits": (co_new - r["outer_capture"]) * outer_pop,
        "outer_ceiling_capture": c_a * d,
        "outer_ceiling_visits": c_a * d * outer_pop,
        "local_ceiling_visits": c_a * local_pop,
    }


def awareness_diffusion_path(start_awareness, p, q, years, c_a, c_u, d, outer_pop):
    """Outer-ring awareness spreads by a Bass process (p = campaign, q = word of mouth)."""
    F = start_awareness
    rows = []
    for yr in years:
        F = min(1.0, F + (p + q * F) * (1 - F))
        capture = (F * c_a + (1 - F) * c_u) * d
        rows.append({"year": yr, "awareness": F, "outer_capture": capture, "outer_visits": capture * outer_pop})
    return rows
