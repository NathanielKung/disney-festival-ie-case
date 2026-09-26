"""Marketing spend and reach expansion.

Four independent estimates of the new-market lift are produced so the deck can show their overlap:
1. Advertising elasticity on the whole spend increase.
2. Awareness x conversion funnel on the 700k new residents.
3. Bass diffusion year-one adoption of a distance-discounted market potential.
4. The calibrated capture-rate model (capture.py).
"""


def elasticity_lift(attendance, spend_increase, elasticity):
    """Percent change in demand = elasticity x percent change in spend."""
    return attendance * elasticity * spend_increase


def local_readings(attendance, elasticity, spend_increase, audience_multiple, spillover=0.005):
    """Reading A holds the local budget and sends the new money outward (spillover only).
    Reading B spreads the whole budget evenly, cutting per-capita intensity."""
    intensity_change = (1 + spend_increase) / audience_multiple - 1   # 1.2 / 2.4 - 1 = -0.5
    return {
        "A_local_lift_pct": spillover,
        "A_local_visits": attendance * spillover,
        "B_intensity_change": intensity_change,
        "B_local_lift_pct": elasticity * intensity_change,
        "B_local_visits": attendance * elasticity * intensity_change,
    }


def awareness_funnel(outer_pop, awareness, conversion):
    return outer_pop * awareness * conversion


def bass_year_one(outer_pop, local_capture_rate, distance_factor, p):
    """Market potential m = outer population x local capture x distance factor; year-one adoption = p x m."""
    m = outer_pop * local_capture_rate * distance_factor
    return {"market_potential": m, "year_one": p * m}


def bass_path(m, p, q, years=5):
    """New adopters per year for a market of potential m starting from zero."""
    cum, out = 0.0, []
    for _ in range(years):
        new = (p + q * cum / m) * (m - cum)
        cum += new
        out.append(new)
    return out
