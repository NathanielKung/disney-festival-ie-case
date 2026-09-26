"""Concert draw as a function of the act's combined Spotify and Apple Music monthly listeners.

Funnel: combined listeners -> listeners inside the festival catchment -> concert-motivated demand
-> split regulars (shift weekends) vs new (incremental) -> who shows up -> seats vs the 8,000 cap.
"""
US_POP = 340e6


def combined_listeners(spotify_ml, apple_ml=None, apple_ratio=0.35, overlap=0.07):
    apple = apple_ml if apple_ml is not None else spotify_ml * apple_ratio
    return (spotify_ml + apple) * (1 - overlap)


def concert(cml, local_pop=500e3, outer_pop=700e3, outer_weight=0.25, us_share=0.60, affinity=1.0,
            engagement=1.0, conv_free=0.10, show_up=0.70, regular_share=0.10, casual_share=0.40,
            grounds_guests=8200, capacity=8000):
    eff_pop = local_pop + outer_pop * outer_weight
    catchment = cml * us_share * (eff_pop / US_POP) * affinity
    demand = catchment * conv_free * engagement
    regulars, new = demand * regular_share, demand * (1 - regular_share)
    new_arrive, reg_arrive = new * show_up, regulars * show_up
    seat_demand = grounds_guests * casual_share + new_arrive + reg_arrive
    return {
        "catchment_listeners": catchment, "demand": demand, "incremental": new_arrive,
        "shifted": reg_arrive, "seat_demand": seat_demand, "overflow": max(0.0, seat_demand - capacity),
        "per_million_incremental": new_arrive / (cml / 1e6) if cml else 0.0,
    }


def listeners_for_incremental(target, **kw):
    """The model is linear in cml, so invert directly."""
    per_million = concert(1e6, **kw)["incremental"]
    return target / per_million * 1e6


def passes_to_issue(capacity, no_show_rate):
    return capacity / (1 - no_show_rate)
