"""Anniversary lift: top-down from milestone comparables, bottom-up from four channels.

Top-down comps (all-in, before stripping the pieces we count elsewhere):
  Disneyland 60th: 16.769M (2014) -> 18.278M (2015) -> 17.943M (2016)
  Disneyland 50th: +8.5% vs top-50 North American parks +4.2%
  Musikfest 40th : 1.33M vs prior record 1.24M
"""


def lift_above_trend(prior, milestone, trend_growth):
    return milestone / prior - 1 - trend_growth


def top_down():
    dl60 = lift_above_trend(16_769_000, 18_278_000, 0.035)
    dl50 = 0.085 - 0.042
    musik = 1_330_000 / 1_240_000 - 1
    giveback = 17_943_000 / 18_278_000 - 1
    return {"disneyland_60th": dl60, "disneyland_50th": dl50, "musikfest_40th": musik, "disneyland_60th_giveback": giveback}


def bottom_up(attendance, local_pop, regular_share=0.80, years_history=10, mover_factor=0.6,
              reactivation=0.03, media_reach=0.20, media_conv=0.01, programming=0.01,
              multi_visit_share=0.30, extra_visit_rate=0.02, overlap=0.45):
    non_regular_visits = attendance * (1 - regular_share)
    lapsed_pool = non_regular_visits * years_history * mover_factor
    reactivated = lapsed_pool * reactivation
    non_attending_locals = local_pop - attendance * 0.70
    earned = non_attending_locals * media_reach * media_conv
    prog = attendance * programming
    extra = attendance * regular_share * multi_visit_share * extra_visit_rate * 5
    stacked = reactivated + earned + prog + extra
    return {
        "lapsed_pool": lapsed_pool, "reactivated": reactivated, "earned_media": earned,
        "programming": prog, "extra_regular_visits": extra, "stacked": stacked,
        "stacked_pct": stacked / attendance, "netted_pct": stacked / attendance * (1 - overlap),
    }
