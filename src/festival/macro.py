"""Macroeconomic backdrop for the 2026 season, turned into three explicit mechanisms.

September 2026 readings (see data/research_anchors.csv for sources):
  UMich consumer sentiment 48.1 vs 55.1 a year earlier; 1-year inflation expectations 4.6%
  AAA national gasoline $4.49, about 25% above a year ago, a record for late September
  Unemployment 4.1%, payrolls +162k, retail and food-service sales +6.0% y/y (August)

Mechanism 1, staycation substitution: a $12 local festival is the cheap alternative to travel.
  In the 2009 recession fairs reported booming business as families skipped out-of-town trips.
  Positive for local-ring attendance.
Mechanism 2, fuel cost on the outer ring: a 90-mile round trip costs a few dollars more at $4.49.
  Short-run gasoline demand elasticity is only -0.03 to -0.08 (Hughes, Knittel, Sperling 2008),
  but discretionary day trips are more elastic, so we apply -0.1 to -0.3 to the outer ring.
Mechanism 3, real-income squeeze: sentiment near record lows, yet spending is still growing in
  nominal terms. Small drag on visits, larger drag on real spend per guest (trade-down).
"""


def fuel_cost_per_trip(round_trip_miles, mpg, price):
    return round_trip_miles / mpg * price


def macro_effects(attendance, local_share=0.70, outer_share=0.25, staycation_lift=0.01,
                  gas_now=4.49, gas_prior=3.60, trip_fuel_elasticity=-0.20, sentiment_drag=-0.003,
                  percap_change=-0.04):
    gas_change = gas_now / gas_prior - 1
    outer_effect = trip_fuel_elasticity * gas_change          # % change in outer-ring visits
    local_effect = staycation_lift
    net = local_share * local_effect + outer_share * outer_effect + sentiment_drag
    return {
        "gas_change_pct": gas_change,
        "outer_visit_effect_pct": outer_effect,
        "outer_visits_effect": attendance * outer_share * outer_effect,
        "local_visits_effect": attendance * local_share * local_effect,
        "sentiment_visits_effect": attendance * sentiment_drag,
        "net_attendance_pct": net,
        "net_attendance_visits": attendance * net,
        "percap_change": percap_change,
        "extra_fuel_cost_per_car": fuel_cost_per_trip(90, 25, gas_now) - fuel_cost_per_trip(90, 25, gas_prior),
    }
