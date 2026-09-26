"""Part 3: growth options. Capacity only pays where demand fills it (load-duration logic)."""

OPTION_ASSUMPTIONS = {
    "A": dict(label="Extend hours to 12 per day", hours_added=16, gross_share_of_season=0.07, redistribution=0.60,
              cost_formula="labor + lighting/security per extended day"),
    "B": dict(label="Add Fridays", hours_added=40, friday_vs_saturday=0.45, redistribution=0.35,
              cost_formula="labor + operations per added day + launch marketing"),
    "C": dict(label="Add a fifth weekend", hours_added=20, weekend_vs_average=0.65, redistribution=0.40,
              cost_formula="labor + operations + extra site week + launch marketing"),
}
COST = dict(staff=50, wage=22.0, lighting_security_per_extended_day=2500, ops_per_added_day=8000,
            extra_site_week=20000, launch_marketing_B=10000, launch_marketing_C=8000, wage_escalator=0.046)


def revenue_per_guest(admission=12.0, parking=1.5, percap_multiplier=1.0, cats=None):
    from .booths import CATEGORIES
    cats = cats or CATEGORIES
    share = sum(c["part"] * c["tx"] * c["spend"] * c["share"] for c in cats.values()) * percap_multiplier
    return admission + parking + share


def options(season, percap_multiplier=1.0, escalate_costs=True):
    rpg = revenue_per_guest(percap_multiplier=percap_multiplier)
    esc = 1 + COST["wage_escalator"] if escalate_costs else 1.0
    avg_sat = season * 0.55 / 4
    avg_wk = season / 4
    out = {}
    a = OPTION_ASSUMPTIONS["A"]
    gA = season * a["gross_share_of_season"]; nA = gA * (1 - a["redistribution"])
    cA = (COST["staff"] * COST["wage"] * esc * 2 * 8) + COST["lighting_security_per_extended_day"] * 8
    b = OPTION_ASSUMPTIONS["B"]
    gB = 4 * b["friday_vs_saturday"] * avg_sat; nB = gB * (1 - b["redistribution"])
    cB = COST["staff"] * COST["wage"] * esc * 11 * 4 + COST["ops_per_added_day"] * 4 + COST["launch_marketing_B"]
    c = OPTION_ASSUMPTIONS["C"]
    gC = c["weekend_vs_average"] * avg_wk; nC = gC * (1 - c["redistribution"])
    cC = COST["staff"] * COST["wage"] * esc * 11 * 2 + COST["ops_per_added_day"] * 2 + COST["extra_site_week"] + COST["launch_marketing_C"]
    for key, g, n, cst in (("A", gA, nA, cA), ("B", gB, nB, cB), ("C", gC, nC, cC)):
        hrs = OPTION_ASSUMPTIONS[key]["hours_added"]
        out[key] = dict(gross=g, net_new=n, net_new_pct=n / season, revenue=n * rpg, cost=cst,
                        net=n * rpg - cst, hours_added=hrs, capacity_factor=g / (hrs * season / 80),
                        peak_relief=g - n)
    return rpg, out


SCORES = {"A": dict(fin=1, gx=2, feas=4, risk=3), "B": dict(fin=5, gx=4, feas=3, risk=3), "C": dict(fin=3, gx=3, feas=4, risk=2)}
WEIGHTS = dict(fin=0.35, gx=0.30, feas=0.20, risk=0.15)


def scorecard(scores=SCORES, weights=WEIGHTS):
    return {k: sum(weights[c] * s[c] for c in weights) for k, s in scores.items()}
