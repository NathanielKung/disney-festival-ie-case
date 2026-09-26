"""Build the Excel workbook: inputs in blue, every derived number a live formula, sources beside inputs.

    python scripts/build_workbook.py  ->  outputs/WDPR_Festival_Model.xlsx

Sheet map
  README      legend and sheet guide
  Sources     every research anchor with value, range, and URL
  Inputs      the assumption register (named cells used by every other sheet)
  Baseline    five trend methods, holdout test, combined baseline, prediction interval, chart
  Capture     capture rates by ring, calibrated model, ceiling, five-year diffusion path, chart
  Reach       elasticity, funnel, Bass, capture-model reconciliation, chart
  Anniversary comparables and bottom-up channels
  Concert     listener funnel and act-size tiers, chart
  Macro       September 2026 readings and the three mechanisms
  Bridge      driver bridge, deterministic low/base/high, Monte Carlo reference, chart
  Sensitivity tornado table (each driver at low and high, others at base), chart
  Weekends    weekend split, Saturday split, design day, chart
  Part2       design-day demand, Erlang C floors, capture curve, three mixes, chart
  Part3       revenue per guest, option economics, scorecard, chart
Monte Carlo percentiles come from src/festival/simulate.py (100,000 draws) and are pasted as
values on the Bridge sheet, labeled as such. Everything else recalculates from Inputs.
"""
import json, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from festival import DATA, OUTPUTS

FONT = "Arial"
BLUE, BLACK, GREEN, MUTED = "0000FF", "000000", "008000", "666666"
HDR_FILL = PatternFill("solid", fgColor="0D2A5C"); HDR_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
KEY_FILL = PatternFill("solid", fgColor="FFF2CC"); SEC_FONT = Font(name=FONT, bold=True, size=12, color="0D2A5C")
THIN = Side(style="thin", color="D9D9D9")
NAMES = {}


def style(ws, cell, kind="formula", fmt=None, bold=False):
    c = ws[cell] if isinstance(cell, str) else cell
    color = {"input": BLUE, "formula": BLACK, "link": GREEN, "muted": MUTED}[kind]
    c.font = Font(name=FONT, color=color, size=10, bold=bold)
    if fmt: c.number_format = fmt
    return c


def header(ws, row, labels, col=1):
    for i, l in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=l); c.fill = HDR_FILL; c.font = HDR_FONT
        c.alignment = Alignment(wrap_text=True, vertical="center")


def section(ws, row, text, col=1):
    c = ws.cell(row=row, column=col, value=text); c.font = SEC_FONT


def note(ws, row, text, col=1):
    c = ws.cell(row=row, column=col, value=text); c.font = Font(name=FONT, italic=True, size=9, color=MUTED)
    c.alignment = Alignment(wrap_text=False)


def widths(ws, w):
    for i, x in enumerate(w, 1): ws.column_dimensions[get_column_letter(i)].width = x


def name(wb, nm, sheet, cell):
    wb.defined_names[nm] = DefinedName(nm, attr_text=f"'{sheet}'!${cell[0]}${cell[1:]}")
    NAMES[nm] = f"'{sheet}'!${cell[0]}${cell[1:]}"


def bar(ws, anchor, title, cats, vals, ytitle="", w=16, h=8, series_title=None, horizontal=False):
    ch = BarChart(); ch.type = "bar" if horizontal else "col"; ch.title = title; ch.style = 10
    ch.y_axis.title = ytitle; ch.legend = None if series_title is None else ch.legend
    ch.add_data(vals, titles_from_data=False); ch.set_categories(cats)
    if series_title: ch.series[0].tx = SeriesLabel(v=series_title)
    ch.series[0].graphicalProperties.solidFill = "1F5FBF"
    ch.width, ch.height = w, h; ws.add_chart(ch, anchor)
    return ch


def build(results_path=OUTPUTS / "results.json", out_path=OUTPUTS / "WDPR_Festival_Model.xlsx"):
    R = json.loads(Path(results_path).read_text())
    anchors = pd.read_csv(DATA / "research_anchors.csv")
    case = pd.read_csv(DATA / "case_parameters.csv")
    att = pd.read_csv(DATA / "attendance_2021_2025.csv")
    wb = Workbook()

    # ------------------------------------------------------------------ README
    ws = wb.active; ws.title = "README"; widths(ws, [110])
    lines = ["WDPR Fall 2026 IE Case, Autumn Harvest Festival model", "",
             "How to read this workbook",
             "Blue text = hardcoded input with a source (edit these). Black = formula. Green = link to another sheet. Yellow fill = key assumption.",
             "Every derived number is a live formula that traces back to the Inputs sheet. Change an input and the bridge, weekends, booths and options all move.",
             "Monte Carlo percentiles (P10/P50/P90) are computed by src/festival/simulate.py in the GitHub repo and pasted as values on the Bridge sheet, labeled as such.",
             "", "Sheet guide", "Sources: every research anchor with value, range and URL.", "Inputs: the assumption register. Named cells drive every other sheet.",
             "Baseline: five trend methods on 2021 to 2025, holdout backtest, weighted combination, 80% prediction interval.",
             "Capture: attendance divided by ring population for the 500k local ring and the 700k neighboring ring; calibrated model; five-year diffusion path.",
             "Reach: four independent estimates of the new-market lift and their overlap.", "Anniversary: milestone comparables and the four bottom-up channels.",
             "Concert: listener funnel by act size.", "Macro: September 2026 readings and the three mechanisms they drive.",
             "Bridge: 2025 actual to 2026 expected, step by step; deterministic low/base/high; Monte Carlo reference.",
             "Sensitivity: tornado table.", "Weekends: weekend split, Saturday split, design day.",
             "Part2: design-day demand, Erlang C queue floors, capture curve, three candidate booth mixes.", "Part3: option economics and scorecard.",
             "", "Code: github.com/NathanielKung/disney-festival-ie-case (models, tests, and this workbook's builder)."]
    for i, l in enumerate(lines, 1):
        c = ws.cell(row=i, column=1, value=l); c.font = Font(name=FONT, size=10, bold=(i in (1, 3, 8)))

    # ------------------------------------------------------------------ Sources
    ws = wb.create_sheet("Sources"); widths(ws, [30, 40, 10, 10, 10, 14, 90, 70])
    header(ws, 1, ["Key", "Parameter", "Low", "Base", "High", "Unit", "Source", "URL"])
    for i, r in anchors.iterrows():
        row = i + 2
        for j, k in enumerate(["key", "label", "low", "base", "high", "unit", "source", "url"], 1):
            c = ws.cell(row=row, column=j, value=r[k]); c.font = Font(name=FONT, size=9, color=BLUE if k in ("low", "base", "high") else BLACK)
            if k == "url": c.hyperlink = r[k]; c.font = Font(name=FONT, size=9, color="1F5FBF", underline="single")
    ws.freeze_panes = "A2"

    # ------------------------------------------------------------------ Inputs
    ws = wb.create_sheet("Inputs"); widths(ws, [34, 14, 14, 14, 14, 70])
    header(ws, 1, ["Input", "Low", "Base", "High", "Unit", "Source / note"])
    inputs = [
        ("Case data", None),
        ("Attendance 2025", 56000, 56000, 56000, "visits", "Case", "Att2025"),
        ("Local ring population", 500000, 500000, 500000, "people", "Case: local metro", "PopLocal"),
        ("Neighboring ring population", 700000, 700000, 700000, "people", "Case: 1.2M reach minus 500k", "PopOuter"),
        ("Marketing spend increase", 0.20, 0.20, 0.20, "share", "Case", "SpendUp"),
        ("Audience multiple (1.2M / 0.5M)", 2.4, 2.4, 2.4, "x", "Case", "AudMult"),
        ("Concert capacity", 8000, 8000, 8000, "people", "Case", "ConcertCap"),
        ("Weekend share 1", 0.20, 0.20, 0.20, "share", "Case", "WkS1"),
        ("Weekend share 2", 0.25, 0.25, 0.25, "share", "Case", "WkS2"),
        ("Weekend share 3", 0.30, 0.30, 0.30, "share", "Case", "WkS3"),
        ("Weekend share 4", 0.25, 0.25, 0.25, "share", "Case", "WkS4"),
        ("Part 1 drivers", None),
        ("Organic baseline 2026", 56800, "=Baseline!C24", 59800, "visits", "Base links to the Baseline sheet combination; low/high bracket the five methods", "Base2026"),
        ("Local marketing spillover", -0.02, 0.005, 0.024, "share", "Reading A: local budget held. Range from elasticity 0.12 x intensity change", "MktLift"),
        ("Anniversary lift (net)", 0.01, 0.035, 0.07, "share", "Sources sheet: anniversary_lift", "AnnivLift"),
        ("Macro net attendance effect", -0.02, "=Macro!B24", 0.015, "share", "Base links to the Macro sheet", "MacroLift"),
        ("Reach expansion lift", 1000, 2000, 4500, "visits", "Overlap of four methods on the Reach sheet", "ReachAdd"),
        ("Concert incremental guests", 400, 800, 2000, "visits", "Concert sheet, 10M-listener tier; range 5M to 25M", "ConcertAdd"),
        ("Concert shifted regulars", 100, 100, 100, "visits", "Concert sheet", "ConcertShift"),
        ("Weather multiplier", 0.90, 1.00, 1.02, "x", "Sources sheet: weather_multiplier", "Weather"),
        ("Saturday share of weekend", 0.52, 0.55, 0.58, "share", "Sources sheet: saturday_share", "SatShare"),
        ("Research anchors", None),
        ("Advertising elasticity (short run)", 0.09, 0.12, 0.24, "elasticity", "Sethuraman, Tellis, Briesch 2011", "AdElast"),
        ("Bass p (innovation)", 0.03, 0.03, 0.10, "rate", "Sultan, Farley, Lehmann 1990", "BassP"),
        ("Bass q (imitation)", 0.30, 0.38, 0.45, "rate", "Sultan, Farley, Lehmann 1990", "BassQ"),
        ("Campaign awareness, new ring", 0.15, 0.25, 0.35, "share", "State tourism ROI studies (36 to 47% at higher spend)", "Aware"),
        ("Conversion among aware", 0.013, 0.017, 0.025, "share", "Washington 77k trips / 2.4M aware households; Wyoming 7% to 14%", "AwareConv"),
        ("Share of 2025 visits from neighboring ring", 0.15, 0.25, 0.35, "share", "Verified origin surveys (Bolton 72% local, Big Sur 70%)", "OuterShare"),
        ("Share of visits from beyond both rings", 0.05, 0.05, 0.05, "share", "Team assumption", "BeyondShare"),
        ("Huff distance exponent", 1.0, 1.5, 2.0, "exponent", "Esri Huff model documentation", "HuffExp"),
        ("Distance local (miles)", 12, 12, 12, "miles", "Team assumption", "DistLocal"),
        ("Distance outer (miles)", 45, 45, 45, "miles", "Team assumption", "DistOuter"),
        ("Distance weight, outer ring", 0.14, 0.25, 0.40, "x", "Huff factor at exponent 1.5 to 1.0", "DistW"),
        ("Aware-vs-unaware conversion lift", 2.0, 2.0, 2.0, "x", "Wyoming: 14% vs 7%", "AwareLiftX"),
        ("Local awareness today", 0.75, 0.75, 0.75, "share", "Team assumption after 25 years of local marketing", "AwareLocal"),
        ("Outer awareness today", 0.20, 0.30, 0.45, "share", "Team assumption: word of mouth only", "AwareOuter"),
        ("Awareness gain from the campaign", 0.10, 0.15, 0.25, "share", "Team assumption at one sixth of local per-capita spend", "AwareGain"),
        ("Regular patrons share", 0.75, 0.80, 0.85, "share", "NY State Fair 2015 exit survey", "RegShare"),
        ("Free event no-show rate", 0.40, 0.50, 0.60, "share", "VenueSight / Eventbrite", "NoShow"),
        ("Concert inputs", None),
        ("Spotify monthly listeners (act)", 7400000, 7400000, 7400000, "listeners", "Placeholder for the booked act; 7.4M Spotify x 1.35 x 0.93 = 9.3M combined, near the 10M tier", "SpotifyML"),
        ("Apple / Spotify listener ratio", 0.25, 0.35, 0.45, "x", "Sources sheet: apple_to_spotify_ratio", "AppleRatio"),
        ("Cross-platform overlap", 0.07, 0.07, 0.07, "share", "Team assumption", "Overlap"),
        ("US share of the act's audience", 0.40, 0.60, 0.80, "share", "Team assumption for a US act", "USShare"),
        ("US population", 340000000, 340000000, 340000000, "people", "Census estimate", "USPop"),
        ("Affinity multiplier", 0.5, 1.0, 2.5, "x", "1.0 national act, up to 4 regional", "Affinity"),
        ("Engagement multiplier", 0.6, 1.0, 1.5, "x", "Follower-to-listener ratio proxy; super listeners drive 50% of sales", "Engage"),
        ("Free-show conversion of catchment listeners", 0.04, 0.10, 0.18, "share", "5x the 2% paid working estimate", "ConvFree"),
        ("Show-up rate", 0.50, 0.70, 0.85, "share", "Team assumption", "ShowUp"),
        ("Regular share of concert fans", 0.10, 0.10, 0.10, "share", "Local penetration", "FanRegular"),
        ("Casual on-grounds share wanting seats", 0.40, 0.40, 0.40, "share", "Team assumption", "Casual"),
        ("Macro inputs", None),
        ("Gas price now ($/gal)", 4.49, 4.49, 4.49, "USD", "AAA Sept 24-25 2026", "GasNow"),
        ("Gas price a year ago ($/gal)", 3.60, 3.60, 3.60, "USD", "AAA/Empower: about 25% below today", "GasPrior"),
        ("Trip fuel elasticity (outer ring)", -0.30, -0.20, -0.10, "elasticity", "Above Hughes-Knittel-Sperling -0.03 to -0.08 for total gasoline", "TripElast"),
        ("Staycation lift (local ring)", 0.00, 0.01, 0.03, "share", "2009 recession fair evidence", "Stay"),
        ("Sentiment drag on visits", -0.01, -0.003, 0.0, "share", "UMich sentiment 48.1 vs 55.1; spending still +6% nominal", "SentDrag"),
        ("Real per-guest spend change", -0.06, -0.04, -0.01, "share", "Trade-down; MN State Fair 2025 survey comments", "PerCap"),
        ("Cost escalator for 2027 options", 0.046, 0.046, 0.046, "share", "UMich one-year inflation expectations 4.6%", "Escalator"),
    ]
    row = 2
    for item in inputs:
        if item[1] is None:
            section(ws, row, item[0]); row += 1; continue
        label, lo, base, hi, unit, src, nm = item
        ws.cell(row=row, column=1, value=label).font = Font(name=FONT, size=10)
        for j, v in enumerate((lo, base, hi), 2):
            c = ws.cell(row=row, column=j, value=v)
            is_formula = isinstance(v, str) and v.startswith("=")
            style(ws, c, "link" if is_formula else "input", "0.0%" if (unit == "share" and not is_formula) else ("#,##0.000" if unit in ("elasticity", "rate", "x", "exponent") else "#,##0.00" if unit == "USD" else "#,##0"))
            if is_formula: c.number_format = "0.0%" if unit == "share" else "#,##0"
        ws.cell(row=row, column=5, value=unit).font = Font(name=FONT, size=9, color=MUTED)
        ws.cell(row=row, column=6, value=src).font = Font(name=FONT, size=9, color=MUTED)
        name(wb, nm, "Inputs", f"C{row}"); name(wb, nm + "_lo", "Inputs", f"B{row}"); name(wb, nm + "_hi", "Inputs", f"D{row}")
        if nm in ("Base2026", "AnnivLift", "ReachAdd", "MacroLift", "OuterShare"): ws.cell(row=row, column=3).fill = KEY_FILL
        row += 1
    ws.freeze_panes = "A2"

    # ------------------------------------------------------------------ Baseline
    ws = wb.create_sheet("Baseline"); widths(ws, [30, 14, 14, 14, 14, 14, 14, 60])
    section(ws, 1, "Historical attendance and trend methods")
    header(ws, 2, ["Year", "x (year index)", "Attendance", "ln(Attendance)", "Holt level", "Holt trend", "Holt 1-step forecast", "Note"])
    for i, r in att.iterrows():
        rr = 3 + i
        ws.cell(row=rr, column=1, value=str(int(r["year"]))).font = Font(name=FONT, size=10)
        style(ws, ws.cell(row=rr, column=2, value=f"=A{rr}-2020"), "formula", "0")
        style(ws, ws.cell(row=rr, column=3, value=int(r["attendance"])), "input", "#,##0")
        style(ws, ws.cell(row=rr, column=4, value=f"=LN(C{rr})"), "formula", "0.0000")
        ws.cell(row=rr, column=8, value=r["note"]).font = Font(name=FONT, size=9, color=MUTED)
    # Holt recursion
    style(ws, ws.cell(row=3, column=5, value="=C3"), "formula", "#,##0"); style(ws, ws.cell(row=3, column=6, value="=C4-C3"), "formula", "#,##0")
    for rr in range(4, 8):
        style(ws, ws.cell(row=rr, column=7, value=f"=E{rr-1}+F{rr-1}"), "formula", "#,##0")
        style(ws, ws.cell(row=rr, column=5, value=f"=$B$11*C{rr}+(1-$B$11)*(E{rr-1}+F{rr-1})"), "formula", "#,##0")
        style(ws, ws.cell(row=rr, column=6, value=f"=$B$12*(E{rr}-E{rr-1})+(1-$B$12)*F{rr-1}"), "formula", "#,##0")
    ws.cell(row=8, column=1, value="2026").font = Font(name=FONT, size=10); style(ws, ws.cell(row=8, column=2, value="=A8-2020"), "formula", "0")
    style(ws, ws.cell(row=8, column=7, value="=E7+F7"), "formula", "#,##0"); note(ws, 8, "Holt 2026 = last level + last trend", 8)
    ws.cell(row=10, column=1, value="Holt parameters").font = SEC_FONT
    ws.cell(row=11, column=1, value="alpha (level smoothing)"); style(ws, ws.cell(row=11, column=2, value=R["baseline"]["holt_params"]["alpha"]), "input", "0.00")
    ws.cell(row=12, column=1, value="beta (trend smoothing)"); style(ws, ws.cell(row=12, column=2, value=R["baseline"]["holt_params"]["beta"]), "input", "0.00")
    note(ws, 11, "Grid-searched to minimize one-step squared error (baseline.py, holt_optimize). 0.9/0.9 means trust the newest data.", 3)
    section(ws, 14, "Five methods for 2026")
    header(ws, 15, ["Method", "Formula", "2026 forecast", "Weight", "Weighted", "Holdout miss on 2025", "", "Read"])
    methods = [
        ("A. OLS linear, all years", "=FORECAST(B8,C3:C7,B3:B7)", 0.15, "=FORECAST(B7,C3:C6,B3:B6)/C7-1", "Constant headcount growth. SLOPE = guests per year."),
        ("B. Log-linear, all years", "=EXP(FORECAST(B8,D3:D7,B3:B7))", 0.15, "=EXP(FORECAST(B7,D3:D6,B3:B6))/C7-1", "Constant percentage growth (CAGR)."),
        ("C. OLS linear, 2023 to 2025", "=FORECAST(B8,C5:C7,B5:B7)", 0.25, "=FORECAST(B7,C4:C6,B4:B6)/C7-1", "Post-pandemic regime only (holdout uses the three years before 2025)."),
        ("D. Log-linear, 2023 to 2025", "=EXP(FORECAST(B8,D5:D7,B5:B7))", 0.20, "n/a", "Recent CAGR."),
        ("E. Holt smoothing", "=G8", 0.25, "n/a", "Level plus trend, recent years weighted more."),
    ]
    for i, (lab, f, w, ho, rd) in enumerate(methods):
        rr = 16 + i
        ws.cell(row=rr, column=1, value=lab).font = Font(name=FONT, size=10)
        ws.cell(row=rr, column=2, value=f.replace("=", "", 1)).font = Font(name=FONT, size=8, color=MUTED)
        style(ws, ws.cell(row=rr, column=3, value=f), "formula", "#,##0")
        style(ws, ws.cell(row=rr, column=4, value=w), "input", "0%")
        style(ws, ws.cell(row=rr, column=5, value=f"=C{rr}*D{rr}"), "formula", "#,##0")
        c = ws.cell(row=rr, column=6, value=ho); style(ws, c, "formula" if ho.startswith("=") else "muted", "0.0%")
        ws.cell(row=rr, column=8, value=rd).font = Font(name=FONT, size=9, color=MUTED)
    ws.cell(row=22, column=1, value="Slope A (guests per year)"); style(ws, ws.cell(row=22, column=3, value="=SLOPE(C3:C7,B3:B7)"), "formula", "#,##0")
    ws.cell(row=23, column=1, value="CAGR B"); style(ws, ws.cell(row=23, column=3, value="=EXP(SLOPE(D3:D7,B3:B7))-1"), "formula", "0.0%")
    ws.cell(row=24, column=1, value="Combined baseline 2026").font = Font(name=FONT, bold=True, size=10)
    c = style(ws, ws.cell(row=24, column=3, value="=SUM(E16:E20)"), "formula", "#,##0", bold=True); c.fill = KEY_FILL
    ws.cell(row=25, column=1, value="Weights sum"); style(ws, ws.cell(row=25, column=3, value="=SUM(D16:D20)"), "formula", "0%")
    section(ws, 27, "80% prediction interval on the linear fit (five points is a short series)")
    ws.cell(row=28, column=1, value="Residual std error"); style(ws, ws.cell(row=28, column=3, value="=STEYX(C3:C7,B3:B7)"), "formula", "#,##0")
    ws.cell(row=29, column=1, value="SE of a new observation at x=6"); style(ws, ws.cell(row=29, column=3, value="=C28*SQRT(1+1/5+(B8-AVERAGE(B3:B7))^2/DEVSQ(B3:B7))"), "formula", "#,##0")
    ws.cell(row=30, column=1, value="t (80%, 3 df)"); style(ws, ws.cell(row=30, column=3, value="=TINV(0.2,3)"), "formula", "0.000")
    ws.cell(row=31, column=1, value="Interval low"); style(ws, ws.cell(row=31, column=3, value="=C16-C30*C29"), "formula", "#,##0")
    ws.cell(row=32, column=1, value="Interval high"); style(ws, ws.cell(row=32, column=3, value="=C16+C30*C29"), "formula", "#,##0")
    note(ws, 33, "Holdout: hide 2025, forecast it from 2021 to 2024, compare. All misses are negative, so growth is accelerating and recent-regime methods get more weight.")
    # chart data block: actual + combined baseline
    ws.cell(row=35, column=1, value="Chart data"); header(ws, 36, ["Year", "Attendance"])
    for i in range(5):
        ws.cell(row=37 + i, column=1, value=f"=A{3+i}"); style(ws, ws.cell(row=37 + i, column=2, value=f"=C{3+i}"), "formula", "#,##0")
    ws.cell(row=42, column=1, value="2026 base"); style(ws, ws.cell(row=42, column=2, value="=C24"), "formula", "#,##0")
    ch = LineChart(); ch.title = "Attendance 2021 to 2025 and 2026 organic baseline"; ch.style = 12; ch.legend = None
    ch.add_data(Reference(ws, min_col=2, min_row=36, max_row=42), titles_from_data=True); ch.set_categories(Reference(ws, min_col=1, min_row=37, max_row=42))
    ch.series[0].marker.symbol = "circle"; ch.series[0].graphicalProperties.line.solidFill = "1F5FBF"; ch.y_axis.scaling.min = 44000
    ch.width, ch.height = 16, 8; ws.add_chart(ch, "J2")

    # ------------------------------------------------------------------ Capture
    ws = wb.create_sheet("Capture"); widths(ws, [40, 16, 16, 16, 16, 60])
    section(ws, 1, "Capture rate = visits from a ring / population of the ring (2025)")
    header(ws, 2, ["Ring", "Population", "Share of visits", "Visits", "Capture rate", "Note"])
    ws.cell(row=3, column=1, value="Local metro (historically advertised)"); style(ws, ws.cell(row=3, column=2, value="=PopLocal"), "link", "#,##0")
    style(ws, ws.cell(row=3, column=3, value="=1-OuterShare-BeyondShare"), "formula", "0%"); style(ws, ws.cell(row=3, column=4, value="=Att2025*C3"), "formula", "#,##0"); style(ws, ws.cell(row=3, column=5, value="=D3/B3"), "formula", "0.00%")
    ws.cell(row=4, column=1, value="Neighboring metros (new reach)"); style(ws, ws.cell(row=4, column=2, value="=PopOuter"), "link", "#,##0")
    style(ws, ws.cell(row=4, column=3, value="=OuterShare"), "link", "0%"); style(ws, ws.cell(row=4, column=4, value="=Att2025*C4"), "formula", "#,##0"); style(ws, ws.cell(row=4, column=5, value="=D4/B4"), "formula", "0.00%")
    ws.cell(row=5, column=1, value="Beyond both rings"); style(ws, ws.cell(row=5, column=3, value="=BeyondShare"), "link", "0%"); style(ws, ws.cell(row=5, column=4, value="=Att2025*C5"), "formula", "#,##0")
    ws.cell(row=6, column=1, value="Outer / local capture ratio"); style(ws, ws.cell(row=6, column=5, value="=E4/E3"), "formula", "0.00")
    ws.cell(row=6, column=6, value="Compare with the Huff distance factor below: distance alone explains most of the gap").font = Font(name=FONT, size=9, color=MUTED)
    section(ws, 8, "Huff gravity check: propensity of an outer-ring resident relative to a local one = (d_outer/d_local)^(-exponent)")
    header(ws, 9, ["Exponent", "Distance factor"])
    for i, e in enumerate((1.0, 1.5, 2.0)):
        style(ws, ws.cell(row=10 + i, column=1, value=e), "input", "0.0"); style(ws, ws.cell(row=10 + i, column=2, value=f"=(DistOuter/DistLocal)^(-A{10+i})"), "formula", "0.00")
    section(ws, 14, "Calibrated model: aware residents convert at AwareLiftX times the unaware rate")
    rows = [("Unaware conversion, local", "=E3/(AwareLocal*AwareLiftX+(1-AwareLocal))", "0.00%", "Backed out of local capture"),
            ("Aware conversion, local", "=B15*AwareLiftX", "0.00%", ""),
            ("Implied distance factor, outer ring", "=E4/(AwareOuter*B16+(1-AwareOuter)*B15)", "0.00", "Backed out of outer capture"),
            ("Outer awareness after campaign", "=MIN(1,AwareOuter+AwareGain)", "0%", ""),
            ("Outer capture after campaign", "=(B18*B16+(1-B18)*B15)*B17", "0.00%", ""),
            ("Incremental outer visits (year one)", "=(B19-E4)*PopOuter", "#,##0", "One of the four reach estimates"),
            ("Outer ceiling capture at full awareness", "=B16*B17", "0.00%", ""),
            ("Outer ceiling visits", "=B21*PopOuter", "#,##0", "About 7,500 above today"),
            ("Local ceiling visits at full awareness", "=B16*PopLocal", "#,##0", "")]
    for i, (lab, f, fmt, nt) in enumerate(rows):
        rr = 15 + i; ws.cell(row=rr, column=1, value=lab).font = Font(name=FONT, size=10); style(ws, ws.cell(row=rr, column=2, value=f), "formula", fmt)
        ws.cell(row=rr, column=6, value=nt).font = Font(name=FONT, size=9, color=MUTED)
    section(ws, 26, "Five-year outer-ring path: awareness spreads by a Bass process (p = campaign at BassP_hi, q = word of mouth)")
    header(ws, 27, ["Year", "Awareness", "Outer capture", "Outer visits", "vs 2025"])
    for i, yr in enumerate(range(2026, 2031)):
        rr = 28 + i; ws.cell(row=rr, column=1, value=str(yr))
        prev = "AwareOuter" if i == 0 else f"B{rr-1}"
        style(ws, ws.cell(row=rr, column=2, value=f"=MIN(1,{prev}+(BassP_hi+BassQ*{prev})*(1-{prev}))"), "formula", "0%")
        style(ws, ws.cell(row=rr, column=3, value=f"=(B{rr}*$B$16+(1-B{rr})*$B$15)*$B$17"), "formula", "0.00%")
        style(ws, ws.cell(row=rr, column=4, value=f"=C{rr}*PopOuter"), "formula", "#,##0"); style(ws, ws.cell(row=rr, column=5, value=f"=D{rr}-$D$4"), "formula", "+#,##0;-#,##0")
    bar(ws, "H2", "Capture rate by ring (2025)", Reference(ws, min_col=1, min_row=3, max_row=4), Reference(ws, min_col=5, min_row=3, max_row=4), "share of ring population")
    bar(ws, "H20", "Outer-ring visits, five-year path", Reference(ws, min_col=1, min_row=28, max_row=32), Reference(ws, min_col=4, min_row=28, max_row=32), "visits")

    # ------------------------------------------------------------------ Reach
    ws = wb.create_sheet("Reach"); widths(ws, [44, 16, 16, 60])
    section(ws, 1, "Marketing arithmetic")
    rows = [("Per-capita intensity change if spread evenly", "=(1+SpendUp)/AudMult-1", "0%", "Money up 1.2x, audience up 2.4x"),
            ("Reading B local effect (elasticity x intensity)", "=AdElast*B2", "0.0%", "Spread evenly: local market loses frequency"),
            ("Reading B local visits", "=B3*Att2025", "#,##0", ""),
            ("Reading A local spillover (held budget)", "=MktLift", "0.0%", "Carried in the bridge"),
            ("Reading A local visits", "=B5*Att2025", "#,##0", "")]
    for i, (lab, f, fmt, nt) in enumerate(rows):
        rr = 2 + i; ws.cell(row=rr, column=1, value=lab); style(ws, ws.cell(row=rr, column=2, value=f), "formula", fmt); ws.cell(row=rr, column=4, value=nt).font = Font(name=FONT, size=9, color=MUTED)
    section(ws, 8, "Four estimates of the new-market lift (visits)")
    header(ws, 9, ["Method", "Low", "High", "How"])
    rows = [("1. Elasticity on whole spend increase", "=Att2025*SpendUp*AdElast", "=Att2025*SpendUp*AdElast_hi", "attendance x 20% x elasticity (0.12 short run, 0.24 long run)"),
            ("2. Awareness x conversion funnel", "=PopOuter*Aware_lo*AwareConv_lo", "=PopOuter*Aware*AwareConv", "700k x awareness x conversion"),
            ("3. Bass diffusion, year one", "=PopOuter*(Att2025/PopLocal)*DistW*BassP", "=PopOuter*(Att2025/PopLocal)*DistW*BassP_hi", "p x market potential; potential = 700k x local rate x distance weight"),
            ("4. Calibrated capture model", "=Capture!B20*0.6", "=Capture!B20*1.67", "Capture sheet, awareness gain 10 to 25 points")]
    for i, (lab, lo, hi, how) in enumerate(rows):
        rr = 10 + i; ws.cell(row=rr, column=1, value=lab); style(ws, ws.cell(row=rr, column=2, value=lo), "formula", "#,##0"); style(ws, ws.cell(row=rr, column=3, value=hi), "formula", "#,##0")
        ws.cell(row=rr, column=4, value=how).font = Font(name=FONT, size=9, color=MUTED)
    ws.cell(row=15, column=1, value="Carried in the bridge").font = Font(name=FONT, bold=True); c = style(ws, ws.cell(row=15, column=2, value="=ReachAdd"), "link", "#,##0", bold=True); c.fill = KEY_FILL
    ws.cell(row=15, column=4, value="The overlap of the four ranges. Change ReachAdd on Inputs to override.").font = Font(name=FONT, size=9, color=MUTED)
    ws.cell(row=17, column=1, value="Bass market potential (outer ring)"); style(ws, ws.cell(row=17, column=2, value="=PopOuter*(Att2025/PopLocal)*DistW"), "formula", "#,##0")
    header(ws, 18, ["Bass path (new adopters per year, p base)", "Year 1", "Year 2", "Year 3", "Year 4", "Year 5"])
    ws.cell(row=19, column=1, value="Cumulative"); ws.cell(row=20, column=1, value="New adopters")
    for j in range(5):
        col = 2 + j; cl = get_column_letter(col); prevc = "0" if j == 0 else f"{get_column_letter(col-1)}19"
        style(ws, ws.cell(row=20, column=col, value=f"=(BassP+BassQ*{prevc}/$B$17)*($B$17-{prevc})"), "formula", "#,##0")
        style(ws, ws.cell(row=19, column=col, value=f"={prevc}+{cl}20"), "formula", "#,##0")
    ws.column_dimensions["E"].width = 16; ws.column_dimensions["F"].width = 16
    bar(ws, "H2", "New-market lift, four methods (high end)", Reference(ws, min_col=1, min_row=10, max_row=13), Reference(ws, min_col=3, min_row=10, max_row=13), "visits")

    # ------------------------------------------------------------------ Anniversary
    ws = wb.create_sheet("Anniversary"); widths(ws, [44, 16, 16, 16, 60])
    section(ws, 1, "Top-down: milestone comparables (all-in lift above trend)")
    header(ws, 2, ["Comparable", "Prior", "Milestone", "Lift above trend", "Source"])
    ws.cell(row=3, column=1, value="Disneyland 60th (2015), trend 3.5%"); style(ws, ws.cell(row=3, column=2, value=16769000), "input", "#,##0"); style(ws, ws.cell(row=3, column=3, value=18278000), "input", "#,##0")
    style(ws, ws.cell(row=3, column=4, value="=C3/B3-1-0.035"), "formula", "0.0%"); ws.cell(row=3, column=5, value="TEA/AECOM Theme Index")
    ws.cell(row=4, column=1, value="Disneyland 50th (2005) vs top-50 parks"); style(ws, ws.cell(row=4, column=2, value=0.042), "input", "0.0%"); style(ws, ws.cell(row=4, column=3, value=0.085), "input", "0.0%")
    style(ws, ws.cell(row=4, column=4, value="=C4-B4"), "formula", "0.0%"); ws.cell(row=4, column=5, value="Amusement Business / ERA 2005")
    ws.cell(row=5, column=1, value="Musikfest 40th (2023) vs prior record"); style(ws, ws.cell(row=5, column=2, value=1240000), "input", "#,##0"); style(ws, ws.cell(row=5, column=3, value=1330000), "input", "#,##0")
    style(ws, ws.cell(row=5, column=4, value="=C5/B5-1"), "formula", "0.0%"); ws.cell(row=5, column=5, value="Morning Call, Aug 2023")
    ws.cell(row=6, column=1, value="Disneyland 2016 giveback"); style(ws, ws.cell(row=6, column=2, value=18278000), "input", "#,##0"); style(ws, ws.cell(row=6, column=3, value=17943000), "input", "#,##0")
    style(ws, ws.cell(row=6, column=4, value="=C6/B6-1"), "formula", "0.0%"); ws.cell(row=6, column=5, value="Year after the milestone")
    ws.cell(row=7, column=1, value="Pieces counted elsewhere (concert + marketing)"); style(ws, ws.cell(row=7, column=4, value="=ConcertAdd/Att2025+MktLift"), "formula", "0.0%")
    ws.cell(row=8, column=1, value="Net milestone lift carried").font = Font(name=FONT, bold=True); c = style(ws, ws.cell(row=8, column=4, value="=AnnivLift"), "link", "0.0%", bold=True); c.fill = KEY_FILL
    section(ws, 10, "Bottom-up: four channels")
    header(ws, 11, ["Channel", "Driver", "Visits", "% of 2025", "Logic"])
    rows = [("Lapsed visitors return", "=Att2025*(1-RegShare)*10*0.6", "=B12*0.03", "Non-regular visits x 10 years x 0.6 movers = pool; 3% reactivate"),
            ("Earned media", "=PopLocal-Att2025*0.7", "=B13*0.2*0.01", "Non-attending locals x 20% reached x 1% convert"),
            ("Special programming", "=Att2025", "=B14*0.01", "1% attractiveness bump"),
            ("Regulars add a visit", "=Att2025*RegShare*0.3", "=B15*0.02*5", "30% multi-visit regulars, +2% visits, 5 extra visit-days")]
    for i, (lab, drv, vis, lg) in enumerate(rows):
        rr = 12 + i; ws.cell(row=rr, column=1, value=lab); style(ws, ws.cell(row=rr, column=2, value=drv), "formula", "#,##0"); style(ws, ws.cell(row=rr, column=3, value=vis), "formula", "#,##0")
        style(ws, ws.cell(row=rr, column=4, value=f"=C{rr}/Att2025"), "formula", "0.0%"); ws.cell(row=rr, column=5, value=lg).font = Font(name=FONT, size=9, color=MUTED)
    ws.cell(row=16, column=1, value="Stacked"); style(ws, ws.cell(row=16, column=3, value="=SUM(C12:C15)"), "formula", "#,##0"); style(ws, ws.cell(row=16, column=4, value="=C16/Att2025"), "formula", "0.0%")
    ws.cell(row=17, column=1, value="Netted for overlap (45%)"); style(ws, ws.cell(row=17, column=3, value="=C16*0.55"), "formula", "#,##0"); style(ws, ws.cell(row=17, column=4, value="=C17/Att2025"), "formula", "0.0%")
    ws.cell(row=17, column=5, value="A lapsed visitor brought back by a news story is channels 1 and 2 at once").font = Font(name=FONT, size=9, color=MUTED)
    bar(ws, "G2", "Milestone lift above trend, comparables", Reference(ws, min_col=1, min_row=3, max_row=5), Reference(ws, min_col=4, min_row=3, max_row=5), "share")

    # ------------------------------------------------------------------ Concert
    ws = wb.create_sheet("Concert"); widths(ws, [44, 16, 16, 60])
    section(ws, 1, "Listener funnel for the booked act")
    rows = [("Combined monthly listeners", "=SpotifyML*(1+AppleRatio)*(1-Overlap)", "#,##0", "Spotify public count + Apple estimate, less overlap"),
            ("Effective catchment population", "=PopLocal+PopOuter*DistW", "#,##0", "Outer ring weighted for distance"),
            ("Catchment listeners", "=B2*USShare*(B3/USPop)*Affinity", "#,##0", "Listeners spread with population, scaled by affinity"),
            ("Concert-motivated demand", "=B4*ConvFree*Engage", "#,##0", "Free-show conversion x engagement"),
            ("Regulars (shift weekends)", "=B5*FanRegular*ShowUp", "#,##0", ""),
            ("Incremental guests", "=B5*(1-FanRegular)*ShowUp", "#,##0", "New to the season"),
            ("Casual on-grounds seat demand", "=Weekends!C9*Casual", "#,##0", "Weekend 4 Saturday guests before the concert"),
            ("Seat demand", "=B8+B6+B7", "#,##0", ""),
            ("Overflow above capacity", "=MAX(0,B9-ConcertCap)", "#,##0", ""),
            ("Passes to issue if reservation-based", "=ConcertCap/(1-NoShow)", "#,##0", "Free events lose 40 to 60% of registrants")]
    for i, (lab, f, fmt, nt) in enumerate(rows):
        rr = 2 + i; ws.cell(row=rr, column=1, value=lab); style(ws, ws.cell(row=rr, column=2, value=f), "formula", fmt); ws.cell(row=rr, column=4, value=nt).font = Font(name=FONT, size=9, color=MUTED)
    ws.cell(row=7, column=1).font = Font(name=FONT, bold=True); ws.cell(row=7, column=2).fill = KEY_FILL
    section(ws, 13, "Act-size tiers (combined listeners)")
    header(ws, 14, ["Combined listeners", "Catchment listeners", "Incremental guests", "Shifted regulars"])
    for i, t in enumerate((1e6, 5e6, 10e6, 20e6, 40e6)):
        rr = 15 + i; style(ws, ws.cell(row=rr, column=1, value=t), "input", "#,##0")
        style(ws, ws.cell(row=rr, column=2, value=f"=A{rr}*USShare*($B$3/USPop)*Affinity"), "formula", "#,##0")
        style(ws, ws.cell(row=rr, column=3, value=f"=B{rr}*ConvFree*Engage*(1-FanRegular)*ShowUp"), "formula", "#,##0")
        style(ws, ws.cell(row=rr, column=4, value=f"=B{rr}*ConvFree*Engage*FanRegular*ShowUp"), "formula", "#,##0")
    ws.cell(row=21, column=1, value="Listeners needed for 2,000 incremental"); style(ws, ws.cell(row=21, column=2, value="=2000/C17*A17"), "formula", "#,##0")
    ws.cell(row=22, column=1, value="Listeners at which WK4 Sat passes WK3 Sat"); style(ws, ws.cell(row=22, column=2, value="=(Weekends!C8-Weekends!C9+ConcertShift)/(C17+D17)*A17"), "formula", "#,##0")
    ws.cell(row=22, column=4, value="Gap between the two Saturdays divided by guests per listener").font = Font(name=FONT, size=9, color=MUTED)
    bar(ws, "F2", "Incremental concert guests by act size", Reference(ws, min_col=1, min_row=15, max_row=19), Reference(ws, min_col=3, min_row=15, max_row=19), "guests")

    # ------------------------------------------------------------------ Macro
    ws = wb.create_sheet("Macro"); widths(ws, [44, 16, 16, 70])
    section(ws, 1, "September 2026 readings")
    header(ws, 2, ["Indicator", "Now", "Year ago", "Source"])
    rows = [("UMich consumer sentiment", 48.1, 55.1, "University of Michigan Surveys of Consumers, Sept 2026 final vs Sept 2025"),
            ("One-year inflation expectations", 0.046, 0.036, "UMich Sept 2026 4.6%; NY Fed Aug 2026 3.6%"),
            ("Gasoline, national average $/gal", 4.49, 3.60, "AAA Sept 24-25 2026; about 25% above a year ago (Empower, AAA)"),
            ("Unemployment rate", 0.041, 0.042, "BLS via Sept 2026 coverage: 4.1%, payrolls +162k"),
            ("Retail and food-service sales, y/y nominal", 0.06, None, "Census, August 2026")]
    for i, (lab, now, ago, src) in enumerate(rows):
        rr = 3 + i; ws.cell(row=rr, column=1, value=lab); fmt = "0.0%" if isinstance(now, float) and now < 1 else "#,##0.00"
        style(ws, ws.cell(row=rr, column=2, value=now), "input", fmt)
        if ago is not None: style(ws, ws.cell(row=rr, column=3, value=ago), "input", fmt)
        ws.cell(row=rr, column=4, value=src).font = Font(name=FONT, size=9, color=MUTED)
    section(ws, 9, "Three mechanisms")
    rows = [("Gas price change y/y", "=GasNow/GasPrior-1", "0.0%", "Input to the fuel-cost mechanism"),
            ("Extra fuel cost per outer-ring car trip (90 mi, 25 mpg)", "=90/25*GasNow-90/25*GasPrior", "$0.00", "A few dollars; the effect is on the margin, not the average"),
            ("Outer-ring visit effect (trip elasticity x gas change)", "=TripElast*B10", "0.0%", "Hughes-Knittel-Sperling: total gasoline demand only -0.03 to -0.08; day trips more elastic"),
            ("Outer-ring visits effect", "=B12*Att2025*OuterShare", "#,##0", ""),
            ("Local staycation lift", "=Stay", "0.0%", "2009: fairs boomed as families skipped out-of-town trips"),
            ("Local visits effect", "=B14*Att2025*(1-OuterShare-BeyondShare)", "#,##0", ""),
            ("Sentiment drag on visits", "=SentDrag", "0.0%", "Sentiment near record lows, yet nominal spending +6%"),
            ("Sentiment visits effect", "=B16*Att2025", "#,##0", ""),
            ("Net attendance effect (visits)", "=B13+B15+B17", "#,##0", ""),
            ("Real per-guest spend change", "=PerCap", "0.0%", "Feeds Part 2 sales per booth and Part 3 revenue per guest"),
            ("Cost escalator for 2027 options", "=Escalator", "0.0%", "Feeds Part 3 labor cost")]
    for i, (lab, f, fmt, nt) in enumerate(rows):
        rr = 10 + i; ws.cell(row=rr, column=1, value=lab); style(ws, ws.cell(row=rr, column=2, value=f), "formula", fmt); ws.cell(row=rr, column=4, value=nt).font = Font(name=FONT, size=9, color=MUTED)
    ws.cell(row=23, column=1, value="Net macro multiplier on attendance").font = Font(name=FONT, bold=True)
    ws.cell(row=24, column=1, value="(local share x staycation + outer share x fuel effect + sentiment)")
    c = style(ws, ws.cell(row=24, column=2, value="=(1-OuterShare-BeyondShare)*Stay+OuterShare*B12+SentDrag"), "formula", "0.00%", bold=True); c.fill = KEY_FILL
    ws.cell(row=24, column=4, value="Linked into Inputs!MacroLift and the bridge").font = Font(name=FONT, size=9, color=MUTED)

    # ------------------------------------------------------------------ Weekends (built before Bridge so references resolve either way)
    ws = wb.create_sheet("Weekends"); widths(ws, [30, 14, 14, 14, 14, 50, 22, 14])
    section(ws, 1, "Weekend split anchored on the expected season total (Monte Carlo P50 on the Bridge sheet; switch B1 to Bridge!B10 for the fair-weather deterministic total)")
    ws.cell(row=1, column=7, value="Season total used"); c = style(ws, ws.cell(row=1, column=8, value="=Bridge!B22"), "link", "#,##0"); c.fill = KEY_FILL
    ws.cell(row=2, column=7, value="Non-concert total"); style(ws, ws.cell(row=2, column=8, value="=H1-ConcertAdd*Weather"), "formula", "#,##0")
    header(ws, 2, ["Weekend", "Share", "Weekend total", "Saturday", "Sunday", "Note"])
    for i in range(4):
        rr = 3 + i; ws.cell(row=rr, column=1, value=f"Weekend {i+1}"); style(ws, ws.cell(row=rr, column=2, value=f"=WkS{i+1}"), "link", "0%")
        base_expr = f"=$H$2*B{rr}"
        if i == 3: base_expr = f"=$H$2*B{rr}+ConcertAdd*Weather+ConcertShift"
        else: base_expr += f"-ConcertShift*{[0.2,0.3,0.5][i]}"
        style(ws, ws.cell(row=rr, column=3, value=base_expr), "formula", "#,##0")
        if i == 3:
            style(ws, ws.cell(row=rr, column=4, value=f"=(C{rr}-ConcertAdd*Weather-ConcertShift)*SatShare+ConcertAdd*Weather+ConcertShift"), "formula", "#,##0")
        else:
            style(ws, ws.cell(row=rr, column=4, value=f"=C{rr}*SatShare"), "formula", "#,##0")
        style(ws, ws.cell(row=rr, column=5, value=f"=C{rr}-D{rr}"), "formula", "#,##0")
    ws.cell(row=6, column=6, value="Concert guests and 100 shifted regulars land on Weekend 4 Saturday").font = Font(name=FONT, size=9, color=MUTED)
    ws.cell(row=8, column=1, value="Weekend 3 Saturday (design day)").font = Font(name=FONT, bold=True); c = style(ws, ws.cell(row=8, column=3, value="=D5"), "formula", "#,##0", bold=True); c.fill = KEY_FILL
    ws.cell(row=9, column=1, value="Weekend 4 Saturday before the concert"); style(ws, ws.cell(row=9, column=3, value="=(C6-ConcertAdd*Weather-ConcertShift)*SatShare"), "formula", "#,##0")
    ws.cell(row=10, column=1, value="Weekend 4 Saturday with the concert"); style(ws, ws.cell(row=10, column=3, value="=D6"), "formula", "#,##0")
    ws.cell(row=11, column=1, value="Peak day"); style(ws, ws.cell(row=11, column=3, value='=IF(D6>D5,"Weekend 4 Saturday","Weekend 3 Saturday")'), "formula")
    section(ws, 13, "Monte Carlo reference (simulate.py, 100,000 draws), pasted values")
    header(ws, 14, ["Weekend", "P10", "P50", "P90"])
    for i in range(4):
        rr = 15 + i; ws.cell(row=rr, column=1, value=f"Weekend {i+1}")
        for j in range(3): style(ws, ws.cell(row=rr, column=2 + j, value=round(R["simulation"]["weekends_p10_p50_p90"][i][j], -2)), "input", "#,##0")
    ws.cell(row=19, column=1, value="Weekend 3 Saturday"); ws.cell(row=20, column=1, value="Weekend 4 Saturday")
    for j in range(3):
        style(ws, ws.cell(row=19, column=2 + j, value=round(R["simulation"]["wk3_sat_p10_p50_p90"][j], -2)), "input", "#,##0")
        style(ws, ws.cell(row=20, column=2 + j, value=round(R["simulation"]["wk4_sat_p10_p50_p90"][j], -2)), "input", "#,##0")
    ws.cell(row=21, column=1, value="P(Weekend 4 Saturday is the peak day)"); style(ws, ws.cell(row=21, column=2, value=R["simulation"]["p_wk4sat_over_wk3sat"]), "input", "0%")
    ch = BarChart(); ch.type = "col"; ch.grouping = "clustered"; ch.title = "Weekend totals, Saturday and Sunday"; ch.style = 10
    ch.add_data(Reference(ws, min_col=4, max_col=5, min_row=2, max_row=6), titles_from_data=True); ch.set_categories(Reference(ws, min_col=1, min_row=3, max_row=6))
    ch.series[0].graphicalProperties.solidFill = "1F5FBF"; ch.series[1].graphicalProperties.solidFill = "F2B632"; ch.width, ch.height = 16, 8; ws.add_chart(ch, "J2")

    # ------------------------------------------------------------------ Bridge
    ws = wb.create_sheet("Bridge", 9); widths(ws, [34, 16, 16, 16, 60])
    section(ws, 1, "Driver bridge: percentages first, headcounts after, weather last")
    header(ws, 2, ["Step", "Cumulative", "Step size", "", "How"])
    steps = [("2025 actual", "=Att2025", "Case"),
             ("+ organic trend", "=Base2026", "Baseline sheet, combined"),
             ("+ local marketing spillover", "=B4*(1+MktLift)", "Reading A"),
             ("+ 25th anniversary", "=B5*(1+AnnivLift)", "Net of marketing and concert"),
             ("+ macro backdrop", "=B6*(1+MacroLift)", "Macro sheet net multiplier"),
             ("+ reach expansion", "=B7+ReachAdd", "Reach sheet overlap"),
             ("+ concert", "=B8+ConcertAdd", "Concert sheet, booked-act tier"),
             ("Fair-weather total", "=B9", ""),
             ("x weather (base)", "=B10*Weather", "Weather multiplier mode 1.00; downside skewed")]
    for i, (lab, f, how) in enumerate(steps):
        rr = 3 + i; ws.cell(row=rr, column=1, value=lab).font = Font(name=FONT, size=10, bold=(i in (7, 8))); style(ws, ws.cell(row=rr, column=2, value=f), "formula", "#,##0", bold=(i in (7, 8)))
        if 0 < i < 7: style(ws, ws.cell(row=rr, column=3, value=f"=B{rr}-B{rr-1}"), "formula", "+#,##0;-#,##0")
        ws.cell(row=rr, column=5, value=how).font = Font(name=FONT, size=9, color=MUTED)
    ws.cell(row=10, column=2).fill = KEY_FILL
    section(ws, 13, "Deterministic scenarios (every driver at its low / base / high)")
    header(ws, 14, ["Scenario", "Total", "vs 2025", "", "Note"])
    for i, (lab, sfx) in enumerate((("Low (all drivers low)", "_lo"), ("Base", ""), ("High (all drivers high)", "_hi"))):
        rr = 15 + i; ws.cell(row=rr, column=1, value=lab)
        f = f"=(Base2026{sfx}*(1+MktLift{sfx})*(1+AnnivLift{sfx})*(1+MacroLift{sfx})+ReachAdd{sfx}+ConcertAdd{sfx})*Weather{sfx}"
        style(ws, ws.cell(row=rr, column=2, value=f), "formula", "#,##0"); style(ws, ws.cell(row=rr, column=3, value=f"=B{rr}/Att2025-1"), "formula", "+0%;-0%")
    ws.cell(row=15, column=5, value="All-adverse: rarely happens together, which is why the Monte Carlo band below is narrower").font = Font(name=FONT, size=9, color=MUTED)
    section(ws, 19, "Monte Carlo (simulate.py, 100,000 draws, triangular and PERT distributions), pasted values")
    header(ws, 20, ["Percentile", "Total", "vs 2025"])
    for i, (lab, v) in enumerate(zip(("P10 (low)", "P50 (expected)", "P90 (high)"), R["simulation"]["total_p10_p50_p90"])):
        rr = 21 + i; ws.cell(row=rr, column=1, value=lab); style(ws, ws.cell(row=rr, column=2, value=round(v, -2)), "input", "#,##0"); style(ws, ws.cell(row=rr, column=3, value=f"=B{rr}/Att2025-1"), "formula", "+0%;-0%")
    ws.cell(row=22, column=2).fill = KEY_FILL
    ws.cell(row=24, column=1, value="P(total > 60,000)"); style(ws, ws.cell(row=24, column=2, value=R["simulation"]["p_over_60k"]), "input", "0%")
    bar(ws, "G2", "Driver bridge, 2025 to 2026 (cumulative)", Reference(ws, min_col=1, min_row=3, max_row=11), Reference(ws, min_col=2, min_row=3, max_row=11), "visits", w=18)

    # ------------------------------------------------------------------ Sensitivity
    ws = wb.create_sheet("Sensitivity", 10); widths(ws, [30, 16, 16, 16, 16, 50])
    section(ws, 1, "Tornado: each driver at its low and high with the others at base (fair-weather total as the anchor)")
    header(ws, 2, ["Driver", "Total at low", "Total at high", "Delta low", "Delta high", "Swing"])
    drivers = [("Baseline trend", "Base2026"), ("Local marketing", "MktLift"), ("Anniversary", "AnnivLift"), ("Macro", "MacroLift"), ("Reach", "ReachAdd"), ("Concert", "ConcertAdd"), ("Weather", "Weather")]
    def total_expr(over, sfx):
        parts = {"Base2026": "Base2026", "MktLift": "MktLift", "AnnivLift": "AnnivLift", "MacroLift": "MacroLift", "ReachAdd": "ReachAdd", "ConcertAdd": "ConcertAdd", "Weather": "Weather"}
        parts[over] = over + sfx
        return f"=({parts['Base2026']}*(1+{parts['MktLift']})*(1+{parts['AnnivLift']})*(1+{parts['MacroLift']})+{parts['ReachAdd']}+{parts['ConcertAdd']})*{parts['Weather']}"
    for i, (lab, nm) in enumerate(drivers):
        rr = 3 + i; ws.cell(row=rr, column=1, value=lab)
        style(ws, ws.cell(row=rr, column=2, value=total_expr(nm, "_lo")), "formula", "#,##0"); style(ws, ws.cell(row=rr, column=3, value=total_expr(nm, "_hi")), "formula", "#,##0")
        style(ws, ws.cell(row=rr, column=4, value=f"=B{rr}-Bridge!$B$10"), "formula", "+#,##0;-#,##0"); style(ws, ws.cell(row=rr, column=5, value=f"=C{rr}-Bridge!$B$10"), "formula", "+#,##0;-#,##0")
        style(ws, ws.cell(row=rr, column=6, value=f"=ABS(C{rr}-B{rr})"), "formula", "#,##0")
    note(ws, 11, "Full-range swing (low to high). simulate.py reports the 10th-to-90th percentile version; the ranking is the same: weather, then reach and anniversary.")
    ch = BarChart(); ch.type = "bar"; ch.grouping = "clustered"; ch.title = "Sensitivity: delta from fair-weather total"; ch.style = 10
    ch.add_data(Reference(ws, min_col=4, max_col=5, min_row=2, max_row=9), titles_from_data=True); ch.set_categories(Reference(ws, min_col=1, min_row=3, max_row=9))
    ch.series[0].graphicalProperties.solidFill = "F2B632"; ch.series[1].graphicalProperties.solidFill = "1F5FBF"; ch.width, ch.height = 16, 9; ws.add_chart(ch, "H2")

    # ------------------------------------------------------------------ Part 2
    ws = wb.create_sheet("Part2"); widths(ws, [36, 14, 14, 14, 14, 14, 14, 50])
    section(ws, 1, "Category parameters (case midpoints in blue; service, fees and floors are team assumptions)")
    header(ws, 2, ["Category", "Spend $", "Participation", "Tx per guest", "Units", "Season fee $", "Revenue share", "Organizer cost $"])
    header(ws, 3, ["", "Vendor floor $", "Service pts", "Min per tx", "Guests/hr per booth", "Peak-hour factor", "Wait target (min)", ""])
    cats = [("F&B", 21.5, 0.775, 1.6, 1, 2000, 0.12, 900, 18000, 2, 2.5, "=60/D6*C6", 0.18, 10),
            ("Merch", 16.0, 0.325, 1.1, 1, 900, 0.10, 500, 5000, 1, 3.0, "=60/D7*C7", 0.15, 5),
            ("A&E", 8.5, 0.40, 1.3, 2, 700, 0.08, 600, 4000, None, None, 60, 0.15, 15)]
    for i, c in enumerate(cats):
        r1, r2 = 4 + 2 * i, 5 + 2 * i
        ws.cell(row=r1, column=1, value=c[0]).font = Font(name=FONT, bold=True)
        for j, v in enumerate(c[1:8], 2): style(ws, ws.cell(row=r1, column=j, value=v), "input", "0.0%" if j in (3, 7) else "#,##0.00" if j == 2 else "#,##0")
        ws.cell(row=r2, column=1, value="").font = Font(name=FONT)
        style(ws, ws.cell(row=r2, column=2, value=c[8]), "input", "#,##0")
        if c[9] is not None:
            style(ws, ws.cell(row=r2, column=3, value=c[9]), "input", "0"); style(ws, ws.cell(row=r2, column=4, value=c[10]), "input", "0.0")
            style(ws, ws.cell(row=r2, column=5, value=f"=60/D{r2}*C{r2}"), "formula", "0.0")
        else:
            ws.cell(row=r2, column=3, value="attraction"); style(ws, ws.cell(row=r2, column=5, value=c[11]), "input", "0.0")
        style(ws, ws.cell(row=r2, column=6, value=c[12]), "input", "0%"); style(ws, ws.cell(row=r2, column=7, value=c[13]), "input", "0")
    ws.cell(row=10, column=1, value="Season attendance (P50)"); c = style(ws, ws.cell(row=10, column=2, value="=Bridge!B22"), "link", "#,##0"); c.fill = KEY_FILL
    ws.cell(row=11, column=1, value="Design day (Weekend 3 Saturday)"); c = style(ws, ws.cell(row=11, column=2, value="=Weekends!C8"), "link", "#,##0"); c.fill = KEY_FILL
    ws.cell(row=12, column=1, value="Real per-guest spend multiplier"); style(ws, ws.cell(row=12, column=2, value="=1+PerCap"), "link", "0.00")
    section(ws, 14, "Design-day demand and Erlang C queue floors")
    header(ws, 15, ["Category", "Daily transactions", "Peak-hour arrivals (lambda)", "Service rate per booth (mu)", "Offered load a = lambda/mu", "Queue floor (booths)", "Wait at floor (min)", "Season sales potential $"])
    for i, cat in enumerate(("F&B", "Merch", "A&E")):
        rr = 16 + i; r1, r2 = 4 + 2 * i, 5 + 2 * i
        ws.cell(row=rr, column=1, value=cat)
        style(ws, ws.cell(row=rr, column=2, value=f"=$B$11*C{r1}*D{r1}"), "formula", "#,##0")
        style(ws, ws.cell(row=rr, column=3, value=f"=B{rr}*F{r2}"), "formula", "#,##0")
        style(ws, ws.cell(row=rr, column=4, value=f"=E{r2}"), "formula", "0.0")
        style(ws, ws.cell(row=rr, column=5, value=f"=C{rr}/D{rr}"), "formula", "0.0")
        style(ws, ws.cell(row=rr, column=6, value=f"=MIN(ErlangC!{['B','E','H'][i]}$3:{['B','E','H'][i]}$102)"), "link", "0")
        style(ws, ws.cell(row=rr, column=7, value=f"=INDEX(ErlangC!{['C','F','I'][i]}$3:{['C','F','I'][i]}$102,MATCH(F{rr},ErlangC!$A$3:$A$102,0))"), "link", "0.0")
        style(ws, ws.cell(row=rr, column=8, value=f"=$B$10*C{r1}*D{r1}*B{r1}*$B$12"), "formula", "$#,##0")
    section(ws, 20, "Candidate mixes (capture = 1 - EXP(-n / n0), n0 set so capture = 85% at the queue floor)")
    header(ws, 21, ["Mix", "F&B booths", "Merch booths", "A&E booths", "Units used", "Organizer contribution $", "", "Check"])
    mixes = [("Profit max (grid search, booths.py)", R["part2"]["mixes"]["profit_max"]["rows"]["F&B"]["booths"], R["part2"]["mixes"]["profit_max"]["rows"]["Merch"]["booths"], R["part2"]["mixes"]["profit_max"]["rows"]["A&E"]["booths"]),
             ("Balanced (recommended)", 60, 28, 16), ("Variety", 56, 32, 16)]
    for i, (lab, F, M, A) in enumerate(mixes):
        rr = 22 + i; ws.cell(row=rr, column=1, value=lab).font = Font(name=FONT, bold=(i == 1))
        for j, v in enumerate((F, M, A), 2): style(ws, ws.cell(row=rr, column=j, value=v), "input", "0")
        style(ws, ws.cell(row=rr, column=5, value=f"=B{rr}+C{rr}+2*D{rr}"), "formula", "0")
        style(ws, ws.cell(row=rr, column=6, value=f"=F{29+4*i}+F{30+4*i}+F{31+4*i}"), "formula", "$#,##0", bold=(i == 1))
        style(ws, ws.cell(row=rr, column=8, value=f'=IF(E{rr}<=120,"fits 120 units","over budget")'), "formula")
    ws.cell(row=23, column=6).fill = KEY_FILL
    section(ws, 27, "Mix detail")
    header(ws, 28, ["Mix / category", "Booths", "Capture", "Season sales $", "Sales per booth $", "Contribution $", "Peak wait (min)", "Above vendor floor?"])
    for i in range(3):
        for k, cat in enumerate(("F&B", "Merch", "A&E")):
            rr = 29 + 4 * i + k; mixrow = 22 + i; floorrow = 16 + k; r1, r2 = 4 + 2 * k, 5 + 2 * k; col = ["B", "C", "D"][k]
            ws.cell(row=rr, column=1, value=f"{mixes[i][0].split(' (')[0]} / {cat}")
            style(ws, ws.cell(row=rr, column=2, value=f"={col}{mixrow}"), "formula", "0")
            style(ws, ws.cell(row=rr, column=3, value=f"=1-EXP(-B{rr}/(F{floorrow}/-LN(1-0.85)))"), "formula", "0%")
            style(ws, ws.cell(row=rr, column=4, value=f"=H{floorrow}*C{rr}"), "formula", "$#,##0")
            style(ws, ws.cell(row=rr, column=5, value=f"=D{rr}/B{rr}"), "formula", "$#,##0")
            style(ws, ws.cell(row=rr, column=6, value=f"=F{r1}*B{rr}+G{r1}*D{rr}-H{r1}*B{rr}"), "formula", "$#,##0")
            ec = ["C", "F", "I"][k]
            style(ws, ws.cell(row=rr, column=7, value=f"=IFERROR(INDEX(ErlangC!{ec}$3:{ec}$102,MATCH(B{rr},ErlangC!$A$3:$A$102,0)),0)"), "link", "0.0")
            style(ws, ws.cell(row=rr, column=8, value=f'=IF(E{rr}>=B{r2},"yes","NO")'), "formula")
    bar(ws, "J2", "Recommended booth mix", Reference(ws, min_col=2, max_col=4, min_row=21, max_row=21), Reference(ws, min_col=2, max_col=4, min_row=23, max_row=23), "booths")

    # ErlangC helper sheet: wait time by number of servers for each category
    we = wb.create_sheet("ErlangC"); widths(we, [10, 12, 12, 12, 12, 12, 12, 12, 12, 12])
    section(we, 1, "Erlang C (M/M/c): mean queue wait in minutes for c booths. Helper for Part2. Column pairs: floor flag, wait.")
    header(we, 2, ["c (booths)", "F&B floor?", "F&B wait", "F&B P(wait)", "Merch floor?", "Merch wait", "Merch P(wait)", "A&E floor?", "A&E wait", "A&E P(wait)"])
    # term table for sum_{k<c} a^k/k! : use K helper columns L..
    we.cell(row=2, column=12, value="k").font = HDR_FONT; we.cell(row=2, column=12).fill = HDR_FILL
    for k in range(0, 100): we.cell(row=3 + k, column=12, value=k)
    for i, (cat, lam_ref) in enumerate((("F&B", "Part2!$C$16"), ("Merch", "Part2!$C$17"), ("A&E", "Part2!$C$18"))):
        mu_ref = ["Part2!$D$16", "Part2!$D$17", "Part2!$D$18"][i]; tgt = ["Part2!$G$5", "Part2!$G$7", "Part2!$G$9"][i]
        flag_c, wait_c, pw_c = 2 + 3 * i, 3 + 3 * i, 4 + 3 * i
        for c in range(1, 101):
            rr = 2 + c; we.cell(row=rr, column=1, value=c)
            a = f"({lam_ref}/{mu_ref})"
            s = f"SUMPRODUCT(($L$3:$L$102<A{rr})*({a}^$L$3:$L$102/FACT($L$3:$L$102)))"
            last = f"({a}^A{rr}/FACT(A{rr})*A{rr}/(A{rr}-{a}))"
            pw = f"=IF(A{rr}<={a},1,{last}/({s}+{last}))"
            we.cell(row=rr, column=pw_c, value=pw).number_format = "0.000"
            we.cell(row=rr, column=wait_c, value=f"=IF(A{rr}<={a},9999,{get_column_letter(pw_c)}{rr}/(A{rr}*{mu_ref}-{lam_ref})*60)").number_format = "0.0"
            we.cell(row=rr, column=flag_c, value=f"=IF({get_column_letter(wait_c)}{rr}<={tgt},A{rr},9999)").number_format = "0"
    note(we, 104, "P(wait) = [a^c/c! * c/(c-a)] / [sum_{k<c} a^k/k! + a^c/c! * c/(c-a)];  Wq = P(wait) / (c*mu - lambda). Floor = smallest c with Wq under the target.")

    # ------------------------------------------------------------------ Part 3
    ws = wb.create_sheet("Part3"); widths(ws, [40, 14, 14, 14, 14, 14, 14, 14, 14, 50])
    section(ws, 1, "Organizer revenue per guest")
    rows = [("Admission $", 12.0, "input"), ("Parking $ per guest", 1.5, "input"),
            ("Revenue share on spend $", "=(Part2!C4*Part2!D4*Part2!B4*Part2!G4+Part2!C6*Part2!D6*Part2!B6*Part2!G6+Part2!C8*Part2!D8*Part2!B8*Part2!G8)*Part2!B12", "formula"),
            ("Revenue per guest $", "=B2+B3+B4", "formula")]
    for i, (lab, v, kind) in enumerate(rows):
        rr = 2 + i; ws.cell(row=rr, column=1, value=lab); style(ws, ws.cell(row=rr, column=2, value=v), kind, "$#,##0.00", bold=(i == 3))
    section(ws, 7, "Cost inputs (team assumptions; replace with the festival's actuals)")
    costs = [("Staff on site", 50, "0"), ("Wage $/hr", 22.0, "$0.00"), ("Cost escalator (2027)", "=Escalator", "0.0%"), ("Lighting and security per extended day $", 2500, "$#,##0"),
             ("Operations per added day $", 8000, "$#,##0"), ("Extra site week $", 20000, "$#,##0"), ("Launch marketing, Fridays $", 10000, "$#,##0"), ("Launch marketing, fifth weekend $", 8000, "$#,##0")]
    for i, (lab, v, fmt) in enumerate(costs):
        rr = 8 + i; ws.cell(row=rr, column=1, value=lab); style(ws, ws.cell(row=rr, column=2, value=v), "link" if isinstance(v, str) else "input", fmt)
    section(ws, 17, "Demand assumptions")
    dem = [("Season attendance", "=Part2!B10", "#,##0"), ("Average Saturday", "=B18*SatShare/4", "#,##0"), ("Average weekend", "=B18/4", "#,##0"), ("Average demand per operating hour", "=B18/80", "#,##0"),
           ("A: share of arrivals in the two added hours", 0.07, "0%"), ("A: redistribution (pulled from open hours)", 0.60, "0%"),
           ("B: Friday as a share of Saturday", 0.45, "0%"), ("B: redistribution from the weekend", 0.35, "0%"),
           ("C: fifth weekend vs average weekend", 0.65, "0%"), ("C: redistribution", 0.40, "0%")]
    for i, (lab, v, fmt) in enumerate(dem):
        rr = 18 + i; ws.cell(row=rr, column=1, value=lab); style(ws, ws.cell(row=rr, column=2, value=v), "formula" if isinstance(v, str) else "input", fmt)
    section(ws, 29, "Option economics")
    header(ws, 30, ["Option", "Hours added", "Gross demand", "Net new", "Net new %", "Revenue $", "Cost $", "Net $", "Capacity factor", "Peak relief (shifted visits)"])
    opts = [("A. Extend to 12 hours", 16, "=B18*B22", "=C31*(1-B23)", "=B8*B9*(1+B10)*2*8+B11*8"),
            ("B. Add Fridays", 40, "=4*B24*B19", "=C32*(1-B25)", "=B8*B9*(1+B10)*11*4+B12*4+B14"),
            ("C. Fifth weekend", 20, "=B26*B20", "=C33*(1-B27)", "=B8*B9*(1+B10)*11*2+B12*2+B13+B15")]
    for i, (lab, hrs, g, n, cst) in enumerate(opts):
        rr = 31 + i; ws.cell(row=rr, column=1, value=lab).font = Font(name=FONT, bold=(i == 1))
        style(ws, ws.cell(row=rr, column=2, value=hrs), "input", "0"); style(ws, ws.cell(row=rr, column=3, value=g), "formula", "#,##0"); style(ws, ws.cell(row=rr, column=4, value=n), "formula", "#,##0")
        style(ws, ws.cell(row=rr, column=5, value=f"=D{rr}/$B$18"), "formula", "0.0%"); style(ws, ws.cell(row=rr, column=6, value=f"=D{rr}*$B$5"), "formula", "$#,##0")
        style(ws, ws.cell(row=rr, column=7, value=cst), "formula", "$#,##0"); style(ws, ws.cell(row=rr, column=8, value=f"=F{rr}-G{rr}"), "formula", "$#,##0;($#,##0)", bold=(i == 1))
        style(ws, ws.cell(row=rr, column=9, value=f"=C{rr}/(B{rr}*$B$21)"), "formula", "0%"); style(ws, ws.cell(row=rr, column=10, value=f"=C{rr}-D{rr}"), "formula", "#,##0")
    ws.cell(row=32, column=8).fill = KEY_FILL
    section(ws, 36, "Weighted scorecard (1 to 5)")
    header(ws, 37, ["Option", "Financial (35%)", "Guest experience (30%)", "Feasibility (20%)", "Risk (15%)", "Weighted score"])
    for i, (lab, s) in enumerate((("A. Extend hours", (1, 2, 4, 3)), ("B. Add Fridays", (5, 4, 3, 3)), ("C. Fifth weekend", (3, 3, 4, 2)))):
        rr = 38 + i; ws.cell(row=rr, column=1, value=lab)
        for j, v in enumerate(s, 2): style(ws, ws.cell(row=rr, column=j, value=v), "input", "0")
        style(ws, ws.cell(row=rr, column=6, value=f"=0.35*B{rr}+0.30*C{rr}+0.20*D{rr}+0.15*E{rr}"), "formula", "0.00", bold=(i == 1))
    ws.cell(row=42, column=1, value="2029 design day without new capacity"); style(ws, ws.cell(row=42, column=2, value="=(B18+3*1500+Capture!E32)*WkS3*SatShare"), "formula", "#,##0")
    ws.cell(row=42, column=3, value="Trend 1,500 a year plus the outer-ring ramp from the Capture sheet").font = Font(name=FONT, size=9, color=MUTED)
    bar(ws, "L2", "Net contribution by option", Reference(ws, min_col=1, min_row=31, max_row=33), Reference(ws, min_col=8, min_row=31, max_row=33), "$")

    # order sheets
    order = ["README", "Sources", "Inputs", "Baseline", "Capture", "Reach", "Anniversary", "Concert", "Macro", "Bridge", "Sensitivity", "Weekends", "Part2", "ErlangC", "Part3"]
    wb._sheets = [wb[n] for n in order]
    out_path.parent.mkdir(exist_ok=True); wb.save(out_path); return out_path


if __name__ == "__main__":
    print(build())
