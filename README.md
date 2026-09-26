# Autumn Harvest Festival, WDPR Fall 2026 IE Case

Attendance forecast, vendor booth mix, and growth-option analysis for a 25-year-old regional fall festival, built for the Walt Disney Parks and Resorts Industrial Engineering case competition.

Three deliverables come out of this repo, and they agree with each other to the dollar:

| Output | Built by | What it is |
|---|---|---|
| `outputs/results.json` | `scripts/run_all.py` | Every model result, Part 1 through Part 3 |
| `outputs/WDPR_Festival_Model.xlsx` | `scripts/build_workbook.py` | Excel workbook: inputs in blue with sources, 1,300 live formulas, native charts |
| `outputs/WDPR_Fall_2026_Case_Deck.pptx` | `deck/build_deck.js` | 21-slide deck written from `results.json` |

## Run it

```bash
pip install -r requirements.txt
python scripts/run_all.py          # models -> outputs/results.json
python scripts/build_workbook.py   # -> outputs/WDPR_Festival_Model.xlsx
python -m pytest -q                # 18 checks
node deck/build_deck.js            # -> outputs/WDPR_Fall_2026_Case_Deck.pptx (needs pptxgenjs)
```

## What the models do

**Part 1, attendance (`src/festival/`)**

- `baseline.py` fits five trend methods to 2021 to 2025 (OLS linear, log-linear, both again on the post-pandemic years, and Holt double exponential smoothing), backtests each by hiding the last year, combines them with weights that favor the recent regime, and reports an 80 percent prediction interval. Combined organic baseline: about 58,600.
- `capture.py` treats attendance divided by ring population as a capture rate. The festival only advertised to the 500,000 local residents, but roughly a quarter of visits already come from the 700,000 in neighboring metros (7.8 percent local capture, 2.0 percent outer). A calibrated model backs conversion out of the local ring and a distance factor out of the outer ring, then projects the outer ring's five-year path as awareness spreads by a Bass process.
- `reach.py` brackets the new-market lift four ways: advertising elasticity (Sethuraman, Tellis and Briesch 2011), an awareness-times-conversion funnel (Longwoods and SMARI tourism ROI studies), Bass diffusion year one (Sultan, Farley and Lehmann 1990), and the capture model. The overlap is about 2,000 visits.
- `anniversary.py` sizes the milestone lift top-down from Disneyland's 50th and 60th, Musikfest's 40th and the Indy 500's 100th, and bottom-up from lapsed-visitor reactivation, earned media, programming and extra regular visits. Both land near 3.5 to 4.5 percent.
- `concert.py` turns the act's combined Spotify and Apple Music monthly listeners into incremental guests through a catchment, conversion, engagement, show-up and seat-cap funnel. A 10M-listener act adds about 750; Weekend 4 only becomes the peak above roughly 20M.
- `macro.py` converts September 2026 readings (UMich sentiment 48.1, one-year inflation expectations 4.6 percent, AAA gasoline $4.49 up 25 percent) into three mechanisms: staycation substitution on the local ring, fuel cost on outer-ring day trips, and a real-income drag. Net about minus 0.8 percent on visits and minus 4 percent on real spend per guest.
- `simulate.py` runs the driver bridge, a 100,000-draw Monte Carlo on triangular and PERT distributions, a tornado sensitivity, and the weekend and Saturday split that produces the design day. Expected 2026 attendance 62,100 (P10 59,100, P90 64,700); design day Weekend 3 Saturday at about 10,100.

**Part 2, booth mix (`booths.py`)**

Design-day demand by category, Erlang C (M/M/c) queue floors that hold wait targets at the lunch peak, a saturating capture curve for booth count, vendor sales floors, and a grid search over every 120-unit mix. Recommendation 60 food, 28 merchandise, 16 activities, within 2 percent of the profit maximum.

**Part 3, growth options (`options.py`)**

Load-duration logic: capacity only pays where demand fills it. Net new demand after redistribution, revenue per guest, escalated costs, capacity factor, peak relief and a weighted scorecard. Option B (add Fridays) wins; pilot on two weekends in 2027.

## Data and sources

`data/research_anchors.csv` lists every research input with its low, base and high value, the source, and the URL. `data/comparables.csv` holds the milestone and fair-day comparables. `data/case_parameters.csv` and `data/attendance_2021_2025.csv` are the case's own numbers.

## Notes

The competition allows AI for research only. The research anchors and the reference implementation here were assembled with that help; the analysis and presentation submitted are the team's own.
