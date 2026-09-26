// Builds the WDPR case deck from outputs/results.json.   node deck/build_deck.js
const pptxgen = require('pptxgenjs');
const fs = require('fs'); const path = require('path');
const R = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'outputs', 'results.json'), 'utf8'));
const OUT = process.argv[2] || path.join(__dirname, '..', 'outputs', 'WDPR_Fall_2026_Case_Deck.pptx');

const pres = new pptxgen(); pres.layout = 'LAYOUT_16x9';
pres.title = 'Autumn Harvest Festival, 25th Anniversary Season';

// Disney palette: navy dominant, Disney blue, gold accent, red for negatives only
const NAVY = '0B2545', BLUE = '1F5FBF', SKY = 'E8EFFA', GOLD = 'F2B632', RED = 'D7263D', TEXT = '1B1B1B', MUTED = '5B6470', WHITE = 'FFFFFF', LINE = 'D5DDEA';
const HFONT = 'Cambria', BFONT = 'Calibri';
const k = (x, d = 0) => Number(x).toLocaleString('en-US', { maximumFractionDigits: d, minimumFractionDigits: d });
const r100 = x => Math.round(x / 100) * 100;
const S = R.simulation, B = R.bridge, P2 = R.part2, P3 = R.part3, M = R.macro, C = R.capture, RE = R.reach, CO = R.concert, AN = R.anniversary, BL = R.baseline;
const season = r100(S.total_p10_p50_p90[1]), p10 = r100(S.total_p10_p50_p90[0]), p90 = r100(S.total_p10_p50_p90[2]);
const design = r100(S.wk3_sat_p10_p50_p90[1]);

function title(s, t) { s.addText(t, { x: 0.5, y: 0.3, w: 9.0, h: 0.7, fontFace: HFONT, fontSize: 24, bold: true, color: NAVY, isTextBox: true, margin: 0, valign: 'top' }); }
function source(s, t) { s.addText(t, { x: 0.5, y: 5.12, w: 9.0, h: 0.38, fontFace: BFONT, fontSize: 8, color: MUTED, isTextBox: true, margin: 0, valign: 'bottom' }); }
function card(s, x, y, w, h, fill = SKY) { s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.08 }); }
function body(s, t, x, y, w, h, o = {}) { s.addText(t, Object.assign({ x, y, w, h, fontFace: BFONT, fontSize: 11, color: TEXT, isTextBox: true, margin: 0.05, valign: 'top' }, o)); }
function bullets(s, items, x, y, w, h, o = {}) {
  const arr = items.map((t, i) => ({ text: t, options: { bullet: { indent: 12 }, breakLine: i < items.length - 1, paraSpaceAfter: 5 } }));
  s.addText(arr, Object.assign({ x, y, w, h, fontFace: BFONT, fontSize: 10.5, color: TEXT, isTextBox: true, margin: 0.05, valign: 'top' }, o));
}
function cardHead(s, t, x, y, w) { s.addText(t, { x, y, w, h: 0.3, fontFace: BFONT, fontSize: 11.5, bold: true, color: BLUE, isTextBox: true, margin: 0 }); }
function big(s, num, label, x, y, w, color = BLUE) {
  s.addText(num, { x, y, w, h: 0.7, fontFace: HFONT, fontSize: 32, bold: true, color, isTextBox: true, margin: 0 });
  s.addText(label, { x, y: y + 0.68, w, h: 0.32, fontFace: BFONT, fontSize: 10, color: MUTED, isTextBox: true, margin: 0 });
}
const chartBase = { chartColors: [BLUE], showLegend: false, showTitle: false, catAxisLabelFontSize: 9, valAxisLabelFontSize: 9, catAxisLabelColor: MUTED, valAxisLabelColor: MUTED, catAxisLabelFontFace: BFONT, valAxisLabelFontFace: BFONT, valGridLine: { color: 'E6EBF3', size: 0.5 }, catGridLine: { style: 'none' }, dataLabelFontSize: 9, dataLabelColor: TEXT, dataLabelFontFace: BFONT };
function tbl(colW, fs = 9) { return { x: 0, y: 0, colW, fontFace: BFONT, fontSize: fs, color: TEXT, border: { type: 'solid', color: LINE, pt: 0.5 }, fill: { color: WHITE }, valign: 'middle', margin: 0.04 }; }
const hdr = cells => cells.map(t => ({ text: t, options: { bold: true, color: WHITE, fill: { color: NAVY } } }));
const bold = (t, color) => ({ text: t, options: Object.assign({ bold: true }, color ? { color } : {}) });

// 1 Title
{ const s = pres.addSlide(); s.background = { color: NAVY };
  s.addText('Autumn Harvest Festival', { x: 0.7, y: 1.3, w: 8.6, h: 0.9, fontFace: HFONT, fontSize: 40, bold: true, color: WHITE, isTextBox: true, margin: 0 });
  s.addText('The 25th anniversary season: attendance forecast, vendor booth mix, and the growth path', { x: 0.7, y: 2.25, w: 8.6, h: 0.9, fontFace: BFONT, fontSize: 18, color: SKY, isTextBox: true, margin: 0 });
  s.addShape(pres.ShapeType.ellipse, { x: 0.7, y: 3.5, w: 0.22, h: 0.22, fill: { color: GOLD }, line: { color: GOLD } });
  s.addText('WDPR Fall 2026 Industrial Engineering Case Study', { x: 1.05, y: 3.45, w: 8, h: 0.35, fontFace: BFONT, fontSize: 13, color: WHITE, isTextBox: true, margin: 0 });
  s.addText('DIsnEOR   |   September 2026', { x: 1.05, y: 3.85, w: 8, h: 0.35, fontFace: BFONT, fontSize: 11, color: GOLD, isTextBox: true, margin: 0 });
  s.addNotes('Every number in this deck traces to outputs/results.json, the Excel workbook, and the GitHub repo. Rebuild in the team\'s own hands before submission; the case limits AI use to research.'); }

// 2 Executive summary
{ const s = pres.addSlide(); title(s, `${k(season)} guests, a 60 / 28 / 16 booth mix, and Fridays in 2027`);
  const xs = [0.5, 3.55, 6.6], cw = 2.9, cy = 1.2, ch = 3.8;
  const heads = ['Part 1. Attendance', 'Part 2. Booth mix', 'Part 3. Growth option'];
  const nums = [k(season), '60 / 28 / 16', 'Option B'];
  const labels = [`expected 2026 attendance (P10 ${k(p10)} to P90 ${k(p90)})`, 'Food & Beverage / Merchandise / Activities booths', 'add Fridays, piloted on Weekends 2 and 3 in 2027'];
  const txt = [[`Up ${Math.round((season / 56000 - 1) * 100)} percent on 2025 from trend, the 25th anniversary, wider reach, and the concert, net of a soft macro backdrop`, `Peak weekend stays Weekend 3 (${k(r100(S.weekends_p10_p50_p90[2][1]))}). Design day is Weekend 3 Saturday at about ${k(design)}`, 'Weekend 4 overtakes only if the concert act exceeds roughly 20M monthly listeners'],
    ['Uses all 120 standard booth units', 'Peak waits under 10 minutes for food and 5 for merchandise on the design day', `Every vendor clears a sales floor. Organizer contribution about $${k(P2.mixes.balanced.total_contribution / 1000)}k, within 2 percent of the profit maximum`],
    [`About ${k(r100(P3.options.B.net_new))} net new guests (+${Math.round(P3.options.B.net_new_pct * 100)} percent) and roughly $${k(P3.options.B.net / 1000)}k net contribution`, `Pulls about ${k(r100(P3.options.B.peak_relief))} visits off the weekend peak, the only option that relieves it and pays`, 'Pair with date based pricing so capacity is added where demand fills it']];
  for (let i = 0; i < 3; i++) { card(s, xs[i], cy, cw, ch);
    s.addText(heads[i], { x: xs[i] + 0.15, y: cy + 0.12, w: cw - 0.3, h: 0.3, fontFace: BFONT, fontSize: 11, bold: true, color: MUTED, isTextBox: true, margin: 0 });
    s.addText(nums[i], { x: xs[i] + 0.15, y: cy + 0.42, w: cw - 0.3, h: 0.6, fontFace: HFONT, fontSize: 28, bold: true, color: BLUE, isTextBox: true, margin: 0 });
    s.addText(labels[i], { x: xs[i] + 0.15, y: cy + 1.02, w: cw - 0.3, h: 0.45, fontFace: BFONT, fontSize: 9.5, color: MUTED, isTextBox: true, margin: 0 });
    bullets(s, txt[i], xs[i] + 0.1, cy + 1.5, cw - 0.2, 2.2, { fontSize: 10 }); }
  s.addNotes('Pyramid Principle: all three answers on one slide. Sources and the register are in the appendix.'); }

// 3 One system
{ const s = pres.addSlide(); title(s, 'One linked model: demand sets capacity, capacity sets the lever');
  const bx = [0.5, 3.75, 7.0], bw = 2.5, by = 1.3, bh = 1.75, fills = [NAVY, BLUE, GOLD], tc = [WHITE, WHITE, NAVY];
  const heads = ['Part 1  Forecast', 'Part 2  Booths', 'Part 3  Growth'];
  const sub = ['Driver based bridge from 2025 actuals. Output is season attendance and the design day.', 'Design day demand into an Erlang C queue model and vendor economics. Output is booths by category.', 'Load duration view of the season. Output is which option adds capacity where demand fills it.'];
  for (let i = 0; i < 3; i++) { s.addShape(pres.ShapeType.roundRect, { x: bx[i], y: by, w: bw, h: bh, fill: { color: fills[i] }, line: { color: fills[i] }, rectRadius: 0.08 });
    s.addText(heads[i], { x: bx[i] + 0.15, y: by + 0.12, w: bw - 0.3, h: 0.35, fontFace: HFONT, fontSize: 15, bold: true, color: tc[i], isTextBox: true, margin: 0 });
    s.addText(sub[i], { x: bx[i] + 0.15, y: by + 0.5, w: bw - 0.3, h: 1.2, fontFace: BFONT, fontSize: 10.5, color: tc[i], isTextBox: true, margin: 0 });
    if (i < 2) s.addShape(pres.ShapeType.rightArrow, { x: bx[i] + bw + 0.15, y: by + 0.65, w: 0.45, h: 0.45, fill: { color: LINE }, line: { color: LINE } }); }
  card(s, 0.5, 3.3, 9.0, 1.7);
  cardHead(s, 'Design day', 0.7, 3.4, 2); body(s, 'The specific high demand day capacity is sized around. We use the expected busiest Saturday rather than the single worst hour, so booths and staffing are built for the day that recurs.', 0.7, 3.7, 4.2, 1.2, { fontSize: 10.5 });
  cardHead(s, 'Assumption register', 5.2, 3.4, 3); body(s, 'Every input carries a source and a low, base and high value, in the appendix and in the Excel workbook. Where two methods disagree, the deck shows both and picks the overlap.', 5.2, 3.7, 4.1, 1.2, { fontSize: 10.5 });
  s.addNotes('The three parts are one workbook: the forecast tab feeds the booth tab, both feed the options tab.'); }

// 4 Factor map
{ const s = pres.addSlide(); title(s, 'Every factor, its direction, and where it is modeled');
  const rows = [hdr(['Factor', 'Direction', 'Size in 2026 (visits)', 'Where it lives']),
    ['Organic trend (2023 to 2025 momentum)', 'up', `+${k(r100(BL.combined - 56000))}`, 'Baseline: five methods, holdout, combination'],
    ['Marketing spend +20%, local budget held', 'up', `+${k(Math.round(RE.local_readings.A_local_visits))}`, 'Reach: elasticity 0.12, effective frequency'],
    ['Reach 500k to 1.2M residents', 'up', `+${k(RE.carried)} (1,000 to 4,500)`, 'Capture and Reach: four methods'],
    ['25th anniversary', 'up', `+${k(r100(56000 * 1.035 - 56000))}`, 'Anniversary: comps and four channels'],
    ['Weekend 4 concert (about a 10M act)', 'up', `+800 incremental, +100 shifted`, 'Concert: listener funnel'],
    ['Gas at $4.49, +25% y/y (outer ring trips)', 'down', `${k(Math.round(M.outer_visits_effect))}`, 'Macro: trip fuel elasticity'],
    ['Staycation substitution (weak sentiment)', 'up', `+${k(Math.round(M.local_visits_effect))}`, 'Macro: 2009 fair evidence'],
    ['Sentiment near record low', 'down', `${k(Math.round(M.sentiment_visits_effect))}`, 'Macro: small drag, spending still growing'],
    ['Inflation, real spend per guest', 'down on spend', '−4% per cap', 'Part 2 sales per booth, Part 3 revenue per guest'],
    ['Cost inflation (4.6% expected)', 'down on margin', '+4.6% on 2027 labor', 'Part 3 option costs'],
    ['Weather', 'down skew', `${k(Math.round(R.tornado.drivers.weather.low_delta))} to +${k(Math.round(R.tornado.drivers.weather.high_delta))}`, 'Simulation: multiplier 0.90 to 1.02'],
    ['Saturday share, distance decay, no-shows', 'shape', 'split and caps', 'Weekends, Capture, Concert'],
    ['2021 to 2022 pandemic regime', 'excluded', 'down-weighted in trend', 'Baseline: recent-regime weights']];
  const t = tbl([2.9, 0.95, 1.9, 3.25], 8.5); t.x = 0.5; t.y = 1.15; s.addTable(rows, t);
  source(s, 'Sizes are base-case deltas from the bridge and macro model (results.json). Macro readings: UMich sentiment 48.1 (Sept 2026) vs 55.1; AAA gasoline $4.49 on Sept 25 2026; UMich one-year inflation expectations 4.6%.');
  s.addNotes('This slide answers "did you factor everything in." Each row points to the sheet in the workbook and the module in the repo.'); }

// 5 Baseline
{ const s = pres.addSlide(); title(s, `Organic trend alone puts 2026 near ${k(r100(BL.combined))}`);
  s.addChart(pres.ChartType.line, [{ name: 'Attendance', labels: ['2021', '2022', '2023', '2024', '2025', '2026 base'], values: [49000, 48500, 50000, 53500, 56000, Math.round(BL.combined)] }],
    Object.assign({}, chartBase, { x: 0.5, y: 1.2, w: 4.6, h: 3.6, lineSize: 2.5, lineDataSymbol: 'circle', lineDataSymbolSize: 7, showValue: true, dataLabelPosition: 't', dataLabelFormatCode: '#,##0', valAxisMinVal: 44000, valAxisMaxVal: 62000, valAxisLabelFormatCode: '#,##0' }));
  const f = BL.fits, h = BL.holdout_2025;
  const rows = [hdr(['Method', '2026', 'Holdout miss on 2025']),
    ['A. Linear trend, all years', k(f.A_ols_all), `${(h.A_ols_all * 100).toFixed(1)}%`], ['B. Log linear (3.7% CAGR)', k(f.B_loglin_all), `${(h.B_loglin_all * 100).toFixed(1)}%`],
    ['C. Linear, 2023 to 2025', k(f.C_ols_recent3), `${(h.C_ols_recent3 * 100).toFixed(1)}%`], ['D. Log linear, 2023 to 2025', k(f.D_loglin_recent3), `${(h.D_loglin_recent3 * 100).toFixed(1)}%`], ['E. Holt smoothing', k(f.E_holt), `${(h.E_holt * 100).toFixed(1)}%`],
    [bold('Combined (15/15/25/20/25)'), bold(k(Math.round(BL.combined)), BLUE), bold(`80% interval ${k(Math.round(BL.pi80_linear[0]))} to ${k(Math.round(BL.pi80_linear[1]))}`)]];
  const t = tbl([2.15, 0.75, 1.45], 9); t.x = 5.4; t.y = 1.2; s.addTable(rows, t);
  bullets(s, ['2021 and 2022 are a pandemic regime, not the trend. Summerfest was still 38 percent below 2019 in 2022.', 'Every method under forecast 2024 and 2025, so growth is accelerating and recent years get more weight.'], 5.4, 3.55, 4.1, 1.4, { fontSize: 10 });
  source(s, 'Holdout: hide 2025, forecast it from 2021 to 2024, compare. Regime evidence: BizTimes, Summerfest attendance 2022 and 2023. Code: src/festival/baseline.py; workbook sheet Baseline.'); }

// 6 Capture
{ const s = pres.addSlide(); title(s, 'Where guests come from: 70 percent local, 25 percent neighboring');
  const rg = C.rings;
  const rows = [hdr(['Ring', 'Population', 'Share of 2025 visits', 'Visits', 'Capture rate']),
    ['Local metro (advertised)', '500,000', '70%', k(rg.local_visits), bold(`${(rg.local_capture * 100).toFixed(1)}%`, BLUE)],
    ['Neighboring metros (new reach)', '700,000', '25%', k(rg.outer_visits), bold(`${(rg.outer_capture * 100).toFixed(1)}%`, BLUE)],
    ['Beyond the reach area', 'n/a', '5%', k(rg.beyond_visits), 'n/a']];
  const t = tbl([2.3, 1.0, 1.35, 0.85, 1.0], 9.5); t.x = 0.5; t.y = 1.2; s.addTable(rows, t);
  s.addChart(pres.ChartType.bar, [{ name: 'Capture rate', labels: ['Local ring', 'Neighboring ring'], values: [+(rg.local_capture * 100).toFixed(1), +(rg.outer_capture * 100).toFixed(1)] }],
    Object.assign({}, chartBase, { x: 7.2, y: 1.2, w: 2.3, h: 2.0, showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '0.0"%"', valAxisHidden: true, valAxisMaxVal: 10 }));
  card(s, 0.5, 2.85, 6.4, 2.2);
  bullets(s, ['Capture rate is visits divided by ring population. The festival already draws from the neighboring ring with no paid media there, through word of mouth and spillover.', `The outer to local ratio of ${rg.ratio_outer_to_local.toFixed(2)} is what a gravity model predicts at a distance exponent near 1.5 (Esri Huff range 1.5 to 2.0). Distance, not awareness alone, explains most of the gap.`, 'So the reach expansion lifts awareness in a market where word of mouth is already running, not a blank market. That matters for how fast it responds.'], 0.6, 2.95, 6.2, 2.05, { fontSize: 10 });
  bullets(s, [`Ceiling at full awareness: ${(C.calibrated.outer_ceiling_capture * 100).toFixed(1)}% outer capture, ${k(r100(C.calibrated.outer_ceiling_visits))} visits.`, 'Origin split is an assumption. A zip code intercept survey at the gate replaces it in one season.'], 7.1, 3.3, 2.4, 1.7, { fontSize: 9.5 });
  source(s, 'Local share 70% matches verified surveys: Bolton Food and Drink Festival 72% local postcode; Big Sur Food and Wine 70% Monterey County. Promoter claims of 60 to 80% out of region are rarely verified (Stratford Analytics). Exponent: Esri Huff model documentation. Sheet Capture; capture.py.'); }

// 7 Marketing and reach
{ const s = pres.addSlide(); title(s, `Marketing and reach: about ${k(RE.carried)} new visits`);
  card(s, 0.5, 1.2, 4.3, 3.85); cardHead(s, 'The dilution arithmetic', 0.65, 1.3, 4);
  bullets(s, ['Money up 1.2x, audience up 2.4x. Spread evenly, spend per person reached falls by half.', `Reading A, local budget held and new money outward: local spillover +0.5 percent, about ${k(Math.round(RE.local_readings.A_local_visits))} visits.`, `Reading B, spread evenly: elasticity 0.12 times minus 50 percent intensity is minus 6 percent locally, about ${k(Math.round(RE.local_readings.B_local_visits))}. Reading A wins by roughly 2,400 visits.`, 'Effective frequency research (Krugman 1972, Naples 1979) says ads need repeated exposure before they work. Hold the local budget and flight the new money into the three weeks before Weekend 1.'], 0.6, 1.65, 4.1, 3.35, { fontSize: 10 });
  s.addChart(pres.ChartType.bar, [{ name: 'High end', labels: ['Elasticity', 'Awareness funnel', 'Bass diffusion', 'Capture model'], values: [Math.round(RE.elasticity_low_high[1]), Math.round(RE.funnel), Math.round(RE.bass_year_one_low_high[1]), Math.round(RE.capture_model * 1.67)] }],
    Object.assign({}, chartBase, { x: 5.0, y: 1.2, w: 4.5, h: 2.5, showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '#,##0', valAxisLabelFormatCode: '#,##0', valAxisMaxVal: 3600, chartColors: [GOLD] }));
  body(s, `Four methods bracket the new market lift (high ends shown). Ranges: elasticity ${k(Math.round(RE.elasticity_low_high[0]))} to ${k(Math.round(RE.elasticity_low_high[1]))}; funnel 700,000 x 25% aware x 1.7% convert; Bass ${k(Math.round(RE.bass_year_one_low_high[0]))} to ${k(Math.round(RE.bass_year_one_low_high[1]))}; capture model ${k(Math.round(RE.capture_model * 0.6))} to ${k(Math.round(RE.capture_model * 1.67))}. Carry ${k(RE.carried)}, range 1,000 to 4,500.`, 5.0, 3.75, 4.5, 1.3, { fontSize: 10 });
  source(s, 'Sethuraman, Tellis and Briesch 2011 (JMR): mean short term elasticity 0.12, long term 0.24. Longwoods: Washington 36% aware, 77,000 trips from 2.4M aware households; Wyoming intent 7% to 14%. Sultan, Farley, Lehmann 1990: p 0.03, q 0.38. Sheet Reach; reach.py.'); }

// 8 Anniversary
{ const s = pres.addSlide(); title(s, 'The 25th anniversary: plus 3.5 percent net');
  const td = AN.top_down;
  s.addChart(pres.ChartType.bar, [{ name: 'Lift above trend', labels: ['Disneyland 50th (2005)', 'Disneyland 60th (2015)', 'Musikfest 40th (2023)'], values: [+(td.disneyland_50th * 100).toFixed(1), +(td.disneyland_60th * 100).toFixed(1), +(td.musikfest_40th * 100).toFixed(1)] }],
    Object.assign({}, chartBase, { x: 0.5, y: 1.2, w: 4.3, h: 2.4, showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '0.0" pts"', valAxisMaxVal: 9, valAxisHidden: true }));
  body(s, `Indy 500's 100th running (2016) sold out for the first time in decades and did not sell out again until 2025. Milestone demand can exceed capacity. Disneyland gave back ${(td.disneyland_60th_giveback * 100).toFixed(1)}% the year after its 60th.`, 0.5, 3.65, 4.3, 0.9, { fontSize: 10 });
  card(s, 5.0, 1.2, 4.5, 3.85); cardHead(s, 'Bottom up, four channels', 5.15, 1.3, 4);
  const bu = AN.bottom_up;
  bullets(s, [`Lapsed visitors return. About 20 percent of visits each year are not annual regulars (New York State Fair, four in five are regulars). A ${k(r100(bu.lapsed_pool))} person lapsed pool at 3 percent is about ${k(r100(bu.reactivated))} visits.`, `Earned media. Anniversary press reaches locals who do not attend. About ${k(r100(bu.earned_media))} visits.`, 'Special programming beyond the concert. About 1 percent.', `Regulars add a weekend. 30 percent already attend multiple times. About ${k(r100(bu.extra_regular_visits))} visits.`, `Stacked ${(bu.stacked_pct * 100).toFixed(1)} percent, netted for overlap ${(bu.netted_pct * 100).toFixed(1)}. Carry 3.5, range 1 to 7.`, 'Giveback. 2026 is a pulse, not the 2027 baseline. Part 3 plans on trend plus the diffusion ramp.'], 5.1, 1.65, 4.3, 3.35, { fontSize: 9.5 });
  source(s, 'TEA/AECOM Theme Index (Disneyland 16.77M, 18.28M, 17.94M). Amusement Business/ERA 2005 (Disneyland +8.5%, top 50 parks +4.2%). Morning Call, Musikfest 2023. Des Moines Register/ESPN on Indy 500 sellouts. NY State Fair 2015 exit survey. Sheet Anniversary; anniversary.py.'); }

// 9 Concert
{ const s = pres.addSlide(); title(s, 'Concert draw depends on the act: a 10M act adds about 800');
  const tiers = CO.tiers;
  s.addChart(pres.ChartType.bar, [{ name: 'Incremental guests', labels: ['1M listeners', '5M', '10M', '20M', '40M'], values: ['1M', '5M', '10M', '20M', '40M'].map(t => Math.round(tiers[t].incremental)) }],
    Object.assign({}, chartBase, { x: 0.5, y: 1.2, w: 4.4, h: 2.6, showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '#,##0', valAxisLabelFormatCode: '#,##0', valAxisMaxVal: 3600, chartColors: [GOLD] }));
  bullets(s, ['Weekend 4 Saturday overtakes Weekend 3 Saturday above roughly 20M combined listeners. A flat 2,000 incremental would need about 27M.', `A regional 1M act with four times the local density (${k(Math.round(CO.regional_1M.incremental))}) beats a playlist driven 10M act (${k(Math.round(CO.playlist_10M.incremental))}). Ask the booker for the Spotify city list, not the national number.`], 0.5, 3.85, 4.4, 1.25, { fontSize: 9.5 });
  card(s, 5.1, 1.2, 4.4, 3.85); cardHead(s, 'Listener funnel', 5.25, 1.3, 4);
  bullets(s, ['Combined listeners: Spotify public count plus Apple at 0.35x (Spotify 37% and Apple 31% of US subscribers, but Spotify counts free users), less 7% overlap.', 'Catchment: the reach area holds 0.23% of a US act\'s listeners with the outer ring weighted 0.25 for distance. Affinity 1.0 national, up to 4x regional.', 'Conversion: 1 to 3% of metro listeners buy paid tickets; free with admission, 10%. Super listeners (2% of audience, 50% of sales) set an engagement multiplier.', 'Show up 70%. Ten percent of fans are regulars who shift weekends. Casual on grounds guests fill 40% of seats.', `If reservations are used, issue about ${k(Math.round(CO.passes_to_issue))} passes: free events lose 40 to 60% to no shows.`], 5.2, 1.65, 4.2, 3.35, { fontSize: 9.5 });
  source(s, 'insights.live (1 to 3% metro conversion; super listeners). Orphiq (2% working estimate). Digital Music News Aug 2025, DemandSage (US shares). VenueSight/Eventbrite (no shows). Google Area 120 DEMAND: 70% of tickets start with a Google search, so Google Trends by metro is a free affinity proxy. Sheet Concert; concert.py.'); }

// 10 Macro
{ const s = pres.addSlide(); title(s, 'Macro backdrop: visits down 0.8 percent, spend down 4 percent');
  const rows = [hdr(['Indicator', 'September 2026', 'A year ago', 'Read']),
    ['UMich consumer sentiment', '48.1', '55.1', 'Near record lows; fell across income groups'],
    ['One-year inflation expectations', '4.6%', '3.6%', 'High prices are the top complaint'],
    ['Gasoline, national average', '$4.49', '$3.60', 'Highest ever for late September; Iran conflict'],
    ['Unemployment / payrolls', '4.1% / +162k', '4.2%', 'Labor market intact'],
    ['Retail and food-service sales, y/y', '+6.0% nominal', '', 'People feel bad and keep spending']];
  const t = tbl([1.95, 1.05, 0.8, 2.6], 8.5); t.x = 0.5; t.y = 1.2; s.addTable(rows, t);
  card(s, 7.05, 1.2, 2.45, 3.85);
  big(s, `${(M.net_attendance_pct * 100).toFixed(1)}%`, 'net on attendance (about ' + k(Math.round(M.net_attendance_visits)) + ' visits)', 7.2, 1.3, 2.2, RED);
  big(s, '−4%', 'real spend per guest (trade-down)', 7.2, 2.45, 2.2, RED);
  big(s, '+4.6%', 'cost escalator on 2027 options', 7.2, 3.6, 2.2, NAVY);
  card(s, 0.5, 3.15, 6.35, 1.9);
  bullets(s, [`Staycation substitution. A $12 local festival is the cheap alternative to travel. In 2009 fairs reported booming business as families skipped out-of-town trips. Local ring +1 percent, about +${k(Math.round(M.local_visits_effect))} visits.`, `Fuel cost on the outer ring. Gas is up ${(M.gas_change_pct * 100).toFixed(0)} percent; a 90-mile round trip costs $${M.extra_fuel_cost_per_car.toFixed(2)} more. Total gasoline demand is barely elastic (Hughes, Knittel, Sperling: minus 0.03 to minus 0.08), but discretionary day trips are more so. At minus 0.2, outer visits fall ${(M.outer_visit_effect_pct * 100).toFixed(1)} percent, about ${k(Math.round(M.outer_visits_effect))}.`, `Real income squeeze. Sentiment barely predicts visits (spending is still growing), so a small drag of ${k(Math.round(M.sentiment_visits_effect))}. The bigger effect is trade-down inside the gate: minus 4 percent real per cap, which Part 2 and Part 3 carry.`], 0.6, 3.22, 6.15, 1.8, { fontSize: 9 });
  source(s, 'UMich Surveys of Consumers Sept 2026; NY Fed SCE Aug 2026; AAA Sept 25 2026; Empower/AAA y/y; BLS and Census via Sept 2026 coverage; NBC News Aug 2009 (Salem Fair ticket and food sales +25%); CBS Minnesota (2009 record); Star Tribune Sept 2025; Hughes, Knittel and Sperling, Energy Journal 2008 (NBER w12530). Sheet Macro; macro.py.'); }

// 11 Bridge
{ const s = pres.addSlide(); title(s, `Bridge: 56,000 in 2025 to an expected ${k(season)} in 2026`);
  const st = B.steps;
  s.addChart(pres.ChartType.bar, [{ name: 'Cumulative', labels: st.map(x => x[0]).concat(['weather adj.']), values: st.map(x => Math.round(x[1])).concat([season]) }],
    Object.assign({}, chartBase, { x: 0.5, y: 1.2, w: 6.0, h: 3.8, showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '#,##0', valAxisLabelFormatCode: '#,##0', valAxisMinVal: 50000, valAxisMaxVal: 66000 }));
  card(s, 6.75, 1.2, 2.75, 3.85);
  big(s, k(p10), 'Low (P10)', 6.95, 1.35, 2.4, MUTED); big(s, k(season), 'Expected (P50)', 6.95, 2.45, 2.4, BLUE); big(s, k(p90), 'High (P90)', 6.95, 3.55, 2.4, MUTED);
  body(s, `${Math.round(S.p_over_60k * 100)}% chance of clearing 60,000`, 6.95, 4.6, 2.4, 0.35, { fontSize: 9.5, color: MUTED });
  source(s, `Percentages applied first (trend, marketing 0.5%, anniversary 3.5%, macro ${(M.net_attendance_pct * 100).toFixed(1)}%), headcounts after (reach 2,000, concert 800). Weather multiplier 0.90 to 1.02, mode 1.00. 100,000 Monte Carlo draws (simulate.py). Fair weather total ${k(Math.round(B.fair_weather_total))}; the downside skewed weather pulls the median lower. Sheet Bridge.`); }

// 12 Weekends
{ const s = pres.addSlide(); title(s, `Weekend 3 Saturday is the design day, at about ${k(design)} guests`);
  const wk = S.weekends_p10_p50_p90;
  s.addChart(pres.ChartType.bar, [{ name: 'Expected', labels: ['Weekend 1', 'Weekend 2', 'Weekend 3', 'Weekend 4'], values: wk.map(w => r100(w[1])) }],
    Object.assign({}, chartBase, { x: 0.5, y: 1.2, w: 5.0, h: 3.0, showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '#,##0', valAxisLabelFormatCode: '#,##0', valAxisMaxVal: 21000 }));
  body(s, `P10 to P90: ${wk.map(w => `${k(r100(w[0]))} to ${k(r100(w[2]))}`).join(' | ')}. Weekend 4 carries the concert guests and 100 shifted regulars.`, 0.5, 4.25, 5.0, 0.8, { fontSize: 9.5, color: MUTED });
  card(s, 5.75, 1.2, 3.75, 3.85); cardHead(s, 'Saturday is the design day', 5.9, 1.3, 3.5);
  bullets(s, ['Saturday runs 55 percent of a weekend. Fair data shows Saturday 8 to 37 percent above Sunday.', `Weekend 3 Saturday: about ${k(design)} (${k(r100(S.wk3_sat_p10_p50_p90[0]))} to ${k(r100(S.wk3_sat_p10_p50_p90[2]))}). Weekend 4 Saturday: about ${k(r100(S.wk4_sat_p10_p50_p90[1]))} with the concert pulse concentrated in the evening.`, `Weekend 4 Saturday wins in ${Math.round(S.p_wk4sat_over_wk3sat * 100)} percent of simulations at the 10M act tier, and in every simulation above roughly a 20M act. Plan a flat top: size Weekends 3 and 4 the same.`, 'Programming moves peaks. The Colorado Renaissance Festival\'s themed pirate weekend is its busiest, and Minnesota\'s standing Sunday record came with a sold out Reba show.'], 5.85, 1.65, 3.55, 3.35, { fontSize: 9.5 });
  source(s, 'Minnesota State Fair daily attendance 2022 to 2024. Colorado Springs Gazette 2026. Star Tribune 2026 attendance tracker. Sheet Weekends.'); }

// 13 Tornado + tracking
{ const s = pres.addSlide(); title(s, 'Weather swings most; reach is the biggest lever');
  const T = R.tornado.drivers; const order = ['weather', 'anniversary', 'reach', 'baseline', 'marketing', 'macro', 'concert'];
  const labs = { weather: 'Weather', anniversary: 'Anniversary', reach: 'Reach', baseline: 'Baseline trend', marketing: 'Local marketing', macro: 'Macro', concert: 'Concert' };
  s.addChart(pres.ChartType.bar, [{ name: 'Low case', labels: order.map(o => labs[o]), values: order.map(o => Math.round(T[o].low_delta)) }, { name: 'High case', labels: order.map(o => labs[o]), values: order.map(o => Math.round(T[o].high_delta)) }],
    Object.assign({}, chartBase, { x: 0.5, y: 1.2, w: 5.2, h: 3.6, barDir: 'bar', barGrouping: 'stacked', chartColors: [GOLD, BLUE], showLegend: true, legendPos: 'b', legendFontSize: 9, legendFontFace: BFONT, valAxisLabelFormatCode: '#,##0', catAxisOrientation: 'maxMin', catAxisLabelPos: 'low' }));
  card(s, 5.95, 1.2, 3.55, 3.85); cardHead(s, 'Treat the forecast as a tracking signal', 6.1, 1.3, 3.3);
  bullets(s, ['Leading indicators: Google Trends for the festival name in August and September versus 2024 and 2025, social follower growth, press mentions, vendor applications, any advance sales.', 'Decision rule: search interest up more than 30 percent year over year by mid September moves the anniversary lift to the high case (7 percent) and Weekend 3 and 4 booths with it. Flat interest drops it to the low case.', 'Precedent: during Disneyland\'s 50th, Anaheim hotel occupancy ran up 9 percent year over year in July and resort trolley ridership rose 24 percent in the launch month, both visible before the year closed.'], 6.05, 1.65, 3.35, 3.35, { fontSize: 9.5 });
  source(s, 'Delta from the fair weather total at each driver\'s 10th and 90th percentile, others at mode. Weather anchored to Minnesota 2024, one rain day at 81,231 against a normal 135,000. Orange County Business Journal 2005 on Anaheim indicators. Sheet Sensitivity.'); }

// 14 Part 2 method
{ const s = pres.addSlide(); title(s, 'Part 2 method: fixed demand, booths as capacity');
  const bx = [0.5, 3.75, 7.0], bw = 2.5, by = 1.2, bh = 1.65, fills = [NAVY, BLUE, GOLD], tc = [WHITE, WHITE, NAVY];
  const pd = P2.peak_hour_demand, fl = P2.floors;
  const heads = ['1. Design day demand', '2. Queue floor', '3. Economics'];
  const sub = [`Peak hour transactions on Weekend 3 Saturday: food ${k(Math.round(pd['F&B']))}, merchandise ${k(Math.round(pd.Merch))}, activities ${k(Math.round(pd['A&E']))}. Peak hour is 15 to 18 percent of the day.`, `Erlang C (M/M/c) for the fewest booths that hold the wait target: ${fl['F&B']} food (10 min), ${fl.Merch} merchandise (5 min), ${fl['A&E']} activities (15 min).`, 'Booth fee plus revenue share minus organizer cost per booth, a vendor sales floor, and a saturating capture curve, the retail space elasticity idea. Real spend per guest carried at minus 4 percent.'];
  for (let i = 0; i < 3; i++) { s.addShape(pres.ShapeType.roundRect, { x: bx[i], y: by, w: bw, h: bh, fill: { color: fills[i] }, line: { color: fills[i] }, rectRadius: 0.08 });
    s.addText(heads[i], { x: bx[i] + 0.15, y: by + 0.1, w: bw - 0.3, h: 0.3, fontFace: HFONT, fontSize: 13, bold: true, color: tc[i], isTextBox: true, margin: 0 });
    s.addText(sub[i], { x: bx[i] + 0.15, y: by + 0.42, w: bw - 0.3, h: 1.2, fontFace: BFONT, fontSize: 9.5, color: tc[i], isTextBox: true, margin: 0 }); }
  const rows = [hdr(['Category', 'Spend', 'Participation', 'Tx per guest', 'Service', 'Fee', 'Share', 'Org. cost', 'Vendor floor']),
    ['Food & Beverage', '$21.50', '77.5%', '1.6', '2 points, 2.5 min', '$2,000', '12%', '$900', '$18,000'], ['Merchandise', '$16.00', '32.5%', '1.1', '1 point, 3 min', '$900', '10%', '$500', '$5,000'], ['Activities (2 units)', '$8.50', '40%', '1.3', '60 guests/hr', '$700', '8%', '$600', '$4,000']];
  const t = tbl([1.55, 0.7, 0.95, 0.9, 1.35, 0.75, 0.65, 0.85, 1.3], 9); t.x = 0.5; t.y = 3.05; s.addTable(rows, t);
  source(s, 'Spend, participation and space from the case (midpoints). Service rates, fees, revenue share, organizer cost and vendor floors are team assumptions; fee and share are per season. Capture reaches 85% at the queue floor and rises with diminishing returns (Eisend 2014 space elasticity near 0.17). Sheets Part2 and ErlangC; booths.py.'); }

// 15 Part 2 recommendation
{ const s = pres.addSlide(); title(s, 'Recommend 60 food, 28 merchandise, 16 activities');
  const mx = P2.mixes; const row = (nm, lab, o) => { const m = mx[nm].rows; const c = [lab, '120', `${Math.round(m['F&B'].capture * 100)}%`, `$${k(r100(m['F&B'].sales_per_booth))}`, `$${k(r100(m.Merch.sales_per_booth))}`, `${m['F&B'].wait_min.toFixed(1)} / ${m.Merch.wait_min.toFixed(1)} / ${m['A&E'].wait_min.toFixed(1)} min`, `$${k(r100(mx[nm].total_contribution))}`]; return o ? c.map((x, i) => (i === 0 || i === 6) ? bold(x, i === 0 ? BLUE : undefined) : x) : c; };
  const rows = [hdr(['Mix (F&B / Merch / A&E)', 'Units', 'F&B capture', 'F&B sales per booth', 'Merch sales per booth', 'Peak wait F&B / Merch / A&E', 'Organizer contribution']),
    row('profit_max', `Profit max  ${mx.profit_max.rows['F&B'].booths} / ${mx.profit_max.rows.Merch.booths} / ${mx.profit_max.rows['A&E'].booths}`), row('balanced', 'Balanced  60 / 28 / 16  (recommended)', true), row('variety', 'Variety  56 / 32 / 16')];
  const t = tbl([2.3, 0.6, 0.85, 1.25, 1.3, 1.6, 1.1], 9); t.x = 0.5; t.y = 1.2; s.addTable(rows, t);
  s.addChart(pres.ChartType.bar, [{ name: 'Booths', labels: ['Food & Beverage', 'Merchandise', 'Activities'], values: [60, 28, 16] }],
    Object.assign({}, chartBase, { x: 0.5, y: 2.6, w: 3.6, h: 2.4, showValue: true, dataLabelPosition: 'outEnd', valAxisHidden: true, valAxisMaxVal: 72 }));
  card(s, 4.3, 2.6, 5.2, 2.45);
  bullets(s, ['Food gets 60 because it is the destination category and the best economics per unit, but the 61st booth adds under 1 percent capture while cutting every vendor\'s take.', 'Merchandise holds at its queue floor of 28. Activities gets 16, two above its floor, for the family identity and for the Weekend 4 concert evening when arrivals bunch.', `Robust to the forecast range and to trade-down: in the low case (${k(p10)}) with real spend down 4 percent, food sales per booth are $${k(r100(P2.low_case_check.rows['F&B'].sales_per_booth))}, still above the $18,000 floor.`, 'Flex rule for Weekends 3 and 4: if food waits pass 10 minutes at lunch, convert two activity slots to food for the weekend.'], 4.4, 2.7, 5.0, 2.3, { fontSize: 9.5 });
  source(s, `Season attendance ${k(season)} and design day ${k(design)} from Part 1. Contribution equals fees plus revenue share minus organizer cost. Vendor floors are sales a vendor needs to return next year. Profit max found by grid search over all 120-unit mixes (booths.py). Sheet Part2.`); }

// 16 Part 3 comparison
{ const s = pres.addSlide(); title(s, 'Option B (Fridays) adds the most per dollar; A loses money');
  const o = P3.options; const $k = v => `$${k(Math.round(v / 1000))}k`;
  const rows = [hdr(['Option', 'Hours added', 'Gross demand', 'Net new (after shifting)', 'Revenue', 'Cost', 'Net', 'Capacity factor', 'Peak relief']),
    ['A. Extend to 12 hours', '+16 (20%)', k(r100(o.A.gross)), `${k(r100(o.A.net_new))} (+${(o.A.net_new_pct * 100).toFixed(1)}%)`, $k(o.A.revenue), $k(o.A.cost), bold(`minus $${k(Math.round(-o.A.net / 1000))}k`, RED), `${Math.round(o.A.capacity_factor * 100)}%`, 'within day only'],
    [bold('B. Add Fridays'), '+40 (50%)', k(r100(o.B.gross)), `${k(r100(o.B.net_new))} (+${(o.B.net_new_pct * 100).toFixed(1)}%)`, $k(o.B.revenue), $k(o.B.cost), bold(`+${$k(o.B.net)}`, BLUE), `${Math.round(o.B.capacity_factor * 100)}%`, `${k(r100(o.B.peak_relief))} off weekends`],
    ['C. Fifth weekend', '+20 (25%)', k(r100(o.C.gross)), `${k(r100(o.C.net_new))} (+${(o.C.net_new_pct * 100).toFixed(1)}%)`, $k(o.C.revenue), $k(o.C.cost), `+${$k(o.C.net)}`, `${Math.round(o.C.capacity_factor * 100)}%`, `${k(r100(o.C.peak_relief))} off weekends`]];
  const t = tbl([1.5, 0.85, 0.9, 1.35, 0.75, 0.7, 0.85, 0.85, 1.25], 9); t.x = 0.5; t.y = 1.2; s.addTable(rows, t);
  s.addChart(pres.ChartType.bar, [{ name: 'Net contribution', labels: ['A. Hours', 'B. Fridays', 'C. Fifth weekend'], values: [Math.round(o.A.net / 1000), Math.round(o.B.net / 1000), Math.round(o.C.net / 1000)] }],
    Object.assign({}, chartBase, { x: 0.5, y: 2.65, w: 3.6, h: 2.35, showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '"$"#,##0"k"', valAxisLabelFormatCode: '"$"#,##0"k"', valAxisMinVal: -20, valAxisMaxVal: 100, chartColors: [GOLD] }));
  card(s, 4.3, 2.65, 5.2, 2.4);
  bullets(s, ['Load duration logic: capacity only pays where demand fills it. Option A adds hours at the thin ends of the arrival curve, and most of what it captures is pulled forward from hours already open, the Thanksgiving retail hours pattern.', 'Option B is Thursday Night Football: weaker per day (Friday at 45 percent of a Saturday), incremental inventory, and the only option that moves real load off the Saturday peak.', 'Option C is the 17th game: it pays only if demand exceeds supply. Late season weather and holiday collisions add risk that A and B do not carry.', 'Costs carry the 4.6 percent inflation expectation; revenue per guest carries the 4 percent real trade-down.'], 4.4, 2.75, 5.0, 2.25, { fontSize: 9.5 });
  source(s, `Organizer revenue $${P3.revenue_per_guest.toFixed(2)} per guest (admission $12, revenue share on spend, parking). Labor 50 staff at $22/hr escalated 4.6%. Redistribution 60% (A), 35% (B), 40% (C). Friday demand 45% of Saturday (fair data shows weekdays consistently lowest). Fifth weekend at 65% of an average weekend. Sheet Part3; options.py.`); }

// 17 Part 3 recommendation
{ const s = pres.addSlide(); title(s, 'Pilot Fridays in 2027, paired with date based pricing');
  const sc = P3.scores;
  s.addChart(pres.ChartType.bar, [{ name: 'Weighted score', labels: ['A. Hours', 'B. Fridays', 'C. Fifth weekend'], values: [sc.A, sc.B, sc.C] }],
    Object.assign({}, chartBase, { x: 0.5, y: 1.2, w: 3.4, h: 2.4, showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '0.0', valAxisHidden: true, valAxisMaxVal: 5, chartColors: [BLUE] }));
  body(s, 'Weights: financial 35%, guest experience 30%, feasibility 20%, risk 15%. Scores 1 to 5.', 0.5, 3.65, 3.4, 0.5, { fontSize: 9, color: MUTED });
  card(s, 4.1, 1.2, 5.4, 3.85);
  bullets(s, ['Why a pilot. Vendor consent and weekday staffing are the binding risk. Two Fridays test demand and vendor economics for about $45k before committing to four. That is a real option, not a half measure.', `Why now. Trend adds about 1,500 a year and the neighboring ring ramps toward a 21,500 visit ceiling as awareness spreads (about +6,500 by 2030). The design day reaches roughly ${k(r100(P3.design_day_2029))} by 2029 without any new capacity.`, 'Demand side first. Date based admission pricing and early entry incentives level load before capacity is built, the same logic utilities use before adding a peaker plant. Disney runs its parks on date based tickets.', 'Macro fit. In a weak-sentiment year the cheap local Friday evening is exactly the staycation product; if sentiment recovers, Fridays inherit the travel-substitution demand rather than lose it.', 'What to watch. Friday gate as a share of Saturday (go if above 40 percent), vendor renewal rate, Saturday peak occupancy, and per cap on Fridays.', 'Do not budget 2027 off 2026. The anniversary is a one year pulse; plan on trend plus the diffusion ramp.'], 4.2, 1.3, 5.2, 3.7, { fontSize: 9.3 });
  source(s, 'Diffusion ramp from the capture rate model with q = 0.38 word of mouth (sheet Capture). 2029 design day = (season + 3 years of trend + outer ring ramp) x 30% x 55%.'); }

// 18 Validate
{ const s = pres.addSlide(); title(s, 'Five things to validate before opening day');
  const items = [['Concert act', 'Pull the act\'s Spotify city list and Google Trends by metro. Above roughly 20M listeners the design day moves to Weekend 4 Saturday.'], ['Guest origin', 'Zip code intercept at the gate. Replaces the 70 / 25 / 5 split and recalibrates both capture rates in one season.'], ['Marketing allocation', 'Confirm the local budget is held. Spreading evenly costs about 2,400 visits. Flight new money before Weekend 1.'], ['Weather and macro', 'Largest swing in the model. Pre agree a rain day playbook. Re-read gas prices and sentiment in early September and move the macro driver.'], ['Anniversary giveback', 'Budget 2027 on trend, not 2026. Disneyland gave back 1.8 percent; Indy waited nine years for its next sellout.']];
  const cw = 1.72, gap = 0.1, cy = 1.2, ch = 3.75, cols = [NAVY, BLUE, GOLD, NAVY, BLUE];
  for (let i = 0; i < 5; i++) { const x = 0.5 + i * (cw + gap); card(s, x, cy, cw, ch);
    s.addShape(pres.ShapeType.ellipse, { x: x + 0.12, y: cy + 0.15, w: 0.4, h: 0.4, fill: { color: cols[i] }, line: { color: cols[i] } });
    s.addText(String(i + 1), { x: x + 0.12, y: cy + 0.15, w: 0.4, h: 0.4, fontFace: HFONT, fontSize: 14, bold: true, color: cols[i] === GOLD ? NAVY : WHITE, align: 'center', valign: 'middle', isTextBox: true, margin: 0 });
    s.addText(items[i][0], { x: x + 0.1, y: cy + 0.65, w: cw - 0.2, h: 0.4, fontFace: BFONT, fontSize: 11, bold: true, color: BLUE, isTextBox: true, margin: 0 });
    body(s, items[i][1], x + 0.05, cy + 1.05, cw - 0.1, 2.6, { fontSize: 9.5 }); } }

// 19 Appendix A
{ const s = pres.addSlide(); title(s, 'Appendix A. Assumption register and sources (Part 1)');
  const rows = [hdr(['Input', 'Value used', 'Range', 'Source']),
    ['Organic baseline', k(r100(BL.combined)), `${k(Math.round(BL.pi80_linear[0]))} to ${k(Math.round(BL.pi80_linear[1]))} (80% PI)`, 'Five methods on 2021 to 2025 actuals, weighted to the recent regime'],
    ['Advertising elasticity', '0.12 short, 0.24 long', 'n/a', 'Sethuraman, Tellis and Briesch, Journal of Marketing Research 2011'],
    ['Campaign awareness, new ring', '25%', '15 to 35%', 'Longwoods and SMARI: Washington 36%, New Mexico 44%, North Dakota 47% at higher per capita spend'],
    ['Conversion among aware', '1.7%', '1.3 to 2.5%', 'Washington 77,000 trips from 2.4M aware households; Wyoming intent 7% to 14% (Longwoods)'],
    ['Historical origin split', '70 / 25 / 5', 'outer 15 to 35%', 'Bolton Food and Drink Festival 72% local; Big Sur Food and Wine 70% Monterey County'],
    ['Distance weight, outer ring', '0.25', '0.14 to 0.4', 'Esri Huff model documentation, exponent 1.5 to 2.0'],
    ['Bass p and q', '0.03, 0.38', 'p up to 0.10 with heavy ads', 'Sultan, Farley and Lehmann, JMR 1990, 213 parameter sets'],
    ['Anniversary lift', '3.5%', '1 to 7%', 'TEA/AECOM Theme Index (Disneyland), Amusement Business 2005, Morning Call (Musikfest), Indy 500 sellout history'],
    ['Macro net effect', `${(M.net_attendance_pct * 100).toFixed(1)}%`, '−2 to +1.5%', 'UMich Sept 2026; AAA Sept 2026; Hughes, Knittel, Sperling 2008; NBC News 2009; Star Tribune 2025'],
    ['Saturday share of weekend', '55%', '52 to 58%', 'Minnesota State Fair daily attendance 2022 to 2024'],
    ['Weather multiplier', '0.90 to 1.02', 'mode 1.00', 'Minnesota 2024 rain day 81,231 versus about 135,000 normal'],
    ['Listener to guest conversion (free)', '10%', '4 to 18%', 'insights.live 1 to 3% paid metro conversion; Orphiq 2% working estimate; super listeners 2% drive 50%'],
    ['Streaming shares', 'Apple = 0.35 x Spotify', '0.25 to 0.45', 'Digital Music News Aug 2025; DemandSage; Spotify 751M MAU vs 290M paid'],
    ['Free event no shows', '40 to 60%', 'n/a', 'VenueSight via eventcloud; Eventbrite up to 50%']];
  const t = tbl([2.1, 1.35, 1.5, 4.05], 7.5); t.x = 0.5; t.y = 1.05; s.addTable(rows, t); }

// 20 Appendix B
{ const s = pres.addSlide(); title(s, 'Appendix B. Assumptions behind Parts 2 and 3');
  card(s, 0.5, 1.1, 4.35, 3.95); cardHead(s, 'Part 2 booth model', 0.65, 1.2, 4);
  bullets(s, [`Design day ${k(design)} (Weekend 3 Saturday P50). Season ${k(season)}. Real spend per guest x 0.96.`, 'Transactions per participating guest: food 1.6, merchandise 1.1, activities 1.3. Peak hour factor 18% food, 15% others.', 'Service: food booth two points at 2.5 min each (48/hr), merchandise one point at 3 min (20/hr), activity booth 60 guests/hr (OHRC).', `Wait targets at peak: 10 min food, 5 min merchandise, 15 min activities. Erlang C floors ${P2.floors['F&B']}, ${P2.floors.Merch}, ${P2.floors['A&E']}.`, 'Capture curve 1 minus exp(minus n / n0), n0 set so capture is 85% at the queue floor.', 'Organizer economics per season: fees $2,000 / $900 / $700, revenue share 12 / 10 / 8%, organizer cost $900 / $500 / $600. Vendor floors $18,000 / $5,000 / $4,000.'], 0.6, 1.55, 4.15, 3.45, { fontSize: 9.5 });
  card(s, 5.15, 1.1, 4.35, 3.95); cardHead(s, 'Part 3 option model', 5.3, 1.2, 4);
  bullets(s, [`Organizer revenue per guest $${P3.revenue_per_guest.toFixed(2)}: admission $12, revenue share on spend after trade-down, parking $1.50.`, 'Labor 50 staff at $22/hr, escalated 4.6% for 2027. Operations $8,000 per added day. Extra site week $20,000. Lighting and security $2,500 per extended day. Launch marketing $8,000 to $10,000.', 'Option A: the two added hours carry 7% of arrivals, 60% pulled forward from open hours.', 'Option B: Friday at 45% of a Saturday (range 35 to 60), 35% shifted from the weekend.', 'Option C: fifth weekend at 65% of an average weekend, 40% shifted.', 'Capacity factor equals gross demand in the added hours divided by average demand per hour times hours added.', 'Scorecard weights 35 / 30 / 20 / 15 for financial, guest experience, feasibility, risk.'], 5.25, 1.55, 4.15, 3.45, { fontSize: 9.5 }); }

// 21 Appendix C: audit trail
{ const s = pres.addSlide(); title(s, 'Appendix C. The workbook and the code');
  card(s, 0.5, 1.1, 4.35, 3.95); cardHead(s, 'Excel workbook (WDPR_Festival_Model.xlsx)', 0.65, 1.2, 4.1);
  bullets(s, ['Blue cells are inputs with a source beside them; black cells are formulas; every derived number recalculates from the Inputs sheet.', 'Sheets follow the deck: Sources, Inputs, Baseline, Capture, Reach, Anniversary, Concert, Macro, Bridge, Sensitivity, Weekends, Part2 (with an ErlangC helper), Part3.', 'Regressions use FORECAST, SLOPE and STEYX on the case data; Holt smoothing is a two-column recursion; Erlang C waits are computed per booth count.', 'Monte Carlo percentiles are pasted from the code and labeled as such; everything else is live.'], 0.6, 1.55, 4.15, 3.45, { fontSize: 9.5 });
  card(s, 5.15, 1.1, 4.35, 3.95); cardHead(s, 'GitHub: NathanielKung/disney-festival-ie-case', 5.3, 1.2, 4.1);
  bullets(s, ['data/: case actuals, case parameters, and every research anchor with its URL.', 'src/festival/: baseline (five trend methods, holdout, prediction interval), capture, reach, concert, anniversary, macro, simulate (100,000 draw Monte Carlo and tornado), booths (Erlang C and mix search), options.', 'tests/: 18 pytest checks, including Erlang C reducing to M/M/1, the capture model reproducing observed rates, and the option ranking.', 'scripts/run_all.py writes outputs/results.json; scripts/build_workbook.py writes the Excel; deck/build_deck.js writes this deck from results.json.'], 5.25, 1.55, 4.15, 3.45, { fontSize: 9.5 }); }

pres.writeFile({ fileName: OUT }).then(f => console.log('wrote', f));
