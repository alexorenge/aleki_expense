const fs = require('fs');
const path = require('path');

const pptxgen = require('pptxgenjs');

// Helpers (layout-safe shadows, etc.)
const {
  safeOuterShadow,
  warnIfSlideHasOverlaps,
  warnIfSlideElementsOutOfBounds,
} = require('/home/oai/share/slides/pptxgenjs_helpers');


const dataPath = path.join(__dirname, 'alex_expense_summary.json');
const outPath = path.join(__dirname, 'Alex_Expense_Presentation.pptx');

if (!fs.existsSync(dataPath)) {
  throw new Error(`Missing data file: ${dataPath}`);
}

const D = JSON.parse(fs.readFileSync(dataPath, 'utf8'));

// --------- Theme constants ---------
const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'Generated with PptxGenJS';

// Wide layout: 13.333 x 7.5 in
const SLIDE_W = 13.333;
const SLIDE_H = 7.5;

const M = 0.6; // margin
const ACCENT = '1F77B4';
const MUTED = '6B7280';
const SOFT_BG = 'F5F7FA';

function addHeader(slide, title, subtitle) {
  // Top bar
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: SLIDE_W,
    h: 0.9,
    fill: { color: SOFT_BG },
    line: { color: SOFT_BG },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0.88,
    w: SLIDE_W,
    h: 0.02,
    fill: { color: ACCENT },
    line: { color: ACCENT },
  });

  slide.addText(title, {
    x: M,
    y: 0.12,
    w: SLIDE_W - 2 * M,
    h: 0.48,
    fontFace: 'Calibri',
    fontSize: 24,
    bold: true,
    color: '111827',
  });

  if (subtitle) {
    slide.addText(subtitle, {
      x: M,
      y: 0.60,
      w: SLIDE_W - 2 * M,
      h: 0.22,
      fontFace: 'Calibri',
      fontSize: 12,
      color: MUTED,
    });
  }
}

function addKpiCard(slide, x, y, w, h, label, value, note) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h,
    fill: { color: 'FFFFFF' },
    line: { color: 'E5E7EB', width: 1 },
    radius: 10,
    shadow: safeOuterShadow('000000', 0.12, 45, 2, 1),
  });
  slide.addText(label, {
    x: x + 0.25,
    y: y + 0.2,
    w: w - 0.5,
    h: 0.3,
    fontFace: 'Calibri',
    fontSize: 12,
    color: MUTED,
  });
  slide.addText(value, {
    x: x + 0.25,
    y: y + 0.55,
    w: w - 0.5,
    h: 0.6,
    fontFace: 'Calibri',
    fontSize: 28,
    bold: true,
    color: '111827',
  });
  if (note) {
    slide.addText(note, {
      x: x + 0.25,
      y: y + h - 0.35,
      w: w - 0.5,
      h: 0.25,
      fontFace: 'Calibri',
      fontSize: 10,
      color: MUTED,
    });
  }
}

function fmtInt(n) {
  return new Intl.NumberFormat('en-US').format(Math.round(n));
}

// ---------- Slide 1: Title ----------
{
  const slide = pptx.addSlide();
  slide.background = { color: 'FFFFFF' };

  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: SLIDE_W,
    h: SLIDE_H,
    fill: { color: 'FFFFFF' },
    line: { color: 'FFFFFF' },
  });
  // Accent block
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: 1.2,
    h: SLIDE_H,
    fill: { color: ACCENT },
    line: { color: ACCENT },
  });
  slide.addText('Alex Expense Overview', {
    x: 1.5,
    y: 2.5,
    w: SLIDE_W - 2.0,
    h: 0.8,
    fontFace: 'Calibri',
    fontSize: 44,
    bold: true,
    color: '111827',
  });
  slide.addText(`Period: ${D.period.start} to ${D.period.end}`, {
    x: 1.55,
    y: 3.35,
    w: SLIDE_W - 2.2,
    h: 0.35,
    fontFace: 'Calibri',
    fontSize: 16,
    color: MUTED,
  });
  slide.addText(
    `Total spend: ${fmtInt(D.kpis.total_spend)} | Transactions: ${fmtInt(D.kpis.transactions)}`,
    {
      x: 1.55,
      y: 3.75,
      w: SLIDE_W - 2.2,
      h: 0.35,
      fontFace: 'Calibri',
      fontSize: 16,
      color: MUTED,
    }
  );
}

// ---------- Slide 2: KPI summary ----------
{
  const slide = pptx.addSlide();
  slide.background = { color: 'FFFFFF' };
  addHeader(
    slide,
    'Snapshot',
    'Key metrics from the provided transaction log (currency not specified; figures shown as raw amounts).'
  );

  const y0 = 1.3;
  const cardW = (SLIDE_W - 2 * M - 0.45) / 2;
  const cardH = 1.55;

  addKpiCard(
    slide,
    M,
    y0,
    cardW,
    cardH,
    'Total spend',
    fmtInt(D.kpis.total_spend),
    `Highest txn: ${fmtInt(D.kpis.max_txn)}`
  );
  addKpiCard(
    slide,
    M + cardW + 0.45,
    y0,
    cardW,
    cardH,
    'Transactions',
    fmtInt(D.kpis.transactions),
    `Avg/txn: ${fmtInt(D.kpis.avg_txn)}`
  );

  // Highlight insights
  slide.addShape(pptx.ShapeType.roundRect, {
    x: M,
    y: y0 + cardH + 0.55,
    w: SLIDE_W - 2 * M,
    h: 3.9,
    fill: { color: SOFT_BG },
    line: { color: 'E5E7EB', width: 1 },
    radius: 10,
  });
  slide.addText('Highlights', {
    x: M + 0.25,
    y: y0 + cardH + 0.75,
    w: SLIDE_W - 2 * M - 0.5,
    h: 0.3,
    fontFace: 'Calibri',
    fontSize: 16,
    bold: true,
    color: '111827',
  });
  const t1 = D.insights.top_types[0];
  const t2 = D.insights.top_types[1];
  const bullets = [
    `Highest-spend month: ${D.insights.highest_month.month} (${fmtInt(D.insights.highest_month.total)})`,
    `Lowest-spend month: ${D.insights.lowest_month.month} (${fmtInt(D.insights.lowest_month.total)})`,
    `Top categories: ${t1.type} (${fmtInt(t1.total)}, ${t1.pct.toFixed(1)}%) and ${t2.type} (${fmtInt(
      t2.total
    )}, ${t2.pct.toFixed(1)}%)`,
    `Car fuel share of car spend: ${D.insights.car_fuel_share_pct.toFixed(1)}%`,
  ];
  slide.addText(bullets.map((b) => ({ text: b, options: { bullet: { indent: 18 } } })), {
    x: M + 0.35,
    y: y0 + cardH + 1.15,
    w: SLIDE_W - 2 * M - 0.7,
    h: 3.2,
    fontFace: 'Calibri',
    fontSize: 14,
    color: '111827',
    paraSpaceAfter: 10,
  });
}

// ---------- Slide 3: Monthly trend (Line) ----------
{
  const slide = pptx.addSlide();
  slide.background = { color: 'FFFFFF' };
  addHeader(slide, 'Monthly spend trend', 'Total spend per month.');

  const chartData = [
    {
      name: 'Total spend',
      labels: D.months,
      values: D.monthly_totals,
    },
  ];

  slide.addChart(pptx.ChartType.line, chartData, {
    x: M,
    y: 1.25,
    w: SLIDE_W - 2 * M,
    h: 4.9,
    showLegend: false,
    valAxisFormatCode: '#,##0',
    catAxisLabelRotation: 0,
    dataLabelPosition: 't',
    lineDataSymbol: 'circle',
  });

  // Callouts
  slide.addShape(pptx.ShapeType.roundRect, {
    x: M,
    y: 6.35,
    w: SLIDE_W - 2 * M,
    h: 0.85,
    fill: { color: SOFT_BG },
    line: { color: 'E5E7EB', width: 1 },
    radius: 10,
  });
  slide.addText(
    `Peak: ${D.insights.highest_month.month} (${fmtInt(D.insights.highest_month.total)})   •   Low: ${D.insights.lowest_month.month} (${fmtInt(
      D.insights.lowest_month.total
    )})`,
    {
      x: M + 0.35,
      y: 6.57,
      w: SLIDE_W - 2 * M - 0.7,
      h: 0.35,
      fontFace: 'Calibri',
      fontSize: 14,
      color: '111827',
    }
  );
  slide.addText('Tip: compare “count of transactions” vs “total” to see whether changes are driven by frequency or big one-offs.', {
    x: M + 0.35,
    y: 6.92,
    w: SLIDE_W - 2 * M - 0.7,
    h: 0.25,
    fontFace: 'Calibri',
    fontSize: 11,
    color: MUTED,
  });
}

// ---------- Slide 4: Category breakdown (Bar) ----------
{
  const slide = pptx.addSlide();
  slide.background = { color: 'FFFFFF' };
  addHeader(slide, 'Spend by category', 'Totals by Type.');

  const topN = 10;
  const labels = D.type_breakdown.labels.slice(0, topN);
  const values = D.type_breakdown.values.slice(0, topN);

  const chartData = [
    {
      name: 'Total spend',
      labels,
      values,
    },
  ];

  slide.addChart(pptx.ChartType.bar, chartData, {
    x: M,
    y: 1.25,
    w: SLIDE_W - 2 * M,
    h: 5.7,
    barDir: 'bar',
    barGrouping: 'clustered',
    showLegend: false,
    valAxisFormatCode: '#,##0',
    dataLabelPosition: 'outEnd',
  });
}

// ---------- Slide 5: Monthly mix (Stacked columns) ----------
{
  const slide = pptx.addSlide();
  slide.background = { color: 'FFFFFF' };
  addHeader(slide, 'Monthly spend mix', 'Top categories stacked by month.');

  slide.addChart(pptx.ChartType.bar, D.monthly_mix_top6, {
    x: M,
    y: 1.25,
    w: SLIDE_W - 2 * M,
    h: 5.6,
    barDir: 'col',
    barGrouping: 'stacked',
    legendPos: 'r',
    showLegend: true,
    valAxisFormatCode: '#,##0',
  });
}

// ---------- Slide 6: Payment method (Doughnut) ----------
{
  const slide = pptx.addSlide();
  slide.background = { color: 'FFFFFF' };
  addHeader(slide, 'Payment methods', 'Share of spend by payment type.');

  const chartData = [
    {
      name: 'Spend',
      labels: D.payment.labels,
      values: D.payment.values,
    },
  ];

  slide.addChart(pptx.ChartType.doughnut, chartData, {
    x: M,
    y: 1.35,
    w: 6.2,
    h: 5.6,
    showLegend: true,
    legendPos: 'r',
    dataLabelPosition: 'bestFit',
  });

  const total = D.payment.values.reduce((a, b) => a + b, 0);
  const mpesaIdx = D.payment.labels.findIndex((s) => s.toLowerCase().includes('mpesa'));
  const visaIdx = D.payment.labels.findIndex((s) => s.toLowerCase().includes('visa'));
  const mpesaPct = mpesaIdx >= 0 ? (D.payment.values[mpesaIdx] / total) * 100 : null;
  const visaPct = visaIdx >= 0 ? (D.payment.values[visaIdx] / total) * 100 : null;

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 7.2,
    y: 1.6,
    w: SLIDE_W - 7.2 - M,
    h: 2.4,
    fill: { color: SOFT_BG },
    line: { color: 'E5E7EB', width: 1 },
    radius: 10,
  });
  slide.addText('Takeaway', {
    x: 7.45,
    y: 1.8,
    w: SLIDE_W - 7.45 - M,
    h: 0.3,
    fontFace: 'Calibri',
    fontSize: 16,
    bold: true,
    color: '111827',
  });
  slide.addText(
    [
      `Mpesa: ${mpesaPct?.toFixed(1) ?? '—'}% of spend`,
      `Visa: ${visaPct?.toFixed(1) ?? '—'}% of spend`,
      'If you want deeper insight, add a “merchant” column and tag recurring bills.',
    ].map((t) => ({ text: t, options: { bullet: { indent: 18 } } })),
    {
      x: 7.45,
      y: 2.2,
      w: SLIDE_W - 7.45 - M,
      h: 1.7,
      fontFace: 'Calibri',
      fontSize: 13,
      color: '111827',
      paraSpaceAfter: 8,
    }
  );
}

// ---------- Slide 7: Car deep dive ----------
{
  const slide = pptx.addSlide();
  slide.background = { color: 'FFFFFF' };
  addHeader(slide, 'Deep dive: Car', 'Breakdown of car-related spending.');

  const chartData = [
    {
      name: 'Car spend',
      labels: D.car.labels,
      values: D.car.values,
    },
  ];

  slide.addChart(pptx.ChartType.bar, chartData, {
    x: M,
    y: 1.25,
    w: 7.2,
    h: 5.9,
    barDir: 'col',
    showLegend: false,
    valAxisFormatCode: '#,##0',
    dataLabelPosition: 'outEnd',
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 8.05,
    y: 1.25,
    w: SLIDE_W - 8.05 - M,
    h: 5.9,
    fill: { color: SOFT_BG },
    line: { color: 'E5E7EB', width: 1 },
    radius: 10,
  });
  slide.addText('What stands out', {
    x: 8.3,
    y: 1.45,
    w: SLIDE_W - 8.3 - M,
    h: 0.35,
    fontFace: 'Calibri',
    fontSize: 16,
    bold: true,
    color: '111827',
  });
  slide.addText(
    [
      `Fuel share of car spend: ${D.insights.car_fuel_share_pct.toFixed(1)}%`,
      'Fuel is frequent; service is occasional (and can be lumpy).',
      'Consider tracking distance (km) to compute cost per km.',
    ].map((t) => ({ text: t, options: { bullet: { indent: 18 } } })),
    {
      x: 8.3,
      y: 1.95,
      w: SLIDE_W - 8.3 - M,
      h: 2.2,
      fontFace: 'Calibri',
      fontSize: 13,
      color: '111827',
      paraSpaceAfter: 8,
    }
  );

  slide.addText('Optional next metrics', {
    x: 8.3,
    y: 4.3,
    w: SLIDE_W - 8.3 - M,
    h: 0.35,
    fontFace: 'Calibri',
    fontSize: 12,
    bold: true,
    color: MUTED,
  });
  slide.addText(
    ['Fuel per week', 'Cost per km', 'Service sinking fund'].map((t) => ({
      text: t,
      options: { bullet: { indent: 18 } },
    })),
    {
      x: 8.3,
      y: 4.7,
      w: SLIDE_W - 8.3 - M,
      h: 1.2,
      fontFace: 'Calibri',
      fontSize: 13,
      color: '111827',
      paraSpaceAfter: 6,
    }
  );
}

// ---------- Slide 8: Recommendations ----------
{
  const slide = pptx.addSlide();
  slide.background = { color: 'FFFFFF' };
  addHeader(slide, 'Recommendations', 'Practical ways to make tracking and budgeting easier.');

  slide.addShape(pptx.ShapeType.roundRect, {
    x: M,
    y: 1.35,
    w: SLIDE_W - 2 * M,
    h: 5.95,
    fill: { color: 'FFFFFF' },
    line: { color: 'E5E7EB', width: 1 },
    radius: 14,
    shadow: safeOuterShadow('000000', 0.10, 45, 2, 1),
  });

  const recs = [
    'Separate fixed obligations (Rent, Tithe) from discretionary spending (Food, Car).',
    'Set monthly targets per discretionary category and review mid-month against actuals.',
    'Batch small purchases when possible (e.g., groceries) to reduce impulse spend.',
    'Add two columns to your tracker: “Month” (auto) and “Notes/Tags” (recurring, one-off, etc.).',
    'For fuel: track odometer/km to derive cost per km and spot changes early.',
  ];

  slide.addText('Next steps', {
    x: M + 0.35,
    y: 1.6,
    w: SLIDE_W - 2 * M - 0.7,
    h: 0.4,
    fontFace: 'Calibri',
    fontSize: 18,
    bold: true,
    color: '111827',
  });

  slide.addText(recs.map((t) => ({ text: t, options: { bullet: { indent: 20 } } })), {
    x: M + 0.45,
    y: 2.15,
    w: SLIDE_W - 2 * M - 0.9,
    h: 4.7,
    fontFace: 'Calibri',
    fontSize: 15,
    color: '111827',
    paraSpaceAfter: 10,
  });
}

// --- Layout checks for severe issues ---
for (const s of pptx._slides) {
  warnIfSlideHasOverlaps(s, pptx);
  warnIfSlideElementsOutOfBounds(s, pptx);
}

// Write the file
pptx.writeFile({ fileName: outPath });
console.log(`Wrote: ${outPath}`);
