const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.defineLayout({ name: 'WIDE', width: 13.333, height: 7.5 });
pptx.layout = 'WIDE';

const C = {
  accent: 'D9530A',
  accentHover: 'B8440A',
  bg: 'FFFFFF',
  bgSecondary: 'FAF7F5',
  bgOffset: 'F3EDE8',
  text: '241C15',
  muted: '6B5F54',
  faint: 'A79C90',
  success: '3F7D33',
  warning: 'B8440A',
  border: 'E0D8D0',
  white: 'FFFFFF',
};

function footer(slide, num) {
  slide.addText('empeo — Performance Insights', {
    x: 0.6, y: 7.05, w: 5, h: 0.35, fontSize: 8, color: C.muted, fontFace: 'Calibri',
  });
  slide.addText(String(num), {
    x: 12.0, y: 7.05, w: 0.7, h: 0.35, fontSize: 8, color: C.muted, fontFace: 'Calibri', align: 'right',
  });
}

function slideTitle(slide, title, num) {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 1.1, fill: { color: C.bg }, border: false });
  slide.addText(title, {
    x: 0.7, y: 0.2, w: 11, h: 0.6, fontSize: 26, color: C.accent, fontFace: 'Calibri', bold: true,
  });
  slide.addShape(pptx.ShapeType.rect, { x: 0.7, y: 0.8, w: 2.5, h: 0.04, fill: { color: C.accent }, border: false });
  footer(slide, num);
}

function pageContent(slide) {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 1.1, w: 13.333, h: 6.4, fill: { color: C.bgSecondary }, border: false });
}

function card(slide, x, y, w, h) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x, y, w, h, fill: { color: C.white },
    shadow: { type: 'outer', blur: 4, offset: 2, color: '000000', opacity: 0.08 },
    rectRadius: 0.1, line: { color: C.border, width: 0.5 },
  });
}

// ===== SLIDE 1: TITLE =====
const s1 = pptx.addSlide();
s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: C.bg } });
s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.35, h: 7.5, fill: { color: C.accent }, border: false });
s1.addShape(pptx.ShapeType.rect, { x: 0.7, y: 2.8, w: 4.5, h: 0.06, fill: { color: C.accent }, border: false });
s1.addText('empeo Performance Results & Insights', {
  x: 0.7, y: 2.0, w: 11, h: 0.8, fontSize: 40, color: C.accent, fontFace: 'Calibri', bold: true,
});
s1.addText('Elevating Performance Evaluation for Confident People Decisions', {
  x: 0.7, y: 3.1, w: 10, h: 0.5, fontSize: 18, color: C.muted, fontFace: 'Calibri',
});
s1.addText('Product Owner  |  ' + new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' }), {
  x: 0.7, y: 5.5, w: 8, h: 0.35, fontSize: 12, color: C.muted, fontFace: 'Calibri',
});

// ===== SLIDE 2: PROBLEM STATEMENT =====
const s2 = pptx.addSlide();
slideTitle(s2, 'The Problem', 2);
pageContent(s2);
s2.addShape(pptx.ShapeType.roundRect, {
  x: 0.7, y: 1.8, w: 11.9, h: 4.5, fill: { color: C.white },
  shadow: { type: 'outer', blur: 6, offset: 3, color: '000000', opacity: 0.1 },
  rectRadius: 0.15, line: false,
});
s2.addText('"The current performance results screen shows a single aggregated score with no depth across the evaluation dimensions, lacks decision-ready insight for salary, bonus, or promotion decisions, and cannot reliably support fair, transparent people outcomes."', {
  x: 1.2, y: 2.0, w: 11.0, h: 1.6, fontSize: 16, color: C.text, fontFace: 'Calibri', italic: true, lineSpacingMultiple: 1.4,
});
s2.addShape(pptx.ShapeType.rect, { x: 1.2, y: 3.8, w: 1.2, h: 0.03, fill: { color: C.accent }, border: false });
// Three impact boxes
const gaps = [
  { label: 'Information Gap', desc: 'Employees see a single number with no dimensional breakdown or trend context.' },
  { label: 'Decision Gap', desc: 'HR cannot confidently recommend salary adjustments, bonuses, or promotions from the data.' },
  { label: 'Trust Gap', desc: 'Without transparency, employees question the fairness and accuracy of their evaluation.' },
];
gaps.forEach((g, i) => {
  const x = 1.2 + i * 3.9;
  s2.addShape(pptx.ShapeType.roundRect, {
    x, y: 4.2, w: 3.5, h: 1.8, fill: { color: 'FFF5ED' }, rectRadius: 0.1, line: { color: C.accent, width: 0.5 },
  });
  s2.addText(g.label, { x: x + 0.2, y: 4.35, w: 3.1, h: 0.4, fontSize: 13, color: C.accent, fontFace: 'Calibri', bold: true });
  s2.addText(g.desc, { x: x + 0.2, y: 4.7, w: 3.1, h: 1.1, fontSize: 11, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.2 });
});

// ===== SLIDE 3: PROBLEM BREAKDOWN =====
const s3 = pptx.addSlide();
slideTitle(s3, 'Problem Breakdown — What\'s Missing', 3);
pageContent(s3);
const cols = [
  { title: 'Information Gaps', items: ['Single-score only', 'No dimension breakdown', 'No historical trend', 'No peer comparison'] },
  { title: 'Decision-Support Gaps', items: ['No eligibility logic', 'No predictive insight', 'No calibration view', 'No bias detection'] },
  { title: 'UX & Role Gaps', items: ['One view for all roles', 'No manager dashboard', 'No employee development view', 'No interactive exploration'] },
];
cols.forEach((col, ci) => {
  const x = 0.7 + ci * 4.2;
  card(s3, x, 1.5, 3.8, 5.2);
  s3.addShape(pptx.ShapeType.rect, { x: x + 0.3, y: 1.7, w: 0.8, h: 0.04, fill: { color: C.accent }, border: false });
  s3.addText(col.title, { x: x + 0.3, y: 1.85, w: 3.2, h: 0.4, fontSize: 15, color: C.accent, fontFace: 'Calibri', bold: true });
  col.items.forEach((item, ii) => {
    s3.addShape(pptx.ShapeType.roundRect, { x: x + 0.3, y: 2.5 + ii * 0.75, w: 0.12, h: 0.12, fill: { color: C.accent }, border: false });
    s3.addText(item, { x: x + 0.6, y: 2.4 + ii * 0.75, w: 2.8, h: 0.35, fontSize: 12, color: C.text, fontFace: 'Calibri' });
  });
});

// ===== SLIDE 4: ROOT CAUSES =====
const s4 = pptx.addSlide();
slideTitle(s4, 'Root Causes', 4);
pageContent(s4);
const causes = [
  { icon: '1', label: 'Single-Score Reliance', desc: 'Averages hide strengths, weaknesses, and patterns across dimensions' },
  { icon: '2', label: 'No Calibration Insight', desc: 'Managers cannot compare ratings across teams to ensure consistency' },
  { icon: '3', label: 'No Context or Benchmarking', desc: 'Scores lack departmental, role, or peer benchmarks for reference' },
  { icon: '4', label: 'One-Size-Fits-All View', desc: 'HR, managers, and employees all see the same limited interface' },
];
causes.forEach((c, i) => {
  const row = Math.floor(i / 2);
  const col = i % 2;
  const x = 0.7 + col * 6.2;
  const y = 1.5 + row * 2.7;
  card(s4, x, y, 5.8, 2.3);
  s4.addShape(pptx.ShapeType.roundRect, {
    x: x + 0.3, y: y + 0.3, w: 0.7, h: 0.7, fill: { color: C.accent }, rectRadius: 0.35, border: false,
  });
  s4.addText(c.icon, { x: x + 0.3, y: y + 0.3, w: 0.7, h: 0.7, fontSize: 18, color: C.white, fontFace: 'Calibri', bold: true, align: 'center', valign: 'middle' });
  s4.addText(c.label, { x: x + 1.2, y: y + 0.25, w: 4.2, h: 0.4, fontSize: 15, color: C.text, fontFace: 'Calibri', bold: true });
  s4.addText(c.desc, { x: x + 1.2, y: y + 0.7, w: 4.2, h: 0.8, fontSize: 11, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3 });
});

// ===== SLIDE 5: TARGET USERS & NEEDS =====
const s5 = pptx.addSlide();
slideTitle(s5, 'Target Users & Their Core Needs', 5);
pageContent(s5);
const users = [
  { role: 'HR', needs: ['Org-wide calibration view', 'Fairness & bias checks', 'Predictive decision support', 'Configurable thresholds'], color: C.accent },
  { role: 'Manager', needs: ['Direct report overview', 'Side-by-side comparison', 'Strengths & gap insights', 'Development suggestions'], color: C.muted },
  { role: 'Employee', needs: ['Personal score breakdown', 'Growth area identification', 'Clear badge rationale', 'Next-step guidance'], color: C.success },
];
users.forEach((u, i) => {
  const x = 0.7 + i * 4.2;
  card(s5, x, 1.5, 3.8, 5.2);
  s5.addShape(pptx.ShapeType.rect, { x, y: 1.5, w: 3.8, h: 0.06, fill: { color: u.color }, border: false });
  s5.addShape(pptx.ShapeType.roundRect, {
    x: x + 1.2, y: 1.8, w: 1.4, h: 1.4, fill: { color: u.color }, rectRadius: 0.7, border: false,
  });
  s5.addText(u.role, { x: x + 1.2, y: 1.8, w: 1.4, h: 1.4, fontSize: 16, color: C.white, fontFace: 'Calibri', bold: true, align: 'center', valign: 'middle' });
  u.needs.forEach((n, ni) => {
    s5.addText('→ ' + n, { x: x + 0.4, y: 3.5 + ni * 0.65, w: 3.0, h: 0.45, fontSize: 12, color: C.text, fontFace: 'Calibri' });
  });
});

// ===== SLIDE 6: INSIGHTS FRAMEWORK =====
const s6 = pptx.addSlide();
slideTitle(s6, 'What Insights Matter — Evaluation Framework', 6);
pageContent(s6);
const groups = [
  { label: 'Results &\nDelivery', metrics: ['Goal Achievement', 'Productivity'], x: 0.7, color: C.accent },
  { label: 'Behaviors &\nCompetencies', metrics: ['Quality of Work', 'Reliability'], x: 3.65, color: C.muted },
  { label: 'Evidence &\nFairness', metrics: ['Problem Solving', 'Integrity'], x: 6.6, color: C.success },
  { label: 'Growth &\nPotential', metrics: ['Peer Feedback', 'Communication'], x: 9.55, color: C.accentHover },
];
groups.forEach((g) => {
  card(s6, g.x, 1.5, 2.75, 5.2);
  s6.addShape(pptx.ShapeType.rect, { x: g.x, y: 1.5, w: 2.75, h: 0.06, fill: { color: g.color }, border: false });
  s6.addText(g.label, { x: g.x + 0.2, y: 1.75, w: 2.35, h: 0.7, fontSize: 13, color: g.color, fontFace: 'Calibri', bold: true, lineSpacingMultiple: 1.1 });
  g.metrics.forEach((m, mi) => {
    const my = 2.7 + mi * 1.7;
    s6.addShape(pptx.ShapeType.roundRect, {
      x: g.x + 0.3, y: my, w: 2.15, h: 1.3, fill: { color: C.bgSecondary }, rectRadius: 0.08, line: false,
    });
    s6.addText(m, { x: g.x + 0.4, y: my + 0.1, w: 1.95, h: 0.4, fontSize: 12, color: C.text, fontFace: 'Calibri', bold: true });
    s6.addText('Scale: 1 – 5', { x: g.x + 0.4, y: my + 0.55, w: 1.95, h: 0.3, fontSize: 10, color: C.muted, fontFace: 'Calibri' });
    s6.addShape(pptx.ShapeType.roundRect, {
      x: g.x + 0.4, y: my + 0.9, w: 1.0, h: 0.16, fill: { color: C.accent }, rectRadius: 0.08, border: false,
    });
    s6.addShape(pptx.ShapeType.roundRect, {
      x: g.x + 1.45, y: my + 0.9, w: 0.8, h: 0.16, fill: { color: C.accent }, rectRadius: 0.08, border: false, opacity: 0.3,
    });
  });
});

// ===== SLIDE 7: PROPOSED SOLUTION OVERVIEW =====
const s7 = pptx.addSlide();
slideTitle(s7, 'Proposed Solution — Three-Layer Architecture', 7);
pageContent(s7);
const layers = [
  { label: 'OUTCOME LAYER', desc: 'Weighted multi-dimensional scores, radar visualization, historical trends, strengths & improvement areas', x: 0.7, y: 1.5 },
  { label: 'INSIGHT LAYER', desc: 'Conditional eligibility badges with transparent rule logic, configurable thresholds, role-based views', x: 0.7, y: 3.5 },
  { label: 'DECISION-SUPPORT LAYER', desc: 'Logistic regression predictions for salary, bonus & promotion with top contributing factors', x: 0.7, y: 5.5 },
];
layers.forEach((l, i) => {
  const layerColors = [C.accent, C.muted, C.success];
  card(s7, l.x, l.y, 11.9, 1.6);
  s7.addShape(pptx.ShapeType.rect, { x: l.x, y: l.y, w: 0.2, h: 1.6, fill: { color: layerColors[i] }, border: false });
  s7.addText(l.label, { x: l.x + 0.5, y: l.y + 0.15, w: 4.5, h: 0.4, fontSize: 14, color: layerColors[i], fontFace: 'Calibri', bold: true });
  s7.addText(l.desc, { x: l.x + 0.5, y: l.y + 0.6, w: 11.0, h: 0.8, fontSize: 12, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3 });
  // arrow between layers
  if (i < 2) {
    s7.addText('▼', { x: 6.0, y: l.y + 1.55, w: 0.6, h: 0.3, fontSize: 14, color: C.faint, fontFace: 'Calibri', align: 'center' });
  }
});

// ===== SLIDE 8: PERFORMANCE OVERVIEW =====
const s8 = pptx.addSlide();
slideTitle(s8, 'Solution Screen — Performance Overview', 8);
pageContent(s8);
card(s8, 0.7, 1.5, 7.5, 5.2);
s8.addShape(pptx.ShapeType.roundRect, {
  x: 1.0, y: 1.8, w: 6.9, h: 4.6, fill: { color: C.bgOffset }, line: { color: C.border, width: 0.5, dashType: 'dash' },
});
s8.addText('[8-Metric Radar / Bar Chart]', { x: 1.0, y: 3.8, w: 6.9, h: 0.5, fontSize: 14, color: C.faint, fontFace: 'Calibri', align: 'center', valign: 'middle' });
s8.addShape(pptx.ShapeType.roundRect, {
  x: 1.0, y: 4.4, w: 6.9, h: 0.5, fill: { color: C.accent, opacity: 0.1 }, rectRadius: 0.08, line: false,
});
s8.addText('Replace with screen capture of the 8-metric radar chart showing all dimension scores for the current quarter', {
  x: 1.0, y: 4.4, w: 6.9, h: 0.5, fontSize: 9, color: C.muted, fontFace: 'Calibri', align: 'center', valign: 'middle',
});
// score card on right
card(s8, 8.5, 1.5, 4.1, 2.2);
s8.addText('OVERALL SCORE', { x: 8.8, y: 1.65, w: 3.5, h: 0.3, fontSize: 10, color: C.muted, fontFace: 'Calibri', bold: true });
s8.addText('4.2', { x: 8.8, y: 1.9, w: 2.0, h: 0.8, fontSize: 48, color: C.accent, fontFace: 'Calibri', bold: true });
s8.addText('/ 5.0', { x: 10.5, y: 2.3, w: 1.5, h: 0.4, fontSize: 16, color: C.faint, fontFace: 'Calibri' });
s8.addText('Cycle Q3 · Current Quarter', { x: 8.8, y: 2.8, w: 3.5, h: 0.3, fontSize: 10, color: C.muted, fontFace: 'Calibri' });
// badges
card(s8, 8.5, 3.9, 4.1, 2.8);
s8.addText('STATUS BADGES', { x: 8.8, y: 4.0, w: 3.5, h: 0.3, fontSize: 10, color: C.muted, fontFace: 'Calibri', bold: true });
s8.addShape(pptx.ShapeType.roundRect, {
  x: 8.8, y: 4.4, w: 2.2, h: 0.4, fill: { color: C.success }, rectRadius: 0.2, border: false,
});
s8.addText('Promotion Watchlist', { x: 8.8, y: 4.4, w: 2.2, h: 0.4, fontSize: 10, color: C.white, fontFace: 'Calibri', align: 'center', valign: 'middle', bold: true });
s8.addShape(pptx.ShapeType.roundRect, {
  x: 8.8, y: 4.9, w: 1.8, h: 0.4, fill: { color: C.success }, rectRadius: 0.2, border: false,
});
s8.addText('Bonus Eligible', { x: 8.8, y: 4.9, w: 1.8, h: 0.4, fontSize: 10, color: C.white, fontFace: 'Calibri', align: 'center', valign: 'middle', bold: true });
s8.addShape(pptx.ShapeType.rect, { x: 8.8, y: 5.7, w: 0.06, h: 0.5, fill: { color: C.accent }, border: false });
s8.addText('Callout: KPI cards show score, department, tenure, and status badges with hover tooltip explaining the "why" behind each badge', {
  x: 9.0, y: 5.55, w: 3.3, h: 0.8, fontSize: 9, color: C.accent, fontFace: 'Calibri', lineSpacingMultiple: 1.2,
});

// ===== SLIDE 9: HISTORICAL TREND =====
const s9 = pptx.addSlide();
slideTitle(s9, 'Solution Screen — Historical Trend', 9);
pageContent(s9);
card(s9, 0.7, 1.5, 11.9, 5.2);
s9.addShape(pptx.ShapeType.roundRect, {
  x: 1.0, y: 1.8, w: 11.3, h: 4.6, fill: { color: C.bgOffset }, line: { color: C.border, width: 0.5, dashType: 'dash' },
});
s9.addText('[6-Quarter Line Chart]', { x: 1.0, y: 3.6, w: 11.3, h: 0.5, fontSize: 14, color: C.faint, fontFace: 'Calibri', align: 'center', valign: 'middle' });
s9.addShape(pptx.ShapeType.roundRect, {
  x: 1.0, y: 4.3, w: 11.3, h: 0.5, fill: { color: C.accent, opacity: 0.1 }, rectRadius: 0.08, line: false,
});
s9.addText('Replace with screen capture of the trend line chart showing overall score progression across 6 quarters with metric toggle dropdown', {
  x: 1.0, y: 4.3, w: 11.3, h: 0.5, fontSize: 9, color: C.muted, fontFace: 'Calibri', align: 'center', valign: 'middle',
});
// feature callout
s9.addShape(pptx.ShapeType.pie, {
  x: 10.5, y: 1.65, w: 0.35, h: 0.35, fill: { color: C.accent },
});
s9.addText('Users can toggle individual metrics on/off and hover data points for exact quarter values', {
  x: 8.5, y: 5.8, w: 3.8, h: 0.7, fontSize: 10, color: C.accent, fontFace: 'Calibri', lineSpacingMultiple: 1.2, align: 'right',
});

// ===== SLIDE 10: ELIGIBILITY LOGIC =====
const s10 = pptx.addSlide();
slideTitle(s10, 'Solution Screen — Conditional Eligibility Logic', 10);
pageContent(s10);
card(s10, 0.7, 1.5, 7.0, 5.2);
s10.addShape(pptx.ShapeType.roundRect, {
  x: 1.0, y: 1.8, w: 6.4, h: 4.6, fill: { color: C.bgOffset }, line: { color: C.border, width: 0.5, dashType: 'dash' },
});
s10.addText('[Badges with Tooltip Explanations]', { x: 1.0, y: 3.8, w: 6.4, h: 0.5, fontSize: 14, color: C.faint, fontFace: 'Calibri', align: 'center', valign: 'middle' });
s10.addText('Replace with screen capture showing badges: Promotion Watchlist, Bonus Eligible, Needs Improvement Plan, Stable Performer', {
  x: 1.0, y: 4.4, w: 6.4, h: 0.5, fontSize: 9, color: C.muted, fontFace: 'Calibri', align: 'center', valign: 'middle',
});
// rules panel
card(s10, 8.0, 1.5, 4.6, 5.2);
s10.addText('ELIGIBILITY RULES', { x: 8.3, y: 1.65, w: 4.0, h: 0.3, fontSize: 11, color: C.accent, fontFace: 'Calibri', bold: true });
const rules = [
  { badge: 'Promotion Watchlist', rule: 'Overall score ≥ 4.2 for 2+ consecutive quarters AND Problem Solving ≥ 4 AND Integrity ≥ 4', color: C.success },
  { badge: 'Bonus Eligible', rule: 'Overall score ≥ 3.8 in current quarter AND Goal Achievement ≥ 4', color: C.success },
  { badge: 'Needs Improvement', rule: 'Overall score < 2.8 OR any metric ≤ 2 for 2+ consecutive quarters', color: C.warning },
  { badge: 'Stable Performer', rule: 'Default when none of the above thresholds are met', color: C.muted },
];
rules.forEach((r, i) => {
  const y = 2.15 + i * 1.1;
  s10.addShape(pptx.ShapeType.roundRect, {
    x: 8.3, y, w: 4.0, h: 0.9, fill: { color: C.bgSecondary }, rectRadius: 0.08, line: false,
  });
  s10.addShape(pptx.ShapeType.rect, { x: 8.3, y, w: 0.05, h: 0.9, fill: { color: r.color }, border: false });
  s10.addText(r.badge, { x: 8.5, y: y + 0.05, w: 3.6, h: 0.3, fontSize: 10, color: r.color, fontFace: 'Calibri', bold: true });
  s10.addText(r.rule, { x: 8.5, y: y + 0.35, w: 3.6, h: 0.5, fontSize: 8, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.1 });
});

// ===== SLIDE 11: PREDICTIVE INSIGHTS =====
const s11 = pptx.addSlide();
slideTitle(s11, 'Solution Screen — Predictive Insights', 11);
pageContent(s11);
card(s11, 0.7, 1.5, 7.5, 5.2);
s11.addShape(pptx.ShapeType.roundRect, {
  x: 1.0, y: 1.8, w: 6.9, h: 4.6, fill: { color: C.bgOffset }, line: { color: C.border, width: 0.5, dashType: 'dash' },
});
s11.addText('[Logistic Regression Output — 3 Probability Bars]', { x: 1.0, y: 3.6, w: 6.9, h: 0.5, fontSize: 14, color: C.faint, fontFace: 'Calibri', align: 'center', valign: 'middle' });
s11.addText('Replace with screen capture showing probability bars for Salary Adjustment, Bonus Award, and Promotion', {
  x: 1.0, y: 4.2, w: 6.9, h: 0.5, fontSize: 9, color: C.muted, fontFace: 'Calibri', align: 'center', valign: 'middle',
});
// right panel
card(s11, 8.5, 1.5, 4.1, 2.8);
s11.addText('PREDICTION OUTPUTS', { x: 8.8, y: 1.65, w: 3.5, h: 0.3, fontSize: 11, color: C.accent, fontFace: 'Calibri', bold: true });
const preds = [
  { label: 'Salary Adjustment', pct: '82%', barW: 3.2, color: C.success },
  { label: 'Bonus Award', pct: '68%', barW: 2.6, color: C.accent },
  { label: 'Promotion', pct: '45%', barW: 1.7, color: C.muted },
];
preds.forEach((p, i) => {
  const y = 2.15 + i * 0.75;
  s11.addText(p.label, { x: 8.8, y, w: 2.5, h: 0.3, fontSize: 10, color: C.text, fontFace: 'Calibri' });
  s11.addShape(pptx.ShapeType.roundRect, {
    x: 8.8, y: y + 0.3, w: 3.0, h: 0.2, fill: { color: C.bgOffset }, rectRadius: 0.1, border: false,
  });
  s11.addShape(pptx.ShapeType.roundRect, {
    x: 8.8, y: y + 0.3, w: p.barW, h: 0.2, fill: { color: p.color }, rectRadius: 0.1, border: false,
  });
  s11.addText(p.pct, { x: 11.0, y, w: 0.6, h: 0.5, fontSize: 12, color: p.color, fontFace: 'Calibri', bold: true });
});
// feature importance
card(s11, 8.5, 4.5, 4.1, 2.2);
s11.addText('TOP CONTRIBUTING FACTORS', { x: 8.8, y: 4.65, w: 3.5, h: 0.3, fontSize: 10, color: C.accent, fontFace: 'Calibri', bold: true });
s11.addText('"Goal Achievement and Problem Solving are the strongest drivers of promotion likelihood for this employee."', {
  x: 8.8, y: 5.05, w: 3.5, h: 0.8, fontSize: 9, color: C.muted, fontFace: 'Calibri', italic: true, lineSpacingMultiple: 1.3,
});
s11.addShape(pptx.ShapeType.rect, { x: 8.8, y: 5.85, w: 0.06, h: 0.5, fill: { color: C.accent }, border: false });
s11.addText('Decision-support estimate — not automatic. Actual outcomes determined by HR review.', {
  x: 9.0, y: 5.85, w: 3.3, h: 0.7, fontSize: 8, color: C.faint, fontFace: 'Calibri', lineSpacingMultiple: 1.1,
});

// ===== SLIDE 12: ROLE-BASED VIEWS =====
const s12 = pptx.addSlide();
slideTitle(s12, 'Role-Based Views — One Platform, Three Perspectives', 12);
pageContent(s12);
const views = [
  { role: 'HR', view: 'Calibration View', img: '[Org-wide table + rating distribution histogram]', desc: 'See all employees, filter by department, spot leniency/strictness patterns, jump into individual overviews', color: C.accent },
  { role: 'Manager', view: 'Team View', img: '[Direct report cards + comparison]', desc: 'List of direct reports with quick badges, one-click drill into any report, side-by-side comparison of two employees', color: C.muted },
  { role: 'Employee', view: 'My Performance', img: '[Simplified overview + next steps]', desc: 'Personal score, trend chart, strengths/gaps, badges with plain-language explanations, auto-generated development suggestions', color: C.success },
];
views.forEach((v, i) => {
  const x = 0.7 + i * 4.2;
  card(s12, x, 1.5, 3.8, 5.2);
  s12.addShape(pptx.ShapeType.rect, { x, y: 1.5, w: 3.8, h: 0.06, fill: { color: v.color }, border: false });
  s12.addText(v.role, { x: x + 0.3, y: 1.7, w: 3.0, h: 0.4, fontSize: 16, color: v.color, fontFace: 'Calibri', bold: true });
  s12.addText(v.view, { x: x + 0.3, y: 2.1, w: 3.0, h: 0.3, fontSize: 11, color: C.muted, fontFace: 'Calibri' });
  // placeholder
  s12.addShape(pptx.ShapeType.roundRect, {
    x: x + 0.3, y: 2.6, w: 3.2, h: 1.2, fill: { color: C.bgOffset }, line: { color: C.border, width: 0.5, dashType: 'dash' },
  });
  s12.addText(v.img, { x: x + 0.3, y: 2.6, w: 3.2, h: 1.2, fontSize: 9, color: C.faint, fontFace: 'Calibri', align: 'center', valign: 'middle' });
  s12.addText(v.desc, { x: x + 0.3, y: 4.0, w: 3.2, h: 2.2, fontSize: 11, color: C.text, fontFace: 'Calibri', lineSpacingMultiple: 1.4 });
});

// ===== SLIDE 13: IMPACT =====
const s13 = pptx.addSlide();
slideTitle(s13, 'Impact — Why This Matters', 13);
pageContent(s13);
const impacts = [
  { num: '86%', label: 'Fairer Decisions', desc: 'Multi-dimensional scoring reduces single-score bias and gives a complete picture of employee performance.' },
  { num: '2.5x', label: 'Faster Calibration', desc: 'HR teams can review, compare, and calibrate across departments in a single consolidated view.' },
  { num: '100%', label: 'Transparent Logic', desc: 'Every badge and prediction includes an explainable "why" — no black-box decisions.' },
  { num: '3', label: 'Role-Specific Views', desc: 'HR, managers, and employees each see the data that matters most to them in the right level of detail.' },
];
impacts.forEach((imp, i) => {
  const row = Math.floor(i / 2);
  const col = i % 2;
  const x = 0.7 + col * 6.2;
  const y = 1.5 + row * 2.7;
  card(s13, x, y, 5.8, 2.3);
  s13.addText(imp.num, { x: x + 0.3, y: y + 0.2, w: 1.5, h: 0.8, fontSize: 36, color: C.accent, fontFace: 'Calibri', bold: true });
  s13.addText(imp.label, { x: x + 1.8, y: y + 0.25, w: 3.6, h: 0.4, fontSize: 15, color: C.text, fontFace: 'Calibri', bold: true });
  s13.addText(imp.desc, { x: x + 1.8, y: y + 0.7, w: 3.6, h: 1.2, fontSize: 11, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.3 });
  s13.addShape(pptx.ShapeType.rect, { x: x + 0.3, y: y + 1.1, w: 1.2, h: 0.03, fill: { color: C.accent }, border: false });
});

// ===== SLIDE 14: NEXT STEPS =====
const s14 = pptx.addSlide();
s14.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: C.bg } });
s14.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.35, h: 7.5, fill: { color: C.accent }, border: false });
s14.addText('Next Steps', { x: 0.7, y: 1.0, w: 11, h: 0.6, fontSize: 26, color: C.accent, fontFace: 'Calibri', bold: true });
s14.addShape(pptx.ShapeType.rect, { x: 0.7, y: 1.65, w: 2.0, h: 0.04, fill: { color: C.accent }, border: false });

const steps = [
  { phase: 'PHASE 1', label: 'Pilot', desc: 'Test with 2–3 departments for 1 review cycle. Gather feedback on usability, badge logic, and prediction accuracy.' },
  { phase: 'PHASE 2', label: 'Calibration Test', desc: 'Expand to all managers for calibration. Validate rating distributions and manager leniency detection.' },
  { phase: 'PHASE 3', label: 'Full Rollout', desc: 'Company-wide launch with employee communication, training materials, and ongoing model refinement.' },
];
steps.forEach((st, i) => {
  const x = 0.7 + i * 4.2;
  card(s14, x, 2.2, 3.8, 3.0);
  s14.addShape(pptx.ShapeType.roundRect, {
    x: x + 0.3, y: 2.4, w: 1.0, h: 0.35, fill: { color: C.accent }, rectRadius: 0.17, border: false,
  });
  s14.addText(st.phase, { x: x + 0.3, y: 2.4, w: 1.0, h: 0.35, fontSize: 10, color: C.white, fontFace: 'Calibri', bold: true, align: 'center', valign: 'middle' });
  s14.addText(st.label, { x: x + 1.5, y: 2.4, w: 2.0, h: 0.35, fontSize: 14, color: C.text, fontFace: 'Calibri', bold: true, valign: 'middle' });
  s14.addText(st.desc, { x: x + 0.3, y: 3.0, w: 3.2, h: 1.8, fontSize: 11, color: C.muted, fontFace: 'Calibri', lineSpacingMultiple: 1.4 });
});

// CTA
s14.addShape(pptx.ShapeType.roundRect, {
  x: 3.5, y: 5.8, w: 6.3, h: 0.8, fill: { color: C.accent }, rectRadius: 0.1, border: false,
});
s14.addText('Try the clickable prototype: http://127.0.0.1:3000', {
  x: 3.5, y: 5.8, w: 6.3, h: 0.8, fontSize: 14, color: C.white, fontFace: 'Calibri', bold: true, align: 'center', valign: 'middle',
});
s14.addText('Thank you — Questions & Feedback Welcome', {
  x: 0.7, y: 6.8, w: 11, h: 0.4, fontSize: 12, color: C.muted, fontFace: 'Calibri', align: 'center',
});
footer(s14, 14);

// SAVE
pptx.writeFile({ fileName: 'C:\\Users\\sihle\\Desktop\\Empeo\\empeo_performance_insights_deck.pptx' })
  .then(() => console.log('Presentation saved successfully'))
  .catch(err => console.log('Error:', err));
