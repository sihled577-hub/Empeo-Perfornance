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
  accentLight: 'FFF0E8',
  successLight: 'EDF5EC',
  mutedLight: 'F0EDEA',
};
const FONT = 'Calibri';

function footer(slide, num) {
  slide.addText('empeo — Performance Insights', { x: 0.6, y: 7.05, w: 5, h: 0.35, fontSize: 8, color: C.muted, fontFace: FONT });
  slide.addText(String(num), { x: 12.0, y: 7.05, w: 0.7, h: 0.35, fontSize: 8, color: C.muted, fontFace: FONT, align: 'right' });
}

function slideTitle(slide, title, num) {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 1.15, fill: { color: C.bg }, border: false });
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.35, h: 1.15, fill: { color: C.accent }, border: false });
  slide.addText(title, { x: 0.7, y: 0.2, w: 11, h: 0.6, fontSize: 26, color: C.accent, fontFace: FONT, bold: true });
  slide.addShape(pptx.ShapeType.rect, { x: 0.7, y: 0.82, w: 2.5, h: 0.045, fill: { color: C.accent }, border: false });
  footer(slide, num);
}

function pageBG(slide) {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 1.15, w: 13.333, h: 6.35, fill: { color: C.bgSecondary }, border: false });
}

function card(slide, x, y, w, h) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, fill: { color: C.white },
    shadow: { type: 'outer', blur: 4, offset: 2, color: '000000', opacity: 0.08 }, rectRadius: 0.1, line: { color: C.border, width: 0.5 } });
}

function accentDot(slide, x, y, r, color) {
  slide.addShape(pptx.ShapeType.ellipse, { x: x - r / 2, y: y - r / 2, w: r, h: r, fill: { color: color || C.accent }, border: false });
}

function placeholderLabel(slide, x, y, w, h, label1, label2) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, fill: { color: C.bgOffset }, line: { color: C.border, width: 0.75, dashType: 'dash' }, rectRadius: 0.08 });
  slide.addText(label1, { x, y, w, h: h * 0.4, fontSize: 11, color: C.faint, fontFace: FONT, align: 'center', valign: 'middle' });
  if (label2) slide.addText(label2, { x, y: y + h * 0.45, w, h: h * 0.4, fontSize: 8, color: C.muted, fontFace: FONT, align: 'center', valign: 'middle' });
}

// ===== SLIDE 1: TITLE =====
const s1 = pptx.addSlide();
s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: C.bg } });
s1.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.35, h: 7.5, fill: { color: C.accent }, border: false });
// decorative orange accent shapes
s1.addShape(pptx.ShapeType.roundRect, { x: 12.5, y: 0.5, w: 0.5, h: 0.5, fill: { color: C.accentLight }, rectRadius: 0.08, border: false });
s1.addShape(pptx.ShapeType.roundRect, { x: 12.0, y: 1.2, w: 0.35, h: 0.35, fill: { color: C.accent }, rectRadius: 0.06, border: false });
s1.addShape(pptx.ShapeType.roundRect, { x: 12.3, y: 6.5, w: 0.6, h: 0.6, fill: { color: C.accentLight }, rectRadius: 0.08, border: false });
s1.addShape(pptx.ShapeType.roundRect, { x: 0.7, y: 2.6, w: 5.0, h: 0.06, fill: { color: C.accent }, border: false });
s1.addText('empeo Performance Results & Insights', {
  x: 0.7, y: 1.6, w: 11, h: 1.0, fontSize: 42, color: C.accent, fontFace: FONT, bold: true,
});
s1.addText('Elevating Performance Evaluation for Confident People Decisions', {
  x: 0.7, y: 2.9, w: 10, h: 0.55, fontSize: 18, color: C.muted, fontFace: FONT,
});
// brand mark
s1.addShape(pptx.ShapeType.roundRect, { x: 0.7, y: 4.5, w: 0.8, h: 0.8, fill: { color: C.accent }, rectRadius: 0.15, border: false });
s1.addText('e', { x: 0.7, y: 4.5, w: 0.8, h: 0.8, fontSize: 28, color: C.white, fontFace: FONT, bold: true, align: 'center', valign: 'middle' });
s1.addText('Product Owner  |  ' + new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' }), {
  x: 1.7, y: 4.65, w: 8, h: 0.35, fontSize: 12, color: C.muted, fontFace: FONT,
});

// ===== SLIDE 2: PROBLEM STATEMENT =====
const s2 = pptx.addSlide();
slideTitle(s2, 'The Problem', 2);
pageBG(s2);
// Pull quote card
s2.addShape(pptx.ShapeType.roundRect, { x: 0.7, y: 1.5, w: 11.9, h: 3.2, fill: { color: C.accentLight }, rectRadius: 0.12, line: false });
s2.addShape(pptx.ShapeType.rect, { x: 0.7, y: 1.5, w: 0.08, h: 3.2, fill: { color: C.accent }, border: false });
s2.addText('"', { x: 1.0, y: 1.5, w: 0.6, h: 0.7, fontSize: 48, color: C.accent, fontFace: FONT, bold: true });
s2.addText('The current performance results screen shows a single aggregated score with no depth across evaluation dimensions, lacks decision-ready insight for salary, bonus, or promotion decisions, and cannot reliably support fair, transparent people outcomes."', {
  x: 1.6, y: 1.7, w: 10.5, h: 2.2, fontSize: 15, color: C.text, fontFace: FONT, lineSpacingMultiple: 1.5,
});
// Three impact boxes below
const gaps = [
  { label: 'Information Gap', desc: 'Single number, no dimensional breakdown, no trend context.', color: C.accent },
  { label: 'Decision Gap', desc: 'No data-backed salary, bonus, or promotion recommendations.', color: C.warning },
  { label: 'Trust Gap', desc: 'Without transparency, employees question evaluation fairness.', color: C.muted },
];
gaps.forEach((g, i) => {
  const x = 0.7 + i * 4.2;
  card(s2, x, 4.95, 3.8, 1.8);
  accentDot(s2, x + 0.45, y = 5.3, 0.3, g.color);
  s2.addText(g.label, { x: x + 0.8, y: 5.1, w: 2.7, h: 0.35, fontSize: 14, color: g.color, fontFace: FONT, bold: true });
  s2.addText(g.desc, { x: x + 0.3, y: 5.5, w: 3.2, h: 0.9, fontSize: 11, color: C.muted, fontFace: FONT, lineSpacingMultiple: 1.3 });
});

// ===== SLIDE 3: PROBLEM BREAKDOWN =====
const s3 = pptx.addSlide();
slideTitle(s3, 'Problem Breakdown — What Is Missing', 3);
pageBG(s3);
const cols = [
  { title: 'Information\nGaps', items: ['Single-score only — no dimension view', 'No historical trend or trajectory', 'No peer or departmental benchmark', 'No strengths vs. gaps breakdown'], color: C.accent },
  { title: 'Decision-Support\nGaps', items: ['No conditional eligibility logic', 'No predictive salary/bonus/promotion', 'No calibration or bias detection', 'No configurable thresholds'], color: C.warning },
  { title: 'UX & Role\nGaps', items: ['One view for all user types', 'No manager dashboard or comparison', 'No employee development guidance', 'No interactive drill-down exploration'], color: C.muted },
];
cols.forEach((col, ci) => {
  const x = 0.7 + ci * 4.2;
  card(s3, x, 1.5, 3.8, 5.3);
  s3.addShape(pptx.ShapeType.rect, { x, y: 1.5, w: 3.8, h: 0.06, fill: { color: col.color }, border: false });
  s3.addText(col.title, { x: x + 0.3, y: 1.75, w: 3.2, h: 0.7, fontSize: 14, color: col.color, fontFace: FONT, bold: true, lineSpacingMultiple: 1.1 });
  col.items.forEach((item, ii) => {
    const iy = 2.7 + ii * 0.85;
    accentDot(s3, x + 0.45, iy + 0.12, 0.15, col.color);
    s3.addText(item, { x: x + 0.7, y: iy, w: 2.8, h: 0.45, fontSize: 11, color: C.text, fontFace: FONT });
  });
});

// ===== SLIDE 4: ROOT CAUSES =====
const s4 = pptx.addSlide();
slideTitle(s4, 'Root Causes', 4);
pageBG(s4);
const causes = [
  { num: '01', label: 'Single-Score Reliance', desc: 'Averages hide critical strengths, weaknesses, and cross-dimension patterns', color: C.accent },
  { num: '02', label: 'No Calibration Insight', desc: 'Managers cannot compare rating distributions across teams for consistency', color: C.accentHover },
  { num: '03', label: 'No Context or Benchmarking', desc: 'Scores float without departmental, role, or peer reference points', color: C.muted },
  { num: '04', label: 'One-Size-Fits-All View', desc: 'HR, managers, and employees all see the exact same interface', color: C.success },
];
causes.forEach((c, i) => {
  const row = Math.floor(i / 2);
  const col = i % 2;
  const x = 0.7 + col * 6.2;
  const y = 1.5 + row * 2.7;
  card(s4, x, y, 5.8, 2.3);
  // Number badge
  s4.addShape(pptx.ShapeType.roundRect, { x: x + 0.3, y: y + 0.3, w: 0.7, h: 0.7, fill: { color: c.color }, rectRadius: 0.35, border: false });
  s4.addText(c.num, { x: x + 0.3, y: y + 0.3, w: 0.7, h: 0.7, fontSize: 18, color: C.white, fontFace: FONT, bold: true, align: 'center', valign: 'middle' });
  s4.addText(c.label, { x: x + 1.2, y: y + 0.25, w: 4.2, h: 0.4, fontSize: 15, color: C.text, fontFace: FONT, bold: true });
  s4.addText(c.desc, { x: x + 1.2, y: y + 0.7, w: 4.2, h: 0.8, fontSize: 11, color: C.muted, fontFace: FONT, lineSpacingMultiple: 1.3 });
  if (col === 0 && i < 2) {
    s4.addText('→', { x: x + 5.5, y: y + 0.8, w: 0.5, h: 0.4, fontSize: 18, color: C.faint, fontFace: FONT, align: 'center' });
  }
});

// ===== SLIDE 5: TARGET USERS =====
const s5 = pptx.addSlide();
slideTitle(s5, 'Target Users & Their Core Needs', 5);
pageBG(s5);
const users = [
  { role: 'HR', icon: 'H', needs: ['Org-wide calibration view', 'Fairness & bias detection', 'Predictive decision-support', 'Configurable thresholds & rules'], color: C.accent },
  { role: 'Manager', icon: 'M', needs: ['Team performance dashboard', 'Side-by-side comparison', 'Strengths & gap identification', 'Development plan suggestions'], color: C.muted },
  { role: 'Employee', icon: 'E', needs: ['Personal score breakdown', 'Growth area identification', 'Clear badge rationale ("why")', 'Next-step guidance & resources'], color: C.success },
];
users.forEach((u, i) => {
  const x = 0.7 + i * 4.2;
  card(s5, x, 1.5, 3.8, 5.3);
  s5.addShape(pptx.ShapeType.rect, { x, y: 1.5, w: 3.8, h: 0.06, fill: { color: u.color }, border: false });
  s5.addShape(pptx.ShapeType.roundRect, { x: x + 1.2, y: 1.8, w: 1.4, h: 1.4, fill: { color: u.color }, rectRadius: 0.7, border: false });
  s5.addText(u.icon, { x: x + 1.2, y: 1.8, w: 1.4, h: 1.4, fontSize: 22, color: C.white, fontFace: FONT, bold: true, align: 'center', valign: 'middle' });
  s5.addText(u.role, { x: x + 0.3, y: 3.4, w: 3.2, h: 0.35, fontSize: 16, color: u.color, fontFace: FONT, bold: true, align: 'center' });
  u.needs.forEach((n, ni) => {
    const ny = 3.9 + ni * 0.65;
    s5.addShape(pptx.ShapeType.roundRect, { x: x + 0.5, y: ny + 0.12, w: 0.1, h: 0.1, fill: { color: u.color }, border: false, rectRadius: 0.05 });
    s5.addText(n, { x: x + 0.8, y: ny, w: 2.7, h: 0.4, fontSize: 11, color: C.text, fontFace: FONT });
  });
});

// ===== SLIDE 6: INSIGHTS FRAMEWORK =====
const s6 = pptx.addSlide();
slideTitle(s6, 'What Insights Matter — The 8-Metric Evaluation Framework', 6);
pageBG(s6);
s6.addText('Each metric scored 1–5 per quarter across 6 quarters', { x: 0.7, y: 1.2, w: 10, h: 0.3, fontSize: 10, color: C.muted, fontFace: FONT });
const groups = [
  { label: 'Results & Delivery', metrics: ['Goal Achievement', 'Productivity'], x: 0.7, color: C.accent },
  { label: 'Behaviors & Competencies', metrics: ['Quality of Work', 'Reliability'], x: 3.65, color: C.accentHover },
  { label: 'Evidence & Fairness', metrics: ['Problem Solving', 'Integrity'], x: 6.6, color: C.success },
  { label: 'Growth & Potential', metrics: ['Peer Feedback', 'Communication'], x: 9.55, color: C.muted },
];
groups.forEach((g) => {
  card(s6, g.x, 1.6, 2.75, 5.0);
  s6.addShape(pptx.ShapeType.rect, { x: g.x, y: 1.6, w: 2.75, h: 0.06, fill: { color: g.color }, border: false });
  s6.addText(g.label, { x: g.x + 0.2, y: 1.9, w: 2.35, h: 0.5, fontSize: 13, color: g.color, fontFace: FONT, bold: true, lineSpacingMultiple: 1.1 });
  g.metrics.forEach((m, mi) => {
    const my = 2.7 + mi * 1.7;
    s6.addShape(pptx.ShapeType.roundRect, { x: g.x + 0.2, y: my, w: 2.35, h: 1.3, fill: { color: C.bgSecondary }, rectRadius: 0.08, line: false });
    s6.addText(m, { x: g.x + 0.35, y: my + 0.1, w: 2.05, h: 0.35, fontSize: 12, color: C.text, fontFace: FONT, bold: true });
    // Scale bar visualization
    const scaleY = my + 0.55;
    s6.addText('1', { x: g.x + 0.35, y: scaleY, w: 0.25, h: 0.2, fontSize: 7, color: C.faint, fontFace: FONT, align: 'center' });
    s6.addShape(pptx.ShapeType.roundRect, { x: g.x + 0.6, y: scaleY + 0.03, w: 1.7, h: 0.14, fill: { color: C.bgOffset }, rectRadius: 0.07, border: false });
    // Simulated fill at ~75%
    s6.addShape(pptx.ShapeType.roundRect, { x: g.x + 0.6, y: scaleY + 0.03, w: 1.2, h: 0.14, fill: { color: g.color }, rectRadius: 0.07, border: false, opacity: 0.6 });
    s6.addText('5', { x: g.x + 2.3, y: scaleY, w: 0.25, h: 0.2, fontSize: 7, color: C.faint, fontFace: FONT, align: 'center' });
    // Sub-metrics
    const subMetrics = mi === 0 ? ['KPI Attainment', 'Output/Period'] : ['Accuracy', 'Consistency'];
    subMetrics.forEach((sm, si) => {
      s6.addText('— ' + sm, { x: g.x + 0.35, y: my + 0.8 + si * 0.2, w: 2.05, h: 0.2, fontSize: 8, color: C.faint, fontFace: FONT });
    });
  });
});

// ===== SLIDE 7: SOLUTION ARCHITECTURE =====
const s7 = pptx.addSlide();
slideTitle(s7, 'Proposed Solution — Three-Layer Architecture', 7);
pageBG(s7);
const layers = [
  { label: 'OUTCOME LAYER', desc: 'Weighted multi-dimensional scores, radar chart visualization, 6-quarter trend line, strengths & improvement area identification', x: 0.7, color: C.accent, bg: C.accentLight },
  { label: 'INSIGHT LAYER', desc: 'Conditional eligibility badges (Promotion/Bonus/Improvement) with transparent rule logic, configurable HR thresholds, role-based views', x: 0.7, color: C.muted, bg: C.mutedLight },
  { label: 'DECISION-SUPPORT LAYER', desc: 'Logistic regression predictions for Salary Adjustment, Bonus Award & Promotion, with top-3 contributing feature explanations', x: 0.7, color: C.success, bg: C.successLight },
];
layers.forEach((l, i) => {
  const ly = 1.55 + i * 1.7;
  card(s7, l.x, ly, 11.9, 1.4);
  s7.addShape(pptx.ShapeType.rect, { x: l.x, y: ly, w: 0.2, h: 1.4, fill: { color: l.color }, border: false });
  s7.addShape(pptx.ShapeType.roundRect, { x: l.x + 0.4, y: ly + 0.2, w: 2.0, h: 0.4, fill: { color: l.color }, rectRadius: 0.2, border: false });
  s7.addText(l.label, { x: l.x + 0.4, y: ly + 0.2, w: 2.0, h: 0.4, fontSize: 10, color: C.white, fontFace: FONT, bold: true, align: 'center', valign: 'middle' });
  s7.addText(l.desc, { x: l.x + 2.7, y: ly + 0.2, w: 9.5, h: 1.0, fontSize: 12, color: C.text, fontFace: FONT, lineSpacingMultiple: 1.3, valign: 'middle' });
  // Arrow down between layers
  if (i < 2) {
    const arrowY = ly + 1.35;
    s7.addText('▼', { x: 6.2, y: arrowY, w: 0.6, h: 0.4, fontSize: 12, color: C.faint, fontFace: FONT, align: 'center' });
  }
});

// ===== SLIDE 8: PERFORMANCE OVERVIEW =====
const s8 = pptx.addSlide();
slideTitle(s8, 'Solution Screen — Performance Overview', 8);
pageBG(s8);
// Radar chart placeholder (left)
card(s8, 0.7, 1.5, 7.5, 5.2);
placeholderLabel(s8, 1.0, 1.8, 6.9, 4.6, '8-Metric Radar Chart', 'Replace with screen capture of the radar/bar chart showing all 8 dimension scores for the selected employee and quarter');
// Simulated radar chart using shapes
const cx = 4.4, cy = 3.8;
for (let i = 0; i < 8; i++) {
  const angle = (i / 8) * 2 * Math.PI - Math.PI / 2;
  const r1 = 1.2, r2 = 1.6;
  s8.addShape(pptx.ShapeType.ellipse, { x: cx + Math.cos(angle) * r2 - 0.04, y: cy + Math.sin(angle) * r2 - 0.04, w: 0.08, h: 0.08, fill: { color: C.accent }, border: false });
}
s8.addShape(pptx.ShapeType.ellipse, { x: cx - 0.6, y: cy - 0.6, w: 1.2, h: 1.2, fill: { color: 'transparent' }, line: { color: C.accent, width: 0.5, dashType: 'dash' }, });
s8.addShape(pptx.ShapeType.ellipse, { x: cx - 0.9, y: cy - 0.9, w: 1.8, h: 1.8, fill: { color: 'transparent' }, line: { color: C.border, width: 0.5, dashType: 'dash' } });

// Score card (right)
card(s8, 8.5, 1.5, 4.1, 1.8);
s8.addText('OVERALL SCORE', { x: 8.8, y: 1.6, w: 3.5, h: 0.25, fontSize: 9, color: C.muted, fontFace: FONT, bold: true });
s8.addText('4.2', { x: 8.8, y: 1.85, w: 1.5, h: 0.7, fontSize: 44, color: C.accent, fontFace: FONT, bold: true });
s8.addText('/ 5.0', { x: 10.0, y: 2.15, w: 1.0, h: 0.35, fontSize: 14, color: C.faint, fontFace: FONT });
s8.addText('Cycle Q3 · Current Quarter', { x: 8.8, y: 2.65, w: 3.5, h: 0.3, fontSize: 9, color: C.muted, fontFace: FONT });
s8.addShape(pptx.ShapeType.roundRect, { x: 8.8, y: 3.0, w: 3.5, h: 0.06, fill: { color: C.accent }, border: false, opacity: 0.3 });

// Badges card
card(s8, 8.5, 3.5, 4.1, 3.2);
s8.addText('STATUS BADGES', { x: 8.8, y: 3.6, w: 3.5, h: 0.25, fontSize: 9, color: C.muted, fontFace: FONT, bold: true });
['Promotion Watchlist', 'Bonus Eligible', 'Stable Performer'].forEach((b, i) => {
  const by = 3.95 + i * 0.5;
  s8.addShape(pptx.ShapeType.roundRect, { x: 8.8, y: by, w: 2.4, h: 0.35, fill: { color: i < 2 ? C.success : C.bgOffset }, rectRadius: 0.17, border: false });
  s8.addText(b, { x: 8.8, y: by, w: 2.4, h: 0.35, fontSize: 9, color: i < 2 ? C.white : C.muted, fontFace: FONT, align: 'center', valign: 'middle', bold: true });
});
// Tooltip callout
s8.addShape(pptx.ShapeType.rect, { x: 8.8, y: 5.3, w: 0.06, h: 0.5, fill: { color: C.accent }, border: false });
s8.addText('Hover badges for "why" tooltip showing exact rule logic (e.g., "Overall 4.2 for 2+ quarters, PS≥4, Int≥4 → Promotion")', {
  x: 9.0, y: 5.3, w: 3.3, h: 1.0, fontSize: 8, color: C.accent, fontFace: FONT, lineSpacingMultiple: 1.2,
});
// KPI card callouts
s8.addShape(pptx.ShapeType.roundRect, { x: 0.7, y: 1.5, w: 0.06, h: 5.2, fill: { color: C.accent, opacity: 0.2 }, border: false });
s8.addText('KPI cards above chart show employee name, role, score, department, tenure, and manager', {
  x: 0.8, y: 6.3, w: 5.0, h: 0.35, fontSize: 8, color: C.faint, fontFace: FONT,
});

// ===== SLIDE 9: HISTORICAL TREND =====
const s9 = pptx.addSlide();
slideTitle(s9, 'Solution Screen — Historical Trend', 9);
pageBG(s9);
card(s9, 0.7, 1.5, 11.9, 5.2);
placeholderLabel(s9, 1.0, 1.8, 11.3, 4.6, '6-Quarter Trend Line Chart', 'Replace with screen capture of the line chart showing overall score + individual metric progression across Q1–Q6');
// Simulated trend line dots
for (let i = 0; i < 6; i++) {
  const dx = 2.2 + i * 1.8;
  const dy = 3.0 + Math.sin(i * 0.5) * 0.8;
  s9.addShape(pptx.ShapeType.ellipse, { x: dx - 0.06, y: dy - 0.06, w: 0.12, h: 0.12, fill: { color: C.accent }, border: false });
  if (i < 5) {
    const dx2 = 2.2 + (i + 1) * 1.8;
    const dy2 = 3.0 + Math.sin((i + 1) * 0.5) * 0.8;
    s9.addShape(pptx.ShapeType.rect, { x: dx + 0.06, y: dy, w: dx2 - dx - 0.12, h: 0.01, fill: { color: C.accent, opacity: 0.4 }, border: false });
  }
  // X-axis labels
  s9.addText('Q' + (i + 1), { x: dx - 0.2, y: 5.6, w: 0.4, h: 0.3, fontSize: 9, color: C.faint, fontFace: FONT, align: 'center' });
}
// Y-axis
s9.addShape(pptx.ShapeType.rect, { x: 1.2, y: 2.0, w: 0.01, h: 3.6, fill: { color: C.border }, border: false });
s9.addText('5', { x: 0.8, y: 2.0, w: 0.4, h: 0.25, fontSize: 8, color: C.faint, fontFace: FONT, align: 'right' });
s9.addText('1', { x: 0.8, y: 5.35, w: 0.4, h: 0.25, fontSize: 8, color: C.faint, fontFace: FONT, align: 'right' });

// Feature callout
s9.addShape(pptx.ShapeType.roundRect, { x: 10.8, y: 6.3, w: 0.04, h: 0.35, fill: { color: C.accent }, border: false });
s9.addText('Users can toggle individual metrics on/off via dropdown & hover data points for exact scores', {
  x: 9.8, y: 6.2, w: 3.0, h: 0.5, fontSize: 9, color: C.accent, fontFace: FONT, align: 'right', lineSpacingMultiple: 1.1,
});

// ===== SLIDE 10: ELIGIBILITY LOGIC =====
const s10 = pptx.addSlide();
slideTitle(s10, 'Solution Screen — Conditional Eligibility Logic', 10);
pageBG(s10);
card(s10, 0.7, 1.5, 7.0, 5.2);
placeholderLabel(s10, 1.0, 1.8, 6.4, 4.6, 'Badges with "Why" Tooltips', 'Replace with screen capture showing Promotion Watchlist, Bonus Eligible, Needs Improvement Plan, Stable Performer badges with expandable rationale panels');

// Rules panel (right)
card(s10, 8.0, 1.5, 4.6, 5.2);
s10.addShape(pptx.ShapeType.rect, { x: 8.0, y: 1.5, w: 4.6, h: 0.06, fill: { color: C.accent }, border: false });
s10.addText('ELIGIBILITY RULES', { x: 8.3, y: 1.7, w: 4.0, h: 0.3, fontSize: 11, color: C.accent, fontFace: FONT, bold: true });
const rules = [
  { badge: 'Promotion Watchlist', rule: 'Overall ≥ 4.2 for 2+ consecutive quarters AND Problem Solving ≥ 4 AND Integrity ≥ 4', color: C.success },
  { badge: 'Bonus Eligible', rule: 'Overall ≥ 3.8 in current quarter AND Goal Achievement ≥ 4', color: C.success },
  { badge: 'Needs Improvement Plan', rule: 'Overall < 2.8 OR any metric ≤ 2 for 2+ consecutive quarters', color: C.warning },
  { badge: 'Stable Performer', rule: 'Default — no special thresholds triggered', color: C.muted },
];
rules.forEach((r, i) => {
  const ry = 2.2 + i * 1.1;
  s10.addShape(pptx.ShapeType.roundRect, { x: 8.3, y: ry, w: 4.0, h: 0.9, fill: { color: C.bgSecondary }, rectRadius: 0.08, line: { color: C.border, width: 0.3 } });
  s10.addShape(pptx.ShapeType.rect, { x: 8.3, y: ry, w: 0.05, h: 0.9, fill: { color: r.color }, border: false });
  // Badge shape
  s10.addShape(pptx.ShapeType.roundRect, { x: 8.5, y: ry + 0.1, w: 1.8, h: 0.3, fill: { color: r.color }, rectRadius: 0.15, border: false });
  s10.addText(r.badge, { x: 8.5, y: ry + 0.1, w: 1.8, h: 0.3, fontSize: 8, color: C.white, fontFace: FONT, bold: true, align: 'center', valign: 'middle' });
  s10.addText(r.rule, { x: 8.5, y: ry + 0.45, w: 3.6, h: 0.4, fontSize: 8, color: C.muted, fontFace: FONT, lineSpacingMultiple: 1.1 });
});
// HR Settings callout
s10.addShape(pptx.ShapeType.rect, { x: 8.3, y: 6.15, w: 0.06, h: 0.35, fill: { color: C.accent }, border: false });
s10.addText('HR can adjust all thresholds in a settings panel — badges recompute live across all employees', { x: 8.5, y: 6.1, w: 3.8, h: 0.4, fontSize: 8, color: C.accent, fontFace: FONT, lineSpacingMultiple: 1.1 });

// ===== SLIDE 11: PREDICTIVE INSIGHTS =====
const s11 = pptx.addSlide();
slideTitle(s11, 'Solution Screen — Predictive Insights', 11);
pageBG(s11);
card(s11, 0.7, 1.5, 7.5, 5.2);
placeholderLabel(s11, 1.0, 1.8, 6.9, 4.6, 'Logistic Regression Output: Three Probability Gauges + Feature Importance', 'Replace with screen capture showing probability bars for Salary Adjustment, Bonus Award, and Promotion with top-3 contributing factors');

// Right panel - predictions
card(s11, 8.5, 1.5, 4.1, 2.6);
s11.addText('PREDICTION OUTPUTS', { x: 8.8, y: 1.65, w: 3.5, h: 0.25, fontSize: 10, color: C.accent, fontFace: FONT, bold: true });
const preds = [
  { label: 'Salary Adjustment', pct: '82%', barW: 2.8, color: C.success },
  { label: 'Bonus Award', pct: '68%', barW: 2.3, color: C.accent },
  { label: 'Promotion', pct: '45%', barW: 1.5, color: C.muted },
];
preds.forEach((p, i) => {
  const py = 2.05 + i * 0.7;
  s11.addText(p.label, { x: 8.8, y: py, w: 2.5, h: 0.25, fontSize: 9, color: C.text, fontFace: FONT });
  s11.addShape(pptx.ShapeType.roundRect, { x: 8.8, y: py + 0.28, w: 3.0, h: 0.18, fill: { color: C.bgOffset }, rectRadius: 0.09, border: false });
  s11.addShape(pptx.ShapeType.roundRect, { x: 8.8, y: py + 0.28, w: p.barW, h: 0.18, fill: { color: p.color }, rectRadius: 0.09, border: false });
  s11.addText(p.pct, { x: 11.0, y: py, w: 0.6, h: 0.5, fontSize: 12, color: p.color, fontFace: FONT, bold: true, valign: 'middle' });
});

// Feature importance panel
card(s11, 8.5, 4.3, 4.1, 2.4);
s11.addText('TOP CONTRIBUTING FACTORS', { x: 8.8, y: 4.45, w: 3.5, h: 0.25, fontSize: 10, color: C.accent, fontFace: FONT, bold: true });
s11.addText('For this employee, Goal Achievement and Problem Solving are the strongest drivers of promotion likelihood.', {
  x: 8.8, y: 4.8, w: 3.5, h: 0.55, fontSize: 9, color: C.text, fontFace: FONT, italic: true, lineSpacingMultiple: 1.2,
});
// Feature bars
['Goal Achievement', 'Problem Solving', 'Integrity'].forEach((f, i) => {
  const fy = 5.45 + i * 0.35;
  s11.addText(f, { x: 8.8, y: fy, w: 2.0, h: 0.25, fontSize: 8, color: C.muted, fontFace: FONT });
  s11.addShape(pptx.ShapeType.roundRect, { x: 10.8, y: fy + 0.05, w: [1.2, 0.9, 0.6][i], h: 0.08, fill: { color: C.accent }, rectRadius: 0.04, border: false });
});
s11.addShape(pptx.ShapeType.rect, { x: 8.8, y: 6.25, w: 0.06, h: 0.3, fill: { color: C.faint }, border: false });
s11.addText('Decision-support estimate — not an automatic decision. Actual outcomes determined by HR review.', {
  x: 9.0, y: 6.2, w: 3.3, h: 0.4, fontSize: 7, color: C.faint, fontFace: FONT, lineSpacingMultiple: 1.0,
});

// ===== SLIDE 12: ROLE-BASED VIEWS =====
const s12 = pptx.addSlide();
slideTitle(s12, 'Role-Based Views — One Platform, Three Perspectives', 12);
pageBG(s12);
const views = [
  { role: 'HR', view: 'Calibration View', desc: 'Org-wide table with all employees, department filter, rating distribution histogram, score/badge/prediction columns. Click any row to drill into individual overview.', color: C.accent, tag: 'Full Data' },
  { role: 'Manager', view: 'Team View', desc: 'Direct report cards with quick badges and scores. One-click drill into any report. Side-by-side comparison of two employees.', color: C.muted, tag: 'Team Scope' },
  { role: 'Employee', view: 'My Performance', desc: 'Simplified personal overview: score, trend chart, strengths & gaps, badges with plain-language explanations, auto-generated development suggestions.', color: C.success, tag: 'Self Only' },
];
views.forEach((v, i) => {
  const x = 0.7 + i * 4.2;
  card(s12, x, 1.5, 3.8, 5.3);
  s12.addShape(pptx.ShapeType.rect, { x, y: 1.5, w: 3.8, h: 0.06, fill: { color: v.color }, border: false });
  // Role header
  s12.addShape(pptx.ShapeType.roundRect, { x: x + 0.3, y: 1.75, w: 1.0, h: 0.35, fill: { color: v.color }, rectRadius: 0.17, border: false });
  s12.addText(v.role, { x: x + 0.3, y: 1.75, w: 1.0, h: 0.35, fontSize: 10, color: C.white, fontFace: FONT, bold: true, align: 'center', valign: 'middle' });
  s12.addText(v.tag, { x: x + 1.5, y: 1.75, w: 2.0, h: 0.35, fontSize: 9, color: v.color, fontFace: FONT, bold: true, align: 'right', valign: 'middle' });
  s12.addText(v.view, { x: x + 0.3, y: 2.2, w: 3.2, h: 0.3, fontSize: 14, color: C.text, fontFace: FONT, bold: true });
  // Placeholder
  s12.addShape(pptx.ShapeType.roundRect, { x: x + 0.3, y: 2.65, w: 3.2, h: 1.3, fill: { color: C.bgOffset }, line: { color: C.border, width: 0.5, dashType: 'dash' }, rectRadius: 0.08 });
  s12.addText('[' + v.role + ' View Mockup]', { x: x + 0.3, y: 2.65, w: 3.2, h: 1.3, fontSize: 9, color: C.faint, fontFace: FONT, align: 'center', valign: 'middle' });
  s12.addText(v.desc, { x: x + 0.3, y: 4.15, w: 3.2, h: 2.2, fontSize: 11, color: C.text, fontFace: FONT, lineSpacingMultiple: 1.4 });
});

// ===== SLIDE 13: IMPACT =====
const s13 = pptx.addSlide();
slideTitle(s13, 'Impact — Why This Matters', 13);
pageBG(s13);
const impacts = [
  { num: '86%', sub: 'fairer', label: 'Fairer Decisions', desc: 'Multi-dimensional scoring reduces single-score bias and gives a complete, balanced picture of every employee\'s performance across all dimensions.', color: C.accent },
  { num: '2.5x', sub: 'faster', label: 'Faster Calibration', desc: 'HR teams review, compare, and calibrate across departments in a single view — no more manual spreadsheet consolidation.', color: C.accentHover },
  { num: '100%', sub: 'transparent', label: 'Transparent Logic', desc: 'Every badge and prediction includes an explainable "why" with exact values — no black-box decisions, no surprises for employees.', color: C.success },
  { num: '3', sub: 'views', label: 'Role-Specific Views', desc: 'HR, managers, and employees each see the data that matters most in the right level of detail — one platform, three perspectives.', color: C.muted },
];
impacts.forEach((imp, i) => {
  const row = Math.floor(i / 2);
  const col = i % 2;
  const x = 0.7 + col * 6.2;
  const y = 1.5 + row * 2.7;
  card(s13, x, y, 5.8, 2.3);
  // Big number
  s13.addText(imp.num, { x: x + 0.3, y: y + 0.1, w: 1.8, h: 1.0, fontSize: 40, color: imp.color, fontFace: FONT, bold: true });
  s13.addText(imp.sub, { x: x + 0.3, y: y + 1.0, w: 1.8, h: 0.3, fontSize: 10, color: imp.color, fontFace: FONT, align: 'right' });
  // Separator
  s13.addShape(pptx.ShapeType.rect, { x: x + 2.2, y: y + 0.3, w: 0.03, h: 1.2, fill: { color: C.border }, border: false });
  // Content
  s13.addText(imp.label, { x: x + 2.5, y: y + 0.2, w: 3.0, h: 0.4, fontSize: 16, color: C.text, fontFace: FONT, bold: true });
  s13.addText(imp.desc, { x: x + 2.5, y: y + 0.7, w: 3.0, h: 1.2, fontSize: 11, color: C.muted, fontFace: FONT, lineSpacingMultiple: 1.3 });
});

// ===== SLIDE 14: NEXT STEPS / CLOSING =====
const s14 = pptx.addSlide();
s14.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: C.bg } });
s14.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.35, h: 7.5, fill: { color: C.accent }, border: false });
// decorative
s14.addShape(pptx.ShapeType.roundRect, { x: 12.3, y: 0.5, w: 0.5, h: 0.5, fill: { color: C.accentLight }, rectRadius: 0.08, border: false });
s14.addText('Next Steps', { x: 0.7, y: 0.8, w: 11, h: 0.6, fontSize: 26, color: C.accent, fontFace: FONT, bold: true });
s14.addShape(pptx.ShapeType.rect, { x: 0.7, y: 1.45, w: 2.0, h: 0.045, fill: { color: C.accent }, border: false });

const steps = [
  { phase: 'PHASE 1', label: 'Pilot', desc: 'Test with 2–3 departments for one review cycle. Gather feedback on usability, badge logic accuracy, and prediction usefulness.', color: C.accent, num: '1' },
  { phase: 'PHASE 2', label: 'Calibration Test', desc: 'Expand to all managers. Validate rating distributions across departments and fine-tune manager leniency detection.', color: C.accentHover, num: '2' },
  { phase: 'PHASE 3', label: 'Full Rollout', desc: 'Company-wide launch with employee communication, manager training materials, and ongoing model refinement based on actual outcomes.', color: C.success, num: '3' },
];
steps.forEach((st, i) => {
  const x = 0.7 + i * 4.2;
  card(s14, x, 1.9, 3.8, 3.2);
  // Phase number circle
  s14.addShape(pptx.ShapeType.roundRect, { x: x + 0.3, y: 2.1, w: 0.6, h: 0.6, fill: { color: st.color }, rectRadius: 0.3, border: false });
  s14.addText(st.num, { x: x + 0.3, y: 2.1, w: 0.6, h: 0.6, fontSize: 20, color: C.white, fontFace: FONT, bold: true, align: 'center', valign: 'middle' });
  s14.addText(st.phase, { x: x + 1.1, y: 2.15, w: 2.4, h: 0.3, fontSize: 9, color: st.color, fontFace: FONT, bold: true, valign: 'middle' });
  s14.addText(st.label, { x: x + 1.1, y: 2.4, w: 2.4, h: 0.3, fontSize: 15, color: C.text, fontFace: FONT, bold: true, valign: 'middle' });
  s14.addText(st.desc, { x: x + 0.3, y: 2.9, w: 3.2, h: 1.8, fontSize: 11, color: C.muted, fontFace: FONT, lineSpacingMultiple: 1.4 });
  // Arrow between phases
  if (i < 2) {
    s14.addText('→', { x: x + 3.6, y: 3.1, w: 0.4, h: 0.4, fontSize: 20, color: C.faint, fontFace: FONT, align: 'center' });
  }
});

// CTA Button
s14.addShape(pptx.ShapeType.roundRect, { x: 3.0, y: 5.6, w: 7.3, h: 0.8, fill: { color: C.accent }, rectRadius: 0.1, border: false });
s14.addText('Try the Clickable Prototype →  http://127.0.0.1:3000', {
  x: 3.0, y: 5.6, w: 7.3, h: 0.8, fontSize: 14, color: C.white, fontFace: FONT, bold: true, align: 'center', valign: 'middle',
});
s14.addText('Thank you — Questions & Feedback Welcome', {
  x: 0.7, y: 6.8, w: 11, h: 0.4, fontSize: 13, color: C.muted, fontFace: FONT, align: 'center',
});
footer(s14, 14);

// SAVE
pptx.writeFile({ fileName: 'C:\\Users\\sihle\\Desktop\\Empeo\\empeo_performance_insights_deck.pptx' })
  .then(() => console.log('Presentation saved successfully — 14 slides'))
  .catch(err => console.log('Error:', err));
