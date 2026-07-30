import { resolve } from 'path';
import { mkdirSync } from 'fs';
import puppeteer from 'puppeteer';
import type { DfmImageMap, DfmResultEntry, DfmResults } from '../decode/types.js';
import { sanitizeName } from '../utils/names.js';

// Display ordering: errors → warnings → info; within a severity,
// fabrication → assembly → general. Kept local so the report has no
// import cycle with the pipeline (which exports the same comparator for
// the image-tile picker).
const SEVERITY_ORDER: Record<string, number> = { error: 0, warning: 1, info: 2 };
const TYPE_ORDER: Record<string, number> = { fabrication: 0, assembly: 1, general: 2 };

// Accepts anything carrying severity + ruleType so the per-rule rollup
// rows (which have no coordinates or description) sort with the same
// comparator as full entries.
type DfmOrderable = Pick<DfmResultEntry, 'severity' | 'ruleType'>;

function compareEntries(a: DfmOrderable, b: DfmOrderable): number {
  const sa = SEVERITY_ORDER[a.severity] ?? 3;
  const sb = SEVERITY_ORDER[b.severity] ?? 3;
  if (sa !== sb) return sa - sb;
  return (TYPE_ORDER[a.ruleType] ?? 3) - (TYPE_ORDER[b.ruleType] ?? 3);
}

function esc(s: unknown): string {
  return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function num(n: number | null | undefined, decimals = 2): string {
  if (n == null || isNaN(n)) return '—';
  return Number(n).toFixed(decimals);
}
function sevColor(severity: string): string {
  return severity === 'error' ? '#dc2626' : severity === 'warning' ? '#d97706' : '#64748b';
}
function sevBadgeClass(severity: string): string {
  return severity === 'error' ? 'badge-err' : severity === 'warning' ? 'badge-warn' : 'badge-info';
}

export type DfmPdfInput = {
  productName: string;
  projectId: string;
  quantity?: number;
  dfm: DfmResults;
  // Per-violation rendered tiles keyed by entry reference.
  images: DfmImageMap;
  outputDir: string;
};

export function buildDfmHtml(input: Omit<DfmPdfInput, 'outputDir'>): string {
  const { productName: pN, projectId: pid, dfm, images } = input;
  const s = dfm.summary;
  const results = dfm.results || [];
  const sorted = [...results].sort(compareEntries);
  const total = s.fabricationErrorCount + s.fabricationWarningCount
    + s.assemblyErrorCount + s.assemblyWarningCount + s.generalCount;
  const now = new Date().toLocaleString();

  // Group by rule for the "Violations by rule" rollup.
  const byRule = new Map<string, { ruleName: string; ruleType: string; severity: string; count: number }>();
  for (const e of sorted) {
    const key = `${e.ruleId}|${e.severity}`;
    const existing = byRule.get(key);
    if (existing) existing.count += 1;
    else byRule.set(key, { ruleName: e.ruleName, ruleType: e.ruleType, severity: e.severity, count: 1 });
  }
  const ruleRows = [...byRule.values()].sort((a, b) => compareEntries(a, b) || b.count - a.count);

  // Gallery: violations that have a rendered tile, in display order.
  const gallery = sorted.filter(e => images.has(e));

  let html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${esc(pN)} — DFM Report</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',Roboto,-apple-system,BlinkMacSystemFont,sans-serif;font-size:11px;color:#1e293b;line-height:1.65}
.page{max-width:820px;margin:0 auto;padding:40px 48px}
h2{font-size:15px;font-weight:700;color:#0f172a;margin-top:32px;margin-bottom:12px;padding-bottom:6px;border-bottom:2px solid #16a34a}
h3{font-size:11px;font-weight:700;color:#334155;margin-top:16px;margin-bottom:8px}
.header-bar{background:linear-gradient(135deg,#0f172a 0%,#1e3a5f 100%);padding:28px 48px;color:white;display:flex;align-items:center;justify-content:space-between}
.header-bar .date{font-size:9px;opacity:0.6;text-align:right}
.header-bar .title-area h1{color:white;font-size:20px;margin-bottom:0}
.header-bar .title-area .sub{font-size:11px;color:rgba(255,255,255,0.7);font-weight:400}
.meta{display:flex;gap:28px;margin:20px 0;flex-wrap:wrap}
.meta-item{font-size:10px;color:#94a3b8}
.meta-item strong{color:#334155;display:block;font-size:12px;font-weight:700}
.kpi-strip{display:flex;gap:12px;margin:20px 0;flex-wrap:wrap}
.kpi{flex:1;min-width:90px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:14px 10px;text-align:center}
.kpi .val{font-size:20px;font-weight:800;color:#0f172a}
.kpi .lbl{font-size:8px;color:#94a3b8;text-transform:uppercase;letter-spacing:0.8px;margin-top:3px;font-weight:600}
.callout{border-radius:8px;padding:12px 16px;margin:16px 0;font-size:10.5px;line-height:1.6}
.callout-ok{background:#dcfce7;border-left:4px solid #16a34a;color:#166534}
.callout-warn{background:#fef9c3;border-left:4px solid #d97706;color:#854d0e}
.callout-err{background:#fef2f2;border-left:4px solid #dc2626;color:#991b1b}
table{width:100%;border-collapse:collapse;margin-bottom:12px;font-size:10px}
th{background:#f1f5f9;text-align:left;padding:7px 10px;font-weight:700;color:#475569;border-bottom:2px solid #e2e8f0;font-size:9px;text-transform:uppercase;letter-spacing:0.3px}
td{padding:6px 10px;border-bottom:1px solid #f1f5f9;vertical-align:top}
tr:nth-child(even){background:#fafbfc}
.right{text-align:right}.center{text-align:center}
.mono{font-family:'SF Mono',Consolas,'Courier New',monospace;font-size:9px}
.badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:8px;font-weight:700;text-transform:uppercase;letter-spacing:0.3px}
.badge-ok{background:#dcfce7;color:#166534}.badge-warn{background:#fef9c3;color:#854d0e}.badge-err{background:#fee2e2;color:#991b1b}.badge-info{background:#e0f2fe;color:#075985}
.gallery{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin:12px 0}
.tile{border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;background:#fff;page-break-inside:avoid}
.tile .img{background:#0b1220;text-align:center;padding:8px}
.tile .img img{max-width:50%;height:auto;image-rendering:pixelated}
.tile .body{padding:10px 12px}
.tile .body .rule{font-weight:700;font-size:11px;color:#0f172a;margin:4px 0 2px}
.tile .body .desc{font-size:9.5px;color:#475569;margin-top:4px}
.tile .body .kv{font-size:9px;color:#64748b;margin-top:4px;font-family:'SF Mono',Consolas,monospace}
.section-intro{font-size:10px;color:#64748b;margin-bottom:10px;line-height:1.6}
.footer{margin-top:40px;padding:16px 0;border-top:2px solid #e2e8f0;font-size:9px;color:#94a3b8;text-align:center;display:flex;justify-content:space-between;align-items:center}
.page-break{page-break-before:always}
@media print{.page{padding:20px 30px;max-width:none}h2{page-break-after:avoid}.page-break{page-break-before:always}.header-bar{-webkit-print-color-adjust:exact;print-color-adjust:exact}.tile .img{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
</style></head><body>`;

  html += `<div class="header-bar"><div class="title-area"><h1>${esc(pN)}</h1><div class="sub">Boardera SpeedDFM — Manufacturing &amp; Assembly Report</div></div><div class="date">Generated<br>${esc(now)}</div></div>`;
  html += `<div class="page">`;

  html += `<div class="meta">`;
  if (input.quantity != null) html += `<div class="meta-item"><strong>${esc(input.quantity)}</strong>Quantity</div>`;
  html += `<div class="meta-item"><strong>${esc(pid?.slice(0, 12) || '—')}</strong>Project ID</div>`;
  html += `<div class="meta-item"><strong>${esc(s.analysisStatus)}</strong>Analysis Status</div>`;
  html += `<div class="meta-item"><strong>${total}</strong>Total Violations</div>`;
  html += `</div>`;

  // Top-level verdict callout.
  const errors = s.fabricationErrorCount + s.assemblyErrorCount;
  const warnings = s.fabricationWarningCount + s.assemblyWarningCount;
  if (total === 0 && s.analysisStatus === 'complete') {
    html += `<div class="callout callout-ok"><b>No DFM violations detected.</b> SpeedDFM completed and found no fabrication or assembly issues against the manufacturer's rule set.</div>`;
  } else if (errors > 0) {
    html += `<div class="callout callout-err"><b>${errors} error${errors === 1 ? '' : 's'}</b> and <b>${warnings} warning${warnings === 1 ? '' : 's'}</b> were found. Errors typically block fabrication or assembly and should be resolved before release.</div>`;
  } else if (warnings > 0 || s.generalCount > 0) {
    html += `<div class="callout callout-warn"><b>No errors</b>, but <b>${warnings} warning${warnings === 1 ? '' : 's'}</b>${s.generalCount ? ` and ${s.generalCount} general note${s.generalCount === 1 ? '' : 's'}` : ''} were found. Review before release.</div>`;
  }

  // Summary KPIs.
  html += `<h2>Summary</h2><div class="kpi-strip">`;
  html += `<div class="kpi"><div class="val" style="color:${s.fabricationErrorCount > 0 ? '#dc2626' : '#a1a1aa'}">${s.fabricationErrorCount}</div><div class="lbl">Fab Errors</div></div>`;
  html += `<div class="kpi"><div class="val" style="color:${s.fabricationWarningCount > 0 ? '#d97706' : '#a1a1aa'}">${s.fabricationWarningCount}</div><div class="lbl">Fab Warnings</div></div>`;
  html += `<div class="kpi"><div class="val" style="color:${s.assemblyErrorCount > 0 ? '#dc2626' : '#a1a1aa'}">${s.assemblyErrorCount}</div><div class="lbl">Asm Errors</div></div>`;
  html += `<div class="kpi"><div class="val" style="color:${s.assemblyWarningCount > 0 ? '#d97706' : '#a1a1aa'}">${s.assemblyWarningCount}</div><div class="lbl">Asm Warnings</div></div>`;
  html += `<div class="kpi"><div class="val">${s.generalCount}</div><div class="lbl">General</div></div>`;
  html += `</div>`;

  // Violations by rule.
  if (ruleRows.length) {
    html += `<h2>Violations by Rule</h2>`;
    html += `<table><tr><th>Severity</th><th>Type</th><th>Rule</th><th class="right">Count</th></tr>`;
    ruleRows.forEach(r => {
      html += `<tr><td><span class="badge ${sevBadgeClass(r.severity)}">${esc(r.severity)}</span></td><td>${esc(r.ruleType)}</td><td>${esc(r.ruleName)}</td><td class="right mono">${r.count}</td></tr>`;
    });
    html += `</table>`;
  }

  // Image gallery.
  if (gallery.length) {
    html += `<h2 class="page-break">Violation Detail (${gallery.length} rendered)</h2>`;
    html += `<p class="section-intro">Each tile is a composite render of the project's PCB layers cropped around the violation, with a magenta crosshair marking the exact location. ${gallery.length < results.length ? `Showing the ${gallery.length} highest-severity violations with coordinates; the full list of ${results.length} follows.` : ''}</p>`;
    html += `<div class="gallery">`;
    gallery.forEach((e, i) => {
      const img = images.get(e)!;
      const coords = e.x != null && e.y != null ? `${num(e.x)}, ${num(e.y)} mm` : '—';
      const meas = e.measuredValue != null && e.threshold != null
        ? `measured ${num(e.measuredValue, 3)} / limit ${num(e.threshold, 3)} ${esc(e.unit || '')}`
        : '';
      html += `<div class="tile"><div class="img"><img src="${img.dataUri}" alt="violation ${i + 1}"></div><div class="body">`;
      html += `<span class="badge ${sevBadgeClass(e.severity)}">${esc(e.severity)}</span> <span class="badge badge-ok" style="background:#eef2ff;color:#3730a3">${esc(e.ruleType)}</span>`;
      html += `<div class="rule">${esc(e.ruleName)}</div>`;
      if (e.description) html += `<div class="desc">${esc(e.description)}</div>`;
      html += `<div class="kv">${esc(e.layerName || e.layerType || 'layer n/a')} · ${coords}</div>`;
      if (meas) html += `<div class="kv">${meas}</div>`;
      html += `</div></div>`;
    });
    html += `</div>`;
  }

  // Full violation table.
  if (results.length) {
    html += `<h2 class="page-break">All Violations (${results.length})</h2>`;
    html += `<table><tr><th>#</th><th>Severity</th><th>Type</th><th>Rule</th><th>Layer</th><th class="right">x / y (mm)</th><th class="right">Measured / Limit</th><th>Description</th></tr>`;
    sorted.forEach((e, i) => {
      const coords = e.x != null && e.y != null ? `${num(e.x)} / ${num(e.y)}` : '—';
      const meas = e.measuredValue != null && e.threshold != null
        ? `${num(e.measuredValue, 3)} / ${num(e.threshold, 3)} ${esc(e.unit || '')}`
        : '—';
      const star = images.has(e) ? ' ★' : '';
      html += `<tr><td class="mono">${i + 1}${star}</td><td style="color:${sevColor(e.severity)};font-weight:600">${esc(e.severity)}</td><td>${esc(e.ruleType)}</td><td class="mono">${esc(e.ruleName)}</td><td>${esc(e.layerName || e.layerType || '—')}</td><td class="right mono">${coords}</td><td class="right mono">${meas}</td><td>${esc(e.description || '')}</td></tr>`;
    });
    html += `</table>`;
    if (gallery.length) html += `<p class="section-intro">★ = a rendered tile appears in the Violation Detail section above.</p>`;
  } else if (s.analysisStatus !== 'complete') {
    html += `<div class="callout callout-warn">SpeedDFM analysis status is <b>${esc(s.analysisStatus)}</b> — no detailed results are available.</div>`;
  }

  html += `<div class="footer"><span>&copy; ${new Date().getFullYear()} Boardera Software Inc.</span><span>${esc(pN)} &middot; ${esc(pid?.slice(0, 12) || '—')}</span><span>Boardera Decode Email Demo</span></div>`;
  html += `</div></body></html>`;
  return html;
}

export async function generateDfmPdf(input: DfmPdfInput): Promise<string> {
  const html = buildDfmHtml(input);
  mkdirSync(input.outputDir, { recursive: true });
  const outPath = resolve(input.outputDir, `${sanitizeName(input.productName)}_DFM_Report.pdf`);

  const browser = await puppeteer.launch({ headless: true, args: ['--no-sandbox'] });
  try {
    const page = await browser.newPage();
    await page.emulateMediaType('print');
    await page.setContent(html, { waitUntil: 'networkidle0' });
    await page.pdf({
      path: outPath,
      format: 'A4',
      printBackground: true,
      margin: { top: '0', right: '0', bottom: '0', left: '0' },
    });
  } finally {
    await browser.close();
  }
  return outPath;
}
