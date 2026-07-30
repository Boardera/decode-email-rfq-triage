import { resolve } from 'path';
import { mkdirSync } from 'fs';
import puppeteer from 'puppeteer';
import type {
  DfmResults,
  ExportProductData,
  ExportPricingLineItem,
  Money,
  PriceResponse,
  ValueWithUnit,
} from '../decode/types.js';
import {
  classifyRow,
  filterRealBom,
  isEol,
  LONG_LEAD_DAYS,
  LOW_STOCK_THRESHOLD,
  orderQtyOf,
  resolveSourcing,
} from './bomAnalysis.js';
import { compareDfmEntries } from '../decode/pipeline.js';
import { sanitizeName } from '../utils/names.js';

function commaFmt(n: number | null | undefined, decimals: number): string {
  if (n == null || isNaN(n as number)) return '—';
  return Number(n).toLocaleString(undefined, { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}
function dollarFmt(n: number | null | undefined, decimals = 2): string {
  if (n == null || isNaN(n as number)) return '—';
  return '$' + commaFmt(n, decimals);
}

function valU(v: ValueWithUnit | null | undefined, decimals = 2): string {
  if (!v || v.value == null) return '—';
  return `${Number(v.value).toFixed(decimals).replace(/\.?0+$/, '')} ${v.unit}`;
}
function valOnly(v: ValueWithUnit | null | undefined): number | null {
  return v?.value ?? null;
}
function money(m: Money | null | undefined): string {
  if (!m) return '—';
  if (m.formatted) return m.formatted;
  return dollarFmt(m.value, 2);
}
function moneyN(m: Money | null | undefined): number {
  return m?.value ?? 0;
}

const VIA_PLUGGING_LABELS: Record<string, string> = {
  'None': 'None',
  'NONE': 'None',
  'Type I Tent': 'Type I — Tent',
  'Type II Tent & Cover': 'Type II — Tent & Cover',
  'Type III Plug': 'Type III — Plug',
  'Type IV Plug & Cover': 'Type IV — Plug & Cover',
  'Type V Fill Epoxy': 'Type V — Fill (Epoxy)',
  'Type VI Fill & Cover Epoxy': 'Type VI — Fill & Cover (Epoxy)',
  'Type VII Fill & Cap Epoxy': 'Type VII — Fill & Cap (Epoxy)',
  'Type VII Fill & Cap Copper': 'Type VII — Fill & Cap (Copper)',
  'Solder Mask Plug & Cover': 'Solder Mask Plug & Cover',
};
function viaPlugLabel(v: string | null | undefined): string | null {
  if (!v || v === 'None' || v === 'NONE') return null;
  return VIA_PLUGGING_LABELS[v] || v;
}

// Exception/warning classification, sourcing resolution, and the
// real-row filter all live in ./bomAnalysis so this report, the BOM
// Excel, and the email reply share one definition.

export type ExecPdfInput = {
  productName: string;
  projectId: string;
  exportData: ExportProductData;
  dfm: DfmResults | null;
  prices: PriceResponse[];
  quantities: number[];
  outputDir: string;
};

export function buildExecHtml(input: Omit<ExecPdfInput, 'outputDir'>): string {
  const { productName: pN, projectId: pid, exportData, dfm, prices, quantities } = input;
  const specs = exportData.specs || {};
  const bb = specs.bareBoard || {};
  const opts = specs.assemblyOptions || {};
  // Same row filter as the BOM Excel and the email reply — drop the
  // empty DNP placeholder rows so all three surfaces agree.
  const bom = filterRealBom(specs.projectBom || []);
  const stackup = specs.stackup || [];
  const pt = specs.padTraceData;
  const dr = specs.drillData;
  const prov = specs.provenance;
  const fileNotes = specs.fileNotes || [];
  const fileTables = specs.fileTables || [];

  const scenario = exportData.costing?.scenarios?.[0];
  const qty = quantities?.[0] || scenario?.quantity || 1;
  const pli: ExportPricingLineItem[] = scenario?.pricingLineItems || [];
  const calcs = scenario?.calculations || [];
  const pp = scenario?.profitProjections || [];
  const subs = scenario?.validationDetails?.assemblyWithSubstitutions || [];

  // Build BOM → pricing join (group === 'BOM' rows).
  const pricingByName: Record<string, ExportPricingLineItem> = {};
  pli.forEach(p => { if (p.group === 'BOM' && p.name) pricingByName[p.name] = p; });

  const joined = bom.map(b => {
    const pricing = pricingByName[b.partNumber || ''] || pricingByName[b.pricedPartNumber || ''];
    return { bomItem: b, pricing, cls: classifyRow(b, pricing) };
  });
  type JoinedRow = typeof joined[number];
  const hasExc = (j: JoinedRow, code: string) => j.cls.exceptions.some(e => e.code === code);
  const hasWarn = (j: JoinedRow, code: string) => j.cls.warnings.some(w => w.code === code);

  const fmt = (v: any) => v != null ? Number(v).toLocaleString() : '—';
  const esc = (s: any) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const otherLabel = (val: any, txt: any) => val && txt ? `${esc(val)} (${esc(txt)})` : esc(val || '—');

  const hasBom = bom.length > 0;
  const smtCount = bom.filter(b => b.mountType === 'Surface Mount').length;
  const thCount = bom.filter(b => b.mountType === 'Through Hole').length;
  const eolItems = bom.filter(b => isEol(b.lifeCycleStatus));
  // Counts come straight from the shared classifier so they match the BOM
  // Excel and the email exactly. sourcingByRow is kept only for rendering
  // the per-row sourcing details (stock / lead / distributor) in sections.
  const sourcingByRow = new Map(joined.map(j => [j, resolveSourcing(j.pricing)]));
  const rowsWithSourcing = [...sourcingByRow.values()].filter(s => s.hasSourcing).length;
  const lowStock = joined.filter(j => hasWarn(j, 'lowStock'));            // warning, not an exception
  const noStock = joined.filter(j => hasExc(j, 'noStock'));
  const insufficientStock = joined.filter(j => hasExc(j, 'insufficientStock'));
  const unfindable = joined.filter(j => hasExc(j, 'unfindable'));
  const longLead = joined.filter(j => hasExc(j, 'longLead'));
  const backordered = joined.filter(j => hasExc(j, 'backordered'));
  const customerSupplied = bom.filter(b => b.customerSupplied);          // now an exception too

  const bomTotal = pli.filter(p => p.group === 'BOM').reduce((s, p) => s + moneyN(p.total), 0);
  // Prefer the scenario-level subtotal (same as the email reply's "Quoted
  // total") so the two surfaces always agree; fall back to summing line items.
  const grandTotal = scenario?.subtotal ?? pli.reduce((s, p) => s + moneyN(p.total), 0);

  const bomSorted = [...joined].sort((a, b) => moneyN(b.pricing?.total) - moneyN(a.pricing?.total));
  const topCost = bomSorted.slice(0, 10);
  const groups = ['PCB', 'ASM', 'BOM'];
  const groupTotals = groups.map(g => {
    const items = pli.filter(p => p.group === g);
    return { group: g, items, total: items.reduce((s, p) => s + moneyN(p.total), 0) };
  });
  const topDrivers = [...pli].sort((a, b) => moneyN(b.total) - moneyN(a.total)).slice(0, 10);

  // A row is an exception row if the shared classifier flagged any
  // exception (customer-supplied, EOL, sourcing, substitution, …).
  // Substitutions are already counted per-row, so we do NOT also add the
  // project-level subs[] count here (that double-counted previously).
  const rowsWithExceptions = joined.filter(j => j.cls.exceptions.length > 0);
  const rowsWithWarnings = joined.filter(j => j.cls.exceptions.length === 0 && j.cls.warnings.length > 0);
  const totalExceptions = rowsWithExceptions.length;
  const totalWarnings = rowsWithWarnings.length;
  const totalDecisions = prov?.totalDecisions || 0;
  const areas = (prov?.areas || []).map(a => ({ area: a.area, entries: a.entries || [], count: (a.entries || []).length }))
    .sort((a, b) => b.count - a.count);
  const totalPlacementsPerBoard = bom.reduce((s, b) => s + (b.quantityPerBoard || 1), 0);

  const ppTotal = pp.find(p => p.type === 'Total');
  const ppSegments = pp.filter(p => p.type !== 'Total');

  const caps: [string, any][] = [
    ['HDI', bb.hdi],
    ['Controlled Impedance', bb.controlledImpedance],
    ['Blind & Buried Vias', bb.blindAndBuriedVias],
    ['Controlled Depth Drilling', bb.controlledDepthDrilling],
    ['Gold Fingers', bb.goldFingers],
    ['Custom Stackup', bb.customStackup],
    ['Electrical Testing', bb.electricalTesting],
    ['Via Plugging', viaPlugLabel(bb.viaPlugging)],
  ];
  const activeCaps = caps.filter(([, v]) => v);

  const pkgMap: Record<string, number> = {};
  bom.forEach(b => {
    const desc = b.description || '';
    const m = desc.match(/\b(0201|0402|0603|0805|1206|1210|1812|2010|2512|SOT-?\d+\w*|SOD-?\d+\w*|QFP\w*|QFN\w*|TQFP\w*|SOP\w*|SOIC\w*|SSOP\w*|DIP\w*|BGA\w*|SMA|SMB|SMC|DO-?\d+|SC-?\d+|1411|TO-?\d+)/i);
    const pkg = m ? m[1].toUpperCase() : 'Other';
    pkgMap[pkg] = (pkgMap[pkg] || 0) + (b.quantityPerBoard || 1);
  });
  const pkgSorted = Object.entries(pkgMap).sort((a, b) => b[1] - a[1]);

  const now = new Date().toLocaleString();

  let html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${esc(pN)} — Analysis Report</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',Roboto,-apple-system,BlinkMacSystemFont,sans-serif;font-size:11px;color:#1e293b;line-height:1.65;padding:0}
.page{max-width:820px;margin:0 auto;padding:40px 48px}
h1{font-size:24px;font-weight:800;color:#0f172a;margin-bottom:2px;letter-spacing:-0.5px}
h2{font-size:15px;font-weight:700;color:#0f172a;margin-top:32px;margin-bottom:12px;padding-bottom:6px;border-bottom:2px solid #16a34a}
h3{font-size:11px;font-weight:700;color:#334155;margin-top:16px;margin-bottom:8px}
.header-bar{background:linear-gradient(135deg,#0f172a 0%,#1e3a5f 100%);padding:28px 48px;color:white;display:flex;align-items:center;justify-content:space-between}
.header-bar .date{font-size:9px;opacity:0.6;text-align:right}
.header-bar .title-area h1{color:white;font-size:20px;margin-bottom:0}
.header-bar .title-area .sub{font-size:11px;color:rgba(255,255,255,0.7);font-weight:400}
.subtitle{font-size:12px;color:#64748b;margin-bottom:20px}
.kpi-strip{display:flex;gap:12px;margin:20px 0;flex-wrap:wrap}
.kpi{flex:1;min-width:100px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:14px 10px;text-align:center}
.kpi .val{font-size:20px;font-weight:800;color:#0f172a}
.kpi .lbl{font-size:8px;color:#94a3b8;text-transform:uppercase;letter-spacing:0.8px;margin-top:3px;font-weight:600}
.exec{background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:20px 24px;margin-bottom:24px;border-left:4px solid #16a34a}
.exec-title{font-size:13px;font-weight:700;color:#0f172a;margin-bottom:10px}
.exec-text{font-size:10.5px;color:#475569;line-height:1.7}
.meta{display:flex;gap:28px;margin-bottom:20px;flex-wrap:wrap}
.meta-item{font-size:10px;color:#94a3b8}
.meta-item strong{color:#334155;display:block;font-size:12px;font-weight:700}
table{width:100%;border-collapse:collapse;margin-bottom:12px;font-size:10px}
th{background:#f1f5f9;text-align:left;padding:7px 10px;font-weight:700;color:#475569;border-bottom:2px solid #e2e8f0;font-size:9px;text-transform:uppercase;letter-spacing:0.3px}
td{padding:6px 10px;border-bottom:1px solid #f1f5f9}
tr:nth-child(even){background:#fafbfc}
.right{text-align:right}
.center{text-align:center}
.mono{font-family:'SF Mono',Consolas,'Courier New',monospace;font-size:9px}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:6px 28px;margin-bottom:12px}
.grid dt{color:#64748b;font-size:10px}.grid dd{font-weight:700;font-size:11px;color:#1e293b}
.badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:8px;font-weight:700;margin-right:4px}
.badge-ok{background:#dcfce7;color:#166534}.badge-warn{background:#fef9c3;color:#854d0e}.badge-err{background:#fee2e2;color:#991b1b}.badge-info{background:#e0f2fe;color:#075985}
.note-box{border:1px solid #e2e8f0;border-radius:8px;padding:10px 14px;margin-bottom:8px;font-size:10.5px;color:#334155;background:#fafbfc}
.note-src{font-size:8px;color:#94a3b8;margin-top:4px;font-family:monospace}
.footer{margin-top:40px;padding:16px 0;border-top:2px solid #e2e8f0;font-size:9px;color:#94a3b8;text-align:center;display:flex;justify-content:space-between;align-items:center}
.footer .copy{opacity:0.7}
.page-break{page-break-before:always}
ul{padding-left:18px;margin-bottom:10px}li{margin-bottom:4px}
.section-intro{font-size:10px;color:#64748b;margin-bottom:10px;line-height:1.6}
@media print{.page{padding:20px 30px;max-width:none}h2{page-break-after:avoid}.page-break{page-break-before:always}.header-bar{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
</style></head><body>`;

  html += `<div class="header-bar"><div class="title-area"><h1>${esc(pN)}</h1><div class="sub">Boardera API — Project Analysis Report</div></div><div class="date">Generated<br>${esc(now)}</div></div>`;
  html += `<div class="page">`;
  html += `<div class="meta">`;
  html += `<div class="meta-item"><strong>${fmt(qty)}</strong>Quantity</div>`;
  if (grandTotal) html += `<div class="meta-item"><strong>${dollarFmt(grandTotal)}</strong>Estimated Total</div>`;
  html += `<div class="meta-item"><strong>${esc(pid?.slice(0, 12) || '—')}</strong>Project ID</div>`;
  html += `</div>`;

  // Executive summary
  html += `<div class="exec"><div class="exec-title">Executive Summary</div>`;
  html += `<div class="kpi-strip">`;
  html += `<div class="kpi"><div class="val">${bb.numOfLayers || '—'}</div><div class="lbl">Layers</div></div>`;
  html += `<div class="kpi"><div class="val">${hasBom ? fmt(totalPlacementsPerBoard) : '—'}</div><div class="lbl">Placements / Board</div></div>`;
  html += `<div class="kpi"><div class="val">${hasBom ? bom.length : '—'}</div><div class="lbl">BOM Parts</div></div>`;
  html += `<div class="kpi"><div class="val">${totalExceptions}</div><div class="lbl">Exceptions</div></div>`;
  html += `<div class="kpi"><div class="val">${totalDecisions}</div><div class="lbl">Provenance Entries</div></div>`;
  html += `</div>`;

  const summaryParts: string[] = [];
  const wMm = valOnly(bb.boardWidth);
  const lMm = valOnly(bb.boardLength);
  summaryParts.push(`This ${bb.numOfLayers || '?'}‑layer ${esc(bb.material || 'FR-4')} board (${wMm ?? '?'}×${lMm ?? '?'} mm) uses ${esc(bb.finishing || 'standard')} finish with ${esc(bb.solderMaskColor || 'standard')} solder mask.`);
  if (hasBom) summaryParts.push(`The BOM contains ${bom.length} unique components (${smtCount} SMT, ${thCount} through-hole) totalling ${dollarFmt(bomTotal)} in parts cost.`);
  if (eolItems.length) summaryParts.push(`<span style="color:#dc2626"><b>${eolItems.length} end-of-life</b></span> components require attention.`);
  if (lowStock.length) summaryParts.push(`${lowStock.length} parts have low stock (<100 units).`);
  if (longLead.length) summaryParts.push(`${longLead.length} parts have long lead times (>${LONG_LEAD_DAYS} days).`);
  if (grandTotal) summaryParts.push(`Quoted total: <b>${dollarFmt(grandTotal)}</b>.`);
  if (ppTotal) summaryParts.push(`Projected margin: <b>${ppTotal.margin?.toFixed(1) || '0.0'}%</b> (gross profit: ${money(ppTotal.grossProfit)}).`);
  if (totalDecisions) summaryParts.push(`The Boardera API engine made ${totalDecisions} automated decisions across ${areas.length} areas.`);
  html += `<p class="exec-text">${summaryParts.join(' ')}</p>`;
  html += `</div>`;

  // === 1. BOARD SPECS ===
  html += `<h2>1. Board Specifications</h2><dl class="grid">`;
  const dimText = (bb.boardWidth && bb.boardLength)
    ? `${valU(bb.boardWidth)} × ${valU(bb.boardLength)}`
    : '—';
  const specRows: [string, any][] = [
    ['Material', otherLabel(bb.material, bb.materialOtherText)],
    ['Layers', fmt(bb.numOfLayers)],
    ['Thickness', valU(bb.boardThickness)],
    ['Dimensions', dimText],
    ['Copper', valU(bb.copperThickness, 0)],
    ['Finishing', otherLabel(bb.finishing, bb.finishingOtherText)],
    ['Solder Mask', otherLabel(bb.solderMaskColor, bb.solderMaskColorOtherText)],
    ['Silkscreen', otherLabel(bb.silkscreenColor, bb.silkscreenColorOtherText)],
    ['IPC Class III', specs.ipcClassIII ? 'Yes' : 'No'],
    ['Industry Sector', esc(specs.industrySector || '—')],
    ['Customer Supplied', bb.customerSupplied ? 'Yes' : 'No'],
  ];
  specRows.forEach(([k, v]) => { html += `<dt>${k}</dt><dd>${v}</dd>`; });
  html += `</dl>`;
  if (activeCaps.length) html += `<h3>Capabilities</h3><p style="font-size:10px">${activeCaps.map(([k, v]) => typeof v === 'string' ? k + ': ' + v : k).join(' · ')}</p>`;
  const panel = bb.panel;
  if (panel && (panel.panelWidth || panel.panelLength)) {
    html += `<h3>Panel Configuration</h3><dl class="grid">`;
    if (panel.panelWidth) html += `<dt>Panel Size</dt><dd>${valU(panel.panelWidth)} × ${valU(panel.panelLength)}</dd>`;
    if (panel.boardsPerAsmPanel) html += `<dt>Boards / Panel</dt><dd>${panel.boardsPerAsmPanel}</dd>`;
    if (panel.designFilesArePanel != null) html += `<dt>Files Are Panel</dt><dd>${panel.designFilesArePanel ? 'Yes' : 'No'}</dd>`;
    html += `</dl>`;
  }
  if (stackup.length) {
    html += `<h3>Stackup (${stackup.length} layers)</h3><table><tr><th>#</th><th>Type</th><th>File Type</th><th class="right">Copper</th><th>Dielectric</th><th class="right">Thickness</th></tr>`;
    stackup.forEach((s, i) => {
      html += `<tr><td>${i + 1}</td><td>${esc(s.type)}</td><td>${esc(s.fileType)}</td><td class="right mono">${valU(s.copperThickness, 0)}</td><td>${esc(s.dielectricType || '—')}</td><td class="right mono">${valU(s.dielectricThickness)}</td></tr>`;
    });
    html += `</table>`;
  }

  // === 2. PCB STATS ===
  if (pt || dr) {
    html += `<h2>2. PCB Design Statistics</h2><div class="kpi-strip">`;
    if (pt) {
      html += `<div class="kpi"><div class="val">${fmt(pt.totalPadCount)}</div><div class="lbl">Total Pads</div></div>`;
      html += `<div class="kpi"><div class="val">${fmt(pt.totalTraceCount)}</div><div class="lbl">Total Traces</div></div>`;
    }
    if (dr) html += `<div class="kpi"><div class="val">${fmt(dr.totalHoleCount)}</div><div class="lbl">Drill Holes</div></div>`;
    html += `</div>`;
    if (pt?.layers?.length) {
      html += `<h3>Layer Details</h3><table><tr><th>Layer</th><th class="right">Pads</th><th class="right">Circular</th><th class="right">Rect</th><th class="right">Traces</th><th class="right">Trace Width</th></tr>`;
      pt.layers.forEach(l => {
        const tw = (l.minTraceWidth && l.maxTraceWidth)
          ? (l.minTraceWidth.value === l.maxTraceWidth.value ? valU(l.minTraceWidth, 3) : `${valU(l.minTraceWidth, 3)} – ${valU(l.maxTraceWidth, 3)}`)
          : '—';
        html += `<tr><td>${esc(l.layerName)}</td><td class="right mono">${fmt(l.totalPadCount)}</td><td class="right mono">${fmt(l.circularPadCount)}</td><td class="right mono">${fmt(l.rectangularPadCount)}</td><td class="right mono">${fmt(l.traceCount)}</td><td class="right mono">${tw}</td></tr>`;
      });
      html += `</table>`;
    }
    if (dr?.drillFiles?.length) {
      html += `<h3>Drill Files</h3><table><tr><th>Description</th><th class="right">Holes</th><th class="right">Diameter</th><th class="right">Hole Area</th></tr>`;
      dr.drillFiles.forEach(d => {
        const diam = (d.minHoleDiameter && d.maxHoleDiameter)
          ? (d.minHoleDiameter.value === d.maxHoleDiameter.value ? valU(d.minHoleDiameter, 3) : `${valU(d.minHoleDiameter, 3)} – ${valU(d.maxHoleDiameter, 3)}`)
          : '—';
        html += `<tr><td>${esc(d.drillDescription)}</td><td class="right mono">${fmt(d.holeCount)}</td><td class="right mono">${diam}</td><td class="right mono">${valU(d.totalHoleSurfaceArea)}</td></tr>`;
      });
      html += `</table>`;
    }
  }

  // === 3. FABRICATION NOTES ===
  if (fileNotes.length || fileTables.length) {
    html += `<h2 class="page-break">3. Fabrication Notes</h2>`;
    if (fileNotes.length) {
      fileNotes.forEach(n => {
        const srcFile = n.foundIn?.map(f => f.fileName).join(', ') || '';
        html += `<div class="note-box">${esc(n.cleanedText || n.text)}${srcFile ? `<div class="note-src">Source: ${esc(srcFile)}</div>` : ''}</div>`;
      });
    }
    if (fileTables.length) {
      html += `<h3>Extracted Tables</h3>`;
      fileTables.forEach(ft => {
        if (ft.foundIn?.length) html += `<p class="note-src" style="margin-bottom:3px">Source: ${ft.foundIn.map(f => esc(f.fileName)).join(', ')}</p>`;
        const rows = ft.table || [];
        if (rows.length > 1) {
          html += `<table style="width:100%;border-collapse:collapse;font-family:monospace;font-size:9px;margin-bottom:8px"><thead><tr style="background:#f4f4f5;border-bottom:2px solid #d4d4d8">${rows[0].map(c => `<th style="padding:4px 8px;text-align:left;font-weight:600;color:#52525b">${esc(c)}</th>`).join('')}</tr></thead><tbody>`;
          rows.slice(1).forEach((cells, ri) => {
            html += `<tr style="border-bottom:1px solid #e4e4e7${ri % 2 ? ';background:#fafafa' : ''}">`;
            cells.forEach(c => { html += `<td style="padding:3px 8px;color:#71717a">${esc(c)}</td>`; });
            html += `</tr>`;
          });
          html += `</tbody></table>`;
        }
      });
    }
  }

  // === 4. ASSEMBLY & BOM ===
  if (hasBom) {
    html += `<h2 class="page-break">4. Assembly & BOM</h2><div class="kpi-strip">`;
    html += `<div class="kpi"><div class="val">${bom.length}</div><div class="lbl">Unique Parts</div></div>`;
    html += `<div class="kpi"><div class="val">${smtCount}</div><div class="lbl">SMT</div></div>`;
    html += `<div class="kpi"><div class="val">${thCount}</div><div class="lbl">Through-Hole</div></div>`;
    html += `<div class="kpi"><div class="val">${fmt(totalPlacementsPerBoard)}</div><div class="lbl">Placements / Board</div></div>`;
    html += `<div class="kpi"><div class="val">${dollarFmt(bomTotal)}</div><div class="lbl">BOM Total</div></div>`;
    html += `</div>`;

    const risks: string[] = [];
    if (eolItems.length) risks.push(`<span class="badge badge-err">${eolItems.length} EOL</span> ${eolItems.map(b => esc(b.partNumber)).join(', ')}`);
    if (noStock.length) risks.push(`<span class="badge badge-err">${noStock.length} No Stock</span>`);
    if (insufficientStock.length) risks.push(`<span class="badge badge-err">${insufficientStock.length} Insufficient Stock</span> (stock &lt; order qty)`);
    if (lowStock.length) risks.push(`<span class="badge badge-warn">${lowStock.length} Low Stock</span> (&lt;100 units)`);
    if (longLead.length) risks.push(`<span class="badge badge-info">${longLead.length} Long Lead</span> (&gt;${LONG_LEAD_DAYS} days)`);
    if (backordered.length) risks.push(`<span class="badge badge-warn">${backordered.length} Backordered</span>`);
    if (unfindable.length) risks.push(`<span class="badge badge-err">${unfindable.length} Unfindable</span>`);
    if (risks.length) html += `<h3>Supply Chain Risks</h3><ul>${risks.map(r => '<li>' + r + '</li>').join('')}</ul>`;
    else html += `<p style="font-size:10px;color:#16a34a">No supply chain risks detected.</p>`;

    const activeOpts = Object.entries(opts).filter(([k, v]) => v && k !== 'otherText');
    if (activeOpts.length) html += `<h3>Assembly Options</h3><p style="font-size:10px">${activeOpts.map(([k]) => k === 'other' && opts.otherText ? 'Other (' + esc(opts.otherText) + ')' : k.replace(/([A-Z])/g, ' $1').trim()).join(' · ')}</p>`;

    if (pkgSorted.length) {
      html += `<h3>Package Type Breakdown</h3><table><tr><th>Package</th><th class="right">Placements / Board</th><th class="right">% of Total</th></tr>`;
      pkgSorted.forEach(([pkg, cnt]) => {
        const pct = totalPlacementsPerBoard > 0 ? (cnt / totalPlacementsPerBoard * 100) : 0;
        html += `<tr><td>${esc(pkg)}</td><td class="right">${cnt}</td><td class="right">${pct.toFixed(1)}%</td></tr>`;
      });
      html += `</table>`;
    }

    if (topCost.length) {
      html += `<h3>Top ${Math.min(10, topCost.length)} BOM Cost Drivers</h3>`;
      html += `<table><tr><th>Part</th><th>Manufacturer</th><th>Description</th><th class="right">Qty</th><th class="right">Unit $</th><th class="right">Total $</th></tr>`;
      topCost.forEach(j => {
        const b = j.bomItem;
        const p = j.pricing;
        html += `<tr><td class="mono">${esc(b.pricedPartNumber || b.partNumber)}</td><td>${esc(b.partManufacturer || '')}</td><td>${esc((b.description || '').slice(0, 40))}</td><td class="right">${fmt(orderQtyOf(p))}</td><td class="right mono">${money(p?.unitPrice)}</td><td class="right mono">${money(p?.total)}</td></tr>`;
      });
      html += `</table>`;
    }

    html += `<h3 class="page-break">Complete Bill of Materials (${bom.length} items)</h3>`;
    html += `<table><tr><th>Part Number</th><th>Manufacturer</th><th>Description</th><th class="center">Qty/Board</th><th class="center">Stock</th><th class="center">Lead</th><th class="center">Lifecycle</th><th class="center">RoHS</th><th class="right">Unit $</th><th class="right">Total $</th></tr>`;
    joined.forEach(j => {
      const b = j.bomItem;
      const p = j.pricing;
      const eol = isEol(b.lifeCycleStatus);
      html += `<tr${eol ? ' style="background:#fef2f2"' : ''}>`;
      html += `<td class="mono">${esc(b.pricedPartNumber || b.partNumber)}${b.pricedPartNumber && b.pricedPartNumber !== b.partNumber ? ' [SUB]' : ''}</td>`;
      html += `<td>${esc(b.partManufacturer || '')}</td>`;
      html += `<td>${esc((b.description || '').slice(0, 35))}</td>`;
      html += `<td class="center">${b.quantityPerBoard || 1}</td>`;
      html += `<td class="center">${p?.specs?.stock != null ? fmt(p.specs.stock) : '—'}</td>`;
      html += `<td class="center">${p?.specs?.leadTime != null ? p.specs.leadTime + 'd' : '—'}</td>`;
      html += `<td class="center">${eol ? 'EOL' : (b.lifeCycleStatus || '?')}</td>`;
      html += `<td class="center">${b.rohsStatus === 'Compliant' ? 'Yes' : (b.rohsStatus ? 'No' : '?')}</td>`;
      html += `<td class="right mono">${money(p?.unitPrice)}</td>`;
      html += `<td class="right mono">${money(p?.total)}</td>`;
      html += `</tr>`;
    });
    html += `</table>`;
  }

  // === 5. EXCEPTIONS ===
  html += `<h2 class="page-break">5. Exceptions & Warnings</h2>`;
  if (hasBom && rowsWithSourcing === 0) {
    html += `<p style="font-size:10.5px;color:#854d0e;background:#fef9c3;border-left:4px solid #d97706;padding:10px 14px;border-radius:4px;margin-bottom:12px">Sourcing data (stock, distributor, lead time) was not returned by the API for any BOM row in this project. Stock-based, distributor-based, and lead-time exceptions are skipped; only EOL / Obsolete, customer-supplied, and substitution exceptions appear below.</p>`;
  } else if (hasBom && rowsWithSourcing < bom.length) {
    html += `<p style="font-size:10.5px;color:#854d0e;background:#fef9c3;border-left:4px solid #d97706;padding:10px 14px;border-radius:4px;margin-bottom:12px">Sourcing data was returned for ${rowsWithSourcing} of ${bom.length} BOM row(s). Stock / distributor / lead-time exceptions are evaluated only on the ${rowsWithSourcing} sourced row(s).</p>`;
  }
  html += `<p class="section-intro"><b>${totalExceptions}</b> row(s) with exceptions (red) and <b>${totalWarnings}</b> with warnings only (amber). Long lead is flagged only when stock won't cover the order; low stock that still covers the order is a warning.</p>`;
  html += `<div class="kpi-strip">`;
  html += `<div class="kpi"><div class="val" style="color:${noStock.length > 0 ? '#dc2626' : '#a1a1aa'}">${noStock.length}</div><div class="lbl">No Stock</div></div>`;
  html += `<div class="kpi"><div class="val" style="color:${insufficientStock.length > 0 ? '#dc2626' : '#a1a1aa'}">${insufficientStock.length}</div><div class="lbl">Insufficient Stock</div></div>`;
  html += `<div class="kpi"><div class="val" style="color:${unfindable.length > 0 ? '#dc2626' : '#a1a1aa'}">${unfindable.length}</div><div class="lbl">Unfindable</div></div>`;
  html += `<div class="kpi"><div class="val" style="color:${longLead.length > 0 ? '#dc2626' : '#a1a1aa'}">${longLead.length}</div><div class="lbl">Long Lead</div></div>`;
  html += `<div class="kpi"><div class="val" style="color:${backordered.length > 0 ? '#dc2626' : '#a1a1aa'}">${backordered.length}</div><div class="lbl">Backordered</div></div>`;
  html += `<div class="kpi"><div class="val" style="color:${eolItems.length > 0 ? '#dc2626' : '#a1a1aa'}">${eolItems.length}</div><div class="lbl">EOL / Obsolete</div></div>`;
  html += `<div class="kpi"><div class="val" style="color:${customerSupplied.length > 0 ? '#dc2626' : '#a1a1aa'}">${customerSupplied.length}</div><div class="lbl">Customer Supplied</div></div>`;
  html += `<div class="kpi"><div class="val">${subs.length}</div><div class="lbl">Substitutions</div></div>`;
  html += `<div class="kpi"><div class="val" style="color:${lowStock.length > 0 ? '#d97706' : '#a1a1aa'}">${lowStock.length}</div><div class="lbl">Low Stock (warning)</div></div>`;
  html += `</div>`;

  if (eolItems.length) {
    html += `<h3>End-of-Life / Obsolete (${eolItems.length})</h3><ul>`;
    eolItems.forEach(b => html += `<li><b class="mono">${esc(b.partNumber)}</b> — ${esc(b.partManufacturer || '?')} — ${esc(b.lifeCycleStatus || '?')}</li>`);
    html += `</ul>`;
  }
  if (noStock.length) {
    html += `<h3>No Stock (${noStock.length})</h3><ul>`;
    noStock.forEach(j => {
      const s = sourcingByRow.get(j)!;
      html += `<li><b class="mono">${esc(j.bomItem.partNumber)}</b> — ${esc(j.bomItem.partManufacturer || '?')}${s.leadTime != null ? ` (lead ${s.leadTime}d)` : ''}</li>`;
    });
    html += `</ul>`;
  }
  if (insufficientStock.length) {
    html += `<h3>Insufficient Stock — stock &lt; order quantity (${insufficientStock.length})</h3><ul>`;
    insufficientStock.forEach(j => {
      const s = sourcingByRow.get(j)!;
      html += `<li><b class="mono">${esc(j.bomItem.partNumber)}</b> — ${esc(j.bomItem.partManufacturer || '?')} — stock <b style="color:#dc2626">${s.stock}</b> / need <b>${orderQtyOf(j.pricing) ?? '?'}</b>${s.distributor ? ` (${esc(s.distributor)})` : ''}</li>`;
    });
    html += `</ul>`;
  }
  if (unfindable.length) {
    html += `<h3>Unfindable — no distributor (${unfindable.length})</h3><ul>`;
    unfindable.forEach(j => html += `<li><b class="mono">${esc(j.bomItem.partNumber)}</b> — ${esc(j.bomItem.partManufacturer || '?')}</li>`);
    html += `</ul>`;
  }
  if (longLead.length) {
    html += `<h3>Long Lead Time (&gt;${LONG_LEAD_DAYS} days, ${longLead.length})</h3><ul>`;
    longLead.forEach(j => {
      const s = sourcingByRow.get(j)!;
      html += `<li><b class="mono">${esc(j.bomItem.partNumber)}</b> — ${s.leadTime}d — ${esc(s.distributor || '?')}</li>`;
    });
    html += `</ul>`;
  }
  if (backordered.length) {
    html += `<h3>Backordered (${backordered.length})</h3><ul>`;
    backordered.forEach(j => {
      const s = sourcingByRow.get(j)!;
      html += `<li><b class="mono">${esc(j.bomItem.partNumber)}</b> — ${esc(j.bomItem.partManufacturer || '?')}${s.distributor ? ` (${esc(s.distributor)})` : ''}</li>`;
    });
    html += `</ul>`;
  }
  if (lowStock.length) {
    html += `<h3 style="color:#854d0e">Low Stock — warning, stock still covers the order (&lt;${LOW_STOCK_THRESHOLD} units, ${lowStock.length})</h3><ul>`;
    lowStock.forEach(j => {
      const s = sourcingByRow.get(j)!;
      html += `<li><b class="mono">${esc(j.bomItem.partNumber)}</b> — ${esc(j.bomItem.partManufacturer || '?')} — stock <b>${s.stock}</b>${s.distributor ? ` (${esc(s.distributor)})` : ''}</li>`;
    });
    html += `</ul>`;
  }
  if (subs.length) {
    html += `<h3>Part Substitutions (${subs.length})</h3>`;
    html += `<table><tr><th>Original Part</th><th>Substituted With</th><th>Ref Des</th></tr>`;
    subs.forEach(s => html += `<tr><td class="mono">${esc(s.primaryPartNumber)}</td><td class="mono">${esc(s.subPartNumber)}</td><td class="mono">${esc(s.refDes || '')}</td></tr>`);
    html += `</table>`;
  }
  if (customerSupplied.length) {
    html += `<h3>Customer-Supplied Components (${customerSupplied.length})</h3>`;
    html += `<p style="font-size:10px;color:#64748b">Flagged as a <b>Customer supplied</b> exception. The customer provides these parts, so they are excluded from stock / distributor / lead-time checks: ${customerSupplied.map(b => `<span class="mono">${esc(b.partNumber)}</span>`).join(', ')}.</p>`;
  }
  if (totalExceptions === 0 && totalWarnings === 0) {
    html += `<p style="color:#16a34a;font-size:11px;font-weight:600">No exceptions or warnings detected.</p>`;
  }

  // === 6. PRICING ===
  html += `<h2 class="page-break">6. Pricing Summary</h2>`;
  if (groupTotals.length) {
    html += `<table><tr><th>Group</th><th class="right">Items</th><th class="right">Total</th></tr>`;
    groupTotals.forEach(g => { html += `<tr><td>${esc(g.group)}</td><td class="right">${g.items.length}</td><td class="right mono">${dollarFmt(g.total)}</td></tr>`; });
    html += `<tr style="font-weight:700;border-top:2px solid #e4e4e7"><td>Grand Total</td><td></td><td class="right mono">${dollarFmt(grandTotal)}</td></tr>`;
    html += `</table>`;
  }

  {
    const sub = grandTotal;
    const unit = scenario?.unitPrice ?? (qty > 0 ? sub / qty : 0);
    html += `<dl class="grid">`;
    html += `<dt>Subtotal</dt><dd>${dollarFmt(sub)}</dd>`;
    html += `<dt>Unit Price</dt><dd>${dollarFmt(unit)}</dd>`;
    html += `<dt>Quantity</dt><dd>${qty}</dd>`;
    if (scenario?.pcbSubtotal != null) html += `<dt>PCB Subtotal</dt><dd>${dollarFmt(scenario.pcbSubtotal)}</dd>`;
    if (scenario?.assemblySubtotal != null) html += `<dt>Assembly Subtotal</dt><dd>${dollarFmt(scenario.assemblySubtotal)}</dd>`;
    if (scenario?.bomSubtotal != null) html += `<dt>BOM Subtotal</dt><dd>${dollarFmt(scenario.bomSubtotal)}</dd>`;
    html += `</dl>`;
  }

  if (topDrivers.length) {
    html += `<h3>Top Pricing Line Items</h3>`;
    html += `<table><tr><th>Name</th><th>Group</th><th class="right">Unit</th><th class="right">Total</th></tr>`;
    topDrivers.forEach(p => {
      html += `<tr><td>${esc(p.name)}</td><td>${esc(p.group)}</td><td class="right mono">${money(p.unitPrice)}</td><td class="right mono">${money(p.total)}</td></tr>`;
    });
    html += `</table>`;
  }

  // === 7. PROFIT PROJECTIONS ===
  if (pp.length) {
    html += `<h2>7. Profit Projections</h2>`;
    if (ppTotal) {
      html += `<div class="kpi-strip">`;
      html += `<div class="kpi"><div class="val">${money(ppTotal.grossProfit)}</div><div class="lbl">Gross Profit</div></div>`;
      html += `<div class="kpi"><div class="val" style="color:${(ppTotal.margin ?? 0) < 0 ? '#dc2626' : '#16a34a'}">${ppTotal.margin?.toFixed(2) || '0.00'}%</div><div class="lbl">Margin</div></div>`;
      html += `</div>`;
    }
    if (ppSegments.length) {
      html += `<h3>Breakdown by Segment</h3>`;
      html += `<table><tr><th>Segment</th><th class="right">Gross Profit</th><th class="right">Margin</th></tr>`;
      ppSegments.forEach(p => {
        html += `<tr><td><b>${esc(p.type)}</b></td><td class="right mono">${money(p.grossProfit)}</td><td class="right" style="color:${(p.margin ?? 0) < 0 ? '#dc2626' : '#16a34a'}">${p.margin != null ? p.margin.toFixed(2) + '%' : '—'}</td></tr>`;
      });
      html += `</table>`;
    }
  }

  // === 8. SPEEDDFM ===
  html += `<h2 class="page-break">8. Manufacturing &amp; Assembly Checks (SpeedDFM)</h2>`;
  if (!dfm) {
    html += `<p style="font-size:10.5px;color:#854d0e;background:#fef9c3;border-left:4px solid #d97706;padding:10px 14px;border-radius:4px;margin-bottom:12px">SpeedDFM analysis was not available for this project. Common causes: the manufacturer your API key is tied to doesn't support SpeedDFM, the analysis errored upstream, or your API plan doesn't include SpeedDFM.</p>`;
  } else {
    const s = dfm.summary;
    const totalDfm = s.fabricationErrorCount + s.fabricationWarningCount + s.assemblyErrorCount + s.assemblyWarningCount + s.generalCount;
    html += `<p class="section-intro">Status: <b>${esc(s.analysisStatus)}</b>. Total violations: <b>${totalDfm}</b>. A full DFM PDF report is attached separately to this email.</p>`;
    html += `<div class="kpi-strip">`;
    html += `<div class="kpi"><div class="val" style="color:${s.fabricationErrorCount > 0 ? '#dc2626' : '#a1a1aa'}">${s.fabricationErrorCount}</div><div class="lbl">Fab Errors</div></div>`;
    html += `<div class="kpi"><div class="val" style="color:${s.fabricationWarningCount > 0 ? '#d97706' : '#a1a1aa'}">${s.fabricationWarningCount}</div><div class="lbl">Fab Warnings</div></div>`;
    html += `<div class="kpi"><div class="val" style="color:${s.assemblyErrorCount > 0 ? '#dc2626' : '#a1a1aa'}">${s.assemblyErrorCount}</div><div class="lbl">Asm Errors</div></div>`;
    html += `<div class="kpi"><div class="val" style="color:${s.assemblyWarningCount > 0 ? '#d97706' : '#a1a1aa'}">${s.assemblyWarningCount}</div><div class="lbl">Asm Warnings</div></div>`;
    html += `<div class="kpi"><div class="val">${s.generalCount}</div><div class="lbl">General</div></div>`;
    html += `</div>`;

    const dfmEntries = dfm.results || [];
    if (dfmEntries.length) {
      // Same display order as the DFM PDF and its image-tile picker.
      const sorted = [...dfmEntries].sort(compareDfmEntries);
      const topN = sorted.slice(0, 25);
      html += `<h3>Top ${topN.length} Violations (of ${dfmEntries.length})</h3>`;
      html += `<table><tr><th>Severity</th><th>Type</th><th>Rule</th><th>Layer</th><th class="right">x / y</th><th class="right">Measured / Threshold</th><th>Description</th></tr>`;
      topN.forEach(e => {
        const sevColor = e.severity === 'error' ? '#dc2626' : e.severity === 'warning' ? '#d97706' : '#64748b';
        const coords = e.x != null && e.y != null ? `${e.x.toFixed(2)} / ${e.y.toFixed(2)}` : '—';
        const meas = e.measuredValue != null && e.threshold != null
          ? `${e.measuredValue.toFixed(3)} / ${e.threshold.toFixed(3)} ${e.unit ?? ''}`
          : '—';
        html += `<tr><td style="color:${sevColor};font-weight:600">${esc(e.severity)}</td><td>${esc(e.ruleType)}</td><td class="mono">${esc(e.ruleName)}</td><td>${esc(e.layerName || e.layerType || '—')}</td><td class="right mono">${coords}</td><td class="right mono">${meas}</td><td>${esc((e.description || '').slice(0, 80))}</td></tr>`;
      });
      html += `</table>`;
      if (dfmEntries.length > topN.length) {
        html += `<p style="font-size:10px;color:#64748b">…${dfmEntries.length - topN.length} more violations in the attached DFM PDF report.</p>`;
      }
    } else if (s.analysisStatus === 'complete') {
      html += `<p style="font-size:11px;color:#16a34a;font-weight:600">No DFM violations detected.</p>`;
    }
  }

  // === 9. MANUFACTURER PRICING COMPARISON ===
  if (prices.length) {
    html += `<h2 class="page-break">9. Manufacturer Pricing Comparison</h2>`;
    html += `<p class="section-intro">Cross-manufacturer pricing for this project from <code>getProjectPrices</code>. Per-row sourcing (stock, distributor, lead time) is not part of this surface — see the BOM Excel for what the per-row export carries.</p>`;
    const priced = prices.filter(p => p.price && !p.errors?.length);
    const errored = prices.filter(p => !p.price || p.errors?.length);
    if (priced.length) {
      const sortedPriced = [...priced].sort((a, b) => (a.price?.subtotal ?? 0) - (b.price?.subtotal ?? 0));
      html += `<h3>Priced (${priced.length})</h3>`;
      html += `<table><tr><th>Manufacturer</th><th class="right">Qty</th><th class="right">Unit Price</th><th class="right">Subtotal</th><th>Currency</th><th>Project Type</th></tr>`;
      sortedPriced.forEach(p => {
        const cur = p.price?.currency ?? '';
        html += `<tr><td>${esc(p.mfrDisplayName || `mfrId=${p.mfrId}`)}</td><td class="right">${fmt(p.price?.quantity)}</td><td class="right mono">${dollarFmt(p.price?.unitPrice, 4)} ${esc(cur)}</td><td class="right mono">${dollarFmt(p.price?.subtotal)} ${esc(cur)}</td><td>${esc(cur || '—')}</td><td>${esc(p.price?.projectType || '—')}</td></tr>`;
      });
      html += `</table>`;
    }
    if (errored.length) {
      html += `<h3>Not Priced (${errored.length})</h3>`;
      html += `<table><tr><th>Manufacturer</th><th>Reason</th></tr>`;
      errored.forEach(p => {
        const reason = (p.errors || []).map(e => e?.code || e?.message || JSON.stringify(e)).join('; ') || 'No price returned';
        html += `<tr><td>${esc(p.mfrDisplayName || `mfrId=${p.mfrId}`)}</td><td style="color:#7f1d1d">${esc(reason)}</td></tr>`;
      });
      html += `</table>`;
    }
  }

  // === 10. PROVENANCE ===
  if (totalDecisions) {
    html += `<h2 class="page-break">10. Provenance — Decision Audit Trail</h2>`;
    html += `<p style="font-size:10px;color:#64748b;margin-bottom:12px">${totalDecisions} automated decisions across ${areas.length} areas. Every specification, sourcing choice, and pricing parameter determined by the Boardera API engine is recorded below.</p>`;
    html += `<table><tr><th>Area</th><th class="right">Decisions</th></tr>`;
    areas.forEach(a => { html += `<tr><td>${esc(a.area)}</td><td class="right">${a.count}</td></tr>`; });
    html += `</table>`;
    areas.forEach(a => {
      html += `<h3>${esc(a.area)} (${a.count})</h3>`;
      html += `<table><tr><th>Decision</th><th>Source</th><th>File</th></tr>`;
      a.entries.forEach(e => { html += `<tr><td>${esc(e.message)}</td><td class="mono">${esc(e.source)}</td><td class="mono">${esc(e.sourceFile || '')}</td></tr>`; });
      html += `</table>`;
    });
  }

  // === 11. PRICING CALCULATIONS ===
  if (calcs.length) {
    html += `<h2>11. Pricing Calculations (${calcs.length} parameters)</h2>`;
    html += `<p style="font-size:10px;color:#64748b;margin-bottom:8px">Named parameters used by the pricing engine.</p>`;
    html += `<table><tr><th>Parameter</th><th class="right">Value</th></tr>`;
    calcs.forEach(c => {
      const v = c.floatValue != null
        ? c.floatValue.toLocaleString(undefined, { maximumFractionDigits: 4 })
        : c.booleanValue != null ? (c.booleanValue ? 'Yes' : 'No') : '—';
      html += `<tr><td class="mono">${esc(c.name)}</td><td class="right mono">${v}</td></tr>`;
    });
    html += `</table>`;
  }

  html += `<div style="margin-top:20px;padding:12px 0;border-top:1px solid #e4e4e7;text-align:center"><p style="font-size:9px;color:#a1a1aa;margin:0">Component pricing and availability data provided by <a href="https://octopart.com" style="color:#71717a">Octopart</a> and <a href="https://www.cofactr.com" style="color:#71717a">Cofactr</a></p></div>`;
  html += `<div class="footer"><span class="copy">&copy; ${new Date().getFullYear()} Boardera Software Inc.</span><span>${esc(pN)} &middot; Project ID: ${esc(pid?.slice(0, 12) || '—')}</span><span>Boardera Decode Email Demo</span></div>`;
  html += `</div></body></html>`;

  return html;
}

export async function generateExecPdf(input: ExecPdfInput): Promise<string> {
  const html = buildExecHtml(input);
  mkdirSync(input.outputDir, { recursive: true });
  const outPath = resolve(input.outputDir, `${sanitizeName(input.productName)}_Analysis_Report.pdf`);

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
