import { resolve } from 'path';
import { mkdirSync } from 'fs';
// `xlsx-js-style` is published as CommonJS. With `"type": "module"` in
// package.json we have to take the default import so the namespace
// (.utils, .writeFile) is reachable.
import XLSX from 'xlsx-js-style';
import type {
  ExportProductData,
  ExportProjectBomItem,
  ExportPricingLineItem,
} from '../decode/types.js';
import {
  classifyRow,
  filterRealBom,
  isEol,
  orderQtyOf,
  resolveSourcing,
  type RowClassification,
} from './bomAnalysis.js';
import { sanitizeName } from '../utils/names.js';

function commaFmt(n: number | null | undefined, decimals: number): string {
  if (n == null || isNaN(n)) return '—';
  return Number(n).toLocaleString(undefined, { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
}
function dollarFmt(n: number | null | undefined, decimals = 2): string {
  if (n == null || isNaN(n)) return '—';
  return '$' + commaFmt(n, decimals);
}

function joinedRefDes(item: ExportProjectBomItem): { joined: string; count: number; dnpCount: number } {
  const arr = item.referenceDesignators || [];
  const dnpCount = arr.filter(r => r.dnp).length;
  return { joined: arr.map(r => r.designator).join(', '), count: arr.length, dnpCount };
}

export type BomReportInput = {
  productName: string;
  projectId: string;
  exportData: ExportProductData;
  quantities: number[];
  outputDir: string;
};

type Joined = {
  bomItem: ExportProjectBomItem;
  pricing?: ExportPricingLineItem;
  cls: RowClassification;
};

export function generateBomExcel(input: BomReportInput): string | null {
  const { productName: pN, projectId: pid, exportData, quantities, outputDir } = input;

  // Drop empty DNP placeholder rows (no part / no placement) the API emits.
  const bom = filterRealBom(exportData.specs?.projectBom || []);
  if (!bom.length) return null;

  const scenario = exportData.costing?.scenarios?.[0];
  const qty = quantities?.[0] || scenario?.quantity || 1;
  const pricingByName: Record<string, ExportPricingLineItem> = {};
  (scenario?.pricingLineItems || []).forEach(pli => {
    if (pli.group === 'BOM' && pli.name) pricingByName[pli.name] = pli;
  });

  const subByPart: Record<string, string> = {};
  (scenario?.validationDetails?.assemblyWithSubstitutions || []).forEach(s => {
    if (s.primaryPartNumber) subByPart[s.primaryPartNumber] = s.subPartNumber;
  });

  const joined: Joined[] = bom.map(b => {
    const pricing = pricingByName[b.partNumber || ''] || pricingByName[b.pricedPartNumber || ''];
    const cls = classifyRow(b, pricing);
    // Enrich the substitution label with the target part when known.
    const target = b.partNumber ? subByPart[b.partNumber] : undefined;
    if (target) {
      const sub = cls.exceptions.find(e => e.code === 'substituted');
      if (sub) sub.label = `Substituted → ${target}`;
      else cls.exceptions.push({ code: 'substituted', label: `Substituted → ${target}` });
    }
    return { bomItem: b, pricing, cls };
  });

  const bomTotal = joined.reduce((s, j) => s + (j.pricing?.total?.value || 0), 0);
  const smtCount = joined.filter(j => j.bomItem.mountType === 'Surface Mount').length;
  const thCount = joined.filter(j => j.bomItem.mountType === 'Through Hole').length;
  const eolCount = joined.filter(j => isEol(j.bomItem.lifeCycleStatus)).length;
  const exceptionRows = joined.filter(j => j.cls.exceptions.length > 0).length;
  const warningRows = joined.filter(j => j.cls.exceptions.length === 0 && j.cls.warnings.length > 0).length;

  // Sort: exception rows first, then warning-only rows, then by total cost desc.
  const rank = (j: Joined) => (j.cls.exceptions.length ? 0 : j.cls.warnings.length ? 1 : 2);
  const sorted = [...joined].sort((a, b) => {
    const r = rank(a) - rank(b);
    if (r !== 0) return r;
    return (b.pricing?.total?.value || 0) - (a.pricing?.total?.value || 0);
  });

  const headers = [
    '#', 'Part Number', 'Priced As', 'Manufacturer', 'Description',
    'Ref Designators', 'Mount Type', 'Qty/Board', 'Order Qty',
    'Unit Price (USD)', 'Total Price (USD)', 'Stock', 'Lead (days)',
    'Lifecycle', 'RoHS', 'Distributor', 'Source URL', 'Package', 'Exceptions / Warnings',
  ];

  const rows: any[][] = [];
  rows.push([pN + ' — BOM Report']);
  rows.push(['Project: ' + (pid?.slice(0, 16) || '—'), '', 'Quantity: ' + qty, '', 'Parts: ' + bom.length, '', 'Generated: ' + new Date().toLocaleDateString()]);
  rows.push(['BOM Total: ' + dollarFmt(bomTotal, 2), '', 'SMT / TH: ' + smtCount + ' / ' + thCount, '', 'EOL Parts: ' + eolCount, '', 'Rows w/ Exceptions: ' + exceptionRows + ' | Warnings: ' + warningRows]);
  rows.push([]);
  rows.push(headers);

  sorted.forEach((j, i) => {
    const b = j.bomItem;
    const p = j.pricing;
    const sourcing = resolveSourcing(p);
    const refDes = joinedRefDes(b);
    const refDesText = refDes.dnpCount
      ? `${refDes.joined.slice(0, 100)} (${refDes.dnpCount} DNP)`
      : refDes.joined.slice(0, 120);

    const orderQty = orderQtyOf(p);

    rows.push([
      i + 1,
      b.partNumber || '',
      b.pricedPartNumber && b.pricedPartNumber !== b.partNumber ? b.pricedPartNumber : '',
      b.partManufacturer || '',
      (b.description || '').slice(0, 80),
      refDesText,
      b.mountType || '',
      b.quantityPerBoard ?? 1,
      orderQty ?? '',
      p?.unitPrice?.value != null ? Number(p.unitPrice.value) : null,
      p?.total?.value != null ? Number(p.total.value) : null,
      sourcing.stock != null ? sourcing.stock : '',
      sourcing.leadTime != null ? sourcing.leadTime : '',
      isEol(b.lifeCycleStatus) ? 'EOL' : (b.lifeCycleStatus || ''),
      b.rohsStatus === 'Compliant' ? 'Yes' : (b.rohsStatus ? 'No' : ''),
      sourcing.distributor || '',
      p?.specs?.clickUrl || '',
      b.potentialPackage || '',
      [...j.cls.exceptions.map(e => e.label), ...j.cls.warnings.map(w => w.label)].join('; '),
    ]);
  });

  const wb = XLSX.utils.book_new();
  const ws: any = XLSX.utils.aoa_to_sheet(rows);

  ws['!cols'] = [
    { wch: 4 }, { wch: 22 }, { wch: 18 }, { wch: 20 }, { wch: 35 },
    { wch: 30 }, { wch: 14 }, { wch: 9 }, { wch: 10 },
    { wch: 14 }, { wch: 14 }, { wch: 10 }, { wch: 9 },
    { wch: 14 }, { wch: 7 }, { wch: 16 }, { wch: 30 }, { wch: 12 }, { wch: 40 },
  ];
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: headers.length - 1 } }];

  const headerRow = 4;
  const dataStart = 5;

  const fillRedExc = { patternType: 'solid', fgColor: { rgb: 'FEE2E2' } };
  const fillAmber = { patternType: 'solid', fgColor: { rgb: 'FFFBEB' } };
  const fillHeader = { fgColor: { rgb: '0F172A' } };
  const fontWhite = { color: { rgb: 'FFFFFF' }, bold: true, sz: 9 };
  const fontBold = { bold: true };
  const borderThin = {
    top: { style: 'thin', color: { rgb: 'D1D5DB' } },
    bottom: { style: 'thin', color: { rgb: 'D1D5DB' } },
    left: { style: 'thin', color: { rgb: 'D1D5DB' } },
    right: { style: 'thin', color: { rgb: 'D1D5DB' } },
  };

  for (let c = 0; c < headers.length; c++) {
    const ref = XLSX.utils.encode_cell({ r: headerRow, c });
    if (!ws[ref]) continue;
    ws[ref].s = { fill: fillHeader, font: fontWhite, alignment: { horizontal: 'center', vertical: 'center' }, border: borderThin };
  }
  const titleRef = XLSX.utils.encode_cell({ r: 0, c: 0 });
  if (ws[titleRef]) ws[titleRef].s = { font: { bold: true, sz: 14, color: { rgb: '0F172A' } }, alignment: { vertical: 'center' } };
  for (let r = 1; r <= 2; r++) {
    for (let c = 0; c < headers.length; c++) {
      const ref = XLSX.utils.encode_cell({ r, c });
      if (ws[ref]) ws[ref].s = { fill: { fgColor: { rgb: 'F8FAFC' } }, font: { sz: 9, color: { rgb: '334155' } }, border: borderThin };
    }
  }

  sorted.forEach((j, i) => {
    const r = dataStart + i;
    const hasException = j.cls.exceptions.length > 0;
    const hasWarning = j.cls.warnings.length > 0;
    const b = j.bomItem;
    const p = j.pricing;
    const rowSourcing = resolveSourcing(p);
    // Red for any exception row; amber for warning-only rows (low stock).
    const rowFill = hasException ? fillRedExc : hasWarning ? fillAmber : null;

    for (let c = 0; c < headers.length; c++) {
      const ref = XLSX.utils.encode_cell({ r, c });
      if (!ws[ref]) continue;
      const style: any = { border: borderThin };
      if (rowFill) style.fill = rowFill;

      if (c === 9 && ws[ref].v != null) { ws[ref].t = 'n'; ws[ref].z = '$#,##0.0000'; }
      if (c === 10 && ws[ref].v != null) { ws[ref].t = 'n'; ws[ref].z = '$#,##0.00'; style.font = fontBold; }
      if (c === 13) {
        if (isEol(b.lifeCycleStatus)) style.font = { bold: true, color: { rgb: '991B1B' } };
        else if (ws[ref].v) style.font = { color: { rgb: '166534' } };
      }
      if (c === 14) {
        if (ws[ref].v === 'No') style.font = { bold: true, color: { rgb: '991B1B' } };
        else if (ws[ref].v === 'Yes') style.font = { color: { rgb: '166534' } };
      }
      if (c === 18 && ws[ref].v) style.font = { color: { rgb: hasException ? 'DC2626' : 'B45309' }, sz: 9 };
      if ([0, 7, 8, 9, 10, 11, 12].includes(c)) style.alignment = { horizontal: 'right' };
      if ([6, 13, 14, 17].includes(c)) style.alignment = { horizontal: 'center' };

      ws[ref].s = style;
    }

    const urlRef = XLSX.utils.encode_cell({ r, c: 16 });
    const url = p?.specs?.clickUrl;
    if (url && ws[urlRef]) {
      ws[urlRef].l = { Target: url, Tooltip: rowSourcing.distributor || 'Open source' };
      ws[urlRef].v = rowSourcing.distributor || url;
      ws[urlRef].s = { ...ws[urlRef].s, font: { color: { rgb: '2563EB' }, underline: true } };
    }
  });

  ws['!freeze'] = { xSplit: 0, ySplit: 5 };
  if (!ws['!views']) ws['!views'] = [{}];
  ws['!views'][0] = { state: 'frozen', ySplit: 5 };
  ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: headerRow, c: 0 }, e: { r: rows.length - 1, c: headers.length - 1 } }) };

  const attrRow = rows.length + 1;
  const attrRef = XLSX.utils.encode_cell({ r: attrRow, c: 0 });
  ws[attrRef] = { v: 'Component pricing and availability data provided by Octopart (octopart.com) and Cofactr (cofactr.com)', t: 's', s: { font: { italic: true, color: { rgb: 'A1A1AA' }, sz: 9 } } };

  XLSX.utils.book_append_sheet(wb, ws, 'BOM Report');

  mkdirSync(outputDir, { recursive: true });
  const outPath = resolve(outputDir, `${sanitizeName(pN)}_BOM_Report.xlsx`);
  XLSX.writeFile(wb, outPath, { cellStyles: true, bookSST: true });
  return outPath;
}
