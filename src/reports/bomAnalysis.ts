import type { ExportPricingLineItem, ExportProjectBomItem } from '../decode/types.js';

// Single source of truth for BOM row filtering + exception/warning
// classification, shared by the BOM Excel, the Exec PDF, and the email
// reply so the three surfaces can never disagree.

// Long lead time only matters when stock won't cover the order. Low
// stock below this many units is a warning (not an exception) as long
// as it still covers the order quantity.
export const LONG_LEAD_DAYS = 21;
export const LOW_STOCK_THRESHOLD = 100;

export function isEol(status?: string | null): boolean {
  return !!status && /eol|obsolete|end[\s_-]?of[\s_-]?life/i.test(status);
}

// True when a projectBom row represents a real, orderable/placeable part.
// The updated Decode API emits a placeholder row per DNP (Do-Not-Populate)
// component: no part number, no description, no reference designators, and
// quantityPerBoard 0. Those are not parts — drop them everywhere so they
// don't inflate part counts, costs, or exception tallies.
export function isRealBomRow(b: ExportProjectBomItem): boolean {
  const hasIdentity = !!(
    b.partNumber || b.pricedPartNumber || b.description || b.partManufacturer
    || b.internalPartNumber || b.customerPartNumber
  );
  const hasPlacement = (b.referenceDesignators?.length ?? 0) > 0 || (b.quantityPerBoard ?? 0) > 0;
  return hasIdentity || hasPlacement;
}

export function filterRealBom(bom: ExportProjectBomItem[]): ExportProjectBomItem[] {
  return bom.filter(isRealBomRow);
}

// Sourcing values can arrive either as top-level fields on `specs` (the
// common case) or only as a per-source breakdown under
// `specs.additionalDetails[].availability`. Prefer the rollup; fall back
// to the breakdown so we never miss data the API returned.
export type ResolvedSourcing = {
  hasSourcing: boolean;
  stock: number | null;
  leadTime: number | null;
  distributor: string | null;
  backordered: boolean;
  uncertainLead: boolean;
};

export function resolveSourcing(pli?: ExportPricingLineItem): ResolvedSourcing {
  const specs = pli?.specs;
  if (!specs) return { hasSourcing: false, stock: null, leadTime: null, distributor: null, backordered: false, uncertainLead: false };

  const addl = specs.additionalDetails || [];
  let stock = specs.stock ?? null;
  if (stock == null && addl.length) {
    const summed = addl.reduce((s, d) => s + (d.availability?.stock ?? 0), 0);
    if (summed > 0) stock = summed;
  }
  let leadTime = specs.leadTime ?? null;
  if (leadTime == null && addl.length) {
    const leads = addl
      .map(d => d.availability?.lead ?? d.best_case_lead ?? null)
      .filter((v): v is number => v != null);
    if (leads.length) leadTime = Math.min(...leads);
  }
  let distributor = specs.distributor ?? null;
  if (!distributor && addl.length) {
    const named = addl.find(d => d.source_name);
    if (named?.source_name) distributor = named.source_name;
  }

  const backordered = addl.some(d => d.is_backordered === true);
  const uncertainLead = addl.some(d => d.is_uncertain_lead === true);
  const hasSourcing = stock != null || leadTime != null || !!distributor || addl.length > 0
    || specs.stock != null || specs.leadTime != null || !!specs.distributor;

  return { hasSourcing, stock, leadTime, distributor, backordered, uncertainLead };
}

// Order quantity for the row = total ÷ unit price (only when unit price is
// a real positive number, so we never divide by a missing/zero price).
export function orderQtyOf(pli?: ExportPricingLineItem): number | null {
  return pli && pli.unitPrice?.value ? Math.round((pli.total?.value ?? 0) / pli.unitPrice.value) : null;
}

export type RowFlag = { code: string; label: string };
export type RowClassification = { exceptions: RowFlag[]; warnings: RowFlag[] };

// Per-row exception/warning rules. Exceptions are red and counted;
// warnings are amber and not counted as exceptions.
//
//   Exceptions: Customer supplied · EOL/Obsolete · Unfindable · No stock ·
//               Insufficient stock · Long lead (only when stock won't
//               cover the order) · Backordered · Uncertain lead · Substituted
//   Warnings:   Low stock (< LOW_STOCK_THRESHOLD but still covers the order)
//
// Lifecycle-unknown is intentionally NOT flagged. Sourcing-derived flags
// only fire when the API returned sourcing data AND the part is not
// customer-supplied (we don't source those).
export function classifyRow(b: ExportProjectBomItem, pli?: ExportPricingLineItem): RowClassification {
  const exceptions: RowFlag[] = [];
  const warnings: RowFlag[] = [];
  const sourcing = resolveSourcing(pli);
  const orderQty = orderQtyOf(pli);
  const stockCoversOrder = sourcing.stock != null && sourcing.stock > 0 && orderQty != null && sourcing.stock >= orderQty;

  if (b.customerSupplied) exceptions.push({ code: 'customerSupplied', label: 'Customer supplied' });
  if (isEol(b.lifeCycleStatus)) exceptions.push({ code: 'eol', label: 'EOL / Obsolete' });

  if (!b.customerSupplied && sourcing.hasSourcing) {
    if (!sourcing.distributor) exceptions.push({ code: 'unfindable', label: 'Unfindable' });
    if (sourcing.stock == null || sourcing.stock === 0) {
      exceptions.push({ code: 'noStock', label: 'No stock' });
    } else if (orderQty != null && sourcing.stock < orderQty) {
      exceptions.push({ code: 'insufficientStock', label: `Insufficient stock (${sourcing.stock} < ${orderQty})` });
    } else if (sourcing.stock < LOW_STOCK_THRESHOLD) {
      warnings.push({ code: 'lowStock', label: `Low stock (${sourcing.stock})` });
    }
    if (!stockCoversOrder && sourcing.leadTime != null && sourcing.leadTime > LONG_LEAD_DAYS) {
      exceptions.push({ code: 'longLead', label: `Long lead (${sourcing.leadTime}d)` });
    }
    if (sourcing.backordered) exceptions.push({ code: 'backordered', label: 'Backordered' });
    if (sourcing.uncertainLead) exceptions.push({ code: 'uncertainLead', label: 'Uncertain lead' });
  }

  if (b.substitutions) exceptions.push({ code: 'substituted', label: 'Substituted' });

  return { exceptions, warnings };
}
