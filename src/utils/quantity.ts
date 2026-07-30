// Matches a quantity in either of two shapes, case-insensitive:
//
//   1. Keyword before the number — keyword (`quantity`, `qty`, `boards`)
//      then an optional connector (`:`, `=`, the word `of`, or just
//      whitespace), then the number. Captured in group 1.
//        "qty 25"  "qty: 25"  "qty=25"  "quantity 25"  "quantity: 25"
//        "quantity of 25"  "boards 25"  "boards: 25"
//
//   2. Unit after the number — the number followed by a pieces unit
//      (`pc`, `pcs`, `piece`, `pieces`). Captured in group 2.
//        "100 pcs"  "100pcs"  "100 pieces"  "100pc"
//
// Leading \b keeps it from firing inside another word (e.g. NOT
// "keyboards", "antiquity", or the "10" in "10PCB").
const QTY_REGEX = /\b(?:quantity|qty|boards)\b\s*(?:of\b|[:=])?\s*(\d{1,5})|\b(\d{1,5})\s*(?:pcs?|pieces?)\b/i;

export function extractQuantity(subject: string | undefined, bodyText: string | undefined, fallback: number): number {
  for (const source of [subject, bodyText]) {
    if (!source) continue;
    const m = source.match(QTY_REGEX);
    if (m) {
      // Group 1 = keyword form, group 2 = "<n> pcs" form; only one matches.
      const n = parseInt(m[1] ?? m[2], 10);
      if (Number.isFinite(n) && n > 0) return n;
    }
  }
  return fallback;
}

export function stripHtml(html: string | undefined): string {
  if (!html) return '';
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}
