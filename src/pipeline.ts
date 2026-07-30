import { mkdirSync, unlinkSync, writeFileSync } from 'fs';
import { basename, resolve } from 'path';
import { config } from './config.js';
import { downloadAttachmentsToTemp, inputDirForMessage, outputDirForMessage, type SavedAttachment } from './outlook/attachments.js';
import { markMessageRead, type GraphMessage } from './outlook/inbox.js';
import { forwardMessageWith } from './outlook/send.js';
import { runDecodePipeline } from './decode/pipeline.js';
import { isArchiveFile } from './decode/pipeline.js';
import { generateBomExcel } from './reports/bom-excel.js';
import { generateExecPdf } from './reports/exec-pdf.js';
import { generateDfmPdf } from './reports/dfm-pdf.js';
import { zipDirectory } from './utils/zip.js';
import { extractQuantity, stripHtml } from './utils/quantity.js';
import { sanitizeName } from './utils/names.js';
import { classifyRow, filterRealBom, isEol } from './reports/bomAnalysis.js';
import type { DfmResults, ExportProductData } from './decode/types.js';

export type ProcessOptions = { dryRun?: boolean };

// Optional BCC on every reply (OUTLOOK_REPLY_BCC). This address is
// blind-copied (the recipient does not see it in the headers).
const REPLY_BCC = config.outlook.replyBcc;

// Resolve where the result/error reply goes. OUTLOOK_REPLY_TO is normally
// the quoting team's mailbox — the reports carry internal pricing and
// margin data, so the RFQ sender (the customer) should not see them by
// default. The literal value "sender" opts into replying to the original
// sender instead.
function resolveReplyRecipient(sender: string): string {
  const configured = config.outlook.replyTo.trim();
  if (configured.toLowerCase() !== 'sender') return configured;
  if (!sender || sender === 'unknown') {
    throw new Error('OUTLOOK_REPLY_TO=sender, but the message has no resolvable sender address.');
  }
  return sender;
}

// Subjects of the tool's own outgoing replies. If the Outlook rule that
// feeds the watched folder also matches the reply recipient (e.g. the
// quoting team's mailbox lives in the same account), our replies land
// back in the queue — processing them would re-analyze our own output
// bundles in an endless loop. Skip them by subject prefix.
const OWN_REPLY_PREFIXES = ['Decode results:', 'Decode error:'];

export async function processMessage(message: GraphMessage, opts: ProcessOptions = {}): Promise<void> {
  const dryRun = !!opts.dryRun;
  const messageId = message.id;
  const subject = message.subject || '(no subject)';
  const sender = message.from?.emailAddress.address || 'unknown';

  if (OWN_REPLY_PREFIXES.some(p => subject.startsWith(p))) {
    console.log(`\n=== Skipping own reply "${subject}" (${messageId.slice(0, 20)}…) — marking read ===`);
    try { await markMessageRead(messageId); }
    catch (markErr) { console.error(`  WARN: failed to mark own reply as read: ${(markErr as Error).message}`); }
    return;
  }

  console.log(`\n=== Processing message ${messageId} ===`);
  console.log(`  From: ${sender}`);
  console.log(`  Subject: ${subject}`);

  try {
    await runOne(message, dryRun);
  } catch (err) {
    const errorMessage = (err as Error).message || String(err);
    console.error(`  Failed to process message: ${errorMessage}`);
    if (dryRun) {
      console.log('  DRY RUN — skipping error reply.');
    } else {
      try {
        await sendErrorReply(message, errorMessage);
      } catch (sendErr) {
        console.error(`  WARN: also failed to send error reply: ${(sendErr as Error).message}`);
      }
    }
  }

  // Mark as read regardless of outcome — the queue must advance. On
  // success the user gets the report email; on failure they get the
  // error email. To retry, the user marks the message unread again in
  // Outlook.
  try {
    await markMessageRead(messageId);
    console.log(`  Marked source message as read`);
  } catch (markErr) {
    console.error(`  WARN: failed to mark message as read: ${(markErr as Error).message}`);
  }
}

async function runOne(message: GraphMessage, dryRun: boolean): Promise<void> {
  const messageId = message.id;
  const subject = message.subject || '(no subject)';
  const sender = message.from?.emailAddress.address || 'unknown';

  const inputDir = inputDirForMessage(messageId);
  const outputDir = outputDirForMessage(messageId);
  mkdirSync(inputDir, { recursive: true });
  mkdirSync(outputDir, { recursive: true });

  // 1. Download attachments
  const saved = await downloadAttachmentsToTemp(messageId);
  if (!saved.length) {
    throw new Error(`Message has no usable file attachments. Please re-send with the design package attached.`);
  }
  console.log(`  Downloaded ${saved.length} attachment(s) (${(saved.reduce((s, a) => s + a.size, 0) / 1024).toFixed(1)} KB total)`);

  // 2. Archive policy: if the email contains one or more archives
  //    (.zip / .tar / .tar.gz / .tgz), keep exactly one — Decode's
  //    archiveExtractor unpacks it server-side. Loose files alongside
  //    the archive are dropped (the archive is assumed to be the full
  //    package). If no archive is present, send all loose files as-is.
  pruneToSingleArchiveIfPresent(saved);

  // 3. Determine quantity
  const bodyText = message.body?.contentType === 'html' ? stripHtml(message.body.content) : (message.body?.content || message.bodyPreview);
  const quantity = extractQuantity(subject, bodyText, config.defaultQuantity);
  console.log(`  Quantity: ${quantity}${quantity !== config.defaultQuantity ? ' (from email)' : ' (default)'}`);

  // 4. Run the Decode pipeline
  const projectName = subject.replace(/[\r\n]+/g, ' ').slice(0, 80) || `Email ${messageId.slice(0, 8)}`;
  const result = await runDecodePipeline({ inputDir, projectName, quantity, outputDir, imageLimit: config.dfm.imageLimit });

  // Surface the case where the uploader sent a BOM but Decode couldn't
  // extract any rows from it. Common cause: non-standard CSV column names
  // (e.g. an embedded comma in a column header) that defeat the BOM
  // classifier upstream.
  const sentBom = (result.files || []).some(f => f.type === 'Bill Of Material');
  const gotBomRows = (result.exportData.specs?.projectBom?.length ?? 0) > 0;
  if (sentBom && !gotBomRows) {
    console.warn(`  WARN: BOM file was uploaded but Decode returned 0 BOM rows. Check the BOM CSV column headers (e.g. embedded commas, non-standard names).`);
  }

  // 5. Save the raw export JSON: the exportProduct payload with the
  //    project _id prepended and the DFM results + manufacturer prices
  //    appended, so one file carries the complete result.
  const safeName = sanitizeName(result.productName);
  const jsonPath = resolve(outputDir, `${safeName}_export.json`);
  const jsonPayload = {
    _id: result.projectId,
    ...result.exportData,
    dfm: result.dfm,
    prices: result.prices,
  };
  writeFileSync(jsonPath, JSON.stringify(jsonPayload, null, 2));
  console.log(`    Wrote ${jsonPath}`);

  // 6. Generate the reports
  console.log('  Generating BOM Excel...');
  const bomPath = generateBomExcel({
    productName: result.productName,
    projectId: result.projectId,
    exportData: result.exportData,
    quantities: [result.quantity],
    outputDir,
  });
  if (!bomPath) {
    console.log('  No BOM in export data — skipping BOM Excel');
  } else {
    console.log(`    Wrote ${bomPath}`);
  }

  console.log('  Generating Executive Summary PDF...');
  const pdfPath = await generateExecPdf({
    productName: result.productName,
    projectId: result.projectId,
    exportData: result.exportData,
    dfm: result.dfm,
    prices: result.prices,
    quantities: [result.quantity],
    outputDir,
  });
  console.log(`    Wrote ${pdfPath}`);

  // 6b. Generate the demo's own DFM report PDF (summary, by-rule rollup,
  //     per-violation image gallery, full violations table). Skipped when
  //     DFM was unavailable. Best-effort: a render failure logs and does
  //     not block the reply.
  let dfmReportPath: string | null = null;
  if (result.dfm) {
    console.log('  Generating DFM Report PDF...');
    try {
      dfmReportPath = await generateDfmPdf({
        productName: result.productName,
        projectId: result.projectId,
        quantity: result.quantity,
        dfm: result.dfm,
        images: result.dfmImages,
        outputDir,
      });
      console.log(`    Wrote ${dfmReportPath}`);
    } catch (err) {
      console.warn(`  WARN: failed to generate DFM report: ${(err as Error).message}`);
    }
  }

  // 7. Bundle the original files as a zip — a local artifact only; the
  //    outgoing forward carries the customer's original attachments
  //    natively, so this is not attached to the email.
  const originalZipPath = resolve(outputDir, `${safeName}_original_files.zip`);
  await zipDirectory(inputDir, originalZipPath);
  console.log(`    Wrote ${originalZipPath}`);

  // 8. Send the results as a FORWARD of the original message to the
  //    configured recipient (the quoting team's mailbox, or the original
  //    sender when OUTLOOK_REPLY_TO=sender), BCC'ing REPLY_BCC. The
  //    forward carries the customer's message and design package
  //    natively, threads with the original in Outlook, and gets our
  //    report files attached on top. In dry-run mode this is skipped —
  //    the reports stay under tmp/<messageId>/out/ for inspection.
  if (dryRun) {
    console.log(`  DRY RUN — skipping reply. Reports are under ${outputDir}`);
    return;
  }
  const recipient = resolveReplyRecipient(sender);
  // One list drives both the real attachments and the "Attached:" section
  // of the reply body, so the listed names always match the files sent.
  const reportAttachments = [
    ...(bomPath ? [{ path: bomPath, desc: 'cleansed/costed BOM noting exceptions' }] : []),
    { path: pdfPath, desc: 'executive summary' },
    ...(dfmReportPath ? [{ path: dfmReportPath, desc: 'SpeedDFM manufacturing & assembly violations' }] : []),
    { path: jsonPath, desc: 'raw Decode API export (includes DFM & manufacturer pricing)' },
  ];
  const replyHtml = buildReplyBody({
    productName: result.productName,
    projectId: result.projectId,
    quantity: result.quantity,
    exportData: result.exportData,
    dfm: result.dfm,
    sender,
    attachments: reportAttachments.map(a => ({ name: basename(a.path), desc: a.desc })),
  });
  const attachments = reportAttachments.map(a => ({ path: a.path }));
  const bccAddresses = REPLY_BCC && recipient !== REPLY_BCC ? [REPLY_BCC] : [];
  await forwardMessageWith({
    originalMessageId: messageId,
    to: recipient,
    bcc: bccAddresses,
    subject: `Decode results: ${subject}`,
    commentHtml: replyHtml,
    attachments,
  });
  console.log(`  Forwarded results to ${recipient}${bccAddresses.length ? ' (bcc: ' + bccAddresses.join(', ') + ')' : ''} with ${attachments.length} report attachment(s)`);
  console.log(`  Project ID: ${result.projectId}`);
}

async function sendErrorReply(message: GraphMessage, errorMessage: string): Promise<void> {
  const sender = message.from?.emailAddress.address || 'unknown';
  let recipient: string;
  try {
    recipient = resolveReplyRecipient(sender);
  } catch (err) {
    console.warn(`  WARN: cannot send error reply — ${(err as Error).message}`);
    return;
  }
  const subject = message.subject || '(no subject)';
  const bccAddresses = REPLY_BCC && recipient !== REPLY_BCC ? [REPLY_BCC] : [];

  await forwardMessageWith({
    originalMessageId: message.id,
    to: recipient,
    bcc: bccAddresses,
    subject: `Decode error: ${subject}`,
    commentHtml: buildErrorReplyBody({ sender, errorMessage }),
    attachments: [],
  });
  console.log(`  Forwarded error notice to ${recipient}${bccAddresses.length ? ' (bcc: ' + bccAddresses.join(', ') + ')' : ''}`);
}

// Body FRAGMENT injected above the forwarded original message (which
// carries the customer's email and files itself — no need to quote it).
function buildErrorReplyBody(args: { sender: string; errorMessage: string }): string {
  const { sender, errorMessage } = args;
  return `<div style="font-family:-apple-system,Segoe UI,Roboto,sans-serif;color:#1e293b;font-size:13px;line-height:1.6">
<p>Hi,</p>
<p>The design package from <b>${escapeHtml(sender)}</b> (forwarded below) could not be processed. The Decode pipeline reported the following error:</p>
<blockquote style="margin:12px 0;padding:12px 16px;background:#fef2f2;border-left:4px solid #dc2626;color:#7f1d1d;font-family:monospace;font-size:12px;white-space:pre-wrap">${escapeHtml(errorMessage)}</blockquote>
<p>A few things to check before re-sending:</p>
<ul>
  <li>The package should include gerber/drill files, an optional BOM (CSV or Excel), and any supporting PDFs — sent as attachments or inside a single .zip.</li>
  <li>BOM CSVs with embedded commas in column header names (e.g. <code>"Reference, Designator"</code>) are known to confuse the BOM classifier. Use simple headers like <code>Part Number</code>, <code>Reference Designators</code>, <code>Description</code>.</li>
  <li>To retry, re-send a fresh email with the corrected package to the intake address — replies to this message are not processed.</li>
</ul>
<p style="color:#94a3b8;font-size:11px;margin-top:24px">Generated automatically by the Boardera Decode email demo. The original message follows below.</p>
</div>`;
}

// Mutates `attachments` in place. If the email had at least one
// archive, keeps the first archive and deletes every other downloaded
// file (archives and loose files alike) from disk. If no archive, all
// loose files are kept.
function pruneToSingleArchiveIfPresent(attachments: SavedAttachment[]): void {
  const archives = attachments.filter(a => isArchiveFile(a.name));
  if (!archives.length) return;

  const kept = archives[0];
  const toRemove = attachments.filter(a => a !== kept);
  for (const a of toRemove) {
    try { unlinkSync(a.path); } catch { /* file may already be gone */ }
  }
  attachments.length = 0;
  attachments.push(kept);

  if (archives.length > 1) {
    console.log(`  Multiple archives in email (${archives.length}); using '${kept.name}', dropping ${archives.length - 1} other archive(s).`);
  }
  if (toRemove.length > archives.length - 1) {
    console.log(`  Dropped ${toRemove.length - (archives.length - 1)} loose attachment(s) alongside the archive.`);
  }
  console.log(`  Forwarding archive '${kept.name}' (${(kept.size / 1024).toFixed(1)} KB) to Decode for server-side extraction.`);
}

function buildReplyBody(args: {
  productName: string;
  projectId: string;
  quantity: number;
  exportData: ExportProductData;
  dfm: DfmResults | null;
  sender: string;
  attachments: { name: string; desc: string }[];
}): string {
  const { productName, projectId, quantity, exportData, dfm, sender, attachments } = args;
  // Drop empty DNP placeholder rows so counts match the reports exactly.
  const bom = filterRealBom(exportData.specs?.projectBom || []);
  const scenario = exportData.costing?.scenarios?.[0];
  const pli = scenario?.pricingLineItems || [];
  const grandTotal = scenario?.subtotal ?? pli.reduce((s, p) => s + (p.total?.value || 0), 0);
  const bomTotal = pli.filter(p => p.group === 'BOM').reduce((s, p) => s + (p.total?.value || 0), 0);

  const pricingByName: Record<string, typeof pli[0]> = {};
  pli.forEach(p => { if (p.group === 'BOM' && p.name) pricingByName[p.name] = p; });
  // Exception/warning counts use the shared classifier so the email
  // matches the BOM Excel and Exec PDF exactly.
  const classified = bom.map(b => classifyRow(b, pricingByName[b.partNumber || ''] || pricingByName[b.pricedPartNumber || '']));
  const exceptions = classified.filter(c => c.exceptions.length > 0).length;
  const warnings = classified.filter(c => c.exceptions.length === 0 && c.warnings.length > 0).length;
  const eolCount = bom.filter(b => isEol(b.lifeCycleStatus)).length;
  const dfmFabErr = dfm?.summary?.fabricationErrorCount ?? 0;
  const dfmAsmErr = dfm?.summary?.assemblyErrorCount ?? 0;
  const dfmFabWarn = dfm?.summary?.fabricationWarningCount ?? 0;
  const dfmAsmWarn = dfm?.summary?.assemblyWarningCount ?? 0;

  const dollar = (n: number) => '$' + (n || 0).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  return `<div style="font-family:-apple-system,Segoe UI,Roboto,sans-serif;color:#1e293b;font-size:13px;line-height:1.6">
<p>Hi,</p>
<p>The Boardera Decode API has finished processing the design package forwarded below. The full reports are attached.</p>
<table style="border-collapse:collapse;margin:16px 0;font-size:12px">
  <tr><td style="padding:4px 12px 4px 0;color:#64748b">Received from</td><td>${escapeHtml(sender)}</td></tr>
  <tr><td style="padding:4px 12px 4px 0;color:#64748b">Project</td><td style="font-weight:600">${escapeHtml(productName)}</td></tr>
  <tr><td style="padding:4px 12px 4px 0;color:#64748b">Project ID</td><td style="font-family:monospace">${escapeHtml(projectId)}</td></tr>
  <tr><td style="padding:4px 12px 4px 0;color:#64748b">Quantity</td><td>${quantity}</td></tr>
  <tr><td style="padding:4px 12px 4px 0;color:#64748b">BOM parts</td><td>${bom.length}</td></tr>
  <tr><td style="padding:4px 12px 4px 0;color:#64748b">BOM total</td><td>${dollar(bomTotal)}</td></tr>
  <tr><td style="padding:4px 12px 4px 0;color:#64748b">Quoted total</td><td><b>${dollar(grandTotal)}</b></td></tr>
  <tr><td style="padding:4px 12px 4px 0;color:#64748b">Exceptions</td><td style="color:${exceptions > 0 ? '#dc2626' : '#16a34a'}">${exceptions}</td></tr>
  <tr><td style="padding:4px 12px 4px 0;color:#64748b">Warnings</td><td style="color:${warnings > 0 ? '#d97706' : '#16a34a'}">${warnings}</td></tr>
  <tr><td style="padding:4px 12px 4px 0;color:#64748b">EOL parts</td><td style="color:${eolCount > 0 ? '#dc2626' : '#16a34a'}">${eolCount}</td></tr>
  ${dfm ? `<tr><td style="padding:4px 12px 4px 0;color:#64748b">DFM Fab Err/Warn</td><td style="color:${dfmFabErr > 0 ? '#dc2626' : '#16a34a'}">${dfmFabErr} / ${dfmFabWarn}</td></tr>
  <tr><td style="padding:4px 12px 4px 0;color:#64748b">DFM Asm Err/Warn</td><td style="color:${dfmAsmErr > 0 ? '#dc2626' : '#16a34a'}">${dfmAsmErr} / ${dfmAsmWarn}</td></tr>` : `<tr><td style="padding:4px 12px 4px 0;color:#64748b">DFM</td><td style="color:#94a3b8">unavailable</td></tr>`}
</table>
<p>Attached:</p>
<ul>
  ${attachments.map(a => `<li><b>${escapeHtml(a.name)}</b> — ${escapeHtml(a.desc)}</li>`).join('\n  ')}
</ul>
<p style="color:#94a3b8;font-size:11px;margin-top:24px">Generated by the Boardera Decode email demo. The original message and design package follow below.</p>
</div>`;
}

function escapeHtml(s: string): string {
  return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
