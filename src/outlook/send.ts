import { readFileSync, statSync } from 'fs';
import { basename } from 'path';
import { graphFetch, graphJson } from './graph.js';

export type Attachment = { path: string; name?: string; contentType?: string };

// Microsoft Graph caps simple (single-POST) attachment uploads at ~3 MB
// per file. Report files rarely approach this; extend to the
// upload-session flow if yours do.
const PER_ATTACHMENT_LIMIT_BYTES = 3_000_000;

function guessContentType(name: string): string {
  const ext = name.toLowerCase().split('.').pop();
  switch (ext) {
    case 'pdf': return 'application/pdf';
    case 'xlsx': return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    case 'xls': return 'application/vnd.ms-excel';
    case 'csv': return 'text/csv';
    case 'zip': return 'application/zip';
    case 'json': return 'application/json';
    default: return 'application/octet-stream';
  }
}

// Sends the results as a FORWARD of the original message: the original
// body and attachments ride along natively (so the recipient sees the
// full customer email and design package), our results HTML is injected
// above the forwarded content, the report files are attached, and the
// subject is replaced with our own (e.g. "Decode results: <original>").
// Forward drafts also keep the conversation thread, so results group
// with the originating request in Outlook.
export async function forwardMessageWith(opts: {
  originalMessageId: string;
  to: string;
  bcc?: string[];
  subject: string;
  commentHtml: string; // body fragment (no <html>/<body> wrapper)
  attachments: Attachment[];
}): Promise<void> {
  for (const a of opts.attachments) {
    const size = statSync(a.path).size;
    if (size > PER_ATTACHMENT_LIMIT_BYTES) {
      throw new Error(
        `Attachment ${a.name ?? basename(a.path)} is ${(size / 1024 / 1024).toFixed(2)} MB, over the ` +
        `${PER_ATTACHMENT_LIMIT_BYTES / 1024 / 1024} MB simple-upload limit. ` +
        `Extend src/outlook/send.ts to the attachment upload-session flow for larger files.`,
      );
    }
  }

  // 1. Create the forward draft (carries the original body + attachments).
  const draft = await graphJson<{ id: string }>(
    `/me/messages/${opts.originalMessageId}/createForward`,
    { method: 'POST', body: '{}' },
  );

  // 2. Inject our results above the forwarded content.
  const current = await graphJson<{ body: { content: string } }>(
    `/me/messages/${draft.id}?$select=body`,
  );
  const original = current.body?.content || '';
  const fragment = `<div style="margin-bottom:16px">${opts.commentHtml}</div>`;
  const merged = /<body[^>]*>/i.test(original)
    ? original.replace(/<body[^>]*>/i, m => m + fragment)
    : fragment + original;

  // 3. Set recipient, subject, and merged body on the draft.
  await graphFetch(`/me/messages/${draft.id}`, {
    method: 'PATCH',
    body: JSON.stringify({
      subject: opts.subject,
      toRecipients: [{ emailAddress: { address: opts.to } }],
      bccRecipients: (opts.bcc ?? []).map(addr => ({ emailAddress: { address: addr } })),
      body: { contentType: 'HTML', content: merged },
    }),
  });

  // 4. Attach the report files.
  for (const a of opts.attachments) {
    const name = a.name ?? basename(a.path);
    const buf = readFileSync(a.path);
    await graphFetch(`/me/messages/${draft.id}/attachments`, {
      method: 'POST',
      body: JSON.stringify({
        '@odata.type': '#microsoft.graph.fileAttachment',
        name,
        contentType: a.contentType ?? guessContentType(name),
        contentBytes: buf.toString('base64'),
      }),
    });
  }

  // 5. Send (lands in Sent Items automatically).
  await graphFetch(`/me/messages/${draft.id}/send`, { method: 'POST' });
}
