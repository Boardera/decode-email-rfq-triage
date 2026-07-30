import { createHash } from 'crypto';
import { mkdirSync, rmSync, writeFileSync } from 'fs';
import { join, resolve } from 'path';
import { TMP_DIR } from '../config.js';
import { graphJson } from './graph.js';

type GraphFileAttachment = {
  '@odata.type': string;
  id: string;
  name: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  contentBytes?: string;
};

export type SavedAttachment = {
  name: string;
  path: string;
  size: number;
};

export function inputDirForMessage(messageId: string): string {
  return resolve(TMP_DIR, messageFolder(messageId), 'in');
}

export function outputDirForMessage(messageId: string): string {
  return resolve(TMP_DIR, messageFolder(messageId), 'out');
}

// Per-message folder name. Microsoft Graph message IDs are ~150 chars and
// share a long common prefix within a mailbox folder, so a plain
// prefix-slice collides across different messages — which would pile
// every package into one shared folder and cross-contaminate uploads.
// We keep a short readable prefix for humans AND append a hash of the
// FULL id so distinct messages always get distinct folders.
function messageFolder(messageId: string): string {
  const slug = messageId.replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 60);
  const hash = createHash('sha1').update(messageId).digest('hex').slice(0, 12);
  return `${slug}-${hash}`;
}

export async function downloadAttachmentsToTemp(messageId: string): Promise<SavedAttachment[]> {
  const inputDir = inputDirForMessage(messageId);
  // Start from a clean input dir. If this message was processed before
  // (e.g. marked unread and retried), stale attachments must not linger
  // and get re-uploaded — the API would analyze them and emit their
  // notes/provenance into this project's results.
  rmSync(inputDir, { recursive: true, force: true });
  mkdirSync(inputDir, { recursive: true });

  // No $select: Graph rejects `contentBytes` in $select because it's only on
  // the derived fileAttachment type. The default representation already
  // includes contentBytes for file attachments.
  const list = await graphJson<{ value: GraphFileAttachment[] }>(
    `/me/messages/${messageId}/attachments`,
  );

  const saved: SavedAttachment[] = [];
  for (const a of list.value) {
    if (a.isInline) continue;
    if (a['@odata.type'] !== '#microsoft.graph.fileAttachment') {
      console.warn(`  Skipping non-file attachment: ${a.name} (${a['@odata.type']})`);
      continue;
    }
    let contentBytes = a.contentBytes;
    if (!contentBytes) {
      // Some mailboxes return list entries without contentBytes; refetch the
      // individual attachment to get the bytes.
      const full = await graphJson<GraphFileAttachment>(`/me/messages/${messageId}/attachments/${a.id}`);
      contentBytes = full.contentBytes;
    }
    if (!contentBytes) {
      console.warn(`  Skipping attachment with no contentBytes: ${a.name}`);
      continue;
    }
    const safeName = a.name.replace(/[/\\]/g, '_');
    const target = join(inputDir, safeName);
    const buf = Buffer.from(contentBytes, 'base64');
    writeFileSync(target, buf);
    saved.push({ name: safeName, path: target, size: buf.length });
  }
  return saved;
}
