import 'dotenv/config';
import { fileURLToPath } from 'url';
import { dirname, resolve } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
export const ROOT_DIR = resolve(__dirname, '..');
export const TMP_DIR = resolve(ROOT_DIR, 'tmp');
export const TOKEN_CACHE_PATH = resolve(ROOT_DIR, 'token-cache.json');
export const FOLDER_CACHE_PATH = resolve(ROOT_DIR, 'folder-cache.json');

function required(name: string): string {
  const value = process.env[name];
  if (!value) {
    throw new Error(`Missing required env var: ${name}. Copy .env.example to .env and fill it in.`);
  }
  return value;
}

export const config = {
  azure: {
    tenantId: required('AZURE_TENANT_ID'),
    clientId: required('AZURE_CLIENT_ID'),
    scopes: ['Mail.Read', 'Mail.ReadWrite', 'Mail.Send', 'offline_access'],
  },
  outlook: {
    userEmail: required('OUTLOOK_USER_EMAIL'),
    watchedFolder: required('OUTLOOK_WATCHED_FOLDER'),
    // Where result/error replies are sent — normally the quoting team's
    // group mailbox. The reports include internal pricing and margin
    // data, so they are typically not for the RFQ sender. Set to the
    // literal value "sender" to reply to the original sender instead.
    replyTo: required('OUTLOOK_REPLY_TO'),
    // Optional: blind-copy every outgoing reply to this address.
    replyBcc: process.env.OUTLOOK_REPLY_BCC || null,
  },
  decode: {
    endpoint: process.env.DECODE_API_ENDPOINT || 'https://api.boardera.io/api/v1',
    apiKey: required('DECODE_API_KEY'),
  },
  dfm: {
    // Max number of per-violation image tiles to render into the DFM
    // report (highest-severity first). 0 disables tile rendering — the
    // report still renders with summary + tables, just no images.
    imageLimit: parseInt(process.env.DFM_IMAGE_LIMIT || '20', 10),
  },
  defaultQuantity: parseInt(process.env.DEFAULT_QUANTITY || '10', 10),
  pollIntervalMs: parseInt(process.env.POLL_INTERVAL_MS || '30000', 10),
};
