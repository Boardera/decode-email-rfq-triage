import { existsSync, readFileSync, writeFileSync } from 'fs';
import { config, FOLDER_CACHE_PATH } from '../config.js';
import { graphJson, graphFetch } from './graph.js';

export type GraphMessage = {
  id: string;
  subject: string;
  bodyPreview?: string;
  body?: { contentType: 'html' | 'text'; content: string };
  from?: { emailAddress: { name?: string; address: string } };
  toRecipients?: { emailAddress: { address: string } }[];
  hasAttachments?: boolean;
  receivedDateTime?: string;
  isRead?: boolean;
};

type FolderCache = { id: string; displayName: string };

function loadFolderCache(): FolderCache | null {
  if (!existsSync(FOLDER_CACHE_PATH)) return null;
  try { return JSON.parse(readFileSync(FOLDER_CACHE_PATH, 'utf8')); }
  catch { return null; }
}

function saveFolderCache(cache: FolderCache): void {
  writeFileSync(FOLDER_CACHE_PATH, JSON.stringify(cache, null, 2), 'utf8');
}

async function findFolderIdByName(displayName: string): Promise<string> {
  // Outlook folders can be nested. Walk the tree shallowly: search at root, then under "Inbox" children.
  const escName = displayName.replace(/'/g, "''");
  const filter = `$filter=displayName eq '${escName}'&$select=id,displayName&$top=5`;

  const rootRes = await graphJson<{ value: FolderCache[] }>(`/me/mailFolders?${filter}`);
  if (rootRes.value.length) return rootRes.value[0].id;

  const inboxRes = await graphJson<{ value: FolderCache[] }>(`/me/mailFolders/inbox/childFolders?${filter}`);
  if (inboxRes.value.length) return inboxRes.value[0].id;

  throw new Error(
    `Outlook folder "${displayName}" not found at the root or under Inbox. ` +
    `Either create the folder and the Outlook rule that moves mail into it, or rename OUTLOOK_WATCHED_FOLDER in .env.`,
  );
}

export async function resolveWatchedFolderId(): Promise<string> {
  const cached = loadFolderCache();
  if (cached?.displayName === config.outlook.watchedFolder && cached.id) return cached.id;
  const id = await findFolderIdByName(config.outlook.watchedFolder);
  saveFolderCache({ id, displayName: config.outlook.watchedFolder });
  return id;
}

export async function listUnreadMessages(limit = 10): Promise<GraphMessage[]> {
  const folderId = await resolveWatchedFolderId();
  const select = '$select=id,subject,bodyPreview,body,from,toRecipients,hasAttachments,receivedDateTime,isRead';
  const filter = '$filter=isRead eq false';
  const order = '$orderby=receivedDateTime asc';
  const top = `$top=${limit}`;
  const path = `/me/mailFolders/${folderId}/messages?${filter}&${order}&${top}&${select}`;
  const res = await graphJson<{ value: GraphMessage[] }>(path);
  return res.value;
}

export async function markMessageRead(messageId: string): Promise<void> {
  await graphFetch(`/me/messages/${messageId}`, {
    method: 'PATCH',
    body: JSON.stringify({ isRead: true }),
  });
}
