import { getAccessToken } from './auth.js';

const GRAPH = 'https://graph.microsoft.com/v1.0';

export async function graphFetch(path: string, init: RequestInit = {}): Promise<Response> {
  const token = await getAccessToken();
  const headers = new Headers(init.headers);
  headers.set('Authorization', `Bearer ${token}`);
  if (!headers.has('Content-Type') && init.body && typeof init.body === 'string') {
    headers.set('Content-Type', 'application/json');
  }
  const url = path.startsWith('http') ? path : `${GRAPH}${path}`;
  const res = await fetch(url, { ...init, headers });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`Graph ${init.method ?? 'GET'} ${path} → ${res.status} ${res.statusText}: ${text}`);
  }
  return res;
}

export async function graphJson<T>(path: string, init: RequestInit = {}): Promise<T> {
  const res = await graphFetch(path, init);
  return res.json() as Promise<T>;
}
