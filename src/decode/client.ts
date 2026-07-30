import { config } from '../config.js';

type GqlError = { message?: string; code?: string; __typename?: string };

export async function gql<T>(query: string, variables: Record<string, unknown> = {}): Promise<T> {
  const res = await fetch(config.decode.endpoint, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'X-Boardera-Key': config.decode.apiKey,
    },
    body: JSON.stringify({ query, variables }),
  });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`Decode HTTP ${res.status}: ${text || res.statusText}`);
  }
  const body = await res.json() as { data?: T; errors?: GqlError[] };
  if (body.errors?.length) {
    throw new Error(`Decode GraphQL errors: ${JSON.stringify(body.errors)}`);
  }
  if (!body.data) throw new Error('Decode response missing data');
  return body.data;
}
