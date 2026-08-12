/**
 * Sanity check: exercises src/decode/client.ts against the configured endpoint
 * and reports whether (a) the URL is reachable and (b) the key authenticates.
 */
import { gql } from '../src/decode/client.js';
import { config } from '../src/config.js';

console.log(`Endpoint: ${config.decode.endpoint}`);
console.log(`Key:      ${config.decode.apiKey ? '(set)' : '(missing)'}`);

try {
  const result = await gql<any>('{ __typename }');
  console.log('\nSUCCESS — server responded:');
  console.log(JSON.stringify(result, null, 2));
} catch (err) {
  console.error('\nFAILED:', (err as Error).message);
  process.exitCode = 1;
}
