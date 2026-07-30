import { existsSync, readFileSync, writeFileSync, chmodSync } from 'fs';
import { PublicClientApplication, type Configuration, LogLevel } from '@azure/msal-node';
import { config, TOKEN_CACHE_PATH } from '../config.js';

const msalConfig: Configuration = {
  auth: {
    clientId: config.azure.clientId,
    authority: `https://login.microsoftonline.com/${config.azure.tenantId}`,
  },
  system: {
    loggerOptions: {
      loggerCallback: () => {},
      piiLoggingEnabled: false,
      logLevel: LogLevel.Warning,
    },
  },
};

const pca = new PublicClientApplication(msalConfig);

function loadCache(): void {
  if (existsSync(TOKEN_CACHE_PATH)) {
    pca.getTokenCache().deserialize(readFileSync(TOKEN_CACHE_PATH, 'utf8'));
  }
}

async function persistCache(): Promise<void> {
  const data = await pca.getTokenCache().serialize();
  writeFileSync(TOKEN_CACHE_PATH, data, 'utf8');
  try { chmodSync(TOKEN_CACHE_PATH, 0o600); } catch { /* best-effort */ }
}

async function trySilent(): Promise<string | null> {
  const accounts = await pca.getTokenCache().getAllAccounts();
  const account = accounts.find(a => a.username?.toLowerCase() === config.outlook.userEmail.toLowerCase()) ?? accounts[0];
  if (!account) return null;
  try {
    const result = await pca.acquireTokenSilent({ account, scopes: config.azure.scopes });
    await persistCache();
    return result?.accessToken ?? null;
  } catch {
    return null;
  }
}

async function deviceCode(): Promise<string> {
  const result = await pca.acquireTokenByDeviceCode({
    scopes: config.azure.scopes,
    deviceCodeCallback: (resp) => {
      console.log('\n=== Microsoft sign-in required ===');
      console.log(resp.message);
      console.log('==================================\n');
    },
  });
  if (!result?.accessToken) throw new Error('Device code flow did not return an access token');
  await persistCache();
  return result.accessToken;
}

let loaded = false;

export async function getAccessToken(): Promise<string> {
  if (!loaded) {
    loadCache();
    loaded = true;
  }
  const silent = await trySilent();
  if (silent) return silent;
  return await deviceCode();
}

export async function forceAuth(): Promise<void> {
  if (!loaded) { loadCache(); loaded = true; }
  await deviceCode();
  console.log(`Signed in. Token cache: ${TOKEN_CACHE_PATH}`);
}
