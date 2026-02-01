import { ConfidentialClientApplication } from '@azure/msal-node';

const tokenCache = new Map();

function buildMsalClient(account) {
  const clientId = process.env.OUTLOOK_CLIENT_ID;
  const clientSecret = process.env.OUTLOOK_CLIENT_SECRET;

  if (!clientId || !clientSecret) {
    throw new Error('Missing OUTLOOK_CLIENT_ID or OUTLOOK_CLIENT_SECRET env vars');
  }

  return new ConfidentialClientApplication({
    auth: {
      clientId,
      clientSecret,
      authority: `https://login.microsoftonline.com/${account.outlook_tenant_id}`,
    },
  });
}

export async function getAccessToken(account) {
  const cacheKey = account.id;
  const cached = tokenCache.get(cacheKey);

  if (cached && cached.expiresAt > Date.now() + 60_000) {
    return cached.token;
  }

  const client = buildMsalClient(account);

  const result = await client.acquireTokenByClientCredential({
    scopes: ['https://graph.microsoft.com/.default'],
  });

  if (!result?.accessToken) {
    throw new Error('Failed to acquire access token from Azure AD');
  }

  tokenCache.set(cacheKey, {
    token: result.accessToken,
    expiresAt: result.expiresOn ? result.expiresOn.getTime() : Date.now() + 3600_000,
  });

  return result.accessToken;
}
