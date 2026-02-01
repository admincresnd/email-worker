import fetch from 'node-fetch';
import { getAccessToken } from './auth.js';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

export async function graphRequest(account, path, options = {}) {
  const token = await getAccessToken(account);
  const url = `${GRAPH_BASE}${path}`;

  const res = await fetch(url, {
    method: options.method || 'GET',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...options.headers,
    },
    body: options.body ? JSON.stringify(options.body) : undefined,
  });

  if (res.status === 204) return null;

  const text = await res.text();

  if (!res.ok) {
    throw new Error(`Graph API ${res.status}: ${text}`);
  }

  return text ? JSON.parse(text) : null;
}

export function userPath(account) {
  return `/users/${encodeURIComponent(account.outlook_user_email)}`;
}
