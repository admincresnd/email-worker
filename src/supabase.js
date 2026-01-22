import fetch from 'node-fetch';
import { mustEnv } from './config.js';

function getHeaders() {
  return {
    apikey: process.env.SUPABASE_SERVICE_ROLE_KEY,
    Authorization: `Bearer ${process.env.SUPABASE_SERVICE_ROLE_KEY}`,
    'Content-Type': 'application/json',
  };
}

export async function getActiveEmailAccounts() {
  const res = await fetch(
    `${mustEnv('SUPABASE_URL')}/rest/v1/email_accounts?active=eq.true`,
    { headers: getHeaders() }
  );
  if (!res.ok) {
    const txt = await res.text().catch(() => '');
    throw new Error(`Failed to fetch email accounts (${res.status}): ${txt}`);
  }
  return res.json();
}

export async function getEmailAccountByVenueId(venueId) {
  const res = await fetch(
    `${mustEnv('SUPABASE_URL')}/rest/v1/email_accounts?venue_id=eq.${venueId}&active=eq.true&limit=1`,
    { headers: getHeaders() }
  );
  if (!res.ok) {
    const txt = await res.text().catch(() => '');
    throw new Error(`Failed to fetch email account (${res.status}): ${txt}`);
  }
  const rows = await res.json();
  if (!rows.length) throw new Error('No active email account found for venue');
  return rows[0];
}

export async function patchEmailAccount(accountId, patch) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 10000);
  
  try {
    const res = await fetch(
      `${mustEnv('SUPABASE_URL')}/rest/v1/email_accounts?id=eq.${accountId}`,
      {
        method: 'PATCH',
        headers: getHeaders(),
        body: JSON.stringify(patch),
        signal: controller.signal,
      }
    );
    if (!res.ok) {
      const txt = await res.text().catch(() => '');
      throw new Error(`Failed to update email account (${res.status}): ${txt}`);
    }
  } finally {
    clearTimeout(timeout);
  }
}

export async function updateLastUid(accountId, lastUid) {
  await patchEmailAccount(accountId, { last_uid: lastUid });
}
