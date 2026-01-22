import express from 'express';
import { ImapFlow } from 'imapflow';
import fetch from 'node-fetch';
import dotenv from 'dotenv';
import { simpleParser } from 'mailparser';

dotenv.config();

const app = express();
app.use(express.json());

/* =========================================================
   Config
========================================================= */

const STARTUP_DELAY_MS = 1500;
const RECONNECT_BASE_MS = 5000;
const RECONNECT_MAX_MS = 60000;
const POLL_INTERVAL_MS = 30_000; // Fallback poll every 30 seconds

// Safety caps to avoid huge mailbox scans on cold-start alignment
const ALIGN_SCAN_BATCH_SIZE = 500; // fetch this many UIDs at a time while aligning
const RESOLVE_UID_SCAN_BATCH_SIZE = 500; // scan 500 at a time for resolve-uid

let listenersStarted = false;

/* =========================================================
   Supabase helpers
========================================================= */

const SUPABASE_HEADERS = {
  apikey: process.env.SUPABASE_SERVICE_ROLE_KEY,
  Authorization: `Bearer ${process.env.SUPABASE_SERVICE_ROLE_KEY}`,
  'Content-Type': 'application/json',
};

function mustEnv(name) {
  const v = process.env[name];
  if (!v) throw new Error(`Missing required env: ${name}`);
  return v;
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

function normalizeFolderPath(p) {
  // Some servers use "/" while others use ".".
  // Keep as provided by caller; no normalization here.
  return p;
}

function parseEmailsSince(value) {
  if (!value) return null;
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return null;
  return d;
}

async function getActiveEmailAccounts() {
  const res = await fetch(
    `${mustEnv('SUPABASE_URL')}/rest/v1/email_accounts?active=eq.true`,
    { headers: SUPABASE_HEADERS }
  );
  if (!res.ok) {
    const txt = await res.text().catch(() => '');
    throw new Error(`Failed to fetch email accounts (${res.status}): ${txt}`);
  }
  return res.json();
}

async function getEmailAccountByVenueId(venueId) {
  const res = await fetch(
    `${mustEnv('SUPABASE_URL')}/rest/v1/email_accounts?venue_id=eq.${venueId}&active=eq.true&limit=1`,
    { headers: SUPABASE_HEADERS }
  );
  if (!res.ok) {
    const txt = await res.text().catch(() => '');
    throw new Error(`Failed to fetch email account (${res.status}): ${txt}`);
  }
  const rows = await res.json();
  if (!rows.length) throw new Error('No active email account found for venue');
  return rows[0];
}

async function patchEmailAccount(accountId, patch) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 10000); // 10s timeout
  
  try {
    const res = await fetch(
      `${mustEnv('SUPABASE_URL')}/rest/v1/email_accounts?id=eq.${accountId}`,
      {
        method: 'PATCH',
        headers: SUPABASE_HEADERS,
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

async function updateLastUid(accountId, lastUid) {
  await patchEmailAccount(accountId, { last_uid: lastUid });
}

/* =========================================================
   Forward to n8n (NON-FATAL, but returns success boolean)
========================================================= */

async function forwardToN8n(payload) {
  const url = process.env.N8N_WEBHOOK_URL;
  if (!url) {
    console.error('[n8n] Missing N8N_WEBHOOK_URL (skipping send)');
    return false;
  }

  console.log(`[n8n] Sending to webhook: ${url}`);

  try {
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });

    if (!res.ok) {
      const body = await res.text().catch(() => '');
      console.error(`[n8n] webhook failed: ${res.status} - ${body}`);
      return false;
    }
    console.log(`[n8n] webhook success: ${res.status}`);
    return true;
  } catch (err) {
    console.error('[n8n] unreachable:', err.message);
    return false;
  }
}

/* =========================================================
   Alignment: derive a safe last_uid from emails_since
   - Purpose: never process emails before emails_since
   - If last_uid is missing/out-of-sync, realign to first UID
     whose INTERNALDATE >= emails_since (server-assigned time).
========================================================= */

async function ensureAlignedLastUid(client, account) {
  const emailsSince = parseEmailsSince(account.emails_since);

  // If emails_since not set, do nothing special.
  if (!emailsSince) return;

  // Opened mailbox required for fetch/search. Caller should hold lock.
  // We align using INTERNALDATE (server time), not message Date header.
  // Strategy:
  // 1) Find the first UID whose INTERNALDATE >= emailsSince by scanning
  //    forward in UID order in bounded batches.
  // 2) If found uidSince:
  //      - If account.last_uid is null/0 or < (uidSince - 1), bump it to uidSince - 1
  //        so processing starts at uidSince.
  //      - If account.last_uid is already >= uidSince - 1, keep it (no backtracking).
  //
  // This guarantees:
  // - Nothing before emailsSince is fetched/processed
  // - We never skip emails after emailsSince due to a stale last_uid (we might re-run
  //   a small tail only if last_uid was ahead and messages are unseen; but commit rules
  //   prevent loss).

  let uidSince = null;

  // Quick shortcut: if mailbox empty, nothing to do
  const mb = client.mailbox;
  if (!mb || !mb.exists || mb.exists === 0) return;

  const maxUid = mb.uidNext ? mb.uidNext - 1 : null;
  if (!maxUid || maxUid <= 0) return;

  // Choose a reasonable starting point.
  // If last_uid exists, start scanning from max(1, last_uid) backwards? No:
  // we need FIRST UID >= emailsSince, so we scan from 1 upwards.
  // To avoid scanning huge mailboxes from 1, we can start from a heuristic:
  // start near last_uid if it exists, otherwise start from maxUid - some window.
  // But safest deterministic approach is to scan forward; we'll do a bounded skip:
  // - Start at max(1, (account.last_uid || 1))
  //   If last_uid is far beyond emailsSince, we'll still be safe.
  // - But if last_uid is stale (too low), scanning from it reduces work vs from 1.
  let start = account.last_uid && account.last_uid > 0 ? account.last_uid : 1;
  if (start > maxUid) start = 1;

  // Scan forward from `start`. If we fail to find (because start is after the date),
  // we then scan from 1 (rare, but handles the case where last_uid jumped too far).
  async function scanFrom(startUid) {
    let cursor = startUid;

    while (cursor <= maxUid) {
      const end = Math.min(cursor + ALIGN_SCAN_BATCH_SIZE - 1, maxUid);
      const range = `${cursor}:${end}`;

      // INTERNALDATE comes from 'internalDate'
      const candidates = [];
      for await (const msg of client.fetch(range, {
        uid: true,
        internalDate: true,
      })) {
        if (!msg?.uid) continue;
        if (msg.internalDate && msg.internalDate >= emailsSince) {
          candidates.push(msg.uid);
        }
      }

      if (candidates.length) {
        return Math.min(...candidates);
      }

      cursor = end + 1;
    }

    return null;
  }

  uidSince = await scanFrom(start);

  // If scanning from last_uid-ish didn't find anything (e.g., last_uid after date),
  // scan from 1 to find earliest UID >= emailsSince.
  if (!uidSince && start !== 1) {
    uidSince = await scanFrom(1);
  }

  // If still not found, then there are no messages with internalDate >= emailsSince.
  // In that case, set last_uid to maxUid to avoid any processing until new mail arrives.
  if (!uidSince) {
    const desiredLast = maxUid;
    if ((account.last_uid || 0) !== desiredLast) {
      account.last_uid = desiredLast;
      try {
        await updateLastUid(account.id, desiredLast);
      } catch (e) {
        console.error('[align] Failed to update last_uid:', e.message);
      }
    }
    return;
  }

  const desiredLastUid = Math.max(0, uidSince - 1);

  // Only move forward (never regress), to avoid replaying history.
  const currentLast = account.last_uid || 0;
  if (currentLast < desiredLastUid) {
    account.last_uid = desiredLastUid;
    try {
      await updateLastUid(account.id, desiredLastUid);
      console.log(
        `[align] ${account.imap_username} set last_uid=${desiredLastUid} (emails_since=${emailsSince.toISOString()})`
      );
    } catch (e) {
      console.error('[align] Failed to update last_uid:', e.message);
    }
  }
}

/* =========================================================
   Process UNSEEN + UID > last_uid
   + emails_since gate using INTERNALDATE
   + commit only after successful forward
========================================================= */

async function processNewUnseen(client, account) {
  // Ensure emails_since gate is enforced even if last_uid is stale.
  // Caller holds INBOX lock and mailbox is open via getMailboxLock('INBOX')
  await ensureAlignedLastUid(client, account);

  const emailsSince = parseEmailsSince(account.emails_since);
  const lastUid = account.last_uid || 0;

  console.log(`[process] ${account.imap_username} searching for emails with uid>${lastUid}...`);

  // First, let's see what's in the mailbox for debugging
  const mailbox = client.mailbox;
  console.log(`[process] ${account.imap_username} mailbox state: exists=${mailbox?.exists}, uidNext=${mailbox?.uidNext}, uidValidity=${mailbox?.uidValidity}`);

  // Search for unseen emails beyond last_uid
  const uids = await client.search({ 
    seen: false, 
    uid: `${lastUid + 1}:*` 
  }, { uid: true });

  console.log(`[process] ${account.imap_username} raw search returned ${uids.length} UID(s): ${uids.slice(0, 10).join(', ')}${uids.length > 10 ? '...' : ''}`);

  // Also search for unseen to compare (debug only)
  const unseenUids = await client.search({ seen: false }, { uid: true });
  console.log(`[process] ${account.imap_username} total unseen in mailbox: ${unseenUids.length} UID(s): ${unseenUids.slice(0, 10).join(', ')}${unseenUids.length > 10 ? '...' : ''}`);

  // Filter out UIDs <= lastUid (IMAP quirk: "1000:*" can return 1000 even if searching for >1000)
  const validUids = uids.filter(uid => uid > lastUid);

  console.log(`[process] ${account.imap_username} after filtering uid>${lastUid}: ${validUids.length} valid UID(s)`);

  if (!validUids.length) return;

  // Sort for deterministic processing (ascending UID)
  validUids.sort((a, b) => a - b);
  console.log(`[process] ${account.imap_username} UIDs to process: ${validUids.join(', ')}`);

  for await (const msg of client.fetch(validUids, {
    uid: true,
    envelope: true,
    source: true,
    internalDate: true,
  }, { uid: true })) {
    // Double-check: skip if UID is not greater than lastUid (safety)
    if (msg.uid <= account.last_uid) {
      console.log(`[process] ${account.imap_username} skipping uid=${msg.uid} (already processed, last_uid=${account.last_uid})`);
      continue;
    }
    
    console.log(`[process] ${account.imap_username} fetched uid=${msg.uid}, internalDate=${msg.internalDate}`);

    // Hard gate: never process emails earlier than emails_since (by INTERNALDATE)
    if (emailsSince && msg.internalDate && msg.internalDate < emailsSince) {
      console.log(`[process] ${account.imap_username} skipping uid=${msg.uid} (before emails_since=${emailsSince.toISOString()})`);
      // We should not advance last_uid based on old emails; just skip.
      continue;
    }

    const parsed = await simpleParser(msg.source);

    console.log(`[process] ${account.imap_username} processing uid=${msg.uid} from="${parsed.from?.text}" subject="${parsed.subject}"`);

    const payload = {
      email_account_id: account.id,
      venue_id: account.venue_id,
      uid: msg.uid,
      from: parsed.from?.text || '',
      to: parsed.to?.text || '',
      subject: parsed.subject || '',
      date: parsed.date,
      internalDate: msg.internalDate || null,
      textPlain: parsed.text || '',
      textHtml: parsed.html || '',
      metadata: parsed.headers ? Object.fromEntries(parsed.headers) : {},
      raw: msg.source.toString(),
    };

    // Send to n8n. If it fails, DO NOT commit.
    console.log(`[process] ${account.imap_username} forwarding uid=${msg.uid} to n8n...`);
    const ok = await forwardToN8n(payload);

    if (!ok) {
      console.error(
        `[ingest] forward failed; not committing uid=${msg.uid} (${account.imap_username})`
      );
      // Stop processing further messages to preserve ordering and avoid skipping.
      break;
    }

    console.log(`[process] ${account.imap_username} ✓ forwarded uid=${msg.uid} successfully`);

    // Commit progress
    console.log(`[process] ${account.imap_username} updating last_uid in Supabase...`);
    account.last_uid = msg.uid;
    await updateLastUid(account.id, msg.uid);
    console.log(`[process] ${account.imap_username} ✓ committed uid=${msg.uid}`);
  }
  
  console.log(`[process] ${account.imap_username} finished processing all messages`);
}

/* =========================================================
   IMAP listener (IDLE + fallback poll + reconnect)
   
   Strategy:
   - Startup: fetch unseen mail immediately
   - IDLE: reacts instantly when server emits events
   - Fallback poll every 30s: guarantees new mail is picked up
     even if IDLE misses events (some servers are unreliable)
   - Safe, idempotent, production-grade
========================================================= */

async function listenToInbox(account) {
  if (!account.imap_username || !account.imap_secret) {
    console.warn(`[imap] Skipping ${account.id} – missing credentials`);
    return;
  }

  console.log(`[imap] Starting listener for ${account.imap_username}`);

  let attempt = 0;

  while (true) {
    const reconnectDelay = Math.min(
      RECONNECT_BASE_MS * Math.pow(2, attempt),
      RECONNECT_MAX_MS
    );

    let client;
    let pollTimer = null;

    try {
      console.log(`[imap] ${account.imap_username} connecting to ${account.imap_host}:${account.imap_port}...`);

      client = new ImapFlow({
        host: account.imap_host,
        port: account.imap_port,
        secure: account.imap_secure,
        auth: {
          user: account.imap_username,
          pass: account.imap_secret,
        },
        logger: false,
        keepalive: {
          interval: 10_000,
          idleInterval: 300_000,
          forceNoop: true,
        },
      });

      await client.connect();
      console.log(`[imap] ${account.imap_username} connected successfully`);
      attempt = 0;

      const lock = await client.getMailboxLock('INBOX');
      console.log(`[imap] ${account.imap_username} INBOX lock acquired, mailbox exists=${client.mailbox?.exists}, uidNext=${client.mailbox?.uidNext}`);

      try {
        // Startup: process any existing unseen mail
        console.log(`[imap] ${account.imap_username} processing unseen on startup...`);
        await processNewUnseen(client, account);

        // Main loop: IDLE with fallback polling
        console.log(`[imap] ${account.imap_username} entering IDLE loop (poll every ${POLL_INTERVAL_MS / 1000}s)`);
        
        while (true) {
          // Create a promise that resolves after POLL_INTERVAL_MS
          let pollResolve;
          const pollPromise = new Promise((resolve) => {
            pollResolve = resolve;
            pollTimer = setTimeout(() => resolve('poll'), POLL_INTERVAL_MS);
          });

          // Race between IDLE (server push) and poll timeout
          // client.idle() resolves when:
          // - Server sends EXISTS/EXPUNGE/etc.
          // - Connection issue occurs
          // - IDLE times out internally (per idleInterval in keepalive)
          const idlePromise = client.idle().then(() => 'idle');

          const reason = await Promise.race([idlePromise, pollPromise]);

          // Clear the poll timer if IDLE won
          if (pollTimer) {
            clearTimeout(pollTimer);
            pollTimer = null;
          }

          console.log(`[imap] ${account.imap_username} woke up (reason: ${reason}), checking for new mail...`);

          // Process new mail regardless of whether IDLE or poll triggered
          // This is idempotent - if no new mail, it's a no-op
          await processNewUnseen(client, account);
        }
      } finally {
        // Clean up poll timer on any exit
        if (pollTimer) {
          clearTimeout(pollTimer);
          pollTimer = null;
        }
        lock.release();
        console.log(`[imap] ${account.imap_username} INBOX lock released`);
      }
    } catch (err) {
      attempt++;
      console.error(
        `[imap] ${account.imap_username} error (reconnect attempt ${attempt}):`,
        err.message
      );

      // Clean up poll timer on error
      if (pollTimer) {
        clearTimeout(pollTimer);
        pollTimer = null;
      }

      console.log(`[imap] ${account.imap_username} waiting ${reconnectDelay / 1000}s before reconnect...`);
      await sleep(reconnectDelay);
    } finally {
      if (client) {
        try {
          await client.logout();
        } catch {}
      }
    }
  }
}

/* =========================================================
   Start listeners
========================================================= */

async function startAllListeners() {
  if (listenersStarted) return;
  listenersStarted = true;

  // Validate required env early (production sanity)
  mustEnv('SUPABASE_URL');
  mustEnv('SUPABASE_SERVICE_ROLE_KEY');

  console.log('[startup] Fetching active email accounts...');
  const accounts = await getActiveEmailAccounts();
  console.log(`[startup] Found ${accounts.length} active email account(s)`);

  if (accounts.length === 0) {
    console.warn('[startup] No active email accounts found! Check email_accounts table.');
    return;
  }

  for (const account of accounts) {
    console.log(`[startup] Will start listener for ${account.imap_username} (venue_id=${account.venue_id}) in ${STARTUP_DELAY_MS}ms`);
    await sleep(STARTUP_DELAY_MS);
    listenToInbox(account);
  }
}

/* =========================================================
   Resolve UID endpoint (kept; made scan safer via batching)
========================================================= */

app.post('/imap/resolve-uid', async (req, res) => {
  const { venue_id, folder, message_id } = req.body;
  if (!venue_id || !folder || !message_id) {
    return res.status(400).json({ error: 'Missing fields' });
  }

  let client;

  try {
    const account = await getEmailAccountByVenueId(venue_id);

    client = new ImapFlow({
      host: account.imap_host,
      port: account.imap_port,
      secure: account.imap_secure,
      auth: {
        user: account.imap_username,
        pass: account.imap_secret,
      },
      keepalive: {
        interval: 10_000,
        idleInterval: 300_000,
        forceNoop: true,
      },
      logger: false,
    });

    await client.connect();

    const folderPath = normalizeFolderPath(folder);

    const mailboxes = await client.list();
    if (!mailboxes.find((m) => m.path === folderPath)) {
      return res.status(404).json({ error: 'Mailbox not found' });
    }

    await client.mailboxOpen(folderPath);

    const target = message_id.replace(/[<>]/g, '');

    // Scan in batches instead of "1:*" in one go (production safety)
    const mb = client.mailbox;
    const exists = mb?.exists || 0;
    if (!exists) return res.status(404).json({ error: 'Message-ID not found' });

    let seqStart = 1;

    while (seqStart <= exists) {
      const seqEnd = Math.min(seqStart + RESOLVE_UID_SCAN_BATCH_SIZE - 1, exists);
      const range = `${seqStart}:${seqEnd}`;

      for await (const msg of client.fetch(range, {
        uid: true,
        envelope: true,
      })) {
        const mid = msg.envelope?.messageId?.replace(/[<>]/g, '');
        if (mid && mid === target) {
          return res.json({ uid: msg.uid });
        }
      }

      seqStart = seqEnd + 1;
    }

    return res.status(404).json({ error: 'Message-ID not found' });
  } catch (err) {
    console.error('[resolve-uid]', err.message);
    return res.status(500).json({ error: err.message });
  } finally {
    if (client) {
      try {
        await client.logout();
      } catch {}
    }
  }
});

/* =========================================================
   Boot
========================================================= */

const PORT = process.env.WORKER_PORT || 3005;

app.listen(PORT, '0.0.0.0', () => {
  console.log(`email-worker listening on port ${PORT}`);
  startAllListeners().catch(console.error);
});
