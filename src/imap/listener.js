import { simpleParser } from 'mailparser';
import { createImapClient } from './client.js';
import { forwardToN8n } from '../n8n.js';
import { updateLastUid } from '../supabase.js';
import {
  STARTUP_DELAY_MS,
  RECONNECT_BASE_MS,
  RECONNECT_MAX_MS,
  POLL_INTERVAL_MS,
  ALIGN_SCAN_BATCH_SIZE,
  sleep,
  parseEmailsSince,
} from '../config.js';

async function ensureAlignedLastUid(client, account) {
  const emailsSince = parseEmailsSince(account.emails_since);

  if (!emailsSince) return;

  let uidSince = null;

  const mb = client.mailbox;
  if (!mb || !mb.exists || mb.exists === 0) return;

  const maxUid = mb.uidNext ? mb.uidNext - 1 : null;
  if (!maxUid || maxUid <= 0) return;

  let start = account.last_uid && account.last_uid > 0 ? account.last_uid : 1;
  if (start > maxUid) start = 1;

  async function scanFrom(startUid) {
    let cursor = startUid;

    while (cursor <= maxUid) {
      const end = Math.min(cursor + ALIGN_SCAN_BATCH_SIZE - 1, maxUid);
      const range = `${cursor}:${end}`;

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

  if (!uidSince && start !== 1) {
    uidSince = await scanFrom(1);
  }

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

async function processNewUnseen(client, account) {
  await ensureAlignedLastUid(client, account);

  const emailsSince = parseEmailsSince(account.emails_since);
  const lastUid = account.last_uid || 0;

  console.log(`[process] ${account.imap_username} searching for emails with uid>${lastUid}...`);

  const mailbox = client.mailbox;
  console.log(`[process] ${account.imap_username} mailbox state: exists=${mailbox?.exists}, uidNext=${mailbox?.uidNext}, uidValidity=${mailbox?.uidValidity}`);

  const uids = await client.search({ 
    seen: false, 
    uid: `${lastUid + 1}:*` 
  }, { uid: true });

  console.log(`[process] ${account.imap_username} raw search returned ${uids.length} UID(s): ${uids.slice(0, 10).join(', ')}${uids.length > 10 ? '...' : ''}`);

  const unseenUids = await client.search({ seen: false }, { uid: true });
  console.log(`[process] ${account.imap_username} total unseen in mailbox: ${unseenUids.length} UID(s): ${unseenUids.slice(0, 10).join(', ')}${unseenUids.length > 10 ? '...' : ''}`);

  const validUids = uids.filter(uid => uid > lastUid);

  console.log(`[process] ${account.imap_username} after filtering uid>${lastUid}: ${validUids.length} valid UID(s)`);

  if (!validUids.length) return;

  validUids.sort((a, b) => a - b);
  console.log(`[process] ${account.imap_username} UIDs to process: ${validUids.join(', ')}`);

  for await (const msg of client.fetch(validUids, {
    uid: true,
    envelope: true,
    source: true,
    internalDate: true,
  }, { uid: true })) {
    if (msg.uid <= account.last_uid) {
      console.log(`[process] ${account.imap_username} skipping uid=${msg.uid} (already processed, last_uid=${account.last_uid})`);
      continue;
    }
    
    console.log(`[process] ${account.imap_username} fetched uid=${msg.uid}, internalDate=${msg.internalDate}`);

    if (emailsSince && msg.internalDate && msg.internalDate < emailsSince) {
      console.log(`[process] ${account.imap_username} skipping uid=${msg.uid} (before emails_since=${emailsSince.toISOString()})`);
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

    console.log(`[process] ${account.imap_username} forwarding uid=${msg.uid} to n8n...`);
    const ok = await forwardToN8n(payload);

    if (!ok) {
      console.error(
        `[ingest] forward failed; not committing uid=${msg.uid} (${account.imap_username})`
      );
      break;
    }

    console.log(`[process] ${account.imap_username} ✓ forwarded uid=${msg.uid} successfully`);

    console.log(`[process] ${account.imap_username} updating last_uid in Supabase...`);
    account.last_uid = msg.uid;
    await updateLastUid(account.id, msg.uid);
    console.log(`[process] ${account.imap_username} ✓ committed uid=${msg.uid}`);
  }
  
  console.log(`[process] ${account.imap_username} finished processing all messages`);
}

export async function listenToInbox(account) {
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

      client = createImapClient(account);

      await client.connect();
      console.log(`[imap] ${account.imap_username} connected successfully`);
      attempt = 0;

      const lock = await client.getMailboxLock('INBOX');
      console.log(`[imap] ${account.imap_username} INBOX lock acquired, mailbox exists=${client.mailbox?.exists}, uidNext=${client.mailbox?.uidNext}`);

      try {
        console.log(`[imap] ${account.imap_username} processing unseen on startup...`);
        await processNewUnseen(client, account);

        console.log(`[imap] ${account.imap_username} entering IDLE loop (poll every ${POLL_INTERVAL_MS / 1000}s)`);
        
        while (true) {
          let pollResolve;
          const pollPromise = new Promise((resolve) => {
            pollResolve = resolve;
            pollTimer = setTimeout(() => resolve('poll'), POLL_INTERVAL_MS);
          });

          const idlePromise = client.idle().then(() => 'idle');

          const reason = await Promise.race([idlePromise, pollPromise]);

          if (pollTimer) {
            clearTimeout(pollTimer);
            pollTimer = null;
          }

          console.log(`[imap] ${account.imap_username} woke up (reason: ${reason}), checking for new mail...`);

          await processNewUnseen(client, account);
        }
      } finally {
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
