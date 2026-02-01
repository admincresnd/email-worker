import { graphRequest, userPath } from './client.js';
import { forwardToN8n } from '../n8n.js';
import { patchOutlookAccount } from '../supabase.js';
import {
  STARTUP_DELAY_MS,
  RECONNECT_BASE_MS,
  RECONNECT_MAX_MS,
  POLL_INTERVAL_MS,
  sleep,
  parseEmailsSince,
} from '../config.js';

async function fetchAttachments(account, messageId) {
  const data = await graphRequest(
    account,
    `${userPath(account)}/messages/${messageId}/attachments`
  );

  if (!data?.value?.length) return [];

  return data.value
    .filter((a) => a['@odata.type'] === '#microsoft.graph.fileAttachment')
    .map((a) => ({
      filename: a.name,
      contentType: a.contentType,
      size: a.size,
      content: a.contentBytes,
    }));
}

const processedIds = new Map();
const PROCESSED_TTL_MS = 24 * 60 * 60 * 1000;
const MAX_PROCESSED_IDS = 10000;

function markProcessed(accountEmail, msgId) {
  processedIds.set(`${accountEmail}:${msgId}`, Date.now());
  if (processedIds.size > MAX_PROCESSED_IDS) {
    const now = Date.now();
    for (const [key, ts] of processedIds) {
      if (now - ts > PROCESSED_TTL_MS) processedIds.delete(key);
    }
  }
}

function isProcessed(accountEmail, msgId) {
  const key = `${accountEmail}:${msgId}`;
  const ts = processedIds.get(key);
  if (!ts) return false;
  if (Date.now() - ts > PROCESSED_TTL_MS) {
    processedIds.delete(key);
    return false;
  }
  return true;
}

async function processMessages(account, messages) {
  const emailsSince = parseEmailsSince(account.emails_since);

  for (const msg of messages) {
    if (msg.isDraft) continue;

    if (isProcessed(account.outlook_user_email, msg.id)) {
      console.log(`[outlook] ${account.outlook_user_email} skipping already-processed id=${msg.id}`);
      continue;
    }

    const receivedDate = msg.receivedDateTime ? new Date(msg.receivedDateTime) : null;

    if (emailsSince && receivedDate && receivedDate < emailsSince) {
      console.log(`[outlook] ${account.outlook_user_email} skipping id=${msg.id} (before emails_since)`);
      continue;
    }

    const attachments = await fetchAttachments(account, msg.id);

    const payload = {
      email_account_id: account.id,
      venue_id: account.venue_id,
      outlook_id: msg.id,
      conversation_id: msg.conversationId,
      from: msg.from?.emailAddress
        ? `${msg.from.emailAddress.name || ''} <${msg.from.emailAddress.address}>`
        : '',
      to: (msg.toRecipients || [])
        .map((r) => `${r.emailAddress?.name || ''} <${r.emailAddress?.address}>`)
        .join(', '),
      subject: msg.subject || '',
      date: msg.receivedDateTime,
      textPlain: msg.body?.contentType === 'text' ? msg.body.content : '',
      textHtml: msg.body?.contentType === 'html' ? msg.body.content : '',
      metadata: {
        internetMessageId: msg.internetMessageId,
        conversationId: msg.conversationId,
        importance: msg.importance,
        categories: msg.categories,
        isRead: msg.isRead,
        hasAttachments: msg.hasAttachments,
      },
      attachments,
    };

    console.log(
      `[outlook] ${account.outlook_user_email} forwarding id=${msg.id} subject="${msg.subject}" to n8n...`
    );

    const ok = await forwardToN8n(payload);

    if (!ok) {
      console.error(
        `[outlook] forward failed; stopping at id=${msg.id} (${account.outlook_user_email})`
      );
      return false;
    }

    markProcessed(account.outlook_user_email, msg.id);
    console.log(`[outlook] ${account.outlook_user_email} forwarded id=${msg.id}`);
  }

  return true;
}

async function pollDelta(account) {
  const base = `${userPath(account)}/mailFolders/inbox/messages/delta`;
  const selectFields = '$select=id,subject,from,toRecipients,body,receivedDateTime,conversationId,internetMessageId,isDraft,isRead,importance,categories,hasAttachments';

  let url;
  if (account.outlook_delta_link) {
    url = account.outlook_delta_link;
  } else {
    url = `${base}?${selectFields}`;
  }

  const allMessages = [];
  let nextLink = url;
  let deltaLink = null;

  while (nextLink) {
    const isFullUrl = nextLink.startsWith('https://');
    let data;

    if (isFullUrl) {
      const fetch_ = (await import('node-fetch')).default;
      const { getAccessToken } = await import('./auth.js');
      const token = await getAccessToken(account);
      const res = await fetch_(nextLink, {
        headers: { Authorization: `Bearer ${token}` },
      });
      const text = await res.text();
      if (!res.ok) throw new Error(`Graph API ${res.status}: ${text}`);
      data = JSON.parse(text);
    } else {
      data = await graphRequest(account, nextLink);
    }

    if (data?.value?.length) {
      allMessages.push(...data.value);
    }

    nextLink = data?.['@odata.nextLink'] || null;
    if (data?.['@odata.deltaLink']) {
      deltaLink = data['@odata.deltaLink'];
    }
  }

  if (deltaLink) {
    account.outlook_delta_link = deltaLink;
    try {
      await patchOutlookAccount(account.id, { outlook_delta_link: deltaLink });
    } catch (e) {
      console.error('[outlook] Failed to save delta link:', e.message);
    }
  }

  return allMessages.filter((m) => !m['@removed']);
}

export async function listenToOutlookInbox(account) {
  if (!account.outlook_tenant_id || !account.outlook_user_email) {
    console.warn(`[outlook] Skipping ${account.id} - missing tenant_id or user_email`);
    return;
  }

  if (!process.env.OUTLOOK_CLIENT_ID || !process.env.OUTLOOK_CLIENT_SECRET) {
    console.warn(`[outlook] Skipping ${account.id} - missing OUTLOOK_CLIENT_ID or OUTLOOK_CLIENT_SECRET env vars`);
    return;
  }

  console.log(`[outlook] Starting listener for ${account.outlook_user_email}`);

  let attempt = 0;

  while (true) {
    const reconnectDelay = Math.min(
      RECONNECT_BASE_MS * Math.pow(2, attempt),
      RECONNECT_MAX_MS
    );

    try {
      console.log(`[outlook] ${account.outlook_user_email} polling for new messages...`);

      const messages = await pollDelta(account);

      if (messages.length > 0) {
        console.log(`[outlook] ${account.outlook_user_email} found ${messages.length} new message(s)`);
        const ok = await processMessages(account, messages);
        if (!ok) {
          console.error(`[outlook] ${account.outlook_user_email} processing failed, will retry`);
        }
      } else {
        console.log(`[outlook] ${account.outlook_user_email} no new messages`);
      }

      attempt = 0;
      await sleep(POLL_INTERVAL_MS);
    } catch (err) {
      attempt++;
      console.error(
        `[outlook] ${account.outlook_user_email} error (attempt ${attempt}):`,
        err.message
      );
      await sleep(reconnectDelay);
    }
  }
}
