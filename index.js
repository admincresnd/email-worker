import express from 'express';
import dotenv from 'dotenv';
import { mustEnv, sleep, STARTUP_DELAY_MS } from './src/config.js';
import { getActiveEmailAccounts, getEmailAccountByVenueId, getActiveOutlookAccounts, getOutlookAccountByVenueId } from './src/supabase.js';
import { listenToInbox } from './src/imap/listener.js';
import { resolveUid, listFolders } from './src/imap/resolve.js';
import { moveEmail } from './src/imap/move.js';
import { sendEmail } from './src/smtp/send.js';
import { createDraft } from './src/imap/draft.js';
import { listenToOutlookInbox } from './src/outlook/listener.js';
import { sendOutlookEmail, replyOutlookEmail } from './src/outlook/send.js';
import { moveOutlookEmail } from './src/outlook/move.js';

dotenv.config();

const app = express();
app.use(express.json({ limit: '25mb' }));

let listenersStarted = false;

async function startAllListeners() {
  if (listenersStarted) return;
  listenersStarted = true;

  mustEnv('SUPABASE_URL');
  mustEnv('SUPABASE_SERVICE_ROLE_KEY');

  console.log('[startup] Fetching active email accounts...');
  const emailAccounts = await getActiveEmailAccounts();
  console.log(`[startup] Found ${emailAccounts.length} active IMAP email account(s)`);

  for (const account of emailAccounts) {
    await sleep(STARTUP_DELAY_MS);
    console.log(`[startup] Starting IMAP listener for ${account.imap_username} (venue_id=${account.venue_id})`);
    listenToInbox(account);
  }

  console.log('[startup] Fetching active outlook accounts...');
  let outlookAccounts = [];
  try {
    outlookAccounts = await getActiveOutlookAccounts();
  } catch (err) {
    console.warn('[startup] Could not fetch outlook accounts:', err.message);
  }
  console.log(`[startup] Found ${outlookAccounts.length} active Outlook account(s)`);

  for (const account of outlookAccounts) {
    await sleep(STARTUP_DELAY_MS);
    console.log(`[startup] Starting Outlook listener for ${account.outlook_user_email} (venue_id=${account.venue_id})`);
    listenToOutlookInbox(account);
  }
}

app.get('/imap/folders', async (req, res) => {
  const { venue_id } = req.query;
  if (!venue_id) {
    return res.status(400).json({ error: 'Missing query param: venue_id' });
  }
  try {
    const account = await getEmailAccountByVenueId(venue_id);
    const folders = await listFolders(account);
    return res.json({ folders });
  } catch (err) {
    console.error('[folders]', err.message);
    return res.status(500).json({ error: err.message });
  }
});

app.post('/imap/resolve-uid', async (req, res) => {
  console.log('[resolve-uid] Request received:', JSON.stringify(req.body));
  const { venue_id, folder, message_id } = req.body;
  if (!venue_id || !folder || !message_id) {
    console.log('[resolve-uid] Missing required fields');
    return res.status(400).json({ error: 'Missing fields: venue_id, folder, message_id required' });
  }

  try {
    console.log(`[resolve-uid] Looking up account for venue_id=${venue_id}`);
    const account = await getEmailAccountByVenueId(venue_id);
    console.log(`[resolve-uid] Searching folder="${folder}" for message_id="${message_id}"`);
    const result = await resolveUid(account, { folder, message_id });
    console.log('[resolve-uid] Found:', result);
    return res.json(result);
  } catch (err) {
    console.error('[resolve-uid] Error:', err.message);
    if (err.message === 'Mailbox not found' || err.message === 'Message-ID not found') {
      return res.status(404).json({ error: err.message });
    }
    return res.status(500).json({ error: err.message });
  }
});

app.post('/outlook/move', async (req, res) => {
  const { venue_id, uid, outlook_id, folder, mark_as_seen, flagged } = req.body;
  const messageId = outlook_id || uid;
  if (!venue_id || !messageId || !folder) {
    return res.status(400).json({ error: 'Missing fields: venue_id, uid/outlook_id, folder required' });
  }

  try {
    const account = await getOutlookAccountByVenueId(venue_id);
    const result = await moveOutlookEmail(account, { outlook_id: messageId, folder, mark_as_seen, flagged });
    return res.json(result);
  } catch (err) {
    console.error('[outlook/move]', err.message);
    return res.status(500).json({ error: err.message });
  }
});

app.post('/imap/move', async (req, res) => {
  const { venue_id, uid, folder, source_folder, mark_as_seen, flagged } = req.body;
  if (!venue_id || !uid || !folder) {
    return res.status(400).json({ error: 'Missing fields: venue_id, uid, folder required' });
  }

  try {
    const account = await getEmailAccountByVenueId(venue_id);
    const result = await moveEmail(account, { uid, folder, source_folder, mark_as_seen, flagged });
    return res.json(result);
  } catch (err) {
    console.error('[move]', err.message);
    return res.status(500).json({ error: err.message });
  }
});

app.post('/outlook/send', async (req, res) => {
  const { venue_id, from, to, subject, html, attachments } = req.body;
  if (!venue_id || !from || !to || !subject) {
    return res.status(400).json({ error: 'Missing fields: venue_id, from, to, subject required' });
  }

  try {
    const account = await getOutlookAccountByVenueId(venue_id);
    const result = await sendOutlookEmail(account, { from, to, subject, html, attachments });
    return res.json(result);
  } catch (err) {
    console.error('[outlook/send]', err.message);
    return res.status(500).json({ error: err.message });
  }
});

app.post('/outlook/reply', async (req, res) => {
  const { venue_id, outlook_id, from, to, subject, html, attachments } = req.body;
  if (!venue_id || !outlook_id) {
    return res.status(400).json({ error: 'Missing fields: venue_id, outlook_id required' });
  }

  try {
    const account = await getOutlookAccountByVenueId(venue_id);
    const result = await replyOutlookEmail(account, { outlook_id, from, to, subject, html, attachments });
    return res.json(result);
  } catch (err) {
    console.error('[outlook/reply]', err.message);
    return res.status(500).json({ error: err.message });
  }
});

app.post('/smtp/send', async (req, res) => {
  const { venue_id, from, to, subject, html, inReplyTo, references, folder_path, attachments } = req.body;
  if (!venue_id || !from || !to || !subject) {
    return res.status(400).json({ error: 'Missing fields: venue_id, from, to, subject required' });
  }

  try {
    const account = await getEmailAccountByVenueId(venue_id);
    const result = await sendEmail(account, { from, to, subject, html, inReplyTo, references, folder_path, attachments });
    return res.json(result);
  } catch (err) {
    console.error('[smtp/send]', err.message);
    return res.status(500).json({ error: err.message });
  }
});

app.post('/imap/draft', async (req, res) => {
  console.log('[draft] Request received:', JSON.stringify(req.body));
  const { venue_id, from, to, subject, html, folder_path, attachments } = req.body;
  if (!venue_id || !from || !to || !subject || !folder_path) {
    console.log('[draft] Missing required fields');
    return res.status(400).json({ error: 'Missing fields: venue_id, from, to, subject, folder_path required' });
  }

  try {
    console.log(`[draft] Looking up account for venue_id=${venue_id}`);
    const account = await getEmailAccountByVenueId(venue_id);
    console.log(`[draft] Creating draft for account: ${account.imap_username}`);
    const result = await createDraft(account, { from, to, subject, html, folder_path, attachments });
    console.log('[draft] Success:', result);
    return res.json(result);
  } catch (err) {
    console.error('[draft] Error:', err.message);
    return res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.WORKER_PORT || 3005;

app.listen(PORT, '0.0.0.0', () => {
  console.log(`email-worker listening on port ${PORT}`);
  startAllListeners().catch(console.error);
});
