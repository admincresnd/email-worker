import { createImapClient } from './client.js';
import { normalizeFolderPath, RESOLVE_UID_SCAN_BATCH_SIZE } from '../config.js';

export async function listFolders(account) {
  const client = createImapClient(account);
  try {
    await client.connect();
    const mailboxes = await client.list();
    return mailboxes.map((m) => m.path);
  } finally {
    try {
      await client.logout();
    } catch {}
  }
}

export async function resolveUid(account, { folder, message_id }) {
  const client = createImapClient(account);

  try {
    console.log('[resolve] Connecting to IMAP...');
    await client.connect();
    console.log('[resolve] Connected');

    const folderPath = normalizeFolderPath(folder);
    console.log(`[resolve] Listing mailboxes...`);

    const mailboxes = await client.list();
    console.log(`[resolve] Found ${mailboxes.length} mailboxes`);
    if (!mailboxes.find((m) => m.path === folderPath)) {
      throw new Error('Mailbox not found');
    }

    console.log(`[resolve] Opening folder: ${folderPath}`);
    await client.mailboxOpen(folderPath);
    console.log('[resolve] Folder opened');

    const target = message_id.replace(/[<>]/g, '');
    console.log(`[resolve] Looking for message_id: ${target}`);

    const mb = client.mailbox;
    const exists = mb?.exists || 0;
    console.log(`[resolve] Folder has ${exists} messages`);
    if (!exists) {
      throw new Error('Message-ID not found');
    }

    let seqStart = 1;

    while (seqStart <= exists) {
      const seqEnd = Math.min(seqStart + RESOLVE_UID_SCAN_BATCH_SIZE - 1, exists);
      const range = `${seqStart}:${seqEnd}`;
      console.log(`[resolve] Fetching batch ${range}...`);

      for await (const msg of client.fetch(range, {
        uid: true,
        envelope: true,
      })) {
        const mid = msg.envelope?.messageId?.replace(/[<>]/g, '');
        console.log(`[resolve] Checking uid=${msg.uid} mid=${mid}`);
        if (mid && mid === target) {
          console.log(`[resolve] FOUND! uid=${msg.uid}`);
          return { uid: msg.uid };
        }
      }

      seqStart = seqEnd + 1;
    }

    console.log('[resolve] Message-ID not found in folder');
    throw new Error('Message-ID not found');
  } finally {
    try {
      console.log('[resolve] Closing connection...');
      client.close();
      console.log('[resolve] Connection closed');
    } catch {}
  }
}
