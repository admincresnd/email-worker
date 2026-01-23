import { createImapClient } from './client.js';

export async function createDraft(account, { from, to, subject, html, folder_path }) {
  console.log(`[draft] Starting draft creation: from="${from}" to="${to}" subject="${subject}" folder="${folder_path}"`);
  
  const client = createImapClient(account);

  try {
    console.log(`[draft] Connecting to IMAP server: ${account.imap_host}:${account.imap_port}`);
    await client.connect();
    console.log(`[draft] Successfully connected to IMAP server`);

    console.log(`[draft] Building raw draft email`);
    const rawDraft = buildRawDraft({ from, to, subject, html });
    console.log(`[draft] Raw draft built, size: ${rawDraft.length} bytes`);

    console.log(`[draft] Appending draft to folder "${folder_path}" with \\Draft flag`);
    const result = await client.append(folder_path, rawDraft, ['\\Draft']);
    console.log(`[draft] Successfully created draft in folder="${folder_path}" uid=${result?.uid || 'unknown'}`);

    return {
      success: true,
      uid: result?.uid,
      folder: folder_path,
    };
  } catch (error) {
    console.error(`[draft] Error creating draft:`, error.message);
    throw error;
  } finally {
    try {
      console.log(`[draft] Closing IMAP connection`);
      await client.logout();
      console.log(`[draft] IMAP connection closed successfully`);
    } catch (logoutError) {
      console.warn(`[draft] Warning: Error during logout:`, logoutError.message);
    }
  }
}

function buildRawDraft({ from, to, subject, html }) {
  const messageId = `<${Date.now()}.${Math.random().toString(36).slice(2)}@draft>`;
  const date = new Date().toUTCString();
  
  let headers = [
    `From: ${from}`,
    `To: ${to}`,
    `Subject: ${subject}`,
    `Date: ${date}`,
    `Message-ID: ${messageId}`,
    `MIME-Version: 1.0`,
    `Content-Type: text/html; charset=utf-8`,
    `Content-Transfer-Encoding: quoted-printable`,
  ];

  const encodedBody = quotedPrintableEncode(html || '');

  return headers.join('\r\n') + '\r\n\r\n' + encodedBody;
}

function quotedPrintableEncode(str) {
  const bytes = Buffer.from(str, 'utf-8');
  let result = '';
  for (const byte of bytes) {
    if ((byte >= 33 && byte <= 126 && byte !== 61) || byte === 32 || byte === 9) {
      result += String.fromCharCode(byte);
    } else {
      result += '=' + byte.toString(16).toUpperCase().padStart(2, '0');
    }
  }
  return result;
}