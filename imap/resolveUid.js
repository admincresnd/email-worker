import { ImapFlow } from 'imapflow';

export async function resolveUid({
  imap_host,
  imap_port,
  imap_secure,
  imap_username,
  imap_secret,
  folder,
  message_id,
}) {
  const client = new ImapFlow({
    host: imap_host,
    port: imap_port,
    secure: imap_secure,
    auth: {
      user: imap_username,
      pass: imap_secret,
    },
  });

  await client.connect();

  try {
    await client.mailboxOpen(folder);

    const uids = await client.search({
      header: ['Message-ID', message_id],
    });

    if (!uids || uids.length === 0) {
      return null;
    }

    return uids[0];
  } finally {
    await client.logout();
  }
}
