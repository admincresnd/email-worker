import { ImapFlow } from 'imapflow';

export function createImapClient(account) {
  return new ImapFlow({
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
}
