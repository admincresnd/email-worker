import express from 'express';
import { ImapFlow } from 'imapflow';
import fetch from 'node-fetch';
import dotenv from 'dotenv';

dotenv.config();

const app = express();
app.use(express.json());

/**
 * Called when a new email is detected
 */
async function forwardToN8n(payload) {
  await fetch(process.env.N8N_WEBHOOK_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });
}

/**
 * IMAP listener (simplified)
 */
async function listenToInbox(account) {
  const client = new ImapFlow({
    host: account.imap_host,
    port: account.imap_port,
    secure: true,
    auth: {
      user: account.imap_user,
      pass: account.imap_password
    }
  });

  await client.connect();
  await client.mailboxOpen('INBOX');

  for await (let msg of client.fetch('1:*', { envelope: true, source: true })) {
    await forwardToN8n({
      email_account_id: account.id,
      venue_id: account.venue_id,
      subject: msg.envelope.subject,
      from: msg.envelope.from,
      raw: msg.source.toString()
    });
  }
}

/**
 * Action endpoint (called by n8n)
 */
app.post('/email-action', async (req, res) => {
  const { action, account } = req.body;

  if (action.type === 'move') {
    // IMAP move logic here
  }

  if (action.type === 'send') {
    // SMTP send logic here
  }

  res.json({ status: 'ok' });
});

app.listen(process.env.WORKER_PORT, () => {
  console.log(`Email worker running on ${process.env.WORKER_PORT}`);
});
