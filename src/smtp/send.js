import nodemailer from 'nodemailer';
import { createImapClient } from '../imap/client.js';

export async function sendEmail(account, { from, to, subject, html, inReplyTo, references, folder_path }) {
  console.log(`[smtp] Creating transporter: host=${account.smtp_host} port=${account.smtp_port} user=${account.smtp_username}`);
  const transporter = nodemailer.createTransport({
    host: account.smtp_host,
    port: account.smtp_port,
    secure: account.smtp_port === 465,
    auth: {
      user: account.smtp_username,
      pass: account.smtp_secret,
    },
  });

  const messageId = `<${Date.now()}.${Math.random().toString(36).slice(2)}@${account.smtp_host}>`;

  const mailOptions = {
    from,
    to,
    subject,
    html,
    messageId,
  };

  if (inReplyTo) {
    mailOptions.inReplyTo = inReplyTo;
  }

  if (references) {
    mailOptions.references = references;
  }

  console.log(`[smtp] Sending email from="${from}" to="${to}" subject="${subject}"`);

  const info = await transporter.sendMail(mailOptions);
  console.log(`[smtp] Email sent: messageId=${info.messageId}`);

  if (folder_path) {
    await appendToSentFolder(account, mailOptions, folder_path);
  }

  return {
    success: true,
    messageId: info.messageId,
    response: info.response,
  };
}

async function appendToSentFolder(account, mailOptions, folderPath) {
  const client = createImapClient(account);

  try {
    await client.connect();

    const rawEmail = buildRawEmail(mailOptions);

    const result = await client.append(folderPath, rawEmail, ['\\Seen']);
    console.log(`[smtp] Appended to Sent folder="${folderPath}" uid=${result?.uid || 'unknown'}`);

    return result;
  } finally {
    try {
      await client.logout();
    } catch {}
  }
}

function buildRawEmail(mailOptions) {
  const boundary = `----=_Part_${Date.now()}_${Math.random().toString(36).slice(2)}`;
  const date = new Date().toUTCString();
  
  let headers = [
    `From: ${mailOptions.from}`,
    `To: ${mailOptions.to}`,
    `Subject: ${mailOptions.subject}`,
    `Date: ${date}`,
    `Message-ID: ${mailOptions.messageId}`,
    `MIME-Version: 1.0`,
    `Content-Type: text/html; charset=utf-8`,
    `Content-Transfer-Encoding: quoted-printable`,
  ];

  if (mailOptions.inReplyTo) {
    headers.push(`In-Reply-To: ${mailOptions.inReplyTo}`);
  }

  if (mailOptions.references) {
    headers.push(`References: ${mailOptions.references}`);
  }

  const encodedBody = quotedPrintableEncode(mailOptions.html || '');

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
