import nodemailer from 'nodemailer';
import { createImapClient } from '../imap/client.js';

export async function sendEmail(account, { from, to, subject, html, inReplyTo, references, folder_path, attachments }) {
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

  if (attachments) {
    let parsed = attachments;
    if (typeof parsed === 'string') {
      try { parsed = JSON.parse(parsed); } catch { parsed = null; }
    }
    if (parsed && typeof parsed === 'object') {
      const list = Array.isArray(parsed) ? parsed : Object.values(parsed);
      if (list.length > 0) {
        mailOptions.attachments = list.map(a => ({
          filename: a.filename,
          content: a.content,
          encoding: a.encoding || 'base64',
          contentType: a.contentType,
        }));
      }
    }
  }

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
  const date = new Date().toUTCString();
  const hasAttachments = mailOptions.attachments && mailOptions.attachments.length > 0;
  const boundary = `----=_Part_${Date.now()}_${Math.random().toString(36).slice(2)}`;

  let headers = [
    `From: ${mailOptions.from}`,
    `To: ${mailOptions.to}`,
    `Subject: ${mailOptions.subject}`,
    `Date: ${date}`,
    `Message-ID: ${mailOptions.messageId}`,
    `MIME-Version: 1.0`,
  ];

  if (mailOptions.inReplyTo) {
    headers.push(`In-Reply-To: ${mailOptions.inReplyTo}`);
  }

  if (mailOptions.references) {
    headers.push(`References: ${mailOptions.references}`);
  }

  if (!hasAttachments) {
    headers.push(`Content-Type: text/html; charset=utf-8`);
    headers.push(`Content-Transfer-Encoding: quoted-printable`);
    const encodedBody = quotedPrintableEncode(mailOptions.html || '');
    return headers.join('\r\n') + '\r\n\r\n' + encodedBody;
  }

  headers.push(`Content-Type: multipart/mixed; boundary="${boundary}"`);

  let body = headers.join('\r\n') + '\r\n\r\n';

  body += `--${boundary}\r\n`;
  body += `Content-Type: text/html; charset=utf-8\r\n`;
  body += `Content-Transfer-Encoding: quoted-printable\r\n\r\n`;
  body += quotedPrintableEncode(mailOptions.html || '') + '\r\n';

  for (const att of mailOptions.attachments) {
    body += `--${boundary}\r\n`;
    body += `Content-Type: ${att.contentType || 'application/octet-stream'}; name="${att.filename}"\r\n`;
    body += `Content-Disposition: attachment; filename="${att.filename}"\r\n`;
    body += `Content-Transfer-Encoding: base64\r\n\r\n`;
    body += att.content + '\r\n';
  }

  body += `--${boundary}--\r\n`;

  return body;
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
