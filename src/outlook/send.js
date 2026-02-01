import { graphRequest, userPath } from './client.js';

function parseAttachments(attachments) {
  if (!attachments) return [];

  let parsed = attachments;
  if (typeof parsed === 'string') {
    try { parsed = JSON.parse(parsed); } catch { return []; }
  }
  if (!parsed || typeof parsed !== 'object') return [];

  const list = Array.isArray(parsed) ? parsed : Object.values(parsed);

  return list.map((a) => ({
    '@odata.type': '#microsoft.graph.fileAttachment',
    name: a.filename,
    contentType: a.contentType || 'application/octet-stream',
    contentBytes: a.content,
  }));
}

export async function sendOutlookEmail(account, { from, to, subject, html, attachments }) {
  const toRecipients = to.split(',').map((addr) => ({
    emailAddress: { address: addr.trim() },
  }));

  const message = {
    subject,
    body: {
      contentType: 'HTML',
      content: html || '',
    },
    from: {
      emailAddress: { address: from },
    },
    toRecipients,
  };

  const graphAttachments = parseAttachments(attachments);

  if (graphAttachments.length > 0) {
    message.attachments = graphAttachments;
    message.hasAttachments = true;
  }

  await graphRequest(account, `${userPath(account)}/sendMail`, {
    method: 'POST',
    body: {
      message,
      saveToSentItems: true,
    },
  });

  console.log(`[outlook/send] Email sent from="${from}" to="${to}" subject="${subject}"`);

  return {
    success: true,
    response: 'Sent via Graph API',
  };
}

export async function replyOutlookEmail(account, { outlook_id, from, to, subject, html, attachments }) {
  const message = {
    body: {
      contentType: 'HTML',
      content: html || '',
    },
  };

  if (to) {
    message.toRecipients = to.split(',').map((addr) => ({
      emailAddress: { address: addr.trim() },
    }));
  }

  const graphAttachments = parseAttachments(attachments);

  if (graphAttachments.length > 0) {
    message.attachments = graphAttachments;
  }

  await graphRequest(account, `${userPath(account)}/messages/${outlook_id}/reply`, {
    method: 'POST',
    body: {
      message,
    },
  });

  console.log(`[outlook/reply] Reply sent to outlook_id="${outlook_id}" subject="${subject}"`);

  return {
    success: true,
    response: 'Replied via Graph API',
  };
}
