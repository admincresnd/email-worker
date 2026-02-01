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

export async function sendOutlookEmail(account, { from, to, subject, html, inReplyTo, references, conversationId, attachments }) {
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

  if (conversationId) {
    message.conversationId = conversationId;
  }

  if (inReplyTo) {
    message.internetMessageHeaders = message.internetMessageHeaders || [];
    message.internetMessageHeaders.push({ name: 'In-Reply-To', value: inReplyTo });
  }

  if (references) {
    message.internetMessageHeaders = message.internetMessageHeaders || [];
    message.internetMessageHeaders.push({ name: 'References', value: references });
  }

  const graphAttachments = parseAttachments(attachments);

  if (graphAttachments.length > 0) {
    message.attachments = graphAttachments;
    message.hasAttachments = true;
  }

  const result = await graphRequest(account, `${userPath(account)}/sendMail`, {
    method: 'POST',
    body: {
      message,
      saveToSentItems: true,
    },
  });

  console.log(`[outlook] Email sent from="${from}" to="${to}" subject="${subject}"`);

  return {
    success: true,
    messageId: null,
    response: 'Sent via Graph API',
  };
}
