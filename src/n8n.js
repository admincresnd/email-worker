import fetch from 'node-fetch';

export async function forwardToN8n(payload) {
  const url = process.env.N8N_WEBHOOK_URL;
  if (!url) {
    console.error('[n8n] Missing N8N_WEBHOOK_URL (skipping send)');
    return false;
  }

  console.log(`[n8n] Sending to webhook: ${url}`);

  try {
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });

    if (!res.ok) {
      const body = await res.text().catch(() => '');
      console.error(`[n8n] webhook failed: ${res.status} - ${body}`);
      return false;
    }
    console.log(`[n8n] webhook success: ${res.status}`);
    return true;
  } catch (err) {
    console.error('[n8n] unreachable:', err.message);
    return false;
  }
}
