export const STARTUP_DELAY_MS = 1500;
export const RECONNECT_BASE_MS = 5000;
export const RECONNECT_MAX_MS = 60000;
export const POLL_INTERVAL_MS = 30_000;

export const ALIGN_SCAN_BATCH_SIZE = 500;
export const RESOLVE_UID_SCAN_BATCH_SIZE = 500;

export function mustEnv(name) {
  const v = process.env[name];
  if (!v) throw new Error(`Missing required env: ${name}`);
  return v;
}

export function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

export function normalizeFolderPath(p) {
  return p;
}

export function parseEmailsSince(value) {
  if (!value) return null;
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return null;
  return d;
}
