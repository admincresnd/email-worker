import { createImapClient } from './client.js';

export async function moveEmail(account, { uid, folder, source_folder, mark_as_seen, flagged }) {
  const client = createImapClient(account);
  const sourceMailbox = source_folder || 'INBOX';

  try {
    await client.connect();

    const lock = await client.getMailboxLock(sourceMailbox);

    try {
      if (mark_as_seen === true) {
        await client.messageFlagsAdd({ uid }, ['\\Seen'], { uid: true });
        console.log(`[move] Marked uid=${uid} as seen`);
      } else if (mark_as_seen === false) {
        await client.messageFlagsRemove({ uid }, ['\\Seen'], { uid: true });
        console.log(`[move] Marked uid=${uid} as unseen`);
      }

      if (flagged === true) {
        await client.messageFlagsAdd({ uid }, ['\\Flagged'], { uid: true });
        console.log(`[move] Marked uid=${uid} as flagged`);
      } else if (flagged === false) {
        await client.messageFlagsRemove({ uid }, ['\\Flagged'], { uid: true });
        console.log(`[move] Marked uid=${uid} as unflagged`);
      }

      if (sourceMailbox === folder) {
        console.log(`[move] Skipping move - uid=${uid} already in folder="${folder}"`);
        return { success: true, destination: folder, uid: uid, skipped_move: true };
      }

      const result = await client.messageMove({ uid }, folder, { uid: true });
      console.log(`[move] Moved uid=${uid} from="${sourceMailbox}" to="${folder}"`);

      return { success: true, destination: folder, uid: result?.destination?.uid || null };
    } finally {
      lock.release();
    }
  } finally {
    try {
      await client.logout();
    } catch {}
  }
}
