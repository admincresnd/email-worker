import { graphRequest, userPath } from './client.js';

async function resolveFolderId(account, folderName) {
  const data = await graphRequest(
    account,
    `${userPath(account)}/mailFolders`
  );

  if (!data?.value) throw new Error('Could not list mail folders');

  const match = data.value.find(
    (f) =>
      f.displayName.toLowerCase() === folderName.toLowerCase() ||
      f.id === folderName
  );

  if (!match) {
    const childFolders = [];
    for (const parent of data.value) {
      try {
        const children = await graphRequest(
          account,
          `${userPath(account)}/mailFolders/${parent.id}/childFolders`
        );
        if (children?.value) childFolders.push(...children.value);
      } catch {}
    }

    const childMatch = childFolders.find(
      (f) =>
        f.displayName.toLowerCase() === folderName.toLowerCase() ||
        f.id === folderName
    );

    if (childMatch) return childMatch.id;
    throw new Error(`Folder "${folderName}" not found`);
  }

  return match.id;
}

export async function moveOutlookEmail(account, { outlook_id, folder, mark_as_seen, flagged }) {
  if (!outlook_id) throw new Error('outlook_id is required for Outlook move');

  const base = `${userPath(account)}/messages/${outlook_id}`;

  if (mark_as_seen === true || mark_as_seen === false) {
    await graphRequest(account, base, {
      method: 'PATCH',
      body: { isRead: mark_as_seen },
    });
    console.log(`[outlook] Marked ${outlook_id} as ${mark_as_seen ? 'read' : 'unread'}`);
  }

  if (flagged === true || flagged === false) {
    await graphRequest(account, base, {
      method: 'PATCH',
      body: {
        flag: { flagStatus: flagged ? 'flagged' : 'notFlagged' },
      },
    });
    console.log(`[outlook] Marked ${outlook_id} as ${flagged ? 'flagged' : 'unflagged'}`);
  }

  if (folder) {
    const destinationId = await resolveFolderId(account, folder);

    const result = await graphRequest(account, `${base}/move`, {
      method: 'POST',
      body: { destinationId },
    });

    console.log(`[outlook] Moved ${outlook_id} to folder="${folder}"`);

    return {
      success: true,
      destination: folder,
      newId: result?.id || null,
    };
  }

  return { success: true, destination: null, newId: null };
}
