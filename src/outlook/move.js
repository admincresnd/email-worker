import { graphRequest, userPath } from './client.js';

async function resolveFolderId(account, folderName) {
  const parts = folderName.split('/').map((p) => p.trim()).filter(Boolean);

  const data = await graphRequest(
    account,
    `${userPath(account)}/mailFolders`
  );

  if (!data?.value) throw new Error('Could not list mail folders');

  const findFolder = (folders, name) =>
    folders.find(
      (f) => f.displayName.toLowerCase() === name.toLowerCase() || f.id === name
    );

  if (parts.length === 1) {
    const match = findFolder(data.value, parts[0]);
    if (match) return match.id;

    for (const parent of data.value) {
      try {
        const children = await graphRequest(
          account,
          `${userPath(account)}/mailFolders/${parent.id}/childFolders`
        );
        if (children?.value) {
          const childMatch = findFolder(children.value, parts[0]);
          if (childMatch) return childMatch.id;
        }
      } catch {}
    }

    throw new Error(`Folder "${folderName}" not found`);
  }

  let currentFolder = findFolder(data.value, parts[0]);
  if (!currentFolder) throw new Error(`Folder "${parts[0]}" not found`);

  for (let i = 1; i < parts.length; i++) {
    const children = await graphRequest(
      account,
      `${userPath(account)}/mailFolders/${currentFolder.id}/childFolders`
    );
    if (!children?.value) throw new Error(`Folder "${parts[i]}" not found under "${parts[i - 1]}"`);
    currentFolder = findFolder(children.value, parts[i]);
    if (!currentFolder) throw new Error(`Folder "${parts[i]}" not found under "${parts[i - 1]}"`);
  }

  return currentFolder.id;
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
