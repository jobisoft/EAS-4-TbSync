/**
 * The name of the local resource (calendar, task list or address book) that
 * is created for a folder, and whether it should follow the folder when the
 * server renames it.
 *
 * Pure, so it can be tested without Thunderbird.
 *
 * `targetName` on a folder row is the live name of the resource: it is set
 * when the resource is created and mirrored from every rename the user makes.
 * A name that differs from the generated one was therefore chosen by the
 * user. The account name is not a factor: it can only be changed while the
 * account is disconnected, which deletes every row and resource, so the
 * account name is always the one the existing names were generated with.
 */

/** What a folder's account setting may be: never follow a server rename,
 *  follow it only while the local name is still the generated one, or
 *  follow it whatever the local name is. */
export const FOLLOW_MODES = ["ignore", "unmodified", "always"];
export const DEFAULT_FOLLOW_MODE = "unmodified";

/** The mode an account's `custom` asks for; anything unknown is the default. */
export function followModeOf(custom) {
  const v = custom?.followServerRenames;
  return FOLLOW_MODES.includes(v) ? v : DEFAULT_FOLLOW_MODE;
}

/** The name a resource is given when it is created for a folder. */
export function genResourceName(displayName, accountName) {
  return `${displayName} (${accountName})`;
}

/** The resources to rename after a folder list was pushed.
 *
 *  `oldRows` are the folder rows as they were before the push and
 *  `newDescriptors` what was pushed; they are matched by `folderId`. A folder
 *  is renamed when its name changed on the server, it has a stored resource
 *  name, and the mode allows it: `always`, or `unmodified` while the stored
 *  name is still the one generated for the old folder name. */
export function planNameFollow({ oldRows, newDescriptors, accountName, mode }) {
  if (mode !== "unmodified" && mode !== "always") return [];
  const byId = new Map((oldRows ?? []).map((row) => [row.folderId, row]));
  const plan = [];
  for (const next of newDescriptors ?? []) {
    const old = byId.get(next.folderId);
    if (!old?.targetName) continue;
    if (!next.displayName || next.displayName === old.displayName) continue;
    if (
      mode === "unmodified" &&
      old.targetName !== genResourceName(old.displayName, accountName)
    ) {
      continue;
    }
    const name = genResourceName(next.displayName, accountName);
    if (name === old.targetName) continue;
    plan.push({
      folderId: old.folderId,
      targetID: old.targetID ?? null,
      targetType: old.targetType,
      name,
    });
  }
  return plan;
}
