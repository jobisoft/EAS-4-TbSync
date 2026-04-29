/**
 * Helpers for the OPTIONS-negotiated `MS-ASProtocolCommands` list,
 * stored on `account.custom.allowedEasCommands` after `connect.mjs`
 * runs the OPTIONS probe.
 *
 * Storage shape is an array of command names (e.g.
 * `["FolderSync", "Sync", "Search", …]`). The legacy add-on stored a
 * comma-separated string; the `eas.legacy-migration` upgrade body
 * normalizes that into the canonical array, so this module never needs
 * to deal with the string form.
 */

function readCommandList(account) {
  const cmds = account?.custom?.allowedEasCommands;
  return Array.isArray(cmds) ? cmds : null;
}

/** Strict check: did the OPTIONS probe explicitly advertise this
 *  command? Returns false when the probe never ran (no list yet) or
 *  when the list is empty. Use this for capabilities that should only
 *  light up when the server confirms support — e.g. Search/GAL or
 *  Settings/DeviceInformation. */
export function easCommandAdvertised(account, command) {
  const cmds = readCommandList(account);
  return !!cmds && cmds.includes(command);
}

/** Permissive check: assume the command is available unless the probe
 *  explicitly says otherwise. Returns true when the list is missing /
 *  empty (no negative information) and falls back to a strict check
 *  when the list is populated. Use this for commands that nearly all
 *  servers support — e.g. GetItemEstimate — where the cost of skipping
 *  a real-world server that simply omits the list is worse than the
 *  cost of attempting an unsupported command and getting a clean
 *  rejection. */
export function easCommandLikelyAvailable(account, command) {
  const cmds = readCommandList(account);
  if (!cmds || cmds.length === 0) return true;
  return cmds.includes(command);
}
