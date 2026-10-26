// The command an AI app starts the tools with, as the AI Tools window shows
// it (ai-tools.js): Rutba Office's own program run as Node on a file the AI
// app can find on disk, with the profile the lines were copied from.
//
// An installed copy names resources\mcp\rutba-office-mcp.cjs, a small real
// file shipped beside the archive (electron-builder extraResources) that
// loads the archive's mcp/node.js. Until 2.18 the lines named
// resources\app.asar\main\mcp\node.js, which only Electron can read: an AI app
// that checked the path first said it was not there (owner, 2026-10-26).
// A copy run from the source names mcp/node.js itself.
//
// Kept apart from ai-tools.js, which needs Electron, so tests can ask it.

import path from 'node:path';

/** The file beside the archive, under resources (resources/mcp in the source). */
export const LAUNCHER = path.join('mcp', 'rutba-office-mcp.cjs');

/** Where the screens say Office is installed, as docs/AI-TOOLS.md prints it. */
export const INSTALLED = 'C:\\Users\\you\\AppData\\Local\\Programs\\Rutba Office';

/**
 * { command, args, env } for the AI app's list of MCP servers.
 *
 * packaged / resourcesPath: an installed copy and its resources folder.
 * appDir: the app's folder when run from the source.
 * profile: the profile settled at the start ({ id, folder, baseNamed }).
 * screens: a screens run for the website, which shows the lines an installed
 * copy shows: the development copy's own paths mean nothing to somebody
 * reading the site.
 */
export function mcpCommand({ execPath, appDir, packaged = false, resourcesPath = '', profile = null, screens = false }) {
  const env = { ELECTRON_RUN_AS_NODE: '1' };
  const named = ['--profile', profile?.id || 'default'];
  if (screens && !packaged) {
    return { command: path.win32.join(INSTALLED, 'Rutba Office.exe'), args: [path.win32.join(INSTALLED, 'resources', LAUNCHER), '--profile', 'default'], env };
  }
  // A base folder named with --user-data-dir (the checks name one) is not where the tools look; say it.
  if (profile?.baseNamed && profile.folder) env.RUTBA_OFFICE_PROFILE = profile.folder;
  if (packaged) return { command: execPath, args: [path.join(resourcesPath, LAUNCHER), ...named], env };
  return { command: execPath, args: [path.join(appDir, 'main', 'mcp', 'node.js'), ...named], env };
}
