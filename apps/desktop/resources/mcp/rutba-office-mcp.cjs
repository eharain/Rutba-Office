// The AI tools' start file for an installed Rutba Office (main/mcp/command.js).
// An AI app runs it with Rutba Office's own program as Node:
//
//   ELECTRON_RUN_AS_NODE=1 "Rutba Office.exe" "<install>\resources\mcp\rutba-office-mcp.cjs" --profile <id>
//
// It is a real file beside the app's archive (electron-builder
// extraResources), so an AI app that checks the path first finds it, and it
// loads the tools from the archive next to it, which Rutba Office's program
// reads when run as Node. The words after it (--profile) reach the tools as
// they are.

'use strict';

const path = require('node:path');
const { pathToFileURL } = require('node:url');

/** The tools in the archive beside this file's folder: resources\app.asar\main\mcp\node.js. */
function archiveEntry(dir = __dirname) {
  return path.join(dir, '..', 'app.asar', 'main', 'mcp', 'node.js');
}

module.exports = { archiveEntry };

if (require.main === module) {
  import(pathToFileURL(archiveEntry()).href).catch((err) => {
    process.stderr.write(`Rutba Office's AI tools could not start: ${err && err.stack ? err.stack : err}\n`);
    process.exit(1);
  });
}
