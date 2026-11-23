// What Windows knows of its printers and Electron no longer says.
//
// Electron's list of printers once named the default one and gave each its status. Electron 44 gives
// a name and nothing else: no default, no status. A print dialog that opens on the first printer of
// that list opens on whichever sorts first (a label printer, the fax), and one that sends the job
// straight to the printer it shows must show the right one. So the default printer and the state of
// each are asked of Windows here.
//
// Two reads, because they cost differently. The default printer is one value in the registry, read
// by reg.exe in a few hundredths of a second: the dialog has it as it opens. The states (offline,
// paused, out of paper) are the spooler's, asked through PowerShell, which takes a second or two to
// start: the dialog asks for them after it has drawn, and marks the printers when they arrive.
//
// Neither read can stop a print. One that fails or takes too long answers with nothing known, and the
// dialog shows the printers unmarked. On another system nothing is asked: Electron's own options are
// read for what CUPS says of a printer, and no default is named.

import { execFile } from 'node:child_process';
import { regRunner } from './browsers.js';

/** The spooler's status bits (PRINTER_STATUS_*), the ones a person can act on. */
const PAUSED = 0x1;
const ERROR = 0x2;
const PAPER_JAM = 0x8;
const PAPER_OUT = 0x10;
const PAPER_PROBLEM = 0x40;
const OFFLINE = 0x80;
const BIN_FULL = 0x800;
const NOT_AVAILABLE = 0x1000;
const NO_TONER = 0x40000;
const USER_INTERVENTION = 0x100000;
const OUT_OF_MEMORY = 0x200000;
const DOOR_OPEN = 0x400000;
const SERVER_UNKNOWN = 0x800000;
const TROUBLE = ERROR | PAPER_JAM | PAPER_PROBLEM | BIN_FULL | NO_TONER | USER_INTERVENTION | OUT_OF_MEMORY | DOOR_OPEN;

/**
 * A printer's state in one word, from what Windows says of it: '' when it is ready (or nothing is
 * known), 'offline', 'paused', 'paper' (it has none) or 'trouble' (it wants a person: a jam, a door
 * open, no toner). `offline` is the printer's own "use offline" setting, `bits` the spooler's status,
 * `extended` and `error` the two fuller readings Windows keeps beside it (7 and 9 are offline there).
 * Offline first: a printer that cannot be reached says nothing true about its paper.
 */
export function printerState({ offline = false, bits = 0, extended = 0, error = 0 } = {}) {
  const b = Number(bits) || 0;
  const x = Number(extended) || 0;
  const e = Number(error) || 0;
  if (offline || b & (OFFLINE | NOT_AVAILABLE | SERVER_UNKNOWN) || x === 7 || x === 11 || e === 9) return 'offline';
  if (b & PAUSED || x === 8) return 'paused';
  if (b & PAPER_OUT || e === 4) return 'paper';
  if (b & TROUBLE || x === 9 || [6, 7, 8, 10, 11].includes(e)) return 'trouble';
  return '';
}

/**
 * The default printer's name, from what `reg query ... /v Device` prints: the value is the name, then
 * ",winspool," and a port. A printer's name holds no comma (Windows does not allow one), so the name
 * is everything before the first. Null when the text names none.
 */
export function defaultPrinterOf(text) {
  const m = /^\s*Device\s+REG_SZ\s+(.+?)\s*$/im.exec(String(text || ''));
  if (!m) return null;
  const name = m[1].split(',')[0].trim();
  return name || null;
}

/** The states Windows gave, by printer name: `{ [name]: { default, state } }`. Text that is not the answer gives null. */
export function printerStatesOf(text) {
  let list;
  try {
    list = JSON.parse(String(text || '').trim());
  } catch {
    return null;
  }
  if (list && !Array.isArray(list)) list = [list];
  if (!Array.isArray(list)) return null;
  const out = {};
  for (const p of list) {
    if (!p || typeof p.n !== 'string' || !p.n) continue;
    out[p.n] = { default: Boolean(p.d), state: printerState({ offline: Boolean(p.w), bits: p.s, extended: p.x, error: p.e }) };
  }
  return out;
}

/** What CUPS says of a printer, in the options Electron hands on (macOS and Linux): the same one word. */
export function cupsState(options = {}) {
  const reasons = String(options?.['printer-state-reasons'] || '');
  if (/offline|connecting-to-device|shutdown/.test(reasons)) return 'offline';
  if (String(options?.['printer-state'] || '') === '5' || /paused/.test(reasons)) return 'paused';
  if (/media-empty|media-needed/.test(reasons)) return 'paper';
  if (/jam|door-open|cover-open|toner-empty|marker-supply-empty/.test(reasons)) return 'trouble';
  return '';
}

// One line of JSON, in UTF-8 whatever the console's code page: a printer's name is whatever its owner typed.
const SCRIPT = [
  '$ErrorActionPreference = "Stop"',
  '[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false',
  '$list = @(Get-CimInstance Win32_Printer -Property Name,Default,WorkOffline,PrinterState,ExtendedPrinterStatus,DetectedErrorState | ForEach-Object { [pscustomobject]@{ n = [string]$_.Name; d = [bool]$_.Default; w = [bool]$_.WorkOffline; s = [int64]$_.PrinterState; x = [int]$_.ExtendedPrinterStatus; e = [int]$_.DetectedErrorState } })',
  'ConvertTo-Json -InputObject $list -Compress',
].join('; ');

/**
 * A PowerShell to ask: `() => Promise<string|null>`, null when it failed or did not answer in time.
 * The time is generous: three seconds on a quiet computer was more than eight on a busy one, and
 * nothing waits on the answer but the states in the list.
 */
export const spoolerRunner = ({ file = 'powershell.exe', before = ['-NoProfile', '-NonInteractive', '-Command'], timeout = 20000 } = {}) => () =>
  new Promise((resolve) => {
    try {
      execFile(file, [...before, SCRIPT], { windowsHide: true, timeout, maxBuffer: 4 * 1024 * 1024, encoding: 'utf8' }, (err, stdout) => resolve(err ? null : String(stdout)));
    } catch {
      resolve(null);
    }
  });

const DEFAULT_KEY = 'HKCU\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Windows';

/**
 * @param {{ platform?: string, reg?: (args: string[]) => Promise<string|null>, spooler?: () => Promise<string|null> }} [opts] for the tests
 */
export function createPrinterFacts({ platform = process.platform, reg = regRunner({ timeout: 4000 }), spooler = spoolerRunner() } = {}) {
  // One question at a time: a dialog opened again while the system is still answering shares the answer.
  let asking = null;
  return {
    /** The default printer's name, quickly; null when Windows does not say, and on another system. */
    defaultName: async () => {
      if (platform !== 'win32') return null;
      try {
        return defaultPrinterOf(await reg(['query', DEFAULT_KEY, '/v', 'Device']));
      } catch {
        return null;
      }
    },
    /** Every printer's state and which is the default, slowly; null when nothing is known. */
    states: () => {
      if (platform !== 'win32') return Promise.resolve(null);
      asking ||= (async () => {
        try {
          return printerStatesOf(await spooler());
        } catch {
          return null;
        } finally {
          asking = null;
        }
      })();
      return asking;
    },
  };
}
