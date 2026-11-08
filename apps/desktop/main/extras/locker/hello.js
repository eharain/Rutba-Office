// Windows Hello, for "this computer only" unlocking. Electron has no Windows
// Hello; the system has it through WinRT, and Windows PowerShell can call
// WinRT without a separate native binary. A key is made in the TPM-backed
// KeyCredentialManager and asked to sign a fixed challenge: the signature is
// the same each time and only that credential can make it, so hashing it gives
// the secret that wraps the locker's key on this computer. The password
// remains the master; Hello is only ever an additional way in.
//
// None of this is reachable from a check run: it needs the person at the
// keyboard, and the checks never touch the desktop.

import crypto from 'node:crypto';
import { execFile } from 'node:child_process';

const CRED = 'RutbaOfficeLocker';
const CHALLENGE = 'rutba-locker:v1:hello-challenge';

const SCRIPT = (body) => `
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Security.Credentials.KeyCredentialManager, Windows.Security.Credentials, ContentType=WindowsRuntime]
$null = [Windows.Security.Cryptography.CryptographicBuffer, Windows.Security.Cryptography, ContentType=WindowsRuntime]
$asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.ToString() -match 'IAsyncOperation' })[0]
function Await($op, [Type]$type) { $t = $asTask.MakeGenericMethod($type).Invoke($null, @($op)); [void]$t.Wait(-1); $t.Result }
${body}
`;

function run(body) {
  return new Promise((resolve, reject) => {
    execFile('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', SCRIPT(body)], { windowsHide: true, timeout: 120000, maxBuffer: 1 << 20 }, (err, stdout) => {
      if (err) reject(err);
      else resolve(String(stdout).trim());
    });
  });
}

const signBody = (create) => `
$mgr = [Windows.Security.Credentials.KeyCredentialManager]
${create ? `$c = Await ($mgr::RequestCreateAsync('${CRED}', [Windows.Security.Credentials.KeyCredentialCreationOption]::ReplaceExisting)) ([Windows.Security.Credentials.KeyCredentialRetrievalResult])` : `$c = Await ($mgr::OpenAsync('${CRED}')) ([Windows.Security.Credentials.KeyCredentialRetrievalResult])`}
if ($c.Status -ne 'Success') { Write-Output "status:$($c.Status)"; exit 0 }
$b = [Windows.Security.Cryptography.CryptographicBuffer]::ConvertStringToBinary('${CHALLENGE}', [Windows.Security.Cryptography.BinaryStringEncoding]::Utf8)
$s = Await ($c.Credential.RequestSignAsync($b)) ([Windows.Security.Credentials.KeyCredentialOperationResult])
if ($s.Status -ne 'Success') { Write-Output "status:$($s.Status)"; exit 0 }
Write-Output ([Windows.Security.Cryptography.CryptographicBuffer]::EncodeToBase64String($s.Result))
`;

async function secretFrom(create) {
  const out = await run(signBody(create));
  if (out.startsWith('status:') || !out) throw new Error(`Windows Hello did not complete (${out || 'no answer'}).`);
  return crypto.createHash('sha256').update(Buffer.from(out, 'base64')).digest();
}

/**
 * @returns {{ available: () => Promise<boolean>, enroll: () => Promise<Buffer>, secret: () => Promise<Buffer> }}
 */
export function helloProvider() {
  let known = null;
  return {
    available: async () => {
      if (process.platform !== 'win32') return false;
      if (known !== null) return known;
      known = await run(`$r = Await ([Windows.Security.Credentials.KeyCredentialManager]::IsSupportedAsync()) ([bool]); Write-Output $r`).then((o) => o.trim() === 'True', () => false);
      return known;
    },
    enroll: () => secretFrom(true),
    secret: () => secretFrom(false),
  };
}

/** A check run's stand-in: a fixed secret in place of the person's PIN. Honoured only by a copy run from source. */
export function fakeHello() {
  const secret = crypto.createHash('sha256').update('rutba-locker:fake-hello').digest();
  return { available: async () => true, enroll: async () => secret, secret: async () => secret };
}
