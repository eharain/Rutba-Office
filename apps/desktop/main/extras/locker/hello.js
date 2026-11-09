// Windows Hello, for "this computer only" unlocking. Electron has no Windows
// Hello; the system has it through WinRT, and Windows PowerShell can call
// WinRT without a separate native binary. A key is made in the TPM-backed
// KeyCredentialManager and asked to sign a fixed challenge: Windows signs with
// RSA, SHA-256 and PKCS#1 padding, so the signature is the same each time and
// only that credential can make it. Hashing it gives the secret that wraps the
// locker's key on this computer. The password remains the master; Hello is only
// ever an additional way in.
//
// Each locker has its own key in Windows, named by an id kept in its keyring,
// so a second profile on this computer cannot replace the first one's, and
// turning Hello off removes the key instead of leaving it behind.
//
// A check run never reaches a prompt: making a key and signing need the person
// at the keyboard. What it can reach without one is everything up to the
// prompt, which tests/locker-hello.test.js runs against the real system.

import crypto from 'node:crypto';
import { execFile } from 'node:child_process';
import { LockerError } from './format.js';

const CHALLENGE = 'rutba-locker:v1:hello-challenge';
const ID = /^[0-9a-f]{16}$/;
const RESULT = '[Windows.Security.Credentials.KeyCredentialRetrievalResult]';

/** How long a "no" about Windows Hello is believed: a person can set up a PIN while Office is open. */
const RECHECK_MS = 60000;

/** The prompt belongs to a Windows process of its own and can open behind Office, so it is brought forward once each time it appears. */
const FRONT = `
try {
  Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class RutbaLockerWindow {
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr FindWindowW(string c, string n);
  [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
  static IntPtr last = IntPtr.Zero;
  public static void Bring() {
    IntPtr h = FindWindowW("Credential Dialog Xaml Host", null);
    if (h != IntPtr.Zero && h != last) { last = h; SetForegroundWindow(h); }
  }
}
'@
  function Front { [RutbaLockerWindow]::Bring() }
} catch { function Front {} }
`;

const SCRIPT = (body, prompts) => `
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Security.Credentials.KeyCredentialManager, Windows.Security.Credentials, ContentType=WindowsRuntime]
$null = [Windows.Security.Cryptography.CryptographicBuffer, Windows.Security.Cryptography, ContentType=WindowsRuntime]
$tasks = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 }
$asTask = @($tasks | Where-Object { $_.GetParameters()[0].ParameterType.Name -like 'IAsyncOperation?1' })[0]
$asAction = @($tasks | Where-Object { $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]
${prompts ? FRONT : 'function Front {}'}
function Await($op, [Type]$type) {
  $t = $asTask.MakeGenericMethod($type).Invoke($null, @($op))
  while (-not $t.IsCompleted) { Front; Start-Sleep -Milliseconds 100 }
  $t.Result
}
$mgr = [Windows.Security.Credentials.KeyCredentialManager]
${body}
`;

function runPowerShell(body, { prompts = false } = {}) {
  return new Promise((resolve, reject) => {
    execFile('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', SCRIPT(body, prompts)], { windowsHide: true, timeout: 120000, maxBuffer: 1 << 20 }, (err, stdout) => {
      if (err) reject(err);
      else resolve(String(stdout).trim());
    });
  });
}

const credName = (id) => `RutbaOfficeLocker-${id}`;

const signBody = (id, create) => `
$c = Await (${create ? `$mgr::RequestCreateAsync('${credName(id)}', [Windows.Security.Credentials.KeyCredentialCreationOption]::ReplaceExisting)` : `$mgr::OpenAsync('${credName(id)}')`}) (${RESULT})
if ($c.Status -ne 'Success') { Write-Output ('status:' + $c.Status); exit 0 }
$b = [Windows.Security.Cryptography.CryptographicBuffer]::ConvertStringToBinary('${CHALLENGE}', [Windows.Security.Cryptography.BinaryStringEncoding]::Utf8)
$s = Await ($c.Credential.RequestSignAsync($b)) ([Windows.Security.Credentials.KeyCredentialOperationResult])
if ($s.Status -ne 'Success') { Write-Output ('status:' + $s.Status); exit 0 }
Write-Output ([Windows.Security.Cryptography.CryptographicBuffer]::EncodeToBase64String($s.Result))
`;

const forgetBody = (id) => `
$t = $asAction.Invoke($null, @($mgr::DeleteAsync('${credName(id)}')))
[void]$t.Wait(10000)
Write-Output 'done'
`;

const GONE = 'Windows no longer has the Windows Hello key for this locker, which can happen after the Windows PIN is reset. Use your password, then turn Windows Hello on again in Locker settings.';

/** What Windows answered, said for the person. A cancelled prompt has its own code, so the window can stay quiet about it. */
export function helloFailure(status) {
  switch (status) {
    case 'UserCanceled':
    case 'UserPrefersPassword':
      return new LockerError('hello-cancelled', 'Windows Hello was cancelled.');
    case 'NotFound':
      return new LockerError('hello-gone', GONE);
    case 'SecurityDeviceLocked':
      return new LockerError('hello-locked', 'Windows Hello is locked after too many wrong tries. Use your password, or try again later.');
    default:
      return new LockerError('hello-failed', 'Windows Hello did not finish. Use your password instead.');
  }
}

/**
 * @param {{ run?: (body: string, o?: { prompts?: boolean }) => Promise<string> }} [o] a test passes its own run
 * @returns {{ available: () => Promise<boolean>, enroll: (id: string) => Promise<Buffer>, secret: (id: string) => Promise<Buffer>, forget: (id: string) => Promise<void> }}
 */
export function helloProvider({ run = runPowerShell } = {}) {
  let known = null;
  let askedAt = 0;

  async function secretFrom(id, create) {
    if (!ID.test(String(id))) throw helloFailure('NotFound');
    const out = await run(signBody(id, create), { prompts: true }).catch(() => '');
    if (out.startsWith('status:')) throw helloFailure(out.slice(7));
    const signature = Buffer.from(out, 'base64');
    if (!signature.length) throw helloFailure('');
    return crypto.createHash('sha256').update(signature).digest();
  }

  return {
    available: async () => {
      if (process.platform !== 'win32') return false;
      if (known === true || (known === false && Date.now() - askedAt < RECHECK_MS)) return known;
      known = await run(`$r = Await ($mgr::IsSupportedAsync()) ([bool]); Write-Output $r`).then((o) => o.trim() === 'True', () => false);
      askedAt = Date.now();
      return known;
    },
    enroll: (id) => secretFrom(id, true),
    secret: (id) => secretFrom(id, false),
    forget: async (id) => {
      if (ID.test(String(id))) await run(forgetBody(id)).catch(() => {});
    },
  };
}

/**
 * A check run's stand-in: a fixed secret in place of the person's PIN. Honoured
 * only by a copy run from source. A check sets globalThis.__rutbaFakeHello to a
 * Windows status for the next call to answer that instead.
 */
export function fakeHello() {
  const secret = crypto.createHash('sha256').update('rutba-locker:fake-hello').digest();
  const answer = async () => {
    const next = globalThis.__rutbaFakeHello;
    globalThis.__rutbaFakeHello = null;
    if (next) throw helloFailure(next);
    return Buffer.from(secret);
  };
  return { available: async () => true, enroll: answer, secret: answer, forget: async () => {} };
}
