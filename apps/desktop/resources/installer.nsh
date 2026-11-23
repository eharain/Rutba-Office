; One copy of Rutba Office on a machine.
;
; electron-builder's installer removes the previous copy in the scope it is
; installing to - per-user, or all-users - and only that one. A copy in the
; other scope (an all-users install from an earlier day beside today's
; per-user one) stayed, and Windows listed two. This runs after the last
; page, once the person has said to go on, and before that step (at the very
; start of a silent install, which has no pages; Cancel on a page must leave
; the copy they have): it finds every registered copy, per-user and
; per-machine, in whichever directory it went to, and runs that copy's own
; uninstaller silently, keeping the documents and the settings. An entry
; whose uninstaller has gone (the folder deleted by hand) is forgotten
; instead of stopping the install with "Uninstall was not successful". A
; copy that cannot be removed stops the install: the one promise here is
; that it never ends with two.
;
; And a removal is all or nothing: a copy's program is never left in its
; folder without the rest (a husk, which every pin made from it starts and
; which ends at once), and an entry is forgotten only once its copy has
; gone. And the install keeps the kind of copy there is, all users or this
; person, and its folder, found however it was installed (rutbaWorkingScope),
; so the pins made from it keep working. The kind it replaces is removed
; first (rutbaRemoveCopies), so a removal that has to stop has not touched
; the copy kept, and a husk of that kind which cannot go is left, said in the
; log, rather than stop the install with no copy that works. A silent
; install (an update) never raises an administrator prompt of its own.
;
; Included ahead of the template, so the macros are defined before .onInit
; inserts customInit; the helpers used are NSIS's own, since the template's
; (installUtil.nsh) are included after .onInit.

!include "LogicLib.nsh"
!include "FileFunc.nsh"

; The company's name on every page of the installer and the uninstaller, in
; the branding line at the foot (owner, 14 Nov: "in installer kindly add
; company name tech style"). The template sets that line to the name and
; version alone, after this file is read; customHeader comes after it, so the
; line set here is the one used. The company is package.json's author, the
; legal name the copyright line uses, which also gives Add or Remove
; Programs its Publisher and the programs their CompanyName. No word of any
; language, since the pages speak the system's: a hyphen.
!macro customHeader
  !ifdef COMPANY_NAME
    BrandingText "${PRODUCT_NAME} ${VERSION} - ${COMPANY_NAME}"
  !endif
!macroend

; Whether DIR holds a copy that works: the program, and beside it
; resources\app.asar, the archive it runs. The program alone (a husk, left
; when a removal was cut short) starts and ends at once. OUT is "1" or "0".
!macro rutbaCopyWorks DIR OUT
  StrCpy ${OUT} "0"
  ${If} "${DIR}" != ""
    ${If} ${FileExists} "${DIR}\${APP_EXECUTABLE_FILENAME}"
    ${AndIf} ${FileExists} "${DIR}\resources\app.asar"
      StrCpy ${OUT} "1"
    ${EndIf}
  ${EndIf}
!macroend

; Whether two folders are the same: compared without a closing backslash,
; and each by its full name where it exists (a short 8.3 name, or one with
; a "..", names the same folder as its long one). OUT is "1" or "0". Uses $0
; and $1; OUT must not be either.
!macro rutbaSamePath A B OUT
  Push $0
  Push $1
  StrCpy $0 "${A}"
  StrCpy $1 "${B}"
  ${If} ${FileExists} "$0\*.*"
    GetFullPathName $0 "$0"
  ${EndIf}
  ${If} ${FileExists} "$1\*.*"
    GetFullPathName $1 "$1"
  ${EndIf}
  ${If} "$0" == ""
    StrCpy $0 "${A}"
  ${EndIf}
  ${If} "$1" == ""
    StrCpy $1 "${B}"
  ${EndIf}
  StrCpy ${OUT} "$0" "" -1
  ${If} ${OUT} == "\"
    StrCpy $0 "$0" -1
  ${EndIf}
  StrCpy ${OUT} "$1" "" -1
  ${If} ${OUT} == "\"
    StrCpy $1 "$1" -1
  ${EndIf}
  StrCpy ${OUT} "0"
  ${If} "$0" != ""
  ${AndIf} "$0" == "$1"
    StrCpy ${OUT} "1"
  ${EndIf}
  ClearErrors
  Pop $1
  Pop $0
!macroend

; Where each kind of copy is, found however it was installed (owner, 16 Nov,
; D-2026-11-16-03: "the installer should detect these changes and
; variations"). For each scope: the folder its InstallLocation names, as
; initMultiUser read it (KNOWN); else the folder its uninstall entry names,
; by InstallLocation or by where its uninstaller is (an entry whose install
; key was lost, in a custom folder too); else the scope's default folder,
; if our program is there with no registration at all (a husk an earlier
; removal left, or a copy whose entries went). "" when there is none.
; rutbaKept is the kind this install keeps ("all" or "user"), for the
; removal; rutbaOldVersion and rutbaOldPublisher what an old entry said,
; for putting it back.
!ifndef BUILD_UNINSTALLER
  Var rutbaMachineDir
  Var rutbaUserDir
  Var rutbaKept
  Var rutbaOldVersion
  Var rutbaOldPublisher
  Var rutbaCancelled
!endif

; The folder electron-builder installs a scope to by default: Program Files
; (the 64-bit one for a 64-bit build), or the person's own programs folder,
; found as setInstallModePerUser finds it (the UserProgramFiles known folder,
; else LOCALAPPDATA\Programs). RUTBA_MACHINE_FOLDER and RUTBA_USER_FOLDER,
; defined, stand in for them (tests/installer-nsis.test.js, which can make
; neither).
!macro rutbaDefaultFolder KIND OUT
  !if "${KIND}" == "machine"
    !ifdef RUTBA_MACHINE_FOLDER
      StrCpy ${OUT} "${RUTBA_MACHINE_FOLDER}"
    !else
      StrCpy ${OUT} "$PROGRAMFILES\${APP_FILENAME}"
      !ifdef APP_64 | APP_ARM64
        ${If} ${RunningX64}
          StrCpy ${OUT} "$PROGRAMFILES64\${APP_FILENAME}"
        ${EndIf}
      !endif
    !endif
  !else
    !ifdef RUTBA_USER_FOLDER
      StrCpy ${OUT} "${RUTBA_USER_FOLDER}"
    !else
      Push $0
      Push $1
      Push $2
      ReadEnvStr $0 LOCALAPPDATA
      StrCpy $0 "$0\Programs"
      StrCpy $2 0
      System::Call 'SHELL32::SHGetKnownFolderPath(g "{5CD7AEE2-2219-4A67-B85D-6C9CE15660CB}", i 0, p 0, *p .r2)i.r1'
      ${If} $1 == 0
        System::Call 'KERNEL32::lstrcpynW(w .r0, p r2, i ${NSIS_MAX_STRLEN})p'
      ${EndIf}
      ${If} $2 != 0
        System::Call 'OLE32::CoTaskMemFree(p r2)'
      ${EndIf}
      Pop $2
      Pop $1
      Exch $0
      Pop ${OUT}
      StrCpy ${OUT} "${OUT}\${APP_FILENAME}"
    !endif
  !endif
!macroend

; Uses $0 and $1.
!macro rutbaScopeFolder ROOT_KEY KNOWN KIND OUT
  StrCpy ${OUT} "${KNOWN}"
  ${If} ${OUT} == ""
    ReadRegStr ${OUT} ${ROOT_KEY} "${UNINSTALL_REGISTRY_KEY}" "InstallLocation"
  ${EndIf}
  ${If} ${OUT} == ""
    ReadRegStr $0 ${ROOT_KEY} "${UNINSTALL_REGISTRY_KEY}" "UninstallString"
    ${If} $0 != ""
      ClearErrors
      ${WordFind} "$0" '"' "+1" $1
      ${If} ${Errors}
        StrCpy $1 $0
      ${EndIf}
      ${GetParent} "$1" ${OUT}
    ${EndIf}
  ${EndIf}
  ${If} ${OUT} == ""
    !insertmacro rutbaDefaultFolder ${KIND} $1
    ${If} ${FileExists} "$1\${APP_EXECUTABLE_FILENAME}"
      StrCpy ${OUT} $1
    ${EndIf}
  ${EndIf}
  ClearErrors
!macroend

!macro rutbaFindCopies
  Push $0
  Push $1
  !insertmacro rutbaScopeFolder HKLM "$perMachineInstallationFolder" machine $rutbaMachineDir
  !insertmacro rutbaScopeFolder HKCU "$perUserInstallationFolder" user $rutbaUserDir
  Pop $1
  Pop $0
!macroend

; Where each copy found is, written where the template reads it. Both the
; page that asks who Office is for (on leaving it) and the elevated inner
; instance set the folder from the scope's install key, or the default
; folder when it has none: a copy whose install key was lost, in a custom
; folder, would be installed over in the default one and left where it was.
; So a scope whose key is empty, with our program found in a folder of its
; own, has the key written: this person's always (not in the inner instance,
; which may run as another account), all users' when elevated (the inner
; instance's .onInit, before its page). Only where our program is: an entry
; whose copy has gone names nothing to keep.
!macro rutbaRecordFolders
  !insertmacro rutbaFindCopies
  ${If} $perUserInstallationFolder == ""
  ${AndIf} $rutbaUserDir != ""
  ${AndIf} ${FileExists} "$rutbaUserDir\${APP_EXECUTABLE_FILENAME}"
  ${AndIfNot} ${UAC_IsInnerInstance}
    WriteRegStr HKCU "${INSTALL_REGISTRY_KEY}" "InstallLocation" "$rutbaUserDir"
    StrCpy $perUserInstallationFolder "$rutbaUserDir"
  ${EndIf}
  ${If} $perMachineInstallationFolder == ""
  ${AndIf} $rutbaMachineDir != ""
  ${AndIf} ${FileExists} "$rutbaMachineDir\${APP_EXECUTABLE_FILENAME}"
  ${AndIf} ${UAC_IsAdmin}
    WriteRegStr HKLM "${INSTALL_REGISTRY_KEY}" "InstallLocation" "$rutbaMachineDir"
    StrCpy $perMachineInstallationFolder "$rutbaMachineDir"
  ${EndIf}
  ClearErrors
!macroend

; Keep the kind of copy there is. electron-builder sees a scope only by its
; InstallLocation, and with both it goes per-user: on the owner's PC (2.26.0)
; that took Office out of Program Files, where 2.25.0 was, from under every
; taskbar pin made from it. So, with the folders rutbaFindCopies finds:
; - one kind works: that kind, whatever else is registered;
; - both work: all users, as a silent update already went: removing that
;   copy would take Office away from every other person on this computer,
;   and the per-user copy is this person's alone;
; - neither works: the kind whose folder still holds our program (a husk,
;   which an install into the same folder makes whole again, and the pins
;   made from it with it), the one Windows still lists first; both, all
;   users, as above;
; - nothing anywhere: OUT is "", and electron-builder decides, as for a
;   first install.
; OUT must not be $0 to $5.
!macro rutbaWorkingScope OUT
  !insertmacro rutbaFindCopies
  Push $0
  Push $1
  Push $2
  Push $3
  Push $4
  Push $5
  !insertmacro rutbaCopyWorks "$rutbaMachineDir" $0
  !insertmacro rutbaCopyWorks "$rutbaUserDir" $1
  StrCpy $2 "0"
  StrCpy $3 "0"
  ${If} $rutbaMachineDir != ""
  ${AndIf} ${FileExists} "$rutbaMachineDir\${APP_EXECUTABLE_FILENAME}"
    StrCpy $2 "1"
  ${EndIf}
  ${If} $rutbaUserDir != ""
  ${AndIf} ${FileExists} "$rutbaUserDir\${APP_EXECUTABLE_FILENAME}"
    StrCpy $3 "1"
  ${EndIf}
  ; Registered: an install key or an uninstall entry for the scope.
  StrCpy $4 "$perMachineInstallationFolder"
  ${If} $4 == ""
    ReadRegStr $4 HKLM "${UNINSTALL_REGISTRY_KEY}" "UninstallString"
  ${EndIf}
  StrCpy $5 "$perUserInstallationFolder"
  ${If} $5 == ""
    ReadRegStr $5 HKCU "${UNINSTALL_REGISTRY_KEY}" "UninstallString"
  ${EndIf}
  ${If} $0 == "1"
    StrCpy ${OUT} "all"
  ${ElseIf} $1 == "1"
    StrCpy ${OUT} "user"
  ${ElseIf} $2 == "1"
  ${AndIf} $4 != ""
    StrCpy ${OUT} "all"
  ${ElseIf} $3 == "1"
  ${AndIf} $5 != ""
    StrCpy ${OUT} "user"
  ${ElseIf} $2 == "1"
    StrCpy ${OUT} "all"
  ${ElseIf} $3 == "1"
    StrCpy ${OUT} "user"
  ${Else}
    StrCpy ${OUT} ""
  ${EndIf}
  ClearErrors
  Pop $5
  Pop $4
  Pop $3
  Pop $2
  Pop $1
  Pop $0
!macroend

; And keep its folder: when the scope's InstallLocation was lost (KNOWN is
; empty) but its copy was found elsewhere, the install goes back there, not
; to the default folder, unless /D= says where. NSIS takes /D= off the
; command line it gives the installer, so it is read as the template reads
; it (GetDParameter). Uses $R0 and $R1.
!macro rutbaKeepFolder KNOWN FOUND
  ${If} "${KNOWN}" == ""
  ${AndIf} "${FOUND}" != ""
    !insertmacro GetDParameter $R1
    ${If} $R1 == ""
      StrCpy $INSTDIR "${FOUND}"
    ${EndIf}
    ClearErrors
  ${EndIf}
!macroend

; The page that asks who Office is for opens on the kind of copy there is
; (rutbaWorkingScope), in its folder, rather than on per-user. It is a
; choice the page still shows, not one forced: forcing all users skips the
; page and goes straight to the administrator prompt, and a No there ends
; the installer without a word, where on the page it leaves the person free
; to pick per-user instead. /allusers and /currentuser still decide, after
; this. The template runs this before the page, and only in the outer
; instance (the elevated inner one is always all users). The uninstaller's
; page asks which copy to remove, and is left alone.
!macro customInstallMode
  !ifndef BUILD_UNINSTALLER
    Push $R9
    Push $R0
    Push $R1
    Push $0
    !insertmacro rutbaWorkingScope $R9
    ${If} $R9 == "all"
      !insertmacro setInstallModePerAllUsers
      !insertmacro rutbaKeepFolder "$perMachineInstallationFolder" "$rutbaMachineDir"
    ${ElseIf} $R9 == "user"
      !insertmacro setInstallModePerUser
      !insertmacro rutbaKeepFolder "$perUserInstallationFolder" "$rutbaUserDir"
    ${EndIf}
    Pop $0
    Pop $R1
    Pop $R0
    Pop $R9
  !endif
!macroend

; The same rule for a silent install (an update has no pages), unless
; /allusers or /currentuser says which. It runs after initMultiUser, which
; sees a scope only by its InstallLocation, and before the install section,
; which elevates and installs for all users when $hasPerMachineInstallation
; is 1. So all users sets it; per-user clears it, so the section does not
; elevate, and sets the per-user mode and folder outright. An elevated
; install sets the all-users mode and folder here too, so the removal knows
; where this install goes. The elevated inner instance is started only
; because the outer one chose all users: it is all users, always, so an
; elevated per-user install, a second elevation or a loop cannot follow.
; And a silent update never asks for an elevation it does not need: when all
; that is left for all users is a husk (no copy works anywhere), an install
; that is not already elevated goes per-user and leaves the husk, said in
; the log, rather than raise an administrator prompt behind an app that has
; just closed for its update.
!macro rutbaSilentScope
  ${If} ${Silent}
    Push $0
    ${If} ${UAC_IsInnerInstance}
      StrCpy $hasPerMachineInstallation "1"
      !insertmacro setInstallModePerAllUsers
    ${Else}
      ${GetParameters} $R0
      ClearErrors
      ${GetOptions} $R0 "/allusers" $R1
      ${If} ${Errors}
        ClearErrors
        ${GetOptions} $R0 "/currentuser" $R1
        ${If} ${Errors}
          !insertmacro rutbaWorkingScope $R1
          ${If} $R1 == "all"
            !insertmacro rutbaCopyWorks "$rutbaMachineDir" $R1
            ${If} $R1 == "1"
            ${OrIf} ${UAC_IsAdmin}
              StrCpy $hasPerMachineInstallation "1"
              ${If} ${UAC_IsAdmin}
                !insertmacro setInstallModePerAllUsers
                !insertmacro rutbaKeepFolder "$perMachineInstallationFolder" "$rutbaMachineDir"
              ${EndIf}
            ${Else}
              DetailPrint "Only the program is left of the Rutba Office for all users, in $rutbaMachineDir; removing it needs an administrator, so it stays, and this install is for this person."
              StrCpy $R1 "user"
            ${EndIf}
          ${EndIf}
          ${If} $R1 == "user"
            StrCpy $hasPerMachineInstallation "0"
            !insertmacro setInstallModePerUser
            !insertmacro rutbaKeepFolder "$perUserInstallationFolder" "$rutbaUserDir"
          ${EndIf}
        ${EndIf}
      ${EndIf}
    ${EndIf}
    ClearErrors
    Pop $0
  ${EndIf}
!macroend

; What a Rutba Office install puts in its folder, by name: the program, its
; uninstaller and that one's icon, the files Electron and the build put
; beside them (two that older Electrons had among them), and, apart, its two
; folders. Nothing else goes with them: a file of anyone else's stays, and
; so does the folder that holds it. MODE delete removes each in turn; MODE
; list adds each, quoted and by its name alone, to $R3, for the del of an
; elevated Command Prompt already in the folder. Only ever used on a folder
; that holds our program, and checked against the build's own folder by
; tests/installer-nsis.test.js.
!macro rutbaOurFile MODE DIR NAME
  !if "${MODE}" == "delete"
    Delete "${DIR}\${NAME}"
  !else
    StrCpy $R3 '$R3 "${NAME}"'
  !endif
!macroend

!macro rutbaOurFiles MODE DIR
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "${APP_EXECUTABLE_FILENAME}"
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "${UNINSTALL_FILENAME}"
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "uninstallerIcon.ico"
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "chrome_100_percent.pak"
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "chrome_200_percent.pak"
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "resources.pak"
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "d3dcompiler_47.dll"
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "dxcompiler.dll"
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "dxil.dll"
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "ffmpeg.dll"
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "libEGL.dll"
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "libGLESv2.dll"
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "vk_swiftshader.dll"
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "vk_swiftshader_icd.json"
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "vulkan-1.dll"
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "icudtl.dat"
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "snapshot_blob.bin"
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "v8_context_snapshot.bin"
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "LICENSE.electron.txt"
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "LICENSE.txt"
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "LICENSES.chromium.html"
  !insertmacro rutbaOurFile ${MODE} "${DIR}" "THIRD-PARTY-NOTICES.txt"
!macroend

; Whether DIR may have its resources and locales folders removed whole: only
; a folder named as ours (the directory page puts our name on the end of any
; folder chosen), so never a drive's root or a folder such as D:\Tools that
; an entry names, where those two names could be anyone's. OUT is "1" or
; "0"; it must not be $0.
!macro rutbaOurFolder DIR OUT
  Push $0
  StrCpy $0 "${DIR}"
  StrCpy ${OUT} "$0" "" -1
  ${If} ${OUT} == "\"
    StrCpy $0 "$0" -1
  ${EndIf}
  ${GetFileName} "$0" $0
  StrCpy ${OUT} "0"
  ${If} "$0" == "${APP_FILENAME}"
    StrCpy ${OUT} "1"
  ${EndIf}
  Pop $0
!macroend

; The processes started from DIR's program, and nothing else: its windows,
; and the helpers that run the same program as Node (the AI tools' helpers
; among them), which can hold the file open after its uninstaller has
; finished. By the program's full path, not its name: a name would take in a
; Rutba Office in the other scope too, and anyone else's on this computer.
; PowerShell is given the path in the environment, not on its command line,
; so no folder name can break the quoting. MODE close asks each to close, as
; its X would (it saves its recovery copies on the way out); end ends them
; and what they started; count sets OUT to how many there are, windows to
; how many have a window up ("error" when PowerShell could not run). Uses $0.
!macro rutbaProcessesIn DIR MODE OUT
  Push $0
  StrCpy $0 "${DIR}\${APP_EXECUTABLE_FILENAME}"
  ${If} ${FileExists} "$0"
    GetFullPathName $0 "$0"
  ${EndIf}
  ${If} $0 == ""
    StrCpy $0 "${DIR}\${APP_EXECUTABLE_FILENAME}"
  ${EndIf}
  ClearErrors
  System::Call 'kernel32::SetEnvironmentVariable(t "RUTBA_LEFTOVER_EXE", t r0) i'
  !if "${MODE}" == "close"
    nsExec::Exec '"$SYSDIR\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -Command "$$p = [IO.Path]::GetFullPath($$env:RUTBA_LEFTOVER_EXE); Get-Process -ErrorAction SilentlyContinue | Where-Object { $$_.Path -and ([IO.Path]::GetFullPath($$_.Path) -eq $$p) } | ForEach-Object { & $$env:SystemRoot\System32\taskkill.exe /T /PID $$_.Id 2>&1 | Out-Null }; exit 0"'
  !else if "${MODE}" == "end"
    nsExec::Exec '"$SYSDIR\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -Command "$$p = [IO.Path]::GetFullPath($$env:RUTBA_LEFTOVER_EXE); Get-Process -ErrorAction SilentlyContinue | Where-Object { $$_.Path -and ([IO.Path]::GetFullPath($$_.Path) -eq $$p) } | ForEach-Object { & $$env:SystemRoot\System32\taskkill.exe /F /T /PID $$_.Id 2>&1 | Out-Null }; exit 0"'
  !else if "${MODE}" == "windows"
    nsExec::Exec '"$SYSDIR\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -Command "$$p = [IO.Path]::GetFullPath($$env:RUTBA_LEFTOVER_EXE); exit [Math]::Min(99, @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $$_.Path -and ([IO.Path]::GetFullPath($$_.Path) -eq $$p) -and $$_.MainWindowTitle }).Count)"'
  !else
    nsExec::Exec '"$SYSDIR\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -Command "$$p = [IO.Path]::GetFullPath($$env:RUTBA_LEFTOVER_EXE); exit [Math]::Min(99, @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $$_.Path -and ([IO.Path]::GetFullPath($$_.Path) -eq $$p) }).Count)"'
  !endif
  Pop $0
  System::Call 'kernel32::SetEnvironmentVariable(t "RUTBA_LEFTOVER_EXE", p 0) i'
  Exch $0
  Pop ${OUT}
!macroend

; Frees DIR's program for removal when its uninstaller left it held open.
; What runs from it is given the same grace as at the start of the removal:
; five seconds to end by itself, then asked to close, five more, then ended;
; a person here is told first. A window still up then may be asking them
; whether to save, so it is never ended: a person is shown the same Retry
; box as the start of the removal (Retry waits again; Cancel leaves it, and
; sets $rutbaCancelled to "1" so the caller shows no second box), and a
; silent install (whose update closed its own app first, so this is a copy
; someone has open) leaves it and says so in the log. Left, nothing of that
; copy is touched, $R3 is "window", and the caller's rules apply: a whole
; copy stops the install, so the one that works is kept, and a husk of the
; kind replaced is kept while the install goes on (stopping there would
; leave no copy that works at all). Otherwise, with no window up, what runs
; from it is ended and the program deleted, three times over a few seconds,
; for a helper that started again; and last, if it is still held, moved out
; of the folder into the temporary folder (a running program's file can be
; renamed), which frees the folder and every pin's path at once. An
; all-users folder this installer is not elevated for is another matter:
; nothing there can be deleted or moved, so nothing of the person's is ended
; for it; only the move is tried. Uses $R3.
!macro rutbaFreeProgram SCOPE DIR
  StrCpy $rutbaCancelled ""
  !if "${SCOPE}" == "allusers"
    ${If} ${UAC_IsAdmin}
  !endif
  !insertmacro rutbaProcessesIn "${DIR}" count $R3
  ${If} $R3 != "0"
  ${AndIf} $R3 != "error"
    ${IfNot} ${Silent}
      MessageBox MB_OK|MB_ICONINFORMATION "Rutba Office is still running from ${DIR}, which this install replaces. Save anything open in it and close it: it is asked to close in a few seconds." /SD IDOK
    ${EndIf}
    Push $0
    StrCpy $0 0
    ${Do}
      ${If} $0 == 5
        DetailPrint "Asking what runs from ${DIR} to close"
        !insertmacro rutbaProcessesIn "${DIR}" close $R3
      ${EndIf}
      Sleep 1000
      IntOp $0 $0 + 1
      !insertmacro rutbaProcessesIn "${DIR}" count $R3
      ${If} $R3 == "0"
        ${ExitDo}
      ${EndIf}
      ${If} $0 >= 10
        ; Only what has no window up may be ended.
        !insertmacro rutbaProcessesIn "${DIR}" windows $R3
        ${If} $R3 == "0"
        ${OrIf} $R3 == "error"
          ${ExitDo}
        ${EndIf}
        ${If} ${Silent}
          DetailPrint "Rutba Office is open from ${DIR}, with a window up; it is left as it is."
          StrCpy $R3 "window"
          ${ExitDo}
        ${EndIf}
        !insertmacro rutbaCopyWorks "${DIR}" $R3
        ${If} $R3 == "1"
          StrCpy $R3 "Cancel stops the install and leaves the copy you have."
        ${Else}
          StrCpy $R3 "Only its program is left there. Cancel leaves it open as it is, and the install goes on without it where it can; where it cannot, it stops."
        ${EndIf}
        ${If} ${Cmd} `MessageBox MB_RETRYCANCEL|MB_ICONEXCLAMATION "Rutba Office (in ${DIR}) is still open, so it cannot be replaced yet.$\r$\n$\r$\nClose every Rutba Office window and press Retry. $R3" /SD IDCANCEL IDRETRY`
          StrCpy $0 0
        ${Else}
          StrCpy $rutbaCancelled "1"
          StrCpy $R3 "window"
          ${ExitDo}
        ${EndIf}
      ${EndIf}
    ${Loop}
    Pop $0
  ${EndIf}
  ${If} $R3 != "window"
    Push $0
    StrCpy $0 0
    ${Do}
      IntOp $0 $0 + 1
      !insertmacro rutbaProcessesIn "${DIR}" end $R3
      Sleep 1000
      ClearErrors
      Delete "${DIR}\${APP_EXECUTABLE_FILENAME}"
      ${IfNot} ${FileExists} "${DIR}\${APP_EXECUTABLE_FILENAME}"
      ${OrIf} $0 >= 3
        ${ExitDo}
      ${EndIf}
    ${Loop}
    Pop $0
  ${EndIf}
  !if "${SCOPE}" == "allusers"
    ${Else}
      StrCpy $R3 ""
    ${EndIf}
  !endif
  ${If} $R3 != "window"
  ${AndIf} ${FileExists} "${DIR}\${APP_EXECUTABLE_FILENAME}"
    System::Call 'kernel32::GetTickCount() i .R3'
    ClearErrors
    Rename "${DIR}\${APP_EXECUTABLE_FILENAME}" "$TEMP\rutba-office-left-$R3.exe"
    ${IfNot} ${Errors}
      DetailPrint "The program in ${DIR} was still held open; it was moved to $TEMP\rutba-office-left-$R3.exe"
    ${EndIf}
  ${EndIf}
  ClearErrors
!macroend

; The registry view reg.exe must use to see the keys this installer writes
; (check64BitAndSetRegView: the 64-bit one for a 64-bit build on 64-bit Windows).
!macro rutbaRegViewFlag OUT
  StrCpy ${OUT} "/reg:32"
  !ifdef APP_64 | APP_ARM64
    ${If} ${RunningX64}
      StrCpy ${OUT} "/reg:64"
    ${EndIf}
  !endif
!macroend

; Removes what is left of a copy in DIR, all of it or nothing: the program
; goes first (freed by rutbaFreeProgram when it is held open), and only once
; it has gone does anything else, so the program is never left without the
; rest (a husk). Its resources and locales go whole only in a folder named
; as ours (rutbaOurFolder); elsewhere only what is ours in them. A copy for
; all users is in Program Files, which this installer cannot change unless
; it is elevated. A silent install then leaves it, and says so in the log:
; an update raises no prompt that nobody is there to answer. With a person
; here, it asks once whether to remove it with an administrator's yes, No
; the default; Yes runs the same removal in a Command Prompt as
; administrator, which also runs $R2 (registry commands of the caller's, ""
; for none) once the program has gone, so one prompt does it all, and
; clears $R2. Whether anything was left is the caller's to judge, by the
; program. Uses $R3.
!macro rutbaClearLeftovers SCOPE DIR
  DetailPrint "Removing what is left of the Rutba Office installed for ${SCOPE}, in ${DIR}"
  ClearErrors
  Delete "${DIR}\${APP_EXECUTABLE_FILENAME}"
  ${If} ${FileExists} "${DIR}\${APP_EXECUTABLE_FILENAME}"
    !insertmacro rutbaFreeProgram ${SCOPE} "${DIR}"
  ${EndIf}
  ${IfNot} ${FileExists} "${DIR}\${APP_EXECUTABLE_FILENAME}"
    !insertmacro rutbaOurFiles delete "${DIR}"
    !insertmacro rutbaOurFolder "${DIR}" $R3
    ${If} $R3 == "1"
      RMDir /r "${DIR}\resources"
      RMDir /r "${DIR}\locales"
    ${Else}
      Delete "${DIR}\resources\app.asar"
      RMDir /r "${DIR}\resources\app.asar.unpacked"
      Delete "${DIR}\resources\app-update.yml"
      Delete "${DIR}\resources\elevate.exe"
      Delete "${DIR}\locales\*.pak"
      RMDir "${DIR}\resources"
      RMDir "${DIR}\locales"
    ${EndIf}
    RMDir "${DIR}"
  ${EndIf}
  !if "${SCOPE}" == "allusers"
    ${If} ${FileExists} "${DIR}\${APP_EXECUTABLE_FILENAME}"
    ${AndIfNot} ${UAC_IsAdmin}
      ${If} ${Silent}
        DetailPrint "What is left of the Rutba Office for all users, in ${DIR}, needs an administrator to remove; this install is not elevated, so it stays."
      ${Else}
        ; In the folder, by name: the whole command stays well inside NSIS's string length.
        !insertmacro rutbaOurFolder "${DIR}" $R3
        ${If} $R3 == "1"
          StrCpy $R3 "del /f /q"
          !insertmacro rutbaOurFiles list "${DIR}"
          StrCpy $R3 '$R3 & rmdir /s /q "resources" & rmdir /s /q "locales"'
        ${Else}
          StrCpy $R3 "del /f /q"
          !insertmacro rutbaOurFiles list "${DIR}"
          StrCpy $R3 '$R3 "resources\app.asar" "resources\app-update.yml" "resources\elevate.exe" "locales\*.pak" & rmdir /s /q "resources\app.asar.unpacked" & rmdir "resources" & rmdir "locales"'
        ${EndIf}
        StrCpy $R3 '$R3 & cd /d "$SYSDIR" & rmdir "${DIR}"'
        ${If} $R2 != ""
          StrCpy $R3 '$R3 & $R2'
        ${EndIf}
        ; The rest only once the program has gone; nothing at all if the folder is not there.
        StrCpy $R3 'cd /d "${DIR}" && (del /f /q "${APP_EXECUTABLE_FILENAME}" & if not exist "${APP_EXECUTABLE_FILENAME}" ($R3))'
        ${If} ${Cmd} `MessageBox MB_YESNO|MB_ICONQUESTION|MB_DEFBUTTON2 "What is left of the Rutba Office installed for all users, in ${DIR}, can only be removed with an administrator's yes. Remove it now?$\r$\n$\r$\nWindows asks for that yes next." IDYES`
          StrCpy $R2 ""
          ClearErrors
          ExecShellWait "runas" "$SYSDIR\cmd.exe" '/d /c "$R3"' SW_HIDE
          Sleep 500
        ${EndIf}
      ${EndIf}
    ${EndIf}
  !endif
  ClearErrors
!macroend

; One copy's removal, by its uninstall entry. ROLE says whose it is: own,
; the kind this install keeps (it is installed over again), or other, the
; kind it replaces. The other kind is removed first (rutbaRemoveCopies), so
; a removal that stops has not touched the copy kept; and what is left of
; the other kind when it is only a husk does not stop the install (it is
; said in the log, and the app's pin repair points the pins away from it):
; stopping would leave no copy that works at all. A whole copy that cannot
; be removed stops it, as ever, rather than leave two.
!macro rutbaRemoveInstalledCopy ROOT_KEY SCOPE ROLE
  ClearErrors
  StrCpy $R9 ""
  StrCpy $rutbaCancelled ""
  ReadRegStr $R0 ${ROOT_KEY} "${UNINSTALL_REGISTRY_KEY}" "UninstallString"
  ${If} $R0 != ""
    ; The uninstaller's path: what is between the quotes, or the whole
    ; string when there are none.
    StrCpy $R1 $R0 1
    ${If} $R1 == '"'
      StrCpy $R1 $R0 "" 1
      StrCpy $R2 0
      ${Do}
        StrCpy $R3 $R1 1 $R2
        ${If} $R3 == ""
        ${OrIf} $R3 == '"'
          ${ExitDo}
        ${EndIf}
        IntOp $R2 $R2 + 1
      ${Loop}
      StrCpy $R1 $R1 $R2
    ${Else}
      StrCpy $R1 $R0
    ${EndIf}

    ${GetParent} "$R1" $R4
    ; An uninstaller for all users that is not elevated asks Windows for an
    ; administrator itself: a silent install leaves that copy instead (no
    ; prompt behind an update), stopping only for one that works.
    StrCpy $R6 "run"
    !if "${SCOPE}" == "allusers"
      ${If} ${Silent}
      ${AndIfNot} ${UAC_IsAdmin}
      ${AndIf} ${FileExists} "$R1"
        StrCpy $R6 "leave"
        !insertmacro rutbaCopyWorks "$R4" $R3
        ${If} $R3 == "1"
          DetailPrint "The Rutba Office installed for all users, in $R4, needs an administrator to remove; this install is not elevated, so it has stopped rather than leave two copies."
          SetErrorLevel 2
          Quit
        ${EndIf}
        DetailPrint "What is left of the Rutba Office for all users, in $R4, needs an administrator to remove; this install is not elevated, so it stays."
        StrCpy $R9 "keep"
      ${EndIf}
    !endif

    ${If} ${FileExists} "$R1"
    ${AndIf} $R6 == "run"
      ; Room for the new copy before the old one goes: a full disk must leave the one that works.
      ${GetRoot} "$R4" $R3
      ${DriveSpace} "$R3\" "/D=F /S=M" $R3
      ${If} $R3 != ""
      ${AndIf} $R3 < 1024
        ${IfNot} ${Silent}
          MessageBox MB_OK|MB_ICONEXCLAMATION "There is not enough free disk space to replace Rutba Office. Free at least 1 GB and run this again. The copy you have was left as it is."
        ${EndIf}
        SetErrorLevel 2
        Quit
      ${EndIf}
      ; Run from a copy: the uninstaller removes the directory it lives in,
      ; and _?= keeps it in place so ExecWait actually waits for it.
      InitPluginsDir
      StrCpy $R5 "$PLUGINSDIR\rutba-old-uninstaller-${SCOPE}.exe"
      CopyFiles /SILENT "$R1" "$R5"
      ${IfNot} ${FileExists} "$R5"
        StrCpy $R5 "$R1"
      ${EndIf}
      ; What the entry says, to put back as it was if the copy cannot go.
      ReadRegStr $rutbaOldVersion ${ROOT_KEY} "${UNINSTALL_REGISTRY_KEY}" "DisplayVersion"
      ReadRegStr $rutbaOldPublisher ${ROOT_KEY} "${UNINSTALL_REGISTRY_KEY}" "Publisher"
      !if "${SCOPE}" == "allusers"
        ; The old uninstaller deletes the key that says where it was installed; an update goes back to the same place.
        ReadRegStr $R8 ${ROOT_KEY} "${INSTALL_REGISTRY_KEY}" "InstallLocation"
      !endif
      DetailPrint "Removing the Rutba Office installed for ${SCOPE} from $R4"
      ; Tried again while the copy is still closing: an update starts this
      ; installer the moment Rutba Office begins to quit, and the old
      ; uninstaller refuses (error 2) until its last window and thread are
      ; gone. Once a second for 15 seconds, as the template's own removal
      ; does, then the person is asked rather than the install stopped.
      StrCpy $R7 0
      ${Do}
        IntOp $R7 $R7 + 1
        ClearErrors
        ExecWait '"$R5" /S /KEEP_APP_DATA /${SCOPE} --updated _?=$R4' $R6
        ${IfNot} ${Errors}
        ${AndIf} $R6 == 0
          ${ExitDo}
        ${EndIf}
        ; Still running after five seconds: ask its windows to close, as the
        ; X would (it saves its recovery copies on the way out). After ten:
        ; end what is left, background processes too, unless a person is here
        ; and a window is still up, which may be asking them whether to save:
        ; then it is theirs to answer, and the Retry box below says so. A
        ; silent install has nobody to ask; unsaved work is copied for Recover
        ; every half minute, so at worst under that much of it. Only what
        ; runs from this copy's program, by its path.
        ${If} $R7 == 5
          DetailPrint "Asking Rutba Office to close"
          !insertmacro rutbaProcessesIn "$R4" close $R6
        ${ElseIf} $R7 == 10
          StrCpy $R6 "0"
          ${IfNot} ${Silent}
            !insertmacro rutbaProcessesIn "$R4" windows $R6
          ${EndIf}
          ${If} $R6 == "0"
          ${OrIf} $R6 == "error"
            DetailPrint "Closing what is left of Rutba Office"
            !insertmacro rutbaProcessesIn "$R4" end $R6
          ${EndIf}
        ${EndIf}
        ${If} $R7 >= 15
          MessageBox MB_RETRYCANCEL|MB_ICONEXCLAMATION "Rutba Office (installed for ${SCOPE}, in $R4) is still open, so it cannot be replaced yet.$\r$\n$\r$\nClose every Rutba Office window and press Retry. Cancel stops the install and leaves the copy you have. (Uninstaller error $R6.)" /SD IDCANCEL IDRETRY +2
          Quit
          StrCpy $R7 0
        ${EndIf}
        Sleep 1000
      ${Loop}
      ; An uninstaller that was refused its elevation says nothing and leaves everything.
      Sleep 1500
      ; One that finished (its entry is gone, which it removes last) but left
      ; the program had it held open, by a window still closing or by a
      ; helper run from it as Node: it is freed and the rest removed, or the
      ; program stays alone in its folder, a husk that every pin made from
      ; this copy starts and that ends at once.
      ReadRegStr $R6 ${ROOT_KEY} "${UNINSTALL_REGISTRY_KEY}" "UninstallString"
      ${If} ${FileExists} "$R4\${APP_EXECUTABLE_FILENAME}"
      ${AndIf} $R6 == ""
        StrCpy $R2 ""
        !insertmacro rutbaClearLeftovers ${SCOPE} "$R4"
      ${EndIf}
      ${If} ${FileExists} "$R4\${APP_EXECUTABLE_FILENAME}"
        !insertmacro rutbaCopyWorks "$R4" $R3
        !if "${ROLE}" == "other"
          ${If} $R3 == "0"
            ; Only the program is left of the kind replaced: the install goes on.
            DetailPrint "Only the program is left of the Rutba Office for ${SCOPE}, in $R4; it could not be removed, and stays."
            StrCpy $R9 "keep"
          ${EndIf}
        !endif
        ${If} $R9 != "keep"
          ; Still there: if its uninstaller took the entry, it is put back as
          ; it was, so the copy left is one Windows still lists and the next
          ; run finds it (with its uninstaller gone, below) and finishes the
          ; removal; for all users, where it is installed too, which
          ; initMultiUser reads. This scope's keys are this install's to
          ; write: the other kind's copy is never left so, only a husk of it.
          ${If} $R6 == ""
            WriteRegStr ${ROOT_KEY} "${UNINSTALL_REGISTRY_KEY}" "UninstallString" "$R0"
            WriteRegStr ${ROOT_KEY} "${UNINSTALL_REGISTRY_KEY}" "InstallLocation" "$R4"
            WriteRegStr ${ROOT_KEY} "${UNINSTALL_REGISTRY_KEY}" "DisplayName" "${PRODUCT_NAME}"
            ${If} $rutbaOldVersion != ""
              WriteRegStr ${ROOT_KEY} "${UNINSTALL_REGISTRY_KEY}" "DisplayVersion" "$rutbaOldVersion"
            ${EndIf}
            ${If} $rutbaOldPublisher != ""
              WriteRegStr ${ROOT_KEY} "${UNINSTALL_REGISTRY_KEY}" "Publisher" "$rutbaOldPublisher"
            ${EndIf}
            !if "${SCOPE}" == "allusers"
              ${If} $R8 != ""
                WriteRegStr ${ROOT_KEY} "${INSTALL_REGISTRY_KEY}" "InstallLocation" "$R8"
              ${EndIf}
            !endif
          ${EndIf}
          ${IfNot} ${Silent}
          ${AndIf} $rutbaCancelled != "1"
            MessageBox MB_OK|MB_ICONEXCLAMATION "The Rutba Office installed for ${SCOPE}, in $R4, was not removed, so this install has stopped rather than leave two copies. Run it again and allow the administrator prompt, or remove that copy first."
          ${EndIf}
          SetErrorLevel 2
          Quit
        ${EndIf}
      ${Else}
        ; The uninstaller run from a copy cannot delete the original of itself.
        Delete "$R1"
        RMDir "$R4"
        !if "${SCOPE}" == "allusers"
          ${If} ${Silent}
            ; Where it was by its install key, else where its uninstaller was:
            ; a custom folder whose install key was lost is kept too.
            ${If} $R8 == ""
              StrCpy $R8 $R4
            ${EndIf}
            WriteRegStr ${ROOT_KEY} "${INSTALL_REGISTRY_KEY}" "InstallLocation" "$R8"
          ${EndIf}
        !endif
      ${EndIf}
    ${ElseIf} $R6 == "run"
      ; The uninstaller has gone, but the copy may not have: one run's old
      ; uninstaller can delete everything but a program that was held open,
      ; and forgetting the entry then left that program alone in its folder
      ; for good (the owner's PC, 2.26.0). So what is left there is removed
      ; first, and the entry forgotten only once it has gone. The folder is
      ; the entry's InstallLocation, else where its uninstaller was (2.18.0
      ; left InstallLocation empty), and it is touched only when it still
      ; holds our program: nothing that is not our install folder.
      ReadRegStr $R4 ${ROOT_KEY} "${UNINSTALL_REGISTRY_KEY}" "InstallLocation"
      ${If} $R4 == ""
        ${GetParent} "$R1" $R4
      ${EndIf}
      ; Where this scope says Office is installed is forgotten with it, when
      ; no working copy is left there; otherwise initMultiUser goes on
      ; seeing this scope registered (2.18.0's stale per-user copy did).
      ; Kept for a silent update for all users: its install section reads
      ; it to go back to the same folder, and the pins made from it.
      ReadRegStr $R5 ${ROOT_KEY} "${INSTALL_REGISTRY_KEY}" "InstallLocation"
      StrCpy $R7 "0"
      ${If} $R5 != ""
        !insertmacro rutbaCopyWorks "$R5" $R6
        !insertmacro rutbaSamePath "$R5" "$R4" $R3
        ${If} $R6 == "0"
        ${OrIf} $R3 == "1"
          StrCpy $R7 "1"
        ${EndIf}
        !if "${SCOPE}" == "allusers"
          ${If} ${Silent}
          ${AndIf} $hasPerMachineInstallation == "1"
            StrCpy $R7 "0"
          ${EndIf}
        !endif
      ${EndIf}
      ; Not elevated, an entry for all users is not this installer's to
      ; change: the commands that forget it go to the administrator's
      ; prompt, after the files, or alone below when no files are left.
      StrCpy $R2 ""
      !if "${SCOPE}" == "allusers"
        ${IfNot} ${UAC_IsAdmin}
          !insertmacro rutbaRegViewFlag $R3
          StrCpy $R2 '"$SYSDIR\reg.exe" delete "HKLM\${UNINSTALL_REGISTRY_KEY}" /f $R3'
          ${If} $R7 == "1"
            StrCpy $R2 '$R2 & "$SYSDIR\reg.exe" delete "HKLM\${INSTALL_REGISTRY_KEY}" /v InstallLocation /f $R3'
          ${EndIf}
        ${EndIf}
      !endif
      StrCpy $R8 "0"
      ${If} $R4 != ""
      ${AndIf} ${FileExists} "$R4\${APP_EXECUTABLE_FILENAME}"
        StrCpy $R8 "1"
        !insertmacro rutbaCopyWorks "$R4" $R6
        !insertmacro rutbaClearLeftovers ${SCOPE} "$R4"
        ; The entry stays, true, while its program does.
        ${If} ${FileExists} "$R4\${APP_EXECUTABLE_FILENAME}"
          StrCpy $R9 "keep"
          ${If} $R6 == "1"
          !if "${ROLE}" == "own"
          ${OrIf} $R6 == "0"
          !endif
            ${IfNot} ${Silent}
            ${AndIf} $rutbaCancelled != "1"
              MessageBox MB_OK|MB_ICONEXCLAMATION "What is left of the Rutba Office installed for ${SCOPE}, in $R4, was not removed, so this install has stopped rather than leave it there. Run it again and allow the administrator prompt, or remove that folder first."
            ${EndIf}
            SetErrorLevel 2
            Quit
          ${EndIf}
          DetailPrint "Only the program is left of the Rutba Office for ${SCOPE}, in $R4; it could not be removed, and stays, with its entry."
        ${EndIf}
      ${EndIf}
      ${If} $R9 != "keep"
        DetailPrint "Forgetting a Rutba Office entry (${SCOPE}) whose uninstaller is gone"
        ${If} $R7 == "1"
          DeleteRegValue ${ROOT_KEY} "${INSTALL_REGISTRY_KEY}" "InstallLocation"
        ${EndIf}
        ; A silent update for all users goes back to the folder the copy was in,
        ; when the install key that said so was lost: a custom one is kept.
        !if "${SCOPE}" == "allusers"
          ${If} ${Silent}
          ${AndIf} $hasPerMachineInstallation == "1"
          ${AndIf} $R5 == ""
          ${AndIf} $R8 == "1"
            WriteRegStr ${ROOT_KEY} "${INSTALL_REGISTRY_KEY}" "InstallLocation" "$R4"
          ${EndIf}
          ; Nothing left but the entry, which needs an administrator: asked
          ; for once when a person is here, No leaving it; never when silent.
          ${If} $R2 != ""
            ${If} ${Silent}
              DetailPrint "Windows lists a Rutba Office for all users that is no longer there; forgetting it needs an administrator, so the entry stays."
            ${Else}
              ${If} ${Cmd} `MessageBox MB_YESNO|MB_ICONQUESTION|MB_DEFBUTTON2 "Windows lists a Rutba Office for all users that is no longer there. Forget it now? Windows asks for an administrator's yes next; No leaves it listed." IDYES`
                ClearErrors
                ExecShellWait "runas" "$SYSDIR\cmd.exe" '/d /c "$R2"' SW_HIDE
              ${EndIf}
            ${EndIf}
          ${EndIf}
        !endif
      ${EndIf}
    ${EndIf}
    ; The uninstall entry goes, unless its program stays. The old uninstaller
    ; took the install key with it; a silent update of an all-users copy has
    ; just put that back.
    ${If} $R9 != "keep"
      DeleteRegKey ${ROOT_KEY} "${UNINSTALL_REGISTRY_KEY}"
    ${EndIf}
  ${EndIf}
  ClearErrors
!macroend

; A copy no uninstall entry names: the program left alone by an earlier
; removal whose entries went with it, or a whole copy with only its install
; key left. The removal above goes by the entries and never sees it, and
; every pin made from a husk starts it and it ends at once. For KIND, its
; folder as rutbaScopeFolder finds it and its default folder are looked in;
; one that holds our program is removed like the rest (rutbaClearLeftovers,
; only our files), unless it is where this install goes (it is installed
; over). A husk that cannot be removed is said in the log and the install
; goes on: it is not the copy kept, and the app's pin repair points the pins
; away from it. A whole copy that cannot be removed stops the install
; rather than leave two. Uses $R2 to $R6 and $R9.
!macro rutbaSweepOne SCOPE DIR
  ${If} "${DIR}" != ""
  ${AndIf} ${FileExists} "${DIR}\${APP_EXECUTABLE_FILENAME}"
    !insertmacro rutbaSamePath "${DIR}" "$INSTDIR" $R5
    ${If} $R5 != "1"
      !insertmacro rutbaCopyWorks "${DIR}" $R9
      StrCpy $R2 ""
      StrCpy $rutbaCancelled ""
      !insertmacro rutbaClearLeftovers ${SCOPE} "${DIR}"
      ${If} ${FileExists} "${DIR}\${APP_EXECUTABLE_FILENAME}"
        ${If} $R9 == "1"
          ${IfNot} ${Silent}
          ${AndIf} $rutbaCancelled != "1"
            MessageBox MB_OK|MB_ICONEXCLAMATION "The Rutba Office in ${DIR} was not removed, so this install has stopped rather than leave two copies. Run it again and allow the administrator prompt, or remove that copy first."
          ${EndIf}
          SetErrorLevel 2
          Quit
        ${EndIf}
        DetailPrint "Only the program is left in ${DIR}; it could not be removed, and stays."
      ${EndIf}
    ${EndIf}
  ${EndIf}
!macroend

!macro rutbaSweepKind KIND
  !if "${KIND}" == "machine"
    ReadRegStr $R6 HKLM "${UNINSTALL_REGISTRY_KEY}" "UninstallString"
  !else
    ReadRegStr $R6 HKCU "${UNINSTALL_REGISTRY_KEY}" "UninstallString"
  !endif
  ${If} $R6 == ""
    Push $0
    Push $1
    !if "${KIND}" == "machine"
      !insertmacro rutbaScopeFolder HKLM "$perMachineInstallationFolder" machine $R4
    !else
      !insertmacro rutbaScopeFolder HKCU "$perUserInstallationFolder" user $R4
    !endif
    Pop $1
    Pop $0
    Push $R4
    !if "${KIND}" == "machine"
      !insertmacro rutbaSweepOne allusers "$R4"
    !else
      !insertmacro rutbaSweepOne currentuser "$R4"
    !endif
    Pop $R6
    !insertmacro rutbaDefaultFolder ${KIND} $R4
    !insertmacro rutbaSamePath "$R4" "$R6" $R5
    ${If} $R5 != "1"
      !if "${KIND}" == "machine"
        !insertmacro rutbaSweepOne allusers "$R4"
      !else
        !insertmacro rutbaSweepOne currentuser "$R4"
      !endif
    ${EndIf}
  ${EndIf}
  ClearErrors
!macroend

; Every copy but the one this install keeps ($rutbaKept, "all" or "user"),
; removed, and that one's old self: the other kind first, by its entry and
; then wherever else it is, so that a removal which has to stop has not yet
; touched the copy kept; then the kind kept. Uses $R0 to $R9.
!macro rutbaRemoveCopies
  ${If} $rutbaKept == "all"
    !insertmacro rutbaRemoveInstalledCopy HKCU currentuser other
    !insertmacro rutbaSweepKind user
    !insertmacro rutbaRemoveInstalledCopy HKLM allusers own
    !insertmacro rutbaSweepKind machine
  ${Else}
    !insertmacro rutbaRemoveInstalledCopy HKLM allusers other
    !insertmacro rutbaSweepKind machine
    !insertmacro rutbaRemoveInstalledCopy HKCU currentuser own
    !insertmacro rutbaSweepKind user
  ${EndIf}
!macroend

!macro customInit
  Push $R0
  Push $R1
  Push $R2
  Push $R3
  Push $R4
  Push $R5
  Push $R6
  Push $R7
  Push $R8
  Push $R9
  ; Where each copy is, written where the template reads it, then which
  ; scope a silent install goes to, before anything is removed.
  !insertmacro rutbaRecordFolders
  !insertmacro rutbaSilentScope
  ; With pages, the removal waits for the last of them (rutbaRemoveOldCopies).
  ; A silent update of an all-users copy elevates in its install section and
  ; comes through here again as administrator: the old copy waits for that
  ; pass, or a refused elevation would leave no Office at all.
  ${If} ${Silent}
    ${If} $hasPerMachineInstallation != "1"
    ${OrIf} ${UAC_IsAdmin}
      StrCpy $rutbaKept "user"
      ${If} $hasPerMachineInstallation == "1"
        StrCpy $rutbaKept "all"
      ${EndIf}
      Call rutbaRemoveCopiesNow
    ${EndIf}
  ${EndIf}
  ; Which app packs to write, before the pages show it (installer-packs.nsh).
  !insertmacro rutbaPacksDecide
  Pop $R9
  Pop $R8
  Pop $R7
  Pop $R6
  Pop $R5
  Pop $R4
  Pop $R3
  Pop $R2
  Pop $R1
  Pop $R0
!macroend

; Choose apps (owner, 2026-10-20: "since the number of applications have
; grown we should let users to choose what they wanted to install and what
; not"). A page after the folder: a box for each app, ticked as the person
; last chose. Office writes the apps on now to HKCU\Software\Rutba Office
; AppsOn, and those off to AppsOff, at every start and on every change in
; Choose apps (owner, 2026-10-21: "the installer should always remember the
; previous choices and currently enabled applications"), so the page opens
; on exactly the apps in use. The apps ticked and unticked are written back
; to the same values, and Office takes the choice from them on its next
; start: their tiles go, and the kinds of file they open are let go of.
; An app turned off here stays installed, and can be turned on again from
; Choose apps; which apps are installed at all is the page before this one
; (Apps to install). A silent install (an update) shows no page and leaves
; the choice as it was.

!include "nsDialogs.nsh"
!include "WordFunc.nsh"
; The page's words in Urdu, Arabic and Hindi, made from the catalogues (tools/make-installer-words.mjs).
!include "${__FILEDIR__}\installer-words.nsh"
; The Languages page's keys and names, made from languages.json (tools/make-installer-languages.mjs).
!include "${__FILEDIR__}\installer-languages.nsh"
!ifdef BUILD_UNINSTALLER
  !insertmacro rutbaUninstallerWords
!endif
; The app packs this installer carries, made from the build before electron-builder runs
; (build/make-app-packs.js): a pack's files are written only when its box is ticked.
!include "${PROJECT_DIR}\build\packs\installer-packs.nsh"

; Apps to install (docs/APP-PACKS.md 5.3). A page before Choose apps: a box
; for every app pack, Text recognition and AI Tools among them, all ticked on
; a first install and ticked as installed on the next. Only the packs ticked
; are written, and only their kinds of file registered; an app not ticked is
; not on the computer at all, and Choose apps below shows it unticked and
; greyed. /PACKS=mail,word on the command line chooses without the page. An
; update, silent, writes the packs installed (Packs), and every pack when
; updating a version with one archive, which had no Packs.

; One box. With the apps on written before (AppsOn, by Office or an earlier
; page): ticked when the app is one of them. With only the apps off (AppsOff,
; from a copy older than AppsOn): ticked unless it is one of those, for an app
; that copy knew of; an app newer than it (NEW 1) takes the first-install
; choice. On a first install: an Office app ticked, a further app (OPTIONAL 1) not.
!macro rutbaAppBox KEY LABEL X Y OPTIONAL NEW
  ${NSD_CreateCheckbox} ${X} ${Y} 46% 12u "${LABEL}"
  Pop $rutbaBox_${KEY}
  ${If} $rutbaAppsOnKnown == "1"
    ${WordFind} ",$rutbaAppsOnBefore," ",${KEY}," "E*" $0
    ${IfNot} ${Errors}
      ${NSD_Check} $rutbaBox_${KEY}
    ${EndIf}
  ${ElseIf} $rutbaAppsKnown == "1"
  ${AndIf} "${NEW}" == "0"
    ${WordFind} ",$rutbaAppsBefore," ",${KEY}," "E*" $0
    ${If} ${Errors}
      ${NSD_Check} $rutbaBox_${KEY}
    ${EndIf}
  ${ElseIf} "${OPTIONAL}" == "0"
    ${NSD_Check} $rutbaBox_${KEY}
  ${EndIf}
  ClearErrors
  ${If} $rutbaPack_${KEY} != "1"
    ${NSD_Uncheck} $rutbaBox_${KEY}
    EnableWindow $rutbaBox_${KEY} 0
  ${EndIf}
!macroend

!macro rutbaAppRead KEY
  ${NSD_GetState} $rutbaBox_${KEY} $1
  ${If} $rutbaPack_${KEY} != "1"
    ; Not installed: on in the choice, so it shows once it is added.
    StrCpy $rutbaAppsOn "$rutbaAppsOn${KEY},"
  ${ElseIf} $1 != ${BST_CHECKED}
    StrCpy $rutbaAppsOff "$rutbaAppsOff${KEY},"
  ${Else}
    StrCpy $rutbaAppsOn "$rutbaAppsOn${KEY},"
    IntOp $2 $2 + 1
  ${EndIf}
!macroend

; The Languages page's slot macros, defined here at the top level: NSIS
; cannot define a macro inside another, and inside customPageAfterChangeDir
; they stopped every installer build ("can't define a macro inside a
; macro", since bac18fb1). They are expanded where they are inserted, so
; the variables that macro declares are theirs as before.
; The Nth language found into its slot. The list starts and ends with a
; comma; WordFind counts no empty word before the first or after the last, so
; slot 1 is the first pack, and a list of "," leaves every slot empty
; (tests/installer-nsis.test.js builds this with NSIS and checks it).
!macro rutbaLangSlot N
  StrCpy $rutbaLang${N} ""
  ClearErrors
  ${WordFind} "$rutbaLangFound" "," "E+${N}" $R2
  ${IfNot} ${Errors}
    StrCpy $rutbaLang${N} $R2
  ${EndIf}
  ClearErrors
!macroend

; The Nth box, ticked, when there is an Nth language.
!macro rutbaLangBox N Y
  ${If} $rutbaLang${N} != ""
    StrCpy $R2 $rutbaLang${N}
    Call rutbaLanguageLabel
    ; A draft pack is marked so, in the page's words.
    ${If} $R8 == "draft"
      ${WordReplace} "$rutbaW_ldraft" "{name}" "$R5" "+" $R5
    ${EndIf}
    ${NSD_CreateCheckbox} 0 ${Y} 100% 12u "$R5"
    Pop $rutbaLangBox${N}
    ${NSD_Check} $rutbaLangBox${N}
  ${EndIf}
!macroend

!macro rutbaLangRead N
  ${If} $rutbaLang${N} != ""
    ${NSD_GetState} $rutbaLangBox${N} $1
    ${If} $1 == ${BST_CHECKED}
      StrCpy $rutbaLangChosen "$rutbaLangChosen$rutbaLang${N},"
    ${EndIf}
  ${EndIf}
!macroend

; The page, its variables and its functions are put in where the template
; asks for a page after the folder: the installer's UI library (MUI) is
; loaded there, and the uninstaller, built without the pages, never sees them.
!macro customPageAfterChangeDir
  Page custom rutbaPacksPageShow rutbaPacksPageLeave
  Page custom rutbaAppsPageShow rutbaAppsPageLeave
  Page custom rutbaLanguagesPageShow rutbaLanguagesPageLeave
  Page custom rutbaRemoveOldCopies

  Var rutbaAppsShown
  Var rutbaAppsOff
  Var rutbaAppsBefore
  Var rutbaAppsOn
  Var rutbaAppsOnBefore
  Var rutbaAppsOnKnown
  Var rutbaBox_mail
  Var rutbaBox_calendar
  Var rutbaBox_contacts
  Var rutbaBox_word
  Var rutbaBox_sheets
  Var rutbaBox_slides
  Var rutbaBox_pdf
  Var rutbaBox_pictures
  Var rutbaBox_image
  Var rutbaBox_video
  Var rutbaBox_recorder
  Var rutbaBox_broadcast
  Var rutbaBox_talk
  Var rutbaBox_base
  Var rutbaBox_ai
  Var rutbaAppsKnown
  Var rutbaAddOns
  Var rutbaAddOnsBefore
  Var rutbaLangFound
  Var rutbaLangShown
  Var rutbaLangChosen
  Var rutbaLang1
  Var rutbaLang2
  Var rutbaLang3
  Var rutbaLang4
  Var rutbaLang5
  Var rutbaLang6
  Var rutbaLangBox1
  Var rutbaLangBox2
  Var rutbaLangBox3
  Var rutbaLangBox4
  Var rutbaLangBox5
  Var rutbaLangBox6
  !insertmacro rutbaInstallerWords
  !insertmacro rutbaLanguageFunctions

; The removal (rutbaRemoveCopies), once in the installer: a silent install
; calls it from .onInit (customInit), one with pages from the moment after
; the last page. $rutbaKept says the kind kept.
Function rutbaRemoveCopiesNow
  Push $R0
  Push $R1
  Push $R2
  Push $R3
  Push $R4
  Push $R5
  Push $R6
  Push $R7
  Push $R8
  Push $R9
  !insertmacro rutbaRemoveCopies
  Pop $R9
  Pop $R8
  Pop $R7
  Pop $R6
  Pop $R5
  Pop $R4
  Pop $R3
  Pop $R2
  Pop $R1
  Pop $R0
FunctionEnd

; Not a page anyone sees: the moment after the last one, when the page that
; asks who Office is for has set the kind. Abort moves on to the install.
Function rutbaRemoveOldCopies
  StrCpy $rutbaKept "user"
  ${If} $installMode == "all"
    StrCpy $rutbaKept "all"
  ${EndIf}
  Call rutbaRemoveCopiesNow
  Abort
FunctionEnd

Function rutbaPacksPageShow
  ${If} ${isUpdated}
    Abort
  ${EndIf}
  Call rutbaPickWords
  !insertmacro MUI_HEADER_TEXT "$rutbaW_ptitle" "$rutbaW_psub"
  nsDialogs::Create 1018
  Pop $0
  ${If} $0 == error
    Abort
  ${EndIf}
  ${NSD_CreateLabel} 0 0 100% 20u "$rutbaW_plead"
  Pop $0
  !insertmacro rutbaPackBoxes
  nsDialogs::Show
FunctionEnd

Function rutbaPacksPageLeave
  StrCpy $2 0
  !insertmacro rutbaPackReads
  ${If} $2 == 0
    MessageBox MB_OK|MB_ICONEXCLAMATION "$rutbaW_none"
    Abort
  ${EndIf}
FunctionEnd

Function rutbaAppsPageShow
  ${If} ${isUpdated}
    Abort
  ${EndIf}
  Call rutbaPickWords
  !insertmacro MUI_HEADER_TEXT "$rutbaW_title" "$rutbaW_sub"
  StrCpy $rutbaAppsKnown "0"
  ReadRegStr $rutbaAppsBefore HKCU "Software\Rutba Office" "AppsOff"
  ${IfNot} ${Errors}
    StrCpy $rutbaAppsKnown "1"
  ${EndIf}
  ClearErrors
  StrCpy $rutbaAppsOnKnown "0"
  ReadRegStr $rutbaAppsOnBefore HKCU "Software\Rutba Office" "AppsOn"
  ${IfNot} ${Errors}
    StrCpy $rutbaAppsOnKnown "1"
  ${EndIf}
  ClearErrors
  ReadRegStr $rutbaAddOnsBefore HKCU "Software\Rutba Office" "AddOns"
  ClearErrors
  nsDialogs::Create 1018
  Pop $0
  ${If} $0 == error
    Abort
  ${EndIf}
  ${NSD_CreateLabel} 0 0 100% 20u "$rutbaW_lead"
  Pop $0
  ; The two sets, in the order Home and Choose apps have them: nine boxes fill the page's height at 12u each.
  ${NSD_CreateLabel} 0 21u 46% 10u "$rutbaW_office"
  Pop $0
  !insertmacro rutbaAppBox mail "$rutbaW_mail" 0 32u 0 0
  !insertmacro rutbaAppBox calendar "$rutbaW_calendar" 0 44u 0 0
  !insertmacro rutbaAppBox contacts "$rutbaW_contacts" 0 56u 0 0
  !insertmacro rutbaAppBox talk "$rutbaW_talk" 0 68u 1 0
  !insertmacro rutbaAppBox word "$rutbaW_word" 0 80u 0 0
  !insertmacro rutbaAppBox sheets "$rutbaW_sheets" 0 92u 0 0
  !insertmacro rutbaAppBox slides "$rutbaW_slides" 0 104u 0 0
  !insertmacro rutbaAppBox pdf "$rutbaW_pdf" 0 116u 0 0
  ; New in 2.9: a 2.8 copy's AppsOff never listed it.
  !insertmacro rutbaAppBox base "$rutbaW_base" 0 128u 1 1
  ${NSD_CreateLabel} 52% 21u 46% 10u "$rutbaW_studio"
  Pop $0
  !insertmacro rutbaAppBox pictures "$rutbaW_pictures" 52% 32u 0 0
  !insertmacro rutbaAppBox image "$rutbaW_image" 52% 44u 1 0
  !insertmacro rutbaAppBox video "$rutbaW_video" 52% 56u 1 0
  !insertmacro rutbaAppBox recorder "$rutbaW_recorder" 52% 68u 1 0
  ; New in 2.22, and Recorder's Go Live before that: ticked wherever Recorder was, unless turned off itself.
  !insertmacro rutbaAppBox broadcast "$rutbaW_broadcast" 52% 80u 1 1
  ${If} $rutbaAppsOnKnown == "1"
  ${AndIf} $rutbaPack_broadcast == "1"
    ${WordFind} ",$rutbaAppsOnBefore," ",recorder," "E*" $0
    ${IfNot} ${Errors}
      ${WordFind} ",$rutbaAppsBefore," ",broadcast," "E*" $0
      ${If} ${Errors}
        ${NSD_Check} $rutbaBox_broadcast
      ${EndIf}
    ${EndIf}
    ClearErrors
  ${EndIf}
  ${NSD_CreateLabel} 52% 97u 46% 10u "$rutbaW_addons"
  Pop $0
  ${NSD_CreateCheckbox} 52% 108u 46% 12u "$rutbaW_ai"
  Pop $rutbaBox_ai
  ${WordFind} ",$rutbaAddOnsBefore," ",ai," "E*" $0
  ${IfNot} ${Errors}
    ${NSD_Check} $rutbaBox_ai
  ${EndIf}
  ClearErrors
  ${If} $rutbaPack_ai != "1"
    ${NSD_Uncheck} $rutbaBox_ai
    EnableWindow $rutbaBox_ai 0
  ${EndIf}
  nsDialogs::Show
FunctionEnd

; Languages (L5). A page after Choose apps, only on a first install and only
; when the computer uses a language Rutba Office has a pack for: a box for
; each, ticked. The packs ticked are written to HKCU\Software\Rutba Office
; LanguagesChosen, and Office adds them at its first start while the computer
; is online. English is always there and never offered.

; $R1 a tag the computer uses: its pack added to the languages found, once.
Function rutbaFoundTag
  Call rutbaPackForTag
  ${If} $R2 != ""
    ${WordFind} "$rutbaLangFound" ",$R2," "E*" $R6
    ${If} ${Errors}
      StrCpy $rutbaLangFound "$rutbaLangFound$R2,"
    ${EndIf}
    ClearErrors
  ${EndIf}
FunctionEnd

; The languages this computer uses, as packs, into the slots: "" in the first when none.
Function rutbaFindLanguages
  Push $R0
  Push $R1
  Push $R2
  Push $R4
  Push $R6
  Push $R7
  StrCpy $rutbaLangFound ","
  StrCpy $R0 0
  ${Do}
    ClearErrors
    EnumRegKey $R1 HKCU "Control Panel\International\User Profile" $R0
    ${If} ${Errors}
    ${OrIf} $R1 == ""
      ${ExitDo}
    ${EndIf}
    Call rutbaFoundTag
    IntOp $R0 $R0 + 1
  ${Loop}
  ClearErrors
  ; The language Windows shows itself in, which a computer set up in it may not list above.
  System::Call 'kernel32::GetUserDefaultUILanguage() i .R7'
  StrCpy $R1 ""
  System::Call 'kernel32::LCIDToLocaleName(i R7, w .R1, i 85, i 0) i'
  ${If} $R1 != ""
    Call rutbaFoundTag
  ${EndIf}
  !insertmacro rutbaDropRegionalBases $rutbaLangFound
  !insertmacro rutbaLangSlot 1
  !insertmacro rutbaLangSlot 2
  !insertmacro rutbaLangSlot 3
  !insertmacro rutbaLangSlot 4
  !insertmacro rutbaLangSlot 5
  !insertmacro rutbaLangSlot 6
  Pop $R7
  Pop $R6
  Pop $R4
  Pop $R2
  Pop $R1
  Pop $R0
FunctionEnd

Function rutbaLanguagesPageShow
  ${If} ${isUpdated}
    Abort
  ${EndIf}
  Call rutbaFindLanguages
  ${If} $rutbaLang1 == ""
    Abort
  ${EndIf}
  Call rutbaPickWords
  !insertmacro MUI_HEADER_TEXT "$rutbaW_ltitle" "$rutbaW_lsub"
  nsDialogs::Create 1018
  Pop $0
  ${If} $0 == error
    Abort
  ${EndIf}
  ${NSD_CreateLabel} 0 0 100% 30u "$rutbaW_llead"
  Pop $0
  Push $R2
  Push $R5
  !insertmacro rutbaLangBox 1 34u
  !insertmacro rutbaLangBox 2 47u
  !insertmacro rutbaLangBox 3 60u
  !insertmacro rutbaLangBox 4 73u
  !insertmacro rutbaLangBox 5 86u
  !insertmacro rutbaLangBox 6 99u
  Pop $R5
  Pop $R2
  nsDialogs::Show
FunctionEnd

Function rutbaLanguagesPageLeave
  StrCpy $rutbaLangChosen ""
  !insertmacro rutbaLangRead 1
  !insertmacro rutbaLangRead 2
  !insertmacro rutbaLangRead 3
  !insertmacro rutbaLangRead 4
  !insertmacro rutbaLangRead 5
  !insertmacro rutbaLangRead 6
  StrCpy $rutbaLangShown "1"
FunctionEnd

Function rutbaAppsPageLeave
  StrCpy $rutbaAppsOff ""
  StrCpy $rutbaAppsOn ""
  StrCpy $2 0
  !insertmacro rutbaAppRead mail
  !insertmacro rutbaAppRead calendar
  !insertmacro rutbaAppRead contacts
  !insertmacro rutbaAppRead talk
  !insertmacro rutbaAppRead word
  !insertmacro rutbaAppRead sheets
  !insertmacro rutbaAppRead slides
  !insertmacro rutbaAppRead pdf
  !insertmacro rutbaAppRead base
  !insertmacro rutbaAppRead pictures
  !insertmacro rutbaAppRead image
  !insertmacro rutbaAppRead video
  !insertmacro rutbaAppRead recorder
  !insertmacro rutbaAppRead broadcast
  ${If} $2 == 0
    MessageBox MB_OK|MB_ICONEXCLAMATION "$rutbaW_none"
    Abort
  ${EndIf}
  StrCpy $rutbaAddOns ""
  ${NSD_GetState} $rutbaBox_ai $1
  ${If} $1 == ${BST_CHECKED}
  ${OrIf} $rutbaPack_ai != "1"
    StrCpy $rutbaAddOns "ai"
  ${EndIf}
  StrCpy $rutbaAppsShown "1"
FunctionEnd
!macroend

; An entry in the other scope, forgotten: its copy was removed (or the
; install stopped) before this one began, so whatever is left there is
; stale and would make initMultiUser see that scope registered again.
!macro rutbaForgetOtherScope ROOT_KEY
  DeleteRegKey ${ROOT_KEY} "${UNINSTALL_REGISTRY_KEY}"
  DeleteRegValue ${ROOT_KEY} "${INSTALL_REGISTRY_KEY}" "InstallLocation"
  ClearErrors
!macroend

; After the files and the template's entry: the entry Windows lists says the
; version installed and names its folder. electron-builder writes the version
; but never InstallLocation, so an entry whose uninstaller later goes could
; be found only by its UninstallString, and a reader of the registry saw the
; folder empty. And no entry is left in the other scope. An all-users entry
; is not this installer's to change unless it is elevated; one for this
; person always is.
!macro rutbaEntryTrue
  WriteRegStr SHELL_CONTEXT "${UNINSTALL_REGISTRY_KEY}" "DisplayVersion" "${VERSION}"
  WriteRegStr SHELL_CONTEXT "${UNINSTALL_REGISTRY_KEY}" "InstallLocation" "$INSTDIR"
  ${If} $installMode == "all"
    !insertmacro rutbaForgetOtherScope HKCU
  ${ElseIf} ${UAC_IsAdmin}
    !insertmacro rutbaForgetOtherScope HKLM
  ${EndIf}
  ClearErrors
!macroend

; The install is judged before it is called done (the owner, 22 Nov 2026:
; "The installer must end with an error when it installs nothing"). NSIS
; goes on past a file it could not write and an entry it could not make, and
; a silent install shows nothing: after 2.26.1 an update run without
; elevation, over what was left of a copy for all users, ended with 0 and no
; Office registered anywhere. So, after the files and the entries, what must
; be true of any install is looked at: the program and its resources are in
; the folder, with the uninstaller; the program is this version's (one a
; helper held open is not replaced); and the entry Windows lists for this
; kind of install names this version and this folder. If not, it says which,
; where there is a window to say it in, and ends with 2, as every other stop
; here does.
!macro rutbaInstallLanded
  Push $R0
  Push $R1
  Push $R2
  Push $R3
  Push $R4
  StrCpy $R0 ""
  !insertmacro rutbaCopyWorks "$INSTDIR" $R1
  ${If} $R1 != "1"
    StrCpy $R0 "the program is not all in $INSTDIR"
  ${Else}
    ${IfNot} ${FileExists} "$INSTDIR\${UNINSTALL_FILENAME}"
      StrCpy $R0 "the uninstaller is not in $INSTDIR"
    ${EndIf}
  ${EndIf}
  ${If} $R0 == ""
    ; The program's own version, where it carries one: the build stamps it
    ; with the version's three numbers and a fourth.
    ClearErrors
    GetDLLVersion "$INSTDIR\${APP_EXECUTABLE_FILENAME}" $R1 $R2
    ${IfNot} ${Errors}
      IntOp $R3 $R1 >> 16
      IntOp $R3 $R3 & 0xFFFF
      IntOp $R4 $R1 & 0xFFFF
      IntOp $R2 $R2 >> 16
      IntOp $R2 $R2 & 0xFFFF
      StrCpy $R1 "$R3.$R4.$R2"
      StrLen $R3 "$R1"
      StrCpy $R4 "${VERSION}" $R3
      ${If} $R4 != $R1
        StrCpy $R0 "the program in $INSTDIR is version $R1, not ${VERSION}"
      ${EndIf}
    ${EndIf}
    ClearErrors
  ${EndIf}
  ${If} $R0 == ""
    ReadRegStr $R1 SHELL_CONTEXT "${UNINSTALL_REGISTRY_KEY}" "DisplayVersion"
    ReadRegStr $R2 SHELL_CONTEXT "${UNINSTALL_REGISTRY_KEY}" "InstallLocation"
    ${If} $R1 != "${VERSION}"
    ${OrIf} $R2 != "$INSTDIR"
      StrCpy $R0 "Windows was not told that version ${VERSION} is in $INSTDIR"
    ${EndIf}
    ClearErrors
  ${EndIf}
  ${If} $R0 != ""
    DetailPrint "Rutba Office ${VERSION} was not installed: $R0."
    ${IfNot} ${Silent}
      MessageBox MB_OK|MB_ICONSTOP "Rutba Office ${VERSION} was not installed: $R0.$\r$\n$\r$\nClose every Rutba Office window and run this again. If Office is installed for everyone on this computer, allow the administrator prompt."
    ${EndIf}
    SetErrorLevel 2
    Quit
  ${EndIf}
  Pop $R4
  Pop $R3
  Pop $R2
  Pop $R1
  Pop $R0
!macroend

!macro customInstall
  !ifndef BUILD_UNINSTALLER
  !insertmacro rutbaEntryTrue
  ; The app packs chosen, and only their kinds of file (installer-packs.nsh).
  !insertmacro rutbaWritePacks
  ${If} $rutbaAppsShown == "1"
    WriteRegStr HKCU "Software\Rutba Office" "AppsOff" "$rutbaAppsOff"
    WriteRegStr HKCU "Software\Rutba Office" "AppsOn" "$rutbaAppsOn"
    WriteRegStr HKCU "Software\Rutba Office" "AddOns" "$rutbaAddOns"
    WriteRegStr HKCU "Software\Rutba Office" "AppsChosenAt" "installer"
  ${EndIf}
  ; The languages ticked on the Languages page, for Office's first start (L5); none ticked is a choice too.
  ${If} $rutbaLangShown == "1"
    WriteRegStr HKCU "Software\Rutba Office" "LanguagesChosen" "$rutbaLangChosen"
  ${EndIf}
  ; Last: an install that put nothing here does not end as a success.
  !insertmacro rutbaInstallLanded
  !endif
!macroend

; "Start with Windows" (the icon by the clock, D-2026-11-22-23). Rutba Office
; asks Windows to start it at sign-in with one value a profile under the
; person's Run key: named by the application's id (appUserModelId in
; office-shell's electron/main.js), and by that id with the profile's after a
; dot for a profile that is not the first. Windows keeps whether each is
; allowed under a second key, by the same name. A copy that is removed takes
; its own with it, or they would start a program that is no longer there; an
; update leaves them, since the setting is the person's. Only a value that
; starts the program in DIR goes: another copy of Rutba Office on this
; computer keeps its own. RUN and APPROVED are the two keys under HKCU (a
; test hands in stand-ins), ID the application's id. It is the Run key of
; whoever removes the copy: where a copy for everyone is removed, the entries
; of other people on this computer stay until they turn the setting off.
!macro rutbaForgetStartup RUN APPROVED ID DIR
  Push $R0
  Push $R1
  Push $R2
  Push $R3
  Push $R4
  StrLen $R4 "${ID}."
  StrCpy $R0 0
  ${Do}
    ClearErrors
    EnumRegValue $R1 HKCU "${RUN}" $R0
    ${If} ${Errors}
      ${ExitDo}
    ${EndIf}
    StrCpy $R2 $R1 $R4
    StrCpy $R3 ""
    ${If} $R1 == "${ID}"
    ${OrIf} $R2 == "${ID}."
      ReadRegStr $R3 HKCU "${RUN}" "$R1"
    ${EndIf}
    ; What it starts: the program in this folder, or it is not this copy's to remove.
    ${If} $R3 != ""
      ClearErrors
      ${WordFind} "$R3" "${DIR}\" "E+1{" $R2
      ${IfNot} ${Errors}
        DeleteRegValue HKCU "${RUN}" "$R1"
        DeleteRegValue HKCU "${APPROVED}" "$R1"
        ; The values after it have moved up one: the same place is read again.
        ${Continue}
      ${EndIf}
    ${EndIf}
    IntOp $R0 $R0 + 1
  ${Loop}
  ClearErrors
  Pop $R4
  Pop $R3
  Pop $R2
  Pop $R1
  Pop $R0
!macroend

; Every kind of file a pack registered is let go of, whichever packs were
; installed: electron-builder no longer registers them, so it no longer
; removes them either.
!macro customUnInstall
  !insertmacro rutbaForgetTypes
  ; "Start with Windows" goes with the copy removed; an update keeps it.
  ${IfNot} ${isUpdated}
    !insertmacro rutbaForgetStartup "Software\Microsoft\Windows\CurrentVersion\Run" "Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run" "co.techstyle.rutba.office" "$INSTDIR"
  ${EndIf}
  ; The voices and languages added (C9): asked, with the room they take,
  ; whether they go too. Kept unless the person says Yes, and never asked on
  ; an update or a silent removal, which replace the program and keep them.
  ; A folder named by RUTBA_ADDONS_DIR, or beside a portable copy, is the
  ; person's own and is left alone. Every copy keeps them in the person's own
  ; folder (D-2026-11-13-35), an install for everyone too: removing one asks
  ; about the add-ons of the person removing it, and those of anyone else on
  ; this computer stay theirs.
  ${IfNot} ${isUpdated}
    Push $R0
    Push $R1
    Push $R2
    Push $R3
    Push $R4
    ReadEnvStr $R0 LOCALAPPDATA
    StrCpy $R0 "$R0\Rutba Office\addons"
    ${If} ${FileExists} "$R0\*.*"
      ${GetSize} "$R0" "/S=0K" $R1 $R2 $R3
      ${If} $R1 < 1024
        StrCpy $R1 "$R1 KB"
      ${Else}
        IntOp $R1 $R1 / 1024
        StrCpy $R1 "$R1 MB"
      ${EndIf}
      Call un.rutbaPickUninstallerWords
      ${WordReplace} "$rutbaW_ukeep" "{size}" "$R1" "+" $R4
      ${If} ${Cmd} `MessageBox MB_YESNO|MB_ICONQUESTION|MB_DEFBUTTON2 "$rutbaW_udrop$\r$\n$\r$\n$R4" /SD IDNO IDYES`
        DetailPrint "Removing the voices and languages added"
        RMDir /r "$R0"
        ; The folder above goes only when nothing else is in it.
        ${GetParent} "$R0" $R2
        RMDir "$R2"
      ${EndIf}
    ${EndIf}
    Pop $R4
    Pop $R3
    Pop $R2
    Pop $R1
    Pop $R0
  ${EndIf}
!macroend

; The Locker's safety copies (owner, 12 Nov 2026: "Uninstall removes the key
; and copies only if the person ticks a box. Otherwise it says they stay").
; A copy holds a folder's files as they were before locking, sealed under a
; key kept beside it in the Office profile (apps/desktop/main/extras/locker/
; safety.js). Uninstalling keeps the profile (deleteAppDataOnUninstall is
; false), and with it the copies, unless the person ticks the box on this
; page; it starts unticked, and a silent uninstall (an update) never ticks it.
; Only the safety folders go, of the first profile and of every other one
; (Profiles\<id>), never the documents, settings or locked folders.
; "Rutba Office" is the profile's folder name, PROFILE_NAME in
; main/template-set-choice.js.
!ifdef BUILD_UNINSTALLER
  !define MUI_COMPONENTSPAGE_TEXT_TOP "Your documents, settings and locked folders are kept. Any safety copies the Locker made are kept too, with their key, unless you tick the box below. Ticked, it removes them from your own Windows account only."
!endif

!macro rutbaRemoveSafetyIn DIR
  ${If} ${FileExists} "${DIR}\locker\safety\*.*"
    DetailPrint "Removing the Locker's safety copies from ${DIR}"
    RMDir /r "${DIR}\locker\safety"
  ${EndIf}
!macroend

!macro customUnInstallSection
  Section /o "un.Also remove the Locker's safety copies and their key" rutbaUnSafety
    Push $R0
    Push $R1
    ${IfNot} ${isUpdated}
      ; Electron keeps every profile in the person's own AppData, whichever way Office was installed.
      ${If} $installMode == "all"
        SetShellVarContext current
      ${EndIf}
      !insertmacro rutbaRemoveSafetyIn "$APPDATA\Rutba Office"
      FindFirst $R0 $R1 "$APPDATA\Rutba Office\Profiles\*"
      ${DoWhile} $R1 != ""
        ${If} $R1 != "."
        ${AndIf} $R1 != ".."
          !insertmacro rutbaRemoveSafetyIn "$APPDATA\Rutba Office\Profiles\$R1"
        ${EndIf}
        FindNext $R0 $R1
      ${Loop}
      FindClose $R0
      ${If} $installMode == "all"
        SetShellVarContext all
      ${EndIf}
    ${EndIf}
    Pop $R1
    Pop $R0
  SectionEnd

  ; Said in the list of what was done when the box was left unticked.
  Section "-un.rutbaSafetyKept"
    ${IfNot} ${SectionIsSelected} ${rutbaUnSafety}
    ${AndIfNot} ${isUpdated}
      DetailPrint "The Locker's safety copies, if any, are kept in your Rutba Office folder in AppData."
    ${EndIf}
  SectionEnd
!macroend
