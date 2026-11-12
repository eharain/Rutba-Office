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
; Included ahead of the template, so the macros are defined before .onInit
; inserts customInit; the helpers used are NSIS's own, since the template's
; (installUtil.nsh) are included after .onInit.

!include "LogicLib.nsh"
!include "FileFunc.nsh"

!macro rutbaRemoveInstalledCopy ROOT_KEY SCOPE
  ClearErrors
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

    ${If} ${FileExists} "$R1"
      ; The directory the uninstaller sits in is the installation.
      ${GetParent} "$R1" $R4
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
        ; every half minute, so at worst under that much of it.
        ${If} $R7 == 5
          DetailPrint "Asking Rutba Office to close"
          nsExec::Exec '"$SYSDIR\taskkill.exe" /T /IM "${APP_EXECUTABLE_FILENAME}"'
          Pop $R6
        ${ElseIf} $R7 == 10
          StrCpy $R6 "1"
          ${IfNot} ${Silent}
            ; find says 0 when a process of ours has a window title.
            nsExec::Exec '"$SYSDIR\cmd.exe" /c ""$SYSDIR\tasklist.exe" /FI "IMAGENAME eq ${APP_EXECUTABLE_FILENAME}" /FI "WINDOWTITLE ne N/A" /NH | "$SYSDIR\find.exe" /I "${APP_EXECUTABLE_FILENAME}""'
            Pop $R6
          ${EndIf}
          ${If} $R6 != "0"
            DetailPrint "Closing what is left of Rutba Office"
            nsExec::Exec '"$SYSDIR\taskkill.exe" /F /T /IM "${APP_EXECUTABLE_FILENAME}"'
            Pop $R6
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
      ${If} ${FileExists} "$R4\${APP_EXECUTABLE_FILENAME}"
        ${IfNot} ${Silent}
          MessageBox MB_OK|MB_ICONEXCLAMATION "The Rutba Office installed for ${SCOPE}, in $R4, was not removed, so this install has stopped rather than leave two copies. Run it again and allow the administrator prompt, or remove that copy first."
        ${EndIf}
        SetErrorLevel 2
        Quit
      ${EndIf}
      ; The uninstaller run from a copy cannot delete the original of itself.
      Delete "$R1"
      RMDir "$R4"
      !if "${SCOPE}" == "allusers"
        ${If} ${Silent}
        ${AndIf} $R8 != ""
          WriteRegStr ${ROOT_KEY} "${INSTALL_REGISTRY_KEY}" "InstallLocation" "$R8"
        ${EndIf}
      !endif
    ${Else}
      DetailPrint "Forgetting a Rutba Office entry (${SCOPE}) whose uninstaller is gone"
    ${EndIf}
    ; The uninstall entry goes. The old uninstaller took the install key with
    ; it; a silent update of an all-users copy has just put that back.
    DeleteRegKey ${ROOT_KEY} "${UNINSTALL_REGISTRY_KEY}"
  ${EndIf}
  ClearErrors
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
  ; With pages, the removal waits for the last of them (rutbaRemoveOldCopies).
  ; A silent update of an all-users copy elevates in its install section and
  ; comes through here again as administrator: the old copy waits for that
  ; pass, or a refused elevation would leave no Office at all.
  ${If} ${Silent}
    ${If} $hasPerMachineInstallation != "1"
    ${OrIf} ${UAC_IsAdmin}
      !insertmacro rutbaRemoveInstalledCopy HKCU currentuser
      !insertmacro rutbaRemoveInstalledCopy HKLM allusers
    ${EndIf}
  ${EndIf}
  ; Which app packs to write, before the pages show it (installer-packs.nsh).
  !insertmacro rutbaPacksDecide
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

; The page, its variables and its functions are put in where the template
; asks for a page after the folder: the installer's UI library (MUI) is
; loaded there, and the uninstaller, built without the pages, never sees them.
!macro customPageAfterChangeDir
  Page custom rutbaPacksPageShow rutbaPacksPageLeave
  Page custom rutbaAppsPageShow rutbaAppsPageLeave
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
  !insertmacro rutbaInstallerWords

; Not a page anyone sees: the moment after the last one. Abort moves on to the install.
Function rutbaRemoveOldCopies
  Push $R0
  Push $R1
  Push $R2
  Push $R3
  Push $R4
  Push $R5
  Push $R6
  Push $R7
  Push $R8
  !insertmacro rutbaRemoveInstalledCopy HKCU currentuser
  !insertmacro rutbaRemoveInstalledCopy HKLM allusers
  Pop $R8
  Pop $R7
  Pop $R6
  Pop $R5
  Pop $R4
  Pop $R3
  Pop $R2
  Pop $R1
  Pop $R0
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

!macro customInstall
  !ifndef BUILD_UNINSTALLER
  ; The app packs chosen, and only their kinds of file (installer-packs.nsh).
  !insertmacro rutbaWritePacks
  ${If} $rutbaAppsShown == "1"
    WriteRegStr HKCU "Software\Rutba Office" "AppsOff" "$rutbaAppsOff"
    WriteRegStr HKCU "Software\Rutba Office" "AppsOn" "$rutbaAppsOn"
    WriteRegStr HKCU "Software\Rutba Office" "AddOns" "$rutbaAddOns"
    WriteRegStr HKCU "Software\Rutba Office" "AppsChosenAt" "installer"
  ${EndIf}
  !endif
!macroend

; Every kind of file a pack registered is let go of, whichever packs were
; installed: electron-builder no longer registers them, so it no longer
; removes them either.
!macro customUnInstall
  !insertmacro rutbaForgetTypes
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
