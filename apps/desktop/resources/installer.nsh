; One copy of Rutba Office on a machine.
;
; electron-builder's installer removes the previous copy in the scope it is
; installing to - per-user, or all-users - and only that one. A copy in the
; other scope (an all-users install from an earlier day beside today's
; per-user one) stayed, and Windows listed two. This runs first, before the
; pages and before that step: it finds every registered copy, per-user and
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
      ; Run from a copy: the uninstaller removes the directory it lives in,
      ; and _?= keeps it in place so ExecWait actually waits for it.
      InitPluginsDir
      StrCpy $R5 "$PLUGINSDIR\rutba-old-uninstaller-${SCOPE}.exe"
      CopyFiles /SILENT "$R1" "$R5"
      ${IfNot} ${FileExists} "$R5"
        StrCpy $R5 "$R1"
      ${EndIf}
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
        ; end whatever is left, background processes too. Unsaved work is
        ; copied for Recover every minute, so at worst under a minute of it.
        ${If} $R7 == 5
          DetailPrint "Asking Rutba Office to close"
          nsExec::Exec '"$SYSDIR\taskkill.exe" /T /IM "${APP_EXECUTABLE_FILENAME}"'
          Pop $R6
        ${ElseIf} $R7 == 10
          DetailPrint "Closing what is left of Rutba Office"
          nsExec::Exec '"$SYSDIR\taskkill.exe" /F /T /IM "${APP_EXECUTABLE_FILENAME}"'
          Pop $R6
        ${EndIf}
        ${If} $R7 >= 15
          MessageBox MB_RETRYCANCEL|MB_ICONEXCLAMATION "Rutba Office (installed for ${SCOPE}, in $R4) is still open, so it cannot be replaced yet.$\r$\n$\r$\nClose every Rutba Office window and press Retry. Cancel stops the install and leaves the copy you have. (Uninstaller error $R6.)" /SD IDCANCEL IDRETRY +2
          Abort
          StrCpy $R7 0
        ${EndIf}
        Sleep 1000
      ${Loop}
      ; The uninstaller run from a copy cannot delete the original of itself.
      Delete "$R1"
      RMDir "$R4"
    ${Else}
      DetailPrint "Forgetting a Rutba Office entry (${SCOPE}) whose uninstaller is gone"
    ${EndIf}
    ; The uninstall entry goes; the install key stays, because it remembers
    ; the directory the next install should go to.
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
  !insertmacro rutbaRemoveInstalledCopy HKCU currentuser
  !insertmacro rutbaRemoveInstalledCopy HKLM allusers
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
; Nothing is left uninstalled: the apps share one program, and any can be
; turned on again from Choose apps. A silent install (an update) shows no
; page and leaves the choice as it was.

!include "nsDialogs.nsh"
!include "WordFunc.nsh"
; The page's words in Urdu, Arabic and Hindi, made from the catalogues (tools/make-installer-words.mjs).
!include "${__FILEDIR__}\installer-words.nsh"

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
!macroend

!macro rutbaAppRead KEY
  ${NSD_GetState} $rutbaBox_${KEY} $1
  ${If} $1 != ${BST_CHECKED}
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
  Page custom rutbaAppsPageShow rutbaAppsPageLeave

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
  Var rutbaBox_talk
  Var rutbaBox_base
  Var rutbaBox_ai
  Var rutbaAppsKnown
  Var rutbaAddOns
  Var rutbaAddOnsBefore
  !insertmacro rutbaInstallerWords

Function rutbaAppsPageShow
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
  ${NSD_CreateLabel} 0 22u 46% 10u "$rutbaW_office"
  Pop $0
  !insertmacro rutbaAppBox mail "$rutbaW_mail" 0 34u 0 0
  !insertmacro rutbaAppBox calendar "$rutbaW_calendar" 0 47u 0 0
  !insertmacro rutbaAppBox contacts "$rutbaW_contacts" 0 60u 0 0
  !insertmacro rutbaAppBox word "$rutbaW_word" 0 73u 0 0
  !insertmacro rutbaAppBox sheets "$rutbaW_sheets" 0 86u 0 0
  !insertmacro rutbaAppBox slides "$rutbaW_slides" 0 99u 0 0
  !insertmacro rutbaAppBox pdf "$rutbaW_pdf" 0 112u 0 0
  !insertmacro rutbaAppBox pictures "$rutbaW_pictures" 0 125u 0 0
  ${NSD_CreateLabel} 52% 22u 46% 10u "$rutbaW_more"
  Pop $0
  !insertmacro rutbaAppBox image "$rutbaW_image" 52% 34u 1 0
  !insertmacro rutbaAppBox video "$rutbaW_video" 52% 47u 1 0
  !insertmacro rutbaAppBox recorder "$rutbaW_recorder" 52% 60u 1 0
  !insertmacro rutbaAppBox talk "$rutbaW_talk" 52% 73u 1 0
  ; New in 2.9: a 2.8 copy's AppsOff never listed it.
  !insertmacro rutbaAppBox base "$rutbaW_base" 52% 86u 1 1
  ${NSD_CreateLabel} 52% 106u 46% 10u "$rutbaW_addons"
  Pop $0
  ${NSD_CreateCheckbox} 52% 118u 46% 12u "$rutbaW_ai"
  Pop $rutbaBox_ai
  ${WordFind} ",$rutbaAddOnsBefore," ",ai," "E*" $0
  ${IfNot} ${Errors}
    ${NSD_Check} $rutbaBox_ai
  ${EndIf}
  ClearErrors
  nsDialogs::Show
FunctionEnd

Function rutbaAppsPageLeave
  StrCpy $rutbaAppsOff ""
  StrCpy $rutbaAppsOn ""
  StrCpy $2 0
  !insertmacro rutbaAppRead mail
  !insertmacro rutbaAppRead calendar
  !insertmacro rutbaAppRead contacts
  !insertmacro rutbaAppRead word
  !insertmacro rutbaAppRead sheets
  !insertmacro rutbaAppRead slides
  !insertmacro rutbaAppRead pdf
  !insertmacro rutbaAppRead pictures
  !insertmacro rutbaAppRead image
  !insertmacro rutbaAppRead video
  !insertmacro rutbaAppRead recorder
  !insertmacro rutbaAppRead talk
  !insertmacro rutbaAppRead base
  ${If} $2 == 0
    MessageBox MB_OK|MB_ICONEXCLAMATION "$rutbaW_none"
    Abort
  ${EndIf}
  StrCpy $rutbaAddOns ""
  ${NSD_GetState} $rutbaBox_ai $1
  ${If} $1 == ${BST_CHECKED}
    StrCpy $rutbaAddOns "ai"
  ${EndIf}
  StrCpy $rutbaAppsShown "1"
FunctionEnd
!macroend

!macro customInstall
  !ifndef BUILD_UNINSTALLER
  ${If} $rutbaAppsShown == "1"
    WriteRegStr HKCU "Software\Rutba Office" "AppsOff" "$rutbaAppsOff"
    WriteRegStr HKCU "Software\Rutba Office" "AppsOn" "$rutbaAppsOn"
    WriteRegStr HKCU "Software\Rutba Office" "AddOns" "$rutbaAddOns"
    WriteRegStr HKCU "Software\Rutba Office" "AppsChosenAt" "installer"
  ${EndIf}
  !endif
!macroend
