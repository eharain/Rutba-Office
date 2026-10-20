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
; not"). A page after the folder: a box for each app, all ticked, or ticked
; as the person last chose (Office keeps HKCU\Software\Rutba Office\AppsOff
; up to date). The apps left unticked are written to that value, and Office
; takes the choice from it on its next start: their tiles go, and the kinds
; of file they open are let go of. Nothing is left uninstalled: the apps
; share one program, and any can be turned on again from Choose apps.
; A silent install (an update) shows no page and leaves the choice as it was.

!include "nsDialogs.nsh"
!include "WordFunc.nsh"

; One box: ticked unless the app is in the list turned off before.
!macro rutbaAppBox KEY LABEL X Y
  ${NSD_CreateCheckbox} ${X} ${Y} 46% 12u "${LABEL}"
  Pop $rutbaBox_${KEY}
  ${WordFind} ",$rutbaAppsBefore," ",${KEY}," "E*" $0
  ${If} ${Errors}
    ${NSD_Check} $rutbaBox_${KEY}
  ${EndIf}
  ClearErrors
!macroend

!macro rutbaAppRead KEY
  ${NSD_GetState} $rutbaBox_${KEY} $1
  ${If} $1 != ${BST_CHECKED}
    StrCpy $rutbaAppsOff "$rutbaAppsOff${KEY},"
  ${Else}
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

Function rutbaAppsPageShow
  !insertmacro MUI_HEADER_TEXT "Choose apps" "The apps you want to see and to open your files."
  ReadRegStr $rutbaAppsBefore HKCU "Software\Rutba Office" "AppsOff"
  ClearErrors
  nsDialogs::Create 1018
  Pop $0
  ${If} $0 == error
    Abort
  ${EndIf}
  ${NSD_CreateLabel} 0 0 100% 22u "Untick an app you do not want. It is not shown and does not open its kinds of file, and it can be turned on again at any time from Choose apps in Rutba Office."
  Pop $0
  !insertmacro rutbaAppBox mail "Mail" 0 28u
  !insertmacro rutbaAppBox calendar "Calendar" 0 42u
  !insertmacro rutbaAppBox contacts "Contacts" 0 56u
  !insertmacro rutbaAppBox word "Documents" 0 70u
  !insertmacro rutbaAppBox sheets "Worksheets" 0 84u
  !insertmacro rutbaAppBox slides "Presentations" 0 98u
  !insertmacro rutbaAppBox pdf "PDF" 52% 28u
  !insertmacro rutbaAppBox pictures "Pictures" 52% 42u
  !insertmacro rutbaAppBox image "Image" 52% 56u
  !insertmacro rutbaAppBox video "Video" 52% 70u
  !insertmacro rutbaAppBox recorder "Recorder" 52% 84u
  nsDialogs::Show
FunctionEnd

Function rutbaAppsPageLeave
  StrCpy $rutbaAppsOff ""
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
  ${If} $2 == 0
    MessageBox MB_OK|MB_ICONEXCLAMATION "Choose one app at least."
    Abort
  ${EndIf}
  StrCpy $rutbaAppsShown "1"
FunctionEnd
!macroend

!macro customInstall
  !ifndef BUILD_UNINSTALLER
  ${If} $rutbaAppsShown == "1"
    WriteRegStr HKCU "Software\Rutba Office" "AppsOff" "$rutbaAppsOff"
    WriteRegStr HKCU "Software\Rutba Office" "AppsChosenAt" "installer"
  ${EndIf}
  !endif
!macroend
