; The installer's Languages page: the keys it matches the computer's languages with, and
; the names it shows. Made by tools/make-installer-languages.mjs from languages.json,
; with the keys language-choice.js uses: change the list there and make this again.

!macro rutbaLanguageFunctions
; $R1 a language tag, as Windows writes it; $R2 the pack for it, or "".
Function rutbaPackForTag
  StrCpy $R2 ""
  StrCpy $R4 $R1 8
  ${If} $R1 == "zh-hans"
  ${OrIf} $R4 == "zh-hans-"
    StrCpy $R2 "lang-zh-hans"
    Return
  ${EndIf}
  StrCpy $R4 $R1 8
  ${If} $R1 == "zh-hant"
  ${OrIf} $R4 == "zh-hant-"
    StrCpy $R2 "lang-zh-hant"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "pt-br"
  ${OrIf} $R4 == "pt-br-"
    StrCpy $R2 "lang-pt-br"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "zh-cn"
  ${OrIf} $R4 == "zh-cn-"
    StrCpy $R2 "lang-zh-hans"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "zh-hk"
  ${OrIf} $R4 == "zh-hk-"
    StrCpy $R2 "lang-zh-hant"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "zh-mo"
  ${OrIf} $R4 == "zh-mo-"
    StrCpy $R2 "lang-zh-hant"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "zh-my"
  ${OrIf} $R4 == "zh-my-"
    StrCpy $R2 "lang-zh-hans"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "zh-sg"
  ${OrIf} $R4 == "zh-sg-"
    StrCpy $R2 "lang-zh-hans"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "zh-tw"
  ${OrIf} $R4 == "zh-tw-"
    StrCpy $R2 "lang-zh-hant"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "de"
  ${OrIf} $R4 == "de-"
    StrCpy $R2 "lang-de"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "es"
  ${OrIf} $R4 == "es-"
    StrCpy $R2 "lang-es"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "fr"
  ${OrIf} $R4 == "fr-"
    StrCpy $R2 "lang-fr"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "it"
  ${OrIf} $R4 == "it-"
    StrCpy $R2 "lang-it"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "nl"
  ${OrIf} $R4 == "nl-"
    StrCpy $R2 "lang-nl"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "pl"
  ${OrIf} $R4 == "pl-"
    StrCpy $R2 "lang-pl"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "pt"
  ${OrIf} $R4 == "pt-"
    StrCpy $R2 "lang-pt-br"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "ro"
  ${OrIf} $R4 == "ro-"
    StrCpy $R2 "lang-ro"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "ru"
  ${OrIf} $R4 == "ru-"
    StrCpy $R2 "lang-ru"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "zh"
  ${OrIf} $R4 == "zh-"
    StrCpy $R2 "lang-zh-hans"
    Return
  ${EndIf}
FunctionEnd

; $R2 a pack; $R5 its name on the page; $R8 "draft" when its pack is a draft.
Function rutbaLanguageLabel
  StrCpy $R5 $R2
  StrCpy $R8 ""
  ${If} $R2 == "lang-fr"
    StrCpy $R5 "Français · French"
    StrCpy $R8 "draft"
  ${EndIf}
  ${If} $R2 == "lang-de"
    StrCpy $R5 "Deutsch · German"
    StrCpy $R8 "draft"
  ${EndIf}
  ${If} $R2 == "lang-es"
    StrCpy $R5 "Español (España) · Spanish (Spain)"
    StrCpy $R8 "draft"
  ${EndIf}
  ${If} $R2 == "lang-zh-hans"
    StrCpy $R5 "简体中文 · Chinese, Simplified"
    StrCpy $R8 "draft"
  ${EndIf}
  ${If} $R2 == "lang-zh-hant"
    StrCpy $R5 "繁體中文（台灣） · Chinese, Traditional (Taiwan)"
    StrCpy $R8 "draft"
  ${EndIf}
  ${If} $R2 == "lang-it"
    StrCpy $R5 "Italiano · Italian"
    StrCpy $R8 "draft"
  ${EndIf}
  ${If} $R2 == "lang-pt-br"
    StrCpy $R5 "Português (Brasil) · Portuguese (Brazil)"
    StrCpy $R8 "draft"
  ${EndIf}
  ${If} $R2 == "lang-pl"
    StrCpy $R5 "Polski · Polish"
    StrCpy $R8 "draft"
  ${EndIf}
  ${If} $R2 == "lang-nl"
    StrCpy $R5 "Nederlands · Dutch"
    StrCpy $R8 "draft"
  ${EndIf}
  ${If} $R2 == "lang-ru"
    StrCpy $R5 "Русский · Russian"
    StrCpy $R8 "draft"
  ${EndIf}
  ${If} $R2 == "lang-ro"
    StrCpy $R5 "Română · Romanian"
    StrCpy $R8 "draft"
  ${EndIf}
FunctionEnd
!macroend

; A regional pack brings its base: the base is not offered beside it.
!macro rutbaDropRegionalBases VAR
!macroend
