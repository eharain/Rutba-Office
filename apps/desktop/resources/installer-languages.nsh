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
  StrCpy $R4 $R1 7
  ${If} $R1 == "es-419"
  ${OrIf} $R4 == "es-419-"
    StrCpy $R2 "lang-es-419"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "en-us"
  ${OrIf} $R4 == "en-us-"
    StrCpy $R2 "lang-en-us"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "es-ar"
  ${OrIf} $R4 == "es-ar-"
    StrCpy $R2 "lang-es-419"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "es-bo"
  ${OrIf} $R4 == "es-bo-"
    StrCpy $R2 "lang-es-419"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "es-cl"
  ${OrIf} $R4 == "es-cl-"
    StrCpy $R2 "lang-es-419"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "es-co"
  ${OrIf} $R4 == "es-co-"
    StrCpy $R2 "lang-es-419"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "es-cr"
  ${OrIf} $R4 == "es-cr-"
    StrCpy $R2 "lang-es-419"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "es-cu"
  ${OrIf} $R4 == "es-cu-"
    StrCpy $R2 "lang-es-419"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "es-do"
  ${OrIf} $R4 == "es-do-"
    StrCpy $R2 "lang-es-419"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "es-ec"
  ${OrIf} $R4 == "es-ec-"
    StrCpy $R2 "lang-es-419"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "es-gt"
  ${OrIf} $R4 == "es-gt-"
    StrCpy $R2 "lang-es-419"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "es-hn"
  ${OrIf} $R4 == "es-hn-"
    StrCpy $R2 "lang-es-419"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "es-mx"
  ${OrIf} $R4 == "es-mx-"
    StrCpy $R2 "lang-es-419"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "es-ni"
  ${OrIf} $R4 == "es-ni-"
    StrCpy $R2 "lang-es-419"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "es-pa"
  ${OrIf} $R4 == "es-pa-"
    StrCpy $R2 "lang-es-419"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "es-pe"
  ${OrIf} $R4 == "es-pe-"
    StrCpy $R2 "lang-es-419"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "es-pr"
  ${OrIf} $R4 == "es-pr-"
    StrCpy $R2 "lang-es-419"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "es-py"
  ${OrIf} $R4 == "es-py-"
    StrCpy $R2 "lang-es-419"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "es-sv"
  ${OrIf} $R4 == "es-sv-"
    StrCpy $R2 "lang-es-419"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "es-us"
  ${OrIf} $R4 == "es-us-"
    StrCpy $R2 "lang-es-419"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "es-uy"
  ${OrIf} $R4 == "es-uy-"
    StrCpy $R2 "lang-es-419"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "es-ve"
  ${OrIf} $R4 == "es-ve-"
    StrCpy $R2 "lang-es-419"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "fr-ca"
  ${OrIf} $R4 == "fr-ca-"
    StrCpy $R2 "lang-fr-ca"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "pt-ao"
  ${OrIf} $R4 == "pt-ao-"
    StrCpy $R2 "lang-pt-pt"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "pt-br"
  ${OrIf} $R4 == "pt-br-"
    StrCpy $R2 "lang-pt-br"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "pt-ch"
  ${OrIf} $R4 == "pt-ch-"
    StrCpy $R2 "lang-pt-pt"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "pt-cv"
  ${OrIf} $R4 == "pt-cv-"
    StrCpy $R2 "lang-pt-pt"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "pt-gq"
  ${OrIf} $R4 == "pt-gq-"
    StrCpy $R2 "lang-pt-pt"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "pt-gw"
  ${OrIf} $R4 == "pt-gw-"
    StrCpy $R2 "lang-pt-pt"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "pt-lu"
  ${OrIf} $R4 == "pt-lu-"
    StrCpy $R2 "lang-pt-pt"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "pt-mo"
  ${OrIf} $R4 == "pt-mo-"
    StrCpy $R2 "lang-pt-pt"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "pt-mz"
  ${OrIf} $R4 == "pt-mz-"
    StrCpy $R2 "lang-pt-pt"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "pt-pt"
  ${OrIf} $R4 == "pt-pt-"
    StrCpy $R2 "lang-pt-pt"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "pt-st"
  ${OrIf} $R4 == "pt-st-"
    StrCpy $R2 "lang-pt-pt"
    Return
  ${EndIf}
  StrCpy $R4 $R1 6
  ${If} $R1 == "pt-tl"
  ${OrIf} $R4 == "pt-tl-"
    StrCpy $R2 "lang-pt-pt"
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
  ${If} $R1 == "ar"
  ${OrIf} $R4 == "ar-"
    StrCpy $R2 "lang-ar"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "bg"
  ${OrIf} $R4 == "bg-"
    StrCpy $R2 "lang-bg"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "cs"
  ${OrIf} $R4 == "cs-"
    StrCpy $R2 "lang-cs"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "da"
  ${OrIf} $R4 == "da-"
    StrCpy $R2 "lang-da"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "de"
  ${OrIf} $R4 == "de-"
    StrCpy $R2 "lang-de"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "el"
  ${OrIf} $R4 == "el-"
    StrCpy $R2 "lang-el"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "es"
  ${OrIf} $R4 == "es-"
    StrCpy $R2 "lang-es"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "et"
  ${OrIf} $R4 == "et-"
    StrCpy $R2 "lang-et"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "fi"
  ${OrIf} $R4 == "fi-"
    StrCpy $R2 "lang-fi"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "fr"
  ${OrIf} $R4 == "fr-"
    StrCpy $R2 "lang-fr"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "ga"
  ${OrIf} $R4 == "ga-"
    StrCpy $R2 "lang-ga"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "he"
  ${OrIf} $R4 == "he-"
    StrCpy $R2 "lang-he"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "hi"
  ${OrIf} $R4 == "hi-"
    StrCpy $R2 "lang-hi"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "hr"
  ${OrIf} $R4 == "hr-"
    StrCpy $R2 "lang-hr"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "hu"
  ${OrIf} $R4 == "hu-"
    StrCpy $R2 "lang-hu"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "it"
  ${OrIf} $R4 == "it-"
    StrCpy $R2 "lang-it"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "lt"
  ${OrIf} $R4 == "lt-"
    StrCpy $R2 "lang-lt"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "lv"
  ${OrIf} $R4 == "lv-"
    StrCpy $R2 "lang-lv"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "mt"
  ${OrIf} $R4 == "mt-"
    StrCpy $R2 "lang-mt"
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
  ${If} $R1 == "sk"
  ${OrIf} $R4 == "sk-"
    StrCpy $R2 "lang-sk"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "sl"
  ${OrIf} $R4 == "sl-"
    StrCpy $R2 "lang-sl"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "sv"
  ${OrIf} $R4 == "sv-"
    StrCpy $R2 "lang-sv"
    Return
  ${EndIf}
  StrCpy $R4 $R1 3
  ${If} $R1 == "ur"
  ${OrIf} $R4 == "ur-"
    StrCpy $R2 "lang-ur"
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
  ${If} $R2 == "lang-ur"
    StrCpy $R5 "اردو · Urdu"
    StrCpy $R8 "draft"
  ${EndIf}
  ${If} $R2 == "lang-ar"
    StrCpy $R5 "العربية · Arabic"
    StrCpy $R8 "draft"
  ${EndIf}
  ${If} $R2 == "lang-hi"
    StrCpy $R5 "हिन्दी · Hindi"
    StrCpy $R8 "draft"
  ${EndIf}
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
  ${If} $R2 == "lang-cs"
    StrCpy $R5 "Čeština · Czech"
  ${EndIf}
  ${If} $R2 == "lang-el"
    StrCpy $R5 "Ελληνικά · Greek"
  ${EndIf}
  ${If} $R2 == "lang-hu"
    StrCpy $R5 "Magyar · Hungarian"
  ${EndIf}
  ${If} $R2 == "lang-sv"
    StrCpy $R5 "Svenska · Swedish"
  ${EndIf}
  ${If} $R2 == "lang-he"
    StrCpy $R5 "עברית · Hebrew"
  ${EndIf}
  ${If} $R2 == "lang-bg"
    StrCpy $R5 "Български · Bulgarian"
  ${EndIf}
  ${If} $R2 == "lang-da"
    StrCpy $R5 "Dansk · Danish"
  ${EndIf}
  ${If} $R2 == "lang-fi"
    StrCpy $R5 "Suomi · Finnish"
  ${EndIf}
  ${If} $R2 == "lang-sk"
    StrCpy $R5 "Slovenčina · Slovak"
  ${EndIf}
  ${If} $R2 == "lang-hr"
    StrCpy $R5 "Hrvatski · Croatian"
  ${EndIf}
  ${If} $R2 == "lang-lt"
    StrCpy $R5 "Lietuvių · Lithuanian"
  ${EndIf}
  ${If} $R2 == "lang-sl"
    StrCpy $R5 "Slovenščina · Slovenian"
  ${EndIf}
  ${If} $R2 == "lang-lv"
    StrCpy $R5 "Latviešu · Latvian"
  ${EndIf}
  ${If} $R2 == "lang-et"
    StrCpy $R5 "Eesti · Estonian"
  ${EndIf}
  ${If} $R2 == "lang-ga"
    StrCpy $R5 "Gaeilge · Irish"
  ${EndIf}
  ${If} $R2 == "lang-mt"
    StrCpy $R5 "Malti · Maltese"
  ${EndIf}
  ${If} $R2 == "lang-pt-pt"
    StrCpy $R5 "Português (Portugal) · Portuguese (Portugal)"
  ${EndIf}
  ${If} $R2 == "lang-es-419"
    StrCpy $R5 "Español (Latinoamérica) · Spanish (Latin America)"
  ${EndIf}
  ${If} $R2 == "lang-fr-ca"
    StrCpy $R5 "Français (Canada) · French (Canada)"
  ${EndIf}
  ${If} $R2 == "lang-en-us"
    StrCpy $R5 "English (US)"
  ${EndIf}
FunctionEnd
!macroend

; A regional pack brings its base: the base is not offered beside it.
!macro rutbaDropRegionalBases VAR
  ${WordFind} "${VAR}" ",lang-pt-pt," "E*" $R6
  ${IfNot} ${Errors}
    ${WordReplace} "${VAR}" ",lang-pt-br," "," "+" ${VAR}
  ${EndIf}
  ClearErrors
  ${WordFind} "${VAR}" ",lang-es-419," "E*" $R6
  ${IfNot} ${Errors}
    ${WordReplace} "${VAR}" ",lang-es," "," "+" ${VAR}
  ${EndIf}
  ClearErrors
  ${WordFind} "${VAR}" ",lang-fr-ca," "E*" $R6
  ${IfNot} ${Errors}
    ${WordReplace} "${VAR}" ",lang-fr," "," "+" ${VAR}
  ${EndIf}
  ClearErrors
!macroend
