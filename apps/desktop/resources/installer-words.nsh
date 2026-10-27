; The Choose apps page in the suite's languages. Made by tools/make-installer-words.mjs
; from the catalogues: change a word there and make this again.

!macro rutbaInstallerWords
  Var rutbaW_title
  Var rutbaW_sub
  Var rutbaW_lead
  Var rutbaW_office
  Var rutbaW_more
  Var rutbaW_addons
  Var rutbaW_ai
  Var rutbaW_none
  Var rutbaW_mail
  Var rutbaW_calendar
  Var rutbaW_contacts
  Var rutbaW_word
  Var rutbaW_sheets
  Var rutbaW_slides
  Var rutbaW_pdf
  Var rutbaW_pictures
  Var rutbaW_image
  Var rutbaW_video
  Var rutbaW_recorder
  Var rutbaW_talk
  Var rutbaW_base
  Var rutbaW_ocr
  Var rutbaW_ptitle
  Var rutbaW_psub
  Var rutbaW_plead

Function rutbaPickWords
  StrCpy $rutbaW_title "Choose apps"
  StrCpy $rutbaW_sub "The apps you want to see and to open your files."
  StrCpy $rutbaW_lead "Tick the apps and add-ons you want. Any of them can be added or taken away later, from Choose apps in Rutba Office."
  StrCpy $rutbaW_office "Office apps"
  StrCpy $rutbaW_more "More apps"
  StrCpy $rutbaW_addons "Add-ons"
  StrCpy $rutbaW_ai "AI Tools"
  StrCpy $rutbaW_none "Choose one app at least."
  StrCpy $rutbaW_mail "Mail"
  StrCpy $rutbaW_calendar "Calendar"
  StrCpy $rutbaW_contacts "Contacts"
  StrCpy $rutbaW_word "Documents"
  StrCpy $rutbaW_sheets "Worksheets"
  StrCpy $rutbaW_slides "Presentations"
  StrCpy $rutbaW_pdf "PDF"
  StrCpy $rutbaW_pictures "Pictures"
  StrCpy $rutbaW_image "Image"
  StrCpy $rutbaW_video "Video"
  StrCpy $rutbaW_recorder "Recorder"
  StrCpy $rutbaW_talk "Talk"
  StrCpy $rutbaW_base "Data Files"
  StrCpy $rutbaW_ocr "Text recognition"
  StrCpy $rutbaW_ptitle "Apps to install"
  StrCpy $rutbaW_psub "Only the apps you tick are put on this computer."
  StrCpy $rutbaW_plead "Tick the apps to install. An app not ticked is not on this computer at all; add it later from Add or remove apps in Rutba Office."
  ; The installer's own language, or, where it fell back to English, the system's.
  IntOp $0 $LANGUAGE & 0x3FF
  ${If} $0 == 9
    System::Call 'kernel32::GetUserDefaultUILanguage() i .r1'
    IntOp $0 $1 & 0x3FF
  ${EndIf}
  ${If} $0 == 1 ; Arabic
    StrCpy $rutbaW_title "اختيار التطبيقات"
    StrCpy $rutbaW_sub "التطبيقات التي تريد أن تراها وأن تفتح ملفاتك."
    StrCpy $rutbaW_lead "حدّد التطبيقات والإضافات التي تريدها. يمكن إضافة أيٍّ منها أو إزالته لاحقًا من «اختيار التطبيقات» في Rutba Office."
    StrCpy $rutbaW_office "تطبيقات أوفيس"
    StrCpy $rutbaW_more "تطبيقات أخرى"
    StrCpy $rutbaW_addons "الإضافات"
    StrCpy $rutbaW_ai "أدوات الذكاء الاصطناعي"
    StrCpy $rutbaW_none "اختر تطبيقًا واحدًا على الأقل."
    StrCpy $rutbaW_mail "البريد"
    StrCpy $rutbaW_calendar "التقويم"
    StrCpy $rutbaW_contacts "جهات الاتصال"
    StrCpy $rutbaW_word "المستندات"
    StrCpy $rutbaW_sheets "أوراق العمل"
    StrCpy $rutbaW_slides "العروض التقديمية"
    StrCpy $rutbaW_pdf "ملف PDF"
    StrCpy $rutbaW_pictures "الصور"
    StrCpy $rutbaW_image "محرر الصور"
    StrCpy $rutbaW_video "فيديو"
    StrCpy $rutbaW_recorder "المسجّل"
    StrCpy $rutbaW_talk "توك"
    StrCpy $rutbaW_base "ملفات البيانات"
    StrCpy $rutbaW_ocr "التعرف على النص"
    StrCpy $rutbaW_ptitle "التطبيقات المراد تثبيتها"
    StrCpy $rutbaW_psub "لا يوضع على هذا الحاسوب إلا التطبيقات التي تحددها."
    StrCpy $rutbaW_plead "حدد التطبيقات المراد تثبيتها. التطبيق غير المحدد لا يوضع على هذا الحاسوب إطلاقاً؛ أضفه لاحقاً من «إضافة التطبيقات أو إزالتها» في Rutba Office."
  ${ElseIf} $0 == 32 ; Urdu
    StrCpy $rutbaW_title "ایپس چنیں"
    StrCpy $rutbaW_sub "وہ ایپس جو آپ دیکھنا اور اپنی فائلیں کھولنے کے لیے چاہتے ہیں۔"
    StrCpy $rutbaW_lead "جو ایپس اور ایڈ آنز آپ چاہتے ہیں ان پر نشان لگائیں۔ ان میں سے کوئی بھی بعد میں Rutba Office میں «ایپس چنیں» سے شامل کی یا ہٹائی جا سکتی ہے۔"
    StrCpy $rutbaW_office "آفس ایپس"
    StrCpy $rutbaW_more "مزید ایپس"
    StrCpy $rutbaW_addons "اضافے"
    StrCpy $rutbaW_ai "AI ٹولز"
    StrCpy $rutbaW_none "کم از کم ایک ایپ چنیں۔"
    StrCpy $rutbaW_mail "میل"
    StrCpy $rutbaW_calendar "کیلنڈر"
    StrCpy $rutbaW_contacts "روابط"
    StrCpy $rutbaW_word "دستاویزات"
    StrCpy $rutbaW_sheets "ورک شیٹس"
    StrCpy $rutbaW_slides "پریزنٹیشنز"
    StrCpy $rutbaW_pdf "پی ڈی ایف"
    StrCpy $rutbaW_pictures "تصاویر"
    StrCpy $rutbaW_image "امیج"
    StrCpy $rutbaW_video "ویڈیو"
    StrCpy $rutbaW_recorder "ریکارڈر"
    StrCpy $rutbaW_talk "ٹاک"
    StrCpy $rutbaW_base "ڈیٹا فائلز"
    StrCpy $rutbaW_ocr "متن کی پہچان"
    StrCpy $rutbaW_ptitle "انسٹال کرنے کے لیے ایپس"
    StrCpy $rutbaW_psub "صرف وہ ایپس اس کمپیوٹر پر رکھی جاتی ہیں جن پر آپ نشان لگائیں۔"
    StrCpy $rutbaW_plead "انسٹال کرنے کے لیے ایپس پر نشان لگائیں۔ جس ایپ پر نشان نہ ہو وہ اس کمپیوٹر پر بالکل نہیں ہوتی؛ اسے بعد میں Rutba Office میں «ایپس شامل کریں یا ہٹائیں» سے شامل کریں۔"
  ${ElseIf} $0 == 57 ; Hindi
    StrCpy $rutbaW_title "ऐप चुनें"
    StrCpy $rutbaW_sub "वे ऐप जिन्हें आप देखना और अपनी फ़ाइलें खोलने के लिए चाहते हैं।"
    StrCpy $rutbaW_lead "जो ऐप और ऐड-ऑन आप चाहते हैं उन पर निशान लगाएँ। इनमें से कोई भी बाद में Rutba Office में «ऐप चुनें» से जोड़ा या हटाया जा सकता है।"
    StrCpy $rutbaW_office "Office ऐप"
    StrCpy $rutbaW_more "और ऐप"
    StrCpy $rutbaW_addons "ऐड-ऑन"
    StrCpy $rutbaW_ai "AI टूल"
    StrCpy $rutbaW_none "कम से कम एक ऐप चुनें।"
    StrCpy $rutbaW_mail "मेल"
    StrCpy $rutbaW_calendar "कैलेंडर"
    StrCpy $rutbaW_contacts "संपर्क"
    StrCpy $rutbaW_word "दस्तावेज़"
    StrCpy $rutbaW_sheets "कार्यपत्रक"
    StrCpy $rutbaW_slides "प्रस्तुतियाँ"
    StrCpy $rutbaW_pdf "पीडीएफ़"
    StrCpy $rutbaW_pictures "चित्र"
    StrCpy $rutbaW_image "इमेज"
    StrCpy $rutbaW_video "वीडियो"
    StrCpy $rutbaW_recorder "रिकॉर्डर"
    StrCpy $rutbaW_talk "टॉक"
    StrCpy $rutbaW_base "डेटा फ़ाइलें"
    StrCpy $rutbaW_ocr "टेक्स्ट पहचान"
    StrCpy $rutbaW_ptitle "इंस्टॉल करने के लिए ऐप"
    StrCpy $rutbaW_psub "इस कंप्यूटर पर केवल वही ऐप रखे जाते हैं जिन पर आप निशान लगाएँ।"
    StrCpy $rutbaW_plead "इंस्टॉल करने के लिए ऐप पर निशान लगाएँ। जिस ऐप पर निशान न हो वह इस कंप्यूटर पर बिल्कुल नहीं होता; उसे बाद में Rutba Office में «ऐप जोड़ें या हटाएँ» से जोड़ें।"
  ${EndIf}
FunctionEnd
!macroend
