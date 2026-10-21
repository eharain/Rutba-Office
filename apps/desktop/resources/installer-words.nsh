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
  StrCpy $rutbaW_base "Databases"
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
    StrCpy $rutbaW_base "قواعد البيانات"
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
    StrCpy $rutbaW_base "ڈیٹابیسز"
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
    StrCpy $rutbaW_base "डेटाबेस"
  ${EndIf}
FunctionEnd
!macroend
