# Arabic glossary for Rutba Office

The words the Arabic catalogue (`ar.js`) uses for the suite's recurring terms,
so the same thing is called the same everywhere. It follows the words people
already know from Microsoft Office and Windows in Arabic, in Modern Standard
Arabic, and plain standard Arabic where Office has no word.

**This translation was drafted by machine and needs a native speaker's
review** before it is called finished. A reviewer changes a word here first,
then in `ar.js`, so the two stay in step.

## How the catalogue is written

- **Commands** take the verbal noun, as Office in Arabic does: حفظ (Save),
  فتح (Open), إغلاق (Close), إدراج (Insert), not the imperative.
- **Sentences** (tooltips, messages, questions) are full Modern Standard
  Arabic, addressing the reader directly as Office does (انقر، اختر، أدخل).
- **Placeholders** such as `{name}` and `{count}` stay exactly as written, in
  braces, untranslated; Arabic word order moves them where Arabic wants them.
- **Keys and shortcuts** stay as they are: Ctrl+S, F7, Shift+F2.
- **File extensions, formats and codes** stay as they are: .docx, PDF, CSV,
  `=SUM(ABOVE)`, `dd/mm/yyyy`.
- **Brand and product names** stay in Latin letters: Rutba Office, Word,
  Excel, PowerPoint, Office, Windows, Google, Microsoft, iCloud, SmartArt,
  WordArt, PivotTable, Power Query.
- **Punctuation**: the Arabic comma ، semicolon ؛ and question mark ؟; a
  sentence ends with a full stop (.); the ellipsis … stays at the end of a
  command that opens a dialog, as in English.
- **Counts** use Arabic's six plural forms, chosen by `Intl.PluralRules`:
  `zero`, `one`, `two`, `few` (3 to 10), `many` (11 to 99) and `other` (100
  and over, and fractions). The noun follows Arabic's number agreement:

  ```js
  '{count} slide': {
    zero: 'لا توجد شرائح',
    one: 'شريحة واحدة',
    two: 'شريحتان',
    few: '{count} شرائح',
    many: '{count} شريحة',
    other: '{count} شريحة',
  }
  ```

  `{count}` may be left out of `zero`, `one` and `two` where Arabic says the
  number in words; every other placeholder is kept in every form.
- **Numbers** are written with the digits 0 to 9 in the catalogue; the window
  shows them in Western or Arabic-Indic digits by the digits setting.

## The apps

| English | Arabic |
| :-- | :-- |
| Rutba Office | Rutba Office |
| Documents (the app) | المستندات |
| Worksheets | أوراق العمل |
| Presentations | العروض التقديمية |
| Mail | البريد |
| Calendar | التقويم |
| Contacts | جهات الاتصال |
| Pictures | الصور |
| Image (the editor) | محرر الصور |
| Video | الفيديو |
| Home (the launcher) | الرئيسية |

## Tabs and frame

| English | Arabic |
| :-- | :-- |
| File | ملف |
| Home | الصفحة الرئيسية |
| Insert | إدراج |
| Draw | رسم |
| Design | تصميم |
| Layout | تخطيط |
| Page Layout | تخطيط الصفحة |
| References | مراجع |
| Mailings | مراسلات |
| Review | مراجعة |
| View | عرض |
| Help | تعليمات |
| Formulas | الصيغ |
| Data | بيانات |
| Automate | أتمتة |
| Transitions | انتقالات |
| Animations | حركات |
| Slide Show | عرض الشرائح |
| Record | تسجيل |
| Table Design | تصميم الجدول |
| Table Layout | تخطيط الجدول |
| Shape Format | تنسيق الشكل |
| Picture Format | تنسيق الصورة |
| Ribbon | الشريط |
| Pane | جزء |
| Status bar | شريط المعلومات |
| Dialog | مربع حوار |
| Settings | الإعدادات |
| Language | اللغة |
| Theme | النسق |

## Commands

| English | Arabic |
| :-- | :-- |
| New | جديد |
| Open | فتح |
| Save | حفظ |
| Save As | حفظ باسم |
| Save as Template | حفظ كقالب |
| Close | إغلاق |
| Print | طباعة |
| Export | تصدير |
| Import | استيراد |
| Share | مشاركة |
| Undo | تراجع |
| Redo | إعادة |
| Cut | قص |
| Copy | نسخ |
| Paste | لصق |
| Paste Special | لصق خاص |
| Delete | حذف |
| Remove | إزالة |
| Rename | إعادة تسمية |
| Select All | تحديد الكل |
| Find | بحث |
| Replace | استبدال |
| Go To | انتقال إلى |
| Insert | إدراج |
| Apply | تطبيق |
| Apply to All | تطبيق على الكل |
| OK | موافق |
| Cancel | إلغاء الأمر |
| Yes / No | نعم / لا |
| Done | تم |
| Next / Previous | التالي / السابق |
| Back | رجوع |
| Refresh | تحديث |
| Update | تحديث |
| Accept / Reject | قبول / رفض |
| Sort | فرز |
| Filter | تصفية |
| Group / Ungroup | تجميع / فك التجميع |
| Merge / Split | دمج / تقسيم |
| Zoom In / Zoom Out | تكبير / تصغير |
| Full screen | ملء الشاشة |
| Minimise / Maximise / Restore | تصغير / تكبير / استعادة |
| Expand / Collapse | توسيع / طي |
| Browse | استعراض |
| Add | إضافة |
| Edit | تحرير |
| Reset | إعادة تعيين |
| Clear | مسح |
| Show / Hide | إظهار / إخفاء |
| More | المزيد |
| None | بلا |
| Automatic | تلقائي |

## Text and documents

| English | Arabic |
| :-- | :-- |
| Document | مستند |
| Page | صفحة |
| Paragraph | فقرة |
| Line | سطر |
| Word (a word) | كلمة |
| Character | حرف |
| Font | الخط |
| Font size | حجم الخط |
| Bold / Italic / Underline | غامق / مائل / تسطير |
| Strikethrough | يتوسطه خط |
| Subscript / Superscript | منخفض / مرتفع |
| Text highlight colour | لون تمييز النص |
| Font colour | لون الخط |
| Colour | لون |
| Style / Styles | نمط / أنماط |
| Heading | عنوان |
| Title | العنوان |
| Bullets / Numbering | تعداد نقطي / تعداد رقمي |
| Multilevel List | قائمة متعددة المستويات |
| List | قائمة |
| Indent | مسافة بادئة |
| Alignment / Align Left / Centre / Align Right / Justify | محاذاة / محاذاة إلى اليسار / توسيط / محاذاة إلى اليمين / ضبط |
| Line spacing | تباعد الأسطر |
| Margins | الهوامش |
| Orientation / Portrait / Landscape | الاتجاه / عمودي / أفقي |
| Size (paper) | الحجم |
| Columns | أعمدة |
| Breaks | فواصل |
| Header / Footer | رأس الصفحة / تذييل الصفحة |
| Footnote / Endnote | حاشية سفلية / تعليق ختامي |
| Table of Contents | جدول المحتويات |
| Caption | تسمية توضيحية |
| Bookmark | إشارة مرجعية |
| Cross-reference | إسناد ترافقي |
| Citation / Bibliography | اقتباس / ثبت المراجع |
| Index | فهرس |
| Comment / Comments | تعليق / تعليقات |
| Track Changes | تعقب التغييرات |
| Spelling | التدقيق الإملائي |
| Thesaurus | قاموس المرادفات |
| Word Count | عدد الكلمات |
| Watermark | علامة مائية |
| Page Colour | لون الصفحة |
| Text Box | مربع نص |
| WordArt | WordArt |
| Equation | معادلة |
| Symbol | رمز |
| Hyperlink / Link | ارتباط تشعبي / ارتباط |
| Mail Merge | دمج المراسلات |
| Envelopes / Labels | مغلفات / تسميات |
| Template | قالب |
| Untitled | بلا عنوان |

## Workbooks

| English | Arabic |
| :-- | :-- |
| Workbook | مصنف |
| Sheet / Worksheet | ورقة / ورقة عمل |
| Cell / Cells | خلية / خلايا |
| Row / Rows | صف / صفوف |
| Column / Columns | عمود / أعمدة |
| Range | نطاق |
| Formula | صيغة |
| Function | دالة |
| Number format | تنسيق الأرقام |
| Currency / Percentage / Date / Time | عملة / نسبة مئوية / تاريخ / وقت |
| Freeze Panes | تجميد الأجزاء |
| Conditional Formatting | التنسيق الشرطي |
| Format as Table | تنسيق كجدول |
| PivotTable | PivotTable |
| Chart | مخطط |
| Sparklines | خطوط المؤشر |
| Slicer | أداة تقسيم البيانات |
| Data Validation | التحقق من صحة البيانات |
| Remove Duplicates | إزالة التكرارات |
| Text to Columns | نص إلى أعمدة |
| Power Query | Power Query |
| Name box | مربع الاسم |
| Formula bar | شريط الصيغة |
| AutoSum | جمع تلقائي |
| Sum / Average / Count / Min / Max | المجموع / المتوسط / العدد / الحد الأدنى / الحد الأقصى |

## Presentations

| English | Arabic |
| :-- | :-- |
| Presentation | عرض تقديمي |
| Slide / Slides | شريحة / شرائح |
| New Slide | شريحة جديدة |
| Layout | تخطيط |
| Section | مقطع |
| Master | رئيسي (Slide Master: الشريحة الرئيسية) |
| Notes | ملاحظات |
| Handout | نشرة |
| Shape / Shapes | شكل / أشكال |
| Picture / Pictures | صورة / صور |
| Fill / Outline | تعبئة / مخطط تفصيلي |
| Shadow / Glow / Reflection | ظل / توهج / انعكاس |
| Transition | انتقال |
| Animation | حركة |
| From Beginning / From Current Slide | من البداية / من الشريحة الحالية |
| Presenter View | طريقة عرض مقدم العرض |
| Rehearse Timings | تمرين على التوقيتات |
| Video / Audio | فيديو / صوت |

## Mail, calendar, contacts, pictures

| English | Arabic |
| :-- | :-- |
| Inbox / Sent / Drafts / Outbox / Junk / Archive / Trash | علبة الوارد / العناصر المرسلة / المسودات / علبة الصادر / البريد غير الهام / الأرشيف / المحذوفات |
| Message | رسالة |
| New message | رسالة جديدة |
| Reply / Reply All / Forward | رد / رد على الكل / إعادة توجيه |
| Send | إرسال |
| Attachment | مرفق |
| Account | حساب |
| Folder | مجلد |
| Search | بحث |
| Unread / Flagged | غير مقروءة / موسومة بعلامة |
| Rules | قواعد |
| Signature | توقيع |
| Event | حدث |
| Meeting | اجتماع |
| Today / Day / Week / Month / Agenda | اليوم / يوم / أسبوع / شهر / جدول الأعمال |
| Repeat | تكرار |
| Reminder | تذكير |
| Contact | جهة اتصال |
| Address book | دفتر العناوين |
| Phone / Email / Address | هاتف / بريد إلكتروني / عنوان |
| Rotate / Crop / Flip | تدوير / اقتصاص / قلب |
| Rating / Tags | تقييم / علامات |
| Slideshow | عرض شرائح |
| Brightness / Contrast | السطوع / التباين |
| Trim / Split | اقتطاع / تقسيم |
| Export video | تصدير الفيديو |

## Errors and plain sentences

Error messages and tooltips are full sentences: translate the meaning, in
clear standard Arabic, not word for word. Keep a sentence's facts (a number,
a name in `{braces}`, a key) exactly. Prefer "تعذر فتح هذا الملف" over a
literal rendering.
