// Invoices, receipts, letters and forms with known fields, for measuring how
// well the fields of a document are read from its picture
// (packages/ocr/src/fields.js). tools/docread-corpus.js draws them with the
// suite's renderer; this module only decides what each page holds and writes
// its HTML, so the same seed always makes the same page.
//
// The invoice layouts follow the gallery's two invoice templates
// (packages/templates: word-invoice, sheet-invoice): the Word one's heading
// table, its bill to / number / date block and its items table ending in the
// subtotal, the tax and the total due; the Worksheets one's labels in one
// column with their values beside them. The receipt is a till roll. Letters
// and forms carry no amounts and are there for the document-type guess and
// the tick boxes.

/** A small seeded generator (mulberry32): the same seed, the same page. */
export function seeded(seed) {
  let a = seed >>> 0;
  const next = () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
  const rng = {
    next,
    int: (lo, hi) => lo + Math.floor(next() * (hi - lo + 1)),
    pick: (list) => list[Math.floor(next() * list.length)],
    chance: (p) => next() < p,
    shuffle: (list) => {
      const out = [...list];
      for (let i = out.length - 1; i > 0; i--) {
        const j = Math.floor(next() * (i + 1));
        [out[i], out[j]] = [out[j], out[i]];
      }
      return out;
    },
  };
  return rng;
}

const DIGIT_SETS = {
  latin: '0123456789',
  arabic: '٠١٢٣٤٥٦٧٨٩',
  urdu: '۰۱۲۳۴۵۶۷۸۹',
  devanagari: '०१२३४५६७८९',
};

/** Western digits written in another set. */
export function inDigits(text, set = 'latin') {
  if (set === 'latin') return text;
  const d = DIGIT_SETS[set];
  return text.replace(/[0-9]/g, (c) => d[c.charCodeAt(0) - 48]);
}

/**
 * Each language: what its documents say, in which faces, in which money, and
 * how it writes numbers and dates. A label has several wordings and a page
 * picks one, so the reader is not measured on one phrase.
 */
export const LANGS = {
  eng: {
    dir: 'ltr',
    fonts: ['Arial', '"Times New Roman"', 'Calibri', 'Georgia'],
    mono: ['Consolas', '"Courier New"'],
    title: { invoice: ['INVOICE', 'Invoice', 'TAX INVOICE'], receipt: ['RECEIPT', 'Sales Receipt'] },
    labels: {
      number: ['Invoice number', 'Invoice No.', 'Invoice #'],
      receiptNumber: ['Receipt No.', 'Receipt number'],
      date: ['Date', 'Invoice date', 'Issue date'],
      dueDate: ['Payment due', 'Due date'],
      billTo: ['Bill to', 'Billed to'],
      description: ['Description', 'Item'],
      quantity: ['Quantity', 'Qty'],
      unitPrice: ['Unit price', 'Price'],
      amount: ['Amount'],
      subtotal: ['Subtotal', 'Sub-total'],
      tax: ['VAT at {rate}', 'VAT', 'Tax', 'Sales tax {rate}'],
      taxRate: ['VAT rate', 'Tax rate'],
      total: ['Total due', 'Total', 'Amount due', 'Grand total'],
      cash: ['Cash', 'Card'],
      change: ['Change'],
      thanks: ['Thank you for your business.', 'Thank you, please come again'],
      howToPay: 'How to pay',
      payText: 'Bank transfer to {supplier}, account 00000000. Please give the invoice number as the reference.',
    },
    suppliers: ['Northgate Supplies Ltd', 'Harbour Print Co.', 'Greenfield Office Ltd', 'Bluebell Catering', 'Ashworth and Sons', 'Kestrel Stationers', 'Meadow Lane Bakery', 'Copperline Electrical'],
    customers: ['Riverside Dental Clinic', 'Oakwood Primary School', 'Jane Patel', 'Linden Architects'],
    addresses: ['12 High Street, Leeds LS1 4AB', 'Unit 4, Mill Road, Bristol BS2 9QT', '88 Queen Street, Glasgow G1 3DN', '5 Station Approach, York YO1 6HT'],
    items: ['Printer paper, A4 box', 'Ink cartridge, black', 'Desk lamp', 'Delivery', 'Design consultation (hours)', 'Website pages', 'Coffee beans, 1 kg', 'Folders, pack of 10', 'USB cable', 'Cleaning service'],
    currencies: [{ sym: '£', code: 'GBP' }, { sym: '$', code: 'USD' }, { sym: '€', code: 'EUR' }],
    decimals: 2,
    taxRates: [20, 5],
    digits: ['latin'],
    dates: ['d MMM yyyy', 'dd/mm/yyyy', 'yyyy-mm-dd', 'MMMM d, yyyy'],
    months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
    letter: { greeting: 'Dear {name},', closing: 'Yours sincerely,', subject: 'Re: your request of last month' },
    form: {
      title: ['Application form', 'Registration form'],
      fields: ['Name', 'Address', 'Date of birth', 'Phone', 'Email', 'Signature'],
      groups: [['Type of request', ['New', 'Renewal', 'Change']], ['Contact by', ['Email', 'Phone', 'Post']]],
    },
  },
  ara: {
    dir: 'rtl',
    fonts: ['NotoNaskh', 'Tahoma', 'Arial'],
    title: { invoice: ['فاتورة', 'فاتورة ضريبية'], receipt: ['إيصال', 'إيصال بيع'] },
    labels: {
      number: ['رقم الفاتورة'],
      receiptNumber: ['رقم الإيصال'],
      date: ['التاريخ', 'تاريخ الفاتورة', 'تاريخ الإصدار'],
      dueDate: ['تاريخ الاستحقاق', 'آخر موعد للدفع'],
      billTo: ['فاتورة إلى', 'العميل'],
      description: ['الوصف', 'البند'],
      quantity: ['الكمية'],
      unitPrice: ['سعر الوحدة', 'السعر'],
      amount: ['المبلغ'],
      subtotal: ['المجموع الفرعي'],
      tax: ['ضريبة القيمة المضافة {rate}', 'الضريبة'],
      taxRate: ['نسبة الضريبة'],
      total: ['الإجمالي', 'المبلغ المستحق', 'المجموع الكلي'],
      cash: ['نقدا', 'بطاقة'],
      change: ['الباقي'],
      thanks: ['شكرا لتعاملكم معنا', 'شكرا لزيارتكم'],
      howToPay: 'طريقة الدفع',
      payText: 'حوالة مصرفية إلى {supplier}، الحساب 00000000، مع ذكر رقم الفاتورة.',
    },
    suppliers: ['مؤسسة النخيل التجارية', 'شركة الأفق للقرطاسية', 'مطبعة الرياض الحديثة', 'شركة الواحة للتموين', 'مكتبة دار المعرفة', 'شركة السلام للإلكترونيات'],
    customers: ['عيادة الشفاء', 'مدرسة الأمل', 'شركة البناء المتحد', 'محمد العلي'],
    addresses: ['شارع الملك فهد، الرياض', 'طريق الشيخ زايد، دبي', 'شارع الجامعة، عمان', 'شارع النيل، القاهرة'],
    items: ['ورق طباعة', 'حبر أسود', 'مصباح مكتب', 'خدمة التوصيل', 'ملفات ورقية', 'قهوة عربية', 'كابل شبكة', 'خدمة تنظيف', 'دفاتر ملاحظات', 'أقلام حبر'],
    currencies: [{ sym: 'ر.س', code: 'SAR', after: true }, { sym: 'د.إ', code: 'AED', after: true }, { sym: 'SAR', code: 'SAR' }],
    decimals: 2,
    taxRates: [15, 5],
    digits: ['latin', 'arabic'],
    dates: ['dd/mm/yyyy', 'd MMMM yyyy', 'yyyy/mm/dd'],
    months: ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'],
    letter: { greeting: 'السيد {name} المحترم،', closing: 'وتفضلوا بقبول فائق الاحترام،', subject: 'الموضوع: طلبكم في الشهر الماضي' },
    form: {
      title: ['استمارة طلب', 'نموذج تسجيل'],
      fields: ['الاسم', 'العنوان', 'تاريخ الميلاد', 'الهاتف', 'البريد الإلكتروني', 'التوقيع'],
      groups: [['نوع الطلب', ['جديد', 'تجديد', 'تعديل']], ['التواصل عبر', ['البريد الإلكتروني', 'الهاتف', 'البريد']]],
    },
  },
  urd: {
    dir: 'rtl',
    fonts: ['NotoNaskh', 'Tahoma', 'NotoNastaliq'],
    title: { invoice: ['انوائس', 'بل'], receipt: ['رسید'] },
    labels: {
      number: ['انوائس نمبر', 'بل نمبر'],
      receiptNumber: ['رسید نمبر'],
      date: ['تاریخ', 'تاریخ اجرا'],
      dueDate: ['ادائیگی کی آخری تاریخ', 'آخری تاریخ'],
      billTo: ['بنام', 'گاہک'],
      description: ['تفصیل'],
      quantity: ['تعداد'],
      unitPrice: ['فی عدد قیمت', 'قیمت'],
      amount: ['رقم'],
      subtotal: ['ذیلی میزان', 'ذیلی کل'],
      tax: ['سیلز ٹیکس {rate}', 'ٹیکس'],
      taxRate: ['ٹیکس کی شرح'],
      total: ['کل رقم', 'واجب الادا رقم', 'کل'],
      cash: ['نقد'],
      change: ['بقایا'],
      thanks: ['خریداری کا شکریہ'],
      howToPay: 'ادائیگی کا طریقہ',
      payText: '{supplier} کے اکاؤنٹ 00000000 میں بینک ٹرانسفر کریں اور حوالے میں بل نمبر لکھیں۔',
    },
    suppliers: ['الفلاح ٹریڈرز', 'نور اسٹیشنری', 'کراچی پرنٹنگ پریس', 'مدینہ کیٹرنگ', 'شاہین الیکٹرونکس', 'رحمان بک ڈپو'],
    customers: ['شفا کلینک', 'گورنمنٹ ہائی اسکول', 'احمد علی', 'سٹی آرکیٹیکٹس'],
    addresses: ['مال روڈ، لاہور', 'شاہراہ فیصل، کراچی', 'بلیو ایریا، اسلام آباد', 'صدر بازار، پشاور'],
    items: ['پرنٹنگ کاغذ', 'سیاہ سیاہی', 'میز کا لیمپ', 'ترسیل', 'فائلیں', 'چائے کی پتی', 'نیٹ ورک کیبل', 'صفائی کی خدمت', 'کاپیاں', 'بال پوائنٹ قلم'],
    currencies: [{ sym: 'Rs', code: 'PKR' }, { sym: 'روپے', code: 'PKR', after: true }],
    decimals: 2,
    taxRates: [17, 18, 16],
    digits: ['latin', 'urdu'],
    dates: ['dd/mm/yyyy', 'd MMMM yyyy'],
    months: ['جنوری', 'فروری', 'مارچ', 'اپریل', 'مئی', 'جون', 'جولائی', 'اگست', 'ستمبر', 'اکتوبر', 'نومبر', 'دسمبر'],
    letter: { greeting: 'محترم {name} صاحب،', closing: 'آپ کا مخلص،', subject: 'موضوع: پچھلے مہینے کی درخواست' },
    form: {
      title: ['درخواست فارم', 'رجسٹریشن فارم'],
      fields: ['نام', 'پتہ', 'تاریخ پیدائش', 'فون', 'ای میل', 'دستخط'],
      groups: [['درخواست کی قسم', ['نئی', 'تجدید', 'تبدیلی']], ['رابطہ', ['ای میل', 'فون', 'ڈاک']]],
    },
  },
  hin: {
    dir: 'ltr',
    fonts: ['NotoDevanagari', '"Nirmala UI"'],
    title: { invoice: ['बीजक', 'टैक्स इनवॉइस', 'चालान'], receipt: ['रसीद'] },
    labels: {
      number: ['बीजक संख्या', 'इनवॉइस नंबर', 'चालान संख्या'],
      receiptNumber: ['रसीद संख्या'],
      date: ['दिनांक', 'तारीख'],
      dueDate: ['देय तिथि', 'भुगतान की अंतिम तिथि'],
      billTo: ['प्राप्तकर्ता', 'ग्राहक'],
      description: ['विवरण'],
      quantity: ['मात्रा'],
      unitPrice: ['दर', 'इकाई मूल्य'],
      amount: ['राशि'],
      subtotal: ['उप-योग', 'उप योग'],
      tax: ['जीएसटी {rate}', 'कर'],
      taxRate: ['कर की दर'],
      total: ['कुल राशि', 'कुल देय', 'कुल'],
      cash: ['नकद'],
      change: ['वापसी'],
      thanks: ['खरीदारी के लिए धन्यवाद'],
      howToPay: 'भुगतान कैसे करें',
      payText: '{supplier} के खाते 00000000 में बैंक ट्रांसफ़र करें और संदर्भ में बीजक संख्या लिखें।',
    },
    suppliers: ['शर्मा ट्रेडर्स', 'गणेश स्टेशनरी', 'नवभारत प्रिंटर्स', 'अन्नपूर्णा कैटरिंग', 'सूर्या इलेक्ट्रॉनिक्स', 'विद्या बुक डिपो'],
    customers: ['आरोग्य क्लिनिक', 'सरस्वती विद्यालय', 'राहुल वर्मा', 'नगर वास्तुकार'],
    addresses: ['महात्मा गांधी मार्ग, पुणे', 'कनॉट प्लेस, नई दिल्ली', 'एम जी रोड, बेंगलुरु', 'हज़रतगंज, लखनऊ'],
    items: ['छपाई का कागज़', 'काली स्याही', 'टेबल लैंप', 'डिलीवरी', 'फ़ाइलें', 'चाय पत्ती', 'नेटवर्क केबल', 'सफ़ाई सेवा', 'कॉपियाँ', 'पेन'],
    currencies: [{ sym: '₹', code: 'INR' }, { sym: 'रु', code: 'INR' }],
    decimals: 2,
    taxRates: [18, 12, 5],
    digits: ['latin', 'devanagari'],
    grouping: 'indian',
    dates: ['dd/mm/yyyy', 'd MMMM yyyy', 'dd-mm-yyyy'],
    months: ['जनवरी', 'फ़रवरी', 'मार्च', 'अप्रैल', 'मई', 'जून', 'जुलाई', 'अगस्त', 'सितंबर', 'अक्टूबर', 'नवंबर', 'दिसंबर'],
    letter: { greeting: 'प्रिय {name},', closing: 'भवदीय,', subject: 'विषय: पिछले महीने का अनुरोध' },
    form: {
      title: ['आवेदन पत्र', 'पंजीकरण फ़ॉर्म'],
      fields: ['नाम', 'पता', 'जन्म तिथि', 'फ़ोन', 'ईमेल', 'हस्ताक्षर'],
      groups: [['अनुरोध का प्रकार', ['नया', 'नवीनीकरण', 'बदलाव']], ['संपर्क', ['ईमेल', 'फ़ोन', 'डाक']]],
    },
  },
  chi_sim: {
    dir: 'ltr',
    fonts: ['SimSun', '"Microsoft YaHei"'],
    title: { invoice: ['发票', '销售发票'], receipt: ['收据', '购物小票'] },
    labels: {
      number: ['发票号码', '发票编号'],
      receiptNumber: ['单号', '收据编号'],
      date: ['日期', '开票日期'],
      dueDate: ['付款期限', '到期日'],
      billTo: ['购买方', '客户'],
      description: ['项目', '品名'],
      quantity: ['数量'],
      unitPrice: ['单价'],
      amount: ['金额'],
      subtotal: ['小计'],
      tax: ['增值税 {rate}', '税额'],
      taxRate: ['税率'],
      total: ['合计', '总计', '应付金额'],
      cash: ['现金'],
      change: ['找零'],
      thanks: ['谢谢惠顾'],
      howToPay: '付款方式',
      payText: '请转账至{supplier}，账号 00000000，并注明发票号码。',
    },
    suppliers: ['华星办公用品有限公司', '长城印刷厂', '绿源餐饮服务公司', '东方电子商行', '新华文具店', '金桥贸易有限公司'],
    customers: ['仁和诊所', '实验小学', '李明', '远景建筑设计'],
    addresses: ['北京市朝阳区建国路88号', '上海市浦东新区世纪大道100号', '广州市天河区天河路20号', '成都市锦江区春熙路5号'],
    items: ['打印纸', '黑色墨盒', '台灯', '送货费', '文件夹', '茶叶', '网线', '清洁服务', '笔记本', '签字笔'],
    currencies: [{ sym: '¥', code: 'CNY' }, { sym: '元', code: 'CNY', after: true, tight: true }],
    decimals: 2,
    taxRates: [13, 6, 9],
    digits: ['latin'],
    dates: ['yyyy年m月d日', 'yyyy-mm-dd', 'yyyy/mm/dd'],
    letter: { greeting: '尊敬的{name}：', closing: '此致敬礼', subject: '关于上月的申请' },
    form: {
      title: ['申请表', '登记表'],
      fields: ['姓名', '地址', '出生日期', '电话', '电子邮件', '签名'],
      groups: [['申请类型', ['新办', '续办', '变更']], ['联系方式', ['电子邮件', '电话', '邮寄']]],
    },
  },
  jpn: {
    dir: 'ltr',
    fonts: ['"Yu Gothic"', '"MS Gothic"'],
    title: { invoice: ['請求書'], receipt: ['領収書', 'レシート'] },
    labels: {
      number: ['請求書番号', '請求番号'],
      receiptNumber: ['領収書番号'],
      date: ['発行日', '請求日', '日付'],
      dueDate: ['お支払期限', '支払期限'],
      billTo: ['請求先', 'お客様'],
      description: ['品名', '内容'],
      quantity: ['数量'],
      unitPrice: ['単価'],
      amount: ['金額'],
      subtotal: ['小計'],
      tax: ['消費税 {rate}', '消費税'],
      taxRate: ['税率'],
      total: ['合計', 'ご請求金額', '合計金額'],
      cash: ['お預り', '現金'],
      change: ['お釣り'],
      thanks: ['ありがとうございました'],
      howToPay: 'お支払方法',
      payText: '{supplier} の口座 00000000 にお振込みください。振込名義に請求書番号をご記入ください。',
    },
    suppliers: ['山田文具店', 'さくら印刷株式会社', 'みどり食堂', '東京電機商会', '青空書店', '富士商事株式会社'],
    customers: ['田中クリニック', '緑が丘小学校', '佐藤花子', '北村設計事務所'],
    addresses: ['東京都千代田区丸の内1-1', '大阪市北区梅田2-3', '名古屋市中区栄4-5', '福岡市博多区博多駅前3-2'],
    items: ['コピー用紙', '黒インク', 'デスクライト', '配送料', 'ファイル', '緑茶', 'LANケーブル', '清掃サービス', 'ノート', 'ボールペン'],
    currencies: [{ sym: '¥', code: 'JPY' }, { sym: '円', code: 'JPY', after: true, tight: true }],
    decimals: 0,
    taxRates: [10, 8],
    digits: ['latin'],
    dates: ['yyyy年m月d日', 'yyyy/mm/dd'],
    letter: { greeting: '{name} 様', closing: '敬具', subject: '先月のお問い合わせについて' },
    form: {
      title: ['申込書', '登録用紙'],
      fields: ['氏名', '住所', '生年月日', '電話番号', 'メール', '署名'],
      groups: [['申込区分', ['新規', '更新', '変更']], ['連絡方法', ['メール', '電話', '郵送']]],
    },
  },
};

const pad2 = (n) => String(n).padStart(2, '0');

/** A number with its thousands grouped and its decimals fixed, in a language's digits. */
export function formatNumber(n, { decimals = 2, grouping = 'western', digits = 'latin' } = {}) {
  const fixed = Math.abs(n).toFixed(decimals);
  const [whole, frac] = fixed.split('.');
  let grouped;
  if (grouping === 'indian' && whole.length > 3) {
    const head = whole.slice(0, -3), tail = whole.slice(-3);
    grouped = head.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + ',' + tail;
  } else grouped = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  let out = (n < 0 ? '-' : '') + grouped + (frac ? '.' + frac : '');
  if (digits === 'arabic') out = out.replace(/,/g, '٬').replace(/\./g, '٫');
  return inDigits(out, digits);
}

/** An amount as a language prints it, with its currency. */
export function formatMoney(n, doc) {
  const num = formatNumber(n, { decimals: doc.decimals, grouping: doc.grouping, digits: doc.digits });
  const c = doc.currency;
  if (c.after) return c.tight ? num + c.sym : `${num} ${c.sym}`;
  return /^[A-Za-z]/.test(c.sym) || /^[ऀ-ॿ]/.test(c.sym) ? `${c.sym} ${num}` : c.sym + num;
}

/** A date (yyyy-mm-dd) in a pattern: d, dd, m, mm, yyyy, MMM, MMMM. */
export function formatDate(iso, pattern, { months = [], digits = 'latin' } = {}) {
  const [y, m, d] = iso.split('-').map(Number);
  const out = pattern.replace(/yyyy|MMMM|MMM|mm|m|dd|d/g, (t) => ({
    yyyy: String(y), MMMM: months[m - 1] || String(m), MMM: (months[m - 1] || String(m)).slice(0, 3),
    mm: pad2(m), m: String(m), dd: pad2(d), d: String(d),
  })[t]);
  return inDigits(out, digits);
}

const addDays = (iso, days) => {
  const t = new Date(`${iso}T00:00:00Z`);
  t.setUTCDate(t.getUTCDate() + days);
  return t.toISOString().slice(0, 10);
};

const round = (n, decimals) => Math.round(n * 10 ** decimals) / 10 ** decimals;

const ACCENTS = ['#00897B', '#1F4E79', '#7B3F98', '#37474F', '#B03A2E'];

/**
 * One document's content and its truth.
 * @param {{ kind: 'invoice'|'receipt'|'letter'|'form', layout?: string, lang: string, seed: number }} spec
 */
export function makeDocument({ kind, layout = kind, lang, seed }, { texts = {} } = {}) {
  const L = LANGS[lang];
  const r = seeded(seed);
  const fontPool = kind === 'receipt' && L.mono && r.chance(0.5) ? L.mono : L.fonts;
  const doc = {
    kind, layout, lang, seed, dir: L.dir,
    font: r.pick(fontPool),
    accent: r.pick(ACCENTS),
    size: r.pick([9.5, 10, 10.5, 11]),
    digits: r.chance(0.35) ? r.pick(L.digits) : 'latin',
    grouping: L.grouping === 'indian' && r.chance(0.6) ? 'indian' : 'western',
    decimals: L.decimals,
    currency: r.pick(L.currencies),
    chosen: {},
    colon: r.chance(0.5),
    upper: lang === 'eng' && r.chance(0.4),
  };
  // One wording a label on each page, chosen the first time it is printed.
  doc.label = (key) => (doc.chosen[key] ??= r.pick(L.labels[key]));
  const year = r.pick([2025, 2026]);
  const date = `${year}-${pad2(r.int(1, 12))}-${pad2(r.int(1, 28))}`;
  const datePattern = r.pick(L.dates);
  doc.fmtDate = (iso) => formatDate(iso, datePattern, { months: L.months, digits: doc.digits });
  doc.supplier = r.pick(L.suppliers);
  doc.address = r.pick(L.addresses);
  doc.customer = r.pick(L.customers);
  doc.customerAddress = r.pick(L.addresses.filter((a) => a !== doc.address));
  const truth = { kind, lang, layout, font: doc.font.replace(/"/g, ''), digits: doc.digits, fields: {}, items: [], checkboxes: [] };

  if (kind === 'invoice' || kind === 'receipt') {
    const count = kind === 'receipt' ? r.int(2, 6) : r.int(2, 5);
    const names = r.shuffle(L.items).slice(0, count);
    const jpy = doc.decimals === 0;
    for (const description of names) {
      const quantity = r.chance(0.5) ? 1 : r.int(2, kind === 'receipt' ? 4 : 12);
      const unitPrice = kind === 'receipt'
        ? (jpy ? r.int(1, 40) * 50 : round(r.int(50, 4000) / 100, 2))
        : (jpy ? r.int(3, 300) * 100 : round(r.int(150, 40000) / 100, 2));
      truth.items.push({ description, quantity, unitPrice, amount: round(quantity * unitPrice, doc.decimals) });
    }
    const subtotal = round(truth.items.reduce((s, it) => s + it.amount, 0), doc.decimals);
    const rate = r.pick(L.taxRates);
    const tax = jpy ? Math.floor(subtotal * rate / 100) : round(subtotal * rate / 100, 2);
    const total = round(subtotal + tax, doc.decimals);
    const number = kind === 'receipt'
      ? String(r.int(10000, 999999))
      : r.pick([`INV-${String(r.int(1, 9999)).padStart(4, '0')}`, `${year}-${String(r.int(1, 999)).padStart(4, '0')}`, String(r.int(100000, 999999))]);
    Object.assign(truth.fields, {
      supplier: doc.supplier, number, date, subtotal, tax, taxRate: rate, total, currency: doc.currency.code,
    });
    if (kind === 'invoice') truth.fields.dueDate = addDays(date, r.pick([14, 30]));
    doc.time = `${pad2(r.int(8, 21))}:${pad2(r.int(0, 59))}`;
    if (kind === 'receipt') {
      const paid = jpy ? Math.ceil(total / 1000) * 1000 : Math.ceil(total / 10) * 10;
      doc.cash = paid > total && r.chance(0.6) ? { paid, change: round(paid - total, doc.decimals) } : null;
    }
  } else if (kind === 'letter') {
    doc.letterText = (texts[lang] || []).filter((_, i) => i !== 1);
    truth.fields.date = date;
    truth.fields.supplier = doc.supplier;
  } else if (kind === 'form') {
    doc.groups = L.form.groups.map(([name, options]) => ({ name, options: options.map((o) => ({ label: o, checked: false })) }));
    for (const g of doc.groups) {
      // One box ticked in most groups; now and then none, or two.
      const ticks = r.chance(0.15) ? 0 : r.chance(0.15) ? 2 : 1;
      for (const i of r.shuffle(g.options.map((_, k) => k)).slice(0, ticks)) g.options[i].checked = true;
    }
    doc.mark = r.pick(['tick', 'cross', 'fill']);
    truth.checkboxes = doc.groups.flatMap((g) => g.options.map((o) => ({ group: g.name, label: o.label, checked: o.checked })));
    doc.formValues = L.form.fields.map((f, i) => (i === 0 ? doc.customer : i === 1 ? doc.customerAddress : i === 2 ? doc.fmtDate(`19${r.int(60, 99)}-${pad2(r.int(1, 12))}-${pad2(r.int(1, 28))}`) : ''));
  }
  doc.date = date;
  doc.truth = truth;
  return doc;
}

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** A label as this page prints it: its wording, its case, and a colon or not. */
function lab(doc, key, { colon = doc.colon, rate } = {}) {
  let text = doc.label(key);
  if (rate !== undefined) text = text.includes('{rate}') ? text.replace('{rate}', `${inDigits(String(rate), doc.digits)}%`) : text;
  else text = text.replace(/\s*\{rate\}/, '');
  if (doc.upper) text = text.toUpperCase();
  const c = doc.lang === 'chi_sim' || doc.lang === 'jpn' ? '：' : doc.dir === 'rtl' ? ':' : ':';
  return colon ? text + c : text;
}

/**
 * The page's HTML at a resolution.
 * @param {object} doc from makeDocument
 * @param {{ dpi: number, fontFaces: string }} how fontFaces: the @font-face rules for the faces not installed
 */
export function documentHtml(doc, { dpi, fontFaces = '' }) {
  const pt = (n) => `${((n / 72) * dpi).toFixed(1)}px`;
  const inch = (n) => `${Math.round(n * dpi)}px`;
  const L = LANGS[doc.lang];
  const money = (n) => formatMoney(n, doc);
  const num = (n) => inDigits(String(n), doc.digits);
  const f = doc.truth.fields;
  const lead = doc.font.includes('Nastaliq') ? 2.0 : 1.35;
  const width = doc.kind === 'receipt' ? 3.15 : 8.27;
  const pageCss = `
${fontFaces}
html, body { margin: 0; background: #fff; }
main { box-sizing: border-box; width: ${inch(width)}; ${doc.kind === 'receipt' ? `padding: ${inch(0.18)} ${inch(0.16)};` : `min-height: ${inch(11.69)}; padding: ${inch(0.75)} ${inch(0.9)};`}
  font: ${pt(doc.size)}/${lead} ${doc.font}, sans-serif; color: #111; }
table { border-collapse: collapse; width: 100%; }
td, th { padding: ${pt(3)} ${pt(5)}; vertical-align: top; text-align: start; font-weight: normal; }
.end { text-align: end; } .b { font-weight: bold; } .small { font-size: ${pt(doc.size * 0.85)}; color: #444; }
.accent { color: ${doc.accent}; } .cap { font-size: ${pt(doc.size * 0.8)}; font-weight: bold; color: ${doc.accent}; letter-spacing: 0.04em; }
`;
  let body = '';
  if (doc.kind === 'invoice' && doc.layout === 'sheet') body = sheetInvoice();
  else if (doc.kind === 'invoice') body = wordInvoice();
  else if (doc.kind === 'receipt') body = receipt();
  else if (doc.kind === 'letter') body = letter();
  else body = form();
  return `<!doctype html><html lang="${doc.lang}" dir="${doc.dir}"><meta charset="utf-8"><style>${pageCss}</style><main>${body}</main></html>`;

  // The Word gallery's invoice: name and address against the title, a rule, the bill to block, the items with the sums under them.
  function wordInvoice() {
    const title = esc(L.title.invoice[doc.seed % L.title.invoice.length]);
    const items = doc.truth.items.map((it) => `<tr style="border-bottom: 1px solid #c9d1d4"><td>${esc(it.description)}</td><td class="end">${num(it.quantity)}</td><td class="end">${money(it.unitPrice)}</td><td class="end">${money(it.amount)}</td></tr>`).join('');
    const sum = (labelHtml, value, strong) => `<tr><td></td><td></td><td class="end b"${strong ? ` style="background:#ECEFF1"` : ''}>${labelHtml}</td><td class="end${strong ? ' b' : ''}"${strong ? ` style="background:#ECEFF1"` : ''}>${value}</td></tr>`;
    return `
<table><tr>
  <td style="width:60%"><div class="b" style="font-size:${pt(14)}">${esc(doc.supplier)}</div><div class="small">${esc(doc.address)}</div><div class="small">00000 000000</div></td>
  <td class="end" style="font-size:${pt(26)}; color:${doc.accent}; letter-spacing:0.08em">${title}</td>
</tr></table>
<div style="border-bottom: ${pt(1.5)} solid ${doc.accent}; margin: ${pt(4)} 0 ${pt(14)}"></div>
<table>
  <tr><td class="cap" style="width:50%">${esc(lab(doc, 'billTo', { colon: false }))}</td><td class="cap end">${esc(lab(doc, 'number', { colon: false }))}</td><td class="cap end">${esc(lab(doc, 'date', { colon: false }))}</td></tr>
  <tr><td>${esc(doc.customer)}</td><td class="end">${esc(inDigits(f.number, /^\d+$/.test(f.number) ? doc.digits : 'latin'))}</td><td class="end">${esc(doc.fmtDate(f.date))}</td></tr>
  <tr><td>${esc(doc.customerAddress)}</td><td class="cap end">${esc(lab(doc, 'dueDate', { colon: false }))}</td><td></td></tr>
  <tr><td></td><td class="end">${esc(doc.fmtDate(f.dueDate))}</td><td></td></tr>
</table>
<div style="height:${pt(16)}"></div>
<table>
  <tr style="background:${doc.accent}; color:#fff"><td style="width:46%">${esc(lab(doc, 'description', { colon: false }))}</td><td class="end">${esc(lab(doc, 'quantity', { colon: false }))}</td><td class="end">${esc(lab(doc, 'unitPrice', { colon: false }))}</td><td class="end">${esc(lab(doc, 'amount', { colon: false }))}</td></tr>
  ${items}
  ${sum(esc(lab(doc, 'subtotal', { colon: false })), money(f.subtotal))}
  ${sum(esc(lab(doc, 'tax', { colon: false, rate: f.taxRate })), money(f.tax))}
  ${sum(esc(lab(doc, 'total', { colon: false })), money(f.total), true)}
</table>
<div style="height:${pt(18)}"></div>
<div class="b" style="font-size:${pt(12)}; color:${doc.accent}">${esc(L.labels.howToPay)}</div>
<p>${esc(L.labels.payText.replace('{supplier}', doc.supplier))}</p>
<p class="small" style="margin-top:${pt(16)}; font-style:italic">${esc(L.labels.thanks[0])}</p>`;
  }

  // The Worksheets gallery's invoice: a column of labels with their values beside them, a gridded items table, the sums in the last two columns.
  function sheetInvoice() {
    const grid = 'border: 1px solid #b8c2c6;';
    const kv = (labelHtml, value, extra = '') => `<tr><td></td><td class="end b accent" style="${extra}">${labelHtml}</td><td class="end" style="${extra}">${value}</td></tr>`;
    const rows = [];
    for (let i = 0; i < Math.max(doc.truth.items.length + 1, 6); i++) {
      const it = doc.truth.items[i];
      rows.push(`<tr><td style="${grid}">${it ? esc(it.description) : '&nbsp;'}</td><td class="end" style="${grid}">${it ? num(it.quantity) : ''}</td><td class="end" style="${grid}">${it ? money(it.unitPrice) : ''}</td><td class="end" style="${grid}">${it ? money(it.amount) : ''}</td></tr>`);
    }
    const title = LANGS[doc.lang].title.invoice[(doc.seed >> 3) % LANGS[doc.lang].title.invoice.length];
    return `
<div class="b accent" style="font-size:${pt(24)}">${esc(title)}</div>
<table style="margin-top:${pt(6)}">
  <tr><td style="width:44%"><div class="b" style="font-size:${pt(12)}">${esc(doc.supplier)}</div><div class="small">${esc(doc.address)}</div></td>
  <td style="width:56%"><table>
    ${kv(esc(lab(doc, 'number')), esc(inDigits(f.number, /^\d+$/.test(f.number) ? doc.digits : 'latin')))}
    ${kv(esc(lab(doc, 'date')), esc(doc.fmtDate(f.date)))}
    ${kv(esc(lab(doc, 'dueDate')), esc(doc.fmtDate(f.dueDate)))}
  </table></td></tr>
</table>
<div class="b accent" style="margin-top:${pt(14)}">${esc(lab(doc, 'billTo'))}</div>
<div>${esc(doc.customer)}</div><div class="small">${esc(doc.customerAddress)}</div>
<table style="margin-top:${pt(14)}">
  <tr style="background:${doc.accent}; color:#fff"><td style="width:46%">${esc(lab(doc, 'description', { colon: false }))}</td><td class="end">${esc(lab(doc, 'quantity', { colon: false }))}</td><td class="end">${esc(lab(doc, 'unitPrice', { colon: false }))}</td><td class="end">${esc(lab(doc, 'amount', { colon: false }))}</td></tr>
  ${rows.join('')}
</table>
<table style="margin-top:${pt(10)}">
  <tr><td style="width:46%"></td><td></td><td class="end b">${esc(lab(doc, 'subtotal'))}</td><td class="end">${money(f.subtotal)}</td></tr>
  <tr><td></td><td></td><td class="end b">${esc(lab(doc, 'taxRate'))}</td><td class="end">${num(f.taxRate)}%</td></tr>
  <tr><td></td><td></td><td class="end b">${esc(lab(doc, 'tax'))}</td><td class="end">${money(f.tax)}</td></tr>
  <tr><td></td><td></td><td class="end b" style="background:#ECEFF1; ${grid}">${esc(lab(doc, 'total'))}</td><td class="end b" style="background:#ECEFF1; ${grid}">${money(f.total)}</td></tr>
</table>
<p class="small" style="margin-top:${pt(18)}">${esc(L.labels.payText.replace('{supplier}', doc.supplier))}</p>
<p class="b accent">${esc(L.labels.thanks[0])}</p>`;
  }

  // A till roll: the shop centred, the number and date, a dashed rule, the items, the sums, how it was paid.
  function receipt() {
    const dash = `<div style="border-top: ${pt(1)} dashed #333; margin: ${pt(5)} 0"></div>`;
    const row = (a, b, style = '') => `<tr style="${style}"><td>${a}</td><td class="end" style="white-space:nowrap">${b}</td></tr>`;
    const title = LANGS[doc.lang].title.receipt[doc.seed % LANGS[doc.lang].title.receipt.length];
    const items = doc.truth.items.map((it) => (it.quantity > 1
      ? row(`${esc(it.description)}<br><span class="small">${num(it.quantity)} x ${money(it.unitPrice)}</span>`, money(it.amount))
      : row(esc(it.description), money(it.amount)))).join('');
    const paid = doc.cash
      ? row(esc(L.labels.cash[0]), money(doc.cash.paid)) + row(esc(L.labels.change[0]), money(doc.cash.change))
      : row(esc(L.labels.cash[L.labels.cash.length - 1]), money(f.total));
    return `
<div style="text-align:center"><div class="b" style="font-size:${pt(doc.size * 1.35)}">${esc(doc.supplier)}</div><div class="small">${esc(doc.address)}</div>
<div class="b" style="margin-top:${pt(5)}">${esc(title)}</div></div>
<table style="margin-top:${pt(5)}">
  ${row(esc(lab(doc, 'receiptNumber')), esc(inDigits(f.number, doc.digits)))}
  ${row(esc(lab(doc, 'date')), `${esc(doc.fmtDate(f.date))} ${inDigits(doc.time, doc.digits)}`)}
</table>${dash}
<table>${items}</table>${dash}
<table>
  ${row(esc(lab(doc, 'subtotal')), money(f.subtotal))}
  ${row(esc(lab(doc, 'tax', { rate: f.taxRate })), money(f.tax))}
  ${row(esc(lab(doc, 'total')), money(f.total), `font-weight:bold; font-size:${pt(doc.size * 1.25)}`)}
</table>${dash}
<table>${paid}</table>
<div style="text-align:center; margin-top:${pt(8)}">${esc(L.labels.thanks[doc.seed % L.labels.thanks.length])}</div>`;
  }

  function letter() {
    const texts = doc.letterText || [];
    return `
<div style="text-align:end"><div class="b" style="font-size:${pt(14)}; color:${doc.accent}">${esc(doc.supplier)}</div><div class="small">${esc(doc.address)}</div></div>
<div style="margin-top:${pt(28)}">${esc(doc.fmtDate(doc.date))}</div>
<div style="margin-top:${pt(14)}">${esc(doc.customer)}<br>${esc(doc.customerAddress)}</div>
<div class="b" style="margin-top:${pt(18)}">${esc(L.letter.subject)}</div>
<p style="margin-top:${pt(12)}">${esc(L.letter.greeting.replace('{name}', doc.customer))}</p>
${texts.map((t) => `<p>${esc(t)}</p>`).join('')}
<p style="margin-top:${pt(16)}">${esc(L.letter.closing)}</p>
<p style="margin-top:${pt(30)}">${esc(doc.supplier)}</p>`;
  }

  function form() {
    const box = (checked) => {
      const s = `display:inline-block; width:${pt(doc.size * 0.95)}; height:${pt(doc.size * 0.95)}; border:${pt(1.1)} solid #222; vertical-align:middle; position:relative; margin-inline-end:${pt(5)}; box-sizing:border-box;`;
      if (!checked) return `<span class="cb" data-checked="0" style="${s}"></span>`;
      if (doc.mark === 'fill') return `<span class="cb" data-checked="1" style="${s} background: radial-gradient(#222 55%, transparent 60%);"></span>`;
      const glyph = doc.mark === 'tick' ? '✓' : '✗';
      return `<span class="cb" data-checked="1" style="${s}"><span style="position:absolute; inset:${pt(-2)} 0 0 ${pt(1)}; font: bold ${pt(doc.size * 1.05)}/1 'Segoe UI Symbol', sans-serif;">${glyph}</span></span>`;
    };
    const fields = L.form.fields.map((name, i) => `<tr><td style="width:28%" class="b">${esc(name)}${doc.colon ? (doc.lang === 'chi_sim' || doc.lang === 'jpn' ? '：' : ':') : ''}</td><td style="border-bottom:1px solid #333">${esc(doc.formValues[i] || '')}&nbsp;</td></tr>`).join('');
    const groups = doc.groups.map((g) => `<div style="margin-top:${pt(12)}"><span class="b">${esc(g.name)}</span>${doc.lang === 'chi_sim' || doc.lang === 'jpn' ? '：' : ':'}&nbsp;&nbsp;${g.options.map((o) => `<span style="margin-inline-end:${pt(18)}; white-space:nowrap">${box(o.checked)}${esc(o.label)}</span>`).join('')}</div>`).join('');
    const title = L.form.title[doc.seed % L.form.title.length];
    return `
<div class="b" style="font-size:${pt(18)}; color:${doc.accent}; text-align:center">${esc(title)}</div>
<div class="small" style="text-align:center">${esc(doc.supplier)}</div>
<table style="margin-top:${pt(20)}">${fields}</table>
${groups}`;
  }
}

/**
 * The corpus: which documents, in which conditions. Every language gets the
 * same set; the seeds differ so wordings, fonts and figures differ.
 */
export function corpusPlan(langs = Object.keys(LANGS)) {
  const plan = [];
  const conditions = ['clean', 'scan', 'photo', 'scan'];
  const shapes = [
    { kind: 'invoice', layout: 'word', count: 4 },
    { kind: 'invoice', layout: 'sheet', count: 3 },
    { kind: 'receipt', layout: 'receipt', count: 4 },
    { kind: 'letter', layout: 'letter', count: 2 },
    { kind: 'form', layout: 'form', count: 2 },
  ];
  langs.forEach((lang, li) => {
    for (const s of shapes) {
      for (let n = 0; n < s.count; n++) {
        const seed = 1000 * (li + 1) + 97 * shapes.indexOf(s) + n * 13 + 1;
        plan.push({ id: `${lang}-${s.layout}-${n + 1}`, kind: s.kind, layout: s.layout, lang, seed, condition: conditions[n % conditions.length] });
      }
    }
  });
  return plan;
}
