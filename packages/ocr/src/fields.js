// The fields of a bill, an invoice or a receipt, from what recognize() read:
// the total, the subtotal, the tax, the number, the date, the due date and
// who it is from, with the lines of items, a check that the sums agree, and
// a guess at what kind of document the page is (invoice, receipt, letter,
// form). Rules and geometry of our own over the engine's words; nothing is
// learnt and nothing is fetched.
//
//   import { documentFields } from '@rutba/ocr/fields';
//   const read = await recognize(picture, { languages: ['ara', 'eng'] });
//   const { kind, fields, items, checks } = documentFields(read);
//   fields.total  // { value: 1744.48, text, box, confidence, label, checked }
//
// How it reads a page:
//   - the words are levelled (the slope of the engine's lines taken out),
//     gathered into rows, and each row cut into segments where a gap wider
//     than a text height opens, as a label stands apart from its value;
//   - a label is the start of a segment matching a phrase of the field's
//     dictionary (six languages, a letter or two of misreading allowed in a
//     longer phrase); its value is the rest of the segment, else the next
//     segment along the row in the page's direction, else the segment below
//     it, never past another label;
//   - amounts and dates are read by locale: Arabic-Indic, Urdu and
//     Devanagari digits, the Arabic decimal and thousands marks, Indian
//     grouping, a comma or a point for the decimals, day-first dates, month
//     names in each language, and the Chinese and Japanese 年月日;
//   - the items are the rows under the items table's heading (or, on a till
//     roll, the rows above the subtotal ending in an amount), and the sums
//     choose among the candidates: items adding up to the subtotal, the
//     subtotal and the tax to the total, the tax being the rate of the
//     subtotal. A field the sums confirm is `checked: true`; a page whose sums
//     agree is `checks.consistent`.

import { joinWords } from './results.js';

// ---- Text ------------------------------------------------------------------

const DIGIT_ZEROS = [0x660, 0x6f0, 0x966, 0xff10];

/** Digits of any of the scripts read (Arabic-Indic, Urdu, Devanagari, full width) as 0 to 9. */
export function westernDigits(text) {
  let out = '';
  for (const ch of String(text)) {
    const c = ch.codePointAt(0);
    const zero = DIGIT_ZEROS.find((z) => c >= z && c <= z + 9);
    out += zero === undefined ? ch : String(c - zero);
  }
  return out.replace(/٫/g, '.').replace(/٬/g, ',');
}

const PUNCT = /[\s.,:;：；、。，·'"`()[\]{}<>«»!?¡¿*#№|/\\_~\-\u2013\u2014=+،؛؟٪%]/gu;
/** Text compared with a dictionary: compatibility forms, lower case, Western digits, no punctuation or spaces. */
const fold = (text) => westernDigits(String(text).normalize('NFKC').toLowerCase())
  // Arabic letters the engine and the fonts write either way.
  .replace(/[أإآٱ]/g, 'ا').replace(/ى/g, 'ي').replace(/ة/g, 'ه').replace(/ك/g, 'ک').replace(/ي/g, 'ی').replace(/[ً-ٰٟ]/g, '')
  .replace(PUNCT, '');

const CJK = /[぀-ヿ㐀-鿿豈-﫿]/u;

/** Edit distance, two strings. */
function distance(a, b) {
  const s = [...a], t = [...b];
  let prev = Array.from({ length: t.length + 1 }, (_, i) => i);
  for (let i = 1; i <= s.length; i++) {
    const row = [i];
    for (let j = 1; j <= t.length; j++) row[j] = Math.min(prev[j] + 1, row[j - 1] + 1, prev[j - 1] + (s[i - 1] === t[j - 1] ? 0 : 1));
    prev = row;
  }
  return prev[t.length];
}

// ---- Dictionaries ----------------------------------------------------------

/**
 * What each field is called, in each language read. More wordings than any
 * one document uses; a phrase is matched at the start of a segment.
 */
export const LABELS = {
  number: {
    eng: ['invoice number', 'invoice no', 'invoice #', 'invoice num', 'inv no', 'receipt number', 'receipt no', 'receipt #', 'bill number', 'bill no', 'transaction no', 'document no'],
    ara: ['رقم الفاتورة', 'فاتورة رقم', 'رقم الإيصال', 'رقم الايصال', 'رقم السند'],
    urd: ['انوائس نمبر', 'بل نمبر', 'رسید نمبر'],
    hin: ['बीजक संख्या', 'इनवॉइस नंबर', 'इनवॉइस संख्या', 'चालान संख्या', 'बिल संख्या', 'बिल नंबर', 'रसीद संख्या', 'रसीद नंबर'],
    chi_sim: ['发票号码', '发票编号', '发票号', '单号', '收据编号', '票据号码', '流水号'],
    jpn: ['請求書番号', '請求番号', '領収書番号', '伝票番号', 'レシート番号'],
  },
  date: {
    eng: ['date', 'invoice date', 'issue date', 'date of issue', 'dated', 'receipt date', 'tax point'],
    ara: ['التاريخ', 'تاريخ الفاتورة', 'تاريخ الإصدار', 'تاريخ'],
    urd: ['تاریخ', 'تاریخ اجرا', 'تاریخ اجراء'],
    hin: ['दिनांक', 'तारीख', 'तिथि', 'बीजक दिनांक'],
    chi_sim: ['日期', '开票日期', '开具日期'],
    jpn: ['発行日', '請求日', '日付', '発行日付', 'お買上日'],
  },
  dueDate: {
    eng: ['due date', 'payment due', 'due by', 'pay by', 'date due', 'payment due date'],
    ara: ['تاريخ الاستحقاق', 'آخر موعد للدفع', 'موعد الاستحقاق'],
    urd: ['ادائیگی کی آخری تاریخ', 'آخری تاریخ', 'واجب الادا تاریخ'],
    hin: ['देय तिथि', 'भुगतान की अंतिम तिथि', 'अंतिम तिथि', 'नियत तिथि'],
    chi_sim: ['付款期限', '到期日', '付款日期', '截止日期'],
    jpn: ['お支払期限', '支払期限', 'お支払期日', '支払期日'],
  },
  subtotal: {
    eng: ['subtotal', 'sub total', 'net total', 'net amount', 'net', 'total before tax'],
    ara: ['المجموع الفرعي', 'الإجمالي الفرعي', 'المبلغ قبل الضريبة'],
    urd: ['ذیلی میزان', 'ذیلی کل', 'ذیلی رقم'],
    hin: ['उप-योग', 'उप योग', 'उपयोग', 'उप-कुल', 'उप कुल'],
    chi_sim: ['小计', '不含税金额'],
    jpn: ['小計', '税抜金額', '税抜合計'],
  },
  tax: {
    eng: ['vat', 'tax', 'sales tax', 'gst', 'vat amount', 'tax amount'],
    ara: ['ضريبة القيمة المضافة', 'الضريبة', 'ضريبة', 'قيمة الضريبة'],
    urd: ['سیلز ٹیکس', 'ٹیکس', 'جی ایس ٹی'],
    hin: ['जीएसटी', 'कर', 'कर राशि', 'वैट'],
    chi_sim: ['增值税', '税额', '税金'],
    jpn: ['消費税', '税額', '内消費税', '消費税等'],
  },
  taxRate: {
    eng: ['vat rate', 'tax rate', 'gst rate'],
    ara: ['نسبة الضريبة', 'معدل الضريبة'],
    urd: ['ٹیکس کی شرح'],
    hin: ['कर की दर', 'कर दर', 'जीएसटी दर'],
    chi_sim: ['税率'],
    jpn: ['税率'],
  },
  total: {
    eng: ['total', 'total due', 'amount due', 'grand total', 'balance due', 'total amount', 'amount payable', 'total payable'],
    ara: ['الإجمالي', 'المبلغ المستحق', 'المجموع الكلي', 'المجموع', 'الإجمالي المستحق', 'المبلغ الإجمالي'],
    urd: ['کل رقم', 'واجب الادا رقم', 'کل', 'کل میزان'],
    hin: ['कुल राशि', 'कुल देय', 'कुल', 'कुल योग', 'देय राशि'],
    chi_sim: ['合计', '总计', '应付金额', '价税合计', '合计金额', '实付'],
    jpn: ['合計', 'ご請求金額', '合計金額', '請求金額', '総合計', 'お支払金額'],
  },
  supplier: {
    eng: ['from', 'supplier', 'seller', 'vendor', 'sold by'],
    ara: ['المورد', 'البائع'],
    urd: ['فروخت کنندہ', 'سپلائر'],
    hin: ['विक्रेता', 'आपूर्तिकर्ता'],
    chi_sim: ['销售方', '卖方', '供应商', '销货单位'],
    jpn: ['発行者', '販売者', '発行元'],
  },
  billTo: {
    eng: ['bill to', 'billed to', 'invoice to', 'customer', 'sold to', 'ship to'],
    ara: ['فاتورة إلى', 'العميل', 'المشتري'],
    urd: ['بنام', 'گاہک', 'خریدار'],
    hin: ['प्राप्तकर्ता', 'ग्राहक', 'क्रेता'],
    chi_sim: ['购买方', '客户', '买方', '购货单位'],
    jpn: ['請求先', 'お客様', '御中'],
  },
  paid: {
    eng: ['cash', 'card', 'paid', 'tendered', 'change', 'change due'],
    ara: ['نقدا', 'بطاقة', 'المدفوع', 'الباقي', 'المتبقي'],
    urd: ['نقد', 'کارڈ', 'ادا شدہ', 'بقایا'],
    hin: ['नकद', 'कार्ड', 'वापसी', 'शेष'],
    chi_sim: ['现金', '实收', '刷卡', '找零'],
    jpn: ['お預り', '現金', 'お預かり', 'お釣り', 'おつり'],
  },
};

/** The items table's headings. */
export const HEADINGS = {
  description: {
    eng: ['description', 'item', 'items', 'product', 'details', 'particulars'], ara: ['الوصف', 'البند', 'الصنف', 'المنتج'], urd: ['تفصیل', 'اشیاء'],
    hin: ['विवरण', 'वस्तु', 'सामान'], chi_sim: ['项目', '品名', '商品', '名称'], jpn: ['品名', '内容', '品目', '商品名', '摘要'],
  },
  quantity: { eng: ['quantity', 'qty', 'units', 'hours'], ara: ['الكمية', 'العدد'], urd: ['تعداد', 'مقدار'], hin: ['मात्रा'], chi_sim: ['数量'], jpn: ['数量'] },
  unitPrice: { eng: ['unit price', 'price', 'rate', 'unit cost'], ara: ['سعر الوحدة', 'السعر'], urd: ['فی عدد قیمت', 'قیمت', 'نرخ'], hin: ['दर', 'इकाई मूल्य', 'मूल्य'], chi_sim: ['单价'], jpn: ['単価'] },
  amount: { eng: ['amount', 'line total', 'value'], ara: ['المبلغ', 'القيمة'], urd: ['رقم'], hin: ['राशि', 'रकम'], chi_sim: ['金额'], jpn: ['金額'] },
};

/** Words that say what a page is, for the guess at its kind. */
export const KINDS = {
  invoice: { eng: ['invoice', 'tax invoice', 'proforma'], ara: ['فاتورة', 'فاتورة ضريبية'], urd: ['انوائس', 'بل'], hin: ['बीजक', 'इनवॉइस', 'चालान', 'टैक्स इनवॉइस'], chi_sim: ['发票', '销售发票', '增值税发票'], jpn: ['請求書', '御請求書'] },
  receipt: { eng: ['receipt', 'sales receipt', 'till receipt'], ara: ['إيصال', 'ايصال', 'إيصال بيع'], urd: ['رسید'], hin: ['रसीद'], chi_sim: ['收据', '小票', '购物小票'], jpn: ['領収書', 'レシート'] },
  greeting: { eng: ['dear'], ara: ['السيد', 'عزيزي', 'تحية طيبة'], urd: ['محترم', 'جناب', 'بخدمت'], hin: ['प्रिय', 'महोदय', 'आदरणीय'], chi_sim: ['尊敬的', '您好'], jpn: ['拝啓', '謹啓', '様'] },
  closing: { eng: ['yours sincerely', 'yours faithfully', 'kind regards', 'best regards', 'sincerely', 'regards'], ara: ['وتفضلوا بقبول', 'مع التحية', 'مع خالص', 'وتفضلوا'], urd: ['آپ کا مخلص', 'مخلص', 'والسلام'], hin: ['भवदीय', 'सादर', 'आपका'], chi_sim: ['此致', '敬礼', '此致敬礼'], jpn: ['敬具', '謹白', 'よろしくお願いいたします'] },
  form: { eng: ['application form', 'registration form', 'form'], ara: ['استمارة', 'نموذج', 'استمارة طلب', 'نموذج تسجيل'], urd: ['فارم', 'درخواست فارم'], hin: ['आवेदन पत्र', 'फ़ॉर्म', 'फॉर्म', 'आवेदन', 'पंजीकरण'], chi_sim: ['申请表', '登记表', '表格'], jpn: ['申込書', '登録用紙', '申請書'] },
  formField: { eng: ['name', 'address', 'date of birth', 'signature', 'phone', 'email'], ara: ['الاسم', 'العنوان', 'تاريخ الميلاد', 'التوقيع', 'الهاتف'], urd: ['نام', 'پتہ', 'تاریخ پیدائش', 'دستخط', 'فون'], hin: ['नाम', 'पता', 'जन्म तिथि', 'हस्ताक्षर', 'फ़ोन'], chi_sim: ['姓名', '地址', '出生日期', '签名', '电话'], jpn: ['氏名', '住所', '生年月日', '署名', '電話番号'] },
};

const LANGS = ['eng', 'ara', 'urd', 'hin', 'chi_sim', 'jpn'];

/** A dictionary's phrases for the languages asked (every language when none), folded, longest first. */
function phrases(dict, languages) {
  const langs = languages?.length ? languages.filter((l) => LANGS.includes(l)) : LANGS;
  const out = [];
  for (const l of langs.length ? langs : LANGS) for (const p of dict[l] || []) out.push({ text: p, folded: fold(p), lang: l });
  return out.sort((a, b) => b.folded.length - a.folded.length);
}

// ---- Values ----------------------------------------------------------------

const CURRENCIES = [
  [/£|gbp/i, 'GBP'], [/€|eur\b/i, 'EUR'], [/\$|usd/i, 'USD'], [/₹|रु|inr/i, 'INR'], [/روپے|\brs\.?|pkr/i, 'PKR'],
  [/ر\.?\s?س|ررس|sar/i, 'SAR'], [/د\.?\s?إ|aed/i, 'AED'], [/円|jpy/i, 'JPY'], [/元|cny|rmb/i, 'CNY'], [/[¥￥]/, 'YEN'],
];

/**
 * An amount in a piece of text: the last number in it that is not a
 * percentage, a time or part of a date, read by its own separators.
 * @returns {{ value: number, decimals: boolean, currency: string|null }|null}
 */
export function parseAmount(text) {
  // The engine can space a figure about its separators ("¥1, 582. 70").
  const t = westernDigits(String(text).normalize('NFKC')).replace(/(\d)\s*([.,])\s+(?=\d)/g, '$1$2').replace(/(\d)\s+([.,])(?=\d)/g, '$1$2');
  let currency = null;
  for (const [re, code] of CURRENCIES) if (re.test(t)) { currency = code; break; }
  const found = [];
  const re = /\d(?:[\d,.'’]|\s(?=\d{3}(?!\d)))*\d|\d/g;
  let m;
  while ((m = re.exec(t))) {
    const after = t.slice(m.index + m[0].length);
    const before = t.slice(0, m.index);
    if (/^\s?%/.test(after) || /[:/-]\s?$/.test(before) && /\d\s?[:/-]\s?$/.test(before) || /^\s?[:/年月]\d?/.test(after) || /^\s?[-/]\s?\d/.test(after)) continue;
    if (/^\d{1,2}\.\d{1,2}\.\d{2,4}$/.test(m[0])) continue;
    found.push(m[0]);
  }
  if (!found.length) return null;
  const raw = found[found.length - 1].replace(/\s/g, '');
  const sep = Math.max(raw.lastIndexOf('.'), raw.lastIndexOf(','));
  let value, decimals = false;
  if (sep >= 0 && raw.length - sep - 1 <= 2 && raw.length - sep - 1 >= 1) {
    value = Number(`${raw.slice(0, sep).replace(/[,.'’]/g, '')}.${raw.slice(sep + 1)}`);
    decimals = true;
  } else value = Number(raw.replace(/[,.'’]/g, ''));
  if (!Number.isFinite(value)) return null;
  return { value, decimals, currency };
}

/** A percentage in a piece of text. */
export function parsePercent(text) {
  const m = westernDigits(String(text).normalize('NFKC')).match(/(\d+(?:[.,]\d+)?)\s?%/);
  return m ? Number(m[1].replace(',', '.')) : null;
}

const MONTHS = {
  eng: ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'],
  ara: ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'],
  levant: ['كانون الثاني', 'شباط', 'آذار', 'نيسان', 'أيار', 'حزيران', 'تموز', 'آب', 'أيلول', 'تشرين الأول', 'تشرين الثاني', 'كانون الأول'],
  urd: ['جنوری', 'فروری', 'مارچ', 'اپریل', 'مئی', 'جون', 'جولائی', 'اگست', 'ستمبر', 'اکتوبر', 'نومبر', 'دسمبر'],
  hin: ['जनवरी', 'फ़रवरी', 'मार्च', 'अप्रैल', 'मई', 'जून', 'जुलाई', 'अगस्त', 'सितंबर', 'अक्टूबर', 'नवंबर', 'दिसंबर'],
};
const MONTH_NAMES = [];
for (const list of Object.values(MONTHS)) list.forEach((name, i) => MONTH_NAMES.push({ folded: fold(name), month: i + 1 }));
MONTHS.eng.forEach((name, i) => MONTH_NAMES.push({ folded: name.slice(0, 3), month: i + 1, short: true }));
MONTH_NAMES.push({ folded: 'sept', month: 9, short: true }, { folded: fold('फरवरी'), month: 2 }, { folded: fold('सितम्बर'), month: 9 }, { folded: fold('नवम्बर'), month: 11 }, { folded: fold('दिसम्बर'), month: 12 });

/** A month named by a word, a letter of misreading allowed in a longer name. */
function monthOf(word) {
  const w = fold(word);
  if (!w) return 0;
  let best = 0, bestD = Infinity;
  for (const m of MONTH_NAMES) {
    const d = m.short ? (w.startsWith(m.folded) && w.length <= 9 ? 0 : Infinity) : distance(w, m.folded);
    const allowed = m.short ? 0 : m.folded.length >= 5 ? 1 : 0;
    if (d <= allowed && d < bestD) { best = m.month; bestD = d; }
  }
  return best;
}

const iso = (y, m, d) => {
  if (y < 100) y += 2000;
  if (!(y >= 1900 && y <= 2100 && m >= 1 && m <= 12 && d >= 1 && d <= 31)) return null;
  const t = new Date(Date.UTC(y, m - 1, d));
  if (t.getUTCMonth() !== m - 1) return null;
  return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
};

/**
 * A date in a piece of text, as yyyy-mm-dd: year first (2026-10-14,
 * 2026/10/14, 2026年10月14日), day first (14/10/2026, 14.10.26; month first
 * only when the first number cannot be a month), or with the month's name in
 * English, Arabic, Urdu or Hindi.
 * @param {{ monthFirst?: boolean }} [options] read 10/04/2026 as October the 4th
 */
export function parseDate(text, { monthFirst = false } = {}) {
  const t = westernDigits(String(text).normalize('NFKC')).replace(/\s+/g, ' ');
  let m = t.match(/(\d{4})\s?[-/.年]\s?(\d{1,2})\s?[-/.月]\s?(\d{1,2})/);
  if (m) return iso(+m[1], +m[2], +m[3]);
  m = t.match(/(\d{1,2})\s?[-/.]\s?(\d{1,2})\s?[-/.]\s?(\d{4}|\d{2})(?!\d)/);
  if (m) {
    let d = +m[1], mo = +m[2];
    if ((monthFirst && d <= 12) || mo > 12) [d, mo] = [mo, d];
    return iso(+m[3], mo, d);
  }
  // A month's name: "14 Oct 2026", "October 14, 2026", "14 اکتوبر 2026".
  const words = t.split(/[\s,.،]+/).filter(Boolean);
  for (let i = 0; i < words.length; i++) {
    const month = monthOf(words[i]);
    if (!month) continue;
    const num = (w) => (w && /^\d{1,4}$/.test(w.replace(/(st|nd|rd|th)$/i, '')) ? +w.replace(/(st|nd|rd|th)$/i, '') : null);
    const before = num(words[i - 1]), after = num(words[i + 1]), after2 = num(words[i + 2]);
    if (before !== null && before <= 31 && after !== null && after > 31) return iso(after, month, before);
    if (after !== null && after <= 31 && after2 !== null && after2 > 31) return iso(after2, month, after);
  }
  return null;
}

/** A document number: the first run of letters and digits holding a digit, three characters or more. */
export function parseNumber(text) {
  const t = westernDigits(String(text).normalize('NFKC')).replace(/^[\s:#№.\-]+/, '').replace(/^(no|nr|num)\b[.:]?\s*/i, '');
  const m = t.match(/[A-Za-z0-9][A-Za-z0-9\-/]*\d[A-Za-z0-9\-/]*|\d[A-Za-z0-9\-/]*/);
  if (!m) return null;
  const v = m[0].replace(/[-/]+$/, '');
  return v.length >= 3 ? v : null;
}

// ---- Geometry --------------------------------------------------------------

const median = (xs) => {
  if (!xs.length) return 0;
  const s = [...xs].sort((a, b) => a - b);
  return s[s.length >> 1];
};

/**
 * The words levelled: the median slope of the engine's lines (a page photographed
 * or scanned a little turned) taken out, so a label and its value far along the
 * row still fall in one row.
 */
function level(read) {
  const words = read.words || [];
  const slopes = [];
  for (const line of read.lines || []) {
    const [a, b] = line.words || [0, 0];
    const ws = words.slice(a, b);
    if (ws.length < 2) continue;
    const xs = ws.map((w) => w.box.x + w.box.width / 2), ys = ws.map((w) => w.box.y + w.box.height / 2);
    const span = Math.max(...xs) - Math.min(...xs);
    if (span < 100) continue;
    const mx = xs.reduce((s, x) => s + x, 0) / xs.length, my = ys.reduce((s, y) => s + y, 0) / ys.length;
    let num = 0, den = 0;
    xs.forEach((x, i) => { num += (x - mx) * (ys[i] - my); den += (x - mx) ** 2; });
    if (den > 0) slopes.push(num / den);
  }
  const slope = slopes.length >= 3 ? median(slopes) : 0;
  if (Math.abs(slope) < 0.003) return { words: words.map((w, i) => ({ ...w, index: i })), slope: 0 };
  return { words: words.map((w, i) => ({ ...w, index: i, original: w.box, box: { ...w.box, y: w.box.y - slope * (w.box.x + w.box.width / 2) } })), slope };
}

/**
 * Words in rows: a word joins the row whose middle band (its words' median
 * top to median bottom) holds its centre and whose words are of a like
 * height. Unlike a band that grows with every word, a large heading beside
 * two lines of small type takes neither of them into its row.
 */
function rowsOf(words) {
  const rows = [];
  for (const w of [...words].sort((a, b) => (a.box.y + a.box.height / 2) - (b.box.y + b.box.height / 2))) {
    const cy = w.box.y + w.box.height / 2;
    let best = null, bestGap = Infinity;
    for (const row of rows) {
      if (cy < row.top || cy > row.bottom) continue;
      const ratio = Math.max(w.box.height, row.height) / Math.max(1, Math.min(w.box.height, row.height));
      if (ratio > 2.2) continue;
      const gap = Math.abs(cy - (row.top + row.bottom) / 2);
      if (gap < bestGap) { best = row; bestGap = gap; }
    }
    if (!best) { best = { words: [] }; rows.push(best); }
    best.words.push(w);
    best.top = median(best.words.map((x) => x.box.y));
    best.bottom = median(best.words.map((x) => x.box.y + x.box.height));
    best.height = median(best.words.map((x) => x.box.height));
  }
  return rows
    .map((row) => {
      const x0 = Math.min(...row.words.map((x) => x.box.x)), x1 = Math.max(...row.words.map((x) => x.box.x + x.box.width));
      const y0 = Math.min(...row.words.map((x) => x.box.y)), y1 = Math.max(...row.words.map((x) => x.box.y + x.box.height));
      return { box: { x: x0, y: y0, width: x1 - x0, height: y1 - y0 }, words: row.words, mid: (row.top + row.bottom) / 2 };
    })
    .sort((a, b) => a.mid - b.mid);
}

/** Rows of segments: words apart by more than a gap wider than `gap` begin a new segment. */
function segment(words, { rtl, gap }) {
  const rows = rowsOf(words);
  return rows.map((row, r) => {
    const byX = [...row.words].sort((a, b) => a.box.x - b.box.x);
    const pieces = [];
    for (const w of byX) {
      const last = pieces[pieces.length - 1];
      if (last && w.box.x - last.x1 <= gap) { last.words.push(w); last.x1 = Math.max(last.x1, w.box.x + w.box.width); }
      else pieces.push({ x0: w.box.x, x1: w.box.x + w.box.width, words: [w] });
    }
    if (rtl) pieces.reverse();
    const segs = pieces.map((p, i) => {
      const ws = rtl ? [...p.words].reverse() : p.words;
      const ys = ws.map((w) => w.box.y), ye = ws.map((w) => w.box.y + w.box.height);
      return {
        row: r, index: i, words: ws, text: joinWords(ws.map((w) => w.text)),
        box: { x: p.x0, y: Math.min(...ys), width: p.x1 - p.x0, height: Math.max(...ye) - Math.min(...ys) },
        height: median(ws.map((w) => w.box.height)),
      };
    });
    return { index: r, box: row.box, segments: segs };
  });
}

/**
 * Whether a segment starts with one of the phrases: the phrase against the
 * segment's first characters with spaces and punctuation aside, a misread
 * letter allowed in five, ending at a word's end (anywhere in Chinese and
 * Japanese, written without spaces).
 * @returns {{ phrase, used: number }|null} used: how many folded characters the label took
 */
function startsWith(seg, list) {
  const tokens = seg.words.map((w) => fold(w.text));
  const s = tokens.join('');
  if (!s) return null;
  const ends = new Set();
  let n = 0;
  for (const t of tokens) { n += t.length; ends.add(n); }
  let best = null;
  for (const p of list) {
    const len = [...p.folded].length;
    if (!len) continue;
    const cjk = CJK.test(p.folded);
    const allowed = cjk ? (len >= 4 ? 1 : 0) : len >= 10 ? 2 : len >= 5 ? 1 : 0;
    for (let e = Math.max(1, len - allowed); e <= len + allowed; e++) {
      if (e > s.length) break;
      if (!cjk && !ends.has(e)) continue;
      const d = distance(s.slice(0, e), p.folded);
      // A longer phrase wins, but a misread letter costs more than a letter of length.
      if (d <= allowed && (!best || len - 4 * d > best.len - 4 * best.d)) best = { phrase: p, used: e, len, d };
    }
  }
  return best;
}

/** The segment's text after its label: whole words after it, and what is left of a word the label ended inside. */
function after(seg, used) {
  let n = 0;
  const out = [];
  for (const w of seg.words) {
    const f = fold(w.text);
    if (n >= used) out.push(w);
    else if (n + f.length > used) {
      // The label ends inside this word: keep its characters past the label.
      let keep = used - n, i = 0;
      const chars = [...w.text.normalize('NFKC')];
      while (i < chars.length && keep > 0) { if (fold(chars[i])) keep--; i++; }
      const rest = chars.slice(i).join('');
      if (rest.replace(PUNCT, '')) out.push({ ...w, text: rest });
    }
    n += f.length;
  }
  return out;
}

// ---- Reading ---------------------------------------------------------------

const AMOUNT_FIELDS = ['subtotal', 'tax', 'total'];
const PARSERS = {
  number: (t) => parseNumber(t),
  date: (t, o) => parseDate(t, o),
  dueDate: (t, o) => parseDate(t, o),
  subtotal: (t) => parseAmount(t),
  tax: (t) => parseAmount(t),
  total: (t) => parseAmount(t),
  taxRate: (t) => parsePercent(t),
  supplier: (t) => (t.replace(PUNCT, '').length >= 2 && /\p{L}/u.test(t) ? t.trim() : null),
  billTo: (t) => (t.replace(PUNCT, '').length >= 2 && /\p{L}/u.test(t) ? t.trim() : null),
  paid: (t) => parseAmount(t),
};

/** A box as a key: its edges in whole pixels. */
const keyOf = (b) => `${Math.round(b.x)},${Math.round(b.y)},${Math.round(b.width)},${Math.round(b.height)}`;

const confidenceOf = (words) => (words.length && words.every((w) => typeof w.confidence === 'number') ? Math.min(...words.map((w) => w.confidence)) : null);
const unionBox = (words) => {
  const bs = words.map((w) => w.original || w.box);
  const x0 = Math.min(...bs.map((b) => b.x)), y0 = Math.min(...bs.map((b) => b.y));
  const x1 = Math.max(...bs.map((b) => b.x + b.width)), y1 = Math.max(...bs.map((b) => b.y + b.height));
  return { x: x0, y: y0, width: x1 - x0, height: y1 - y0 };
};

/**
 * The fields of the document a page holds.
 *
 * @param {{ words: object[], lines?: object[], width?: number, height?: number, languages?: string[] }} read
 *   what recognize() answered (the words and lines are what is used)
 * @param {{ languages?: string[], monthFirst?: boolean, checkboxes?: object[] }} [options]
 *   languages: the dictionaries to use (the read's own, else all six);
 *   monthFirst: read 10/04/2026 as October the 4th;
 *   checkboxes: tick boxes found on the page (checkboxes.js), a sign of a form
 * @returns {{
 *   kind: 'invoice'|'receipt'|'letter'|'form'|'other', scores: object,
 *   fields: { [name]: { value, text, box, confidence, label, checked } },
 *   items: { description, quantity, unitPrice, amount, box }[],
 *   checks: { itemsAddUp, totalAddsUp, taxRateHolds, consistent }|null,
 *   currency: string|null
 * }}
 */
export function documentFields(read, options = {}) {
  const languages = options.languages || read.languages;
  let { words, slope } = level(read);
  const empty = { kind: 'other', scores: {}, fields: {}, items: [], checks: null, currency: null, slope };
  if (!words.length) return empty;
  // Chinese and Japanese characters are about as tall as they are wide; the engine boxes some of them
  // far taller (a character on a line it merged with a large title), which would split their rows.
  words = words.map((w) => {
    const n = [...w.text].length;
    if (!CJK.test(w.text) || !n) return w;
    const size = w.box.width / n;
    if (w.box.height <= size * 1.5) return w;
    const cy = w.box.y + w.box.height / 2;
    return { ...w, box: { ...w.box, y: cy - size / 2, height: size } };
  });
  const h0 = median(words.map((w) => w.box.height)) || 20;
  // A word a third of the text's height or thinner is a rule read as letters (a dashed line on a till roll).
  words = words.filter((w) => w.box.height > h0 * 0.3);
  const h = median(words.map((w) => w.box.height)) || h0;
  const rtlWords = words.filter((w) => /[֐-ࣿיִ-﷿ﹰ-ﻼ]/u.test(w.text)).length;
  const ltrWords = words.filter((w) => /[A-Za-zऀ-ॿ぀-ヿ一-鿿]/u.test(w.text)).length;
  const rtl = rtlWords > ltrWords;
  const rows = segment(words, { rtl, gap: h * 1.1 });
  const segs = rows.flatMap((r) => r.segments);
  const dict = Object.fromEntries(Object.entries(LABELS).map(([k, d]) => [k, phrases(d, languages)]));
  const headingDict = Object.fromEntries(Object.entries(HEADINGS).map(([k, d]) => [k, phrases(d, languages)]));

  // The items table's heading row: a row naming two or more of its columns.
  let headingRow = -1;
  for (const row of rows) {
    const named = new Set();
    for (const seg of row.segments) {
      for (const [k, list] of Object.entries(headingDict)) {
        const hit = startsWith(seg, list);
        if (hit && fold(seg.text).length <= hit.used + 2) named.add(k);
      }
    }
    if (named.size >= 2) { headingRow = row.index; break; }
  }

  // Labels: every segment that begins with a field's name (the heading row's names are column names, not labels).
  const hits = [];
  for (const seg of segs) {
    if (seg.row === headingRow) continue;
    let best = null;
    for (const [field, list] of Object.entries(dict)) {
      const hit = startsWith(seg, list);
      if (hit && (!best || hit.len - 4 * hit.d > best.len - 4 * best.d)) best = { ...hit, field };
    }
    if (best) {
      seg.label = best;
      hits.push({ field: best.field, seg, used: best.used, phrase: best.phrase.text });
    }
  }

  // A figure read again: the part of the picture it lies in, as one line, by the readings page.js made
  // (options.again); the boxes asked for are gathered in options.collect for it to read.
  const again = new Map((options.again || []).map((a) => [keyOf(a.box), a.readings]));
  const pageW = read.width || Infinity, pageH = read.height || Infinity;
  const readAgain = (box) => {
    const pad = h * 0.5;
    const x = Math.max(0, box.x - pad), y = Math.max(0, box.y - pad * 0.5);
    const b = { x: Math.round(x), y: Math.round(y), width: Math.round(Math.min(pageW, box.x + box.width + pad) - x), height: Math.round(Math.min(pageH, box.y + box.height + pad * 0.5) - y) };
    if (options.collect && !options.collect.some((c) => keyOf(c) === keyOf(b))) options.collect.push(b);
    return again.get(keyOf(b)) || [];
  };
  const FIGURES = new Set(['number', 'date', 'dueDate', 'subtotal', 'tax', 'total']);

  // A label's candidate values: the rest of its segment, the next segments along its row, the segment below;
  // a figure's each read again too, and where none lies beside its label, the stretch of the row it would lie in.
  const parseOpts = { monthFirst: options.monthFirst };
  const candidatesOf = (hit) => {
    const out = [];
    const parse = PARSERS[hit.field];
    const figure = FIGURES.has(hit.field);
    let looked = false;
    const take = (ws, how) => {
      if (!ws.length) return;
      looked = true;
      const text = joinWords(ws.map((w) => w.text));
      const v = parse(text, parseOpts);
      const own = [];
      if (v !== null && v !== undefined) own.push({ parsed: v, text, words: ws, how, confidence: confidenceOf(ws) ?? 0 });
      if (figure) {
        for (const r of readAgain(unionBox(ws))) {
          const v2 = parse(r.text, parseOpts);
          if (v2 !== null && v2 !== undefined) own.push({ parsed: v2, text: r.text, words: ws, how, again: true, confidence: r.confidence ?? 0 });
        }
      }
      // Of one segment's readings, the surest first.
      own.sort((p, q) => q.confidence - p.confidence);
      out.push(...own);
    };
    const rest = after(hit.seg, hit.used);
    if (rest.length && (!figure || rest.some((w) => /\d/.test(westernDigits(w.text))))) take(rest, 'same');
    const row = rows[hit.seg.row];
    for (const next of row.segments.slice(hit.seg.index + 1, hit.seg.index + 3)) {
      if (next.label) break;
      take(next.words, 'beside');
    }
    const b = hit.seg.box;
    for (const below of rows.slice(hit.seg.row + 1, hit.seg.row + 3)) {
      if (below.box.y - (b.y + b.height) > h * 2.5) break;
      const under = below.segments.filter((sg) => Math.min(sg.box.x + sg.box.width, b.x + b.width) - Math.max(sg.box.x, b.x) > -h * 0.5)
        .sort((p, q) => Math.abs((p.box.x + p.box.width / 2) - (b.x + b.width / 2)) - Math.abs((q.box.x + q.box.width / 2) - (b.x + b.width / 2)))[0];
      if (!under) continue;
      if (under.label) break;
      if (!out.length) take(under.words, 'below');
      break;
    }
    // Nothing read beside the label: the row onward from it, which the reading may have passed over.
    if (figure && !out.length && !looked) {
      const lb = unionBox(hit.seg.words);
      const span = Math.min(pageW === Infinity ? h * 20 : pageW * 0.45, h * 20);
      const box = rtl ? { x: Math.max(0, lb.x - span), y: lb.y, width: Math.min(span, lb.x), height: lb.height } : { x: lb.x + lb.width, y: lb.y, width: span, height: lb.height };
      for (const r of readAgain(box)) {
        const v2 = parse(r.text, parseOpts);
        if (v2 !== null && v2 !== undefined) out.push({ parsed: v2, text: r.text, words: [], how: 'row', again: true, confidence: r.confidence ?? 0, box });
      }
    }
    return out;
  };
  for (const hit of hits) hit.candidates = candidatesOf(hit);
  const valueSegs = new Set(hits.flatMap((hit) => hit.candidates.flatMap((c) => c.words)));

  // Items: rows under the heading, or above the first sum on a till roll, ending in an amount.
  const firstSumRow = Math.min(...hits.filter((x) => ['subtotal', 'tax', 'total', 'taxRate'].includes(x.field) && x.seg.row > headingRow).map((x) => x.seg.row), rows.length);
  const items = [];
  const CURRENCY_WORDS = /rs|sar|aed|inr|pkr|usd|gbp|eur|ررس|ر\.س|د\.إ|روپے|रु|円|元/giu;
  const itemOf = (row) => {
    const text = row.segments.map((sg) => sg.text).join('  ');
    if (/(?:^|\s)\d+\s?[x×X*@]\s?\S/.test(westernDigits(text)) && row.segments.length <= 1) return 'detail';
    const lastSeg = row.segments[row.segments.length - 1];
    if (!lastSeg) return null;
    let amount = parseAmount(lastSeg.text);
    if (amount && row.segments.length < 2 && !/\p{L}/u.test(lastSeg.text.replace(CURRENCY_WORDS, ''))) return 'detail';
    // The amount read again (also when the first reading made no figure of it, if the row has a description
    // beside it); the sums choose between the readings.
    const alternatives = amount ? [amount.value] : [];
    if (amount || (row.segments.length >= 2 && /\p{L}/u.test(row.segments[0].text))) {
      for (const r of readAgain(unionBox(lastSeg.words))) {
        const a = parseAmount(r.text);
        if (a && !alternatives.includes(a.value)) { alternatives.push(a.value); amount ||= a; }
      }
    }
    if (!amount) return null;
    const nums = row.segments.slice(1).map((sg) => parseAmount(sg.text)).filter(Boolean).map((a) => a.value);
    const ws = row.segments.flatMap((sg) => sg.words);
    return {
      description: row.segments.length > 1 ? row.segments[0].text : lastSeg.text.replace(/[\d.,\s]+$/, '').trim(),
      quantity: nums.length >= 3 ? nums[nums.length - 3] : null,
      unitPrice: nums.length >= 2 ? nums[nums.length - 2] : null,
      amount: amount.value,
      alternatives,
      decimals: amount.decimals,
      box: unionBox(ws),
      confidence: confidenceOf(ws),
    };
  };
  if (headingRow >= 0) {
    for (const row of rows.slice(headingRow + 1, firstSumRow)) {
      const it = itemOf(row);
      if (it && typeof it === 'object') items.push(it);
    }
  } else if (firstSumRow < rows.length) {
    for (let r = firstSumRow - 1; r >= 0; r--) {
      const row = rows[r];
      if (row.segments.some((sg) => sg.label)) break;
      // A row of unsure words only is noise (a rule, a stain), not the end of the items.
      if (row.segments.every((sg) => sg.words.every((w) => (w.confidence ?? 100) < 45))) continue;
      const it = itemOf(row);
      if (it === 'detail') continue;
      if (!it) {
        // A description on a row of its own above its amount's row is part of the item below it.
        if (items.length && row.segments.length === 1 && /\p{L}/u.test(row.segments[0].text) && !items[0].joined) {
          items[0].description = `${row.segments[0].text} ${items[0].description}`.trim();
          items[0].joined = true;
          continue;
        }
        break;
      }
      items.unshift(it);
    }
  }
  for (const it of items) delete it.joined;

  // The sums choose among the candidates.
  const decimalPage = items.some((it) => it.decimals) || AMOUNT_FIELDS.some((f) => hits.some((x) => x.field === f && x.candidates.some((c) => c.parsed.decimals)));
  const amountsOf = (field) => {
    const out = [];
    const seen = new Set();
    hits.filter((x) => x.field === field).forEach((hit, hi) => hit.candidates.forEach((c, ci) => {
      const add = (value, extra = {}) => {
        if (seen.has(value)) return;
        seen.add(value);
        out.push({ ...c, value, rank: hi + ci + (extra.repaired ? 2 : 0), label: hit.phrase, labelSeg: hit.seg, ...extra });
      };
      add(c.parsed.value);
      // A decimal point the reading lost: 12345 for 123.45.
      if (!c.parsed.decimals && decimalPage) add(c.parsed.value / 100, { repaired: true });
    }));
    return out;
  };
  const cands = Object.fromEntries(AMOUNT_FIELDS.map((f) => [f, amountsOf(f)]));
  const rateHit = hits.find((x) => x.field === 'taxRate' && x.candidates.length);
  // A rate written in the tax's own label ("VAT at 20%").
  const rate = rateHit ? rateHit.candidates[0].parsed : (() => {
    for (const x of hits.filter((y) => y.field === 'tax')) { const p = parsePercent(x.seg.text); if (p !== null) return p; }
    return null;
  })();
  const near = (a, b) => Math.abs(a - b) < 0.005 + Math.abs(b) * 1e-9;
  const cents = (v) => Math.round(v * 100);
  // Every sum the items can make, one reading of each item's amount a sum, the first readings preferred.
  const sums = new Map();
  if (items.length) {
    let states = new Map([[0, []]]);
    const many = items.length > 10;
    for (const it of items) {
      const next = new Map();
      for (const [sum, picks] of states) {
        (many ? [it.alternatives[0]] : it.alternatives).forEach((v, i) => {
          const k = sum + cents(v);
          if (!next.has(k)) next.set(k, [...picks, i]);
        });
      }
      states = next;
    }
    for (const [k, picks] of states) sums.set(k, picks);
  }
  const none = { value: null, rank: 3 };
  let best = null;
  for (const st of [...cands.subtotal, none]) {
    for (const tx of [...cands.tax, none]) {
      for (const tot of [...cands.total, none]) {
        let score = -0.1 * (st.rank + tx.rank + tot.rank);
        const checks = { itemsAddUp: null, totalAddsUp: null, taxRateHolds: null };
        let picks = null;
        if (st.value !== null && tx.value !== null && tot.value !== null) checks.totalAddsUp = near(st.value + tx.value, tot.value);
        const itemsTo = st.value !== null ? st.value : tx.value === null && tot.value !== null ? tot.value : null;
        if (items.length && itemsTo !== null) {
          picks = sums.get(cents(itemsTo)) || null;
          checks.itemsAddUp = Boolean(picks);
        }
        if (rate !== null && st.value !== null && tx.value !== null) {
          const exact = st.value * rate / 100;
          checks.taxRateHolds = Math.abs(exact - tx.value) <= 0.0051 || Math.floor(exact) === tx.value;
        }
        for (const v of Object.values(checks)) if (v !== null) score += v ? 3 : -1;
        if (st.value !== null) score += 0.5;
        if (tx.value !== null) score += 0.5;
        if (tot.value !== null) score += 1;
        if (!best || score > best.score) best = { score, st, tx, tot, checks, picks };
      }
    }
  }
  // The items as the sums read them.
  if (best.picks) best.picks.forEach((i, k) => { items[k].amount = items[k].alternatives[i]; });

  const fields = {};
  const put = (field, c, checked = null) => {
    if (!c || c.value === null && c.parsed === undefined) return;
    fields[field] = {
      value: c.value !== undefined ? c.value : c.parsed,
      text: c.text, box: c.words?.length ? unionBox(c.words) : c.box || null, confidence: c.words ? confidenceOf(c.words) : null,
      label: c.label ?? null, checked, ...(c.repaired ? { repaired: true } : {}), ...(c.derived ? { derived: true } : {}),
    };
  };
  const ch = best.checks;
  const anyCheck = Object.values(ch).some((v) => v !== null);
  const consistent = anyCheck && !Object.values(ch).includes(false);
  if (best.st.value !== null) put('subtotal', best.st, ch.itemsAddUp ?? ch.totalAddsUp ?? ch.taxRateHolds);
  if (best.tx.value !== null) put('tax', best.tx, ch.totalAddsUp ?? ch.taxRateHolds);
  if (best.tot.value !== null) put('total', best.tot, ch.totalAddsUp ?? (best.st.value === null && best.tx.value === null ? ch.itemsAddUp : null));
  // A total the page lost but whose parts it kept, when the parts agree with the items.
  if (!fields.total && fields.subtotal && fields.tax && ch.itemsAddUp) {
    fields.total = { value: Math.round((fields.subtotal.value + fields.tax.value) * 100) / 100, text: null, box: null, confidence: null, label: null, checked: true, derived: true };
  }
  if (rate !== null) fields.taxRate = { value: rate, text: rateHit?.candidates[0].text ?? null, box: rateHit ? unionBox(rateHit.candidates[0].words) : null, confidence: rateHit ? confidenceOf(rateHit.candidates[0].words) : null, label: rateHit?.phrase ?? null, checked: ch.taxRateHolds };

  // The other fields: the first label with a value; dates must not be each other.
  for (const field of ['number', 'date', 'dueDate', 'supplier']) {
    for (const hit of hits.filter((x) => x.field === field)) {
      const c = hit.candidates[0];
      if (!c) continue;
      put(field, { ...c, label: hit.phrase });
      break;
    }
  }

  // Who it is from, with no label: the largest writing near the top that is not a title, a label, a value or the customer.
  const kindDicts = Object.fromEntries(Object.entries(KINDS).map(([k, d]) => [k, phrases(d, languages)]));
  const titles = [...kindDicts.invoice, ...kindDicts.receipt, ...kindDicts.form].filter((p) => p.folded.length >= 2);
  const cjkPage = words.filter((w) => CJK.test(w.text)).length > words.length / 3;
  if (!fields.supplier) {
    const billTo = hits.filter((x) => x.field === 'billTo');
    const customer = new Set(billTo.flatMap((x) => [...x.candidates.flatMap((c) => c.words), ...rows.slice(x.seg.row + 1, x.seg.row + 3).flatMap((r) => r.segments.filter((s) => Math.abs(s.box.x - x.seg.box.x) < h * 3 || Math.abs(s.box.x + s.box.width - x.seg.box.x - x.seg.box.width) < h * 3).flatMap((s) => s.words))]));
    const firstLabelRow = Math.min(...hits.filter((x) => x.field !== 'supplier').map((x) => x.seg.row), rows.length);
    const top = Math.max(rows.length ? rows[Math.min(firstLabelRow, rows.length - 1)].box.y : 0, (read.height || 0) * 0.12);
    let pick = null;
    for (const seg of segs) {
      if (seg.box.y > top + h || seg.label || seg.words.some((w) => valueSegs.has(w) || customer.has(w))) continue;
      const f = fold(seg.text);
      const letters = (seg.text.match(/\p{L}/gu) || []).length;
      const digits = (westernDigits(seg.text).match(/\d/g) || []).length;
      if (letters < 2 || digits > letters / 2) continue;
      // A title set in spaced capitals ("I N V O I C E") reads as letters one by one (not so Chinese and Japanese).
      const toks = seg.words.map((w) => fold(w.text)).filter(Boolean);
      if (!CJK.test(f) && toks.length >= 2 && toks.reduce((n, t) => n + t.length, 0) / toks.length < 2) continue;
      // The title, whole, in part or misread by a letter; on a page in Chinese or Japanese, its characters alone
      // (the engine reads a large title's strokes as stray Latin letters around them).
      const core = cjkPage ? f.replace(/[^぀-ヿ㐀-鿿豈-﫿]/gu, '') : f;
      if (core && titles.some((p) => p.folded.includes(core) || core.includes(p.folded) || distance(core, p.folded) <= (p.folded.length >= 5 ? 1 : 0))) continue;
      // Writing the reading is unsure of is a stain or a picture read as letters.
      if (seg.words.reduce((n, w) => n + (w.confidence ?? 100), 0) / seg.words.length < 60) continue;
      // Writing in the page's own script first, then the largest, then the highest.
      const own = cjkPage ? CJK.test(f) : rtl ? /[؀-ۿ]/u.test(seg.text) : true;
      const size = seg.height;
      if (!pick || (own && !pick.own) || (own === pick.own && (size > pick.height * 1.08 || (size > pick.height * 0.92 && seg.box.y < pick.box.y)))) pick = { ...seg, own };
    }
    if (pick) {
      // A name the row's gaps cut in two: the pieces beside it of its own size, closer than three letters' height.
      const row = rows[pick.row];
      const ws = [...pick.words];
      for (let i = pick.index + 1; i < row.segments.length; i++) {
        const sg = row.segments[i], prev = row.segments[i - 1];
        const gap = rtl ? prev.box.x - (sg.box.x + sg.box.width) : sg.box.x - (prev.box.x + prev.box.width);
        if (sg.label || gap > h * 3 || Math.abs(sg.height - pick.height) > pick.height * 0.25) break;
        ws.push(...sg.words);
      }
      for (let i = pick.index - 1; i >= 0; i--) {
        const sg = row.segments[i], next = row.segments[i + 1];
        const gap = rtl ? sg.box.x - (next.box.x + next.box.width) : next.box.x - (sg.box.x + sg.box.width);
        if (sg.label || gap > h * 3 || Math.abs(sg.height - pick.height) > pick.height * 0.25) break;
        ws.unshift(...sg.words);
      }
      const text = joinWords(ws.map((w) => w.text)).trim();
      put('supplier', { parsed: text, text, words: ws, label: null });
    }
  }

  // Currency: the one the amounts name most.
  const named = {};
  for (const f of AMOUNT_FIELDS) { const c = fields[f]?.text && parseAmount(fields[f].text)?.currency; if (c) named[c] = (named[c] || 0) + 1; }
  let currency = Object.entries(named).sort((a, b) => b[1] - a[1])[0]?.[0] || null;
  if (currency === 'YEN') currency = (languages || []).includes('jpn') ? 'JPY' : (languages || []).includes('chi_sim') ? 'CNY' : null;

  // What kind of page this is.
  const allTokens = segs.map((s) => s.words.map((w) => fold(w.text)).filter(Boolean));
  const has = (list) => {
    const found = new Set();
    for (const p of list) {
      for (const toks of allTokens) {
        const s = toks.join('');
        if (CJK.test(p.folded)) { if (s.includes(p.folded)) found.add(p.folded); continue; }
        // Whole words: the phrase must begin and end at word edges.
        let at = 0;
        for (let i = 0; i < toks.length; i++) {
          if (s.slice(at).startsWith(p.folded)) {
            let n = at, j = i;
            while (j < toks.length && n < at + p.folded.length) n += toks[j++].length;
            if (n === at + p.folded.length) { found.add(p.folded); break; }
          }
          at += toks[i].length;
        }
      }
    }
    return found.size;
  };
  const fieldHits = (f) => hits.some((x) => x.field === f);
  const width = read.width || 0, height = read.height || 0;
  const time = /(?:^|\D)([01]?\d|2[0-3]):[0-5]\d(?!\d)/.test(westernDigits(segs.map((s) => s.text).join(' ')));
  const longRows = rows.filter((r) => r.segments.some((s) => s.words.length >= 8 || (width && s.box.width > width * 0.55 && s.words.length >= 5))).length;
  const scores = {
    invoice: (has(kindDicts.invoice) ? 3 : 0) + (fieldHits('dueDate') ? 2 : 0) + (fieldHits('billTo') ? 1 : 0) + (fieldHits('number') ? 1 : 0) + (headingRow >= 0 ? 1 : 0) + (fields.total ? 1 : 0),
    receipt: (has(kindDicts.receipt) ? 3 : 0) + (fieldHits('paid') ? 2 : 0) + (width && height && width / height < 0.6 ? 2 : 0) + (time ? 1 : 0) + (fields.total ? 1 : 0),
    letter: (has(kindDicts.greeting) ? 2 : 0) + (has(kindDicts.closing) ? 2 : 0) + (longRows >= 3 ? 1 : 0) - (fields.total ? 2 : 0),
    form: (has(kindDicts.form) ? 2 : 0) + Math.min(3, has(kindDicts.formField)) + ((options.checkboxes || []).length ? 1 : 0) - (fields.total ? 2 : 0),
  };
  const order = ['invoice', 'receipt', 'form', 'letter'];
  let kind = order.reduce((a, b) => (scores[b] > scores[a] ? b : a));
  if (scores[kind] < 2) kind = 'other';

  return {
    kind, scores, fields,
    items: items.map(({ decimals, alternatives, ...it }) => it),
    checks: anyCheck ? { ...ch, consistent } : (fields.total ? { ...ch, consistent: false } : null),
    currency,
    slope,
  };
}
