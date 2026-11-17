import type { Lang } from './i18n';
import type { DocData, OpenFile } from './types';
import { addr, newSheet, type Cell, type Workbook } from './editors/sheet/model';
import { makeSlide, type Deck, type TextEl } from './editors/slides/model';

/** Starter files for the start screen. Content follows the interface language. */
export interface Template {
  id: string;
  kind: 'doc' | 'sheet' | 'slides';
  name: Record<Lang, string>;
  build: (lang: Lang) => Omit<OpenFile, 'id'>;
}

const L = (lang: Lang, en: string, ar: string) => (lang === 'ar' ? ar : en);
const doc = (name: string, html: string): Omit<OpenFile, 'id'> => ({ kind: 'doc', name, data: { html, page: { orientation: 'portrait', margin: 'normal' } } as DocData });
const today = (lang: Lang) => new Date().toLocaleDateString(lang === 'ar' ? 'ar' : 'en-GB', { day: 'numeric', month: 'long', year: 'numeric' });

function letter(lang: Lang) {
  const a = lang === 'ar';
  const html = a
    ? `<p style="text-align: right">${today(lang)}</p><p><strong>إلى:</strong> اسم المستلم<br>المسمى الوظيفي، اسم الجهة<br>العنوان</p><p><strong>الموضوع: </strong>اكتب موضوع الرسالة هنا</p><p>السيد/السيدة المحترم/ة،</p><p>تحية طيبة وبعد،</p><p>اكتب الفقرة الأولى هنا لتعريف سبب الرسالة. اجعلها قصيرة وواضحة.</p><p>استخدم الفقرة الثانية للتفاصيل: ما الذي تطلبه أو تقترحه، ولماذا، وما الخطوة التالية.</p><p>شكرًا لوقتكم واهتمامكم.</p><p>مع خالص التحية،</p><p><strong>اسمك</strong><br>المسمى الوظيفي · رقم الهاتف · البريد الإلكتروني</p>`
    : `<p style="text-align: right">${today(lang)}</p><p><strong>To:</strong> Recipient name<br>Title, Organization<br>Street address, City</p><p><strong>Subject: </strong>Write the subject of your letter here</p><p>Dear Recipient,</p><p>Open with one or two sentences that say why you are writing. Keep it short and clear.</p><p>Use the second paragraph for the details: what you are asking for or proposing, why it matters, and what the next step is.</p><p>Thank you for your time and attention.</p><p>Kind regards,</p><p><strong>Your Name</strong><br>Title · Phone · Email</p>`;
  return doc(L(lang, 'Letter', 'رسالة'), html);
}

function report(lang: Lang) {
  const a = lang === 'ar';
  const html = a
    ? `<h1>عنوان التقرير</h1><p><em>أعده: اسمك · ${today(lang)}</em></p><h2>الملخص</h2><p>اكتب في ثلاث جمل ما الذي وجدته وما الذي توصي به.</p><h2>الخلفية</h2><p>ما المشكلة أو السؤال؟ ولماذا يهم الآن؟</p><h2>النتائج</h2><ul><li><p>النتيجة الأولى، مع رقم يدعمها.</p></li><li><p>النتيجة الثانية.</p></li><li><p>النتيجة الثالثة.</p></li></ul><table><tbody><tr><th><p>المؤشر</p></th><th><p>الربع السابق</p></th><th><p>هذا الربع</p></th></tr><tr><td><p>الإيرادات</p></td><td><p>120,000</p></td><td><p>138,000</p></td></tr><tr><td><p>العملاء</p></td><td><p>840</p></td><td><p>910</p></td></tr></tbody></table><h2>التوصيات</h2><ol><li><p>التوصية الأولى.</p></li><li><p>التوصية الثانية.</p></li></ol>`
    : `<h1>Report title</h1><p><em>Prepared by Your Name · ${today(lang)}</em></p><h2>Summary</h2><p>In three sentences, say what you found and what you recommend.</p><h2>Background</h2><p>What is the problem or question, and why does it matter now?</p><h2>Findings</h2><ul><li><p>First finding, with a number that backs it up.</p></li><li><p>Second finding.</p></li><li><p>Third finding.</p></li></ul><table><tbody><tr><th><p>Metric</p></th><th><p>Last quarter</p></th><th><p>This quarter</p></th></tr><tr><td><p>Revenue</p></td><td><p>120,000</p></td><td><p>138,000</p></td></tr><tr><td><p>Customers</p></td><td><p>840</p></td><td><p>910</p></td></tr></tbody></table><h2>Recommendations</h2><ol><li><p>First recommendation.</p></li><li><p>Second recommendation.</p></li></ol>`;
  return doc(L(lang, 'Report', 'تقرير'), html);
}

function resume(lang: Lang) {
  const a = lang === 'ar';
  const html = a
    ? `<h1>اسمك الكامل</h1><p>المسمى الوظيفي · المدينة · رقم الهاتف · البريد الإلكتروني</p><hr><h2>نبذة</h2><p>جملتان عن خبرتك وما تبرع فيه وما تبحث عنه.</p><h2>الخبرات</h2><h3>المسمى الوظيفي — اسم الشركة</h3><p><em>2022 – الآن</em></p><ul><li><p>إنجاز قابل للقياس بدأته بفعل قوي.</p></li><li><p>إنجاز آخر مع رقم.</p></li></ul><h3>المسمى الوظيفي — اسم الشركة</h3><p><em>2019 – 2022</em></p><ul><li><p>إنجاز.</p></li></ul><h2>التعليم</h2><p><strong>الدرجة العلمية</strong> — الجامعة، 2019</p><h2>المهارات</h2><p>مهارة · مهارة · مهارة · اللغات</p>`
    : `<h1>Your Full Name</h1><p>Job title · City · Phone · Email</p><hr><h2>Profile</h2><p>Two sentences on your experience, what you are great at, and what you are looking for.</p><h2>Experience</h2><h3>Job title — Company</h3><p><em>2022 – present</em></p><ul><li><p>A measurable achievement that starts with a strong verb.</p></li><li><p>Another achievement, with a number.</p></li></ul><h3>Job title — Company</h3><p><em>2019 – 2022</em></p><ul><li><p>An achievement.</p></li></ul><h2>Education</h2><p><strong>Degree</strong> — University, 2019</p><h2>Skills</h2><p>Skill · Skill · Skill · Languages</p>`;
  return doc(L(lang, 'Résumé', 'سيرة ذاتية'), html);
}

function meeting(lang: Lang) {
  const a = lang === 'ar';
  const html = a
    ? `<h1>ملاحظات الاجتماع</h1><p><strong>التاريخ:</strong> ${today(lang)} · <strong>الحضور:</strong> الأسماء</p><h2>جدول الأعمال</h2><ol><li><p>البند الأول</p></li><li><p>البند الثاني</p></li></ol><h2>القرارات</h2><ul><li><p>القرار ومن يملكه.</p></li></ul><h2>المهام</h2><table><tbody><tr><th><p>المهمة</p></th><th><p>المسؤول</p></th><th><p>الموعد</p></th></tr><tr><td><p>…</p></td><td><p>…</p></td><td><p>…</p></td></tr></tbody></table>`
    : `<h1>Meeting notes</h1><p><strong>Date:</strong> ${today(lang)} · <strong>Attendees:</strong> Names</p><h2>Agenda</h2><ol><li><p>First topic</p></li><li><p>Second topic</p></li></ol><h2>Decisions</h2><ul><li><p>The decision and who owns it.</p></li></ul><h2>Action items</h2><table><tbody><tr><th><p>Task</p></th><th><p>Owner</p></th><th><p>Due</p></th></tr><tr><td><p>…</p></td><td><p>…</p></td><td><p>…</p></td></tr></tbody></table>`;
  return doc(L(lang, 'Meeting notes', 'ملاحظات اجتماع'), html);
}

function sheetFrom(name: string, sheetName: string, rows: (string | number | null)[][], opts: { head?: boolean; widths?: number[]; money?: number[]; totalRow?: number } = {}): Omit<OpenFile, 'id'> {
  const sh = newSheet(sheetName);
  rows.forEach((row, r) =>
    row.forEach((v, c) => {
      if (v === null || v === '') return;
      const cell: Cell = { v: String(v) };
      if (opts.head && r === 0) cell.st = { b: true, bg: '#e8f6ef', color: '#065f46', border: true };
      else if (opts.totalRow === r) cell.st = { b: true, border: true };
      if (r > 0 && opts.money?.includes(c)) cell.st = { ...cell.st, fmt: 'number', dec: 2 };
      sh.cells[addr(r, c)] = cell;
    }),
  );
  opts.widths?.forEach((w, c) => (sh.colW[c] = w));
  sh.freezeTop = !!opts.head;
  const wb: Workbook = { sheets: [sh], active: 0 };
  return { kind: 'sheet', name, data: wb };
}

function budget(lang: Lang) {
  const a = lang === 'ar';
  const rows: (string | number | null)[][] = [
    [L(lang, 'Category', 'البند'), L(lang, 'Planned', 'المخطط'), L(lang, 'Actual', 'الفعلي'), L(lang, 'Difference', 'الفرق')],
    ...(a
      ? [['السكن', 1500, 1500], ['الطعام', 600, 645], ['المواصلات', 250, 210], ['الفواتير', 300, 288], ['الترفيه', 200, 260], ['الادخار', 500, 500]]
      : [['Housing', 1500, 1500], ['Food', 600, 645], ['Transport', 250, 210], ['Utilities', 300, 288], ['Fun', 200, 260], ['Savings', 500, 500]]
    ).map((r, i) => [...r, `=B${i + 2}-C${i + 2}`]),
    [L(lang, 'Total', 'المجموع'), '=SUM(B2:B7)', '=SUM(C2:C7)', '=SUM(D2:D7)'],
  ];
  return sheetFrom(L(lang, 'Monthly budget', 'ميزانية شهرية'), L(lang, 'Budget', 'الميزانية'), rows, { head: true, widths: [160, 110, 110, 110], money: [1, 2, 3], totalRow: 7 });
}

function invoice(lang: Lang) {
  const rows: (string | number | null)[][] = [
    [L(lang, 'Item', 'الصنف'), L(lang, 'Qty', 'الكمية'), L(lang, 'Unit price', 'سعر الوحدة'), L(lang, 'Amount', 'المبلغ')],
    [L(lang, 'Design work', 'أعمال التصميم'), 10, 45, '=B2*C2'],
    [L(lang, 'Development', 'التطوير'), 24, 60, '=B3*C3'],
    [L(lang, 'Hosting (1 year)', 'الاستضافة (سنة)'), 1, 120, '=B4*C4'],
    [null, null, L(lang, 'Subtotal', 'المجموع الفرعي'), '=SUM(D2:D4)'],
    [null, null, L(lang, 'Tax 15%', 'الضريبة 15%'), '=D5*0.15'],
    [null, null, L(lang, 'Total', 'الإجمالي'), '=D5+D6'],
  ];
  return sheetFrom(L(lang, 'Invoice', 'فاتورة'), L(lang, 'Invoice', 'الفاتورة'), rows, { head: true, widths: [200, 80, 120, 120], money: [2, 3], totalRow: 6 });
}

function tracker(lang: Lang) {
  const rows: (string | number | null)[][] = [
    [L(lang, 'Task', 'المهمة'), L(lang, 'Owner', 'المسؤول'), L(lang, 'Due', 'الموعد'), L(lang, 'Status', 'الحالة'), L(lang, 'Done', 'منجز')],
    [L(lang, 'Write the plan', 'كتابة الخطة'), 'Sara', '2026-12-01', L(lang, 'In progress', 'قيد التنفيذ'), 'FALSE'],
    [L(lang, 'Review budget', 'مراجعة الميزانية'), 'Omar', '2026-12-05', L(lang, 'Not started', 'لم تبدأ'), 'FALSE'],
    [L(lang, 'Send invitations', 'إرسال الدعوات'), 'Lina', '2026-11-28', L(lang, 'Done', 'منجزة'), 'TRUE'],
    [null],
    [L(lang, 'Completed', 'المنجز'), '=COUNTIF(E2:E4,TRUE)&" / "&COUNTA(A2:A4)'],
  ];
  return sheetFrom(L(lang, 'Task tracker', 'متابعة المهام'), L(lang, 'Tasks', 'المهام'), rows, { head: true, widths: [200, 110, 110, 130, 70] });
}

function fill(deck: Deck, layout: Parameters<typeof makeSlide>[0], texts: string[]) {
  const s = makeSlide(layout, deck);
  s.els.filter((e): e is TextEl => e.type === 'text').forEach((e, i) => (e.text = texts[i] ?? ''));
  deck.slides.push(s);
}

function pitch(lang: Lang): Omit<OpenFile, 'id'> {
  const deck: Deck = { w: 960, h: 540, theme: 'night', slides: [] };
  fill(deck, 'title', [L(lang, 'Company name', 'اسم الشركة'), L(lang, 'One line on what you do and for whom', 'سطر واحد عمّا تقدمه ولمن')]);
  fill(deck, 'content', [L(lang, 'The problem', 'المشكلة'), L(lang, 'Who has it\nHow they solve it today\nWhy that is not good enough', 'من يعاني منها\nكيف يحلونها اليوم\nلماذا لا يكفي ذلك')]);
  fill(deck, 'content', [L(lang, 'Our solution', 'الحل'), L(lang, 'What the product does\nThe moment it clicks for users\nWhat makes it hard to copy', 'ماذا يفعل المنتج\nاللحظة التي يقتنع فيها المستخدم\nما الذي يصعب تقليده')]);
  fill(deck, 'two', [L(lang, 'Market and traction', 'السوق والنمو'), L(lang, 'Market size\nTarget customers', 'حجم السوق\nالعملاء المستهدفون'), L(lang, 'Users and revenue\nGrowth per month', 'المستخدمون والإيرادات\nالنمو الشهري')]);
  fill(deck, 'content', [L(lang, 'The ask', 'ما نطلبه'), L(lang, 'How much we are raising\nWhat it pays for\nMilestones it reaches', 'المبلغ المطلوب\nفيمَ سيُصرف\nالمراحل التي سنبلغها')]);
  return { kind: 'slides', name: L(lang, 'Pitch deck', 'عرض استثماري'), data: deck };
}

function lesson(lang: Lang): Omit<OpenFile, 'id'> {
  const deck: Deck = { w: 960, h: 540, theme: 'mint', slides: [] };
  fill(deck, 'title', [L(lang, 'Lesson title', 'عنوان الدرس'), L(lang, 'Class · Teacher name', 'الصف · اسم المعلم')]);
  fill(deck, 'content', [L(lang, 'Learning goals', 'أهداف الدرس'), L(lang, 'By the end, you will be able to…\n…explain\n…apply', 'في نهاية الدرس ستكون قادرًا على…\n…أن تشرح\n…أن تطبق')]);
  fill(deck, 'two', [L(lang, 'Key idea', 'الفكرة الرئيسية'), L(lang, 'Definition\nExample', 'التعريف\nمثال'), L(lang, 'Common mistake\nHow to avoid it', 'خطأ شائع\nكيف نتجنبه')]);
  fill(deck, 'content', [L(lang, 'Practice', 'تمرين'), L(lang, 'Question 1\nQuestion 2\nQuestion 3', 'السؤال 1\nالسؤال 2\nالسؤال 3')]);
  return { kind: 'slides', name: L(lang, 'Lesson', 'درس'), data: deck };
}

export const TEMPLATES: Template[] = [
  { id: 'letter', kind: 'doc', name: { en: 'Letter', ar: 'رسالة' }, build: letter },
  { id: 'report', kind: 'doc', name: { en: 'Report', ar: 'تقرير' }, build: report },
  { id: 'resume', kind: 'doc', name: { en: 'Résumé', ar: 'سيرة ذاتية' }, build: resume },
  { id: 'meeting', kind: 'doc', name: { en: 'Meeting notes', ar: 'ملاحظات اجتماع' }, build: meeting },
  { id: 'budget', kind: 'sheet', name: { en: 'Monthly budget', ar: 'ميزانية شهرية' }, build: budget },
  { id: 'invoice', kind: 'sheet', name: { en: 'Invoice', ar: 'فاتورة' }, build: invoice },
  { id: 'tracker', kind: 'sheet', name: { en: 'Task tracker', ar: 'متابعة المهام' }, build: tracker },
  { id: 'pitch', kind: 'slides', name: { en: 'Pitch deck', ar: 'عرض استثماري' }, build: pitch },
  { id: 'lesson', kind: 'slides', name: { en: 'Lesson', ar: 'درس' }, build: lesson },
];
