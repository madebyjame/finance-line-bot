/**
 * LINE Bot สำหรับบันทึกรายรับ-รายจ่าย อัตโนมัติลงใน Google Sheets
 * - บันทึกรายรับ/รายจ่ายจากแชต (รองรับเลขใหญ่/คอมม่า)
 * - วิเคราะห์ด้วย Gemini (ต่อผู้ใช้)
 * - แดชบอร์ดพร้อมปุ่มช่วงเวลา (ต่อผู้ใช้)
 * - แยกข้อมูลเป็น "แท็บชีตต่อผู้ใช้" อัตโนมัติ
 */

const fs = require("fs");
if (process.env.GOOGLE_CREDENTIALS_JSON) {
  const path = "/tmp/google.json";
  fs.writeFileSync(path, process.env.GOOGLE_CREDENTIALS_JSON);
  process.env.GOOGLE_SERVICE_ACCOUNT_FILE = path;
}

require('dotenv').config();
const express = require('express');
const { Client, middleware } = require('@line/bot-sdk');
const { google } = require('googleapis');
const axios = require('axios');

const app = express();

const lineConfig = {
  channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.CHANNEL_SECRET
};
const lineClient = new Client(lineConfig);

// ===== THEME (Modern UI) =====
const THEME = {
  accent: '#4F46E5',
  accentSoft: '#EEF2FF',
  danger: '#EF4444',
  textMuted: '#8B95A1',
  textStrong: '#111827',
};

// -------- Google Sheets (ฐานกลาง) --------
const auth = new google.auth.GoogleAuth({
  keyFile: process.env.GOOGLE_SERVICE_ACCOUNT_FILE,
  scopes: ['https://www.googleapis.com/auth/spreadsheets']
});
const sheets = google.sheets({ version: 'v4', auth });
const SHEET_ID = process.env.GOOGLE_SHEET_ID;

// (ยังเก็บ helper เก่าไว้ เผื่อใช้ endpoint อื่นในอนาคต)
async function readRecentRows(limit = 120, range = 'Sheet1!A:E') {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range
  });
  const rows = res.data.values || [];
  const dataRows = rows.length > 1 ? rows.slice(1) : [];
  return dataRows.slice(-limit);
}

// ============================ PER-USER SHEET HELPERS ============================
// ตั้งชื่อแท็บจาก userId (ปลอดภัยและสั้น)
function sheetTitleForUser(userId) {
  const uid = String(userId || 'unknown');
  const tail = uid.slice(-8).replace(/[^A-Za-z0-9_-]/g, '');
  return `U_${tail}`; // เช่น U_a1B2c3D4
}

// รายชื่อแท็บทั้งหมด
async function listSheetTitles() {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  return (meta.data.sheets || []).map(s => s.properties?.title).filter(Boolean);
}

// สร้างแท็บ + หัวตาราง ถ้ายังไม่มี
async function ensureUserSheet(userId) {
  const title = sheetTitleForUser(userId);
  const titles = await listSheetTitles();
  if (!titles.includes(title)) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title } } }] }
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${title}!A1:F1`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [['วันที่', 'ประเภท', 'จำนวน', 'หมวดหมู่', 'รายละเอียด', 'userId']] }
    });
  }
  return title;
}

// บันทึกลงแท็บของผู้ใช้
async function appendRowToUser(userId, values) {
  const title = await ensureUserSheet(userId);
  return sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: `${title}!A:F`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [values] }
  });
}

// อ่านข้อมูลเฉพาะผู้ใช้
async function readRecentRowsForUser(userId, limit = 1000) {
  const title = await ensureUserSheet(userId);
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${title}!A:F`
  });
  const rows = res.data.values || [];
  const dataRows = rows.length > 1 ? rows.slice(1) : [];
  return dataRows.slice(-limit);
}
// ============================ END PER-USER SHEET HELPERS ============================

// แปลงวันที่ไทย → YYYY-MM (คีย์รายเดือน)
function monthKeyFromThaiDate(s) {
  const d = parseThaiDate(s);
  if (!d) return 'unknown';
  const y = d.getFullYear();
  const m = (d.getMonth()+1).toString().padStart(2, '0');
  return `${y}-${m}`;
}

// สร้างข้อเท็จจริง (facts) สำหรับช่วง N เดือนล่าสุด
async function buildFinanceFactsForUser(userId, months = 3) {
  const rows = await readRecentRowsForUser(userId, 1000); // ใช้ข้อมูลพอประมาณ
  if (!rows || rows.length === 0) return { rows: [], facts: null, text: 'NO_DATA' };

  // เลือกเฉพาะ N เดือนล่าสุด
  const now = new Date();
  const cutoff = new Date(now.getFullYear(), now.getMonth() - (months - 1), 1);
  const filtered = rows.filter(r => {
    const d = parseThaiDate(r[0]);
    return d && d >= cutoff;
  });

  // สรุปตัวเลข
  let income = 0, expense = 0;
  const byCategory = {};      // {cat: sumExpense}
  const byMonth = {};         // {YYYY-MM: {income, expense}}
  const singles = [];         // รายการเดี่ยว (สำหรับ outlier)

  for (const r of filtered) {
    const [date, type, amountRaw, categoryRaw] = r;
    const amt = Number(String(amountRaw || '0').toString().replace(/,/g, '')) || 0;
    const cat = categoryRaw || 'อื่นๆ';
    const mk = monthKeyFromThaiDate(date);
    if (!byMonth[mk]) byMonth[mk] = { income: 0, expense: 0 };

    if (type === 'รายรับ') {
      income += amt;
      byMonth[mk].income += amt;
    } else if (type === 'รายจ่าย') {
      expense += amt;
      byMonth[mk].expense += amt;
      byCategory[cat] = (byCategory[cat] || 0) + amt;
      singles.push({ date, category: cat, amount: amt });
    }
  }

  // จัดอันดับหมวดรายจ่ายตามสัดส่วน
  const categoryShares = Object.entries(byCategory)
    .map(([cat, sum]) => ({
      category: cat,
      sum,
      share: expense > 0 ? sum / expense : 0
    }))
    .sort((a,b) => b.sum - a.sum);

  // ค่ากลางของรายจ่ายเดี่ยว (สำหรับดู outlier)
  const amounts = singles.map(s => s.amount).sort((a,b) => a-b);
  const median = amounts.length ? (amounts[Math.floor(amounts.length/2)] + amounts[Math.ceil(amounts.length/2)-1]) / 2 : 0;
  const mean = amounts.length ? (amounts.reduce((a,b)=>a+b,0)/amounts.length) : 0;

  // ระบุ “หมวดที่น่าพิจารณา” ด้วยกติกา
  const interestingCats = categoryShares.filter(x =>
    x.share >= FINANCE_RULES.CATEGORY_SHARE_ALERT ||
    x.sum >= FINANCE_RULES.MIN_MONTHLY_CATEGORY_SUM
  );

  // คัดรายการเดี่ยวที่ “ใหญ่ผิดปกติ” (ข้ามยอดเล็ก)
  const outliers = singles.filter(s =>
    s.amount >= FINANCE_RULES.MIN_SINGLE_IGNORE &&
    (median > 0 ? s.amount >= FINANCE_RULES.OUTLIER_MULTIPLIER * median : s.amount >= FINANCE_RULES.MIN_MONTHLY_CATEGORY_SUM)
  ).sort((a,b) => b.amount - a.amount).slice(0, 10);

  // month-over-month (ล่าสุดเทียบก่อนหน้า)
  const keys = Object.keys(byMonth).sort();
  const lastKey = keys[keys.length - 1];
  const prevKey = keys[keys.length - 2];
  const mom = (lastKey && prevKey) ? {
    last: { key: lastKey, ...byMonth[lastKey] },
    prev: { key: prevKey, ...byMonth[prevKey] },
    diff: {
      income: (byMonth[lastKey].income - byMonth[prevKey].income),
      expense: (byMonth[lastKey].expense - byMonth[prevKey].expense),
    }
  } : null;

  const facts = {
    monthsAnalyzed: months,
    totalIncome: income,
    totalExpense: expense,
    balance: income - expense,
    medianSingleExpense: median,
    meanSingleExpense: mean,
    categoryShares,        // เรียงจากมากไปน้อย
    interestingCats,       // หมวดที่เข้าหลักเกณฑ์
    outliers,              // รายการเดี่ยวที่ผิดปกติ
    monthOverMonth: mom,   // เทียบเดือนล่าสุดกับก่อนหน้า
    rules: FINANCE_RULES,
  };

  // พร้อมข้อความแถวๆ สำหรับ AI (แบบกะทัดรัด)
  const text = rows.map(r => {
    const [date, type, amount, category] = r;
    return `${date} | ${type} | ${amount} | ${category || '-'}`;
  }).join('\n');

  return { rows: filtered, facts, text };
}


// -------- Helpers: Date / Format --------
function todayTH() {
  return new Date().toLocaleDateString('th-TH');
}

// -------- Classifier: Type & Category (TH) --------
const KW = {
  income: [ 'รายรับ','รับ','ได้','โอนเข้า','เงินเดือน','โบนัส','ทิป','ขายได้','ดอกเบี้ย','ปันผล' ],
  expense: [ 'รายจ่าย','จ่าย','ซื้อ','โอนออก','เติม','ค่าผ่อน','ผ่อน','ค่าสมาชิก','ค่าใช้จ่าย' ],
};

const CATS_EXPENSE = [
  { name: 'อาหาร/กาแฟ', kws: ['ข้าว','อาหาร','ข้าวเที่ยง','ข้าวเย็น','ของกิน','กาแฟ','คาเฟ่','ชานม','ของหวาน','ส้มตำ','ก๋วยเตี๋ยว','หมูกระทะ'] },
  { name: 'เดินทาง', kws: ['รถ','น้ำมัน','เติมน้ำมัน','มอเตอร์ไซค์','แท็กซี่','แกร็บ','บีทีเอส','mrt','ตั๋ว','ค่าทางด่วน','ที่จอด'] },
  { name: 'บิล/สาธารณูปโภค', kws: ['ค่าไฟ','ค่าน้ำ','ค่าเน็ต','เน็ตบ้าน','โทรศัพท์','ค่าโทร','ค่าสาธารณูปโภค','ค่าเช่า','ค่าเช่าบ้าน','คอนโด','ค่าบ้าน'] },
  { name: 'ช้อปปิ้ง/ของใช้', kws: ['ช้อป','สั่งของ','shopee','lazada','tiktok','เสื้อ','กางเกง','รองเท้า','ของใช้','อุปกรณ์','แก็ดเจ็ต'] },
  { name: 'สุขภาพ/ประกัน', kws: ['ยา','โรงพยาบาล','คลินิก','ประกัน','ฟิตเนส','ตรวจสุขภาพ','วิตามิน'] },
  { name: 'การศึกษา/งาน', kws: ['คอร์ส','เรียน','หนังสือ','หลักสูตร','อบรม','ค่าสอบ','ซอฟต์แวร์','ไลเซนส์','microsoft','adobe'] },
  { name: 'บันเทิง', kws: ['หนัง','netflix','spotify','เกม','สตรีม','คอนเสิร์ต'] },
];

const CATS_INCOME = [
  { name: 'เงินเดือน', kws: ['เงินเดือน','salary','เดือนนี้','ค่าจ้าง'] },
  { name: 'โบนัส/ทิป', kws: ['โบนัส','bonus','ทิป','tip'] },
  { name: 'ขายของ', kws: ['ขาย','ยอดขาย','ออเดอร์','commission','ค่าคอม'] },
  { name: 'การเงิน/ลงทุน', kws: ['ดอกเบี้ย','ปันผล','หุ้น','คริปโต'] }
];

// Utils
function normalize(s) { return String(s || '').toLowerCase(); }

// ดึงจำนวนเงิน (รองรับ 1,234,567.89 และ 25000)
function extractAmount(text) {
  if (!text) return null;
  const cleaned = String(text).replace(/\s+/g, '');
  const m = cleaned.match(/(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?/);
  if (!m) return null;
  const raw = m[0].replace(/,/g, '');
  const n = Number(raw);
  return Number.isFinite(n) ? n : null;
}

function detectType(text) {
  const t = normalize(text);
  const isIncome = KW.income.some(k => t.includes(k.toLowerCase()));
  const isExpense = KW.expense.some(k => t.includes(k.toLowerCase()));
  if (isIncome && !isExpense) return 'รายรับ';
  if (isExpense && !isIncome) return 'รายจ่าย';
  if (t.includes('เงินเดือน') || t.includes('ได้') || t.includes('รับ')) return 'รายรับ';
  return 'รายจ่าย';
}

function detectCategory(text, type) {
  const t = normalize(text);
  const list = type === 'รายรับ' ? CATS_INCOME : CATS_EXPENSE;
  for (const cat of list) {
    if (cat.kws.some(k => t.includes(k.toLowerCase()))) return cat.name;
  }
  return 'อื่นๆ';
}

function stripNote(text) {
  return text.replace(/\s*(รายรับ|รายจ่าย)\b/gi, '').trim();
}

// ============================ DASHBOARD HELPERS ============================
// แปลงวันที่ไทย → Date JS
function parseThaiDate(s) {
  if (!s) return null;
  const parts = String(s).split('/');
  if (parts.length !== 3) return new Date(s);
  let [d, m, y] = parts.map(x => parseInt(x, 10));
  if (y > 2400) y -= 543;
  return new Date(y, m - 1, d);
}

// อยู่ในช่วงวันหรือไม่
function isWithin(d, start, end) {
  if (!d) return false;
  const dd = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const ss = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const ee = new Date(end.getFullYear(), end.getMonth(), end.getDate());
  return dd >= ss && dd <= ee;
}

// โค้ดช่วงเวลา → ช่วงวันที่
function getRangeDates(rangeCode) {
  const end = new Date();
  const start = new Date(end);
  switch (rangeCode) {
    case '1y': start.setFullYear(end.getFullYear() - 1); break;
    case '6m': start.setMonth(end.getMonth() - 6); break;
    case '3m': start.setMonth(end.getMonth() - 3); break;
    case '1m': start.setMonth(end.getMonth() - 1); break;
    case '1w': start.setDate(end.getDate() - 7); break;
    default:   start.setMonth(end.getMonth() - 1);
  }
  return { start, end };
}

// สร้าง URL กราฟ QuickChart
function buildQuickChartUrl(config, { w = 800, h = 400 } = {}) {
  const base = 'https://quickchart.io/chart';
  return `${base}?c=${encodeURIComponent(JSON.stringify(config))}&w=${w}&h=${h}`;
}

// สร้างภาพสรุปแดชบอร์ดจากข้อมูลของผู้ใช้
async function buildDashboardImages(userId, rangeCode = '1m') {
  const rows = await readRecentRowsForUser(userId, 1000);
  if (!rows || rows.length === 0) return { note: 'ยังไม่มีข้อมูลเลยครับ ลองบันทึกก่อนนะ' };

  const { start, end } = getRangeDates(rangeCode);
  const inRange = rows.filter(r => isWithin(parseThaiDate(r[0]), start, end));
  if (inRange.length === 0) return { note: 'ช่วงเวลานี้ยังไม่มีรายการครับ' };

  let sumIncome = 0, sumExpense = 0;
  const catExpense = {};
  for (const r of inRange) {
    const type = r[1];
    const amt = Number(String(r[2] ?? '0').toString().replace(/,/g, '')) || 0;
    const cat = r[3] || 'อื่นๆ';
    if (type === 'รายรับ') sumIncome += amt;
    else if (type === 'รายจ่าย') { sumExpense += amt; catExpense[cat] = (catExpense[cat] || 0) + amt; }
  }
  const balance = sumIncome - sumExpense;

  const barConfig = {
    type: 'bar',
    data: { labels: ['รวม'], datasets: [{ label: 'รายรับ', data: [sumIncome] }, { label: 'รายจ่าย', data: [sumExpense] }] },
    options: { plugins: { legend: { position: 'bottom' } } }
  };
  const pieLabels = Object.keys(catExpense);
  const pieData = Object.values(catExpense);
  const pieConfig = {
    type: 'pie',
    data: { labels: pieLabels.length ? pieLabels : ['ไม่มีรายจ่าย'], datasets: [{ data: pieData.length ? pieData : [1] }] },
    options: { plugins: { legend: { position: 'bottom' } } }
  };

  const barUrl = buildQuickChartUrl(barConfig);
  const pieUrl = buildQuickChartUrl(pieConfig, { w: 600, h: 600 });
  const pretty = (n) => Number(n).toLocaleString();
  const note = [
    `ช่วง: ${start.toLocaleDateString('th-TH')} – ${end.toLocaleDateString('th-TH')}`,
    `รายรับรวม: ${pretty(sumIncome)} บาท`,
    `รายจ่ายรวม: ${pretty(sumExpense)} บาท`,
    `คงเหลือ: ${pretty(balance)} บาท`
  ].join('\n');

  return { note, barUrl, pieUrl };
}
// ============================ END DASHBOARD HELPERS ============================

// -------- Gemini --------
const ENV_MODEL = (process.env.GEMINI_MODEL || '').trim();
const DEFAULT_MODELS = [
  'models/gemini-2.0-flash-lite-001',
  'models/gemini-2.5-flash',
  'models/gemini-2.0-flash',
  'models/gemini-2.5-pro'
];
const MODEL_LIST = ENV_MODEL ? [ENV_MODEL, ...DEFAULT_MODELS] : DEFAULT_MODELS;
const AI_TIMEOUT_MS = Number(process.env.AI_TIMEOUT_MS || 45000);

async function callGeminiModel({ model, apiKey, prompt }) {
  const url = `https://generativelanguage.googleapis.com/v1/${model}:generateContent`;
  const body = { contents: [{ parts: [{ text: prompt }]}] };
  let lastErr = null;

  for (let i = 0; i < 3; i++) {
    try {
      const res = await axios.post(url, body, {
        headers: { 'Content-Type': 'application/json', 'x-goog-api-key': apiKey },
        timeout: AI_TIMEOUT_MS
      });
      const c = res.data?.candidates?.[0];
      const text = c?.content?.parts?.[0]?.text;
      if (text) return text.trim();
      lastErr = new Error('empty response');
    } catch (err) {
      lastErr = err;
      const status = err?.response?.status;
      if ([429,500,502,503,504].includes(status)) {
        await new Promise(r => setTimeout(r, Math.pow(2, i) * 500));
        continue;
      } else {
        throw err;
      }
    }
  }
  throw lastErr;
}

async function analyzeWithGemini() {
  const rows = await readRecentRows(120);
  if (rows.length === 0) return 'ยังไม่มีข้อมูล ลองพิมพ์ "รายจ่าย 120 คาเฟ่" ก่อนนะ';
  const lines = rows.map(r => {
    const [date, type, amount, category] = r;
    return `${date} | ${type} | ${amount} | ${category || '-'}`;
  }).join('\n');

  const prompt = `
ข้อมูลรายรับ-รายจ่ายล่าสุด:
${lines}

โจทย์:
1) สรุปเดือนล่าสุด: ใช้จ่ายหมวดไหนเยอะสุดและประมาณเท่าไหร่
2) แนะนำแบบทำได้จริง 3 ข้อ
3) ถ้าต้องการ DCA เดือนละ 3,000 บาท ควรตัดจากหมวดใดจึงกระทบน้อยที่สุด
ย่อ กระชับ เป็น bullet และใส่ตัวเลขประมาณการ
`.trim();

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) return 'ยังไม่ได้ตั้งค่า GEMINI_API_KEY ใน .env นะ';

  for (const model of MODEL_LIST) {
    try {
      const text = await callGeminiModel({ model, apiKey, prompt });
      return text;
    } catch (e) {
      const status = e?.response?.status;
      if ([404,429,500,502,503,504].includes(status)) continue;
      return `เรียก AI ไม่ได้ (${model}): ${e?.response?.data?.error?.message || e.message}`;
    }
  }
  return 'Gemini ช้า/ล่มชั่วคราว ลองพิมพ์ "วิเคราะห์" ใหม่ครับ';
}

// วิเคราะห์เฉพาะของผู้ใช้ (เวอร์ชัน per-user)
// วิเคราะห์เฉพาะของผู้ใช้ (แฟกต์เบส + กติกา)
async function analyzeWithGeminiForUser(userId) {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) return 'ยังไม่ได้ตั้งค่า GEMINI_API_KEY ใน .env นะ';

  // รวบรวม facts 3 เดือนล่าสุด (ปรับได้)
  const { facts, text } = await buildFinanceFactsForUser(userId, 3);
  if (!facts) return 'ยังไม่มีข้อมูลของคุณเลยครับ ลองบันทึกก่อนนะ';

  // บล็อกป้องกัน “แนะนำลดยอดเล็กๆ”
  const rules = facts.rules;
  const safeCats = rules.ESSENTIAL_CATEGORIES;

  const prompt = `
คุณเป็นผู้จัดการการเงินส่วนบุคคล (Personal Finance Manager) หน้าที่คุณคือสรุปจาก "ข้อเท็จจริง" ที่ให้เท่านั้น
ห้ามเดา/ห้ามกุข้อมูลใหม่ และห้ามแนะนำให้ลดรายจ่ายเล็กน้อยที่ต่ำกว่า ${rules.MIN_SINGLE_IGNORE} บาท เว้นแต่รวมหมวดนั้น > ${rules.MIN_MONTHLY_CATEGORY_SUM} บาท/เดือน
ให้ความสำคัญกับหมวดที่กินสัดส่วนรายจ่าย > ${(rules.CATEGORY_SHARE_ALERT*100).toFixed(0)}% หรือโผล่สูงกว่า median/เฉลี่ยมาก

= ข้อเท็จจริง=
- รวมรายรับล่าสุด (ช่วง ${facts.monthsAnalyzed} เดือน): ${facts.totalIncome.toLocaleString()} บาท
- รวมรายจ่ายล่าสุด: ${facts.totalExpense.toLocaleString()} บาท
- คงเหลือ: ${(facts.balance).toLocaleString()} บาท
- ค่ากลางรายจ่ายเดี่ยว (median): ${Math.round(facts.medianSingleExpense).toLocaleString()} บาท
- ค่ากลางเฉลี่ย (mean): ${Math.round(facts.meanSingleExpense).toLocaleString()} บาท
- หมวดที่น่าพิจารณา (ตามเกณฑ์): ${facts.interestingCats.map(c => `${c.category} ${Math.round(c.sum).toLocaleString()}บ (${Math.round(c.share*100)}%)`).join(', ') || '—'}
- รายการเดี่ยวที่ดูสูงผิดปกติ (top): ${facts.outliers.map(o => `${o.date}:${o.category} ${Math.round(o.amount).toLocaleString()}บ`).join(', ') || '—'}
- เทียบเดือนล่าสุดกับก่อนหน้า: ${
    facts.monthOverMonth
      ? `รายรับ ${facts.monthOverMonth.last.key} ${facts.monthOverMonth.last.income.toLocaleString()}บ (Δ ${facts.monthOverMonth.diff.income>=0?'+':''}${facts.monthOverMonth.diff.income.toLocaleString()}บ), ` +
        `รายจ่าย ${facts.monthOverMonth.last.key} ${facts.monthOverMonth.last.expense.toLocaleString()}บ (Δ ${facts.monthOverMonth.diff.expense>=0?'+':''}${facts.monthOverMonth.diff.expense.toLocaleString()}บ)`
      : 'ข้อมูลไม่พอ'
  }

= ตารางดิบ (เพื่ออ้างอิงเท่านั้น ไม่ต้องคัดลอกทั้งหมดในการตอบ)=
วันที่ | ประเภท | จำนวน | หมวด
${text}

= กติกาการให้คำแนะนำ =
1) ข้ามรายการเดี่ยวที่ต่ำกว่า ${rules.MIN_SINGLE_IGNORE} บาท เว้นแต่หมวดนั้นรวมเดือนเกิน ${rules.MIN_MONTHLY_CATEGORY_SUM} บาท
2) พุ่งเป้าที่หมวดที่ share > ${(rules.CATEGORY_SHARE_ALERT*100).toFixed(0)}% หรือรวมต่อเดือนสูง
3) สำหรับหมวดจำเป็น (${safeCats.join(', ')}) ให้เสนอ "บริหาร/ต่อรอง/จัดตาราง" แทนการตัดจนเกินจริง
4) อย่าเสนอให้ลดสิ่งที่ไม่ส่งผลต่อภาพรวม
5) ถ้า “ไม่มีอะไรน่าลด” ให้พูดตามจริงว่า "ปกติดี"
6) ตอบเป็นภาษาไทย ลำดับหัวข้อชัดเจน อ่านง่าย (bullet) และใส่ตัวเลขประมาณให้เสมอ

= งานที่ต้องตอบ =
- สรุปภาพรวม 1–2 บรรทัด
- หมวดที่ควรให้ความสนใจ (เหตุผลสั้น + ตัวเลข)
- ข้อเสนอ 3 ข้อที่ทำได้จริง (เช่น เป้า/ความถี่/ขั้นตอน)
- สรุปสถานะ: ปกติ / ควรเฝ้าดู / ควรปรับ
`.trim();

  for (const model of MODEL_LIST) {
    try {
      const text = await callGeminiModel({ model, apiKey, prompt });
      return text;
    } catch (e) {
      const status = e?.response?.status;
      if ([404,429,500,502,503,504].includes(status)) continue;
      return `เรียก AI ไม่ได้ (${model}): ${e?.response?.data?.error?.message || e.message}`;
    }
  }
  return 'Gemini ช้า/ล่มชั่วคราว ลองใหม่อีกครั้งครับ';
}


// 🎨 Flex UI
function chip(text, bg = THEME.accentSoft, color = THEME.accent) {
  return {
    type: 'box', layout: 'baseline', backgroundColor: bg, cornerRadius: '12px', paddingAll: '6px',
    contents: [{ type: 'text', text, size: '12px', weight: 'bold', color }]
  };
}

// ===== FINANCE RULES (เกณฑ์เพื่อกันแนะนำมั่ว) =====
const FINANCE_RULES = {
  MIN_SINGLE_IGNORE: 100,            // ข้ามรายการเดี่ยวที่ต่ำกว่า X บาท (เว้นแต่มียอดรวมหมวดนี้สูง)
  MIN_MONTHLY_CATEGORY_SUM: 1000,    // หมวดไหนรวมต่อเดือนต่ำกว่า X ให้ถือว่า "ปกติ"
  CATEGORY_SHARE_ALERT: 0.15,        // ถ้าหมวดกินสัดส่วน > 15% ของรายจ่ายรวม → น่าสนใจ
  OUTLIER_MULTIPLIER: 1.3,           // มากกว่า median/เฉลี่ย 1.3 เท่า → ถือว่าโผล่
  ESSENTIAL_CATEGORIES: ['บิล/สาธารณูปโภค', 'ค่าเช่า', 'ที่อยู่อาศัย', 'การเดินทางจำเป็น'],
};


function confirmFlex({ type, amount, category, note, date, payload }) {
  const isIncome = type === 'รายรับ';
  const title = isIncome ? 'บันทึกรายรับ นี้ใช่ไหมครับ?' : 'บันทึกรายจ่าย นี้ใช่ไหมครับ ?';
  const icon = isIncome ? '💸' : '🧾';
  const amountTxt = `${amount.toLocaleString()} บาท`;
  const chips = [ chip(type), chip(category) ];

  return {
    type: 'flex',
    altText: title,
    contents: {
      type: 'bubble', size: 'mega', styles: { body: { backgroundColor: '#FFFFFF' } },
      header: {
        type: 'box', layout: 'horizontal', contents: [
          { type: 'text', text: icon + ' ' + title, weight: 'bold', size: 'md', color: THEME.textStrong },
        ], justifyContent: 'space-between'
      },
      body: {
        type: 'box', layout: 'vertical', spacing: 'md', contents: [
          { type: 'text', text: amountTxt, size: 'xxl', weight: 'bold', color: THEME.textStrong },
          { type: 'text', text: note || '-', size: 'sm', color: THEME.textMuted },
          { type: 'text', text: date, size: '12px', color: THEME.textMuted },
          { type: 'box', layout: 'horizontal', spacing: 'sm', contents: chips }
        ]
      },
      footer: {
        type: 'box', layout: 'horizontal', spacing: 'md', contents: [
          { type: 'button', style: 'primary', color: THEME.accent, height: 'sm',
            action: { type: 'postback', label: 'บันทึก', data: payload, displayText: 'บันทึก' } },
          { type: 'button', style: 'secondary', color: THEME.accentSoft, height: 'sm',
            action: { type: 'postback', label: 'ยกเลิก', data: 'action=cancel', displayText: 'ยกเลิก' } }
        ]
      }
    }
  };
}

function buildDashboardMenuFlex() {
  const GREEN = '#16A34A';
  const TEXT = '#111111';
  const MUTED = '#8B95A1';
  return {
    type: 'flex',
    altText: 'แดชบอร์ดการเงิน',
    contents: {
      type: 'bubble',
      size: 'kilo',
      header: {
        type: 'box',
        layout: 'vertical',
        paddingAll: '16px',
        contents: [
          { type: 'text', text: 'แดชบอร์ดการเงิน', weight: 'bold', size: 'lg', color: TEXT },
          { type: 'text', text: 'เลือกช่วงเวลาที่ต้องการดูสรุป', size: 'xs', color: MUTED, margin: 'sm' }
        ]
      },
      body: {
        type: 'box',
        layout: 'vertical',
        spacing: '12px',
        paddingAll: '16px',
        contents: [
          { type: 'box', layout: 'horizontal', spacing: '8px', contents: [
            { type: 'button', style: 'secondary', height: 'sm', action: { type: 'postback', label: '1 year',  data: 'action=dash&range=1y' } },
            { type: 'button', style: 'secondary', height: 'sm', action: { type: 'postback', label: '6 เดือน', data: 'action=dash&range=6m' } }
          ]},
          { type: 'box', layout: 'horizontal', spacing: '8px', contents: [
            { type: 'button', style: 'secondary', height: 'sm', action: { type: 'postback', label: '3 เดือน', data: 'action=dash&range=3m' } },
            { type: 'button', style: 'secondary', height: 'sm', action: { type: 'postback', label: '1 เดือน', data: 'action=dash&range=1m' } }
          ]},
          { type: 'box', layout: 'horizontal', contents: [
            { type: 'button', style: 'secondary', height: 'sm', action: { type: 'postback', label: '1 week', data: 'action=dash&range=1w' } }
          ]}
        ]
      },
      footer: {
        type: 'box',
        layout: 'vertical',
        paddingAll: '16px',
        contents: [
          { type: 'button', style: 'primary', height: 'sm', color: GREEN,
            action: { type: 'postback', label: 'Export เป็น Excel', data: 'action=dash&do=export_excel' } }
        ]
      },
      styles: { header: { backgroundColor: '#FFFFFF' }, body: { backgroundColor: '#FFFFFF' }, footer: { backgroundColor: '#FFFFFF' } }
    }
  };
}

// -------- Payload Builder --------
function buildSavePayload({ type, amount, category, note }) {
  const amt = Number(String(amount).replace(/,/g, ''));
  const shortNote = String(note || '').slice(0, 40);
  const params = new URLSearchParams({ action: 'save', type, amount: String(amt), category, note: shortNote });
  return params.toString();
}

// -------- Routes --------
app.get('/webhook', (req, res) => {
  res.status(200).send('OK');
});

app.post('/webhook', middleware(lineConfig), (req, res) => {
  try {
    res.sendStatus(200); // ตอบ 200 ทันที กัน timeout
    const body = (req && req.body && typeof req.body === 'object') ? req.body : {};
    const events = Array.isArray(body.events) ? body.events : [];
    Promise.all(events.map(e => handleEvent(e))).catch(err => {
      console.error('handleEvent error:', err?.response?.data || err.message || err);
    });
  } catch (err) {
    console.error('webhook handler crash:', err);
  }
});

app.get('/', (req, res) => { res.send('Bot พร้อมทำงาน'); });

app.get('/debug/models', async (req, res) => {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) return res.status(400).json({ error: 'GEMINI_API_KEY not set' });

  const url = 'https://generativelanguage.googleapis.com/v1/models';
  let lastErr = null;
  for (let i = 0; i < 3; i++) {
    try {
      const r = await axios.get(url, { headers: { 'x-goog-api-key': apiKey }, timeout: Number(process.env.AI_TIMEOUT_MS || 45000) });
      const models = (r.data?.models || []).map(m => ({
        name: m.name, displayName: m.displayName, supportedGenerationMethods: m.supportedGenerationMethods
      }));
      return res.json({ models });
    } catch (e) {
      lastErr = e;
      const status = e?.response?.status;
      if ([502,503,504].includes(status)) {
        await new Promise(r => setTimeout(r, Math.pow(2, i) * 500));
        continue;
      }
      return res.status(500).json({ error: e?.response?.data?.error?.message || e.message });
    }
  }
  return res.status(503).json({ error: 'The service is currently unavailable (after retries).' });
});

// -------- LINE Handler --------
async function handleEvent(event) {
  // Postback ก่อน
  if (event.type === 'postback' && event.postback?.data) {
    const data = String(event.postback.data || '');
    const p = new URLSearchParams(data);
    const action = p.get('action');

    if (action === 'cancel' && event.replyToken) {
      return lineClient.replyMessage(event.replyToken, { type: 'text', text: 'ยกเลิกแล้วครับ' });
    }

    if (action === 'save') {
      const type = p.get('type') || 'รายจ่าย';
      const amount = Number(String(p.get('amount') || '0').replace(/,/g, ''));
      const category = p.get('category') || 'อื่นๆ';
      const note = p.get('note') || '-';
      const date = todayTH();
      const uid = event.source?.userId || '-';
      try {
        await appendRowToUser(uid, [date, type, amount, category, note, uid]);
        if (event.replyToken) {
          return lineClient.replyMessage(event.replyToken, {
            type: 'text',
            text: `บันทึกแล้ว: ${type} ${amount.toLocaleString()} บ. • ${category} • ${date} ✅`
          });
        }
      } catch (err) {
        if (event.replyToken) {
          return lineClient.replyMessage(event.replyToken, {
            type: 'text',
            text: 'บันทึกไม่ผ่าน เช็กสิทธิ์ชีต/ค่า .env ก่อนครับ'
          });
        }
      }
      return;
    }

    // ===== Dash range & export =====
    if (action === 'dash') {
      const range = p.get('range');   // '1y' | '6m' | '3m' | '1m' | '1w'
      const doing = p.get('do');      // 'export_excel' หรือ undefined

      if (doing === 'export_excel') {
        return lineClient.replyMessage(event.replyToken, {
          type: 'text',
          text: 'กำลังสร้างไฟล์ Excel ให้ครับ... (จะส่งลิงก์ดาวน์โหลดให้เร็ว ๆ นี้)'
        });
      }

      const userId = event.source?.userId || 'anonymous';
      const dash = await buildDashboardImages(userId, range || '1m');

      const msgs = [];
      if (dash.note)   msgs.push({ type: 'text',  text: dash.note });
      if (dash.barUrl) msgs.push({ type: 'image', originalContentUrl: dash.barUrl, previewImageUrl: dash.barUrl });
      if (dash.pieUrl) msgs.push({ type: 'image', originalContentUrl: dash.pieUrl, previewImageUrl: dash.pieUrl });

      if (msgs.length === 0) msgs.push({ type: 'text', text: 'ไม่มีข้อมูลในช่วงเวลาที่เลือกครับ' });
      return lineClient.replyMessage(event.replyToken, msgs);
    }

    return; // postback แต่ไม่เข้ากรณีไหน
  }

  // ข้อความธรรมดา
  if (event.type !== 'message' || event.message.type !== 'text') return;
  const text = event.message.text.trim();
  const today = todayTH();

  // เมนูแดชบอร์ด
  if (/^(แดชบอร์ด|dashboard|สรุป|รายงาน)$/i.test(text)) {
    const flex = buildDashboardMenuFlex();
    return lineClient.replyMessage(event.replyToken, flex);
  }

  // ฟอร์แมตชัดเจน (รองรับคอมม่า/ทศนิยม)
  const spendRegex  = /^(รายจ่าย)\s+((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?)\s+(.+)$/i;
  const incomeRegex = /^(รายรับ)\s+((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?)\s+(.+)$/i;

  if (/^(รายจ่าย|รายรับ)\b/i.test(text) && !(spendRegex.test(text) || incomeRegex.test(text))) {
    return lineClient.replyMessage(event.replyToken, {
      type: 'text',
      text: 'รูปแบบผิด ลองแก้เป็น "รายจ่าย 120 คาเฟ่" หรือ "รายรับ 15000 เงินเดือน" ครับ'
    });
  }

  if (spendRegex.test(text) || incomeRegex.test(text)) {
    const isSpend = spendRegex.test(text);
    const [, type, amountRaw, detail] = (isSpend ? spendRegex : incomeRegex).exec(text);
    const amount = Number(String(amountRaw).replace(/,/g, ''));
    const category = detectCategory(detail, type);
    const note = detail;
    const payload = buildSavePayload({ type, amount, category, note });
    const flex = confirmFlex({ type, amount, category, note, date: today, payload });
    return lineClient.replyMessage(event.replyToken, flex);
  }

  // โหมดเดาอัตโนมัติ
  const amt = extractAmount(text);
  if (amt && amt > 0) {
    const type = detectType(text);
    const category = detectCategory(text, type);
    const note = stripNote(text).trim();
    const payload = buildSavePayload({ type, amount: amt, category, note });
    const flex = confirmFlex({ type, amount: amt, category, note, date: today, payload });
    return lineClient.replyMessage(event.replyToken, flex);
  }

  // วิเคราะห์ (เฉพาะข้อมูลของผู้ใช้)
  if (text === 'วิเคราะห์') {
    const uid = event.source?.userId;
    if (uid) {
      await lineClient.replyMessage(event.replyToken, { type: 'text', text: 'กำลังวิเคราะห์ข้อมูลของคุณ แป๊บเดียวครับ...' });
      analyzeWithGeminiForUser(uid)
        .then(msg => lineClient.pushMessage(uid, { type: 'text', text: msg }))
        .catch(err => lineClient.pushMessage(uid, {
          type: 'text',
          text: `เรียก AI ไม่ได้: ${err?.response?.data?.error?.message || err.message || 'unknown'}`
        }));
      return;
    } else {
      const advice = await analyzeWithGemini();
      return lineClient.replyMessage(event.replyToken, { type: 'text', text: advice });
    }
  }

  // เมนูช่วยเหลือ
  const help = [
    'สวัสดีครับ! บันทึกรายรับ รายจ่ายพิมพ์ได้เลยครับ',
    'ตัวอย่างการบันทึก',
    'กาแฟ 120 ',
    'เงินเดือน 15000'
  ].join('\n');
  return lineClient.replyMessage(event.replyToken, { type: 'text', text: help });
}

// -------- Start Server --------
const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
