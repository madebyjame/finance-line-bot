/**
 * LINE Bot สำหรับบันทึกรายรับ-รายจ่าย อัตโนมัติลงใน Google Sheets
 * Features:
 * - บันทึกรายรับ/รายจ่ายจากข้อความแชท (เช่น "กาแฟ 60", "เงินเดือน 30000")
 * - วิเคราะห์ค่าใช้จ่ายด้วย Gemini AI
 * - แสดงแดชบอร์ดสรุปและกราฟวิเคราะห์
 * - ส่งออกข้อมูลเป็น Excel ได้ (placeholder)
 */

const fs = require("fs");
// สร้าง Google Service Account credentials เป็นไฟล์ชั่วคราว (สำหรับ Railway)
if (process.env.GOOGLE_CREDENTIALS_JSON) {
  const path = "/tmp/google.json";
  fs.writeFileSync(path, process.env.GOOGLE_CREDENTIALS_JSON);
  process.env.GOOGLE_SERVICE_ACCOUNT_FILE = path;
}

// โหลดค่า environment variables จากไฟล์ .env
require('dotenv').config();

// โหลด dependencies หลัก
const express = require('express');                     // เว็บเซิร์ฟเวอร์
const { Client, middleware } = require('@line/bot-sdk'); // LINE Bot SDK
const { google } = require('googleapis');               // Google Sheets API
const axios = require('axios');                         // HTTP client

const app = express();

const lineConfig = {
  channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.CHANNEL_SECRET
};
const lineClient = new Client(lineConfig);

// ===== THEME (Modern UI) =====
const THEME = {
  accent: '#4F46E5',       // indigo-600
  accentSoft: '#EEF2FF',   // indigo-50
  danger: '#EF4444',       // red-500
  textMuted: '#8B95A1',
  textStrong: '#111827',
};

// -------- Google Sheets --------
const auth = new google.auth.GoogleAuth({
  keyFile: process.env.GOOGLE_SERVICE_ACCOUNT_FILE,
  scopes: ['https://www.googleapis.com/auth/spreadsheets']
});
const sheets = google.sheets({ version: 'v4', auth });

const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const SHEET_RANGE = process.env.SHEET_RANGE || 'Sheet1!A:E';

function headRangeFrom(range) {
  const sheetTitle = (range.includes('!') ? range.split('!')[0] : 'Sheet1');
  return `${sheetTitle}!A1:E1`;
}

async function ensureHeader() {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: headRangeFrom(SHEET_RANGE)
  });
  const hasHeader = res.data.values && res.data.values.length > 0;
  if (!hasHeader) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: headRangeFrom(SHEET_RANGE),
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [['วันที่', 'ประเภท', 'จำนวน', 'หมวดหมู่', 'รายละเอียด']] }
    });
  }
}

async function appendRow(values) {
  await ensureHeader();
  return sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: SHEET_RANGE,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [values] }
  });
}

async function readRecentRows(limit = 120) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: SHEET_RANGE
  });
  const rows = res.data.values || [];
  const dataRows = rows.length > 1 ? rows.slice(1) : [];
  return dataRows.slice(-limit);
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

// 🛠️ Utilities
function normalize(s) { return String(s || '').toLowerCase(); }

// รองรับเลขใหญ่/คอมม่า/ทศนิยม เช่น 1,234,567.89 หรือ 25000
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
// แปลงวันที่รูปแบบไทย (เช่น 26/10/2568) เป็น Date JS (แปลง พ.ศ. → ค.ศ.)
function parseThaiDate(s) {
  if (!s) return null;
  const parts = String(s).split('/');
  if (parts.length !== 3) return new Date(s);
  let [d, m, y] = parts.map(x => parseInt(x, 10));
  if (y > 2400) y -= 543; // พ.ศ. → ค.ศ.
  return new Date(y, m - 1, d);
}

// ตรวจว่า date อยู่ในช่วง [start, end] หรือไม่ (ไม่สนเวลา)
function isWithin(d, start, end) {
  if (!d) return false;
  const dd = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const ss = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const ee = new Date(end.getFullYear(), end.getMonth(), end.getDate());
  return dd >= ss && dd <= ee;
}

// แปลงโค้ดช่วงเวลา (1y, 6m, 3m, 1m, 1w) → วันที่เริ่ม/สิ้นสุด
function getRangeDates(rangeCode) {
  const end = new Date(); // วันนี้
  const start = new Date(end);
  switch (rangeCode) {
    case '1y': start.setFullYear(end.getFullYear() - 1); break;
    case '6m': start.setMonth(end.getMonth() - 6); break;
    case '3m': start.setMonth(end.getMonth() - 3); break;
    case '1m': start.setMonth(end.getMonth() - 1); break;
    case '1w': start.setDate(end.getDate() - 7); break;
    default:   start.setMonth(end.getMonth() - 1); // ดีฟอลต์ 1 เดือน
  }
  return { start, end };
}

// สร้าง URL กราฟด้วย QuickChart
function buildQuickChartUrl(config, { w = 800, h = 400 } = {}) {
  const base = 'https://quickchart.io/chart';
  return `${base}?c=${encodeURIComponent(JSON.stringify(config))}&w=${w}&h=${h}`;
}

// อ่านข้อมูลล่าสุด (ยังไม่แยก user) — ใช้ข้อมูลรวมไปก่อน
async function readRecentRowsForUser(userId, limit = 1000) {
  return readRecentRows(limit);
}

// คำนวณสรุป และสร้างลิงก์กราฟ (แท่ง: รับ/จ่าย รวม, พาย: รายจ่ายตามหมวด)
async function buildDashboardImages(userId, rangeCode = '1m') {
  const rows = await readRecentRowsForUser(userId, 1000);
  if (!rows || rows.length === 0) return { note: 'ยังไม่มีข้อมูลเลยครับ ลองบันทึกก่อนนะ' };

  const { start, end } = getRangeDates(rangeCode);

  // กรองตามช่วงเวลา
  const inRange = rows.filter(r => {
    const d = parseThaiDate(r[0]); // คอลัมน์ A = วันที่ (th-TH)
    return isWithin(d, start, end);
  });

  if (inRange.length === 0) {
    return { note: 'ช่วงเวลานี้ยังไม่มีรายการครับ' };
  }

  // รวมยอด และแตกหมวดหมู่รายจ่าย
  let sumIncome = 0;
  let sumExpense = 0;
  const catExpense = {};
  for (const r of inRange) {
    const type = r[1];
    const amt = Number(String(r[2] ?? '0').toString().replace(/,/g, '')) || 0;
    const cat = r[3] || 'อื่นๆ';
    if (type === 'รายรับ') sumIncome += amt;
    else if (type === 'รายจ่าย') {
      sumExpense += amt;
      catExpense[cat] = (catExpense[cat] || 0) + amt;
    }
  }
  const balance = sumIncome - sumExpense;

  // กราฟแท่ง: รายรับ/รายจ่ายรวม
  const barConfig = {
    type: 'bar',
    data: {
      labels: ['รวม'],
      datasets: [
        { label: 'รายรับ', data: [sumIncome] },
        { label: 'รายจ่าย', data: [sumExpense] }
      ]
    },
    options: { plugins: { legend: { position: 'bottom' } } }
  };

  // กราฟพาย: รายจ่ายตามหมวด
  const pieLabels = Object.keys(catExpense);
  const pieData = Object.values(catExpense);
  const pieConfig = {
    type: 'pie',
    data: {
      labels: pieLabels.length ? pieLabels : ['ไม่มีรายจ่าย'],
      datasets: [{ data: pieData.length ? pieData : [1] }]
    },
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
      if (status === 429 || status === 500 || status === 502 || status === 503 || status === 504) {
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
  if (rows.length === 0) {
    return 'ยังไม่มีข้อมูล ลองพิมพ์ "รายจ่าย 120 คาเฟ่" ก่อนนะ';
  }

  const lines = rows.map(r => {
    const [date, type, amount, category] = r;
    return `${date} | ${type} | ${amount} | ${category || '-'}`;
  }).join('\n');

  const prompt = `
ข้อมูลรายรับ-รายจ่ายล่าสุด:
${lines}

โจทย์:
1) สรุปเดือนล่าสุด: ใช้จ่ายหมวดไหนเยอะสุดและประมาณเท่าไหร่
2) แนะนำแบบทำได้จริง 3 ข้อ (เช่น ลดหมวดไหนกี่ครั้ง/สัปดาห์ จะออมเพิ่มได้ประมาณเท่าไร/เดือน)
3) ถ้าต้องการ DCA เดือนละ 3,000 บาท ควรตัดจากหมวดใดจึงกระทบน้อยที่สุด
ย่อ กระชับ เป็น bullet และใส่ตัวเลขประมาณการให้ด้วย
`.trim();

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) return 'ยังไม่ได้ตั้งค่า GEMINI_API_KEY ใน .env นะ';

  for (const model of MODEL_LIST) {
    try {
      const text = await callGeminiModel({ model, apiKey, prompt });
      return text;
    } catch (e) {
      const status = e?.response?.status;
      if (status === 404) continue;
      if (status === 503 || status === 502 || status === 504 || status === 500 || status === 429) continue;
      return `เรียก AI ไม่ได้ (${model}): ${e?.response?.data?.error?.message || e.message}`;
    }
  }
  return 'Gemini ช้า/ล่มชั่วคราว ลองพิมพ์ "วิเคราะห์" ใหม่ หรือตั้ง GEMINI_MODEL เป็น models/gemini-2.0-flash-lite-001';
}

// 🎨 Flex UI
function chip(text, bg = THEME.accentSoft, color = THEME.accent) {
  return {
    type: 'box', layout: 'baseline', backgroundColor: bg, cornerRadius: '12px', paddingAll: '6px',
    contents: [{ type: 'text', text, size: '12px', weight: 'bold', color }]
  };
}

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
          { type: 'button', style: 'primary', color: THEME.accent, height: 'sm', action: { type: 'postback', label: 'บันทึก', data: payload, displayText: 'บันทึก' } },
          { type: 'button', style: 'secondary', color: THEME.accentSoft, height: 'sm', action: { type: 'postback', label: 'ยกเลิก', data: 'action=cancel', displayText: 'ยกเลิก' } }
        ]
      }
    }
  };
}

// เมนูแดชบอร์ด (ช่วงเวลา + Export)
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

app.get('/', (req, res) => {
  res.send('Bot พร้อมทำงาน');
});

app.get('/debug/models', async (req, res) => {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) return res.status(400).json({ error: 'GEMINI_API_KEY not set' });

  const url = 'https://generativelanguage.googleapis.com/v1/models';
  let lastErr = null;
  for (let i = 0; i < 3; i++) {
    try {
      const r = await axios.get(url, { headers: { 'x-goog-api-key': apiKey }, timeout: Number(process.env.AI_TIMEOUT_MS || 45000) });
      const models = (r.data?.models || []).map(m => ({
        name: m.name,
        displayName: m.displayName,
        supportedGenerationMethods: m.supportedGenerationMethods
      })); 
      return res.json({ models });
    } catch (e) {
      lastErr = e;
      const status = e?.response?.status;
      if (status === 503 || status === 502 || status === 504) {
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
      try {
        await appendRow([date, type, amount, category, note]);
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

    // ถ้าเป็น postback แต่ไม่เข้ากรณีไหนเลย
    return;
  }

  // ข้อความธรรมดา
  if (event.type !== 'message' || event.message.type !== 'text') return;
  const text = event.message.text.trim();
  const today = todayTH();

  // เมนูแดชบอร์ด
  if (/^แดชบอร์ด$/i.test(text)) {
    const flex = buildDashboardMenuFlex();
    return lineClient.replyMessage(event.replyToken, flex);
  }

  // คำสั่งแบบกำหนดชัดเจน (รองรับเลขมีคอมม่า/ทศนิยม)
  const spendRegex  = /^(รายจ่าย)\s+((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?)\s+(.+)$/i;
  const incomeRegex = /^(รายรับ)\s+((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?)\s+(.+)$/i;

  if (/^(รายจ่าย|รายรับ)\b/i.test(text) && !(spendRegex.test(text) || incomeRegex.test(text))) {
    return lineClient.replyMessage(event.replyToken, {
      type: 'text',
      text: 'รูปแบบไม่ครบ ลองแบบนี้ → "รายจ่าย 120 คาเฟ่" หรือ "รายรับ 15000 เงินเดือน"'
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

  // โหมดเดาอัตโนมัติ: พิมพ์ "กาแฟ 65" หรือ "เงินเดือน 15000"
  const amt = extractAmount(text);
  if (amt && amt > 0) {
    const type = detectType(text);
    const category = detectCategory(text, type);
    const note = stripNote(text).trim();
    const payload = buildSavePayload({ type, amount: amt, category, note });
    const flex = confirmFlex({ type, amount: amt, category, note, date: today, payload });
    return lineClient.replyMessage(event.replyToken, flex);
  }

  // วิเคราะห์ด้วย AI
  if (text === 'วิเคราะห์') {
    if (event.source?.userId) {
      await lineClient.replyMessage(event.replyToken, { type: 'text', text: 'กำลังวิเคราะห์ แป๊บเดียวครับ...' });
      analyzeWithGemini()
        .then(msg => lineClient.pushMessage(event.source.userId, { type: 'text', text: msg }))
        .catch(err => lineClient.pushMessage(event.source.userId, {
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
    'พิมพ์เร็ว ๆ ได้แบบนี้:',
    '• กาแฟ 65  → (ระบบเดา: รายจ่าย/อาหาร)',
    '• เงินเดือน 15000  → (ระบบเดา: รายรับ/เงินเดือน)',
    '• รายจ่าย 120 คาเฟ่  → (กำหนดเอง)',
    '• รายรับ 15000 เงินเดือน  → (กำหนดเอง)',
    '• แดชบอร์ด  → (ดูกราฟสรุป)',
    '• วิเคราะห์  → (สรุปด้วย AI)'
  ].join('\n');
  return lineClient.replyMessage(event.replyToken, { type: 'text', text: help });
}

// -------- Start Server --------
const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
