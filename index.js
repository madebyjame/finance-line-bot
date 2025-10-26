/**
 * LINE Bot สำหรับบันทึกรายรับ-รายจ่าย อัตโนมัติลงใน Google Sheets
 * Features:
 * - บันทึกรายรับ/รายจ่ายจากข้อความแชท (เช่น "กาแฟ 60", "เงินเดือน 30000")
 * - วิเคราะห์ค่าใช้จ่ายด้วย Gemini AI
 * - แสดงแดชบอร์ดสรุปและกราฟวิเคราะห์
 * - ส่งออกข้อมูลเป็น Excel ได้
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
const express = require('express');          // เว็บเซิร์ฟเวอร์
const { Client, middleware } = require('@line/bot-sdk');  // LINE Bot SDK
const { google } = require('googleapis');    // Google Sheets API
const axios = require('axios');              // HTTP client สำหรับเรียก APIs

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
// 🔍 คำสำคัญสำหรับระบบเดาอัตโนมัติว่าเป็นรายรับหรือรายจ่าย
// ใช้ตอนผู้ใช้พิมพ์แบบสั้น เช่น "กาแฟ 60" (รายจ่าย) หรือ "ได้เงิน 500" (รายรับ)
const KW = {
  income: [ 'รายรับ','รับ','ได้','โอนเข้า','เงินเดือน','โบนัส','ทิป','ขายได้','ดอกเบี้ย','ปันผล' ],
  expense: [ 'รายจ่าย','จ่าย','ซื้อ','โอนออก','เติม','ค่าผ่อน','ผ่อน','ค่าสมาชิก','ค่าใช้จ่าย' ],
};

// 📂 หมวดหมู่รายจ่าย พร้อมคำสำคัญสำหรับระบบเดาอัตโนมัติ
// เมื่อผู้ใช้พิมพ์ "กาแฟ 60" ระบบจะเดาว่าเป็นหมวด "อาหาร/กาแฟ" เป็นต้น
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

// 🛠️ Utility functions สำหรับประมวลผลข้อความ

// แปลงข้อความเป็นตัวพิมพ์เล็ก และจัดการกรณี null/undefined
function normalize(s) { return String(s || '').toLowerCase(); }

// แยกจำนวนเงินออกจากข้อความ รองรับรูปแบบต่างๆ เช่น
// - "1234" -> 1234
// - "1,234.56" -> 1234.56
// - "1 234.5" -> 1234.50
function extractAmount(text) {
  const m = String(text).match(/(\d{1,3}(?:[, ]\d{3})*|\d+)(?:[.,](\d{1,2}))?/);
  if (!m) return null;
  const num = (m[1] || '').replace(/[ ,]/g, '');
  const dec = m[2] ? '.' + m[2] : '';
  const val = Number(num + dec);
  return Number.isFinite(val) ? val : null;
}

// 🤖 ระบบเดาอัตโนมัติว่าเป็นรายรับหรือรายจ่าย
// ใช้คำสำคัญจาก KW.income และ KW.expense ในการเดา
// ถ้าไม่ชัดเจน จะเดาว่าเป็นรายจ่าย (default)
function detectType(text) {
  const t = normalize(text);
  const isIncome = KW.income.some(k => t.includes(k.toLowerCase()));
  const isExpense = KW.expense.some(k => t.includes(k.toLowerCase()));
  if (isIncome && !isExpense) return 'รายรับ';
  if (isExpense && !isIncome) return 'รายจ่าย';
  if (t.includes('เงินเดือน') || t.includes('ได้') || t.includes('รับ')) return 'รายรับ';
  return 'รายจ่าย';
}

// 🏷️ ระบบเดาหมวดหมู่อัตโนมัติจากคำในข้อความ
// ใช้คำสำคัญจาก CATS_INCOME และ CATS_EXPENSE
// ถ้าไม่ตรงกับหมวดไหนเลย จะเป็น "อื่นๆ"
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

// 🎨 ส่วนสร้าง UI ด้วย Flex Message
// สร้างป้ายชิป (chip) แสดงประเภทและหมวดหมู่
// เช่น [รายจ่าย] [อาหาร/กาแฟ]
function chip(text, bg = THEME.accentSoft, color = THEME.accent) {
  return {
    type: 'box', layout: 'baseline', backgroundColor: bg, cornerRadius: '12px', paddingAll: '6px',
    contents: [{ type: 'text', text, size: '12px', weight: 'bold', color }]
  };
}

// 🎯 สร้างหน้ายืนยันการบันทึกด้วย Flex Message
// แสดงรายละเอียดการบันทึก พร้อมปุ่มยืนยัน/ยกเลิก
// UI จะปรับเปลี่ยนตามประเภท (รายรับ/รายจ่าย)
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

// 🔄 สร้าง payload สำหรับส่งข้อมูลกลับมาเมื่อกดปุ่มบันทึก
// แปลงข้อมูลเป็น URL encoded string เพื่อส่งผ่าน LINE postback
// จำกัดความยาว note ไม่เกิน 40 ตัวอักษร
function buildSavePayload({ type, amount, category, note }) {
  const shortNote = String(note || '').slice(0, 40);
  const params = new URLSearchParams({ action: 'save', type, amount: String(amount), category, note: shortNote });
  return params.toString();
}

// -------- Routes --------
app.get('/webhook', (req, res) => {
  res.status(200).send('OK');
});

app.post('/webhook', middleware(lineConfig), (req, res) => {
  try {
    res.sendStatus(200);
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
      const amount = Number(p.get('amount') || '0');
      const category = p.get('category') || 'อื่นๆ';
      const note = p.get('note') || '-';
      const date = todayTH();
      try {
        await appendRow([date, type, amount, category, note]);
        if (event.replyToken) {
          return lineClient.replyMessage(event.replyToken, { type: 'text', text: `บันทึกแล้ว: ${type} ${amount.toLocaleString()} บ. • ${category} • ${date} ✅` });
        }
      } catch (err) {
        if (event.replyToken) {
          return lineClient.replyMessage(event.replyToken, { type: 'text', text: 'บันทึกไม่ผ่าน เช็กสิทธิ์ชีต/ค่า .env ก่อนครับ' });
        }
      }
      return;
    }
  }

  // ข้อความธรรมดา
  if (event.type !== 'message' || event.message.type !== 'text') return;
  const text = event.message.text.trim();
  const today = todayTH();

  // คำสั่งเก่าแบบกำหนดชัดเจน
  const spendRegex = /^(รายจ่าย)\s+(\d+(?:[.,]\d+)?)\s+(.+)$/i;
  const incomeRegex = /^(รายรับ)\s+(\d+(?:[.,]\d+)?)\s+(.+)$/i;

  if (/^(รายจ่าย|รายรับ)\b/i.test(text) && !(spendRegex.test(text) || incomeRegex.test(text))) {
    return lineClient.replyMessage(event.replyToken, { type: 'text', text: 'รูปแบบไม่ครบ ลองแบบนี้ → "รายจ่าย 120 คาเฟ่" หรือ "รายรับ 15000 เงินเดือน"' });
  }

  if (spendRegex.test(text) || incomeRegex.test(text)) {
    const isSpend = spendRegex.test(text);
    const [, type, amountRaw, detail] = (isSpend ? spendRegex : incomeRegex).exec(text);
    const amount = Number(String(amountRaw).replace(',', ''));
    const category = detectCategory(detail, type);
    const note = detail;
    const payload = buildSavePayload({ type, amount, category, note });
    const flex = confirmFlex({ type, amount, category, note, date: today, payload });
    return lineClient.replyMessage(event.replyToken, flex);
  }

  // โหมดอัตโนมัติ: แค่พิมพ์ "กาแฟ 65" หรือ "เงินเดือน 15000"
  const amt = extractAmount(text);
  if (amt && amt > 0) {
    const type = detectType(text);
    const category = detectCategory(text, type);
    const note = stripNote(text).trim();
    const payload = buildSavePayload({ type, amount: amt, category, note });
    const flex = confirmFlex({ type, amount: amt, category, note, date: today, payload });
    return lineClient.replyMessage(event.replyToken, flex);
  }

  // วิเคราะห์ (เดิม)
  if (text === 'วิเคราะห์') {
    if (event.source?.userId) {
      await lineClient.replyMessage(event.replyToken, { type: 'text', text: 'กำลังวิเคราะห์ แป๊บเดียวครับ...' });
      analyzeWithGemini()
        .then(msg => lineClient.pushMessage(event.source.userId, { type: 'text', text: msg }))
        .catch(err => lineClient.pushMessage(event.source.userId, { type: 'text', text: `เรียก AI ไม่ได้: ${err?.response?.data?.error?.message || err.message || 'unknown'}` }));
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
    '• วิเคราะห์  → (สรุปด้วย AI)'
  ].join('\n');
  return lineClient.replyMessage(event.replyToken, { type: 'text', text: help });
}

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});