const fs = require("fs");
// ✅ ถ้ามีข้อมูล GOOGLE_CREDENTIALS_JSON จาก Railway จะเขียนเป็นไฟล์ชั่วคราว เพื่อให้ Google ใช้ตรวจสิทธิ์
if (process.env.GOOGLE_CREDENTIALS_JSON) {
const path = "/tmp/google.json"; // ที่อยู่ไฟล์ชั่วคราวในเซิร์ฟเวอร์
fs.writeFileSync(path, process.env.GOOGLE_CREDENTIALS_JSON);
process.env.GOOGLE_SERVICE_ACCOUNT_FILE = path; // ตั้งค่าให้ Google ใช้ไฟล์นี้
}


// ✅ โหลดค่าตัวแปรจากไฟล์ .env เช่น Channel Token, Sheet ID, API Key
require('dotenv').config();

// ✅ โหลดโมดูลที่ต้องใช้
const express = require('express'); // สร้างเว็บเซิร์ฟเวอร์
const { Client, middleware } = require('@line/bot-sdk'); // เชื่อมต่อกับ LINE Messaging API
const { google } = require('googleapis'); // เชื่อมต่อ Google Sheets
const axios = require('axios'); // เรียก API ภายนอก

// ✅ สร้างแอป Express
const app = express();

// ✅ ตั้งค่าการเชื่อมต่อกับ LINE จากค่าใน Railway (.env)
const lineConfig = {
  channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.CHANNEL_SECRET
};
const lineClient = new Client(lineConfig);

// 🎨 THEME: สีและสไตล์ UI สำหรับ Flex Message (โทนมินิมอล)
const THEME = {
accent: '#4F46E5', // สีหลัก (ม่วง)
accentSoft: '#EEF2FF', // สีพื้นอ่อน
danger: '#EF4444', // สีแดงแจ้งเตือน
textMuted: '#8B95A1', // สีข้อความจาง
textStrong: '#111827', // สีข้อความเข้ม
};

// 📊 เชื่อมกับ Google Sheets ด้วยบัญชี Service Account
const auth = new google.auth.GoogleAuth({
  keyFile: process.env.GOOGLE_SERVICE_ACCOUNT_FILE, //ใช้ไฟล์คีย์ที่สร้างไว้ข้างบน
  scopes: ['https://www.googleapis.com/auth/spreadsheets'] // สิทธิ์เข้าถึงชีต
});
const sheets = google.sheets({ version: 'v4', auth });

// 🗂️ กำหนด ID ของชีต และช่วงข้อมูล (เช่น A:E)
const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const SHEET_RANGE = process.env.SHEET_RANGE || 'Sheet1!A:E';

// ✅ ฟังก์ชันช่วยตรวจหัวตาราง ถ้าไม่มีจะสร้างให้
function headRangeFrom(range) {
  const sheetTitle = (range.includes('!') ? range.split('!')[0] : 'Sheet1');
  return `${sheetTitle}!A1:E1`;
}

// 🧾 ตรวจว่ามีหัวตารางหรือยัง ถ้ายังไม่มีให้เพิ่ม
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

// ➕ เพิ่มข้อมูล 1 แถวในชีต (เช่น บันทึกรายรับ/รายจ่าย)
async function appendRow(values) {
  await ensureHeader();
  return sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: SHEET_RANGE,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [values] }
  });
}

// 📖 อ่านข้อมูลจากชีต (เอาเฉพาะล่าสุด)
async function readRecentRows(limit = 120) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: SHEET_RANGE
  });
  const rows = res.data.values || [];
  const dataRows = rows.length > 1 ? rows.slice(1) : [];
  return dataRows.slice(-limit);
}

// 📅 ฟังก์ชันวันที่ภาษาไทย
function todayTH() {
  return new Date().toLocaleDateString('th-TH');
}

// 🔍 กำหนดคำที่ใช้จำแนกรายรับรายจ่าย
const KW = {
  income: [ 'รายรับ','รับ','ได้','โอนเข้า','เงินเดือน','โบนัส','ทิป','ขายได้','ดอกเบี้ย','ปันผล' ],
  expense: [ 'รายจ่าย','จ่าย','ซื้อ','โอนออก','เติม','ค่าผ่อน','ผ่อน','ค่าสมาชิก','ค่าใช้จ่าย' ],
};

// 📂 หมวดหมู่หลักของรายจ่ายและรายรับ (ใช้เดาอัตโนมัติ)
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

// 🧮 ฟังก์ชันช่วยประมวลผลข้อความ
function normalize(s) { return String(s || '').toLowerCase(); }

function extractAmount(text) {
  const m = String(text).match(/(\d{1,3}(?:[, ]\d{3})*|\d+)(?:[.,](\d{1,2}))?/);
  if (!m) return null;
  const num = (m[1] || '').replace(/[ ,]/g, '');
  const dec = m[2] ? '.' + m[2] : '';
  const val = Number(num + dec);
  return Number.isFinite(val) ? val : null;
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

// 🤖 ฟังก์ชันวิเคราะห์ด้วย Gemini (AI)
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

// 🎨 Flex Message: UI สวย ๆ ให้ผู้ใช้กดยืนยัน
function chip(text, bg = THEME.accentSoft, color = THEME.accent) { /* ... สร้างกล่องแสดงป้ายหมวด ... */
  return {
    type: 'box', layout: 'baseline', backgroundColor: bg, cornerRadius: '12px', paddingAll: '6px',
    contents: [{ type: 'text', text, size: '12px', weight: 'bold', color }]
  };
}

function confirmFlex({ type, amount, category, note, date, payload }) { /* ... หน้ายืนยันก่อนบันทึก ... */
  const isIncome = type === 'รายรับ';
  const title = isIncome ? 'บันทึกรายรับ?' : 'บันทึกรายจ่าย?';
  const icon = isIncome ? '💸' : '🧾';
  const amountTxt = `${amount.toLocaleString()} บ.`;
  const chips = [ chip(type), chip(category) ];

  return {
    type: 'flex',
    altText: title,
    contents: {
      type: 'bubble', size: 'mega', styles: { body: { backgroundColor: '#FFFFFF' } },
      header: {
        type: 'box', layout: 'horizontal', contents: [
          { type: 'text', text: icon + ' ' + title, weight: 'bold', size: 'md', color: THEME.textStrong },
          { type: 'button', action: { type: 'postback', label: '✖', data: 'action=cancel', displayText: 'ยกเลิก' }, style: 'secondary', color: THEME.danger, height: 'sm' }
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

function buildSavePayload({ type, amount, category, note }) { /* ... แปลงข้อมูลเป็น URL สำหรับ postback ... */
  const shortNote = String(note || '').slice(0, 40);
  const params = new URLSearchParams({ action: 'save', type, amount: String(amount), category, note: shortNote });
  return params.toString();
}

// 🌐 เส้นทางหลักของเว็บเซิร์ฟเวอร์ (Route)
app.get('/webhook', (req, res) => { // ตรวจสุขภาพ
  res.status(200).send('OK');
});

// 📨 LINE จะยิงมาที่นี่เมื่อมีข้อความหรือปุ่มกดจากผู้ใช้
app.post('/webhook', middleware(lineConfig), (req, res) => {
  try {
    res.sendStatus(200); // ตอบกลับทันทีเพื่อไม่ให้ Timeout
    const body = (req && req.body && typeof req.body === 'object') ? req.body : {};
    const events = Array.isArray(body.events) ? body.events : [];
    Promise.all(events.map(e => handleEvent(e))).catch(err => {
      console.error('handleEvent error:', err?.response?.data || err.message || err);
    });
  } catch (err) {
    console.error('webhook handler crash:', err);
  }
});

// ✅ หน้าแรก แสดงข้อความเวลาบอทออนไลน์แล้ว
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

// ============================ FLEX DASHBOARD MENU ============================
// ฟังก์ชันนี้สร้าง Flex Message เมนูแดชบอร์ดมินิมอล
// มีปุ่มเลือกช่วงเวลา: 1 year / 6 เดือน / 3 เดือน / 1 เดือน / 1 week
// และปุ่ม Export Excel สีเขียวเด่น
function buildDashboardMenuFlex() {
  const GREEN = '#16A34A'; // สีเขียวเด่นสำหรับปุ่ม Export
  const TEXT = '#111111'; // สีข้อความหลัก
  const MUTED = '#8B95A1'; // สีข้อความรอง

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
          {
            type: 'box',
            layout: 'horizontal',
            spacing: '8px',
            contents: [
              { type: 'button', style: 'secondary', height: 'sm', action: { type: 'postback', label: '1 year', data: 'dash?range=1y' } },
              { type: 'button', style: 'secondary', height: 'sm', action: { type: 'postback', label: '6 เดือน', data: 'dash?range=6m' } }
            ]
          },
          {
            type: 'box',
            layout: 'horizontal',
            spacing: '8px',
            contents: [
              { type: 'button', style: 'secondary', height: 'sm', action: { type: 'postback', label: '3 เดือน', data: 'dash?range=3m' } },
              { type: 'button', style: 'secondary', height: 'sm', action: { type: 'postback', label: '1 เดือน', data: 'dash?range=1m' } }
            ]
          },
          {
            type: 'box',
            layout: 'horizontal',
            contents: [
              { type: 'button', style: 'secondary', height: 'sm', action: { type: 'postback', label: '1 week', data: 'dash?range=1w' } }
            ]
          }
        ]
      },
      footer: {
        type: 'box',
        layout: 'vertical',
        paddingAll: '16px',
        contents: [
          {
            type: 'button',
            style: 'primary',
            height: 'sm',
            color: GREEN,
            action: { type: 'postback', label: 'Export เป็น Excel', data: 'dash?action=export_excel' }
          }
        ]
      },
      styles: {
        header: { backgroundColor: '#FFFFFF' },
        body:   { backgroundColor: '#FFFFFF' },
        footer: { backgroundColor: '#FFFFFF' }
      }
    }
  };
}
// ============================ END FLEX DASHBOARD MENU ============================



// 🧠 ส่วนหลักที่ใช้ตอบกลับข้อความผู้ใช้
async function buildDashboardImages(userId) {
  const rows = await readRecentRowsForUser(userId, 500);
  if (rows.length === 0) return null;

  const now = new Date();
  const thisMonthRows = rows.filter(r => {
    const d = parseThaiDate(r[0]);
    return d && isSameMonth(d, now);
  });
  if (thisMonthRows.length === 0)
    return { note: 'เดือนนี้ยังไม่มีรายการเลยครับ ลองบันทึกเพิ่มก่อนนะ' };

  let sumIncome = 0;
  let sumExpense = 0;
  const catExpense = {};
  for (const r of thisMonthRows) {
    const type = r[1];
    const amount = Number(r[2] || 0);
    const cat = normalizeCategory(r[3]);
    if (type === 'รายรับ') sumIncome += amount;
    else if (type === 'รายจ่าย') {
      sumExpense += amount;
      catExpense[cat] = (catExpense[cat] || 0) + amount;
    }
  }
  const balance = sumIncome - sumExpense;

  // ---------- Chart ----------
  const barConfig = {
    type: 'bar',
    data: {
      labels: ['เดือนนี้'],
      datasets: [
        { label: 'รายรับ', data: [sumIncome], backgroundColor: '#4CAF50' },
        { label: 'รายจ่าย', data: [sumExpense], backgroundColor: '#F44336' }
      ]
    }
  };

  const pieConfig = {
    type: 'pie',
    data: {
      labels: Object.keys(catExpense),
      datasets: [{ data: Object.values(catExpense) }]
    }
  };

  const barUrl = buildQuickChartUrl(barConfig);
  const pieUrl = buildQuickChartUrl(pieConfig, { w: 600, h: 600 });

  return {
    note: `รายรับรวม: ${sumIncome.toLocaleString()} บาท\nรายจ่ายรวม: ${sumExpense.toLocaleString()} บาท\nคงเหลือ: ${balance.toLocaleString()} บาท`,
    barUrl,
    pieUrl
  };
}


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

  // ถ้าผู้ใช้พิมพ์ "แดชบอร์ด" ให้เรียกเมนู Flex ที่สร้างไว้
  if (/^แดชบอร์ด$/i.test(text)) {
  const flex = buildDashboardMenuFlex();
  return lineClient.replyMessage(event.replyToken, flex);
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

  if (/^แดชบอร์ด$/i.test(text)) {
  const dash = await buildDashboardImages(userId);
  if (!dash) {
    return lineClient.replyMessage(event.replyToken, { type: 'text', text: 'ยังไม่มีข้อมูลครับ' });
  }
  if (dash.note) {
    return lineClient.replyMessage(event.replyToken, {
      type: 'text',
      text: `${dash.note}\n\n${dash.barUrl}\n${dash.pieUrl}`
    });
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

// ============================ FLEX DASHBOARD POSTBACK ============================
// ฟังก์ชันนี้ทำงานเมื่อผู้ใช้กดปุ่มใน Flex Message
function parsePostbackData(data) {
  const params = {};
  const [, qs] = String(data).split('?');
  if (qs) qs.split('&').forEach(kv => {
    const [k, v] = kv.split('=');
    params[k] = v;
  });
  return params;
}

// ใช้ใน handleEvent(event) เมื่อ event.type === 'postback'
async function handlePostback(event) {
  const data = String(event.postback?.data || '');
  const params = parsePostbackData(data);

  // ✅ ถ้าผู้ใช้เลือกช่วงเวลา
  if (data.startsWith('dash?range=')) {
    const range = params.range;
    return lineClient.replyMessage(event.replyToken, {
      type: 'text',
      text: `กำลังเตรียมแดชบอร์ดช่วงเวลา: ${range}`
    });
  }

  // ✅ ถ้าผู้ใช้กดปุ่ม Export Excel
  if (params.action === 'export_excel') {
    return lineClient.replyMessage(event.replyToken, {
      type: 'text',
      text: 'กำลังสร้างไฟล์ Excel ให้ครับ...'
    });
  }
}
// ============================ END FLEX DASHBOARD POSTBACK ============================


// 🚀 เริ่มรันเซิร์ฟเวอร์บนพอร์ตที่กำหนด (Railway จะใช้ PORT อัตโนมัติ)
const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});