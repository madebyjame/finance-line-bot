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
  return 'Gemini ล่มชั่วคราว รอสักครู่ แล้วค่อย"วิเคราะห์" ใหม่น่ะ';
}

// -------- Routes --------
app.get('/webhook', (req, res) => {
  res.status(200).send('OK');
});

// ปลอดภัย + ไม่ timeout: ตรวจลายเซ็นด้วย middleware และตอบ 200 ไว
app.post('/webhook', middleware(lineConfig), (req, res) => {
  try {
    // ตอบ LINE ให้ก่อน
    res.sendStatus(200);

    // ประมวลผลหลังบ้าน
    const body = (req && req.body && typeof req.body === 'object') ? req.body : {};
    const events = Array.isArray(body.events) ? body.events : [];

    Promise.all(events.map(e => handleEvent(e)))
      .catch(err => {
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
      const r = await axios.get(url, { headers: { 'x-goog-api-key': apiKey }, timeout: AI_TIMEOUT_MS });
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
  if (event.type !== 'message' || event.message.type !== 'text') return;
  const text = event.message.text.trim();
  const today = new Date().toLocaleDateString('th-TH');

  const spendRegex = /^(รายจ่าย)\s+(\d+(?:[.,]\d+)?)\s+(.+)$/i;
  const incomeRegex = /^(รายรับ)\s+(\d+(?:[.,]\d+)?)\s+(.+)$/i;

  // พิมพ์ขึ้นต้นถูก แต่รูปแบบไม่ครบ
  if (/^(รายจ่าย|รายรับ)\b/i.test(text) && !(spendRegex.test(text) || incomeRegex.test(text))) {
    return lineClient.replyMessage(event.replyToken, {
      type: 'text',
      text: 'รูปแบบไม่ครบ ลองแบบนี้ "รายจ่าย 120 คาเฟ่" หรือ "รายรับ 15000 เงินเดือน"'
    });
  }

  // จัดการคำสั่งบันทึก
  if (spendRegex.test(text) || incomeRegex.test(text)) {
    const isSpend = spendRegex.test(text);
    const [, type, amountRaw, category] = (isSpend ? spendRegex : incomeRegex).exec(text);
    const amount = Number(String(amountRaw).replace(',', ''));
    if (!Number.isFinite(amount) || amount <= 0) {
      return lineClient.replyMessage(event.replyToken, {
        type: 'text',
        text: 'ตัวเลขไม่ถูกนะ ลองแบบ "รายจ่าย 120 คาเฟ่"'
      });
    }
    try {
      await appendRow([today, type, amount, category, '-']);
      return lineClient.replyMessage(event.replyToken, {
        type: 'text',
        text: `บันทึกแล้ว: ${type} ${amount.toLocaleString()} บ. • ${category} • ${today} ✅`
      });
    } catch (err) {
      return lineClient.replyMessage(event.replyToken, {
        type: 'text',
        text: 'บันทึกไม่ผ่าน เช็กสิทธิ์ชีต/ค่า .env ก่อนครับ'
      });
    }
  }

  // คำสั่งวิเคราะห์
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
    'พิมพ์ได้แบบนี้:',
    '• รายจ่าย 120 คาเฟ่',
    '• รายรับ 15000 เงินเดือน',
    '• วิเคราะห์ (สรุปให้พร้อมแนวทาง)'
  ].join('\n');
  return lineClient.replyMessage(event.replyToken, { type: 'text', text: help });
}

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
