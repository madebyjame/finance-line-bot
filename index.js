/**
 * LINE Bot สำหรับบันทึกรายรับ-รายจ่าย อัตโนมัติลงใน Google Sheets
 * - บันทึกรายรับ/รายจ่ายจากแชต (รองรับเลขใหญ่/คอมม่า)
 * - วิเคราะห์ด้วย Gemini (ต่อผู้ใช้)
 * - แดชบอร์ดพร้อมปุ่มช่วงเวลา (ต่อผู้ใช้)
 * - แยกข้อมูลเป็น "แท็บชีตต่อผู้ใช้" อัตโนมัติ
 */
require('dotenv').config();

const fs = require("fs");
const express = require('express');
const { Client, middleware } = require('@line/bot-sdk');
const { google } = require('googleapis');        // ✅ ต้องมี
const axios = require('axios');                  // ✅ ต้องมี

const app = express();

if (process.env.GOOGLE_CREDENTIALS_JSON) {
  const path = "/tmp/google.json";
  fs.writeFileSync(path, process.env.GOOGLE_CREDENTIALS_JSON);
  process.env.GOOGLE_SERVICE_ACCOUNT_FILE = path;
}

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

// ========= MONTHLY FACTS HELPERS =========
// แปลงวันที่ไทย -> key เดือน YYYY-MM
function monthKeyFromThaiDate(s) {
  const d = parseThaiDate(s);
  if (!d) return 'unknown';
  const y = d.getFullYear();
  const m = (d.getMonth()+1).toString().padStart(2, '0');
  return `${y}-${m}`;
}

// รวมยอดเป็นรายเดือน พร้อม breakdown
async function buildMonthlyFactsForUser(userId, months = 2) {
  const rows = await readRecentRowsForUser(userId, 2000);
  if (!rows || rows.length === 0) return { monthsData: [], keys: [] };

  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth() - (months - 1), 1);

  const monthly = {}; // { 'YYYY-MM': { income, expense, byCategory: {cat: sumExpense} } }
  for (const r of rows) {
    const [date, type, amountRaw, categoryRaw] = r;
    const d = parseThaiDate(date);
    if (!d || d < start) continue;

    const mk = monthKeyFromThaiDate(date);
    const amt = Number(String(amountRaw || '0').toString().replace(/,/g, '')) || 0;
    const cat = categoryRaw || 'อื่นๆ';
    if (!monthly[mk]) monthly[mk] = { income: 0, expense: 0, byCategory: {} };

    if (type === 'รายรับ') {
      monthly[mk].income += amt;
    } else if (type === 'รายจ่าย') {
      monthly[mk].expense += amt;
      monthly[mk].byCategory[cat] = (monthly[mk].byCategory[cat] || 0) + amt;
    }
  }

  const keys = Object.keys(monthly).sort().slice(-months);
  const monthsData = keys.map(k => {
    const m = monthly[k];
    const balance = m.income - m.expense;
    const catSorted = Object.entries(m.byCategory)
      .map(([cat, sum]) => ({ category: cat, sum }))
      .sort((a,b) => b.sum - a.sum);
    return { key: k, income: m.income, expense: m.expense, balance, topCats: catSorted.slice(0, 5) };
  });

  return { monthsData, keys };
}

// ============================ END PER-USER SHEET HELPERS ============================

// สร้างข้อเท็จจริง (facts) สำหรับช่วง N เดือนล่าสุด
async function buildFinanceFactsForUser(userId, months = 3) {
  const rows = await readRecentRowsForUser(userId, 1000);
  if (!rows || rows.length === 0) return { rows: [], facts: null, text: 'NO_DATA' };

  const now = new Date();
  const cutoff = new Date(now.getFullYear(), now.getMonth() - (months - 1), 1);
  const filtered = rows.filter(r => {
    const d = parseThaiDate(r[0]);
    return d && d >= cutoff;
  });

  let income = 0, expense = 0;
  const byCategory = {};
  const byMonth = {};
  const singles = [];

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

  const categoryShares = Object.entries(byCategory)
    .map(([cat, sum]) => ({ category: cat, sum, share: expense > 0 ? sum / expense : 0 }))
    .sort((a,b) => b.sum - a.sum);

  const amounts = singles.map(s => s.amount).sort((a,b) => a-b);
  const median = amounts.length ? (amounts[Math.floor(amounts.length/2)] + amounts[Math.ceil(amounts.length/2)-1]) / 2 : 0;
  const mean = amounts.length ? (amounts.reduce((a,b)=>a+b,0)/amounts.length) : 0;

  const interestingCats = categoryShares.filter(x =>
    x.share >= FINANCE_RULES.CATEGORY_SHARE_ALERT ||
    x.sum >= FINANCE_RULES.MIN_MONTHLY_CATEGORY_SUM
  );

  const outliers = singles.filter(s =>
    s.amount >= FINANCE_RULES.MIN_SINGLE_IGNORE &&
    (median > 0 ? s.amount >= FINANCE_RULES.OUTLIER_MULTIPLIER * median : s.amount >= FINANCE_RULES.MIN_MONTHLY_CATEGORY_SUM)
  ).sort((a,b) => b.amount - a.amount).slice(0, 10);

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
    categoryShares,
    interestingCats,
    outliers,
    monthOverMonth: mom,
    rules: FINANCE_RULES,
  };

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

// ============================ Modern Mobile Dashboard + Category Summary ============================
// ============================ Dashboard + Category Summary Text ============================
async function buildDashboardImages(userId, rangeCode = '1m') {
  const rows = await readRecentRowsForUser(userId, 1000);
  if (!rows || rows.length === 0) return { note: 'ยังไม่มีข้อมูลเลยครับ ลองบันทึกก่อนนะ' };

  const { start, end } = getRangeDates(rangeCode);
  const inRange = rows.filter(r => isWithin(parseThaiDate(r[0]), start, end));
  if (inRange.length === 0) return { note: 'ช่วงเวลานี้ยังไม่มีรายการครับ' };

  let sumIncome = 0, sumExpense = 0;
  const catExpense = {};
  const catIncome  = {};

  for (const r of inRange) {
    const type = r[1];
    const amt = Number(String(r[2] ?? '0').toString().replace(/,/g, '')) || 0;
    const cat = r[3] || 'อื่นๆ';
    if (type === 'รายรับ') {
      sumIncome += amt;
      catIncome[cat]  = (catIncome[cat]  || 0) + amt;
    } else if (type === 'รายจ่าย') {
      sumExpense += amt;
      catExpense[cat] = (catExpense[cat] || 0) + amt;
    }
  }

  const balance = sumIncome - sumExpense;
  const pretty = (n) => Number(n).toLocaleString();

  // ==== เรียงหมวดและรวมยอด ====
  const expEntries = Object.entries(catExpense).sort((a,b) => b[1]-a[1]);
  const incEntries = Object.entries(catIncome ).sort((a,b) => b[1]-a[1]);
  const totalExp = expEntries.reduce((a,[,v]) => a+v, 0);
  const totalInc = incEntries.reduce((a,[,v]) => a+v, 0);

  // ==== สรุปแบบข้อความก้อนเดียว (รายหมวด) ====
  // รวมทุกหมวดเป็นรายการเดียวกันเพื่ออ่านง่ายบนมือถือ
  const allCats = Array.from(new Set([...Object.keys(catExpense), ...Object.keys(catIncome)]));
  // เรียงตาม “จ่าย” ก่อน แล้วตาม “รับ”
  const both = allCats
    .map(name => ({ name, inc: catIncome[name]||0, exp: catExpense[name]||0 }))
    .sort((a,b) => (b.exp - a.exp) || (b.inc - a.inc));

  const MAX_LINES = 12; // กันยาวเกินใน LINE
  const lines = both.slice(0, MAX_LINES).map((x, i) => {
    const pExp = totalExp ? Math.round((x.exp/totalExp)*100) : 0;
    const pInc = totalInc ? Math.round((x.inc/totalInc)*100) : 0;
    return `${i+1}. ${x.name} — รับ ${pretty(x.inc)} บ. (${pInc}%) | จ่าย ${pretty(x.exp)} บ. (${pExp}%)`;
  });
  if (both.length > MAX_LINES) lines.push(`… และอีก ${both.length - MAX_LINES} หมวด`);

  const catSummaryText =
    (both.length
      ? [
          '📊 สรุปตามหมวด (รับ | จ่าย)',
          ...lines
        ].join('\n')
      : 'ยังไม่มีข้อมูลรายรับ/รายจ่ายตามหมวดในช่วงเวลานี้');

  // 🎨 โทนสี + กราฟ (เหมือนเดิมที่ปรับอ่านง่าย)
  const COLORS = {
    income: '#22C55E',
    expense: '#EF4444',
    text: '#111827',
    grid: '#E5E7EB',
    bg: '#FFFFFF'
  };

  const barConfig = {
    type: 'bar',
    data: {
      labels: ['สรุปรายรับ/รายจ่ายรวม'],
      datasets: [
        { label: 'รายรับ', data: [sumIncome],  backgroundColor: COLORS.income,  borderWidth: 3 },
        { label: 'รายจ่าย', data: [sumExpense], backgroundColor: COLORS.expense, borderWidth: 3 }
      ]
    },
    options: {
      plugins: {
        legend: { position: 'bottom', labels: { font: { size: 32 }, color: COLORS.text } },
        datalabels: {
          display: true, color: COLORS.text, anchor: 'end', align: 'top',
          font: { size: 36, weight: 'bold' }, formatter: v => v.toLocaleString()
        }
      },
      scales: {
        y: { ticks: { color: COLORS.text, font: { size: 28 } }, grid: { color: COLORS.grid } },
        x: { ticks: { color: COLORS.text, font: { size: 28 } } }
      },
      layout: { padding: 30 },
      backgroundColor: COLORS.bg
    }
  };

  const pieLabels = expEntries.map(([name]) => name);
  const pieData   = expEntries.map(([,val]) => val);
  const pieTotal  = pieData.reduce((a,b)=>a+b,0);
  const pieColors = ['#60A5FA','#34D399','#FBBF24','#F87171','#A78BFA','#F472B6','#FCD34D','#4ADE80','#93C5FD','#C084FC'];

  const pieConfig = {
    type: 'pie',
    data: {
      labels: pieLabels.length ? pieLabels : ['ไม่มีรายจ่าย'],
      datasets: [{
        data: pieData.length ? pieData : [1],
        backgroundColor: pieLabels.length ? pieColors.slice(0, pieLabels.length) : ['#E5E7EB'],
        borderColor: '#FFFFFF', borderWidth: 3
      }]
    },
    options: {
      plugins: {
        legend: { display: false },
        datalabels: {
          display: true, color: '#FFFFFF', font: { size: 36, weight: 'bold' },
          formatter: (v, ctx) => {
            const label = ctx.chart.data.labels[ctx.dataIndex];
            const percent = pieTotal ? Math.round((v / pieTotal) * 100) : 0;
            return `${label}\n${percent}%`;
          }
        },
        tooltip: {
          callbacks: {
            label: (context) => {
              const label = context.label || '';
              const value = context.parsed || 0;
              return `${label}: ${pretty(value)} บาท`;
            }
          }
        }
      },
      layout: { padding: 20 },
      backgroundColor: COLORS.bg
    }
  };

  const catBarLabels = pieLabels;
  const catBarData   = pieData;
  const catBarConfig = {
    type: 'bar',
    data: {
      labels: catBarLabels.length ? catBarLabels : ['ไม่มีรายจ่าย'],
      datasets: [{
        label: 'ยอดรายจ่ายตามหมวด (บาท)',
        data: catBarData.length ? catBarData : [1],
        backgroundColor: '#3B82F6',
        borderWidth: 2
      }]
    },
    options: {
      indexAxis: 'y',
      plugins: {
        legend: { position: 'bottom', labels: { font: { size: 28 }, color: COLORS.text } },
        datalabels: {
          display: true, color: COLORS.text, anchor: 'end', align: 'right',
          font: { size: 32, weight: 'bold' }, formatter: v => pretty(v)
        }
      },
      scales: {
        x: { ticks: { color: COLORS.text, font: { size: 26 } }, grid: { color: COLORS.grid } },
        y: { ticks: { color: COLORS.text, font: { size: 30 } }, grid: { color: COLORS.grid } }
      },
      layout: { padding: 24 },
      backgroundColor: COLORS.bg
    }
  };

  const barUrl    = buildQuickChartUrl(barConfig,    { w: 900,  h: 700  });
  const pieUrl    = buildQuickChartUrl(pieConfig,    { w: 900,  h: 900  });
  const catBarUrl = buildQuickChartUrl(catBarConfig, { w: 1000, h: 1100 });

// === สร้างข้อความสรุปหมวดหมู่ แยกเป็น รายรับ / รายจ่าย ===
const prettyFormat = (n) => Number(n).toLocaleString();

// เรียงหมวดตามยอดมากไปน้อย
const incEntriesSorted = Object.entries(catIncome).sort((a, b) => b[1] - a[1]);
const expEntriesSorted = Object.entries(catExpense).sort((a, b) => b[1] - a[1]);

// จำกัดจำนวนแถวให้พอดีบนมือถือ (ถ้าเยอะเกิน)
const MAX_SHOW = 8;
const incLines = incEntriesSorted.slice(0, MAX_SHOW).map(([name, val]) => `• ${name}: +${prettyFormat(val)} บ.`);
const expLines = expEntriesSorted.slice(0, MAX_SHOW).map(([name, val]) => `• ${name}: -${prettyFormat(val)} บ.`);

// สรุปข้อความทั้งหมด (รวมเข้า note เดิม)
const note = [
  `📅 ช่วง: ${start.toLocaleDateString('th-TH')} – ${end.toLocaleDateString('th-TH')}`,
  `💚 รายรับรวม: ${pretty(sumIncome)} บาท`,
  `❤️ รายจ่ายรวม: ${pretty(sumExpense)} บาท`,
  `💰 คงเหลือ: ${pretty(balance)} บาท`,
  '',
  '💵 รายรับ:',
  ...(incLines.length ? incLines : ['(ไม่มีรายรับในช่วงนี้)']),
  '',
  '🧾 รายจ่าย:',
  ...(expLines.length ? expLines : ['(ไม่มีรายจ่ายในช่วงนี้)'])
].join('\n');


  return {
    note,
    barUrl,
    pieUrl,
    catBarUrl,
    catSummaryText // <<==== เพิ่มก้อนข้อความสรุปรายหมวด
  };
}

// =========================== END DASHBOARD HELPERS ============================

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

// วิเคราะห์เฉพาะของผู้ใช้ (แฟกต์เบส + กติกา)
async function analyzeWithGeminiForUser(userId) {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) return 'ยังไม่ได้ตั้งค่า GEMINI_API_KEY ใน .env นะ';

  const { facts, text } = await buildFinanceFactsForUser(userId, 3);
  if (!facts) return 'ยังไม่มีข้อมูลของคุณเลยครับ ลองบันทึกก่อนนะ';

  const monthsKeys = facts?.monthOverMonth ? [facts.monthOverMonth.prev?.key, facts.monthOverMonth.last?.key].filter(Boolean).join(' → ') : '-';
  const momInc = facts?.monthOverMonth ? facts.monthOverMonth.diff.income : 0;
  const momExp = facts?.monthOverMonth ? facts.monthOverMonth.diff.expense : 0;

  const topCats = (facts.categoryShares || []).slice(0, 5)
    .map(x => `${x.category}:${Math.round(x.sum).toLocaleString()}(${Math.round(x.share*100)}%)`).join(', ') || '-';

  const outlierLines = (facts.outliers || []).slice(0,5)
    .map(o => `${o.date} ${o.category} ${o.amount.toLocaleString()}`).join('\n') || '-';

  const prompt = `
คุณเป็นผู้ช่วยวิเคราะห์การเงินส่วนบุคคล ให้สรุปข้อมูลล่าสุดแบบอ่านง่าย ไม่ใช้ Markdown
ข้อมูลแถวต่อแถว:
${text}

สรุปตัวเลขรวม (3 เดือนล่าสุด):
- รายรับรวม: ${facts.totalIncome.toLocaleString()}
- รายจ่ายรวม: ${facts.totalExpense.toLocaleString()}
- คงเหลือรวม: ${(facts.balance).toLocaleString()}
- เทียบเดือนต่อเดือน (keys: ${monthsKeys}) → รายรับ Δ ${momInc.toLocaleString()}, รายจ่าย Δ ${momExp.toLocaleString()}
- หมวดใช้จ่ายนำ: ${topCats}
- รายการเดี่ยวที่สูงกว่าปกติ (ตัวอย่าง): 
${outlierLines}

กรุณาตอบ:
1) เดือนล่าสุดควรจับตาหมวดไหน เพราะอะไร (สั้น)
2) คำแนะนำเชิงปฏิบัติ 2-3 ข้อ (ตั้งเพดาน/ลดความถี่/ย้ายเป็นคงที่ ฯลฯ)
3) สรุปรวมว่า “ภาพรวมปกติ” หรือ “ควรปรับ” พร้อมเหตุผลสั้น ๆ
`.trim();

  for (const model of MODEL_LIST) {
    try {
      const msg = await callGeminiModel({ model, apiKey, prompt });
      return msg;
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
    type: 'box',
    layout: 'baseline',
    backgroundColor: bg,
    cornerRadius: 'lg',       // เดิม '12px'
    paddingAll: 'sm',         // เดิม '6px'
    contents: [
      { type: 'text', text, size: 'xs', weight: 'bold', color } // เดิม '12px'
    ]
  };
}

// ===== FINANCE RULES (เกณฑ์เพื่อกันแนะนำมั่ว) =====
const FINANCE_RULES = {
  MIN_SINGLE_IGNORE: 100,
  MIN_MONTHLY_CATEGORY_SUM: 1000,
  CATEGORY_SHARE_ALERT: 0.15,
  OUTLIER_MULTIPLIER: 1.3,
  ESSENTIAL_CATEGORIES: ['บิล/สาธารณูปโภค', 'ค่าเช่า', 'ที่อยู่อาศัย', 'การเดินทางจำเป็น'],
};


function confirmFlex({ type, amount, category, note, date, payload }) {
  const isIncome = type === 'รายรับ';
  const title = isIncome ? 'บันทึกรายรับ นี้ใช่ไหมครับ?' : 'บันทึกรายจ่าย นี้ใช่ไหมครับ ?';
  const icon = isIncome ? '💸' : '🧾';
  const amountTxt = `${amount.toLocaleString()} บาท`;
  const chips = [ chip(type), chip(category) ];

  // ตั้งสีปุ่มตามประเภท
  const buttonColor = isIncome ? '#16A34A' : '#DC2626'; // เขียว/แดง

  return {
    type: 'flex',
    altText: title,
    contents: {
      type: 'bubble',
      size: 'kilo',
      styles: {
        header: { backgroundColor: '#FFFFFF' },
        body:   { backgroundColor: '#FFFFFF' },
        footer: { backgroundColor: '#FFFFFF' }
      },
      header: {
        type: 'box',
        layout: 'horizontal',
        justifyContent: 'flex-start',
        paddingAll: 'md',
        contents: [
          { type: 'text', text: icon + ' ' + title, weight: 'bold', size: 'md', color: THEME.textStrong }
        ]
      },
      body: {
        type: 'box',
        layout: 'vertical',
        spacing: 'md',
        paddingAll: 'md',
        contents: [
          { type: 'text', text: amountTxt, size: 'xl', weight: 'bold', color: THEME.textStrong },
          { type: 'text', text: note || '-', size: 'sm', color: THEME.textMuted, wrap: true },
          { type: 'text', text: date, size: 'xs', color: THEME.textMuted },
          {
            type: 'box',
            layout: 'horizontal',
            spacing: 'sm',
            contents: chips
          }
        ]
      },
      footer: {
        type: 'box',
        layout: 'horizontal',
        spacing: 'md',
        paddingAll: 'md',
        contents: [
          {
            type: 'button',
            style: 'primary',
            color: buttonColor, // ✅ ใช้สีตามประเภท
            height: 'sm',
            action: { type: 'postback', label: 'บันทึก', data: payload, displayText: 'บันทึก' }
          },
          {
            type: 'button',
            style: 'link',
            height: 'sm',
            action: { type: 'postback', label: 'ยกเลิก', data: 'action=cancel', displayText: 'ยกเลิก' }
          }
        ]
      }
    }
  };
}



function buildDashboardMenuFlex() {
  const GREEN = '#16A34A';
  const TEXT = '#111111';
  const MUTED = '#8B95A1';

  // ปุ่มแบ่ง 2 แถว (3 + 2) กันล้นแนวนอนจอเล็ก
  const row1 = [
    { type: 'button', style: 'link', action: { type: 'postback', label: '1 year',  data: 'action=dash&range=1y' } },
    { type: 'button', style: 'link', action: { type: 'postback', label: '6 เดือน', data: 'action=dash&range=6m' } },
    { type: 'button', style: 'link', action: { type: 'postback', label: '3 เดือน', data: 'action=dash&range=3m' } },
  ];
  const row2 = [
    { type: 'button', style: 'link', action: { type: 'postback', label: '1 เดือน', data: 'action=dash&range=1m' } },
    { type: 'button', style: 'link', action: { type: 'postback', label: '1 week', data: 'action=dash&range=1w' } },
  ];

  return {
    type: 'flex',
    altText: 'แดชบอร์ดการเงิน',
    contents: {
      type: 'bubble',
      size: 'kilo',
      styles: {
        header: { backgroundColor: '#FFFFFF' },
        body:   { backgroundColor: '#FFFFFF' },
        footer: { backgroundColor: '#FFFFFF' }
      },
      header: {
        type: 'box',
        layout: 'vertical',
        paddingAll: 'md',     // เดิม '16px'
        contents: [
          { type: 'text', text: 'แดชบอร์ดการเงิน', weight: 'bold', size: 'lg', color: TEXT },
          { type: 'text', text: 'เลือกช่วงเวลาที่ต้องการดูสรุป', size: 'xs', color: MUTED, margin: 'sm', wrap: true }
        ]
      },
      body: {
        type: 'box',
        layout: 'vertical',
        spacing: 'md',        // เดิม '12px'
        paddingAll: 'md',     // เดิม '16px'
        contents: [
          { type: 'box', layout: 'horizontal', spacing: 'sm', contents: row1 }, // เดิม spacing '8px'
          { type: 'box', layout: 'horizontal', spacing: 'sm', contents: row2 }
        ]
      },
      footer: {
        type: 'box',
        layout: 'vertical',
        paddingAll: 'md',     // เดิม '16px'
        contents: [
          {
            type: 'button',
            style: 'primary',
            height: 'sm',
            color: GREEN,
            action: { type: 'postback', label: 'Export เป็น Excel', data: 'action=dash&do=export_excel' }
          }
        ]
      }
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

app.get('/healthz', (_req, res) => res.status(200).json({ ok: true, ts: Date.now() })); // ✅ health check

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

  // ===== สรุปรายเดือน + เทียบเดือนก่อน =====
  if (/^(สรุปรายเดือน|สรุปเดือนนี้|รายเดือน|สรุปเดือนล่าสุด)$/i.test(text)) {
    const uid = event.source?.userId;
    if (!uid) {
      return lineClient.replyMessage(event.replyToken, { type: 'text', text: 'ตรวจไม่พบ userId ของคุณ' });
    }
    await lineClient.replyMessage(event.replyToken, { type: 'text', text: 'กำลังสรุปรายเดือนให้ครับ...' });
    analyzeWithGeminiForUser(uid) // ✅ ใช้ฟังก์ชันที่มีอยู่จริง
      .then(msg => lineClient.pushMessage(uid, { type: 'text', text: msg }))
      .catch(err => lineClient.pushMessage(uid, {
        type: 'text',
        text: `เรียก AI ไม่ได้: ${err?.response?.data?.error?.message || err.message || 'unknown'}`
      }));
    return;
  }

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
