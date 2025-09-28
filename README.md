# Finance Line Bot

บอท LINE สำหรับบันทึก **รายรับ-รายจ่าย** และ **วิเคราะห์การเงินด้วย Gemini AI**  
ข้อมูลทั้งหมดจะถูกบันทึกลง **Google Sheet** เพื่อสะดวกต่อการดูย้อนหลังและทำ Dashboard เพิ่มเติม

---

## 🚀 Features
- บันทึกรายรับ/รายจ่าย ผ่าน LINE Chat เช่น  
  - `รายจ่าย 120 คาเฟ่`  
  - `รายรับ 15000 เงินเดือน`
- วิเคราะห์พฤติกรรมการเงินด้วยคำสั่ง `วิเคราะห์`  
  - AI จะสรุปรายรับ-รายจ่าย  
  - แนะนำ actionable steps (ลด/เพิ่ม/จัดการ DCA)
- ข้อมูลถูกเก็บลง Google Sheets แบบ Real-time  
- รองรับการขยายต่อ เช่น ตั้งเป้าหมาย, แจ้งเตือน, สรุปรายเดือน

---

## 🛠️ Tech Stack
- **Node.js** + **Express**
- **LINE Messaging API**
- **Google Sheets API** (Service Account)
- **Gemini AI (Google Generative Language API)**

