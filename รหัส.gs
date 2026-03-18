function doGet(e) {
  const MY_SECRET_PASS = "1751252"; // ตั้งรหัสลับของคุณเองที่นี่

  // เช็คว่ามี Parameter "text" และ "key" ส่งมาถูกต้องไหม
  if (e && e.parameter && e.parameter.text) {
    if (e.parameter.key !== MY_SECRET_PASS) {
      return ContentService.createTextOutput("Error: Unauthorized Access")
        .setMimeType(ContentService.MimeType.TEXT);
    }

    const result = parseInput(e.parameter.text);
    return ContentService.createTextOutput("บันทึกสำเร็จ: " + e.parameter.text)
      .setMimeType(ContentService.MimeType.TEXT);
  }

  // แสดงหน้าเว็บปกติ
  return HtmlService.createHtmlOutputFromFile("index")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, minimum-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function parseInput(text){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ledger");
  text = text.trim().toLowerCase();
  const now = new Date();

  if(text.includes(">")){
    const [left, right] = text.split(">").map(s => s.trim().split(" "));
    const amount      = Number(left[0]);
    const fromAccount = left[1], fromWallet = left[2];
    const toAccount   = right[0], toWallet  = right[1];
    const ref         = Utilities.getUuid();
    sheet.appendRow([Utilities.getUuid(), ref, "transfer", now, fromAccount, fromWallet, amount, -amount, "", "transfer", now]);
    sheet.appendRow([Utilities.getUuid(), ref, "transfer", now, toAccount,   toWallet,   amount,  amount, "", "transfer", now]);
    return "transfer ok";
  }

  const parts   = text.split(" ");
  let rawAmount = parts[0];
  const note    = parts[1];
  const account = parts[2];
  const wallet  = parts[3] || "spend";
  let type, delta, amount;

  if(rawAmount.startsWith("+")){
    type   = "income";
    amount = Number(rawAmount.substring(1));
    delta  = amount;
  } else {
    type   = "expense";
    amount = Number(rawAmount);
    delta  = -amount;
  }

  sheet.appendRow([Utilities.getUuid(), "", type, now, account, wallet, amount, delta, "", note, now]);
  return "ok";
}

function getBalances(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ledger");
  const data  = sheet.getDataRange().getValues();
  const map   = {};
  for(let i = 1; i < data.length; i++){
    const acc = String(data[i][4]).trim().toLowerCase();
    const wal = String(data[i][5]).trim().toLowerCase();
    const key = acc + "_" + wal;
    map[key]  = (map[key] || 0) + data[i][7];
  }
  return map;
}

function getMeta(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const accounts = ss.getSheetByName("accounts").getRange("A2:A").getValues().flat().filter(String).map(s=>s.trim().toLowerCase());
  const wallets  = ss.getSheetByName("wallets").getRange("A2:A").getValues().flat().filter(String).map(s=>s.trim().toLowerCase());
  return { accounts, wallets };
}

// --- ฟังก์ชันใหม่สำหรับ AI ---
function getAISummary() {
  // ต้องมี " " คร่อม Key แบบนี้นะครับ
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('GEMINI_KEY');
  
  if (!apiKey) return "ยังไม่ได้ใส่ API Key ครับ";

  try {
    const balances = getBalances();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ledger");
    if (!sheet) return "ไม่พบ Sheet ชื่อ ledger ครับ"; // ป้องกันกรณีหาชีทไม่เจอ
    
    const data = sheet.getDataRange().getValues();
    const now = new Date();
    const currentMonth = now.getMonth();
    
    const recentTransactions = data.slice(1)
      .filter(row => {
        const d = new Date(row[3]);
        return d instanceof Date && !isNaN(d) && d.getMonth() === currentMonth;
      })
      .slice(-15)
      .map(row => `${row[2]}: ${row[6]} (${row[9]})`)
      .join(", ");

    const context = `ยอดคงเหลือ: ${JSON.stringify(balances)}. รายการล่าสุด: ${recentTransactions}`;
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=${apiKey}`;
    
    const payload = {
      "contents": [{ "parts": [{ "text": `คุณคือ AI ผู้ช่วยการเงิน ข้อมูลคือ ${context}. อ่านข้อมูลการเงินแล้วสรุป พร้อมบอกค่าใช้จ่ายล่าสุดย้อนไป 5 ครั้ง` }] }]
    };

    const res = UrlFetchApp.fetch(url, {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true // ใส่ไว้เพื่อดู Error ชัดๆ ถ้าพัง
    });
    
    const json = JSON.parse(res.getContentText());
    
    if (json.candidates && json.candidates[0].content) {
       return json.candidates[0].content.parts[0].text;
    } else {
       return "API ตอบกลับแต่ไม่มีข้อมูล: " + res.getContentText();
    }

  } catch (e) {
    // ถ้ายังขึ้นอันนี้อีก ให้ลองเปลี่ยนเป็น return e.toString(); เพื่อดู Error จริงครับ
    return "เกิดข้อผิดพลาด: " + e.toString();
  }
}
function listMyModels() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;
  
  const res = UrlFetchApp.fetch(url);
  Logger.log(res.getContentText()); // ก๊อปปี้สิ่งที่ขึ้นใน Log มาวางให้ผมดูครับ
}
