/**
 * โค้ดฉบับสมบูรณ์ (อัปเดตจัดหน้าตาราง 2_Report_Form ให้เหมือนตัวอย่างเป๊ะ + รูปภาพขึ้นตรงกลางช่อง 100%)
 * สำหรับ Google Sheet: https://docs.google.com/spreadsheets/d/1t44kkGGjRiUCRCYmySL62ohrPX5EBY8PJWU0CCTSmvQ/edit
 *
 * วิธีอัปเดต:
 * 1. ใน Google Sheet ไปที่ Extensions > Apps Script ลบโค้ดเก่าออกทั้งหมด แล้วกด Ctrl + V วางโค้ดนี้ลงไป กด Save (แผ่นดิสก์)
 * 2. เลือกฟังก์ชัน "setupPMGoogleSheet" ด้านบน แล้วกดปุ่ม "Run"
 *    (โค้ดใหม่นี้จะ "เก็บข้อมูลเดิมใน 1_Raw_Data ที่คุณกรอกไว้ทั้งหมด" ไม่ลบทิ้ง และจะสร้าง 2_Report_Form ใหม่ให้สวยงามเหมือนตัวอย่างเป๊ะ!)
 * 3. กด Deploy > Manage deployments (จัดการรายการที่ใช้งาน) > กดรูปดินสอ (Edit) > ช่อง Version เลือก "New version" (เวอร์ชันใหม่) > กด Deploy
 */

const DRIVE_FOLDER_NAME = "PM_Report_Photos_UU";

/**
 * เปิดหน้าแอพกรอกข้อมูลผ่านมือถือโดยตรงจากลิงก์ถาวรของ Google (ทำงานได้ตลอด 24 ชม. แม้ปิดคอมพิวเตอร์)
 */
function doGet(e) {
  return HtmlService.createHtmlOutput(getMobileAppHtml())
    .setTitle("ระบบบันทึกงาน PM มือถือ (Universal Utilities GROUP)")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * แปลงลิงก์ Google Drive ทุกแบบให้เป็นลิงก์ Thumbnail ความละเอียดสูงที่สูตร =IMAGE() ใน Google Sheets แสดงผลได้ 100%
 */
function toSheetImageUrl(urlOrId) {
  if (!urlOrId) return "";
  const str = String(urlOrId).trim();
  if (!str) return "";
  const idMatch = str.match(/[?&]id=([a-zA-Z0-9_-]+)/) || str.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (idMatch && idMatch[1]) {
    return `https://drive.google.com/thumbnail?id=${idMatch[1]}&sz=w1000`;
  }
  return str;
}

/**
 * รับข้อมูลจากแอพมือถือเมื่อกดปุ่ม "บันทึก" แล้วอัปเดตลง Google Sheet ทันที
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(15000);
  try {
    const payload = JSON.parse(e.postData.contents);

    if (payload.action === "rebuild_report") {
      setupPMGoogleSheet();
      return ContentService.createTextOutput(JSON.stringify({ status: "success", rebuilt: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let rawSheet = ss.getSheetByName("1_Raw_Data");
    if (!rawSheet) {
      setupPMGoogleSheet();
      rawSheet = ss.getSheetByName("1_Raw_Data");
    }

    if (payload.action === "clear_all") {
      const master = getMasterEquipmentList();
      const roundStr = payload.round || "2/69";
      const emptyRows = master.map(item => [
        "", roundStr, item.id, item.cat, item.name,
        "", "", "", "", "", "", "", "", "", "", "", "", "", ""
      ]);
      rawSheet.getRange(2, 1, emptyRows.length, 19).setValues(emptyRows);
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({ status: "success", clearedAll: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const dataRange = rawSheet.getDataRange().getValues();
    let targetRow = -1;
    for (let r = 1; r < dataRange.length; r++) {
      if (String(dataRange[r][2]).trim() === String(payload.id).trim()) {
        targetRow = r + 1;
        break;
      }
    }
    if (targetRow === -1) {
      targetRow = rawSheet.getLastRow() + 1;
    }

    // ดึง URL รูปเดิมในแถวนั้นไว้ก่อน (กรณีไม่ได้อัพรูปใหม่ทับ)
    const existingRow = targetRow <= dataRange.length ? dataRange[targetRow - 1] : [];
    const photoUrls = [
      toSheetImageUrl(existingRow[14] || ""),
      toSheetImageUrl(existingRow[15] || ""),
      toSheetImageUrl(existingRow[16] || ""),
      toSheetImageUrl(existingRow[17] || ""),
      toSheetImageUrl(existingRow[18] || "")
    ];

    // ถ้ามีการส่งรูปภาพมาจากมือถือ
    if (payload.photos && Array.isArray(payload.photos) && payload.photos.length > 0) {
      let folder = null;
      for (let i = 0; i < 5; i++) {
        const pData = payload.photos[i];
        if (!pData) {
          photoUrls[i] = "";
        } else if (pData.startsWith("data:image")) {
          if (!folder) {
            const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
            folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
            folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          }
          const parts = pData.split(",");
          const mimeMatch = parts[0].match(/:(.*?);/);
          const mimeType = mimeMatch ? mimeMatch[1] : "image/jpeg";
          const bytes = Utilities.base64Decode(parts[1]);
          const blob = Utilities.newBlob(bytes, mimeType, `Equip_${payload.id}_Photo_${i + 1}_${new Date().getTime()}.jpg`);
          const file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          // ใช้ลิงก์ thumbnail?id=...&sz=w1000 เพื่อให้สูตร =IMAGE() ใน Google Sheets แสดงรูปได้ทันที 100%
          photoUrls[i] = `https://drive.google.com/thumbnail?id=${file.getId()}&sz=w1000`;
        } else if (pData.startsWith("http")) {
          photoUrls[i] = toSheetImageUrl(pData);
        }
      }
    }

    const checks = payload.checks || ["", "", "", "", ""];
    const rowValues = [
      new Date(),
      payload.round || "2/69",
      payload.id,
      payload.category || (existingRow[3] || ""),
      payload.name || (existingRow[4] || ""),
      payload.actionDate || "",
      payload.ampValue !== undefined ? payload.ampValue : "",
      checks[0] !== undefined ? checks[0] : "",
      checks[1] !== undefined ? checks[1] : "",
      checks[2] !== undefined ? checks[2] : "",
      checks[3] !== undefined ? checks[3] : "",
      checks[4] !== undefined ? checks[4] : "",
      payload.calPercent !== undefined ? payload.calPercent : "",
      payload.note !== undefined ? payload.note : "",
      photoUrls[0],
      photoUrls[1],
      photoUrls[2],
      photoUrls[3],
      photoUrls[4]
    ];

    rawSheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
    SpreadsheetApp.flush();

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      equipId: payload.id,
      row: targetRow,
      photoUrls: photoUrls
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function getMasterEquipmentList() {
  const subPumpSubs = [
    { no: 1, title: "ตรวจสอบและบันทึกกระแสไฟฟ้า (Amp.)", freq: "ทุกเดือน", color: "#ffff00", type: "amp" },
    { no: 2, title: "ตรวจเช็คทางกายภาพ", freq: "ทุกเดือน", color: "#ffff00", type: "check", def: "✓" },
    { no: 3, title: "ตรวจเช็ค Guide Rail", freq: "ทุกเดือน", color: "#ffff00", type: "check", def: "✓" },
    { no: 4, title: "ยกทำความสะอาด", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "" },
    { no: 5, title: "ถ่ายน้ำมันเครื่อง", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "" },
    { no: 6, title: "ตรวจเช็คใบพัด", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "" }
  ];
  const selfPrimingSubs = [
    { no: 1, title: "ตรวจสอบและบันทึกกระแสไฟฟ้า (Amp.)", freq: "ทุกเดือน", color: "#ffff00", type: "amp" },
    { no: 2, title: "ตรวจเช็คทางกายภาพ", freq: "ทุกเดือน", color: "#ffff00", type: "check", def: "✓" },
    { no: 3, title: "ทำความสะอาด", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "" },
    { no: 4, title: "ถ่ายน้ำมันเครื่อง", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "" },
    { no: 5, title: "ตรวจเช็คใบพัด", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "" }
  ];
  const chemPumpSubs = [
    { no: 1, title: "ตรวจสอบและบันทึกกระแสไฟฟ้า (Amp.)", freq: "ทุกเดือน", color: "#ffff00", type: "amp" },
    { no: 2, title: "ตรวจเช็ค Diaphragm", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "-" },
    { no: 3, title: "ตรวจเช็ค Valve", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "-" },
    { no: 4, title: "ตรวจเช็ค Seal", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "-" }
  ];
  const valveSubs = [
    { no: 1, title: "ตรวจเช็คการคลายตัวของน็อต", freq: "ทุกเดือน", color: "#ffff00", type: "check", def: "✓" }
  ];
  const meterSubs = [
    { no: 1, title: "ทำความสะอาด", freq: "ทุกเดือน", color: "#ffff00", type: "check", def: "✓" },
    { no: 2, title: "Calibration (No Cer)", freq: "ทุก 3 เดือน", color: "#ed7d31", type: "check", def: "✓" }
  ];
  const levelSwitchSubs = [
    { no: 1, title: "ฉีดล้าง ทำความสะอาด", freq: "ทุก 3 เดือน", color: "#ed7d31", type: "check", def: "-" },
    { no: 2, title: "ตรวจเช็คฟังก์ชั่นการทำงาน", freq: "ทุก 6 เดือน", color: "#7030a0", type: "check", def: "-" }
  ];
  const radarSubs = [
    { no: 1, title: "ตรวจเช็คค่าระดับที่หน้าจอ Smart Phone", freq: "ทุกเดือน", color: "#ffff00", type: "check", def: "✓" },
    { no: 2, title: "ตรวจเช็คค่าระดับที่หน้าจอ Scada", freq: "ทุกเดือน", color: "#ffff00", type: "check", def: "✓" }
  ];

  return [
    { id: 1, page: 1, cat: "เครื่องเติมอากาศ", showCat: true, name: "Aeration Blower (AB-SBR-01)", amp: "40.55", note: "สลับจาก AB-SBR-03 เป็น AB-SBR-01", minRows: 6, subs: [
      { no: 1, title: "ตรวจสอบและบันทึกกระแสไฟฟ้า (Amp.)", freq: "ทุกเดือน", color: "#ffff00", type: "amp" },
      { no: 2, title: "เป่าทำความสะอาด Filter", freq: "ทุกเดือน", color: "#ffff00", type: "check", def: "✓" },
      { no: 3, title: "เปลี่ยน Air Filter", freq: "ทุก 3 เดือน", color: "#ed7d31", type: "check", def: "-" }
    ]},
    { id: 2, page: 1, cat: "เครื่องเติมอากาศ", showCat: false, name: "Aeration Blower (AB-SBR-02)", amp: "-", note: "", minRows: 6, subs: [
      { no: 1, title: "ตรวจสอบและบันทึกกระแสไฟฟ้า (Amp.)", freq: "ทุกเดือน", color: "#ffff00", type: "amp" },
      { no: 2, title: "เป่าทำความสะอาด Filter", freq: "ทุกเดือน", color: "#ffff00", type: "check", def: "✓" },
      { no: 3, title: "เปลี่ยน Air Filter", freq: "ทุก 3 เดือน", color: "#ed7d31", type: "check", def: "-" }
    ]},
    { id: 3, page: 1, cat: "เครื่องเติมอากาศ", showCat: false, name: "Aeration Blower (AB-SBR-03)", amp: "41.64", note: "", minRows: 6, subs: [
      { no: 1, title: "ตรวจสอบและบันทึกกระแสไฟฟ้า (Amp.)", freq: "ทุกเดือน", color: "#ffff00", type: "amp" },
      { no: 2, title: "เป่าทำความสะอาด Filter", freq: "ทุกเดือน", color: "#ffff00", type: "check", def: "✓" },
      { no: 3, title: "เปลี่ยน Air Filter", freq: "ทุก 3 เดือน", color: "#ed7d31", type: "check", def: "-" }
    ]},
    { id: 4, page: 2, cat: "เครื่องเติมอากาศผิวน้ำ", showCat: true, name: "Surface Aeration (SA-EQ-01)", amp: "23.56", minRows: 6, subs: [
      { no: 1, title: "ตรวจสอบและบันทึกกระแสไฟฟ้า (Amp.)", freq: "ทุกเดือน", color: "#ffff00", type: "amp" },
      { no: 2, title: "ถ่ายน้ำมันเครื่อง", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "-" },
      { no: 3, title: "ตรวจเช็คใบพัดเติมอากาศ", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "-" },
      { no: 4, title: "ตรวจสอบจุดยึดต่อและความตึงของเชือกใยยัก", freq: "ทุกเดือน", color: "#ffff00", type: "check", def: "✓" }
    ]},
    { id: 5, page: 2, cat: "เครื่องเติมอากาศผิวน้ำ", showCat: false, name: "Surface Aeration (SA-EQ-02)", amp: "23.42", minRows: 6, subs: [
      { no: 1, title: "ตรวจสอบและบันทึกกระแสไฟฟ้า (Amp.)", freq: "ทุกเดือน", color: "#ffff00", type: "amp" },
      { no: 2, title: "ถ่ายน้ำมันเครื่อง", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "-" },
      { no: 3, title: "ตรวจเช็คใบพัดเติมอากาศ", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "-" },
      { no: 4, title: "ตรวจสอบจุดยึดต่อและความตึงของเชือกใยยัก", freq: "ทุกเดือน", color: "#ffff00", type: "check", def: "✓" }
    ]},
    { id: 6, page: 2, cat: "เครื่องสูบน้ำ (Submersible Pump)", showCat: true, name: "Sump Pump (P-CS-01)", amp: "15.65", minRows: 6, subs: subPumpSubs },
    { id: 7, page: 2, cat: "เครื่องสูบน้ำ (Submersible Pump)", showCat: false, name: "Sump Pump (P-CS-02)", amp: "15.14", minRows: 6, subs: subPumpSubs },
    { id: 8, page: 3, cat: "เครื่องสูบน้ำ (Submersible Pump)", showCat: false, name: "Sump Pump (P-CS-03)", amp: "15.54", minRows: 6, subs: subPumpSubs },
    { id: 9, page: 3, cat: "เครื่องสูบน้ำ (Submersible Pump)", showCat: false, name: "Equalizing Pump (P-EQ-01)", amp: "15.31", minRows: 6, subs: subPumpSubs },
    { id: 10, page: 3, cat: "เครื่องสูบน้ำ (Submersible Pump)", showCat: false, name: "Equalizing Pump (P-EQ-02)", amp: "15.22", minRows: 6, subs: subPumpSubs },
    { id: 11, page: 3, cat: "เครื่องสูบน้ำ (Submersible Pump)", showCat: false, name: "Decant Pump (DP-SBR-01)", amp: "17.00", minRows: 6, subs: subPumpSubs },
    { id: 12, page: 4, cat: "เครื่องสูบน้ำ (Submersible Pump)", showCat: false, name: "Decant Pump (DP-SBR-02)", amp: "17.21", minRows: 6, subs: subPumpSubs },
    { id: 13, page: 4, cat: "เครื่องสูบน้ำ (Submersible Pump)", showCat: false, name: "Decant Pump (DP-SBR-03)", amp: "17.19", minRows: 6, subs: subPumpSubs },
    { id: 14, page: 4, cat: "เครื่องสูบน้ำ (Submersible Pump)", showCat: false, name: "Decant Pump (DP-SBR-04)", amp: "17.46", minRows: 6, subs: subPumpSubs },
    { id: 15, page: 4, cat: "เครื่องสูบน้ำ (Submersible Pump)", showCat: false, name: "Excess Sludge Pump (EP-SBR-01)", amp: "2.05", minRows: 6, subs: subPumpSubs },
    { id: 16, page: 5, cat: "เครื่องสูบน้ำ (Submersible Pump)", showCat: false, name: "Excess Sludge Pump (EP-SBR-02)", amp: "1.95", minRows: 6, subs: subPumpSubs },
    { id: 17, page: 5, cat: "เครื่องสูบน้ำ (Submersible Pump)", showCat: false, name: "Excess Sludge Pump (EP-SBR-03)", amp: "2.04", minRows: 6, subs: subPumpSubs },
    { id: 18, page: 5, cat: "เครื่องสูบน้ำ (Submersible Pump)", showCat: false, name: "Excess Sludge Pump (EP-SBR-04)", amp: "2.05", minRows: 6, subs: subPumpSubs },
    { id: 19, page: 5, cat: "เครื่องสูบน้ำ (Submersible Pump)", showCat: false, name: "Effluent Pump (P-EFF-01)", amp: "18.55", minRows: 6, subs: subPumpSubs },
    { id: 20, page: 6, cat: "เครื่องสูบน้ำ (Submersible Pump)", showCat: false, name: "Effluent Pump (P-EFF-02)", amp: "17.30", minRows: 6, subs: subPumpSubs },
    { id: 21, page: 6, cat: "เครื่องสูบน้ำ (Submersible Pump)", showCat: false, name: "Emergency Pump (P-EM-01)", amp: "18.42", minRows: 6, subs: subPumpSubs },
    { id: 22, page: 6, cat: "เครื่องสูบน้ำ (Submersible Pump)", showCat: false, name: "Emergency Pump (P-EM-02)", amp: "17.46", minRows: 6, subs: subPumpSubs },
    { id: 23, page: 6, cat: "เครื่องสูบน้ำ (Submersible Pump)", showCat: false, name: "Sludge Transfer Pump (P-ST-01)", amp: "1.96", minRows: 6, subs: subPumpSubs },
    { id: 24, page: 7, cat: "เครื่องสูบน้ำ (Submersible Pump)", showCat: false, name: "Sludge Transfer Pump (P-ST-02)", amp: "1.93", minRows: 6, subs: subPumpSubs },
    { id: 25, page: 7, cat: "เครื่องสูบน้ำ (Submersible Pump)", showCat: false, name: "Over Flow Sludge pump (P-SS-01)", amp: "1.88", minRows: 6, subs: subPumpSubs },
    { id: 26, page: 7, cat: "เครื่องสูบน้ำ (Submersible Pump)", showCat: false, name: "Over Flow Sludge pump (P-SS-02)", amp: "1.70", minRows: 6, subs: subPumpSubs },
    { id: 27, page: 7, cat: "เครื่องสูบน้ำ (Self Priming Centifugal Pump)", showCat: true, name: "Post Aeration to Wet land Transfer Pump (P-PP-01)", amp: "10.9", minRows: 5, subs: selfPrimingSubs },
    { id: 28, page: 8, cat: "เครื่องสูบน้ำ (Self Priming Centifugal Pump)", showCat: false, name: "Post Aeration to Wet land Transfer Pump (P-PP-02)", amp: "11.96", minRows: 5, subs: selfPrimingSubs },
    { id: 29, page: 8, cat: "เครื่องสูบน้ำ (Self Priming Centifugal Pump)", showCat: false, name: "Retreat Water Pump 2 (P-EP-01)", amp: "6.64", minRows: 5, subs: selfPrimingSubs },
    { id: 30, page: 8, cat: "เครื่องสูบน้ำ (Self Priming Centifugal Pump)", showCat: false, name: "Retreat Water Pump 2 (P-EP-02)", amp: "6.68", minRows: 5, subs: selfPrimingSubs },
    { id: 31, page: 8, cat: "เครื่องสูบจ่ายสารเคมี", showCat: true, name: "H2SO4 Feed Pump (CP-A-01)", amp: "0.92", minRows: 4, subs: chemPumpSubs },
    { id: 32, page: 9, cat: "เครื่องสูบจ่ายสารเคมี", showCat: false, name: "H2SO4 Feed Pump (CP-A-02)", amp: "0.90", minRows: 4, subs: chemPumpSubs },
    { id: 33, page: 9, cat: "เครื่องสูบจ่ายสารเคมี", showCat: false, name: "NaOH Feed Pump (CP-B-01)", amp: "0.29", minRows: 4, subs: chemPumpSubs },
    { id: 34, page: 9, cat: "เครื่องสูบจ่ายสารเคมี", showCat: false, name: "NaOH Feed Pump (CP-B-02)", amp: "0.24", minRows: 4, subs: chemPumpSubs },
    { id: 35, page: 9, cat: "เครื่องสูบจ่ายสารเคมี", showCat: false, name: "Urea Feed Pump (CP-N-01)", amp: "0.51", minRows: 4, subs: chemPumpSubs },
    { id: 36, page: 9, cat: "เครื่องสูบจ่ายสารเคมี", showCat: false, name: "Urea Feed Pump (CP-N-02)", amp: "0.55", minRows: 4, subs: chemPumpSubs },
    { id: 37, page: 10, cat: "เครื่องสูบจ่ายสารเคมี", showCat: false, name: "H3PO4 Feed Pump (CP-P-01)", amp: "0.20", minRows: 4, subs: chemPumpSubs },
    { id: 38, page: 10, cat: "เครื่องสูบจ่ายสารเคมี", showCat: false, name: "H3PO4 Feed Pump (CP-P-02)", amp: "0.20", minRows: 4, subs: chemPumpSubs },
    { id: 39, page: 10, cat: "เครื่องสูบจ่ายสารเคมี", showCat: false, name: "NaOCL feed pump (CP-CL-01)", amp: "0.50", minRows: 4, subs: chemPumpSubs },
    { id: 40, page: 10, cat: "เครื่องสูบจ่ายสารเคมี", showCat: false, name: "NaOCL feed pump (CP-CL-02)", amp: "0.54", minRows: 4, subs: chemPumpSubs },
    { id: 41, page: 10, cat: "เครื่องสูบจ่ายสารเคมี", showCat: false, name: "Polymer Feed pump (P-DS-01)", amp: "0.92", minRows: 4, subs: chemPumpSubs },
    { id: 42, page: 11, cat: "เครื่องสูบจ่ายสารเคมี", showCat: false, name: "Polymer Feed pump (P-DS-02)", amp: "0.89", minRows: 4, subs: chemPumpSubs },
    { id: 43, page: 11, cat: "วาล์ว", showCat: true, name: "Motorized Valve For Inlet (MV-SBR-01B)", minRows: 2, photoGroup: "valve_inlet", subs: valveSubs },
    { id: 44, page: 11, cat: "วาล์ว", showCat: false, name: "Motorized Valve For Inlet (MV-SBR-02B)", minRows: 2, photoGroup: "valve_inlet", subs: valveSubs },
    { id: 45, page: 11, cat: "วาล์ว", showCat: false, name: "Motorized Valve For Inlet (MV-SBR-03B)", minRows: 2, photoGroup: "valve_inlet", subs: valveSubs },
    { id: 46, page: 11, cat: "วาล์ว", showCat: false, name: "Motorized Valve For Inlet (MV-SBR-04B)", minRows: 2, photoGroup: "valve_inlet", subs: valveSubs },
    { id: 47, page: 11, cat: "วาล์ว", showCat: false, name: "Motorized Valve For Air (MV-SBR-01A)", minRows: 2, photoGroup: "valve_air", subs: valveSubs },
    { id: 48, page: 11, cat: "วาล์ว", showCat: false, name: "Motorized Valve For Air (MV-SBR-02A)", minRows: 2, photoGroup: "valve_air", subs: valveSubs },
    { id: 49, page: 11, cat: "วาล์ว", showCat: false, name: "Motorized Valve For Air (MV-SBR-03A)", minRows: 2, photoGroup: "valve_air", subs: valveSubs },
    { id: 50, page: 11, cat: "วาล์ว", showCat: false, name: "Motorized Valve For Air (MV-SBR-04A)", minRows: 2, photoGroup: "valve_air", subs: valveSubs },
    { id: 51, page: 12, cat: "เครื่องมือวัดความเป็นกรด-ด่าง และออกซิเจนละลายน้ำ", showCat: true, name: "pH Meter (pH-EQ-01)", calRange: "Slop 80 - 103 %", calVal: "95.01", minRows: 4, subs: meterSubs },
    { id: 52, page: 12, cat: "เครื่องมือวัดความเป็นกรด-ด่าง และออกซิเจนละลายน้ำ", showCat: false, name: "DO meter (DO-SBR-01)", calRange: "Slop (-30) - (-110) %", calVal: "- 86.83", minRows: 4, subs: meterSubs },
    { id: 53, page: 12, cat: "เครื่องมือวัดความเป็นกรด-ด่าง และออกซิเจนละลายน้ำ", showCat: false, name: "DO meter (DO-SBR-02)", calRange: "Slop (-30) - (-110) %", calVal: "-93.71", minRows: 4, subs: meterSubs },
    { id: 54, page: 12, cat: "เครื่องมือวัดความเป็นกรด-ด่าง และออกซิเจนละลายน้ำ", showCat: false, name: "DO meter (DO-SBR-03)", calRange: "Slop (-30) - (-110) %", calVal: "-89.68", minRows: 4, subs: meterSubs },
    { id: 55, page: 12, cat: "เครื่องมือวัดความเป็นกรด-ด่าง และออกซิเจนละลายน้ำ", showCat: false, name: "DO meter (DO-SBR-04)", calRange: "Slop (-30) - (-110) %", calVal: "-74.04", note: "หัวที่ 4 ต้องทำการ Cal 3-4 ครั้งค่าถึงจะขึ้น", minRows: 4, subs: meterSubs },
    { id: 56, page: 13, cat: "เครื่องวัดอัตราการไหล", showCat: true, name: "Flow Meter (FM-CS-01)", minRows: 2, photoGroup: "flow_meter", subs: [{ no: 1, title: "Calibration (Cer)", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "-" }] },
    { id: 57, page: 13, cat: "เครื่องวัดอัตราการไหล", showCat: false, name: "Flow Meter (FM-EQ-01)", minRows: 2, photoGroup: "flow_meter", subs: [{ no: 1, title: "Calibration (Cer)", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "-" }] },
    { id: 58, page: 13, cat: "ชุดดักขยะและกรวดทราย", showCat: true, name: "Grit Remover Unit (GR-01)", amp: "1.81", minRows: 2, photoGroup: "grit_remover", subs: [
      { no: 1, title: "ตรวจสอบและบันทึกกระแสไฟฟ้า (Amp.)", freq: "ทุกเดือน", color: "#ffff00", type: "amp" },
      { no: 2, title: "เปลี่ยนน้ำมันเกียร์", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "-" }
    ]},
    { id: 59, page: 13, cat: "ชุดดักขยะและกรวดทราย", showCat: false, name: "Grit Remover Unit (GR-02)", amp: "1.88", minRows: 2, photoGroup: "grit_remover", subs: [
      { no: 1, title: "ตรวจสอบและบันทึกกระแสไฟฟ้า (Amp.)", freq: "ทุกเดือน", color: "#ffff00", type: "amp" },
      { no: 2, title: "เปลี่ยนน้ำมันเกียร์", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "-" }
    ]},
    { id: 60, page: 13, cat: "ชุดดักขยะและกรวดทราย", showCat: false, name: "Grit Remover Unit (GR-03)", amp: "1.85", minRows: 2, photoGroup: "grit_remover", subs: [
      { no: 1, title: "ตรวจสอบและบันทึกกระแสไฟฟ้า (Amp.)", freq: "ทุกเดือน", color: "#ffff00", type: "amp" },
      { no: 2, title: "เปลี่ยนน้ำมันเกียร์", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "-" }
    ]},
    { id: 61, page: 13, cat: "หัวจ่ายอากาศ", showCat: true, name: "Fine Bubble Diffuser", minRows: 4, subs: [
      { no: 1, title: "ตรวจสอบและบันทึกค่าความดัน (bar)", freq: "ทุกเดือน", color: "#ffff00", type: "check", def: "✓" },
      { no: 2, title: "ล้างด้วยกรดอ่อน", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "-" }
    ]},
    { id: 62, page: 13, cat: "ลูกลอย", showCat: true, name: "Level switch (LS-EFF-01,02)", minRows: 3, subs: levelSwitchSubs },
    { id: 63, page: 13, cat: "ลูกลอย", showCat: false, name: "Level switch (LS-EM-01,02)", minRows: 3, subs: levelSwitchSubs },
    { id: 64, page: 14, cat: "ลูกลอย", showCat: false, name: "Level switch (LS-ST-01,02)", minRows: 3, subs: levelSwitchSubs },
    { id: 65, page: 14, cat: "ลูกลอย", showCat: false, name: "Level switch (LS-SS-01,02,03)", minRows: 3, subs: levelSwitchSubs },
    { id: 66, page: 14, cat: "ลูกลอย", showCat: false, name: "Level switch (LS-PP-01,02,03)", minRows: 3, subs: levelSwitchSubs },
    { id: 67, page: 14, cat: "ลูกลอย", showCat: false, name: "Level switch (LS-EP-01,02,03)", minRows: 3, subs: levelSwitchSubs },
    { id: 68, page: 14, cat: "Radar", showCat: true, name: "Level Transmitter (LT-CS-01)", minRows: 3, photoGroup: "radar_1", subs: radarSubs },
    { id: 69, page: 14, cat: "Radar", showCat: false, name: "Level Transmitter (LT-EQ-01)", minRows: 3, photoGroup: "radar_1", subs: radarSubs },
    { id: 70, page: 14, cat: "Radar", showCat: false, name: "Level Transmitter (LT-SBR-01,02,03,04)", minRows: 3, photoGroup: "radar_2", subs: radarSubs },
    { id: 71, page: 14, cat: "Radar", showCat: false, name: "Level Transmitter (LT-PP-01)", minRows: 3, photoGroup: "radar_2", subs: radarSubs },
    { id: 72, page: 15, cat: "ระบบรีดตะกอน", showCat: true, name: "Dewatering Machine (DM-01)", minRows: 5, subs: [
      { no: 1, title: "ตรวจสอบการปรับระยะ End Plate", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "-" },
      { no: 2, title: "อัดจาระบี", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "-" },
      { no: 3, title: "เปลี่ยน Moving ring", freq: "5,000 ชั่วโมง", color: "", type: "check", def: "-" },
      { no: 4, title: "เปลี่ยน Screw shaft", freq: "10,000 ชั่วโมง", color: "", type: "check", def: "-" },
      { no: 5, title: "เปลี่ยนถ่ายน้ำมันหล่อลื่น", freq: "2 ปี", color: "", type: "check", def: "-" }
    ]},
    { id: 73, page: 15, cat: "ระบบรีดตะกอน", showCat: false, name: "Agitator (MX-01)", minRows: 4, subs: [
      { no: 1, title: "เปลี่ยนถ่ายน้ำมันเกียร์", freq: "2 ปี", color: "", type: "check", def: "-" }
    ]},
    { id: 74, page: 15, cat: "ระบบรีดตะกอน", showCat: false, name: "Washing Pump (WP-01)", minRows: 3, photoGroup: "washing_pump", subs: [
      { no: 1, title: "เช็ครอยรั่วตามวาล์วแต่ละจุด", freq: "ทุก 3 เดือน", color: "#ed7d31", type: "check", def: "-" },
      { no: 2, title: "เช็คลมในถังแรงดัน", freq: "ทุก 3 เดือน", color: "#ed7d31", type: "check", def: "-" },
      { no: 3, title: "เปลี่ยนถุง Diaphragm", freq: "2-3 ปี", color: "", type: "check", def: "-" }
    ]},
    { id: 75, page: 15, cat: "ระบบรีดตะกอน", showCat: false, name: "Washing Pump (WP-02)", minRows: 3, photoGroup: "washing_pump", subs: [
      { no: 1, title: "เช็ครอยรั่วตามวาล์วแต่ละจุด", freq: "ทุก 3 เดือน", color: "#ed7d31", type: "check", def: "-" },
      { no: 2, title: "เช็คลมในถังแรงดัน", freq: "ทุก 3 เดือน", color: "#ed7d31", type: "check", def: "-" },
      { no: 3, title: "เปลี่ยนถุง Diaphragm", freq: "2-3 ปี", color: "", type: "check", def: "-" }
    ]},
    { id: 76, page: 15, cat: "ตู้ควบคุมไฟฟ้า", showCat: true, name: "MCC Panel", minRows: 5, subs: [
      { no: 1, title: "ทำความสะอาดตู้ควบคุมไฟฟ้า", freq: "ทุกเดือน", color: "#ffff00", type: "check", def: "✓" },
      { no: 2, title: "ตรวจสอบจุดขันแน่นที่เทอร์มินอล", freq: "ทุก 1 ปี", color: "#70ad47", type: "check", def: "-" }
    ]}
  ];
}

/**
 * สร้างหรือจัดรูปแบบตาราง 2_Report_Form ใหม่ทั้งหมดให้เหมือนไฟล์ตัวอย่างเป๊ะ (โดยไม่ลบข้อมูลที่กรอกไว้แล้วใน 1_Raw_Data)
 */
function setupPMGoogleSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = getMasterEquipmentList();

  // 1. ตรวจสอบ 1_Raw_Data (ถ้ามีข้อมูลเดิมอยู่แล้ว ให้รักษาไว้และแปลงลิงก์รูปเป็น thumbnail เพื่อให้แสดงผลได้ทันที)
  let rawSheet = ss.getSheetByName("1_Raw_Data");
  if (!rawSheet) {
    rawSheet = ss.insertSheet("1_Raw_Data");
    const rawHeaders = [
      "Timestamp", "Report_Round", "Equip_No", "Category", "Equip_Name",
      "Action_Date (วันที่ดำเนินการ)", "Amp_Value (กระแสไฟ A)",
      "Check_1", "Check_2", "Check_3", "Check_4", "Check_5",
      "Cal_Percent (ค่า % Slope)", "Note (หมายเหตุ)",
      "Photo_1_URL", "Photo_2_URL", "Photo_3_URL", "Photo_4_URL", "Photo_5_URL"
    ];
    rawSheet.getRange(1, 1, 1, rawHeaders.length).setValues([rawHeaders])
      .setBackground("#0f4c81").setFontColor("#ffffff").setFontWeight("bold");

    const rawRows = master.map(item => {
      const checkSubs = item.subs.filter(s => s.type !== "amp");
      return [
        new Date(), "2/69", item.id, item.cat, item.name,
        "21/07/2026", item.amp !== undefined ? item.amp : "",
        checkSubs[0] ? checkSubs[0].def : "",
        checkSubs[1] ? checkSubs[1].def : "",
        checkSubs[2] ? checkSubs[2].def : "",
        checkSubs[3] ? checkSubs[3].def : "",
        checkSubs[4] ? checkSubs[4].def : "",
        item.calVal || "", item.note || "",
        "", "", "", "", ""
      ];
    });
    rawSheet.getRange(2, 1, rawRows.length, rawHeaders.length).setValues(rawRows);
    rawSheet.setFrozenRows(1);
  } else {
    // แปลง URL รูปภาพในคอลัมน์ O..S (index 14..18) ของ 1_Raw_Data ให้เป็น thumbnail URL ทั้งหมดเพื่อให้รูปขึ้นทันที
    const vals = rawSheet.getDataRange().getValues();
    let changed = false;
    for (let r = 1; r < vals.length; r++) {
      for (let c = 14; c <= 18; c++) {
        if (vals[r][c] && String(vals[r][c]).includes("drive.google.com/uc")) {
          vals[r][c] = toSheetImageUrl(vals[r][c]);
          changed = true;
        }
      }
    }
    if (changed) {
      rawSheet.getRange(1, 1, vals.length, vals[0].length).setValues(vals);
    }
  }

  // 2. สร้างและจัดรูปแบบ 2_Report_Form แบบ Batch (รวดเร็วและตรงตามแบบฟอร์มตัวอย่าง 100%)
  let reportSheet = ss.getSheetByName("2_Report_Form");
  if (!reportSheet) reportSheet = ss.insertSheet("2_Report_Form");
  reportSheet.clear();

  // ตั้งความกว้างคอลัมน์ A ถึง K ให้ได้สัดส่วนเดียวกับแบบฟอร์ม A4 แนวนอน
  reportSheet.setColumnWidth(1, 55);   // A: ลำดับที่ / Note
  reportSheet.setColumnWidth(2, 32);   // B: ข้อย่อย 1, 2, 3
  reportSheet.setColumnWidth(3, 235);  // C: รายการตรวจสอบ
  reportSheet.setColumnWidth(4, 95);   // D: ความถี่ / วันที่ดำเนินการ
  reportSheet.setColumnWidth(5, 48);   // E: Plan (ช่องสี)
  reportSheet.setColumnWidth(6, 95);   // F: Actual
  for (let c = 7; c <= 11; c++) {
    reportSheet.setColumnWidth(c, 138); // G-K: ภาพประกอบ 5 ช่อง
  }

  // เตรียม Array 2 มิติสำหรับใส่ค่า/สูตร พื้นหลัง สีตัวอักษร น้ำหนักตัวอักษร และการจัดวาง
  const gridVals = [];
  const gridBgs = [];
  const gridFontColors = [];
  const gridFontWeights = [];
  const gridHAligns = [];
  const mergeTasks = []; // { r, c, numR, numC }
  const borderBlocks = []; // { r, c, numR, numC }

  function pushRow(vals, bgs, fColors, fWeights, hAligns) {
    gridVals.push(vals);
    gridBgs.push(bgs || Array(11).fill(null));
    gridFontColors.push(fColors || Array(11).fill("#000000"));
    gridFontWeights.push(fWeights || Array(11).fill("normal"));
    gridHAligns.push(hAligns || Array(11).fill("left"));
    return gridVals.length; // 1-based row number
  }

  // แถวที่ 1: ชื่อรายงาน
  pushRow(
    ['="รายงานการบำรุงรักษาเครื่องจักรและเครื่องมือวัดระบบบำบัดน้ำเสีย ครั้ง  " & \'1_Raw_Data\'!B2', "", "", "", "", "", "", "", "", "Universal Utilities", ""],
    null,
    ["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#1e3a5f","#1e3a5f"],
    ["bold","normal","normal","normal","normal","normal","normal","normal","normal","bold","bold"],
    ["center","center","center","center","center","center","center","center","center","center","center"]
  );
  mergeTasks.push({ r: 1, c: 1, numR: 1, numC: 9 });
  mergeTasks.push({ r: 1, c: 10, numR: 1, numC: 2 });

  // แถวที่ 2: ชื่อโครงการ
  pushRow(
    ["โครงการ ระบบบำบัดน้ำเสีย นิคมอุตสาหกรรมหลักชัยเมืองยาง ขนาด 3,750 ลูกบาศก์เมตรต่อวัน", "", "", "", "", "", "", "", "", "GROUP", ""],
    null,
    ["#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#475569","#475569"],
    ["bold","normal","normal","normal","normal","normal","normal","normal","normal","bold","bold"],
    ["center","center","center","center","center","center","center","center","center","center","center"]
  );
  mergeTasks.push({ r: 2, c: 1, numR: 1, numC: 9 });
  mergeTasks.push({ r: 2, c: 10, numR: 1, numC: 2 });

  // แถวที่ 3: เว้นบรรทัดว่าง
  pushRow(Array(11).fill(""));

  // แถวที่ 4: คำอธิบายสีความถี่ (Legend)
  pushRow(
    ["", "", "", "ทุกเดือน", "", "ทุก 3 เดือน", "", "ทุก 6 เดือน", "", "ทุก 1 ปี", ""],
    [null, null, "#ffff00", null, "#ed7d31", null, "#7030a0", null, "#70ad47", null, null],
    null,
    Array(11).fill("bold"),
    ["center","center","center","left","center","left","center","left","center","left","center"]
  );
  borderBlocks.push({ r: 4, c: 3, numR: 1, numC: 1 });
  borderBlocks.push({ r: 4, c: 5, numR: 1, numC: 1 });
  borderBlocks.push({ r: 4, c: 7, numR: 1, numC: 1 });
  borderBlocks.push({ r: 4, c: 9, numR: 1, numC: 1 });

  // แถวที่ 5: เว้นบรรทัดว่าง
  pushRow(Array(11).fill(""));

  // แถวที่ 6: หัวตารางหลัก
  pushRow(
    ["ลำดับที่", "รายการตรวจสอบ", "", "ความถี่", "Plan", "Actual", "ภาพประกอบ", "", "", "", ""],
    Array(11).fill("#a6a6a6"),
    Array(11).fill("#000000"),
    Array(11).fill("bold"),
    Array(11).fill("center")
  );
  mergeTasks.push({ r: 6, c: 2, numR: 1, numC: 2 });
  mergeTasks.push({ r: 6, c: 7, numR: 1, numC: 5 });
  borderBlocks.push({ r: 6, c: 1, numR: 1, numC: 11 });

  const checkCols = ["H", "I", "J", "K", "L"];
  const photoCols = ["O", "P", "Q", "R", "S"];

  let idx = 0;
  while (idx < master.length) {
    const item = master[idx];

    // กรณีเป็นกลุ่มที่ใช้ช่องรูปภาพร่วมกัน (เช่น วาล์ว 43-46, 47-50, Flow 56-57, Grit 58-60, Radar 68-69, 70-71, Washing 74-75)
    if (item.photoGroup) {
      const groupName = item.photoGroup;
      const groupItems = [];
      while (idx < master.length && master[idx].photoGroup === groupName) {
        groupItems.push(master[idx]);
        idx++;
      }

      if (groupItems[0].showCat) {
        const catRow = pushRow(
          [groupItems[0].cat, "", "", "", "", "", "", "", "", "", ""],
          Array(11).fill("#d9d9d9"),
          null,
          Array(11).fill("bold"),
          ["left","left","left","left","left","left","left","left","left","left","left"]
        );
        mergeTasks.push({ r: catRow, c: 1, numR: 1, numC: 11 });
        borderBlocks.push({ r: catRow, c: 1, numR: 1, numC: 11 });
      }

      const groupBlockStart = gridVals.length + 1;
      const firstRawRow = groupItems[0].id + 1; // แถวใน 1_Raw_Data ของเครื่องแรกในกลุ่ม

      groupItems.forEach((g, gIdx) => {
        const r = g.id + 1;
        // แถวชื่อเครื่องจักร
        const titleRow = pushRow(
          [g.id, g.name, "", "วันที่ดำเนินการ :", "", `='1_Raw_Data'!F${r}`, "", "", "", "", ""],
          null,
          null,
          ["bold","bold","bold","bold","normal","normal","normal","normal","normal","normal","normal"],
          ["center","left","left","center","center","center","center","center","center","center","center"]
        );
        mergeTasks.push({ r: titleRow, c: 2, numR: 1, numC: 2 });
        if (gIdx === 0) {
          mergeTasks.push({ r: titleRow, c: 7, numR: 1, numC: 5 });
        }

        const totalSubRows = Math.max(g.minRows || 2, g.subs.length);
        let cIdx = 0;
        for (let s = 0; s < totalSubRows; s++) {
          const sub = g.subs[s];
          let subNo = "", subTitle = "", subFreq = "", planBg = null, actualFormula = "";
          if (sub) {
            subNo = sub.no;
            subTitle = sub.title;
            subFreq = sub.freq;
            planBg = sub.color || null;
            if (sub.type === "amp") {
              actualFormula = `=IF('1_Raw_Data'!G${r}="", "", IF('1_Raw_Data'!G${r}="-", "-", '1_Raw_Data'!G${r} & " A"))`;
            } else {
              const colLet = checkCols[cIdx++];
              actualFormula = `='1_Raw_Data'!${colLet}${r}`;
            }
          }

          // วางสูตร =IMAGE() ที่แถวย่อยแรกของเครื่องแรกในกลุ่ม
          const pFormulas = ["", "", "", "", ""];
          if (gIdx === 0 && s === 0) {
            for (let p = 0; p < 5; p++) {
              pFormulas[p] = `=IF('1_Raw_Data'!${photoCols[p]}${firstRawRow}<>"", IMAGE('1_Raw_Data'!${photoCols[p]}${firstRawRow}, 1), "")`;
            }
          }

          pushRow(
            ["", subNo, subTitle, subFreq, "", actualFormula, pFormulas[0], pFormulas[1], pFormulas[2], pFormulas[3], pFormulas[4]],
            [null, null, null, null, planBg, null, null, null, null, null, null],
            null,
            ["normal","normal","normal","normal","normal","bold","normal","normal","normal","normal","normal"],
            ["center","center","left","center","center","center","center","center","center","center","center"]
          );
        }
      });

      const groupBlockEnd = gridVals.length;
      const photoStartRow = groupBlockStart + 1;
      const photoSpanRows = groupBlockEnd - photoStartRow + 1;
      for (let p = 0; p < 5; p++) {
        if (photoSpanRows > 1) {
          mergeTasks.push({ r: photoStartRow, c: 7 + p, numR: photoSpanRows, numC: 1 });
        }
      }
      borderBlocks.push({ r: groupBlockStart, c: 1, numR: groupBlockEnd - groupBlockStart + 1, numC: 11 });
      continue;
    }

    // กรณีเครื่องจักรปกติ (มีกรอบตาราง 6 แถวย่อย + แถว Note เหมือนตัวอย่างเป๊ะ)
    const r = item.id + 1; // แถวใน 1_Raw_Data (ลำดับ 1 อยู่แถว 2)
    if (item.showCat) {
      const catRow = pushRow(
        [item.cat, "", "", "", "", "", "", "", "", "", ""],
        Array(11).fill("#d9d9d9"),
        null,
        Array(11).fill("bold"),
        ["left","left","left","left","left","left","left","left","left","left","left"]
      );
      mergeTasks.push({ r: catRow, c: 1, numR: 1, numC: 11 });
      borderBlocks.push({ r: catRow, c: 1, numR: 1, numC: 11 });
    }

    const blockStart = gridVals.length + 1;

    // แถวหัวชื่อเครื่องจักร + วันที่ดำเนินการ + แถบยาวเหนือรูป (G:K)
    const titleRow = pushRow(
      [item.id, item.name, "", "วันที่ดำเนินการ :", "", `='1_Raw_Data'!F${r}`, "", "", "", "", ""],
      null,
      null,
      ["bold","bold","bold","bold","normal","normal","normal","normal","normal","normal","normal"],
      ["center","left","left","center","center","center","center","center","center","center","center"]
    );
    mergeTasks.push({ r: titleRow, c: 2, numR: 1, numC: 2 });
    mergeTasks.push({ r: titleRow, c: 7, numR: 1, numC: 5 });

    const extraRows = item.calRange ? 2 : Math.max(0, (item.minRows || 6) - item.subs.length);
    const totalSubRows = item.subs.length + extraRows;
    let cIdx = 0;

    for (let s = 0; s < totalSubRows; s++) {
      const sub = item.subs[s];
      let subNo = "", subTitle = "", subFreq = "", planBg = null, planText = "", actualFormula = "";
      let fColors = Array(11).fill("#000000");

      if (sub) {
        subNo = sub.no;
        subTitle = sub.title;
        subFreq = sub.freq;
        planBg = sub.color || null;
        if (sub.type === "amp") {
          actualFormula = `=IF('1_Raw_Data'!G${r}="", "", IF('1_Raw_Data'!G${r}="-", "-", '1_Raw_Data'!G${r} & " A"))`;
        } else {
          const colLet = checkCols[cIdx++];
          actualFormula = `='1_Raw_Data'!${colLet}${r}`;
        }
      } else if (item.calRange && s === item.subs.length) {
        planText = item.calRange;
      } else if (item.calRange && s === item.subs.length + 1) {
        planText = `="อยู่ในช่วง = " & '1_Raw_Data'!M${r} & " %"`;
        fColors[4] = "#2e75b6";
        fColors[5] = "#2e75b6";
      }

      const pFormulas = ["", "", "", "", ""];
      if (s === 0) {
        for (let p = 0; p < 5; p++) {
          pFormulas[p] = `=IF('1_Raw_Data'!${photoCols[p]}${r}<>"", IMAGE('1_Raw_Data'!${photoCols[p]}${r}, 1), "")`;
        }
      }

      const subRowNum = pushRow(
        ["", subNo, subTitle, subFreq, planText, actualFormula, pFormulas[0], pFormulas[1], pFormulas[2], pFormulas[3], pFormulas[4]],
        [null, null, null, null, planBg, null, null, null, null, null, null],
        fColors,
        ["normal","normal","normal","normal","normal","bold","normal","normal","normal","normal","normal"],
        ["center","center","left","center","center","center","center","center","center","center","center"]
      );

      if (item.calRange && (s === item.subs.length || s === item.subs.length + 1)) {
        mergeTasks.push({ r: subRowNum, c: 5, numR: 1, numC: 2 });
      }
    }

    const photoStartRow = blockStart + 1;
    for (let p = 0; p < 5; p++) {
      if (totalSubRows > 1) {
        mergeTasks.push({ r: photoStartRow, c: 7 + p, numR: totalSubRows, numC: 1 });
      }
    }

    // แถว Note ท้ายบล็อกเครื่องจักร (แสดงคำว่า Note และข้อความหมายเหตุสีเขียวเมื่อมีการกรอก Note ในเครื่องนั้น)
    if (item.note !== undefined || item.id <= 3 || item.id === 55) {
      const noteRowNum = pushRow(
        [`=IF('1_Raw_Data'!N${r}<>"", "Note", "")`, `='1_Raw_Data'!N${r}`, "", "", "", "", "", "", "", "", ""],
        null,
        ["#548235","#548235","#548235","#000000","#000000","#000000","#000000","#000000","#000000","#000000","#000000"],
        null,
        ["left","left","left","center","center","center","center","center","center","center","center"]
      );
      mergeTasks.push({ r: noteRowNum, c: 2, numR: 1, numC: 2 });
    }

    const blockEnd = gridVals.length;
    borderBlocks.push({ r: blockStart, c: 1, numR: blockEnd - blockStart + 1, numC: 11 });
    idx++;
  }

  // เขียนข้อมูลและรูปแบบทั้งหมดลง Sheet ในคำสั่งเดียว (รวดเร็วและแม่นยำ 100%)
  const totalRows = gridVals.length;
  const fullRange = reportSheet.getRange(1, 1, totalRows, 11);
  fullRange.setValues(gridVals);
  fullRange.setBackgrounds(gridBgs);
  fullRange.setFontColors(gridFontColors);
  fullRange.setFontWeights(gridFontWeights);
  fullRange.setHorizontalAlignments(gridHAligns);
  fullRange.setVerticalAlignment("middle"); // จัดกึ่งกลางแนวตั้งทุกช่อง (รวมถึงรูปภาพ!)
  fullRange.setFontFamily("Sarabun");
  fullRange.setFontSize(10);

  // ปรับขนาดตัวอักษรหัวกระดาษ
  reportSheet.getRange("A1:K2").setFontSize(12);
  reportSheet.setRowHeights(6, totalRows - 5, 24); // ตั้งความสูงแถวให้โปร่งสวยงามเหมือนไฟล์ตัวอย่าง

  // ผสานเซลล์ทั้งหมด
  mergeTasks.forEach(m => {
    reportSheet.getRange(m.r, m.c, m.numR, m.numC).merge();
  });

  // ตีเส้นตารางทั้งหมด
  borderBlocks.forEach(b => {
    reportSheet.getRange(b.r, b.c, b.numR, b.numC).setBorder(true, true, true, true, true, true);
  });

  SpreadsheetApp.flush();
}

/**
 * สร้างหน้าเว็บแอพมือถือแบบถาวรบนเซิร์ฟเวอร์ Google (เปิดผ่านลิงก์ /exec ได้ตลอด 24 ชม. แม้ปิดคอมพิวเตอร์)
 */
function getMobileAppHtml() {
  const itemsJson = JSON.stringify(getMasterEquipmentList());
  return `<!DOCTYPE html>
<html lang="th">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>บันทึกงาน PM มือถือ (Universal Utilities GROUP)</title>
  <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;600;700&display=swap" rel="stylesheet">
  <style>
    * { box-sizing: border-box; }
    body { margin: 0; font-family: 'Sarabun', sans-serif; background: #f1f5f9; color: #1e293b; }
    .wrap { max-width: 520px; margin: 0 auto; background: #fff; min-height: 100vh; box-shadow: 0 4px 20px rgba(0,0,0,0.08); }
    .hdr { background: #0f4c81; color: #fff; padding: 14px 16px; position: sticky; top: 0; z-index: 10; }
    .hdr h2 { margin: 0 0 4px 0; font-size: 16px; }
    .hdr p { margin: 0; font-size: 12px; opacity: 0.9; }
    .body { padding: 16px; }
    .lbl { display: block; font-size: 13px; font-weight: 700; margin-bottom: 4px; color: #334155; }
    .ctrl { width: 100%; padding: 9px 10px; border: 1.5px solid #cbd5e1; border-radius: 8px; font-family: 'Sarabun', sans-serif; font-size: 15px; }
    .card { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px; margin-bottom: 8px; }
    .tgroup { display: grid; grid-template-columns: repeat(3, 1fr); gap: 6px; margin-top: 6px; }
    .tbtn { padding: 8px; border: 1.5px solid #cbd5e1; background: #fff; border-radius: 6px; font-family: 'Sarabun', sans-serif; font-size: 13px; font-weight: 600; cursor: pointer; }
    .tbtn.ok { background: #dcfce7; border-color: #16a34a; color: #15803d; }
    .tbtn.dash { background: #fef3c7; border-color: #d97706; color: #b45309; }
    .pgrid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; margin-top: 8px; }
    .pslot { aspect-ratio: 4/3; border: 1.5px dashed #94a3b8; border-radius: 8px; display: flex; flex-direction: column; align-items: center; justify-content: center; cursor: pointer; position: relative; overflow: hidden; background: #f8fafc; font-size: 11px; color: #64748b; }
    .pslot img { width: 100%; height: 100%; object-fit: cover; }
    .btn-save { width: 100%; padding: 13px; background: #10b981; color: #fff; border: none; border-radius: 10px; font-family: 'Sarabun', sans-serif; font-size: 15px; font-weight: 700; cursor: pointer; margin-top: 10px; }
    .toast { display: none; background: #059669; color: #fff; padding: 10px; border-radius: 8px; margin-bottom: 10px; text-align: center; font-weight: 600; font-size: 13px; }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="hdr">
      <div style="display:flex;justify-content:space-between;align-items:center;">
        <h2>📋 บันทึกผลตรวจ PM (ครบ 76 เครื่อง)</h2>
        <span id="cnt" style="background:rgba(255,255,255,0.2);padding:2px 8px;border-radius:10px;font-size:12px;">0/76</span>
      </div>
      <p>ลิงก์ถาวรบน Google Cloud เชื่อมต่อ Google Sheet อัตโนมัติ</p>
    </div>
    <div class="body">
      <div id="toast" class="toast"></div>
      <div style="display:grid;grid-template-columns:95px 1fr;gap:8px;background:#f0f7ff;padding:10px;border-radius:10px;margin-bottom:12px;">
        <div>
          <label class="lbl">ครั้งที่</label>
          <input id="round" class="ctrl" value="2/69">
        </div>
        <div>
          <label class="lbl">วันที่ดำเนินการ (จำค่าต่อให้เครื่องถัดไป)</label>
          <input id="actDate" class="ctrl" value="21/07/2026" style="font-weight:700;color:#0f4c81;">
        </div>
      </div>
      <div style="margin-bottom:10px;">
        <label class="lbl">เลือกเครื่องจักร (1–76):</label>
        <select id="eqSel" class="ctrl" style="font-weight:700;color:#0f4c81;" onchange="selEq(Number(this.value))"></select>
      </div>
      <div id="fields"></div>
      <div style="margin-top:12px;">
        <label class="lbl">📸 อัพรูปประกอบ (สูงสุด 5 รูป):</label>
        <div id="pgrid" class="pgrid"></div>
        <input type="file" id="fInput" accept="image/*" multiple style="display:none" onchange="onFiles(this.files)">
      </div>
      <button class="btn-save" onclick="saveAndNext()">💾 บันทึกลง Google Sheet & ไปเครื่องถัดไป ➔</button>
    </div>
  </div>
  <script>
    const ITEMS = ${itemsJson}.map(x => ({
      ...x,
      actionDate: "21/07/2026",
      isSaved: false,
      ampValue: x.amp || "",
      calPercent: x.calVal || "",
      noteVal: x.note || "",
      subVals: x.subs.map(s => s.def || ""),
      photos: []
    }));
    let curIdx = 0;
    let lastDate = "21/07/2026";
    let slotIdx = 0;

    function render() {
      const cur = ITEMS[curIdx];
      document.getElementById("cnt").textContent = "บันทึกแล้ว " + ITEMS.filter(x => x.isSaved).length + "/76";
      document.getElementById("actDate").value = cur.actionDate;
      document.getElementById("eqSel").innerHTML = ITEMS.map((it, i) =>
        \`<option value="\${i}" \${i===curIdx?'selected':''}>\${it.isSaved?'🔒':'⏳'} #\${it.id} \${it.name} (หน้า \${it.page})</option>\`
      ).join("");

      let h = "";
      cur.subs.forEach((s, idx) => {
        if (s.type === "amp") {
          h += \`<div class="card"><label class="lbl">⚡ 1. \${s.title} (\${s.freq})</label>
                <input id="ampIn" class="ctrl" value="\${cur.ampValue}" placeholder="กรอกตัวเลขแอมป์ หรือ -"></div>\`;
        } else {
          const v = cur.subVals[idx];
          h += \`<div class="card"><div style="font-size:13px;font-weight:600;">\${s.no}. \${s.title} <span style="font-size:11px;color:#64748b;">(\${s.freq})</span></div>
                <div class="tgroup">
                  <button type="button" class="tbtn \${v==='✓'?'ok':''}" onclick="setChk(\${idx},'✓')">✓ ทำแล้ว</button>
                  <button type="button" class="tbtn \${v==='-'?'dash':''}" onclick="setChk(\${idx},'-')">- ไม่ทำ</button>
                  <button type="button" class="tbtn" onclick="setChk(\${idx},'')">เว้นว่าง</button>
                </div></div>\`;
        }
      });
      if (cur.calRange) {
        h += \`<div class="card"><label class="lbl">🌡️ ค่า % Slope (\${cur.calRange})</label>
              <input id="calIn" class="ctrl" value="\${cur.calPercent}"></div>\`;
      }
      h += \`<div class="card"><label class="lbl">📝 หมายเหตุ (Note)</label>
            <input id="noteIn" class="ctrl" value="\${cur.noteVal}" placeholder="เว้นว่างได้"></div>\`;
      document.getElementById("fields").innerHTML = h;

      let ph = "";
      for (let i = 0; i < 5; i++) {
        ph += \`<div class="pslot" onclick="slotIdx=\${i};document.getElementById('fInput').click()">
          \${cur.photos[i] ? \`<img src="\${cur.photos[i]}">\` : \`📷 รูปที่ \${i+1}\`}
        </div>\`;
      }
      document.getElementById("pgrid").innerHTML = ph;
    }

    function selEq(i) {
      curIdx = i;
      if (!ITEMS[curIdx].isSaved) ITEMS[curIdx].actionDate = lastDate;
      render();
    }
    function setChk(idx, val) {
      ITEMS[curIdx].subVals[idx] = val;
      render();
    }
    function onFiles(files) {
      Array.from(files).forEach((f, i) => {
        const r = new FileReader();
        r.onload = e => {
          const img = new Image();
          img.onload = () => {
            const c = document.createElement("canvas");
            const sc = Math.min(1, 550 / img.width);
            c.width = img.width * sc; c.height = img.height * sc;
            c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
            ITEMS[curIdx].photos[Math.min(4, slotIdx + i)] = c.toDataURL("image/jpeg", 0.84);
            render();
          };
          img.src = e.target.result;
        };
        r.readAsDataURL(f);
      });
    }
    function saveAndNext() {
      const cur = ITEMS[curIdx];
      cur.actionDate = document.getElementById("actDate").value.trim();
      if (document.getElementById("ampIn")) cur.ampValue = document.getElementById("ampIn").value.trim();
      if (document.getElementById("calIn")) cur.calPercent = document.getElementById("calIn").value.trim();
      if (document.getElementById("noteIn")) cur.noteVal = document.getElementById("noteIn").value.trim();
      cur.isSaved = true;
      lastDate = cur.actionDate;
      ITEMS.forEach(x => { if (!x.isSaved) x.actionDate = lastDate; });

      const t = document.getElementById("toast");
      t.style.display = "block";
      t.textContent = "☁️ กำลังบันทึกลง Google Sheet...";

      const checks = cur.subs.map((s, idx) => s.type !== "amp" ? cur.subVals[idx] : null).filter(x => x !== null);
      const payload = {
        round: document.getElementById("round").value,
        id: cur.id, category: cur.cat, name: cur.name,
        actionDate: cur.actionDate, ampValue: cur.ampValue,
        checks: checks, calPercent: cur.calPercent, note: cur.noteVal,
        photos: cur.photos
      };

      fetch(window.location.href, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(payload)
      }).then(() => {
        t.textContent = "✅ บันทึกเครื่อง #" + cur.id + " ลง Google Sheet สำเร็จ!";
        setTimeout(() => t.style.display = "none", 3000);
      }).catch(() => {});

      curIdx = (curIdx + 1) % ITEMS.length;
      if (!ITEMS[curIdx].isSaved) ITEMS[curIdx].actionDate = lastDate;
      render();
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
    render();
  </script>
</body>
</html>`;
}

