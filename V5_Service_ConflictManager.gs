/**
 * V5_Service_ConflictManager.gs
 * Version: 5.0
 * Description: Handles the Conflict Queue logic and serves data to the HTML UI.
 */

/**
 * แสดงหน้าต่าง UI จัดการ Conflict
 */
function V5_ShowConflictQueueUI() {
  const html = HtmlService.createHtmlOutputFromFile('ConflictQueueUI')
      .setWidth(600)
      .setHeight(500)
      .setTitle('จัดการรายการติดขัด');
  SpreadsheetApp.getUi().showModalDialog(html, 'Conflict Queue Manager');
}

/**
 * ดึงข้อมูลจากชีต Conflict_Queue เพื่อส่งให้ UI
 * @returns {Array} Array of objects containing queue details
 */
function V5_GetConflictQueueData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(V5_CONFIG.SHEETS.CONFLICT_QUEUE);
  
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // หา Index คอลัมน์
  const col = {
    id: headers.indexOf("Queue_ID") + 1,
    name: headers.indexOf("Incoming_Name") + 1,
    latlng: headers.indexOf("Incoming_LatLong") + 1,
    address: headers.indexOf("Incoming_Address") + 1,
    reason: headers.indexOf("Conflict_Reason") + 1,
    entId: headers.indexOf("Suggested_Entity_ID") + 1,
    locId: headers.indexOf("Suggested_Location_ID") + 1,
    status: headers.indexOf("Action_Required") + 1
  };

  const queueItems = [];
  
  // วนลูปเฉพาะแถวที่ยังสถานะเป็น REVIEW_REQUIRED
  for (let i = 1; i < data.length; i++) {
    if (data[i][col.status - 1] === "REVIEW_REQUIRED") {
      queueItems.push({
        rowId: i + 1, // เก็บเลขแถวจริงไว้สำหรับอัปเดต
        id: data[i][col.id - 1],
        name: data[i][col.name - 1],
        latlng: data[i][col.latlng - 1],
        address: data[i][col.address - 1],
        reason: data[i][col.reason - 1],
        entityId: data[i][col.entId - 1],
        locId: data[i][col.locId - 1]
      });
    }
  }
  
  return queueItems;
}

/**
 * ประมวลผลเมื่อกดปุ่มใน UI (Approve/Reject)
 * @param {number} rowIndex - Index ใน Array ที่ส่งไป (ต้องแปลงเป็นแถวจริง)
 * @param {string} action - 'APPROVE' หรือ 'REJECT'
 */
function V5_ProcessConflictAction(arrayIndex, action) {
  // เนื่องจาก HTML ส่ง index ของ Array มา เราต้องดึงข้อมูลใหม่เพื่อหาแถวจริง
  // วิธีที่ปลอดภัยคือดึงข้อมูลทั้งหมดมาใหม่อีกครั้งเพื่อจับคู่
  const currentQueue = V5_GetConflictQueueData();
  
  if (arrayIndex >= currentQueue.length) return currentQueue; // Out of sync
  
  const item = currentQueue[arrayIndex];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(V5_CONFIG.SHEETS.CONFLICT_QUEUE);
  const mapSheet = ss.getSheetByName(V5_CONFIG.SHEETS.MAP);
  const realRow = item.rowId;

  const user = Session.getActiveUser().getEmail();
  const now = new Date();

  if (action === 'APPROVE') {
    // 1. สร้างความสัมพันธ์ใน Entity_Loc_Map
    V5_CreateMapRelation(mapSheet, item.entityId, item.locId, "MANUAL_APPROVED");
    
    // 2. อัปเดตสถานะใน Queue เป็น APPROVED
    sheet.getRange(realRow, sheet.getHeaderRows() + 1).setValue("APPROVED_BY_USER"); // Action_Required
    sheet.getRange(realRow, sheet.getHeaderRows() + 2).setValue(user); // Resolved_By (สมมติคอลัมน์ถัดไป)
    sheet.getRange(realRow, sheet.getHeaderRows() + 3).setValue(now);   // Resolved_At
  } else {
    // REJECT / IGNORE
    // แค่เปลี่ยนสถานะ ไม่สร้างความสัมพันธ์
    sheet.getRange(realRow, sheet.getHeaderRows() + 1).setValue("REJECTED_BY_USER");
    sheet.getRange(realRow, sheet.getHeaderRows() + 2).setValue(user);
    sheet.getRange(realRow, sheet.getHeaderRows() + 3).setValue(now);
  }

  // ส่ง返回列表ที่อัปเดตแล้วกลับไปให้ UI แสดงผลใหม่
  return V5_GetConflictQueueData();
}
