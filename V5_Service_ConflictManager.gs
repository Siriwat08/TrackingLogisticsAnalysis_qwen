/**
 * V5_Service_ConflictManager.gs
 * Version: 5.0
 * Description: Manages the Conflict Queue, provides a UI for reviewing conflicts, 
 *              and processes user decisions (Approve/Reject) to update the Master Data.
 */

// --- 1. ฟังก์ชันเปิดหน้า UI (Dialog) ---

function V5_ShowConflictQueueUI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName(V5_CONFIG.SHEETS.CONFLICT_QUEUE);
  
  if (!queueSheet) {
    Browser.msgBox("❌ ไม่พบชีต 'Conflict_Queue'");
    return;
  }

  const data = queueSheet.getDataRange().getValues();
  const headers = data[0];
  
  // กรองเฉพาะแถวที่ยังไม่แก้ไข (Status = REVIEW_REQUIRED)
  const colStatus = headers.indexOf("Action_Required") + 1;
  const pendingRows = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][colStatus - 1] === "REVIEW_REQUIRED") {
      pendingRows.push({
        rowIndex: i + 1, // Row number in sheet (1-based)
        id: data[i][0],
        name: data[i][2],
        latLng: data[i][3],
        address: data[i][4],
        suggestedEntId: data[i][5],
        suggestedLocId: data[i][6],
        reason: data[i][7]
      });
    }
  }

  if (pendingRows.length === 0) {
    Browser.msgBox("✅ เยี่ยมมาก! ไม่มีรายการติดขัดค้างอยู่");
    return;
  }

  // สร้าง HTML อย่างง่ายสำหรับแสดงใน Dialog
  let htmlContent = `
    <div style="font-family: sans-serif; padding: 10px;">
      <h3>🚨 รายการติดขัดที่ต้องตรวจสอบ (${pendingRows.length})</h3>
      <p>กรุณาตรวจสอบและเลือกการจัดการ:</p>
      <hr/>
  `;

  pendingRows.forEach((item, index) => {
    htmlContent += `
      <div style="border: 1px solid #ccc; padding: 10px; margin-bottom: 10px; border-radius: 5px; background: #f9f9f9;">
        <strong>#${index + 1}: ${item.name}</strong><br/>
        <small>พิกัด: ${item.latLng}</small><br/>
        <small>ที่อยู่: ${item.address || '-'}</small><br/>
        <span style="color: red; font-weight: bold;">⚠️ เหตุผล: ${item.reason}</span><br/>
        <br/>
        <button onclick="resolveConflict(${item.rowIndex}, 'APPROVE')" style="background:#4CAF50; color:white; border:none; padding:5px 10px; cursor:pointer;">✅ อนุมัติ (เชื่อมตามแนะนำ)</button>
        <button onclick="resolveConflict(${item.rowIndex}, 'REJECT')" style="background:#f44336; color:white; border:none; padding:5px 10px; cursor:pointer;">❌ ปฏิเสธ (ไม่เชื่อม)</button>
      </div>
    `;
  });

  htmlContent += `
      <script>
        function resolveConflict(rowIndex, action) {
          google.script.run
            .withSuccessHandler(function() {
              // Refresh หรือปิด dialog หลังจากทำสำเร็จ
              google.script.host.close(); 
              // เรียกฟังก์ชันเดิมอีกครั้งเพื่อโหลดใหม่ (หรือให้ผู้ใช้กดเปิดใหม่)
              google.script.run.V5_ShowConflictQueueUI();
            })
            .processConflictDecision(rowIndex, action);
        }
      </script>
      <hr/>
      <button onclick="google.script.host.close()" style="padding:5px 10px;">ปิดหน้าต่าง</button>
    </div>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(600)
    .setHeight(800)
    .setTitle('จัดการรายการติดขัด (Conflict Queue)');

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'จัดการรายการติดขัด');
}

// --- 2. ฟังก์ชันประมวลผลการตัดสินใจจาก UI ---

function V5_ProcessConflictDecision(rowIndex, action) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName(V5_CONFIG.SHEETS.CONFLICT_QUEUE);
  const mapSheet = ss.getSheetByName(V5_CONFIG.SHEETS.MAP);
  const entSheet = ss.getSheetByName(V5_CONFIG.SHEETS.ENTITIES);
  const locSheet = ss.getSheetByName(V5_CONFIG.SHEETS.LOCATIONS);

  const row = queueSheet.getRange(rowIndex, 1, 1, 10).getValues()[0];
  const queueId = row[0];
  const entId = row[5];
  const locId = row[6];
  const incomingName = row[2];
  
  const user = Session.getActiveUser().getEmail();
  const now = new Date();

  if (action === 'APPROVE') {
    // 1. สร้างความสัมพันธ์ใน Map Sheet
    V5_CreateMapRelation(mapSheet, entId, locId, "MANUAL_APPROVED");
    
    // 2. อัปเดตสถานะใน Queue
    queueSheet.getRange(rowIndex, 9).setValue("RESOLVED_APPROVED"); // Action_Required
    queueSheet.getRange(rowIndex, 10).setValue(user); // Resolved_By
    queueSheet.getRange(rowIndex, 11).setValue(now); // Resolved_At
    queueSheet.getRange(rowIndex, 12).setValue("User approved the suggested link."); // Note
    
    Logger.log(`Approved conflict for ${incomingName} -> Link ${entId} & ${locId}`);
    
  } else if (action === 'REJECT') {
    // กรณีปฏิเสธ: ไม่สร้างความสัมพันธ์
    // อาจจะทำเครื่องหมายว่าชื่อนี้ไม่ควรจับคู่กับพิกัดนี้ (อาจต้องเพิ่ม Logic Blacklist ในอนาคต)
    // ปัจจุบันแค่ปิดจบว่าเป็น Manual Reject
    
    queueSheet.getRange(rowIndex, 9).setValue("RESOLVED_REJECTED");
    queueSheet.getRange(rowIndex, 10).setValue(user);
    queueSheet.getRange(rowIndex, 11).setValue(now);
    queueSheet.getRange(rowIndex, 12).setValue("User rejected the link. Entities remain separate.");
    
    Logger.log(`Rejected conflict for ${incomingName}. No link created.`);
  }

  return true;
}

// --- 3. ฟังก์ชันเสริม: ดูสถิติ Queue ---

function V5_ShowConflictStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName(V5_CONFIG.SHEETS.CONFLICT_QUEUE);
  
  if (!queueSheet) {
    Browser.msgBox("ไม่พบชีต Conflict_Queue");
    return;
  }
  
  const data = queueSheet.getDataRange().getValues();
  const total = data.length - 1;
  const pending = data.filter(r => r[8] === "REVIEW_REQUIRED").length; // Col 9 is Action_Required (index 8)
  const resolved = total - pending;
  
  Browser.msgBox(`📊 สถิติ Conflict Queue:\nทั้งหมด: ${total}\n⏳ รอตรวจสอบ: ${pending}\n✅ แก้ไขแล้ว: ${resolved}`);
}
