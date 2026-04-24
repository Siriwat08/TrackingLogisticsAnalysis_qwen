/**
 * V5_Service_ConflictManager.gs
 * Version: 5.0
 * Description: Manages the Conflict Queue with a simple UI to Approve/Reject suggestions.
 */

function V5_ShowConflictQueueUI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName(V5_CONFIG.SHEETS.CONFLICT_QUEUE);
  
  if (!queueSheet || queueSheet.getLastRow() < 2) {
    Browser.msgBox("✅ ไม่มีรายการติดขัดในคิว (Conflict Queue ว่าง)");
    return;
  }

  const html = HtmlService.createHtmlOutputFromFile('ConflictQueueUI')
      .setWidth(800)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'จัดการรายการติดขัด (Conflict Queue)');
}

function V5_GetConflictItems() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName(V5_CONFIG.SHEETS.CONFLICT_QUEUE);
  const data = queueSheet.getDataRange().getValues();
  const headers = data[0];
  const items = [];

  // กรองเฉพาะแถวที่ยังไม่ถูกแก้ไข (Action_Required = REVIEW_REQUIRED)
  const colAction = headers.indexOf("Action_Required") + 1;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][colAction - 1] === "REVIEW_REQUIRED") {
      items.push({
        rowIndex: i + 1,
        id: data[i][0],
        name: data[i][2],
        latLong: data[i][3],
        address: data[i][4],
        suggestedEntId: data[i][5],
        suggestedLocId: data[i][6],
        reason: data[i][7]
      });
    }
  }
  return items;
}

function V5_ResolveConflict(rowIndex, action, note) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName(V5_CONFIG.SHEETS.CONFLICT_QUEUE);
  const mapSheet = ss.getSheetByName(V5_CONFIG.SHEETS.MAP);
  const user = Session.getActiveUser().getEmail();
  const now = new Date();

  const headers = queueSheet.getRange(1, 1, 1, queueSheet.getLastColumn()).getValues()[0];
  const colAction = headers.indexOf("Action_Required") + 1;
  const colResolvedBy = headers.indexOf("Resolved_By") + 1;
  const colResolvedAt = headers.indexOf("Resolved_At") + 1;
  const colNote = headers.indexOf("Resolution_Note") + 1;
  const colEntId = headers.indexOf("Suggested_Entity_ID") + 1;
  const colLocId = headers.indexOf("Suggested_Location_ID") + 1;

  const row = queueSheet.getRange(rowIndex, 1, 1, queueSheet.getLastColumn()).getValues()[0];
  const entId = row[colEntId - 1];
  const locId = row[colLocId - 1];

  if (action === "APPROVE") {
    // อนุมัติ: สร้างความสัมพันธ์ใน Map Sheet
    V5_CreateMapRelation(mapSheet, entId, locId, "MANUAL_APPROVED");
    queueSheet.getRange(rowIndex, colAction).setValue("APPROVED");
  } else if (action === "REJECT") {
    // ปฏิเสธ: แค่บันทึกสถานะ
    queueSheet.getRange(rowIndex, colAction).setValue("REJECTED");
  }

  // อัปเดตข้อมูลผู้แก้ไขและหมายเหตุ
  queueSheet.getRange(rowIndex, colResolvedBy).setValue(user);
  queueSheet.getRange(rowIndex, colResolvedAt).setValue(now);
  queueSheet.getRange(rowIndex, colNote).setValue(note || "");

  return "success";
}

// สร้างไฟล์ HTML แบบ Inline สำหรับ UI ง่ายๆ (ไม่ต้องสร้างไฟล์ .html แยก)
function createConflictUIHtml() {
  const htmlContent = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: sans-serif; padding: 10px; }
          .item { border: 1px solid #ccc; padding: 10px; margin-bottom: 10px; border-radius: 5px; background: #fff; }
          .header { font-weight: bold; color: #d93025; }
          .btn { padding: 8px 15px; margin: 5px; cursor: pointer; border: none; border-radius: 3px; color: white; }
          .btn-approve { background-color: #188038; }
          .btn-reject { background-color: #d93025; }
          .btn:hover { opacity: 0.9; }
          input[type="text"] { width: 90%; padding: 5px; margin-top: 5px; }
          .loading { text-align: center; color: #666; }
        </style>
      </head>
      <body>
        <h3>🛑 รายการติดขัดที่ต้องตรวจสอบ</h3>
        <div id="content">กำลังโหลด...</div>

        <script>
          window.onload = function() {
            google.script.run.withSuccessHandler(renderItems).V5_GetConflictItems();
          };

          function renderItems(items) {
            const div = document.getElementById('content');
            if (items.length === 0) {
              div.innerHTML = '<p>✅ ไม่มีรายการค้างชำระ</p>';
              return;
            }

            let html = '';
            items.forEach(item => {
              html += \`
                <div class="item">
                  <div class="header">ชื่อ: \${item.name}</div>
                  <div>พิกัด: \${item.latLong}</div>
                  <div>ที่อยู่: \${item.address}</div>
                  <div style="color:#666; font-size:0.9em;">เหตุผล: \${item.reason}</div>
                  <div style="font-size:0.8em; color:#888;">Entity ID: \${item.suggestedEntId}<br>Loc ID: \${item.suggestedLocId}</div>
                  <input type="text" id="note-\${item.rowIndex}" placeholder="หมายเหตุ (ถ้ามี)">
                  <br>
                  <button class="btn btn-approve" onclick="resolve(\${item.rowIndex}, 'APPROVE')">✅ อนุมัติ (เชื่อมโยง)</button>
                  <button class="btn btn-reject" onclick="resolve(\${item.rowIndex}, 'REJECT')">❌ ปฏิเสธ</button>
                </div>
              \`;
            });
            div.innerHTML = html;
          }

          function resolve(rowIndex, action) {
            const note = document.getElementById('note-' + rowIndex).value;
            const btns = event.target.parentElement.querySelectorAll('button');
            btns.forEach(b => b.disabled = true);
            event.target.innerText = "กำลังบันทึก...";

            google.script.run.withSuccessHandler(function(res) {
              if (res === 'success') {
                event.target.parentElement.parentElement.style.background = action === 'APPROVE' ? '#e6f4ea' : '#fce8e6';
                event.target.parentElement.innerHTML = '<b>ดำเนินการแล้ว:</b> ' + action;
                // โหลดใหม่ทั้งหมดเพื่อให้เห็นความเปลี่ยนแปลง
                setTimeout(() => google.script.run.withSuccessHandler(renderItems).V5_GetConflictItems(), 1000);
              }
            }).V5_ResolveConflict(rowIndex, action, note);
          }
        </script>
      </body>
    </html>
  `;
  
  // บันทึกลงไฟล์ HTML จริงเพื่อให้เรียกใช้ได้
  // หมายเหตุ: ใน Apps Script เราต้องสร้างไฟล์ .html แยก หรือใช้วิธีสร้าง Blob
  // เพื่อความง่ายในคู่มือนี้ ผมจะแนะนำให้คุณสร้างไฟล์ HTML ชื่อ 'ConflictQueueUI.html' 
  // แล้ววางโค้ดด้านล่างนี้แทนครับ (ดูขั้นตอนถัดไป)
  return htmlContent;
}
