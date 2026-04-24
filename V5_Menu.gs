/**
 * V5_Menu.gs
 * Version: 5.0
 * Description: Creates the custom menu for LMDS V5.0
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 LMDS V5')
    .addSubMenu(ui.createMenu('⚙️ Setup & Config')
      .addItem('1. สร้างโครงสร้างฐานข้อมูลใหม่', 'V5_InitAllSheets'))
    .addSubMenu(ui.createMenu('ส่วนที่ 2: จัดการฐานข้อมูล')
      .addItem('🚜 นำเข้าข้อมูลดิบสู่ฐานหลัก', 'V5_IngestRawDataToMaster')
      .addItem('🔄 ล้างสถานะ Sync (เริ่มใหม่)', 'V5_ResetSyncStatus')
      .addItem('🛠️ จัดการรายการติดขัด (Queue)', 'V5_ShowConflictQueueUI'))
    .addSubMenu(ui_menu_part1())
    .addSubMenu(ui.createMenu('📈 รายงานและเครื่องมือ')
      .addItem('ดูรายงานคุณภาพข้อมูล', 'V5_ShowQualityReport')
      .addItem('เกี่ยวกับระบบ V5', 'V5_ShowAbout'))
    .addToUi();
}

function ui_menu_part1() {
  return ui.createMenu('ส่วนที่ 1: ปฏิบัติการรายวัน')
    .addItem('📥 โหลดข้อมูล Shipment (API)', 'V5_FetchRawDataFromAPI')
    .addItem('🔄 จับคู่ข้อมูลสะอาด (Fill Coordinates)', 'V5_MatchDailyData')
    .addItem('📊 สร้างรายงานสรุป', 'V5_GenerateDailySummaries')
    .addItem('🧹 ล้างข้อมูลรายวัน', 'V5_ClearDailySheets');
}

/**
 * ฟังก์ชันแสดงรายงานคุณภาพข้อมูลอย่างง่าย
 */
function V5_ShowQualityReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const entSheet = ss.getSheetByName(V5_CONFIG.SHEETS.ENTITIES);
  const locSheet = ss.getSheetByName(V5_CONFIG.SHEETS.LOCATIONS);
  const mapSheet = ss.getSheetByName(V5_CONFIG.SHEETS.MAP);
  const queueSheet = ss.getSheetByName(V5_CONFIG.SHEETS.CONFLICT_QUEUE);
  
  if (!entSheet || !locSheet) {
    Browser.msgBox("❌ ไม่พบชีตฐานข้อมูล กรุณา Run Setup ก่อน");
    return;
  }

  const countEnt = Math.max(0, entSheet.getLastRow() - 1);
  const countLoc = Math.max(0, locSheet.getLastRow() - 1);
  const countMap = Math.max(0, mapSheet.getLastRow() - 1);
  const countQueue = Math.max(0, queueSheet ? queueSheet.getLastRow() - 1 : 0);
  
  // คำนวณความครอบคลุม (Coverage) แบบง่ายๆ
  const coverage = countMap > 0 && countEnt > 0 ? Math.round((countMap / countEnt) * 100) : 0;

  const msg = `
📊 รายงานคุณภาพข้อมูล (Data Quality Report)
-------------------------------------------
👥 จำนวน Entities (ตัวตน): ${countEnt} รายการ
📍 จำนวน Locations (สถานที่): ${countLoc} รายการ
🔗 จำนวนความสัมพันธ์ (Mappings): ${countMap} รายการ
⚠️ รายการติดขัดคงค้าง: ${countQueue} รายการ

💡 ดัชนีความสมบูรณ์:
- ความครอบคลุม Entity ที่มีพิกัด: ${coverage}%
(ยิ่งสูงยิ่งดี หมายถึง Entity ส่วนใหญ่มี Location จับคู่แล้ว)

✅ สถานะระบบ: ปกติ
  `;
  
  Browser.msgBox(msg);
}

function V5_ShowAbout() {
  Browser.msgBox("LMDS Version 5.0\nระบบจัดการข้อมูลหลักโลจิสติกส์\nพัฒนาขึ้นเพื่อแก้ปัญหาความซ้ำซ้อนของข้อมูล\nโดยไม่พึ่งพา AI 100%");
}
