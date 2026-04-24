/**
 * V5_Menu.gs
 * Version: 5.0
 * Description: Creates the custom menu '🚀 LMDS V5' in Google Sheets, organizing all system functions 
 *              into logical groups for easy access by Admin and Users.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🚀 LMDS V5')
    
    // --- กลุ่มที่ 1: การตั้งค่าเบื้องต้น (ใช้ครั้งเดียวหรือเมื่อมีการเปลี่ยนแปลงโครงสร้าง) ---
    .addSubMenu(ui.createMenu('⚙️ Setup & Config')
      .addItem('1. สร้างโครงสร้างฐานข้อมูลใหม่ (Init Schema)', 'V5_InitAllSheets')
      .addSeparator()
      .addItem('🔍 ตรวจสอบความถูกต้องของระบบ', 'V5_CheckSystemIntegrity'))
    
    // --- กลุ่มที่ 2: ส่วนที่ 2 - เครื่องจักรจัดการข้อมูลดิบ (Master Data Engine) ---
    .addSubMenu(ui.createMenu('🏭 Part 2: จัดการฐานข้อมูล')
      .addItem('🚜 1. นำเข้าข้อมูลดิบ -> สร้าง Master Data', 'V5_IngestRawDataToMaster')
      .addSeparator()
      .addItem('📋 2. ดูรายการติดขัด (Conflict Queue)', 'V5_ShowConflictQueueUI')
      .addSeparator()
      .addItem('🔄 รีเซ็ตสถานะ Sync ข้อมูลดิบ', 'V5_ResetSyncStatus'))
    
    // --- กลุ่มที่ 3: ส่วนที่ 1 - ปฏิบัติการรายวัน (Daily Workflow) ---
    .addSubMenu(ui.createMenu('📅 Part 1: ปฏิบัติการรายวัน')
      .addItem('📥 1. โหลดข้อมูลจาก SCG API (ลงชีต Data)', 'V5_FetchRawDataFromAPI')
      .addSeparator()
      .addItem('✨ 2. จับคู่ข้อมูลสะอาด & เติมพิกัด/อีเมล', 'V5_MatchDailyData')
      .addSeparator()
      .addItem('📊 3. สร้างรายงานสรุป (Shipment/Owner)', 'V5_GenerateDailySummaries')
      .addSeparator()
      .addItem('🧹 ล้างข้อมูลรายวัน (Input/Data/Summary)', 'V5_ClearDailySheets'))
    
    // --- กลุ่มที่ 4: เครื่องมือเสริมและการบำรุงรักษา ---
    .addSubMenu(ui.createMenu('🛠️ Maintenance & Tools')
      .addItem('🗑️ ล้างเฉพาะรายงานสรุป', 'V5_ClearSummarySheets')
      .addSeparator()
      .addItem('ℹ️ เกี่ยวกับระบบ LMDS V5', 'V5_ShowAboutInfo'))
    
    .addToUi();
    
  Logger.log("✅ Menu '🚀 LMDS V5' loaded successfully.");
}

/**
 * ฟังก์ชันแสดง Dialog สำหรับดูรายการ Conflict (แบบง่าย)
 * ในอนาคตสามารถพัฒนาเป็น Web App หรือ Sidebar ที่ซับซ้อนกว่านี้ได้
 */
function V5_ShowConflictQueueUI() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName(V5_CONFIG.SHEETS.CONFLICT_QUEUE);
  
  if (!queueSheet || queueSheet.getLastRow() < 2) {
    Browser.msgBox("✅ ไม่มีรายการติดขัดใน Queue ปัจจุบัน\nข้อมูลสะอาดเรียบร้อย!");
    return;
  }
  
  const count = queueSheet.getLastRow() - 1;
  const msg = `⚠️ พบรายการที่ต้องตรวจสอบ ${count} รายการ\n\n` +
              `กรุณาเปิดชีต 'Conflict_Queue' เพื่อตรวจสอบรายละเอียด\n` +
              `และดำเนินการอนุมัติ (Approve) หรือ แก้ไข (Edit) ด้วยตนเอง`;
              
  Browser.msgBox(msg);
  // เปิดชีต Conflict_Queue ให้ทันที
  ss.setActiveSheet(queueSheet);
}

/**
 * ฟังก์ชันแสดงข้อมูลเกี่ยวกับระบบ
 */
function V5_ShowAboutInfo() {
  const info = `📘 ระบบจัดการข้อมูลหลัก LMDS Version 5.0\n\n` +
               `🔹 ส่วนที่ 1 (Daily Ops): ดึง API -> จับคู่ -> สรุปผล\n` +
               `🔹 ส่วนที่ 2 (Master Data): ทำความสะอาดข้อมูลดิบ -> สร้าง Entity/Location\n\n` +
               `ผู้พัฒนา: AI Assistant\n` +
               `สถานะ: พร้อมใช้งาน ✅`;
               
  Browser.msgBox(info);
}

/**
 * ฟังก์ชันตรวจสอบความถูกต้องของระบบ (เรียกจาก Config)
 */
function V5_CheckSystemIntegrity() {
  try {
    V5_CONFIG.validateSystemIntegrity();
    Browser.msgBox("✅ ระบบผ่านการตรวจสอบความถูกต้อง!\nชีตสำคัญครบถ้วน พร้อมใช้งาน");
  } catch (e) {
    Browser.msgBox(`❌ พบข้อผิดพลาดในการตรวจสอบระบบ:\n${e.message}`);
  }
}
