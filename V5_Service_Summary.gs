/**
 * V5_Service_Summary.gs
 * Version: 5.0
 * Description: Generates summary reports ('สรุป_Shipment' and 'สรุป_เจ้าของสินค้า') 
 *              based on the fully processed data in the 'Data' sheet.
 */

function V5_GenerateDailySummaries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(V5_CONFIG.SHEETS.DATA);
  const sumShipmentSheet = ss.getSheetByName(V5_CONFIG.SHEETS.SUMMARY_SHIPMENT);
  const sumOwnerSheet = ss.getSheetByName(V5_CONFIG.SHEETS.SUMMARY_OWNER);

  if (!dataSheet || dataSheet.getLastRow() < 2) {
    Browser.msgBox("⚠️ ไม่พบข้อมูลในชีต 'Data' กรุณาดึงข้อมูลและจับคู่ให้เสร็จสิ้นก่อนสร้างรายงาน");
    return;
  }

  // ดึงข้อมูลทั้งหมดจากชีต Data
  const data = dataSheet.getDataRange().getValues();
  const headers = data[0];
  
  // Mapping Column Index (หาตำแหน่งคอลัมน์แบบ Dynamic)
  const col = {
    shipmentNo: headers.indexOf("ShipmentNo") + 1,
    truckLicense: headers.indexOf("TruckLicense") + 1,
    planDelivery: headers.indexOf("PlanDelivery") + 1,
    soldToName: headers.indexOf("SoldToName") + 1,
    deliveryStatus: headers.indexOf("DeliveryStatus") + 1,
    itemQuantity: headers.indexOf("ItemQuantity") + 1,
    itemWeight: headers.indexOf("ItemWeight") + 1
  };

  Logger.log("เริ่มสร้างรายงานสรุป...");

  // --- 1. สร้างสรุป Shipment ---
  const shipmentGroups = {};
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const sNo = row[col.shipmentNo - 1] || "UNKNOWN";
    const tLicense = row[col.truckLicense - 1] || "N/A";
    const key = `${sNo}|${tLicense}`; // Unique Key สำหรับ Shipment

    if (!shipmentGroups[key]) {
      shipmentGroups[key] = {
        shipmentNo: sNo,
        truckLicense: tLicense,
        planDelivery: row[col.planDelivery - 1],
        totalCount: 0,
        epodCount: 0,
        lastUpdated: new Date()
      };
    }

    // นับจำนวน
    shipmentGroups[key].totalCount++;
    
    // นับ E-POD (สมมติว่าถ้า DeliveryStatus มีคำว่า 'Completed' หรือ 'E-POD' ถือว่าเป็น E-POD)
    const status = row[col.deliveryStatus - 1] ? row[col.deliveryStatus - 1].toString().toLowerCase() : "";
    if (status.includes("completed") || status.includes("epod") || status.includes("สำเร็จ")) {
      shipmentGroups[key].epodCount++;
    }
  }

  // เขียนลงชีต สรุป_Shipment
  const shipmentHeaders = ["ShipmentKey", "ShipmentNo", "TruckLicense", "PlanDelivery", "จำนวน_ทั้งหมด", "จำนวน_E-POD_ทั้งหมด", "LastUpdated"];
  const shipmentRows = Object.values(shipmentGroups).map(g => [
    `${g.shipmentNo}|${g.truckLicense}`,
    g.shipmentNo,
    g.truckLicense,
    g.planDelivery,
    g.totalCount,
    g.epodCount,
    g.lastUpdated
  ]);

  _writeSummarySheet(sumShipmentSheet, shipmentHeaders, shipmentRows);

  // --- 2. สร้างสรุป เจ้าของสินค้า ---
  const ownerGroups = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const owner = row[col.soldToName - 1] || "UNKNOWN";
    const plan = row[col.planDelivery - 1];
    const key = `${owner}|${plan}`;

    if (!ownerGroups[key]) {
      ownerGroups[key] = {
        ownerName: owner,
        planDelivery: plan,
        totalCount: 0,
        epodCount: 0,
        lastUpdated: new Date()
      };
    }

    ownerGroups[key].totalCount++;
    
    const status = row[col.deliveryStatus - 1] ? row[col.deliveryStatus - 1].toString().toLowerCase() : "";
    if (status.includes("completed") || status.includes("epod") || status.includes("สำเร็จ")) {
      ownerGroups[key].epodCount++;
    }
  }

  // เขียนลงชีต สรุป_เจ้าของสินค้า
  const ownerHeaders = ["SummaryKey", "SoldToName", "PlanDelivery", "จำนวน_ทั้งหมด", "จำนวน_E-POD_ทั้งหมด", "LastUpdated"];
  const ownerRows = Object.values(ownerGroups).map(g => [
    `${g.ownerName}|${g.planDelivery}`,
    g.ownerName,
    g.planDelivery,
    g.totalCount,
    g.epodCount,
    g.lastUpdated
  ]);

  _writeSummarySheet(sumOwnerSheet, ownerHeaders, ownerRows);

  Browser.msgBox("✅ สร้างรายงานสรุปเสร็จสิ้น!\n- สรุป_Shipment\n- สรุป_เจ้าของสินค้า\n\nตรวจสอบข้อมูลได้ที่ชีตที่เกี่ยวข้อง");
}

/**
 * Helper Function: เขียนข้อมูลลงชีตสรุป (จัดรูปแบบพื้นฐาน)
 */
function _writeSummarySheet(sheet, headers, rows) {
  sheet.clearContents();
  
  if (rows.length === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  // เขียน Header
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#4285F4").setFontColor("white");
  
  // เขียน Data
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  
  // จัดรูปแบบ
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  
  // Format วันที่และตัวเลข (ถ้าจำเป็น)
  // ตัวอย่าง: sheet.getRange(2, 4, rows.length, 1).setNumberFormat("dd/mm/yyyy");
}

/**
 * ฟังก์ชันเสริม: ล้างชีตสรุป
 */
function V5_ClearSummarySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sumShipmentSheet = ss.getSheetByName(V5_CONFIG.SHEETS.SUMMARY_SHIPMENT);
  const sumOwnerSheet = ss.getSheetByName(V5_CONFIG.SHEETS.SUMMARY_OWNER);

  if (Browser.msgBox("ต้องการล้างข้อมูลในชีตสรุปทั้งหมดหรือไม่?", Browser.Buttons.YES_NO) === Browser.Buttons.YES) {
    if (sumShipmentSheet) sumShipmentSheet.clearContents();
    if (sumOwnerSheet) sumOwnerSheet.clearContents();
    Browser.msgBox("✅ ล้างรายงานสรุปเรียบร้อยแล้ว");
  }
}
