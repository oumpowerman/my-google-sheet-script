// =========================================================================
// ฟังก์ชัน: กระโดดไปแท็บ Calendar ดึงหัวเดือนขึ้นบนสุด และโฟกัสช่องสีเหลือง
// =========================================================================
function goToTodayCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calSheet = ss.getSheetByName("Calendar");
  if (!calSheet) return;

  const today = new Date();
  const todayDate = today.getDate();
  const monthNames = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
  ];
  const currentMonthTag = `${monthNames[today.getMonth()]} ${today.getFullYear()}`;

  // 1. ค้นหาแถวหัวตารางของเดือนปัจจุบัน
  const currentMonthFinder = calSheet.createTextFinder(currentMonthTag).findNext();

  if (currentMonthFinder) {
    const monthHeaderRow = currentMonthFinder.getRow();
    const startRow = monthHeaderRow + 2;
    const currentMonthValues = calSheet.getRange(startRow, 1, 55, 7).getValues();

    // 2. ค้นหาพิกัดช่องวันที่ของวันนี้ (ช่องสีเหลือง)
    let targetCell = null;
    for (let r = 0; r < currentMonthValues.length; r++) {
      for (let c = 0; c < 7; c++) {
        if (currentMonthValues[r][c] == todayDate) {
          targetCell = calSheet.getRange(startRow + r, c + 1);
          break;
        }
      }
      if (targetCell) break;
    }

    // 3. สลับแท็บไปหน้า Calendar
    ss.setActiveSheet(calSheet);

    // 4. เทคนิคดึงหน้าจอ: แกล้งกระโดดไปแถวล่างสุดของเดือนก่อน
    const bottomRow = Math.min(monthHeaderRow + 55, calSheet.getMaxRows());
    calSheet.setActiveRange(calSheet.getRange(bottomRow, 1));
    SpreadsheetApp.flush(); // สั่งให้หน้าจอเลื่อนลงไปล่างสุดทันที

    // 5. ดีดกลับขึ้นมาเลือกที่หัวเดือน (การเลื่อนขึ้นจะบังคับให้หัวเดือนอยู่ขอบบนสุดของจอ)
    calSheet.setActiveRange(calSheet.getRange(monthHeaderRow, 1));
    SpreadsheetApp.flush();

    // 6. เลื่อนเคอร์เซอร์ไปชี้ที่ช่องสีเหลืองของวันนี้
    if (targetCell) {
      calSheet.setCurrentCell(targetCell);
    }
  } else {
    ss.setActiveSheet(calSheet);
  }
}