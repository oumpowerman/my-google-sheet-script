// =========================================================================
// 2. ฟังก์ชันใหม่: เคลียร์เด็กฝึกงานเก่าที่สิ้นสุดการฝึกงานเกินกำหนดไปแท็บ "เด็กฝึกงานเก่า" (Dynamic)
// =========================================================================
function archivePastInterns() {
  // -----------------------------------------------------------------------
  // ⚙️️ [จุดตั้งค่าจุดเดียว] กำหนดจำนวนเดือนย้อนหลังที่ต้องการตัดรอบ (เช่น 1, 2 หรือ 3 เดือน)
  // แก้ไขตัวเลขตรงนี้เพียงจุดเดียว ทั้งการคำนวณและข้อความแจ้งเตือนจะเปลี่ยนตามอัตโนมัติ!
  // -----------------------------------------------------------------------
  const ARCHIVE_MONTHS = 1; // <-- เปลี่ยนเป็น 1 หรือ 2 ได้ตรงนี้เลย

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("Input");
  if (!inputSheet) return;

  // ตรวจสอบแท็บเป้าหมาย "เด็กฝึกงานเก่า" หากยังไม่มีให้สร้างใหม่อัตโนมัติ
  let archiveSheet = ss.getSheetByName("เด็กฝึกงานเก่า");
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet("เด็กฝึกงานเก่า");
  }

  // หากแท็บ "เด็กฝึกงานเก่า" ยังว่างอยู่ ให้คัดลอกหัวตารางจากแท็บ Input มาใส่ไว้
  if (archiveSheet.getLastRow() < 1) {
    inputSheet.getRange(1, 1, 1, 6).copyTo(archiveSheet.getRange(1, 1, 1, 6));
    archiveSheet.setColumnWidths(1, 6, 130);
  }

  const lastRow = inputSheet.getLastRow();
  // หากไม่มีแถวข้อมูล (มีแค่หัวตาราง) ให้แจ้งเตือนแล้วหยุดการทำงาน
  if (lastRow < 2) {
    ss.toast("ไม่มีข้อมูลเด็กฝึกงานในตาราง Input ให้ตรวจสอบ", "แจ้งเตือน", 4);
    return;
  }

  // คำนวณวันกำหนดตัดรอบ (Cutoff Date) แบบไดนามิกตามจำนวนเดือนที่กำหนดไว้ข้างต้น
  const today = new Date();
  const cutoffDate = new Date(today.getFullYear(), today.getMonth() - ARCHIVE_MONTHS, today.getDate());

  // ดึงข้อมูลแถวทั้งหมดจากแท็บ Input (คอลัมน์ A ถึง E)
  const internData = inputSheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const rowsToArchive = [];

  for (let i = 0; i < internData.length; i++) {
    const row = internData[i];
    const name = row[0];
    let startDate = new Date(row[1]);
    let endDate = new Date(row[2]);

    // หากไม่มีชื่อ หรือวันที่จบไม่ถูกต้อง ให้ข้ามไป
    if (!name || isNaN(endDate.getTime())) continue;

    // แปลง พ.ศ. เป็น ค.ศ. หากมี
    if (endDate.getFullYear() > 2400) {
      endDate.setFullYear(endDate.getFullYear() - 543);
    }
    if (!isNaN(startDate.getTime()) && startDate.getFullYear() > 2400) {
      startDate.setFullYear(startDate.getFullYear() - 543);
    }

    // กรณีวันที่สิ้นสุดพิมพ์แค่เดือน-วัน แล้วระบบใส่ปีเดียวกัน ทำให้ endDate < startDate (เช่น เริ่ม Nov จบ Mar)
    if (!isNaN(startDate.getTime()) && endDate < startDate) {
      endDate.setFullYear(endDate.getFullYear() + 1);
    }

    // ตรวจสอบว่าวันที่จบการฝึกงาน ผ่านมาเกินจำนวนเดือนที่ตั้งไว้หรือไม่
    if (endDate < cutoffDate) {
      // บันทึกหมายเลขแถวจริงบนแผ่นงาน Input (แถวแรกของข้อมูลคือแถวที่ 2)
      rowsToArchive.push(i + 2);
    }
  }

  // หากไม่มีใครเข้าเกณฑ์ ให้แจ้งเตือนแบบไดนามิกตามตัวเลขเดือน
  if (rowsToArchive.length === 0) {
    ss.toast(`ไม่พบเด็กฝึกงานที่จบการฝึกงานเกิน ${ARCHIVE_MONTHS} เดือน`, "ตรวจสอบเรียบร้อย", 4);
    return;
  }

  // 1. คัดลอกข้อมูลทั้งแถว (คอลัมน์ A ถึง F รวมสีและพรีวิว) ไปยังแท็บ "เด็กฝึกงานเก่า"
  rowsToArchive.forEach(rowIndex => {
    const nextArchiveRow = archiveSheet.getLastRow() + 1;
    inputSheet.getRange(rowIndex, 1, 1, 6).copyTo(archiveSheet.getRange(nextArchiveRow, 1, 1, 6));
  });

  // 2. ลบแถวออกจากหน้า Input จากล่างขึ้นบน เพื่อป้องกันลำดับแถวคลาดเคลื่อน
  for (let i = rowsToArchive.length - 1; i >= 0; i--) {
    inputSheet.deleteRow(rowsToArchive[i]);
  }

  // 3. สั่งรันซิงก์ปฏิทินใหม่ทันที เพื่อให้แท็บ Calendar สะอาดและอัปเดตตามรายชื่อล่าสุด
  syncInternData();

  // แจ้งเตือนสถานะความสำเร็จแบบไดนามิก
  ss.toast(`ย้ายประวัติเด็กฝึกงานเก่า (จบเกิน ${ARCHIVE_MONTHS} เดือน) ไปแท็บ "เด็กฝึกงานเก่า" เรียบร้อยแล้ว ${rowsToArchive.length} คน`, "สำเร็จ", 5);
}