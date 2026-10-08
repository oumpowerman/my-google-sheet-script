// =========================================================================
// 1. ฟังก์ชันอัตโนมัติ (ทำงานทันทีเมื่อมีการแก้ไขข้อมูลบนหน้า Sheet)
// =========================================================================
function onEdit(e) {
  // ดึงข้อมูลตำแหน่งของเซลล์ที่มีการแก้ไข
  const range = e.range;
  const sheet = range.getSheet();
  const col = range.getColumn();
  const row = range.getRow();

  // ตรวจสอบว่าทำงานเฉพาะบนแท็บ "Input" เท่านั้น ถ้าแก้แท็บอื่นให้หยุดการทำงานทันที
  if (sheet.getName() !== "Input") return;

  // ตรวจสอบส่วน Checkbox ซิงก์ข้อมูลที่เซลล์ F1 (แถว 1 คอลัมน์ 6 คือช่อง TRUE/Checkbox)
  if (row === 1 && col === 6) {
    if (range.getValue() === true) { 
      // เรียกใช้งานฟังก์ชันหลักเพื่อสร้างปฏิทินและซิงก์ข้อมูลเด็กฝึกงานทั้งหมด
      syncInternData();
      // เมื่อซิงก์เสร็จแล้ว ให้คืนค่า Checkbox กลับมาเป็น false (เอาเครื่องหมายถูกออก) เพื่อรอการกดครั้งต่อไป
      range.setValue(false);
    }
    return;
  }

  // ดักจับเมื่อมีการแก้ไขชื่อเด็กฝึกงาน (Col 1: คอลัมน์ A) หรือแก้ไขรหัสสี (Col 5: คอลัมน์ E)
  if (row > 1 && (col === 1 || col === 5)) {
    const nameValue = sheet.getRange(row, 1).getValue(); 
    const colorValue = sheet.getRange(row, 5).getValue().toString().trim(); // ดึงรหัสสีจากคอลัมน์ E (Col 5)
    const previewCell = sheet.getRange(row, 6); // นำมาแสดงผลพรีวิวตัวอย่างที่คอลัมน์ F (Col 6)

    // ตรวจสอบว่ามีรหัสสีและขึ้นต้นด้วยเครื่องหมาย # หรือไม่
    if (colorValue && colorValue.startsWith("#")) {
      previewCell.setBackground(colorValue)
                 .setValue(nameValue)
                 .setFontColor("#000000")
                 .setFontWeight("bold")
                 .setHorizontalAlignment("center");
    }
  }
}


// =========================================================================
// 4. ฟังก์ชันหลักสำหรับ Sync ข้อมูลเด็กฝึกงานลงบนปฏิทิน
// =========================================================================
function syncInternData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("Input");
  const calSheet = ss.getSheetByName("Calendar") || ss.insertSheet("Calendar");
  
  // ตรวจสอบจำนวนแถวที่มีข้อมูล หากมีข้อมูลน้อยกว่าแถวที่ 2 (มีแค่หัวตาราง) ให้หยุดการทำงาน
  const lastRow = inputSheet.getLastRow();
  if (lastRow < 2) return;

  // -----------------------------------------------------------------------
  // ดึงข้อมูลเด็กฝึกงานทั้งหมด (คอลัมน์ A ถึง E รวม 5 คอลัมน์: ชื่อ, วันที่เริ่ม, วันที่จบ, KEY CARD, รหัสสี)
  // -----------------------------------------------------------------------
  const internData = inputSheet.getRange(2, 1, lastRow - 1, 5).getValues();

  // -----------------------------------------------------------------------
  // หาช่วงเดือนและปีที่จำเป็นจริง (Start Month/Year ถึง End Month/Year)
  // -----------------------------------------------------------------------
  const currentSystemDate = new Date();
  let currentSystemYear = currentSystemDate.getFullYear();
  if (currentSystemYear > 2400) currentSystemYear -= 543;

  let minDate = null;
  let maxDate = null;

  internData.forEach(row => {
    let startDate = new Date(row[1]);
    let endDate = new Date(row[2]);

    // แปลง พ.ศ. เป็น ค.ศ. หากมี
    if (!isNaN(startDate.getTime()) && startDate.getFullYear() > 2400) {
      startDate.setFullYear(startDate.getFullYear() - 543);
    }
    if (!isNaN(endDate.getTime()) && endDate.getFullYear() > 2400) {
      endDate.setFullYear(endDate.getFullYear() - 543);
    }

    // ป้องกันกรณีวันที่สิ้นสุดพิมพ์แค่เดือน-วัน แล้วระบบใส่ปีเดียวกัน ทำให้ endDate < startDate
    // เช่น เริ่ม 10-Nov-2026 จบ 6-Mar-2026 ให้ปรับเป็น 6-Mar-2027 อัตโนมัติ
    if (!isNaN(startDate.getTime()) && !isNaN(endDate.getTime())) {
      if (endDate < startDate) {
        endDate.setFullYear(endDate.getFullYear() + 1);
      }
    }

    if (!isNaN(startDate.getTime())) {
      if (!minDate || startDate < minDate) minDate = new Date(startDate);
      if (!maxDate || startDate > maxDate) maxDate = new Date(startDate);
    }

    if (!isNaN(endDate.getTime())) {
      if (!minDate || endDate < minDate) minDate = new Date(endDate);
      if (!maxDate || endDate > maxDate) maxDate = new Date(endDate);
    }
  });

  // กำหนดเดือนและปีเริ่มต้น-สิ้นสุดที่จะสร้าง (เฉพาะช่วงที่มีเด็กฝึกงาน)
  let startYear = minDate ? minDate.getFullYear() : currentSystemYear;
  let startMonth = minDate ? minDate.getMonth() : currentSystemDate.getMonth();
  let endYear = maxDate ? maxDate.getFullYear() : startYear;
  let endMonth = maxDate ? maxDate.getMonth() : startMonth;

  // 1. สั่งสร้างโครงสร้างปฏิทินเฉพาะเดือนที่จำเป็นตั้งแต่ (startYear, startMonth) ถึง (endYear, endMonth)
  createInternCalendar(startYear, startMonth, endYear, endMonth); 
  
  // 2. ดึงข้อมูลปฏิทินที่เพิ่งสร้างเสร็จทั้งหมดขึ้นมาไว้ใน Memory (Array 2 มิติ)
  const fullRange = calSheet.getDataRange();
  const fullValues = fullRange.getValues(); 
  
  // กำหนด Array สำหรับเก็บกฎ Conditional Formatting
  const rules = [];
  // กำหนดขอบเขตสีให้ครอบคลุมแถวทั้งหมดของ Sheet โดยไม่เกินขนาดจริงของ Sheet
  const rangeToApply = calSheet.getRange(1, 1, calSheet.getMaxRows(), 7);

  // -----------------------------------------------------------------------
  // เพิ่มกฎเหล็ก: เดือนที่ผ่านมาแล้วทั้งหมดต้องเป็นสีเทา (Priority สูง ชนะทุกสี)
  // -----------------------------------------------------------------------
  const today = new Date();
  const monthNames = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
  ];
  let curYear = today.getFullYear();
  if (curYear > 2400) curYear -= 543;
  const currentMonthTag = `${monthNames[today.getMonth()]} ${curYear}`;
  const currentMonthFinder = calSheet.createTextFinder(currentMonthTag).findNext();

  if (currentMonthFinder) {
    const lastRowOfPast = currentMonthFinder.getRow() - 1; 
    if (lastRowOfPast > 0 && lastRowOfPast <= calSheet.getMaxRows()) {
      const pastRange = calSheet.getRange(1, 1, lastRowOfPast, 7);
      const pastRule = SpreadsheetApp.newConditionalFormatRule()
        .whenCellNotEmpty() 
        .setBackground("#d1d1d1") // สีเทาเข้ม
        .setFontColor("#7f8c8d") // ตัวอักษรสีเทาหม่น
        .setRanges([pastRange])
        .build();
      rules.push(pastRule); 
    }
  }

  // -----------------------------------------------------------------------
  // เพิ่มกฎ: ไฮไลท์ "วันปัจจุบัน" ให้เป็นสีเหลืองทอง
  // -----------------------------------------------------------------------
  const todayDate = today.getDate();
  if (currentMonthFinder) {
    const startRow = currentMonthFinder.getRow() + 2;
    // อ่านค่าเฉพาะภายในตารางของเดือนปัจจุบัน โดยจำกัดแถวไม่ให้หลุดขนาดจริงของ Sheet
    const rowsToRead = Math.min(50, Math.max(1, calSheet.getMaxRows() - startRow + 1));
    const currentMonthValues = calSheet.getRange(startRow, 1, rowsToRead, 7).getValues();
    let todayFound = false;
    for (let r = 0; r < currentMonthValues.length; r++) {
      for (let c = 0; c < 7; c++) {
        if (currentMonthValues[r][c] == todayDate) {
          const todayRange = calSheet.getRange(startRow + r, c + 1); 
          const todayRule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberEqualTo(todayDate) 
            .setBackground("#fff700") // สีเหลืองทอง
            .setFontColor("#000000")
            .setBold(true)
            .setRanges([todayRange])
            .build();
          rules.unshift(todayRule); 
          todayFound = true;
          break;
        }
      }
      if (todayFound) break;
    }
  }

  // -----------------------------------------------------------------------
  // 3. วนลูปรายชื่อเด็กฝึกงานแต่ละคน เพื่อนำชื่อไปหยอดลงใน Array (ใน Memory)
  // -----------------------------------------------------------------------
  internData.forEach(row => {
    const name = row[0];
    let startDate = new Date(row[1]);
    let endDate = new Date(row[2]);
    const colorCode = row[4]; 

    // หากไม่มีชื่อ หรือวันที่เริ่มไม่ถูกต้อง ให้ข้ามแถวนี้ไป
    if (!name || isNaN(startDate.getTime())) return;

    // ตรวจสอบและปรับกรณีวันที่จบข้ามปี
    if (!isNaN(endDate.getTime()) && endDate < startDate) {
      endDate = new Date(endDate.getFullYear() + 1, endDate.getMonth(), endDate.getDate());
    }

    // สร้างกฎสีประจำตัว (Conditional Formatting) สำหรับชื่อของเด็กฝึกงานคนนี้
    if (colorCode) {
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(name)
        .setBackground(colorCode)
        .setFontColor("#000000")
        .setRanges([rangeToApply])
        .build();
      rules.push(rule);
    }

    // วนลูปทีละวันตั้งแต่วันเริ่มต้น จนถึงวันสิ้นสุดการฝึกงาน
    for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
      const timeDiff = endDate.getTime() - d.getTime();
      const daysLeft = Math.round(timeDiff / (1000 * 3600 * 24)); // จำนวนวันที่เหลือก่อนจบ

      let currentDYear = d.getFullYear();
      if (currentDYear > 2400) currentDYear -= 543;
      const searchTag = `${monthNames[d.getMonth()]} ${currentDYear}`;
      const day = d.getDate();

      // ค้นหาตำแหน่งหัวข้อเดือนใน Memory Array (fullValues)
      for (let i = 0; i < fullValues.length; i++) {
        if (fullValues[i][0] == searchTag) {
          const startRowInArray = i + 2; 
          let found = false;

          // วนหาช่องตัวเลขวันที่ภายในกรอบของเดือนนั้น
          for (let r = 0; r < 45 && (startRowInArray + r) < fullValues.length; r++) {
            for (let c = 0; c < 7; c++) {
              if (fullValues[startRowInArray + r] && fullValues[startRowInArray + r][c] == day) {
                // เมื่อเจอเลขวันที่แล้ว ให้หยอดชื่อลงในแถวว่างใต้เลขวันที่ (วนหาไม่เกิน 10 แถว)
                for (let offset = 1; offset <= 10; offset++) {
                  const targetRowIdx = startRowInArray + r + offset;
                  if (targetRowIdx < fullValues.length && fullValues[targetRowIdx] && fullValues[targetRowIdx][c] === "") {
                    fullValues[targetRowIdx][c] = name; 

                    // --- สร้างกฎสีไล่เฉดเฉพาะจุด (สำหรับ 3 วันสุดท้ายของการฝึกงาน) ---
                    if (daysLeft <= 2 && daysLeft >= 0 && (targetRowIdx + 1) <= calSheet.getMaxRows()) {
                      const targetRange = calSheet.getRange(targetRowIdx + 1, c + 1);
                      let fadeColor = daysLeft === 0 ? "#c0392b" : (daysLeft === 1 ? "#e74c3c" : "#fab1a0");
                      let fontColor = daysLeft === 2 ? "#000000" : "#ffffff";
                      
                      const fadeRule = SpreadsheetApp.newConditionalFormatRule()
                        .whenTextEqualTo(name)
                        .setBackground(fadeColor)
                        .setFontColor(fontColor)
                        .setRanges([targetRange])
                        .build();
                      rules.unshift(fadeRule); 
                    }

                    found = true;
                    break;
                  }
                }
              }
              if (found) break;
            }
            if (found) break;
          }
          break;
        }
      }
    }
  });

  // -----------------------------------------------------------------------
  // ชุดสีกรณีมีแท็กเคานต์ดาวน์แบบย่อ [E1], [E2], [E3]
  // -----------------------------------------------------------------------
  const countdownColors = [
    {tag: "[E1]", color: "#c0392b", text: "#ffffff"}, // แดงเข้ม - วันสุดท้าย
    {tag: "[E2]", color: "#e74c3c", text: "#ffffff"}, // แดงสว่าง - 1 วันก่อนจบ
    {tag: "[E3]", color: "#fab1a0", text: "#000000"}  // ชมพูพาสเทล - 2 วันก่อนจบ
  ];

  countdownColors.forEach(item => {
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains(item.tag)
      .setBackground(item.color)
      .setFontColor(item.text)
      .setBold(true)
      .setRanges([rangeToApply])
      .build();
    rules.unshift(rule); 
  });

  // -----------------------------------------------------------------------
  // 4. บันทึกข้อมูลทั้งก้อนจาก Memory กลับลงบนหน้า Sheet พร้อมจัดรูปแบบฟอนต์
  // -----------------------------------------------------------------------
  fullRange.setValues(fullValues)
           .setFontFamily("Anuphan")
           .setFontSize(10)
           .setVerticalAlignment("middle")
           .setHorizontalAlignment("center");
  
  // 5. บันทึกกฎ Conditional Formatting ทั้งหมดลงบนแผ่นงาน Calendar
  calSheet.setConditionalFormatRules(rules);
  
  // แจ้งเตือนเสร็จสิ้น
  ss.toast("ซิงก์ข้อมูลปฏิทินสำเร็จเรียบร้อยแล้ว!", "สำเร็จ", 4);
}

// =========================================================================
// 5. ฟังก์ชันสร้างโครงปฏิทิน (สร้างเฉพาะเดือนที่จำเป็น และขยายขนาด Sheet อัตโนมัติ)
// =========================================================================
function createInternCalendar(startYear, startMonth, endYear, endMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Calendar") || ss.insertSheet("Calendar");
  sheet.clear();
  
  // กำหนดปีและเดือนปัจจุบันสำหรับการคำนวณ Past Month
  const today = new Date();
  const currentMonth = today.getMonth(); 
  let currentYear = today.getFullYear();
  if (currentYear > 2400) currentYear -= 543;

  let sYear = (typeof startYear === "number" && !isNaN(startYear)) ? startYear : currentYear;
  let sMonth = (typeof startMonth === "number" && !isNaN(startMonth)) ? startMonth : 0;
  let eYear = (typeof endYear === "number" && !isNaN(endYear)) ? endYear : sYear;
  let eMonth = (typeof endMonth === "number" && !isNaN(endMonth)) ? endMonth : 11;

  // ป้องกันกรณีปีที่ส่งมาเป็น พ.ศ.
  if (sYear > 2400) sYear -= 543;
  if (eYear > 2400) eYear -= 543;

  // ป้องกันกรณีส่งค่าช่วงเวลาสลับกัน
  if (sYear > eYear || (sYear === eYear && sMonth > eMonth)) {
    eYear = sYear;
    eMonth = sMonth;
  }

  // -----------------------------------------------------------------------
  // คำนวณจำนวนแถวที่จำเป็นตามจำนวนเดือนที่ต้องสร้างจริง
  // -----------------------------------------------------------------------
  const totalMonths = ((eYear - sYear) * 12) + (eMonth - sMonth + 1);
  const rowsNeeded = (totalMonths * 72) + 100; // 1 เดือนใช้ ~70 แถว
  const currentMaxRows = sheet.getMaxRows();
  if (currentMaxRows < rowsNeeded) {
    sheet.insertRowsAfter(currentMaxRows, rowsNeeded - currentMaxRows);
  }

  // ตรวจสอบและเพิ่มคอลัมน์ให้มีครบอย่างน้อย 7 คอลัมน์ (A ถึง G)
  const currentMaxCols = sheet.getMaxColumns();
  if (currentMaxCols < 7) {
    sheet.insertColumnsAfter(currentMaxCols, 7 - currentMaxCols);
  }

  const months = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
  ];
  const rowsPerDay = 11; // 1 แถวเลขวันที่ + 10 แถวสำหรับใส่ชื่อ
  let currentRow = 1;

  // วนลูปสร้างเฉพาะเดือนที่จำเป็นตั้งแต่ (sYear, sMonth) ถึง (eYear, eMonth)
  for (let year = sYear; year <= eYear; year++) {
    const mStart = (year === sYear) ? sMonth : 0;
    const mEnd = (year === eYear) ? eMonth : 11;

    for (let monthIndex = mStart; monthIndex <= mEnd; monthIndex++) {
      const monthName = months[monthIndex];

      // ตรวจสอบความปลอดภัยของจำนวนแถว หากใกล้เต็มให้ขยายเพิ่มทันที
      if (currentRow + rowsPerDay + 5 > sheet.getMaxRows()) {
        sheet.insertRowsAfter(sheet.getMaxRows(), 100);
      }

      // ตรวจสอบว่าเดือนนี้เป็นเดือนที่ผ่านไปแล้วหรือไม่ (Past Month)
      let isPastMonth = false;
      if (year < currentYear) {
        isPastMonth = true;
      } else if (year === currentYear && monthIndex < currentMonth) {
        isPastMonth = true;
      }

      // --- ส่วนหัวเดือนและปี (เช่น "กันยายน 2026", "มีนาคม 2027") ---
      const headerBg = isPastMonth ? "#2c3e50" : "#34495e"; 
      sheet.getRange(currentRow, 1, 1, 7).merge().setValue(`${monthName} ${year}`)
           .setBackground(headerBg).setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
      currentRow++;

      // --- ส่วนแถวชื่อวัน (อาทิตย์ - เสาร์) ---
      const subHeaderBg = isPastMonth ? "#7f8c8d" : "#ecf0f1";
      sheet.getRange(currentRow, 1, 1, 7).setValues([["อาทิตย์", "จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์"]])
           .setBackground(subHeaderBg).setFontWeight("bold").setHorizontalAlignment("center");
      currentRow++;

      // คำนวณวันแรกของเดือน และจำนวนวันทั้งหมดในเดือนนั้น
      const firstDay = new Date(year, monthIndex, 1).getDay();
      const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
      let dateCounter = 1;

      // วนลูปสร้างสัปดาห์ (ไม่เกิน 6 สัปดาห์ต่อเดือน)
      for (let week = 0; week < 6; week++) {
        // เพิ่มแถวสำรองหากใกล้ชนเพดาน
        if (currentRow + rowsPerDay > sheet.getMaxRows()) {
          sheet.insertRowsAfter(sheet.getMaxRows(), 50);
        }

        for (let day = 0; day < 7; day++) {
          if (!(week === 0 && day < firstDay || dateCounter > daysInMonth)) {
            const cellRange = sheet.getRange(currentRow, day + 1, rowsPerDay, 1);
            
            // ถ้าเป็นเดือนที่ผ่านไปแล้ว ให้ระบายพื้นหลังเป็นสีเทาเข้ม
            if (isPastMonth) {
              cellRange.setBackground("#d1d1d1"); 
            } else {
              cellRange.setBackground("#ffffff");
            }

            sheet.getRange(currentRow, day + 1).setValue(dateCounter).setFontWeight("bold").setVerticalAlignment("top");
            cellRange.setBorder(true, true, true, true, null, null);
            dateCounter++;
          }
        }
        currentRow += rowsPerDay;
        if (dateCounter > daysInMonth) break;
      }
      currentRow += 2; // เว้นวรรค 2 บรรทัดระหว่างเดือน
    }
  }

  // ปรับความกว้างของคอลัมน์ A ถึง G ให้เท่ากัน
  sheet.setColumnWidths(1, 7, 130);
}

// =========================================================================
// 6. ฟังก์ชันเสริมสำหรับการหยอดชื่อ (คงไว้ตามโครงสร้างเดิม)
// =========================================================================
function fillNameOnly(sheet, monthLabel, day, name) {
  const textFinder = sheet.createTextFinder(monthLabel).findNext();
  if (!textFinder) return;
  
  const startRow = textFinder.getRow() + 2;
  const rowsToRead = Math.min(60, Math.max(1, sheet.getMaxRows() - startRow + 1));
  const rangeValues = sheet.getRange(startRow, 1, rowsToRead, 7).getValues();
  
  for (let r = 0; r < rangeValues.length; r++) {
    for (let c = 0; c < 7; c++) {
      if (rangeValues[r][c] == day) {
        for (let i = 1; i < 8; i++) {
          if (startRow + r + i <= sheet.getMaxRows()) {
            const cell = sheet.getRange(startRow + r + i, c + 1);
            if (cell.getValue() === "" || cell.getValue() === name) {
              cell.setValue(name).setFontWeight("bold").setHorizontalAlignment("center");
              return;
            }
          }
        }
      }
    }
  }
}

function fastFill(sheet, fullValues, monthLabel, day, name, color) {
  let startRowIndex = -1;
  for (let i = 0; i < fullValues.length; i++) {
    if (fullValues[i][0] == monthLabel) {
      startRowIndex = i + 2;
      break;
    }
  }

  if (startRowIndex === -1) return;

  for (let r = 0; r < 50; r++) {
    for (let c = 0; c < 7; c++) {
      if (fullValues[startRowIndex + r] && fullValues[startRowIndex + r][c] == day) {
        for (let i = 1; i <= 7; i++) {
          const targetCell = sheet.getRange(startRowIndex + r + i + 1, c + 1);
          if (targetCell.getValue() === "") {
            targetCell.setValue(name).setBackground(color).setFontColor("#000000").setFontWeight("bold").setHorizontalAlignment("center");
            return;
          }
        }
      }
    }
  }
}

function fillDateInCalendar(sheet, monthLabel, day, name, color) {
  const textFinder = sheet.createTextFinder(monthLabel).findNext();
  if (!textFinder) return;
  
  const startRow = textFinder.getRow() + 2; 
  const rowsToRead = Math.min(60, Math.max(1, sheet.getMaxRows() - startRow + 1));
  const range = sheet.getRange(startRow, 1, rowsToRead, 7).getValues();
  
  for (let r = 0; r < range.length; r++) {
    for (let c = 0; c < 7; c++) {
      if (range[r][c] == day) {
        for (let i = 1; i < 8; i++) {
          if (startRow + r + i <= sheet.getMaxRows()) {
            const cell = sheet.getRange(startRow + r + i, c + 1);
            if (cell.getValue() === "" || cell.getValue() === name) {
              cell.setValue(name);
              cell.setFontColor("#000000"); 
              cell.setFontWeight("bold");
              cell.setHorizontalAlignment("center");
              
              if (color && color !== "") {
                cell.setBackground(color); 
              } else {
                cell.setBackground("#3498db"); 
              }
              return;
            }
          }
        }
      }
    }
  }
}