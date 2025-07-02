function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function submitSurvey(formData) {
  try {
    // เปิด Google Sheets หรือสร้างใหม่
    const spreadsheetId = '10BbdDZXxUn_x8CUV2zwASjbQxH2ahNdbkW9Z170yauc'; // ใส่ ID ของ Google Sheets
    let sheet;
    
    try {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      sheet = spreadsheet.getActiveSheet();
    } catch (e) {
      // ถ้าไม่มี Spreadsheet ให้สร้างใหม่
      const spreadsheet = SpreadsheetApp.create('แบบประเมินความพึงพอใจ');
      sheet = spreadsheet.getActiveSheet();
      console.log('Spreadsheet URL: ' + spreadsheet.getUrl());
      
      // สร้างหัวข้อคอลัมน์
      const headers = [
        'วันที่ส่ง', 'เพศ', 'อายุ', 'การศึกษา', 'อาชีพ',
        // ด้านวิทยากร
        'การเตรียมตัวและความพร้อมของวิทยากร', 'การถ่ายทอดของวิทยากร', 'อธิบายเนื้อหาชัดเจนและตรงประเด็น', 
        'ใช้ภาษาเหมาะสมและเข้าใจง่าย', 'การตอบคำถามและการให้คำปรึกษา', 'เอกสารประกอบการบรรยายเหมาะสม',
        // ด้านสถานที่ อุปกรณ์ อาหาร
        'สถานที่สะอาดปลอดภัยบรรยากาศดี', 'อุปกรณ์ที่ใช้มีความพร้อม', 'ระยะเวลาที่จัดกิจกรรมเหมาะสม', 'อาหารและเครื่องดื่มเหมาะสม',
        // ด้านการบริการของเจ้าหน้าที่
        'บริการของเจ้าหน้าที่ตลอดระยะเวลาการจัดงาน', 'การประสานงานจัดการการอบรม', 'การอำนวยความสะดวกของเจ้าหน้าที่', 'การให้คำแนะนำเพิ่มเติมและตอบข้อซักถาม',
        // ด้านความรู้ความเข้าใจ
        'ความรู้เรื่องนี้ก่อนเข้ารับการอบรม', 'ความรู้เรื่องนี้หลังเข้ารับการอบรม', 'สามารถบอกประโยชน์ของการอบรม', 
        'สามารถบอกข้อดีข้อเสียของการอบรม', 'สามารถอธิบายรายละเอียดการอบรม', 'สามารถจัดระบบความคิดรู้ที่ได้รับ', 'สามารถบูรณาการความรู้ที่ได้รับกับความรู้เดิม',
        // ด้านการนำไปใช้
        'สามารถประยุกต์ใช้ความรู้ที่ได้รับในการทำงาน', 'สามารถเผยแพร่ความรู้ที่ได้รับให้ผู้อื่น', 'สามารถให้คำปรึกษาเกี่ยวกับเรื่องนี้แก่ผู้อื่น', 'มั่นใจในการใช้ความรู้ที่ได้รับจากการอบรม',
        // ข้อเสนอแนะ
        'ข้อเสนอแนะ'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      
      // จัดรูปแบบ header
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#E5E7EB');
      headerRange.setFontWeight('bold');
      headerRange.setWrap(true);
      headerRange.setVerticalAlignment('middle');
      
      // ปรับความกว้างคอลัมน์
      sheet.setColumnWidth(1, 120); // วันที่ส่ง
      for (let i = 2; i <= 5; i++) {
        sheet.setColumnWidth(i, 100); // ข้อมูลทั่วไป
      }
      for (let i = 6; i <= headers.length - 1; i++) {
        sheet.setColumnWidth(i, 80); // คะแนนประเมิน
      }
      sheet.setColumnWidth(headers.length, 200); // ข้อเสนอแนะ
    }
    
    // จัดเตรียมข้อมูลสำหรับบันทึก
    const timestamp = new Date();
    const rowData = [
      timestamp,
      formData.gender || '',
      formData.age || '',
      formData.education || '',
      formData.occupation || '',
      // ด้านวิทยากร
      parseInt(formData['instructor-1']) || 0,
      parseInt(formData['instructor-2']) || 0,
      parseInt(formData['instructor-3']) || 0,
      parseInt(formData['instructor-4']) || 0,
      parseInt(formData['instructor-5']) || 0,
      parseInt(formData['instructor-6']) || 0,
      // ด้านสถานที่ อุปกรณ์ อาหาร
      parseInt(formData['venue-1']) || 0,
      parseInt(formData['venue-2']) || 0,
      parseInt(formData['venue-3']) || 0,
      parseInt(formData['venue-4']) || 0,
      // ด้านการบริการของเจ้าหน้าที่
      parseInt(formData['service-1']) || 0,
      parseInt(formData['service-2']) || 0,
      parseInt(formData['service-3']) || 0,
      parseInt(formData['service-4']) || 0,
      // ด้านความรู้ความเข้าใจ
      parseInt(formData['knowledge-1']) || 0,
      parseInt(formData['knowledge-2']) || 0,
      parseInt(formData['knowledge-3']) || 0,
      parseInt(formData['knowledge-4']) || 0,
      parseInt(formData['knowledge-5']) || 0,
      parseInt(formData['knowledge-6']) || 0,
      parseInt(formData['knowledge-7']) || 0,
      // ด้านการนำไปใช้
      parseInt(formData['application-1']) || 0,
      parseInt(formData['application-2']) || 0,
      parseInt(formData['application-3']) || 0,
      parseInt(formData['application-4']) || 0,
      // ข้อเสนอแนะ
      formData.suggestions || ''
    ];
    
    // เพิ่มข้อมูลลงใน Sheet
    sheet.appendRow(rowData);
    
    // จัดรูปแบบแถวที่เพิ่มใหม่
    const lastRow = sheet.getLastRow();
    const dataRange = sheet.getRange(lastRow, 1, 1, rowData.length);
    dataRange.setBorder(true, true, true, true, true, true);
    
    // จัดรูปแบบข้อมูลวันที่
    const dateCell = sheet.getRange(lastRow, 1);
    dateCell.setNumberFormat('dd/mm/yyyy hh:mm:ss');
    
    return {
      success: true,
      message: 'บันทึกข้อมูลเรียบร้อยแล้ว ขอบคุณสำหรับการประเมิน'
    };
    
  } catch (error) {
    console.error('Error submitting survey:', error);
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการบันทึกข้อมูล: ' + error.toString()
    };
  }
}

function getSheetUrl() {
  const spreadsheetId = '10BbdDZXxUn_x8CUV2zwASjbQxH2ahNdbkW9Z170yauc';
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    return spreadsheet.getUrl();
  } catch (e) {
    return null;
  }
}

// ฟังก์ชันสำหรับสร้างรายงานสรุปผล
function createSummaryReport() {
  const spreadsheetId = '10BbdDZXxUn_x8CUV2zwASjbQxH2ahNdbkW9Z170yauc';
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const dataSheet = spreadsheet.getActiveSheet();
  
  // สร้างชีทใหม่สำหรับรายงานสรุป
  let summarySheet;
  try {
    summarySheet = spreadsheet.getSheetByName('สรุปผลการประเมิน');
    summarySheet.clear();
  } catch (e) {
    summarySheet = spreadsheet.insertSheet('สรุปผลการประเมิน');
  }
  
  const data = dataSheet.getDataRange().getValues();
  if (data.length <= 1) return; // ไม่มีข้อมูล
  
  const headers = data[0];
  const responses = data.slice(1);
  
  // สร้างรายงานสรุป
  summarySheet.getRange(1, 1).setValue('รายงานสรุปผลการประเมินความพึงพอใจ');
  summarySheet.getRange(1, 1, 1, 8).merge();
  summarySheet.getRange(1, 1).setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  
  // ข้อมูลทั่วไปของผู้ตอบแบบประเมิน
  let row = 3;
  summarySheet.getRange(row, 1).setValue('จำนวนผู้ตอบแบบประเมิน: ' + responses.length + ' คน');
  summarySheet.getRange(row, 1, 1, 8).merge();
  row += 2;
  
  // สร้างตารางสรุป
  summarySheet.getRange(row, 1).setValue('สรุปผลการประเมินแยกตามหมวดหมู่');
  summarySheet.getRange(row, 1, 1, 8).merge();
  summarySheet.getRange(row, 1).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#E5E7EB');
  row++;
  
  // หัวตาราง
  const tableHeaders = ['หมวดหมู่', 'ข้อคำถาม', 'จำนวนผู้ประเมิน', 'คะแนนรวม', 'ค่าเฉลี่ย', 'S.D.', 'ระดับความพึงพอใจ'];
  summarySheet.getRange(row, 1, 1, tableHeaders.length).setValues([tableHeaders]);
  summarySheet.getRange(row, 1, 1, tableHeaders.length).setFontWeight('bold').setBackground('#F3F4F6');
  row++;
  
  // จัดกลุ่มข้อมูล
  const categories = [
    { name: 'ด้านวิทยากร', start: 5, count: 6, prefix: 'instructor' },
    { name: 'ด้านสถานที่ อุปกรณ์ อาหาร', start: 11, count: 4, prefix: 'venue' },
    { name: 'ด้านการบริการของเจ้าหน้าที่', start: 15, count: 4, prefix: 'service' },
    { name: 'ด้านความรู้ความเข้าใจ', start: 19, count: 7, prefix: 'knowledge' },
    { name: 'ด้านการนำไปใช้', start: 26, count: 4, prefix: 'application' }
  ];
  
  const ratingLevels = [
    {min: 4.51, max: 5.00, text: 'มากที่สุด'},
    {min: 3.51, max: 4.50, text: 'มาก'},
    {min: 2.51, max: 3.50, text: 'ปานกลาง'},
    {min: 1.51, max: 2.50, text: 'น้อย'},
    {min: 1.00, max: 1.50, text: 'น้อยที่สุด'}
  ];
  
  // สร้างสรุปผลแยกตามหมวดหมู่
  for (const category of categories) {
    let firstRow = row;
    
    for (let i = 0; i < category.count; i++) {
      const col = category.start + i;
      const questionText = headers[col];
      const scores = responses.map(r => parseInt(r[col]) || 0).filter(s => s > 0);
      const count = scores.length;
      const sum = scores.reduce((a, b) => a + b, 0);
      const avg = count > 0 ? sum / count : 0;
      
      // คำนวณ S.D. (Standard Deviation)
      let sd = 0;
      if (count > 0) {
        const variance = scores.reduce((total, x) => {
          return total + Math.pow(x - avg, 2);
        }, 0) / count;
        sd = Math.sqrt(variance);
      }
      
      // หาระดับความพึงพอใจ
      let ratingLevel = '';
      for (const level of ratingLevels) {
        if (avg >= level.min && avg <= level.max) {
          ratingLevel = level.text;
          break;
        }
      }
      
      const rowData = [
        i === 0 ? category.name : '', // แสดงชื่อหมวดหมู่เฉพาะแถวแรกของแต่ละหมวด
        questionText,
        count,
        sum,
        avg.toFixed(2),
        sd.toFixed(2),
        ratingLevel
      ];
      
      summarySheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
      row++;
    }
    
    // ถ้ามีข้อคำถามมากกว่า 1 ข้อ ให้ merge ช่องของชื่อหมวดหมู่
    if (category.count > 1) {
      summarySheet.getRange(firstRow, 1, category.count, 1).merge();
      summarySheet.getRange(firstRow, 1).setVerticalAlignment('middle');
    }
    
    // คำนวณค่าเฉลี่ยรวมของหมวด
    const categoryStart = category.start;
    const categoryEnd = category.start + category.count - 1;
    
    let categoryScores = [];
    for (let r = 0; r < responses.length; r++) {
      for (let c = categoryStart; c <= categoryEnd; c++) {
        const score = parseInt(responses[r][c]) || 0;
        if (score > 0) categoryScores.push(score);
      }
    }
    
    const categoryCount = categoryScores.length;
    const categorySum = categoryScores.reduce((a, b) => a + b, 0);
    const categoryAvg = categoryCount > 0 ? categorySum / categoryCount : 0;
    
    // คำนวณ S.D. ของทั้งหมวด
    let categorySd = 0;
    if (categoryCount > 0) {
      const categoryVariance = categoryScores.reduce((total, x) => {
        return total + Math.pow(x - categoryAvg, 2);
      }, 0) / categoryCount;
      categorySd = Math.sqrt(categoryVariance);
    }
    
    // หาระดับความพึงพอใจของทั้งหมวด
    let categoryRating = '';
    for (const level of ratingLevels) {
      if (categoryAvg >= level.min && categoryAvg <= level.max) {
        categoryRating = level.text;
        break;
      }
    }
    
    // เพิ่มแถวสรุปของหมวด
    const categorySummaryRow = [
      'สรุป' + category.name,
      '',
      categoryCount,
      categorySum,
      categoryAvg.toFixed(2),
      categorySd.toFixed(2),
      categoryRating
    ];
    
    summarySheet.getRange(row, 1, 1, categorySummaryRow.length).setValues([categorySummaryRow]);
    summarySheet.getRange(row, 1, 1, categorySummaryRow.length).setBackground('#E5E7EB');
    summarySheet.getRange(row, 1, 1, categorySummaryRow.length).setFontWeight('bold');
    row++;
    
    // เพิ่มช่องว่างระหว่างหมวด
    row++;
  }
  
  // สรุปผลรวมทั้งหมด
  row++;
  summarySheet.getRange(row, 1).setValue('สรุปผลการประเมินโดยรวม');
  summarySheet.getRange(row, 1, 1, 7).merge();
  summarySheet.getRange(row, 1).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#D1FAE5');
  row++;
  
  // รวมทุกคะแนน
  let allScores = [];
  for (let r = 0; r < responses.length; r++) {
    for (const category of categories) {
      for (let i = 0; i < category.count; i++) {
        const score = parseInt(responses[r][category.start + i]) || 0;
        if (score > 0) allScores.push(score);
      }
    }
  }
  
  const totalCount = allScores.length;
  const totalSum = allScores.reduce((a, b) => a + b, 0);
  const totalAvg = totalCount > 0 ? totalSum / totalCount : 0;
  
  // คำนวณ S.D. ของทั้งหมด
  let totalSd = 0;
  if (totalCount > 0) {
    const totalVariance = allScores.reduce((total, x) => {
      return total + Math.pow(x - totalAvg, 2);
    }, 0) / totalCount;
    totalSd = Math.sqrt(totalVariance);
  }
  
  // หาระดับความพึงพอใจของทั้งหมด
  let totalRating = '';
  for (const level of ratingLevels) {
    if (totalAvg >= level.min && totalAvg <= level.max) {
      totalRating = level.text;
      break;
    }
  }
  
  const totalSummaryRow = [
    'ผลการประเมินโดยรวม',
    '',
    totalCount,
    totalSum,
    totalAvg.toFixed(2),
    totalSd.toFixed(2),
    totalRating
  ];
  
  summarySheet.getRange(row, 1, 1, totalSummaryRow.length).setValues([totalSummaryRow]);
  summarySheet.getRange(row, 1, 1, 2).merge();
  summarySheet.getRange(row, 1, 1, totalSummaryRow.length).setBackground('#D1FAE5');
  summarySheet.getRange(row, 1, 1, totalSummaryRow.length).setFontWeight('bold');
  row += 2;
  
  // เกณฑ์การประเมิน
  row++;
  summarySheet.getRange(row, 1).setValue('เกณฑ์การประเมินระดับความพึงพอใจ');
  summarySheet.getRange(row, 1, 1, 3).merge();
  summarySheet.getRange(row, 1).setFontWeight('bold').setBackground('#F3F4F6');
  row++;
  
  summarySheet.getRange(row, 1).setValue('ค่าเฉลี่ย');
  summarySheet.getRange(row, 2).setValue('ความหมาย');
  summarySheet.getRange(row, 1, 1, 2).setFontWeight('bold').setBackground('#F3F4F6');
  row++;
  
  const criteriaRows = [
    ['4.51 - 5.00', 'มากที่สุด'],
    ['3.51 - 4.50', 'มาก'],
    ['2.51 - 3.50', 'ปานกลาง'],
    ['1.51 - 2.50', 'น้อย'],
    ['1.00 - 1.50', 'น้อยที่สุด']
  ];
  
  for (const criteriaRow of criteriaRows) {
    summarySheet.getRange(row, 1, 1, 2).setValues([criteriaRow]);
    row++;
  }
  
  // จัดรูปแบบตารางโดยรวม
  const totalRange = summarySheet.getRange(1, 1, row - 1, 7);
  totalRange.setBorder(true, true, true, true, true, true);
  
  // ปรับความกว้างคอลัมน์
  summarySheet.setColumnWidth(1, 180);
  summarySheet.setColumnWidth(2, 300);
  for (let i = 3; i <= 7; i++) {
    summarySheet.setColumnWidth(i, 120);
  }
  
  // ปรับ format ของตัวเลข
  summarySheet.getRange(5, 5, row - 5, 1).setNumberFormat('0.00');
  summarySheet.getRange(5, 6, row - 5, 1).setNumberFormat('0.00');
  
  return {
    success: true,
    url: spreadsheet.getUrl() + '#gid=' + summarySheet.getSheetId()
  };
}

// ฟังก์ชันเพิ่มข้อมูลและสร้างรายงานสรุปทันที
function submitSurveyAndCreateReport(formData) {
  const result = submitSurvey(formData);
  if (result.success) {
    try {
      const reportResult = createSummaryReport();
      if (reportResult.success) {
        result.reportUrl = reportResult.url;
      }
    } catch (error) {
      console.error('Error creating report:', error);
    }
  }
  return result;
}

function createEmptyEvaluationSheet() {
  const formData = {
    gender: "",
    age: "",
    education: "",
    occupation: "",
    suggestions: ""
  };
  
  return submitSurvey(formData);
}