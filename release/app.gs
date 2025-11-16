// ---------- UI / Sidebar ----------
function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('點名系統')
    .addItem('開啟點名面板','showSidebar')
    .addToUi();
}

function showSidebar(){
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('廣青雲端全廣大課點名系統')
    .setWidth(800);
  SpreadsheetApp.getUi().showSidebar(html);
}

function doGet(e){
  return HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('廣青雲端全廣大課點名系統');
}

// ---------- 後端 API ----------

// 取得組別清單
function getGroups(){
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('學員名單資料_連動後台');
  if(!sh) return [];
  var vals = sh.getRange(2,1, sh.getLastRow()-1,1).getValues().flat();
  return Array.from(new Set(vals)).filter(String).sort();
}

// 取得月份清單
function getMonths(){
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('上課日期維護');
  if(!sh) return [];
  var lastCol = sh.getLastColumn();
  var months = sh.getRange(2,2,1,lastCol-1).getValues()[0];
  return months.filter(String);
}

// 取得學生名單
function getStudentsByGroup(group){
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('學員名單資料_連動後台');
  if(!sh) return [];
  var data = sh.getRange(2,1,sh.getLastRow()-1,4).getValues();
  return data.filter(r=>r[0]+''===group+'')
             .map(r=>({id:r[1]+'', name:r[2]+'', email:r[3]+''}));
}


// 取得學生名單 V2
function getStudentsByGroup_v2(group){
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('學員名單資料_連動後台');
  if(!sh) return [];
  
  // 取得所有資料
  var data = sh.getDataRange().getValues();
  if(data.length < 2) return []; // 至少要有標題行和一行資料
  
  // 找出標題行中各欄位的位置（更靈活）
  var headers = data[0];
  var groupCol = headers.indexOf('組別') >= 0 ? headers.indexOf('組別') : 0;
  var idCol = headers.indexOf('學籍編號') >= 0 ? headers.indexOf('學籍編號') : 1;
  var nameCol = headers.indexOf('姓名') >= 0 ? headers.indexOf('姓名') : 2;
  var emailCol = headers.indexOf('電子郵件') >= 0 ? headers.indexOf('電子郵件') : 3;
  
  // 篩選和轉換資料
  var students = [];
  for(var i = 1; i < data.length; i++) {
    var row = data[i];
    if(row[groupCol] && row[groupCol].toString() === group.toString()) {
      students.push({
        id: row[idCol] ? row[idCol].toString() : '',
        name: row[nameCol] ? row[nameCol].toString() : '',
        email: row[emailCol] ? row[emailCol].toString() : ''
      });
    }
  }
  
  return students;
}

// 取得既有出席記錄 - 適應實際工作表格式
function getExistingAttendance(group, month){
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('出席紀錄彙總');
    
    if(!sheet) {
      console.log('出席紀錄彙總工作表不存在');
      return [];
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if(lastRow <= 3) { // 至少要有3行（總數、標題、星期）
      console.log('出席紀錄彙總工作表無足夠資料');
      return [];
    }
    
    // 取得日期標題行（第2行，從D欄開始）
    const dateHeaders = sheet.getRange(2, 4, 1, lastCol - 3).getValues()[0];
    console.log('日期標題:', dateHeaders);
    
    // 取得學生資料（從第4行開始）
    const studentData = sheet.getRange(4, 1, lastRow - 3, lastCol).getValues();
    
    console.log(`查詢組別: ${group}, 月份: ${month}`);
    
    const records = [];
    
    // 處理每個學生的出席記錄
    studentData.forEach(row => {
      const studentGroup = row[0] ? row[0].toString() : '';
      const studentId = row[1] ? row[1].toString() : '';
      const studentName = row[2] ? row[2].toString() : '';
      
      // 檢查是否為目標組別的學生
      if(studentGroup !== group) {
        return; // 跳過不屬於此組的學生
      }
      
      // 處理該學生的各日期出席記錄
      dateHeaders.forEach((dateHeader, dateIndex) => {
        if(!dateHeader) return; // 跳過空日期
        
        // 轉換日期格式
        const dateStr = dateHeader.toString();
        
        // 檢查是否為目標月份的日期
        // 支援 "10/01" 格式和其他可能的格式
        const isTargetMonth = dateStr.startsWith(month + '/') || 
                             dateStr.includes('/' + month + '/') ||
                             (dateStr.includes('/') && dateStr.split('/')[0] === month);
        
        if(!isTargetMonth) {
          return; // 跳過非目標月份的日期
        }
        
        // 取得出席狀態值（從D欄開始，所以是 dateIndex + 3）
        const statusValue = row[dateIndex + 3];
        
        // 轉換狀態值
        let status = '';
        if(statusValue === 0 || statusValue === '0') status = '請假';
        else if(statusValue === 1 || statusValue === '1') status = '出席';
        else if(statusValue === 2 || statusValue === '2') status = '補課';
        
        // 只有有狀態值才加入記錄
        if(status) {
          records.push({
            studentId: studentId,
            date: dateStr,
            status: status
          });
        }
      });
    });
    
    console.log(`找到 ${records.length} 筆既有記錄 (${group}-${month})`);
    return records;
    
  } catch (error) {
    console.error('getExistingAttendance 錯誤:', error);
    return [];
  }
}
// 取得該月份上課日期
function getDatesByMonth(month){
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('上課日期維護');
  if(!sh) return [];
  var lastCol = sh.getLastColumn();
  var months = sh.getRange(2,2,1,lastCol-1).getValues()[0];
  var idx = months.findIndex(m=>m+''===month+'');
  if(idx===-1) return [];
  var col = 2 + idx; // B起算
  return sh.getRange(3,col,14,1).getValues().flat().filter(String);
}

// 儲存點名回「出席紀錄彙總」
function saveAttendance(payload){
  if(!payload || !payload.records) return {success:false,message:'payload empty'};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summary = ss.getSheetByName('出席紀錄彙總');
  if(!summary) return {success:false,message:'找不到出席紀錄彙總分頁'};
  
  var allDates = summary.getRange(2,3,1,summary.getLastColumn()-2).getValues()[0];
  var idRange = summary.getRange(4,1,summary.getLastRow()-3,1).getValues().flat();
  var idToRow = {};
  idRange.forEach((id,i)=>{ if(id) idToRow[id+'']=4+i; });

  var valMap = {'請假':0,'出席':1,'補課':2,'': ''};
  payload.records.forEach(rec=>{
    var sid = rec.studentId+'';
    var dt = rec.date+'';
    var targetRow = idToRow[sid]; if(!targetRow) return;
    var colIdx = allDates.findIndex(d=>d+''===dt);
    if(colIdx===-1) return;
    summary.getRange(targetRow, 3+colIdx).setValue(valMap[rec.status]!==undefined?valMap[rec.status]:'');
  });

  return {success:true,message:'已回填 '+payload.records.length+' 筆資料'};
}



function saveAttendance_v2(payload){
    try {
      console.log('收到儲存請求:', payload);
      
      if(!payload || !payload.records) {
        return {success: false, message: 'payload empty'};
      }
      
      const {group, month, records} = payload;
      
      // 取得或建立儲存工作表
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName('出席紀錄彙總');
      
      if(!sheet) {
        sheet = ss.insertSheet('出席紀錄彙總');
        // 建立標題行
        sheet.getRange(1, 1, 1, 5).setValues([['組別', '月份', '學籍編號', '日期', '狀態']]);
      }
      
      // 取得現有資料
      let existingData = [];
      if(sheet.getLastRow() > 0) {
        existingData = sheet.getDataRange().getValues();
      }
      
      // 確保標題行存在
      if(existingData.length === 0) {
        existingData = [['組別', '月份', '學籍編號', '日期', '狀態']];
      }
      
      // 清除該組別該月份的舊記錄，保留其他記錄
      const filteredData = [existingData[0]]; // 保留標題行
      
      for(let i = 1; i < existingData.length; i++) {
        const row = existingData[i];
        // 確保每列只有5個欄位，且不是要刪除的記錄
        if(row.length >= 2 && (row[0] !== group || row[1] !== month)) {
          // 只取前5個欄位，防止資料異常
          filteredData.push(row.slice(0, 5));
        }
      }
      
      // 添加新記錄
      records.forEach(record => {
        // 確保每筆記錄都是5個欄位
        const newRow = [
          group || '',
          month || '',
          record.studentId || '',
          record.date || '',
          record.status || ''
        ];
        filteredData.push(newRow);
      });
      
      // 清除整個工作表
      sheet.clear();
      
      // 寫入資料（確保所有列都是5欄）
      if(filteredData.length > 0) {
        // 驗證所有列都是5欄
        const cleanedData = filteredData.map(row => {
          if(Array.isArray(row)) {
            // 確保每列都是5欄，不足的補空字串，多的截斷
            const cleanRow = [];
            for(let i = 0; i < 5; i++) {
              cleanRow[i] = (row[i] !== undefined && row[i] !== null) ? row[i].toString() : '';
            }
            return cleanRow;
          }
          return ['', '', '', '', '']; // 防止非陣列資料
        });
        
        console.log('準備寫入的資料:', cleanedData);
        console.log('資料行數:', cleanedData.length, '欄數:', cleanedData[0].length);
        
        sheet.getRange(1, 1, cleanedData.length, 5).setValues(cleanedData);
      }
      
      return {
        success: true, 
        message: `成功儲存 ${records.length} 筆出席記錄`
      };
      
    } catch (error) {
      console.error('saveAttendance_v2 錯誤:', error);
      return {
        success: false, 
        message: '儲存失敗: ' + error.message
      };
    }
}

function updateAttendanceSummary(updateData) {
  try {
    console.log('開始更新出席紀錄彙總 (僅更新出席記錄):', updateData);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('出席紀錄彙總');
    
    // 如果工作表不存在，直接回傳錯誤
    if (!sheet) {
      return {
        success: false,
        message: '出席紀錄彙總工作表不存在，請先手動建立工作表結構'
      };
    }
    
    const { group, month, students, dates, attendanceGrid } = updateData;
    
    // 🎯 1. 只讀取現有日期標題 (完全不修改)
    console.log('讀取現有日期標題...');
    
    const lastCol = sheet.getLastColumn();
    let existingDates = [];
    
    if (lastCol >= 4) {
      const existingDateRange = sheet.getRange(2, 4, 1, lastCol - 3);
      existingDates = existingDateRange.getValues()[0].filter(date => date && date.toString().trim() !== '');
    }
    
    console.log('現有日期標題:', existingDates);
    console.log('前端傳入的日期:', dates);
    
    // 建立日期對應表 (前端日期 -> 工作表欄位位置)
    const dateColumnMap = {};
    existingDates.forEach((existingDate, index) => {
      const existingDateStr = existingDate.toString();
      
      // 尋找前端日期中匹配的項目
      const matchedFrontendDate = dates.find(frontendDate => {
        return frontendDate === existingDateStr || 
               frontendDate.replace(/^20\d{2}\//, '') === existingDateStr; // 處理 "2025/10/1" vs "10/1"
      });
      
      if (matchedFrontendDate) {
        dateColumnMap[matchedFrontendDate] = 4 + index; // D欄開始
        console.log(`日期對應: 前端"${matchedFrontendDate}" -> 工作表第${4 + index}欄"${existingDateStr}"`);
      }
    });
    
    console.log('日期欄位對應表:', dateColumnMap);
    
    // 🎯 2. 只讀取現有學生資料 (完全不修改學生資訊)
    console.log(`只更新組別 "${group}" 的出席記錄...`);
    
    const lastRow = sheet.getLastRow();
    
    if (lastRow < 4) {
      return {
        success: false,
        message: '工作表中沒有學生資料，請先手動建立學生資料'
      };
    }
    
    // 取得現有的所有學生資料
    const existingRange = sheet.getRange(4, 1, lastRow - 3, Math.max(lastCol, 3 + existingDates.length));
    const existingData = existingRange.getValues();
    
    console.log(`讀取到 ${existingData.length} 行既有學生資料`);
    
    let updatedRecordCount = 0;
    let processedStudents = [];
    
    // 🎯 只更新出席記錄，完全不碰學生資料 (A, B, C 欄)
    students.forEach((student, studentIndex) => {
      console.log(`處理學生: ${student.id} (${student.name}) - 組別: ${group}`);
      
      // 在現有資料中找到該學生
      let targetRowIndex = -1;
      for (let i = 0; i < existingData.length; i++) {
        const existingGroup = existingData[i][0] ? existingData[i][0].toString() : '';
        const existingStudentId = existingData[i][1] ? existingData[i][1].toString() : '';
        
        // 必須同時匹配組別和學生ID
        if (existingGroup === group && existingStudentId === student.id) {
          targetRowIndex = i;
          break;
        }
      }
      
      if (targetRowIndex === -1) {
        console.warn(`找不到學生: ${student.id} (組別: ${group})`);
        return; // 跳過不存在的學生
      }
      
      processedStudents.push(student.id);
      
      // 🎯 只更新該學生的出席記錄 (D欄以後)
      let studentUpdatedCount = 0;
      attendanceGrid[studentIndex].forEach((value, dateIndex) => {
        const frontendDate = dates[dateIndex];
        const targetColumn = dateColumnMap[frontendDate];
        
        if (targetColumn) {
          const actualRowIndex = targetRowIndex + 4; // 轉換為實際行號
          
          // 直接更新工作表中的單一儲存格
          sheet.getRange(actualRowIndex, targetColumn).setValue(value);
          studentUpdatedCount++;
          
          if (value !== '') {
            console.log(`  更新記錄: ${student.id} ${frontendDate} = ${value} (第${actualRowIndex}行第${targetColumn}欄)`);
            updatedRecordCount++;
          }
        } else {
          console.warn(`  找不到對應欄位: ${student.id} ${frontendDate}`);
        }
      });
      
      console.log(`  學生 ${student.id} 更新了 ${studentUpdatedCount} 個日期的記錄`);
    });
    
    const result = {
      success: true,
      message: `成功更新 ${group} 組 ${month} 月出席記錄！共更新 ${updatedRecordCount} 筆記錄，處理 ${processedStudents.length} 位學生`,
      details: {
        group: group,
        month: month,
        studentsProcessed: processedStudents.length,
        recordsUpdated: updatedRecordCount,
        datesMatched: Object.keys(dateColumnMap).length,
        existingDates: existingDates.length,
        processedStudents: processedStudents
      }
    };
    
    console.log('更新完成:', result);
    return result;
    
  } catch (error) {
    console.error('updateAttendanceSummary 錯誤:', error);
    return {
      success: false,
      message: '更新出席紀錄彙總失敗: ' + error.message,
      error: error.toString()
    };
  }
}

function updateAttendanceSummary_optimized(updateData) {
  try {
    console.log('開始更新出席紀錄彙總 (批量更新優化版):', updateData);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('出席紀錄彙總');
    
    if (!sheet) {
      return {
        success: false,
        message: '出席紀錄彙總工作表不存在，請先手動建立工作表結構'
      };
    }
    
    const { group, month, students, dates, attendanceGrid } = updateData;
    
    // 1. 讀取現有日期標題
    const lastCol = sheet.getLastColumn();
    let existingDates = [];
    
    if (lastCol >= 4) {
      const existingDateRange = sheet.getRange(2, 4, 1, lastCol - 3);
      existingDates = existingDateRange.getValues()[0].filter(date => date && date.toString().trim() !== '');
    }
    
    // 建立日期對應表
    const dateColumnMap = {};
    existingDates.forEach((existingDate, index) => {
      const existingDateStr = existingDate.toString();
      const matchedFrontendDate = dates.find(frontendDate => {
        return frontendDate === existingDateStr || 
               frontendDate.replace(/^20\d{2}\//, '') === existingDateStr;
      });
      
      if (matchedFrontendDate) {
        dateColumnMap[matchedFrontendDate] = 4 + index;
      }
    });
    
    // 2. 讀取現有學生資料
    const lastRow = sheet.getLastRow();
    if (lastRow < 4) {
      return {
        success: false,
        message: '工作表中沒有學生資料，請先手動建立學生資料'
      };
    }
    
    const existingRange = sheet.getRange(4, 1, lastRow - 3, Math.max(lastCol, 3 + existingDates.length));
    const existingData = existingRange.getValues();
    
    // 3. 準備批量更新資料
    const updatesData = [];
    let processedStudents = [];
    
    students.forEach((student, studentIndex) => {
      // 找到該學生在工作表中的行號
      let targetRowIndex = -1;
      for (let i = 0; i < existingData.length; i++) {
        const existingGroup = existingData[i][0] ? existingData[i][0].toString() : '';
        const existingStudentId = existingData[i][1] ? existingData[i][1].toString() : '';
        
        if (existingGroup === group && existingStudentId === student.id) {
          targetRowIndex = i;
          break;
        }
      }
      
      if (targetRowIndex === -1) {
        console.warn(`找不到學生: ${student.id} (組別: ${group})`);
        return;
      }
      
      processedStudents.push(student.id);
      
      // 收集該學生的所有更新
      attendanceGrid[studentIndex].forEach((value, dateIndex) => {
        const frontendDate = dates[dateIndex];
        const targetColumn = dateColumnMap[frontendDate];
        
        if (targetColumn) {
          const actualRowIndex = targetRowIndex + 4; // 轉換為實際行號
          
          updatesData.push({
            row: actualRowIndex,
            col: targetColumn,
            value: value
          });
        }
      });
    });
    
    // 4. 🚀 批量更新 - 按範圍分組更新
    if (updatesData.length > 0) {
      // 將更新按行分組
      const rowGroups = {};
      updatesData.forEach(update => {
        if (!rowGroups[update.row]) {
          rowGroups[update.row] = {};
        }
        rowGroups[update.row][update.col] = update.value;
      });
      
      // 批量更新每一行
      Object.keys(rowGroups).forEach(row => {
        const rowNum = parseInt(row);
        const colUpdates = rowGroups[row];
        
        // 找出該行的最小和最大欄位
        const cols = Object.keys(colUpdates).map(c => parseInt(c)).sort((a, b) => a - b);
        const minCol = cols[0];
        const maxCol = cols[cols.length - 1];
        
        // 讀取該行的現有資料
        const currentRowData = sheet.getRange(rowNum, minCol, 1, maxCol - minCol + 1).getValues()[0];
        
        // 更新需要變更的儲存格
        cols.forEach(col => {
          const colIndex = col - minCol;
          currentRowData[colIndex] = colUpdates[col];
        });
        
        // 一次性寫入整行
        sheet.getRange(rowNum, minCol, 1, maxCol - minCol + 1).setValues([currentRowData]);
      });
    }
    
    const result = {
      success: true,
      message: `成功更新 ${group} 組 ${month} 月出席記錄！共更新 ${updatesData.length} 筆記錄，處理 ${processedStudents.length} 位學生`,
      details: {
        group: group,
        month: month,
        studentsProcessed: processedStudents.length,
        recordsUpdated: updatesData.length,
        datesMatched: Object.keys(dateColumnMap).length,
        processedStudents: processedStudents
      }
    };
    
    console.log('批量更新完成:', result);
    return result;
    
  } catch (error) {
    console.error('updateAttendanceSummary_optimized 錯誤:', error);
    return {
      success: false,
      message: '更新出席紀錄彙總失敗: ' + error.message,
      error: error.toString()
    };
  }
}



// --------------------------------------
// 測試函數
// --------------------------------------

// 取得學生名單 V2 測試選取特定組別名單
function testGetStudentsByGroup() {
  var group = 'B02'; 
  var result = getStudentsByGroup_v2(group);
  console.log(group + '組學生數量:', result.length);
  console.log('學生資料:', result);
  return result;
}


function testGetDatesByMonth() {
  var month = '10';
  var result = getDatesByMonth(month);
  console.log(month + '月上課日期:', result);
  return result;
}

// 測試取得既有出席記錄 - 針對實際工作表格式
function testGetExistingAttendance() {
  var group = 'A01'; // 測試 A01 組別
  var month = '10';  // 測試10月份
  
  console.log(`測試取得既有出席記錄 - 組別: ${group}, 月份: ${month}`);
  
  var result = getExistingAttendance(group, month);
  
  console.log('找到的記錄數量:', result.length);
  
  if(result.length > 0) {
    console.log('範例記錄:');
    result.slice(0, 15).forEach((record, index) => { // 顯示前15筆
      console.log(`  ${index + 1}. 學籍編號: ${record.studentId}, 日期: ${record.date}, 狀態: ${record.status}`);
    });
  } else {
    console.log('沒有找到任何記錄');
  }
  
  return result;
}


// 檢查工作表實際結構
function checkActualSheetStructure() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('出席紀錄彙總');
    
    if(!sheet) {
      console.log('❌ 出席紀錄彙總工作表不存在');
      return;
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    console.log(`工作表大小: ${lastRow} 行 x ${lastCol} 欄`);
    
    // 檢查第1行 (總數資訊)
    if(lastRow >= 1) {
      const row1 = sheet.getRange(1, 1, 1, Math.min(5, lastCol)).getValues()[0];
      console.log('第1行 (總數):', row1);
    }
    
    // 檢查第2行 (日期標題行)
    if(lastRow >= 2) {
      const row2 = sheet.getRange(2, 1, 1, Math.min(10, lastCol)).getValues()[0];
      console.log('第2行 (日期標題):', row2);
    }
    
    // 檢查第3行 (星期)
    if(lastRow >= 3) {
      const row3 = sheet.getRange(3, 1, 1, Math.min(10, lastCol)).getValues()[0];
      console.log('第3行 (星期):', row3);
    }
    
    // 檢查前幾個學生資料
    if(lastRow >= 4) {
      const studentRows = sheet.getRange(4, 1, Math.min(5, lastRow - 3), Math.min(8, lastCol)).getValues();
      console.log('學生資料範例:');
      studentRows.forEach((row, index) => {
        console.log(`  學生${index + 1}: 組別=${row[0]}, 學籍編號=${row[1]}, 姓名=${row[2]}, 出席狀態=${row.slice(3, 7)}`);
      });
    }
    
    // 檢查 A01 組別的學生
    console.log('\n=== A01 組別學生 ===');
    if(lastRow >= 4) {
      const allStudents = sheet.getRange(4, 1, lastRow - 3, 3).getValues();
      const a01Students = allStudents.filter(row => row[0] === 'A01');
      console.log('A01 組學生數量:', a01Students.length);
      a01Students.forEach((student, index) => {
        console.log(`  ${index + 1}. ${student[1]} - ${student[2]}`);
      });
    }
    
  } catch (error) {
    console.error('檢查工作表結構錯誤:', error);
  }
}

// 測試多個組別的既有記錄
function testMultipleGroupsAttendance() {
  var testCases = [
    {group: 'A01', month: '10'},
    {group: 'B01', month: '10'},
    {group: 'B02', month: '10'},
    {group: 'A01', month: '11'}
  ];
  
  testCases.forEach(testCase => {
    console.log(`\n=== 測試 ${testCase.group} 組 ${testCase.month} 月 ===`);
    var result = getExistingAttendance(testCase.group, testCase.month);
    console.log(`記錄數量: ${result.length}`);
    
    if(result.length > 0) {
      var students = [...new Set(result.map(r => r.studentId))];
      var dates = [...new Set(result.map(r => r.date))];
      console.log(`學生數: ${students.length}, 日期數: ${dates.length}`);
    }
  });
}


function testAttendanceSheetStructure_v2() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('出席紀錄彙總');
    
    if(!sheet) {
      console.log('❌ 出席紀錄彙總工作表不存在');
      return false;
    }
    
    console.log('✅ 出席紀錄彙總工作表存在');
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    console.log(`工作表大小: ${lastRow} 行 x ${lastCol} 欄`);
    
    if(lastRow >= 1) {
      // 檢查第1行 (學員總數資訊)
      const row1 = sheet.getRange(1, 1, 1, Math.min(10, lastCol)).getValues()[0];
      console.log('第1行 (學員總數):', row1);
    }
    
    if(lastRow >= 2) {
      // 檢查第2行 (組別、學籍編號、姓名、日期標題)
      const row2 = sheet.getRange(2, 1, 1, Math.min(15, lastCol)).getValues()[0];
      console.log('第2行 (標題行):', row2.slice(0, 10), '...'); // 只顯示前10個
      
      // 分析日期標題
      const dateHeaders = row2.slice(3); // 從D欄開始是日期
      const validDates = dateHeaders.filter(d => d && d.toString().includes('/'));
      console.log(`共有 ${validDates.length} 個日期欄位`);
      console.log('前5個日期:', validDates.slice(0, 5));
    }
    
    if(lastRow >= 3) {
      // 檢查第3行 (星期資訊)
      const row3 = sheet.getRange(3, 1, 1, Math.min(15, lastCol)).getValues()[0];
      console.log('第3行 (星期):', row3.slice(0, 10), '...'); // 只顯示前10個
    }
    
    if(lastRow >= 4) {
      // 檢查學生資料
      const studentRows = sheet.getRange(4, 1, Math.min(10, lastRow - 3), Math.min(10, lastCol)).getValues();
      console.log('\n學生資料範例:');
      studentRows.forEach((row, index) => {
        const group = row[0] || '';
        const studentId = row[1] || '';
        const studentName = row[2] || '';
        const attendanceData = row.slice(3, 8); // 前5個出席資料
        console.log(`  ${index + 1}. 組別:${group}, 學號:${studentId}, 姓名:${studentName}, 出席:${attendanceData}`);
      });
      
      // 統計各組別學生數量
      console.log('\n=== 組別統計 ===');
      const allStudents = sheet.getRange(4, 1, lastRow - 3, 3).getValues();
      const groupStats = {};
      
      allStudents.forEach(row => {
        const group = row[0] ? row[0].toString() : '';
        if(group) {
          groupStats[group] = (groupStats[group] || 0) + 1;
        }
      });
      
      Object.keys(groupStats).forEach(group => {
        console.log(`  ${group} 組: ${groupStats[group]} 位學生`);
      });
      
      // 檢查 A01 組的詳細資料
      console.log('\n=== A01 組詳細資料 ===');
      const a01Students = allStudents.filter(row => row[0] === 'A01');
      console.log(`A01 組共 ${a01Students.length} 位學生:`);
      a01Students.forEach((student, index) => {
        console.log(`  ${index + 1}. ${student[1]} - ${student[2]}`);
      });
    }
    
    return true;
    
  } catch (error) {
    console.error('檢查工作表結構時發生錯誤:', error);
    return false;
  }
}

// 專門測試日期欄位的函數
function testDateColumns() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('出席紀錄彙總');
    
    if(!sheet) {
      console.log('❌ 出席紀錄彙總工作表不存在');
      return;
    }
    
    const lastCol = sheet.getLastColumn();
    
    // 取得第2行的日期標題
    const dateHeaders = sheet.getRange(2, 4, 1, lastCol - 3).getValues()[0];
    
    console.log('=== 日期欄位分析 ===');
    console.log(`總共 ${dateHeaders.length} 個日期欄位`);
    
    // 分析各月份的日期
    const monthGroups = {};
    dateHeaders.forEach((date, index) => {
      if(date && date.toString().includes('/')) {
        const dateStr = date.toString();
        const month = dateStr.split('/')[0];
        if(!monthGroups[month]) monthGroups[month] = [];
        monthGroups[month].push({date: dateStr, colIndex: index + 4});
      }
    });
    
    Object.keys(monthGroups).forEach(month => {
      console.log(`\n${month}月份:`, monthGroups[month].length, '個日期');
      monthGroups[month].forEach(item => {
        console.log(`  ${item.date} (第${item.colIndex}欄)`);
      });
    });
    
  } catch (error) {
    console.error('測試日期欄位錯誤:', error);
  }
}

// 測試特定學生的出席記錄
function testStudentAttendance() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('出席紀錄彙總');
    
    if(!sheet) {
      console.log('❌ 出席紀錄彙總工作表不存在');
      return;
    }
    
    const studentId = 'A250101'; // 測試這位學生
    console.log(`=== 測試學生 ${studentId} 的出席記錄 ===`);
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    // 找到該學生的行
    const allStudents = sheet.getRange(4, 1, lastRow - 3, lastCol).getValues();
    const studentRow = allStudents.find(row => row[1] === studentId);
    
    if(!studentRow) {
      console.log(`找不到學生 ${studentId}`);
      return;
    }
    
    console.log(`學生資料: 組別=${studentRow[0]}, 學號=${studentRow[1]}, 姓名=${studentRow[2]}`);
    
    // 取得日期標題
    const dateHeaders = sheet.getRange(2, 4, 1, lastCol - 3).getValues()[0];
    
    // 顯示該學生的出席記錄
    console.log('\n出席記錄:');
    dateHeaders.forEach((date, index) => {
      if(date && date.toString().includes('/')) {
        const status = studentRow[index + 3];
        let statusText = '';
        if(status === 0) statusText = '請假';
        else if(status === 1) statusText = '出席';
        else if(status === 2) statusText = '補課';
        else statusText = '未填';
        
        if(statusText !== '未填') {
          console.log(`  ${date}: ${statusText}`);
        }
      }
    });
    
  } catch (error) {
    console.error('測試學生出席記錄錯誤:', error);
  }
}

// 綜合測試函數 - 針對實際工作表格式
function runAttendanceTests_v2() {
  console.log('='.repeat(60));
  console.log('開始執行出席記錄相關測試 (適用於實際工作表格式)');
  console.log('='.repeat(60));
  
  // 1. 檢查工作表結構
  console.log('\n1. 檢查工作表結構');
  console.log('-'.repeat(30));
  testAttendanceSheetStructure_v2();
  
  // 2. 測試日期欄位
  console.log('\n2. 測試日期欄位');
  console.log('-'.repeat(30));
  testDateColumns();
  
  // 3. 測試特定學生記錄
  console.log('\n3. 測試特定學生記錄');
  console.log('-'.repeat(30));
  testStudentAttendance();
  
  // 4. 測試 getExistingAttendance 函數
  console.log('\n4. 測試取得既有記錄函數');
  console.log('-'.repeat(30));
  testGetExistingAttendance();
  
  console.log('\n' + '='.repeat(60));
  console.log('測試完成');
  console.log('='.repeat(60));
}

// 測試 getExistingAttendance 函數
function testGetExistingAttendanceDebug() {
  var group = 'A01';
  var month = '10';
  
  console.log('=== 詳細除錯 getExistingAttendance ===');
  console.log(`測試參數: 組別=${group}, 月份=${month}`);
  
  var result = getExistingAttendance(group, month);
  
  console.log('回傳結果:');
  console.log('記錄數量:', result.length);
  
  if(result.length > 0) {
    console.log('前5筆記錄:');
    result.slice(0, 5).forEach((record, index) => {
      console.log(`  ${index + 1}. 學號: ${record.studentId}, 日期: ${record.date}, 狀態: ${record.status}`);
    });
  }
  
  return result;
}