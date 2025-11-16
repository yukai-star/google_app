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
