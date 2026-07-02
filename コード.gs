const SS_ID = PropertiesService.getScriptProperties().getProperty('Spsheet_ID');

function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('冷凍機械 運転日報')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 初期データ取得（マスタと担当者）
function getInitialData() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    
    const masterSheet = ss.getSheetByName('マスタ');
    if(!masterSheet) return {master:[], staff:[], error:"マスタシートなし"};
    const masterData = masterSheet.getDataRange().getDisplayValues();
    if (masterData.length > 0) masterData.shift();

    const staffSheet = ss.getSheetByName('担当者マスタ');
    const staffData = staffSheet ? staffSheet.getDataRange().getDisplayValues() : [];
    if (staffData.length > 0) staffData.shift();
    
    return { master: masterData, staff: staffData };
  } catch (e) {
    return { master: [], staff: [], error: e.message };
  }
}

// ログイン判定
function checkLogin(id, password) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('事業所マスタ');
  const data = sheet.getDataRange().getDisplayValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id && String(data[i][2]) === password) {
      return { success: true, siteName: data[i][1],targetSsId:data[i][4]};
    }
  }
  return { success: false };
}

// 事業所リスト取得
function getSiteList() {
 try{ 
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('事業所マスタ');
  const data = sheet.getDataRange().getDisplayValues();
  data.shift();
  return data.map(row => ({ id:String(row[0]),name:String(row[1])}));
}catch(e){
  console.error("getSiteListエラー:"+e.message);
  return[];
  }
}

// ==========================================
// 🌟 各工場の「保存用スプシ」を狙い撃ちで開く処理
// （引数の最後に targetSsId を追加して動的に切り替えます）
// ==========================================

// 既存データを取得する処理（エリア別・フォーム反映用）
function getDailyData(dateStr, siteName, building, area, targetSsId) {
  try {
    const ss = SpreadsheetApp.openById(targetSsId);
    const sheet = ss.getSheetByName('データ保存');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if(!row[0]) continue;
      
      let dStr = "";
      if (row[0] instanceof Date) {
         dStr = Utilities.formatDate(row[0], Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
         dStr = String(row[0]).replace(/\//g, '-');
      }
      
      // 🌟前後の見えない空白などが原因でデータが一致しないのを防ぐため、全てtrim()で空白除去
      if (dStr === dateStr && 
          String(row[1]).trim() === String(siteName).trim() && 
          String(row[4]).trim() === String(building).trim() && 
          String(row[5]).trim() === String(area).trim()) {
          
        const eqName = String(row[6]).replace(/^'/, '').trim(); 
        result.push({
          no: eqName,
          item: String(row[7]).trim(), 
          value: String(row[8]).trim(), 
          weather: String(row[3]).trim(),
          staff: String(row[2]).trim() 
        });
      }
    }
    return result;
  } catch(e) {
    console.error("getDailyDataエラー:", e);
    return [];
  }
}

// ==========================================
// 追加: 既存データを取得する処理（棟全体・未入力バッジ計算用）
// ==========================================
function getBuildingData(dateStr, siteName, building, targetSsId) {
  try {
    const ss = SpreadsheetApp.openById(targetSsId); // 🌟各工場のスプシを開く
    const sheet = ss.getSheetByName('データ保存');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if(!row[0]) continue;
      
      let dStr = "";
      if (row[0] instanceof Date) {
         dStr = Utilities.formatDate(row[0], Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
         dStr = String(row[0]).replace(/\//g, '-');
      }
      
      if (dStr === dateStr && row[1] === siteName && row[4] === building) {
        const eqName = String(row[6]).replace(/^'/, '');
        result.push({
          area: String(row[5]),
          no: eqName,
          item: String(row[7]),
          value: String(row[8]),
          weather: String(row[3]),
          staff:String(row[2])
        });
      }
    }
    return result;
  } catch(e) {
    console.error(e);
    return [];
  }
}

// データ保存（上書き処理対応）
function saveData(results, targetSsId) {
  try {
    const ss = SpreadsheetApp.openById(targetSsId); // 🌟各工場のスプシを開く
    const sheet = ss.getSheetByName('データ保存');
    const now = new Date();
    
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    
    const rowMap = new Map();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if(!row[0]) continue;
      
      let dateStr = "";
      if (row[0] instanceof Date) {
         dateStr = Utilities.formatDate(row[0], Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
         dateStr = String(row[0]).replace(/\//g, '-');
      }
      
      const eqName = String(row[6]).replace(/^'/, '');
      const key = `${dateStr}|${row[1]}|${row[4]}|${row[5]}|${eqName}|${row[7]}`;
      rowMap.set(key, i + 1);
    }

    const newRows = [];
    results.forEach(res => {
      const key = `${res.date}|${res.site}|${res.building}|${res.area}|${res.no}|${res.item}`;
      
      if (rowMap.has(key)) {
        const rowNum = rowMap.get(key);
        sheet.getRange(rowNum, 1, 1, 11).setValues([[
          res.date,      // A: 点検日
          res.site,      // B: 事業所
          res.staff,     // C: 担当者
          res.weather,   // D: 天候
          res.building,  // E: 棟
          res.area,      // F: エリア
          "'" + res.no,  // G: 設備名
          res.item,      // H: 項目
          res.value,     // I: 値
          now,            // J: 書込時間
          `=A${rowNum}&G${rowNum}&H${rowNum}` //検索値
        ]]);
      } else {
        newRows.push([
          res.date,
          res.site,
          res.staff,
          res.weather,
          res.building,
          res.area,
          "'" + res.no,
          res.item,
          res.value,
          now,
          ""
        ]);
      }
    });

    if (newRows.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      for (let i = 0; i < newRows.length; i++) {
        const rowNum = startRow + i;
        newRows[i][10] = `=A${rowNum}&G${rowNum}&H${rowNum}`;
      }
      sheet.getRange(startRow, 1, newRows.length, 11).setValues(newRows);
    }
    
    return "報告データの保存（上書き）が完了しました！";
  } catch (e) {
    return "保存エラー: " + e.message;
  }
}

// 各種日報・一覧表データの取得
function getReportTableData(dateStr, siteName,targetSsId) {
    return getSpreadsheetTableData('運転日報', dateStr, 'reportTableContent', 'reportTableSpinner', targetSsId);
}

function getListTableData(dateStr, siteName, sheetName, targetSsId) {
  return getSpreadsheetTableData(sheetName, dateStr, 'listTableContent', 'listTableSpinner', targetSsId);
}

// 各工場のシートからデータを引っ張って計算・表示用データを返す
function getSpreadsheetTableData(sheetName, targetDateStr, containerId, spinnerId, targetSsId) {
  try {
    const ss = SpreadsheetApp.openById(targetSsId); // 🌟各工場のスプシを開く
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return { success: false, error: "シート「" + sheetName + "」が見つかりません", containerId: containerId, spinnerId: spinnerId };
    }
    
    if (targetDateStr) {
      const formattedDate = targetDateStr.replace(/-/g,'/');
      if (sheetName === '運転日報') {
        sheet.getRange("B2").setValue(formattedDate);
      } else if (sheetName === '大伸運輸') {
        sheet.getRange("A1").setValue(formattedDate);
      } else if (sheetName.includes('一覧表')) {
        sheet.getRange("A1").setValue(formattedDate);
      }
      SpreadsheetApp.flush();
    }
    
    const displayValues = sheet.getDataRange().getDisplayValues();
    return { success: true, data: displayValues, containerId: containerId, spinnerId: spinnerId };
  } catch(e) {
    return { success: false, error: e.toString(), containerId: containerId, spinnerId: spinnerId };
  }
}
