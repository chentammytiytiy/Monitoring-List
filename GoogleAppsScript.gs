// ============================================================
// Google Apps Script - 監控申請表單後端 (v6 - POST+GET 混合版)
// GET  → 讀取資料 / 小資料寫入 (向下相容)
// POST → 大資料寫入 / 追加資料 (解決 URL 長度限制)
// ============================================================

var SPREADSHEET_ID = '1suou6jcWJTuUh_mbKp9c9nlpoNtQSyqxxxvxgHgtpX8';

/**
 * GET 請求：
 *   - 無參數 → 讀取資料
 *   - ?action=save&data=... → 寫入（小資料向下相容）
 */
function doGet(e) {
  if (e.parameter && e.parameter.action === 'save') {
    return handleSave(e.parameter.data || '');
  }
  return handleRead();
}

/**
 * POST 請求：
 *   Content-Type: application/json
 *   Body: { "action": "save", "data": [...] }
 *         { "action": "append", "data": [...] }  ← 只追加，不覆蓋
 */
function doPost(e) {
  try {
    var jsonString = '';

    // 優先從 postData.contents 讀取（application/json body）
    if (e.postData && e.postData.contents) {
      jsonString = e.postData.contents;
    } else if (e.parameter && e.parameter.data) {
      // 相容 form-encoded 方式
      jsonString = e.parameter.data;
    }

    if (!jsonString) {
      return jsonResponse({ error: '未收到資料' });
    }

    var payload = JSON.parse(jsonString);

    // 支援 { action, data } 結構
    if (payload && typeof payload === 'object' && !Array.isArray(payload)) {
      var action = payload.action || 'save';
      var records = payload.data;

      if (action === 'append') {
        return handleAppend(records);
      } else {
        // action === 'save' 或預設：整批覆蓋
        return handleSaveArray(records);
      }
    }

    // 向下相容：直接傳陣列
    if (Array.isArray(payload)) {
      return handleSaveArray(payload);
    }

    return jsonResponse({ error: '格式錯誤' });
  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

/** 整批覆蓋寫入 */
function handleSave(jsonString) {
  try {
    if (!jsonString) {
      return jsonResponse({ error: '未收到資料' });
    }
    var records = JSON.parse(jsonString);
    if (!Array.isArray(records)) {
      return jsonResponse({ error: '格式錯誤，應為陣列' });
    }
    return handleSaveArray(records);
  } catch (err) {
    Logger.log('handleSave error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

/** 實際儲存陣列到 A1 */
function handleSaveArray(records) {
  try {
    if (!Array.isArray(records)) {
      return jsonResponse({ error: '格式錯誤，應為陣列' });
    }
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheets()[0];
    sheet.clearContents();
    var cell = sheet.getRange('A1');
    cell.setNumberFormat('@');
    var jsonString = JSON.stringify(records);
    cell.setValue(jsonString);
    Logger.log('handleSaveArray: 儲存 ' + records.length + ' 筆');
    return jsonResponse({ success: true, count: records.length });
  } catch (err) {
    Logger.log('handleSaveArray error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

/**
 * append 模式：讀取現有資料，合併後再寫入
 * 以 _id 為 key 去重（相同 _id 則更新，否則新增）
 */
function handleAppend(newRecords) {
  try {
    if (!Array.isArray(newRecords) || newRecords.length === 0) {
      return jsonResponse({ error: '無新增資料' });
    }

    // 讀取現有資料
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheets()[0];
    var val = sheet.getRange('A1').getValue();
    var existing = [];
    if (val) {
      try {
        existing = JSON.parse(val.toString().trim());
        if (!Array.isArray(existing)) existing = [];
      } catch(e) {
        existing = [];
      }
    }

    // 合併：以 _id 去重
    newRecords.forEach(function(r) {
      var idx = r._id ? existing.findIndex(function(h) { return h._id && h._id === r._id; }) : -1;
      if (idx !== -1) {
        existing[idx] = r; // 更新
      } else {
        existing.push(r); // 新增
      }
    });

    return handleSaveArray(existing);
  } catch (err) {
    Logger.log('handleAppend error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

/** 讀取 A1 的 JSON */
function handleRead() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheets()[0];
    var val = sheet.getRange('A1').getValue();
    var jsonStr = val ? val.toString().trim() : '';

    if (!jsonStr) {
      return ContentService.createTextOutput('[]')
        .setMimeType(ContentService.MimeType.JSON);
    }
    JSON.parse(jsonStr); // 驗證格式
    return ContentService.createTextOutput(jsonStr)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('handleRead error: ' + err.toString());
    return ContentService.createTextOutput('[]')
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/** 回傳 JSON 回應 */
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// 測試用函數
// ============================================================
function testWrite() {
  var testData = [{ "活動名稱": "TEST測試", "_id": "test001", "申請時間": "2026/03/13" }];
  var result = handleSaveArray(testData);
  Logger.log('testWrite result: ' + result.getContent());
}

function testAppend() {
  var newData = [{ "活動名稱": "APPEND測試", "_id": "test002", "申請時間": "2026/03/17" }];
  var result = handleAppend(newData);
  Logger.log('testAppend result: ' + result.getContent());
}
