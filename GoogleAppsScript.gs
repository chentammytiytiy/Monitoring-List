// ============================================================
// Google Apps Script - 監控申請表單後端 (v7 - 表格格式版)
// GET  → 讀取資料（回傳 JSON 陣列）
// POST → 寫入 / 追加資料（以表格欄列方式儲存）
// 第1列：欄位標題；第2列起：資料內容
// ============================================================

var SPREADSHEET_ID = '1suou6jcWJTuUh_mbKp9c9nlpoNtQSyqxxxvxgHgtpX8';

/**
 * GET 請求：
 *   - 無參數 → 讀取資料
 *   - ?action=save&data=... → 寫入（小資料向下相容）
 */

// ============================================================
// Helper to get the main sheet (ignoring config sheets)
// ============================================================
function getMainSheet(ss) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name !== '排隊設定' && name !== '帳號') {
      return sheets[i];
    }
  }
  return sheets[0];
}

function doGet(e) {
  if (e.parameter && e.parameter.action === 'get_users') {
    return handleGetUsers();
  }
  if (e.parameter && e.parameter.action === 'get_queue') {
    return handleGetQueue();
  }
  if (e.parameter && e.parameter.action === 'save') {
    return handleSave(e.parameter.data || '');
  }
  return handleRead();
}

/**
 * POST 請求：
 *   Content-Type: application/json
 *   Body: { "action": "save",   "data": [...] }  ← 整批覆蓋
 *         { "action": "append", "data": [...] }  ← 合併（_id 去重）
 */
function doPost(e) {
  try {
    var jsonString = '';

    if (e.postData && e.postData.contents) {
      jsonString = e.postData.contents;
    } else if (e.parameter && e.parameter.data) {
      jsonString = e.parameter.data;
    }

    if (!jsonString) {
      return jsonResponse({ error: '未收到資料' });
    }

    var payload = JSON.parse(jsonString);

    if (payload && typeof payload === 'object' && !Array.isArray(payload)) {
      var action  = payload.action || 'save';
      var records = payload.data;
      if (action === 'append') {
        return handleAppend(records);
      } else if (action === 'save_queue') {
        return handleSaveQueue(records);
      } else if (action === 'log_account') {
        return handleLogAccount(records);
      } else if (action === 'sync_users') {
        return handleSyncUsers(records);
      } else {
        return handleSaveArray(records);
      }
    }

    if (Array.isArray(payload)) {
      return handleSaveArray(payload);
    }

    return jsonResponse({ error: '格式錯誤' });
  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

// ============================================================
// 寫入相關函數
// ============================================================

/**
 * 紀錄帳號操作
 */
function handleLogAccount(records) {
  try {
    if (!Array.isArray(records)) {
      records = [records];
    }
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetName = '帳號紀錄';
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName, ss.getSheets().length);
      sheet.appendRow(['建立時間', '動作', '操作者', '目標帳號', '目標姓名', '目標角色', '權限']);
      sheet.setFrozenRows(1);
    }
    
    records.forEach(function(r) {
      var time = r.time || new Date().toISOString();
      var action = r.action || '新增';
      var operator = r.operator || '';
      var targetUser = r.targetUser || '';
      var targetName = r.targetName || '';
      var targetRole = r.targetRole || '';
      var perms = r.permissions ? r.permissions.join(', ') : '';
      sheet.appendRow([time, action, operator, targetUser, targetName, targetRole, perms]);
    });
    
    return jsonResponse({ success: true, count: records.length });
  } catch (err) {
    Logger.log('handleLogAccount error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

/**
 * 寫入系統帳號（整批覆寫）
 */
function handleSyncUsers(records) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('系統帳號');
    if (!sheet) {
      sheet = ss.insertSheet('系統帳號');
    }
    sheet.clearContents();
    if (records.length === 0) return jsonResponse({ success: true, count: 0 });
    
    var headers = ['username', 'password', 'name', 'role', 'permissions', 'enabled', 'createdAt'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    var rows = records.map(function(r) {
      return [
        r.username || '',
        r.password || '',
        r.name || '',
        r.role || '',
        r.permissions ? JSON.stringify(r.permissions) : '[]',
        r.enabled !== false,
        r.createdAt || new Date().toISOString()
      ];
    });
    
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.setFrozenRows(1);
    
    return jsonResponse({ success: true, count: records.length });
  } catch(err) {
    Logger.log('handleSyncUsers error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

/** GET save 向下相容 */
function handleSave(jsonString) {
  try {
    if (!jsonString) return jsonResponse({ error: '未收到資料' });
    var records = JSON.parse(jsonString);
    if (!Array.isArray(records)) return jsonResponse({ error: '格式錯誤，應為陣列' });
    return handleSaveArray(records);
  } catch (err) {
    Logger.log('handleSave error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

/**
 * 整批覆蓋寫入（以表格格式）
 * - 第1列：所有欄位名稱（union of all keys）
 * - 第2列起：各筆資料
 */
function handleSaveArray(records) {
  try {
    if (!Array.isArray(records)) {
      return jsonResponse({ error: '格式錯誤，應為陣列' });
    }

    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = getMainSheet(ss);
    sheet.clearContents();

    if (records.length === 0) {
      Logger.log('handleSaveArray: 清空完成（0 筆）');
      return jsonResponse({ success: true, count: 0 });
    }

    // 收集所有欄位（保留順序，_id 優先排首）
    var headers = [];
    records.forEach(function(r) {
      Object.keys(r).forEach(function(k) {
        if (headers.indexOf(k) === -1) headers.push(k);
      });
    });
    // 確保 _id 在第一欄
    var idIdx = headers.indexOf('_id');
    if (idIdx > 0) {
      headers.splice(idIdx, 1);
      headers.unshift('_id');
    }

    // 寫入標題列
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // 建立資料矩陣
    var rows = records.map(function(r) {
      return headers.map(function(h) {
        var v = r[h];
        return (v !== undefined && v !== null) ? String(v) : '';
      });
    });

    // 寫入資料（從第2列開始）
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

    // 凍結標題列
    sheet.setFrozenRows(1);

    Logger.log('handleSaveArray: 儲存 ' + records.length + ' 筆 × ' + headers.length + ' 欄');
    return jsonResponse({ success: true, count: records.length, columns: headers.length });
  } catch (err) {
    Logger.log('handleSaveArray error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

function handleAppend(newRecords) {
  try {
    if (!Array.isArray(newRecords) || newRecords.length === 0) {
      return jsonResponse({ error: '無新增資料' });
    }

    var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = getMainSheet(ss);
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    // 如果表格為空，直接用 handleSaveArray 寫入
    if (lastRow < 2) {
      return handleSaveArray(newRecords);
    }

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var idIndex = headers.indexOf('_id');

    if (idIndex === -1) {
      // 若無 _id，退回舊的寫入模式
      var existing = readTableAsArray();
      newRecords.forEach(function(r) {
        existing.push(r);
      });
      return handleSaveArray(existing);
    }

    // 取得所有現有 _id
    var idValues = sheet.getRange(2, idIndex + 1, lastRow - 1, 1).getValues().map(function(row) {
      return row[0] ? row[0].toString() : '';
    });

    newRecords.forEach(function(r) {
      var targetId = r._id ? r._id.toString() : '';
      var rowIdx = targetId ? idValues.indexOf(targetId) : -1;

      var rowData = headers.map(function(h) {
        var v = r[h];
        return (v !== undefined && v !== null) ? String(v) : '';
      });

      if (rowIdx !== -1) {
        // 更新：將單列資料寫入對應範圍
        sheet.getRange(rowIdx + 2, 1, 1, headers.length).setValues([rowData]);
      } else {
        // 新增
        sheet.appendRow(rowData);
        idValues.push(targetId); // 預防同批有重複新增
      }
    });

    return jsonResponse({ success: true, count: newRecords.length });
  } catch (err) {
    Logger.log('handleAppend error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

// ============================================================
// 讀取相關函數
// ============================================================

/**
 * 讀取試算表 → 回傳 JSON 陣列
 * 自動相容舊版（A1 存放 JSON 字串）與新版（表格格式）
 */
function handleRead() {
  try {
    var records = readTableAsArray();
    return ContentService
      .createTextOutput(JSON.stringify(records))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('handleRead error: ' + err.toString());
    return ContentService.createTextOutput('[]')
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 核心讀取函數：將試算表轉為 JSON 物件陣列
 * - 新版：第1列為欄位名稱，第2列起為資料
 * - 舊版相容：若 A1 內容為 JSON 字串，自動解析並遷移
 */
function readTableAsArray() {
  var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getMainSheet(ss);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 1 || lastCol < 1) return [];

  var a1Value = sheet.getRange(1, 1).getValue();
  var a1Str   = a1Value ? a1Value.toString().trim() : '';

  // ── 舊版相容：A1 存的是 JSON 陣列字串 ──
  if (a1Str.charAt(0) === '[') {
    try {
      var oldData = JSON.parse(a1Str);
      if (Array.isArray(oldData) && oldData.length > 0) {
        Logger.log('readTableAsArray: 偵測到舊版 JSON 格式，自動遷移至表格格式');
        handleSaveArray(oldData); // 遷移寫回新格式
        return oldData;
      }
    } catch(e) {
      // 解析失敗，當作一般表格處理
    }
  }

  // ── 新版：表格格式 ──
  if (lastRow < 2) return []; // 只有標題列，無資料

  var headers  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var dataRows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  return dataRows
    .filter(function(row) {
      // 過濾掉完全空白的列
      return row.some(function(cell) { return cell !== '' && cell !== null; });
    })
    .map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) {
        obj[h] = (row[i] !== undefined && row[i] !== null) ? row[i].toString() : '';
      });
      return obj;
    });
}

/**
 * 讀取系統帳號
 */
function handleGetUsers() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('系統帳號');
    if (!sheet) return jsonResponse([]);
    
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return jsonResponse([]);
    
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    var users = data.map(function(row) {
      var u = {};
      headers.forEach(function(h, i) {
        if (h === 'permissions') {
          try { u[h] = JSON.parse(row[i]); } catch(e) { u[h] = []; }
        } else if (h === 'enabled') {
          u[h] = (row[i] === true || row[i] === 'TRUE' || row[i] === 'true');
        } else {
          u[h] = row[i];
        }
      });
      return u;
    });
    return jsonResponse(users);
  } catch(err) {
    Logger.log('handleGetUsers error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

// ============================================================
// 工具函數
// ============================================================

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// 測試用函數
// ============================================================
function testMigrate() {
  // 測試：讀取 → 若為舊版會自動遷移為表格格式
  var records = readTableAsArray();
  Logger.log('testMigrate: 讀取 ' + records.length + ' 筆');
  Logger.log(JSON.stringify(records.slice(0, 2)));
}

function testWrite() {
  var testData = [
    { "_id": "test001", "活動名稱": "TEST測試A", "申請時間": "2026/03/18", "負責業務": "小明" },
    { "_id": "test002", "活動名稱": "TEST測試B", "申請時間": "2026/03/18", "負責業務": "小華" }
  ];
  var result = handleSaveArray(testData);
  Logger.log('testWrite result: ' + result.getContent());
}

function testAppend() {
  var newData = [
    { "_id": "test003", "活動名稱": "APPEND測試C", "申請時間": "2026/03/18", "負責業務": "小李" }
  ];
  var result = handleAppend(newData);
  Logger.log('testAppend result: ' + result.getContent());
}
function handleGetQueue() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('排隊設定');
    if (!sheet) return jsonResponse([]);
    
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return jsonResponse([]);
    
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    var queues = data.map(function(row) {
      var u = {};
      headers.forEach(function(h, i) {
        u[h] = row[i];
      });
      return u;
    });
    return jsonResponse(queues);
  } catch(err) {
    Logger.log('handleGetQueue error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

function handleSaveQueue(records) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('排隊設定');
    var isNew = false;
    if (!sheet) {
      sheet = ss.insertSheet('排隊設定', ss.getSheets().length);
      isNew = true;
    }
    if (records.length === 0) return jsonResponse({ success: true, count: 0 });
    
    var headers = [];
    if (!isNew && sheet.getLastRow() > 0) {
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    } else {
      records.forEach(function(r) {
        Object.keys(r).forEach(function(k) {
          if (headers.indexOf(k) === -1) headers.push(k);
        });
      });
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
    
    var idIndex = headers.indexOf('_id');
    var existingIds = [];
    if (!isNew && sheet.getLastRow() > 1 && idIndex !== -1) {
      existingIds = sheet.getRange(2, idIndex + 1, sheet.getLastRow() - 1, 1).getValues().map(function(row) {
        return row[0] ? row[0].toString() : '';
      });
    }

    records.forEach(function(r) {
      var rowData = headers.map(function(h) {
        var v = r[h];
        return (v !== undefined && v !== null) ? String(v) : '';
      });
      
      var targetId = r._id ? r._id.toString() : '';
      var rowIdx = targetId ? existingIds.indexOf(targetId) : -1;
      
      if (rowIdx !== -1) {
        sheet.getRange(rowIdx + 2, 1, 1, headers.length).setValues([rowData]);
      } else {
        sheet.appendRow(rowData);
        if (targetId) existingIds.push(targetId);
      }
    });
    
    return jsonResponse({ success: true, count: records.length });
  } catch(err) {
    Logger.log('handleSaveQueue error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

function handleLogAccount(records) {
  try {
    if (!Array.isArray(records)) {
      records = [records];
    }
    
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheetName = '帳號紀錄';
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName, ss.getSheets().length);
      sheet.appendRow(['建立時間', '動作', '操作者', '目標帳號', '目標姓名', '目標角色', '權限']);
      sheet.setFrozenRows(1);
    }
    
    records.forEach(function(r) {
      var time = r.time || new Date().toISOString();
      var action = r.action || '新增';
      var operator = r.operator || '';
      var targetUser = r.targetUser || '';
      var targetName = r.targetName || '';
      var targetRole = r.targetRole || '';
      var perms = r.permissions ? r.permissions.join(', ') : '';
      sheet.appendRow([time, action, operator, targetUser, targetName, targetRole, perms]);
    });
    
    return jsonResponse({ success: true, count: records.length });
  } catch (err) {
    Logger.log('handleLogAccount error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

/**
 * 寫入系統帳號（整批覆寫）
 */
function handleSyncUsers(records) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('系統帳號');
    if (!sheet) {
      sheet = ss.insertSheet('系統帳號');
    }
    sheet.clearContents();
    if (records.length === 0) return jsonResponse({ success: true, count: 0 });
    
    var headers = ['username', 'password', 'name', 'role', 'permissions', 'enabled', 'createdAt'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    var rows = records.map(function(r) {
      return [
        r.username || '',
        r.password || '',
        r.name || '',
        r.role || '',
        r.permissions ? JSON.stringify(r.permissions) : '[]',
        r.enabled !== false,
        r.createdAt || new Date().toISOString()
      ];
    });
    
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.setFrozenRows(1);
    
    return jsonResponse({ success: true, count: records.length });
  } catch(err) {
    Logger.log('handleSyncUsers error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

/** GET save 向下相容 */
function handleSave(jsonString) {
  try {
    if (!jsonString) return jsonResponse({ error: '未收到資料' });
    var records = JSON.parse(jsonString);
    if (!Array.isArray(records)) return jsonResponse({ error: '格式錯誤，應為陣列' });
    return handleSaveArray(records);
  } catch (err) {
    Logger.log('handleSave error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

/**
 * 整批覆蓋寫入（以表格格式）
 * - 第1列：所有欄位名稱（union of all keys）
 * - 第2列起：各筆資料
 */
function handleSaveArray(records) {
  try {
    if (!Array.isArray(records)) {
      return jsonResponse({ error: '格式錯誤，應為陣列' });
    }

    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = getMainSheet(ss);
    sheet.clearContents();

    if (records.length === 0) {
      Logger.log('handleSaveArray: 清空完成（0 筆）');
      return jsonResponse({ success: true, count: 0 });
    }

    // 收集所有欄位（保留順序，_id 優先排首）
    var headers = [];
    records.forEach(function(r) {
      Object.keys(r).forEach(function(k) {
        if (headers.indexOf(k) === -1) headers.push(k);
      });
    });
    // 確保 _id 在第一欄
    var idIdx = headers.indexOf('_id');
    if (idIdx > 0) {
      headers.splice(idIdx, 1);
      headers.unshift('_id');
    }

    // 寫入標題列
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // 建立資料矩陣
    var rows = records.map(function(r) {
      return headers.map(function(h) {
        var v = r[h];
        return (v !== undefined && v !== null) ? String(v) : '';
      });
    });

    // 寫入資料（從第2列開始）
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

    // 凍結標題列
    sheet.setFrozenRows(1);

    Logger.log('handleSaveArray: 儲存 ' + records.length + ' 筆 × ' + headers.length + ' 欄');
    return jsonResponse({ success: true, count: records.length, columns: headers.length });
  } catch (err) {
    Logger.log('handleSaveArray error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

function handleAppend(newRecords) {
  try {
    if (!Array.isArray(newRecords) || newRecords.length === 0) {
      return jsonResponse({ error: '無新增資料' });
    }

    var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = getMainSheet(ss);
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();

    // 如果表格為空，直接用 handleSaveArray 寫入
    if (lastRow < 2) {
      return handleSaveArray(newRecords);
    }

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var idIndex = headers.indexOf('_id');

    if (idIndex === -1) {
      // 若無 _id，退回舊的寫入模式
      var existing = readTableAsArray();
      newRecords.forEach(function(r) {
        existing.push(r);
      });
      return handleSaveArray(existing);
    }

    // 取得所有現有 _id
    var idValues = sheet.getRange(2, idIndex + 1, lastRow - 1, 1).getValues().map(function(row) {
      return row[0] ? row[0].toString() : '';
    });

    newRecords.forEach(function(r) {
      var targetId = r._id ? r._id.toString() : '';
      var rowIdx = targetId ? idValues.indexOf(targetId) : -1;

      var rowData = headers.map(function(h) {
        var v = r[h];
        return (v !== undefined && v !== null) ? String(v) : '';
      });

      if (rowIdx !== -1) {
        // 更新：將單列資料寫入對應範圍
        sheet.getRange(rowIdx + 2, 1, 1, headers.length).setValues([rowData]);
      } else {
        // 新增
        sheet.appendRow(rowData);
        idValues.push(targetId); // 預防同批有重複新增
      }
    });

    return jsonResponse({ success: true, count: newRecords.length });
  } catch (err) {
    Logger.log('handleAppend error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

// ============================================================
// 讀取相關函數
// ============================================================

/**
 * 讀取試算表 → 回傳 JSON 陣列
 * 自動相容舊版（A1 存放 JSON 字串）與新版（表格格式）
 */
function handleRead() {
  try {
    var records = readTableAsArray();
    return ContentService
      .createTextOutput(JSON.stringify(records))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log('handleRead error: ' + err.toString());
    return ContentService.createTextOutput('[]')
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 核心讀取函數：將試算表轉為 JSON 物件陣列
 * - 新版：第1列為欄位名稱，第2列起為資料
 * - 舊版相容：若 A1 內容為 JSON 字串，自動解析並遷移
 */
function readTableAsArray() {
  var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getMainSheet(ss);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 1 || lastCol < 1) return [];

  var a1Value = sheet.getRange(1, 1).getValue();
  var a1Str   = a1Value ? a1Value.toString().trim() : '';

  // ── 舊版相容：A1 存的是 JSON 陣列字串 ──
  if (a1Str.charAt(0) === '[') {
    try {
      var oldData = JSON.parse(a1Str);
      if (Array.isArray(oldData) && oldData.length > 0) {
        Logger.log('readTableAsArray: 偵測到舊版 JSON 格式，自動遷移至表格格式');
        handleSaveArray(oldData); // 遷移寫回新格式
        return oldData;
      }
    } catch(e) {
      // 解析失敗，當作一般表格處理
    }
  }

  // ── 新版：表格格式 ──
  if (lastRow < 2) return []; // 只有標題列，無資料

  var headers  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var dataRows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  return dataRows
    .filter(function(row) {
      // 過濾掉完全空白的列
      return row.some(function(cell) { return cell !== '' && cell !== null; });
    })
    .map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) {
        obj[h] = (row[i] !== undefined && row[i] !== null) ? row[i].toString() : '';
      });
      return obj;
    });
}

/**
 * 讀取系統帳號
 */
function handleGetUsers() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('系統帳號');
    if (!sheet) return jsonResponse([]);
    
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return jsonResponse([]);
    
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    var users = data.map(function(row) {
      var u = {};
      headers.forEach(function(h, i) {
        if (h === 'permissions') {
          try { u[h] = JSON.parse(row[i]); } catch(e) { u[h] = []; }
        } else if (h === 'enabled') {
          u[h] = (row[i] === true || row[i] === 'TRUE' || row[i] === 'true');
        } else {
          u[h] = row[i];
        }
      });
      return u;
    });
    return jsonResponse(users);
  } catch(err) {
    Logger.log('handleGetUsers error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

// ============================================================
// 工具函數
// ============================================================

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// 測試用函數
// ============================================================
function testMigrate() {
  // 測試：讀取 → 若為舊版會自動遷移為表格格式
  var records = readTableAsArray();
  Logger.log('testMigrate: 讀取 ' + records.length + ' 筆');
  Logger.log(JSON.stringify(records.slice(0, 2)));
}

function testWrite() {
  var testData = [
    { "_id": "test001", "活動名稱": "TEST測試A", "申請時間": "2026/03/18", "負責業務": "小明" },
    { "_id": "test002", "活動名稱": "TEST測試B", "申請時間": "2026/03/18", "負責業務": "小華" }
  ];
  var result = handleSaveArray(testData);
  Logger.log('testWrite result: ' + result.getContent());
}

function testAppend() {
  var newData = [
    { "_id": "test003", "活動名稱": "APPEND測試C", "申請時間": "2026/03/18", "負責業務": "小李" }
  ];
  var result = handleAppend(newData);
  Logger.log('testAppend result: ' + result.getContent());
}
function handleGetQueue() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('排隊設定');
    if (!sheet) return jsonResponse([]);
    
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return jsonResponse([]);
    
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    var queues = data.map(function(row) {
      var u = {};
      headers.forEach(function(h, i) {
        u[h] = row[i];
      });
      return u;
    });
    return jsonResponse(queues);
  } catch(err) {
    Logger.log('handleGetQueue error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

function handleSaveQueue(records) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('排隊設定');
    var isNew = false;
    if (!sheet) {
      sheet = ss.insertSheet('排隊設定', ss.getSheets().length);
      isNew = true;
    }
    if (records.length === 0) return jsonResponse({ success: true, count: 0 });
    
    var headers = [];
    if (!isNew && sheet.getLastRow() > 0) {
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    } else {
      records.forEach(function(r) {
        Object.keys(r).forEach(function(k) {
          if (headers.indexOf(k) === -1) headers.push(k);
        });
      });
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
    
    var rows = records.map(function(r) {
      return headers.map(function(h) {
        var v = r[h];
        return (v !== undefined && v !== null) ? String(v) : '';
      });
    });
    
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
    
    return jsonResponse({ success: true, count: records.length });
  } catch(err) {
    Logger.log('handleSaveQueue error: ' + err.toString());
    return jsonResponse({ error: err.toString() });
  }
}

