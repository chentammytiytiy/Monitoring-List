import re

with open('GoogleAppsScript.gs', 'r', encoding='utf-8') as f:
    content = f.read()

# Replace the old handleSaveQueue using regex to be safe
old_func_pattern = r'function handleSaveQueue\(records\)\s*\{.*?\n\}'
# Actually regex with .*? can be tricky with newlines. Let's do it manually.

start_idx = content.find('function handleSaveQueue(records) {')
end_idx = content.find('function handleLogAccount(records) {')

if start_idx != -1 and end_idx != -1:
    old_str = content[start_idx:end_idx]
    
    new_handle_save_queue = '''function handleSaveQueue(records) {
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

'''
    # We replace exactly the slice
    content = content[:start_idx] + new_handle_save_queue + content[end_idx:]
    with open('GoogleAppsScript.gs', 'w', encoding='utf-8') as f:
        f.write(content)
    print("Replaced!")
else:
    print("Not found", start_idx, end_idx)
