/**
 * ------------------------------------------------------------------
 * Multi-Model API Key Test Console v1.02
 * Force Teaching Edition
 * Google Sheet Setup Script
 * ------------------------------------------------------------------
 * Author: Force Cheng
 * Framework: FALO (Formosa AI Life Outlook)
 * License: Internal Knowledge Asset / Educational Use
 * Copyright: © Force Cheng × FALO 2026/04/01. All Rights Reserved.
 * ------------------------------------------------------------------
 * 用途說明
 * ------------------------------------------------------------------
 * 這份檔案專門負責初始化 Google Sheet 的教學資料。
 *
 * 會建立或重建三個工作表：
 * - README
 * - users
 * - logs
 * - forced_output
 * - runtime_config
 *
 * 預設資料包含：
 * - users：admin / 123、user / 123
 * - logs：一筆示範紀錄
 * - README：操作說明、部署建議、環境變數說明
 * - forced_output：前言與結語
 * - runtime_config：後端即時 systemPrompt 與 outputCharLimit 設定
 *
 * 使用方式：
 * 1. 將此檔與 Code.gs 放在同一個 Apps Script 專案
 * 2. 專案需綁定在目標 Google Sheet 上
 * 3. 執行 setupTeachingSheets()
 * 4. 第一次執行會要求授權
 * ------------------------------------------------------------------
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("API Key Teach")
    .addItem("初始化教學資料表", "setupTeachingSheets")
    .addToUi();
}

function setupTeachingSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("找不到綁定中的 Google Sheet，請確認此專案是綁定式 Apps Script。");
  }

  var inspection = inspectTeachingSheets_(ss);
  if (inspection.valid) {
    SpreadsheetApp.getUi().alert("資料表結構檢查完成：README / users / logs / forced_output / runtime_config 結構正確，已保留既有資料。");
    return;
  }

  var readmeSheet = resetSheet_(ss, README_SHEET_NAME || "README");
  var usersSheet = resetSheet_(ss, USERS_SHEET_NAME || "users");
  var logsSheet = resetSheet_(ss, LOGS_SHEET_NAME || "logs");
  var forcedOutputSheet = resetSheet_(ss, FORCED_OUTPUT_SHEET_NAME || "forced_output");
  var runtimeConfigSheet = resetSheet_(ss, RUNTIME_CONFIG_SHEET_NAME || "runtime_config");

  setupReadmeSheet_(readmeSheet);
  setupUsersSheet_(usersSheet);
  setupLogsSheet_(logsSheet);
  setupForcedOutputSheet_(forcedOutputSheet);
  setupRuntimeConfigSheet_(runtimeConfigSheet);

  SpreadsheetApp.getUi().alert(
    "偵測到資料表結構不完整或內容錯誤，已清空並重建 README / users / logs / forced_output / runtime_config。\n\n原因：\n- " +
    inspection.reasons.join("\n- ")
  );
}

function inspectTeachingSheets_(ss) {
  var reasons = [];

  validateReadmeSheet_(ss, reasons);
  validateUsersSheet_(ss, reasons);
  validateLogsSheet_(ss, reasons);
  validateForcedOutputSheet_(ss, reasons);
  validateRuntimeConfigSheet_(ss, reasons);

  return {
    valid: reasons.length === 0,
    reasons: reasons
  };
}

function resetSheet_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.clear();
  sheet.clearFormats();
  sheet.clearNotes();
  sheet.clearConditionalFormatRules();

  return sheet;
}

function setupReadmeSheet_(sheet) {
  var values = [
    ["Multi-Model API Key Test Console v1.02", "Force Teaching Edition"],
    ["檔案用途", "這是一份教學用 Google Sheet 範例。可搭配 Apps Script 做成超簡單的 API Key 測試、登入驗證、角色分流、訊息回傳與紀錄保存範例。"],
    ["工作表說明", "README：操作說明 / users：帳號密碼與角色 / logs：輸入輸出紀錄"],
    ["建議使用方式", "1. 先把這份 Google Sheet 準備好\n2. 將 Apps Script 綁定在這份 Google Sheet\n3. 把 Code.gs 與 SheetSetup.gs 貼到 Apps Script 專案\n4. 先執行 setupTeachingSheets()\n5. 到 Script Properties 設定三組 API Key 與兩組驗證機制\n6. 部署為 Web App"],
    ["建議環境變數", "GEMINI_API_KEY / MINIMAX_API_KEY / KIMI_API_KEY / SHARE_EXEC_CODE / SHARE_EXEC_URL / ADMIN_PASSWORD_CODE"],
    ["設定位置", "Apps Script > 專案設定 > 指令碼屬性（Script Properties）"],
    ["模型預設規則", "程式預設使用 gemini；若前端有傳 aiModel，則會以使用者選擇覆蓋。"],
    ["後端即時設定來源", "systemPrompt 與 outputCharLimit 會先從 runtime_config 工作表讀取。若前端有傳入同名欄位，則以前端值覆蓋。"],
    ["runtime_config 預設 systemPrompt", "請用清楚、簡短、教學導向的方式回覆，並保留足夠資訊讓學生理解系統運作。"],
    ["runtime_config 預設輸出限制", "outputCharLimit 預設為 500 字。"],
    ["帳號角色規則", "只有在 Use Server Key 模式下才套用角色限制。admin 為管理者帳號，可取得完整教學版回覆；user 為一般使用者帳號，進入受限教學回覆，不產生完整回答。若使用 Use My API Key，則不受 admin / user 限制。"],
    ["users sheet 用途", "存放簡單教學版帳號密碼與角色。預設帳號：admin / 123、user / 123。這些角色限制只在 Use Server Key 模式下生效；管理者密碼回填功能也會從這張表讀取 admin 的目前密碼。"],
    ["logs sheet 用途", "記錄使用者輸入、系統輸出、時間、模型、金鑰來源與其他教學資訊。"],
    ["forced_output sheet 用途", "用來控制強制輸出的前言與結語，讓學生觀察 Google Sheet 內容如何影響回覆結果。"],
    ["runtime_config sheet 用途", "用來控制後端即時 systemPrompt 與 outputCharLimit，方便學生直接從 Google Sheet 調整系統限制。"],
    ["分享功能說明", "A 機制：設定 SHARE_EXEC_CODE 與 SHARE_EXEC_URL，前端輸入第一組驗證碼後，後端驗證成功才回傳範例 exec 網址並回填欄位。B 機制：設定 ADMIN_PASSWORD_CODE，前端輸入第二組驗證碼後，後端驗證成功才會從 users sheet 讀取 admin 的目前密碼並回填。"],
    ["重要提醒", "這是教學版範例，不建議直接用於正式產品。正式環境請至少加入密碼雜湊、權限控管與錯誤處理。"],
    ["對應 Code.gs 固定工作表名稱", "README / users / logs / forced_output / runtime_config"],
    ["台北時間格式", "yyyy-MM-dd HH:mm:ss"],
    ["版權宣告", "© Force Cheng × FALO 2026/04/01. All Rights Reserved."],
    ["SheetSetup.gs 檢查規則", "SheetSetup.gs 會先檢查目前 Google Sheet 是否存在 README / users / logs / forced_output / runtime_config 五張工作表，並確認必要表頭與欄位順序是否正確。"],
    ["何時會重建", "若缺少工作表、表頭錯誤、欄位順序錯誤，或必要預設資料缺失到無法使用，則會清空既有資料並重新建立完整結構。"],
    ["何時保留資料", "若五張工作表結構正確，SheetSetup.gs 就不主動清除既有資料，避免把已經可用的教學紀錄洗掉。"],
    ["教學用途說明", "這種做法適合 Excel 上傳後轉存為 Google Sheet 的教學流程。即使使用者手動改壞 sheet，也能透過初始化腳本快速修復到可用狀態。"]
  ];

  sheet.getRange(1, 1, values.length, 2).setValues(values);
  sheet.getRange("A1:B1").merge();
  sheet.getRange("A1").setFontWeight("bold").setFontSize(14).setBackground("#1F4E78").setFontColor("#FFFFFF").setHorizontalAlignment("center");
  sheet.getRange(2, 1, values.length - 1, 1).setFontWeight("bold").setBackground("#DCE6F1");
  sheet.getRange(2, 2, values.length - 1, 1).setBackground("#F7FBFF");
  sheet.getRange(1, 1, values.length, 2).setBorder(true, true, true, true, true, true, "#C9D2E3", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(1, 1, values.length, 2).setWrap(true).setVerticalAlignment("top");
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 760);
  sheet.setFrozenRows(1);
}

function setupUsersSheet_(sheet) {
  var headers = [["username", "password", "role", "status", "display_name", "created_at", "notes"]];
  var rows = [
    ["admin", "123", "admin", "active", "System Admin", "2026-04-07 16:10:00", "管理者預設帳號，在 Use Server Key 模式下可取得完整教學版回覆"],
    ["user", "123", "user", "active", "General User", "2026-04-07 16:12:00", "一般使用者預設帳號，在 Use Server Key 模式下進入受限教學回覆"]
  ];

  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  styleTable_(sheet, 1, headers[0].length, 2);
  setColumnWidths_(sheet, [140, 140, 110, 100, 180, 180, 220]);
}

function setupLogsSheet_(sheet) {
  var headers = [[
    "log_id", "timestamp_taipei", "username", "action", "user_input",
    "system_output", "ai_model", "key_source", "success", "client_ip",
    "user_agent", "notes"
  ]];

  var rows = [[
    "LOG-0001",
    "2026-04-07 16:10:00",
    "admin",
    "chat",
    "Please reply with: API connection successful.",
    "API connection successful. 回應時間（台北）：2026-04-07 16:10:00",
    "gemini",
    "server",
    "TRUE",
    "127.0.0.1",
    "Teaching Sample Browser",
    "這是一筆示範資料，可保留也可刪除。"
  ]];

  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  styleTable_(sheet, 1, headers[0].length, 2);
  setColumnWidths_(sheet, [110, 180, 120, 100, 260, 320, 110, 110, 90, 120, 180, 220]);
}

function setupForcedOutputSheet_(sheet) {
  var headers = [["key", "content", "enabled", "notes"]];
  var rows = [
    ["preface", "【教學版前言】這是示範用回覆，用來幫助學生理解 GAS 如何組合輸出內容。", "TRUE", "會加在回覆最前面"],
    ["conclusion", "【教學版結語】你可以修改 forced_output 工作表，觀察回覆文字如何一起改變。", "TRUE", "會加在回覆最後面"]
  ];

  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  styleTable_(sheet, 1, headers[0].length, rows.length + 1);
  setColumnWidths_(sheet, [120, 520, 90, 220]);
}

function setupRuntimeConfigSheet_(sheet) {
  var headers = [["key", "value", "enabled", "notes"]];
  var rows = [
    ["system_prompt", "請用清楚、簡短、教學導向的方式回覆，並保留足夠資訊讓學生理解系統運作。", "TRUE", "後端預設 systemPrompt"],
    ["output_char_limit", "500", "TRUE", "後端預設輸出字數限制"]
  ];

  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  styleTable_(sheet, 1, headers[0].length, rows.length + 1);
  setColumnWidths_(sheet, [160, 520, 90, 220]);
}

function styleTable_(sheet, headerRow, totalColumns, totalRows) {
  sheet.getRange(headerRow, 1, 1, totalColumns)
    .setFontWeight("bold")
    .setBackground("#1F4E78")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");

  sheet.getRange(headerRow, 1, Math.max(totalRows, 2), totalColumns)
    .setBorder(true, true, true, true, true, true, "#C9D2E3", SpreadsheetApp.BorderStyle.SOLID)
    .setWrap(true)
    .setVerticalAlignment("top");

  sheet.setFrozenRows(1);
}

function setColumnWidths_(sheet, widths) {
  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
}

function validateReadmeSheet_(ss, reasons) {
  var sheet = ss.getSheetByName(README_SHEET_NAME || "README");
  if (!sheet) {
    reasons.push("缺少 README 工作表");
    return;
  }

  var expectedLabels = getReadmeLabels_();
  var actualTitle = String(sheet.getRange("A1").getDisplayValue() || "").trim();
  if (actualTitle !== "Multi-Model API Key Test Console v1.02") {
    reasons.push("README 標題不正確");
    return;
  }

  var actualLabels = sheet.getRange(2, 1, expectedLabels.length, 1).getDisplayValues()
    .map(function(row) { return String(row[0] || "").trim(); });

  for (var i = 0; i < expectedLabels.length; i++) {
    if (actualLabels[i] !== expectedLabels[i]) {
      reasons.push("README 結構或欄位順序不正確");
      return;
    }
  }
}

function validateUsersSheet_(ss, reasons) {
  var sheet = ss.getSheetByName(USERS_SHEET_NAME || "users");
  if (!sheet) {
    reasons.push("缺少 users 工作表");
    return;
  }

  var expectedHeaders = getUsersHeaders_();
  var actualHeaders = sheet.getRange(1, 1, 1, expectedHeaders.length).getDisplayValues()[0]
    .map(function(value) { return String(value || "").trim(); });

  if (!arraysEqual_(expectedHeaders, actualHeaders)) {
    reasons.push("users 表頭或欄位順序不正確");
    return;
  }

  var values = sheet.getDataRange().getDisplayValues();
  var hasDefaultAdmin = values.some(function(row, index) {
    return index > 0 &&
      String(row[0] || "").trim() === "admin" &&
      String(row[1] || "").trim() === "123";
  });
  var hasDefaultUser = values.some(function(row, index) {
    return index > 0 &&
      String(row[0] || "").trim() === "user" &&
      String(row[1] || "").trim() === "123" &&
      String(row[2] || "").trim().toLowerCase() === "user";
  });

  if (!hasDefaultAdmin) {
    reasons.push("users 缺少預設 admin / 123 帳號");
  }

  if (!hasDefaultUser) {
    reasons.push("users 缺少預設 user / 123 帳號");
  }
}

function validateLogsSheet_(ss, reasons) {
  var sheet = ss.getSheetByName(LOGS_SHEET_NAME || "logs");
  if (!sheet) {
    reasons.push("缺少 logs 工作表");
    return;
  }

  var expectedHeaders = getLogsHeaders_();
  var actualHeaders = sheet.getRange(1, 1, 1, expectedHeaders.length).getDisplayValues()[0]
    .map(function(value) { return String(value || "").trim(); });

  if (!arraysEqual_(expectedHeaders, actualHeaders)) {
    reasons.push("logs 表頭或欄位順序不正確");
  }
}

function validateForcedOutputSheet_(ss, reasons) {
  var sheet = ss.getSheetByName(FORCED_OUTPUT_SHEET_NAME || "forced_output");
  if (!sheet) {
    reasons.push("缺少 forced_output 工作表");
    return;
  }

  var expectedHeaders = getForcedOutputHeaders_();
  var actualHeaders = sheet.getRange(1, 1, 1, expectedHeaders.length).getDisplayValues()[0]
    .map(function(value) { return String(value || "").trim(); });

  if (!arraysEqual_(expectedHeaders, actualHeaders)) {
    reasons.push("forced_output 表頭或欄位順序不正確");
    return;
  }

  var values = sheet.getDataRange().getDisplayValues();
  var hasPreface = false;
  var hasConclusion = false;

  for (var i = 1; i < values.length; i++) {
    var key = String(values[i][0] || "").trim().toLowerCase();
    if (key === "preface") hasPreface = true;
    if (key === "conclusion") hasConclusion = true;
  }

  if (!hasPreface || !hasConclusion) {
    reasons.push("forced_output 缺少 preface 或 conclusion 預設資料");
  }
}

function validateRuntimeConfigSheet_(ss, reasons) {
  var sheet = ss.getSheetByName(RUNTIME_CONFIG_SHEET_NAME || "runtime_config");
  if (!sheet) {
    reasons.push("缺少 runtime_config 工作表");
    return;
  }

  var expectedHeaders = getRuntimeConfigHeaders_();
  var actualHeaders = sheet.getRange(1, 1, 1, expectedHeaders.length).getDisplayValues()[0]
    .map(function(value) { return String(value || "").trim(); });

  if (!arraysEqual_(expectedHeaders, actualHeaders)) {
    reasons.push("runtime_config 表頭或欄位順序不正確");
    return;
  }

  var values = sheet.getDataRange().getDisplayValues();
  var hasSystemPrompt = false;
  var hasOutputCharLimit = false;

  for (var i = 1; i < values.length; i++) {
    var key = String(values[i][0] || "").trim().toLowerCase();
    if (key === "system_prompt") hasSystemPrompt = true;
    if (key === "output_char_limit") hasOutputCharLimit = true;
  }

  if (!hasSystemPrompt || !hasOutputCharLimit) {
    reasons.push("runtime_config 缺少 system_prompt 或 output_char_limit 預設資料");
  }
}

function getReadmeLabels_() {
  return [
    "檔案用途",
    "工作表說明",
    "建議使用方式",
    "建議環境變數",
    "設定位置",
    "模型預設規則",
    "後端即時設定來源",
    "runtime_config 預設 systemPrompt",
    "runtime_config 預設輸出限制",
    "帳號角色規則",
    "users sheet 用途",
    "logs sheet 用途",
    "forced_output sheet 用途",
    "runtime_config sheet 用途",
    "分享功能說明",
    "重要提醒",
    "對應 Code.gs 固定工作表名稱",
    "台北時間格式",
    "版權宣告",
    "SheetSetup.gs 檢查規則",
    "何時會重建",
    "何時保留資料",
    "教學用途說明"
  ];
}

function getUsersHeaders_() {
  return ["username", "password", "role", "status", "display_name", "created_at", "notes"];
}

function getLogsHeaders_() {
  return [
    "log_id", "timestamp_taipei", "username", "action", "user_input",
    "system_output", "ai_model", "key_source", "success", "client_ip",
    "user_agent", "notes"
  ];
}

function getForcedOutputHeaders_() {
  return ["key", "content", "enabled", "notes"];
}

function getRuntimeConfigHeaders_() {
  return ["key", "value", "enabled", "notes"];
}

function arraysEqual_(a, b) {
  if (!a || !b || a.length !== b.length) return false;
  for (var i = 0; i < a.length; i++) {
    if (String(a[i] || "").trim() !== String(b[i] || "").trim()) {
      return false;
    }
  }
  return true;
}
