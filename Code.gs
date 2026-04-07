/**
 * ------------------------------------------------------------------
 * Multi-Model API Key Test Console v1.02
 * Force Teaching Edition
 * Minimal GAS Web App Sample
 * ------------------------------------------------------------------
 * Author: Force Cheng
 * Framework: FALO (Formosa AI Life Outlook)
 * License: Internal Knowledge Asset / Educational Use
 * Copyright: © Force Cheng × FALO 2026/04/01. All Rights Reserved.
 * ------------------------------------------------------------------
 * 用途說明
 * ------------------------------------------------------------------
 * 1. 這是一份「最小可部署」的 Google Apps Script Web App 範例。
 * 2. 目前版本已可實際串接外部 AI API，目的除了驗證前後端通訊，也讓學員理解：
 *    - 前端欄位如何組成 payload
 *    - POST payload 如何送到後端
 *    - 後端如何選擇模型、帶入 system prompt、再把模型回覆組裝成最終輸出
 * 3. 此版本會依 keySource 與帳號角色分流：
 *    - Use Server Key：才套用 admin / user 角色限制
 *    - 管理者：admin，可取得完整教學版回覆
 *    - 一般使用者：user，在 Server Key 模式下只會收到受限教學回覆，不進行完整回答
 *    - Use My API Key：不受 admin / user 限制，直接走完整回覆
 *    - 所有回應仍保留台北時間欄位供教學觀察
 * 4. 等你之後提供正式部署網址後，我們再把前端 HTML 改成對接這支服務。
 * 5. 建議搭配同專案中的 SheetSetup.gs 使用，用來初始化預設資料表。
 *
 * ------------------------------------------------------------------
 * 環境變數 / Script Properties（可選）
 * ------------------------------------------------------------------
 * 請到 Apps Script：
 * 專案設定 > 指令碼屬性
 *
 * 本教學版建議設定六個主要指令碼屬性：
 *
 * GEMINI_API_KEY=你的 Gemini API Key
 * MINIMAX_API_KEY=你的 MiniMax API Key
 * KIMI_API_KEY=你的 Kimi API Key
 * SHARE_EXEC_CODE=取得範例網址時要輸入的驗證碼
 * SHARE_EXEC_URL=通過驗證後要回傳的 exec 網址
 * ADMIN_PASSWORD_CODE=取得 admin 密碼時要輸入的第二組驗證碼
 * AUDIT_TOTAL_EXECUTIONS=內部稽核用總執行次數計數器（系統每次執行 doGet / doPost 自動 +1）
 *
 * 其他資料都固定放在同一份 Google Sheet 中：
 * - README：說明文件與設定提示
 * - users：使用者帳號與密碼
 * - logs：輸入輸出與測試紀錄
 * - forced_output：強制輸出的前言與結語
 * - runtime_config：後端即時設定欄位
 *
 * 其他可選欄位：
 *
 * APP_TIMEZONE=Asia/Taipei
 * REPLY_PREFIX=你剛剛輸入的是：
 *
 * 說明：
 * - 如果沒有設定，程式會使用內建預設值。
 * - 程式預設模型是 `gemini`。
 * - 若前端 payload 有帶 `aiModel`，就會用使用者指定的模型覆蓋預設值。
 * - 三組 API Key 可先設定好，之後再逐步接到正式模型 API。
 * - AUDIT_TOTAL_EXECUTIONS 為內部稽核變數，建議讓系統自動維護，不要手動修改。
 * - 建議將 Apps Script 直接綁定在這份 Google Sheet 上使用。
 * - 工作表名稱固定為 README、users、logs、forced_output、runtime_config。
 * - 這一版偏向「偵錯模式」，會盡量輸出更多可教學資訊。
 * - 這一版是「通訊驗證 + Google Sheet 教學版」的基礎骨架。
 *
 * ------------------------------------------------------------------
 * 部署方式（教學版）
 * ------------------------------------------------------------------
 * 1. 建立或上傳本教學用 Google Sheet
 * 2. 將 Apps Script 綁定到該 Google Sheet
 * 3. 將本檔內容貼到 Code.gs
 * 4. 將初始化檔貼到 SheetSetup.gs
 * 5. 先執行一次 setupTeachingSheets()
 * 6. 儲存專案
 * 7. 點選「部署」>「新增部署作業」
 * 5. 類型選「網頁應用程式」
 * 6. 執行身分：自己
 * 7. 存取權限：任何人
 * 8. 部署後取得 Web App URL
 *
 * ------------------------------------------------------------------
 * 前端送法建議（重要）
 * ------------------------------------------------------------------
 * 由於 GAS Web App 在瀏覽器環境下，若你直接用 application/json，
 * 很容易觸發 CORS 預檢（OPTIONS）問題。
 *
 * 建議前端使用：
 * - method: POST
 * - Content-Type: text/plain;charset=utf-8
 * - body: JSON.stringify(payload)
 *
 * 範例 payload：
 * {
 *   "action": "chat",
 *   "message": "Hello GAS",
 *   "systemPrompt": "請用清楚、簡短、教學導向的方式回覆。",
 *   "outputCharLimit": 500,
 *   "aiModel": "gemini",
 *   "keySource": "server",
 *   "userApiKey": ""
 * }
 *
 * ------------------------------------------------------------------
 * 版權宣導 / 使用提醒
 * ------------------------------------------------------------------
 * 本檔為 Force Cheng × FALO 教學與內部知識資產範例。
 * 可作為教學展示、內部測試、客戶說明用途。
 * 若要對外散佈、商業整合或移除署名，請先取得授權。
 * ------------------------------------------------------------------
 */

var README_SHEET_NAME = "README";
var USERS_SHEET_NAME = "users";
var LOGS_SHEET_NAME = "logs";
var FORCED_OUTPUT_SHEET_NAME = "forced_output";
var RUNTIME_CONFIG_SHEET_NAME = "runtime_config";
var DEFAULT_SYSTEM_PROMPT = "請用清楚、簡短、教學導向的方式回覆，並保留足夠資訊讓學生理解系統運作。除非使用者明確指定其他語言，否則一律使用繁體中文（台灣）回答。";
var DEFAULT_OUTPUT_CHAR_LIMIT = 500;

function doGet(e) {
  var auditTotalExecutions = incrementAuditExecutionCounter_();
  var config = getAppConfig_();
  var nowTaipei = getTaipeiTime_(config.timezone);
  var runtimeConfig = getRuntimeConfig_();

  return jsonOutput_({
    success: true,
    mode: "GET",
    message: "GAS Web App is running in teaching debug mode.",
    data: {
      appName: config.appName,
      version: config.version,
      edition: config.edition,
      author: config.author,
      framework: config.framework,
      timezone: config.timezone,
      apiKeys: {
        gemini: config.geminiApiKeyConfigured,
        minimax: config.minimaxApiKeyConfigured,
        kimi: config.kimiApiKeyConfigured
      },
      readmeSheetName: README_SHEET_NAME,
      usersSheetName: USERS_SHEET_NAME,
      logsSheetName: LOGS_SHEET_NAME,
      forcedOutputSheetName: FORCED_OUTPUT_SHEET_NAME,
      runtimeConfigSheetName: RUNTIME_CONFIG_SHEET_NAME,
      auditTotalExecutions: auditTotalExecutions,
      responseTimeTaipei: nowTaipei,
      debugMode: true,
      runtimeConfig: runtimeConfig,
      usage: {
        method: "POST",
        contentType: "text/plain;charset=utf-8",
        action: "chat"
      },
      shareExample: {
        shareExecCodeConfigured: config.shareExecCodeConfigured,
        shareExecUrlConfigured: config.shareExecUrlConfigured,
        adminPasswordCodeConfigured: config.adminPasswordCodeConfigured
      },
      samplePayload: {
        action: "chat",
        message: "Please reply with: API connection successful.",
        systemPrompt: runtimeConfig.systemPrompt,
        outputCharLimit: runtimeConfig.outputCharLimit,
        aiModel: "gemini",
        keySource: "server",
        userApiKey: ""
      }
    }
  });
}

function doPost(e) {
  var auditTotalExecutions = incrementAuditExecutionCounter_();
  var config = getAppConfig_();
  var nowTaipei = getTaipeiTime_(config.timezone);
  var payload = parseRequestPayload_(e);
  var runtimeConfig = getRuntimeConfig_();

  if (!payload.ok) {
    return jsonOutput_({
      success: false,
      message: payload.error,
      data: {
        responseTimeTaipei: nowTaipei,
        auditTotalExecutions: auditTotalExecutions
      }
    });
  }

  var input = payload.data || {};
  var action = String(input.action || "chat").trim();

  if (action === "get_share_url") {
    return handleShareUrlRequest_(input, config, nowTaipei);
  }

  if (action === "get_admin_password") {
    return handleAdminPasswordRequest_(input, config, nowTaipei);
  }
  var message = String(input.message || "").trim();
  var systemPrompt = normalizeSystemPrompt_(input.systemPrompt, runtimeConfig.systemPrompt);
  var outputCharLimit = normalizeOutputCharLimit_(input.outputCharLimit, runtimeConfig.outputCharLimit);
  var systemPromptSource = hasValue_(input.systemPrompt) ? "payload" : "runtime_config";
  var outputLimitSource = hasValue_(input.outputCharLimit) ? "payload" : "runtime_config";
  var aiModel = normalizeAiModel_(input.aiModel);
  var keySource = String(input.keySource || "server").trim();
  var userApiKey = String(input.userApiKey || "").trim();
  var username = String(input.username || "").trim();
  var password = String(input.password || "").trim();
  var apiKeyStatus = getApiKeyStatus_(config, aiModel);
  var forcedOutput = getForcedOutputText_();
  var authResult;

  if (keySource === "server" && !apiKeyStatus.configured) {
    var missingEnvMessage = "未填寫環境變數：" + apiKeyStatus.keyName;

    appendLog_({
      timestampTaipei: nowTaipei,
      username: username || "anonymous",
      action: action || "chat",
      userInput: message,
      systemOutput: missingEnvMessage,
      aiModel: aiModel,
      keySource: keySource,
      success: false,
      notes: "缺少對應模型的後端環境變數"
    });

    return jsonOutput_({
      success: false,
      message: missingEnvMessage,
      data: {
        responseTimeTaipei: nowTaipei,
        aiModel: aiModel,
        keySource: keySource,
        auditTotalExecutions: auditTotalExecutions,
        missingEnvVar: apiKeyStatus.keyName,
        serverApiKeyConfigured: false
      }
    });
  }

  if (keySource === "user" && !userApiKey) {
    var missingUserKeyMessage = "未填寫使用者 API Key";

    appendLog_({
      timestampTaipei: nowTaipei,
      username: username || "anonymous",
      action: action || "chat",
      userInput: message,
      systemOutput: missingUserKeyMessage,
      aiModel: aiModel,
      keySource: keySource,
      success: false,
      notes: "使用者選擇自填 API Key，但未輸入"
    });

    return jsonOutput_({
      success: false,
      message: missingUserKeyMessage,
      data: {
        responseTimeTaipei: nowTaipei,
        aiModel: aiModel,
        keySource: keySource,
        auditTotalExecutions: auditTotalExecutions,
        userApiKeyProvided: false
      }
    });
  }

  authResult = getAuthResult_(username, password);

  if (!authResult.ok) {
    appendLog_({
      timestampTaipei: nowTaipei,
      username: username || "anonymous",
      action: action || "chat",
      userInput: message,
      systemOutput: authResult.message,
      aiModel: aiModel,
      keySource: keySource,
      success: false,
      notes: "登入驗證失敗"
    });

    return jsonOutput_({
      success: false,
      message: authResult.message,
      data: {
        responseTimeTaipei: nowTaipei,
        username: username || "",
        aiModel: aiModel,
        auditTotalExecutions: auditTotalExecutions,
        loginRequired: true
      }
    });
  }

  if (action !== "chat") {
    return jsonOutput_({
      success: false,
      message: "Unsupported action: " + action,
      data: {
        responseTimeTaipei: nowTaipei,
        auditTotalExecutions: auditTotalExecutions,
        expectedAction: "chat"
      }
    });
  }

  if (!message) {
    return jsonOutput_({
      success: false,
      message: "message 欄位不可為空。",
      data: {
        responseTimeTaipei: nowTaipei,
        auditTotalExecutions: auditTotalExecutions
      }
    });
  }

  var responseMode = getResponseMode_(authResult.user.role, keySource);
  var replyResult;

  try {
    // 這裡開始進入真正的「完整迴路」：
    // 1. 先判斷這次是否允許走完整回答模式
    // 2. 若允許，就實際呼叫對應 AI API
    // 3. 最後再把前言、系統限制提示、使用者輸入、模型回覆、結語組回教學版輸出
    replyResult = buildReplyByRole_({
      role: authResult.user.role,
      keySource: keySource,
      message: message,
      aiModel: aiModel,
      systemPrompt: systemPrompt,
      outputCharLimit: outputCharLimit,
      nowTaipei: nowTaipei,
      forcedOutput: forcedOutput,
      replyPrefix: config.replyPrefix,
      userApiKey: userApiKey
    });
  } catch (error) {
    appendLog_({
      timestampTaipei: nowTaipei,
      username: authResult.user.username,
      action: action,
      userInput: message,
      systemOutput: "AI 呼叫失敗：" + error.message,
      aiModel: aiModel,
      keySource: keySource,
      success: false,
      notes: "AI provider request failed"
    });

    return jsonOutput_({
      success: false,
      message: "AI 呼叫失敗：" + error.message,
      data: {
        responseTimeTaipei: nowTaipei,
        username: authResult.user.username,
        role: authResult.user.role,
        responseMode: responseMode,
        aiModel: aiModel,
        keySource: keySource,
        auditTotalExecutions: auditTotalExecutions
      }
    });
  }

  var logResult = appendLog_({
    timestampTaipei: nowTaipei,
    username: authResult.user.username,
    action: action,
    userInput: message,
    systemOutput: replyResult.fullReply,
    aiModel: aiModel,
    keySource: keySource,
    success: true,
    notes: "教學版 Web App 回應成功 / role=" + authResult.user.role + " / responseMode=" + responseMode + " / provider=" + replyResult.provider + " / outputCharLimit=" + outputCharLimit
  });

  return jsonOutput_({
    success: true,
    message: "Request processed successfully.",
    data: {
      reply: replyResult.fullReply,
      username: authResult.user.username,
      role: authResult.user.role,
      displayName: authResult.user.displayName,
      responseMode: responseMode,
      provider: replyResult.provider,
      modelRequestName: replyResult.modelRequestName,
      coreReply: replyResult.coreReply,
      coreReplyRaw: replyResult.coreReplyRaw || replyResult.coreReply,
      replySections: replyResult.sections || {},
      userMessage: message,
      systemPrompt: systemPrompt,
      outputCharLimit: outputCharLimit,
      aiModel: aiModel,
      keySource: keySource,
      userApiKeyProvided: !!userApiKey,
      serverApiKeyConfigured: apiKeyStatus.configured,
      serverApiKeyName: apiKeyStatus.keyName,
      auditTotalExecutions: auditTotalExecutions,
      responseTimeTaipei: nowTaipei,
      apiKeys: {
        gemini: config.geminiApiKeyConfigured,
        minimax: config.minimaxApiKeyConfigured,
        kimi: config.kimiApiKeyConfigured
      },
      readmeSheetName: README_SHEET_NAME,
      usersSheetName: USERS_SHEET_NAME,
      logsSheetName: LOGS_SHEET_NAME,
      forcedOutputSheetName: FORCED_OUTPUT_SHEET_NAME,
      runtimeConfigSheetName: RUNTIME_CONFIG_SHEET_NAME,
      logSaved: logResult.ok,
      logMessage: logResult.message,
      appName: config.appName,
      version: config.version,
      edition: config.edition,
      debug: {
        debugMode: true,
        timestampTaipei: nowTaipei,
        selectedModel: aiModel,
        selectedServerKeyName: apiKeyStatus.keyName,
        selectedServerKeyConfigured: apiKeyStatus.configured,
        systemPrompt: systemPrompt,
        outputCharLimit: outputCharLimit,
        systemPromptSource: systemPromptSource,
        outputCharLimitSource: outputLimitSource,
        runtimeConfig: runtimeConfig,
        receivedPayload: sanitizePayloadForDebug_(input),
        authResult: {
          ok: authResult.ok,
          username: authResult.user.username,
          role: authResult.user.role
        },
        responseMode: responseMode,
        provider: replyResult.provider,
        modelRequestName: replyResult.modelRequestName,
        providerRequest: replyResult.providerRequest,
        providerRawResponse: replyResult.providerRawResponse,
        forcedOutput: forcedOutput,
        replyBeforeLimit: replyResult.coreReplyRaw || replyResult.fullReply,
        replyAfterLimit: replyResult.fullReply,
        originalReplyLength: String(replyResult.coreReplyRaw || replyResult.fullReply || "").length,
        finalReplyLength: String(replyResult.fullReply || "").length,
        outputLimitApplied: String(replyResult.coreReplyRaw || replyResult.coreReply || "") !== String(replyResult.coreReply || ""),
        sheets: {
          readme: README_SHEET_NAME,
          users: USERS_SHEET_NAME,
          logs: LOGS_SHEET_NAME,
          forcedOutput: FORCED_OUTPUT_SHEET_NAME,
          runtimeConfig: RUNTIME_CONFIG_SHEET_NAME
        }
      }
    }
  });
}

function getAppConfig_() {
  var props = PropertiesService.getScriptProperties();

  return {
    appName: props.getProperty("APP_NAME") || "Multi-Model API Key Test Console",
    version: props.getProperty("APP_VERSION") || "v1.02",
    edition: props.getProperty("APP_EDITION") || "Force Teaching Edition",
    author: props.getProperty("APP_AUTHOR") || "Force Cheng",
    framework: props.getProperty("APP_FRAMEWORK") || "FALO (Formosa AI Life Outlook)",
    shareExecCodeConfigured: !!(props.getProperty("SHARE_EXEC_CODE") || "").trim(),
    shareExecUrlConfigured: !!(props.getProperty("SHARE_EXEC_URL") || "").trim(),
    adminPasswordCodeConfigured: !!(props.getProperty("ADMIN_PASSWORD_CODE") || "").trim(),
    geminiApiKeyConfigured: !!(props.getProperty("GEMINI_API_KEY") || "").trim(),
    minimaxApiKeyConfigured: !!(props.getProperty("MINIMAX_API_KEY") || "").trim(),
    kimiApiKeyConfigured: !!(props.getProperty("KIMI_API_KEY") || "").trim(),
    timezone: props.getProperty("APP_TIMEZONE") || "Asia/Taipei",
    replyPrefix: props.getProperty("REPLY_PREFIX") || "你剛剛輸入的是："
  };
}

function incrementAuditExecutionCounter_() {
  var lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    var props = PropertiesService.getScriptProperties();
    var currentValue = parseInt(String(props.getProperty("AUDIT_TOTAL_EXECUTIONS") || "0").trim(), 10);
    var nextValue = (isNaN(currentValue) ? 0 : currentValue) + 1;
    props.setProperty("AUDIT_TOTAL_EXECUTIONS", String(nextValue));
    return nextValue;
  } finally {
    lock.releaseLock();
  }
}

function handleShareUrlRequest_(input, config, nowTaipei) {
  var props = PropertiesService.getScriptProperties();
  var expectedCode = String(props.getProperty("SHARE_EXEC_CODE") || "").trim();
  var shareExecUrl = String(props.getProperty("SHARE_EXEC_URL") || "").trim();
  var accessCode = String(input.accessCode || "").trim();

  if (!expectedCode) {
    return jsonOutput_({
      success: false,
      message: "請跟管理者索取",
      data: {
        responseTimeTaipei: nowTaipei
      }
    });
  }

  if (!shareExecUrl) {
    return jsonOutput_({
      success: false,
      message: "請跟管理者索取",
      data: {
        responseTimeTaipei: nowTaipei
      }
    });
  }

  if (!accessCode) {
    return jsonOutput_({
      success: false,
      message: "請跟管理者索取",
      data: {
        responseTimeTaipei: nowTaipei
      }
    });
  }

  if (accessCode !== expectedCode) {
    return jsonOutput_({
      success: false,
      message: "請跟管理者索取",
      data: {
        responseTimeTaipei: nowTaipei
      }
    });
  }

  return jsonOutput_({
    success: true,
    message: "驗證成功，已提供範例網址。",
    data: {
      responseTimeTaipei: nowTaipei,
      shareExecUrl: shareExecUrl
    }
  });
}

function handleAdminPasswordRequest_(input, config, nowTaipei) {
  var props = PropertiesService.getScriptProperties();
  var expectedCode = String(props.getProperty("ADMIN_PASSWORD_CODE") || "").trim();
  var accessCode = String(input.accessCode || "").trim();
  var adminPassword = getAdminPasswordFromUsers_();

  if (!expectedCode || !adminPassword) {
    return jsonOutput_({
      success: false,
      message: "請跟管理者索取",
      data: {
        responseTimeTaipei: nowTaipei
      }
    });
  }

  if (!accessCode || accessCode !== expectedCode) {
    return jsonOutput_({
      success: false,
      message: "請跟管理者索取",
      data: {
        responseTimeTaipei: nowTaipei
      }
    });
  }

  return jsonOutput_({
    success: true,
    message: "驗證成功，已提供管理者密碼。",
    data: {
      responseTimeTaipei: nowTaipei,
      username: "admin",
      password: adminPassword
    }
  });
}

function getAdminPasswordFromUsers_() {
  var sheet = getSheetByName_(USERS_SHEET_NAME);
  if (!sheet) return "";

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return "";

  var headers = values[0];
  var usernameIndex = headers.indexOf("username");
  var passwordIndex = headers.indexOf("password");

  if (usernameIndex === -1 || passwordIndex === -1) return "";

  for (var i = 1; i < values.length; i++) {
    var rowUsername = String(values[i][usernameIndex] || "").trim().toLowerCase();
    if (rowUsername === "admin") {
      return String(values[i][passwordIndex] || "").trim();
    }
  }

  return "";
}

function getTaipeiTime_(timezone) {
  return Utilities.formatDate(new Date(), timezone || "Asia/Taipei", "yyyy-MM-dd HH:mm:ss");
}

function normalizeAiModel_(aiModel) {
  var normalized = String(aiModel || "gemini").trim().toLowerCase();
  if (normalized === "minimax" || normalized === "kimi") {
    return normalized;
  }
  return "gemini";
}

function normalizeSystemPrompt_(systemPrompt, runtimeValue) {
  var normalized = String(systemPrompt || "").trim();
  if (normalized) return normalized;
  var sheetValue = String(runtimeValue || "").trim();
  return sheetValue || DEFAULT_SYSTEM_PROMPT;
}

function normalizeOutputCharLimit_(value, runtimeValue) {
  if (String(value || "").trim() === "0") return 0;
  var parsed = parseInt(value, 10);
  if (parsed && parsed > 0) return parsed;
  if (String(runtimeValue || "").trim() === "0") return 0;
  var runtimeParsed = parseInt(runtimeValue, 10);
  if (runtimeParsed && runtimeParsed > 0) return runtimeParsed;
  return DEFAULT_OUTPUT_CHAR_LIMIT;
}

function getApiKeyStatus_(config, aiModel) {
  if (aiModel === "minimax") {
    return {
      model: "minimax",
      keyName: "MINIMAX_API_KEY",
      configured: config.minimaxApiKeyConfigured
    };
  }

  if (aiModel === "kimi") {
    return {
      model: "kimi",
      keyName: "KIMI_API_KEY",
      configured: config.kimiApiKeyConfigured
    };
  }

  return {
    model: "gemini",
    keyName: "GEMINI_API_KEY",
    configured: config.geminiApiKeyConfigured
  };
}

function getForcedOutputText_() {
  var defaults = {
    preface: "【教學版前言】這是示範用回覆，用來幫助學生理解 GAS 如何組合輸出內容。",
    conclusion: "【教學版結語】你可以修改 forced_output 工作表，觀察回覆文字如何一起改變。"
  };

  var sheet = getSheetByName_(FORCED_OUTPUT_SHEET_NAME);
  if (!sheet) {
    return defaults;
  }

  var values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return defaults;
  }

  var result = {
    preface: defaults.preface,
    conclusion: defaults.conclusion
  };

  for (var i = 1; i < values.length; i++) {
    var key = String(values[i][0] || "").trim().toLowerCase();
    var content = String(values[i][1] || "").trim();
    var enabled = String(values[i][2] || "TRUE").trim().toUpperCase() !== "FALSE";

    if (key === "preface" && content && enabled) {
      result.preface = content;
    }

    if (key === "conclusion" && content && enabled) {
      result.conclusion = content;
    }
  }

  return result;
}

function getRuntimeConfig_() {
  var defaults = {
    systemPrompt: DEFAULT_SYSTEM_PROMPT,
    outputCharLimit: DEFAULT_OUTPUT_CHAR_LIMIT
  };

  var sheet = getSheetByName_(RUNTIME_CONFIG_SHEET_NAME);
  if (!sheet) {
    return defaults;
  }

  var values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return defaults;
  }

  var result = {
    systemPrompt: defaults.systemPrompt,
    outputCharLimit: defaults.outputCharLimit
  };

  for (var i = 1; i < values.length; i++) {
    var key = String(values[i][0] || "").trim().toLowerCase();
    var value = String(values[i][1] || "").trim();
    var enabled = String(values[i][2] || "TRUE").trim().toUpperCase() !== "FALSE";
    if (!enabled) continue;

    if (key === "system_prompt" && value) {
      result.systemPrompt = value;
    }

    if (key === "output_char_limit" && value) {
      result.outputCharLimit = normalizeOutputCharLimit_(value, defaults.outputCharLimit);
    }
  }

  return result;
}

function applyOutputCharLimit_(text, limit) {
  var source = String(text || "");
  var maxLength = normalizeOutputCharLimit_(limit);

  if (maxLength === 0) {
    return {
      text: source,
      wasTrimmed: false
    };
  }

  if (source.length <= maxLength) {
    return {
      text: source,
      wasTrimmed: false
    };
  }

  if (maxLength <= 1) {
    return {
      text: source.substring(0, maxLength),
      wasTrimmed: true
    };
  }

  return {
    text: source.substring(0, maxLength - 1) + "…",
    wasTrimmed: true
  };
}

function hasValue_(value) {
  return String(value || "").trim() !== "";
}

function isAdminRole_(role) {
  return String(role || "").trim().toLowerCase() === "admin";
}

function getResponseMode_(role, keySource) {
  var normalizedKeySource = String(keySource || "server").trim().toLowerCase();
  if (normalizedKeySource === "user") {
    return "full_reply_user_key";
  }
  return isAdminRole_(role) ? "admin_full_reply" : "limited_info_reply";
}

function buildReplyByRole_(options) {
  var role = String(options.role || "user").trim().toLowerCase();
  var keySource = String(options.keySource || "server").trim().toLowerCase();
  var message = String(options.message || "").trim();
  var aiModel = String(options.aiModel || "gemini").trim().toLowerCase();
  var systemPrompt = String(options.systemPrompt || "").trim();
  var outputCharLimit = String(options.outputCharLimit || "").trim();
  var nowTaipei = String(options.nowTaipei || "").trim();
  var replyPrefix = String(options.replyPrefix || "你剛剛輸入的是：").trim();
  var forcedOutput = options.forcedOutput || {};

  if (keySource === "server" && !isAdminRole_(role)) {
    var limitedSections = {
      preface: "",
      systemPromptLine: "",
      userInputLine: replyPrefix + message,
      systemReplyLine: "【一般使用者模式】目前使用的是 Server Key，因此這個帳號不走完整管理者回覆，而是回傳教學型資訊內容。\n【系統說明】系統已成功收到你的輸入，並完成帳號驗證、模型辨識與後端設定讀取。\n【模式差異】若要取得完整教學版回覆，可改用 admin 帳號，或切換成 Use My API Key。\n【目前模型】一般使用者在 Server Key 模式下，仍可觀察模型選擇、輸出限制與回應時間等系統資訊。",
      outputLimitLine: "",
      responseTimeLine: "回應時間（台北）： " + nowTaipei,
      conclusion: ""
    };
    return {
      fullReply: [
        limitedSections.systemReplyLine,
        limitedSections.userInputLine,
        limitedSections.responseTimeLine
      ].filter(function(item) {
        return String(item || "").trim() !== "";
      }).join("\n"),
      coreReply: "一般使用者在 Server Key 模式下，不執行完整 AI 回覆。",
      coreReplyRaw: "一般使用者在 Server Key 模式下，不執行完整 AI 回覆。",
      sections: limitedSections,
      provider: "teaching-limited",
      modelRequestName: aiModel,
      providerRequest: {
        skipped: true,
        reason: "server key + non-admin role"
      },
      providerRawResponse: {}
    };
  }

  // 這裡才是真正會送到 AI 的完整模式：
  // - admin + Use Server Key
  // - 任意角色 + Use My API Key
  var aiExecution = executeAiSingleTurn_({
    aiModel: aiModel,
    keySource: keySource,
    userApiKey: options.userApiKey || "",
    systemPrompt: systemPrompt,
    outputCharLimit: outputCharLimit,
    message: message
  });

  var limitedCoreReply = applyOutputCharLimit_(aiExecution.reply, outputCharLimit);
  var structuredSections = {
    preface: String(forcedOutput.preface || "").trim(),
    systemPromptLine: "【系統限制提示】" + systemPrompt,
    userInputLine: "【使用者輸入】" + message,
    systemReplyLine: "【系統回覆】" + limitedCoreReply.text,
    outputLimitLine: "【輸出限制】最多 " + outputCharLimit + " 字",
    responseTimeLine: "回應時間（台北）： " + nowTaipei,
    conclusion: String(forcedOutput.conclusion || "").trim()
  };

  return {
    fullReply: [
      structuredSections.preface,
      structuredSections.systemPromptLine,
      structuredSections.userInputLine,
      structuredSections.systemReplyLine,
      structuredSections.outputLimitLine,
      structuredSections.responseTimeLine,
      structuredSections.conclusion
    ].filter(function(item) {
      return String(item || "").trim() !== "";
    }).join("\n"),
    coreReply: limitedCoreReply.text,
    coreReplyRaw: aiExecution.reply,
    sections: structuredSections,
    provider: aiExecution.provider,
    modelRequestName: aiExecution.modelRequestName,
    providerRequest: aiExecution.providerRequest,
    providerRawResponse: aiExecution.providerRawResponse
  };
}

function executeAiSingleTurn_(options) {
  var aiModel = String(options.aiModel || "gemini").trim().toLowerCase();
  var apiKey = getApiKeyValue_(aiModel, options.keySource, options.userApiKey);
  var systemPrompt = String(options.systemPrompt || "").trim();
  var message = String(options.message || "").trim();
  var outputCharLimit = normalizeOutputCharLimit_(options.outputCharLimit);

  if (!apiKey) {
    throw new Error("找不到可用的 API Key，無法執行完整 AI 回覆。");
  }

  if (aiModel === "kimi") {
    return callOpenAiCompatibleModel_({
      provider: "moonshot",
      endpoint: "https://api.moonshot.ai/v1/chat/completions",
      model: "kimi-k2.5",
      apiKey: apiKey,
      systemPrompt: systemPrompt,
      message: message,
      outputCharLimit: outputCharLimit,
      temperature: 1
    });
  }

  if (aiModel === "minimax") {
    return callOpenAiCompatibleModel_({
      provider: "minimax",
      endpoint: "https://api.minimax.io/v1/chat/completions",
      model: "MiniMax-M2.7",
      apiKey: apiKey,
      systemPrompt: systemPrompt,
      message: message,
      outputCharLimit: outputCharLimit,
      reasoningSplit: true
    });
  }

  return callGeminiModel_({
    apiKey: apiKey,
    systemPrompt: systemPrompt,
    message: message,
    outputCharLimit: outputCharLimit
  });
}

function getApiKeyValue_(aiModel, keySource, userApiKey) {
  var normalizedKeySource = String(keySource || "server").trim().toLowerCase();
  if (normalizedKeySource === "user") {
    return String(userApiKey || "").trim();
  }

  var props = PropertiesService.getScriptProperties();
  if (aiModel === "minimax") {
    return String(props.getProperty("MINIMAX_API_KEY") || "").trim();
  }

  if (aiModel === "kimi") {
    return String(props.getProperty("KIMI_API_KEY") || "").trim();
  }

  return String(props.getProperty("GEMINI_API_KEY") || "").trim();
}

function callGeminiModel_(options) {
  // Gemini 走 Google 官方 generateContent REST API。
  // 這裡把 systemPrompt 放在 systemInstruction，
  // 再把使用者問題與字數限制提示一起送進內容區。
  var endpoint = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" +
    encodeURIComponent(options.apiKey);
  var payload = {
    systemInstruction: {
      parts: [{
        text: options.systemPrompt
      }]
    },
    contents: [{
      role: "user",
      parts: [{
        text: buildAiUserPrompt_(options.message, options.outputCharLimit)
      }]
    }],
    generationConfig: {
      temperature: 0.2,
      maxOutputTokens: estimateMaxOutputTokens_(options.outputCharLimit)
    }
  };

  var rawResponse = fetchJsonFromApi_(endpoint, {
    method: "post",
    contentType: "application/json",
    headers: {},
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var text = extractGeminiText_(rawResponse.json);
  if (!text) {
    throw new Error("Gemini 沒有返回可讀文字。");
  }

  return {
    reply: text,
    provider: "gemini",
    modelRequestName: "gemini-2.5-flash",
    providerRequest: payload,
    providerRawResponse: rawResponse.json
  };
}

function callOpenAiCompatibleModel_(options) {
  // Kimi 與 MiniMax 目前都採 OpenAI 相容格式，因此共用同一個 helper。
  // systemPrompt 直接放 system role，使用者訊息放 user role。
  // 但不同供應商仍可能有自己的參數限制，例如某些 Kimi 模型只接受 temperature=1。
  var payload = {
    model: options.model,
    messages: [
      {
        role: "system",
        content: buildAiSystemPrompt_(options.systemPrompt, options.outputCharLimit)
      },
      {
        role: "user",
        content: options.message
      }
    ]
  };

  if (typeof options.temperature !== "undefined" && options.temperature !== null) {
    payload.temperature = options.temperature;
  } else {
    payload.temperature = 0.2;
  }

  if (options.reasoningSplit) {
    payload.reasoning_split = true;
  }

  var rawResponse = fetchJsonFromApi_(options.endpoint, {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + options.apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var text = extractOpenAiCompatibleText_(rawResponse.json);
  if (!text) {
    throw new Error(options.provider + " 沒有返回可讀文字。");
  }

  return {
    reply: stripThinkTags_(text),
    provider: options.provider,
    modelRequestName: options.model,
    providerRequest: payload,
    providerRawResponse: rawResponse.json
  };
}

function fetchJsonFromApi_(url, requestOptions) {
  var response = UrlFetchApp.fetch(url, requestOptions);
  var statusCode = response.getResponseCode();
  var bodyText = response.getContentText();
  var json;

  try {
    json = JSON.parse(bodyText);
  } catch (error) {
    throw new Error("API 回傳不是合法 JSON。HTTP " + statusCode + " / body=" + bodyText);
  }

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error("API 請求失敗。HTTP " + statusCode + " / " + bodyText);
  }

  return {
    statusCode: statusCode,
    json: json
  };
}

function extractGeminiText_(json) {
  var candidates = json && json.candidates ? json.candidates : [];
  if (!candidates.length) return "";

  var parts = (((candidates[0] || {}).content || {}).parts || []);
  return parts.map(function(part) {
    return String((part && part.text) || "");
  }).join("").trim();
}

function extractOpenAiCompatibleText_(json) {
  var choices = json && json.choices ? json.choices : [];
  if (!choices.length) return "";

  var message = choices[0] && choices[0].message ? choices[0].message : {};
  var content = message.content;

  if (Array.isArray(content)) {
    return content.map(function(item) {
      if (typeof item === "string") return item;
      return String((item && item.text) || "");
    }).join("").trim();
  }

  return String(content || "").trim();
}

function stripThinkTags_(text) {
  return String(text || "").replace(/<think>[\s\S]*?<\/think>/gi, "").trim();
}

function buildAiSystemPrompt_(systemPrompt, outputCharLimit) {
  var limit = normalizeOutputCharLimit_(outputCharLimit);
  return [
    String(systemPrompt || "").trim(),
    "除非使用者明確指定其他語言，否則一律使用繁體中文（台灣）回答。",
    "請直接輸出最終答案，不要顯示思考過程。",
    limit === 0 ? "本次回覆不限制字數，但仍請保持結構清楚、重點明確。" : "請盡量控制在 " + limit + " 字以內。"
  ].filter(function(item) {
    return item !== "";
  }).join("\n");
}

function buildAiUserPrompt_(message, outputCharLimit) {
  var limit = normalizeOutputCharLimit_(outputCharLimit);
  return [
    String(message || "").trim(),
    "",
    "補充要求：若我沒有指定語言，請使用繁體中文（台灣）回答。",
    "",
    limit === 0
      ? "補充要求：請直接回答，不要描述推理過程，本次不限制字數，但請維持清楚易讀。"
      : "補充要求：請直接回答，不要描述推理過程，並盡量控制在 " + limit + " 字以內。"
  ].join("\n");
}

function estimateMaxOutputTokens_(outputCharLimit) {
  var limit = normalizeOutputCharLimit_(outputCharLimit);
  if (limit === 0) {
    return 2048;
  }
  return Math.max(64, Math.min(1024, limit * 2));
}

function sanitizePayloadForDebug_(payload) {
  var sanitized = JSON.parse(JSON.stringify(payload || {}));
  if (sanitized.userApiKey) {
    sanitized.userApiKey = maskApiKey_(sanitized.userApiKey);
  }
  return sanitized;
}

function maskApiKey_(value) {
  var source = String(value || "").trim();
  if (!source) return "";
  if (source.length <= 8) return "****";
  return source.substring(0, 4) + "..." + source.substring(source.length - 4);
}

function getAuthResult_(username, password) {
  if (!username && !password) {
    return {
      ok: true,
      user: {
        username: "guest",
        role: "guest",
        displayName: "Guest User"
      }
    };
  }

  if (!username || !password) {
    return {
      ok: false,
      message: "若要登入，請同時提供 username 與 password。"
    };
  }

  var sheet = getSheetByName_(USERS_SHEET_NAME);
  if (!sheet) {
    return {
      ok: false,
      message: "找不到 users 工作表，請先執行 setupTeachingSheets()。"
    };
  }

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return {
      ok: false,
      message: "users 工作表沒有任何可用帳號。"
    };
  }

  var headers = values[0];
  var usernameIndex = headers.indexOf("username");
  var passwordIndex = headers.indexOf("password");
  var roleIndex = headers.indexOf("role");
  var statusIndex = headers.indexOf("status");
  var displayNameIndex = headers.indexOf("display_name");

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rowUsername = String(row[usernameIndex] || "").trim();
    var rowPassword = String(row[passwordIndex] || "").trim();
    var rowRole = String(row[roleIndex] || "user").trim();
    var rowStatus = String(row[statusIndex] || "active").trim().toLowerCase();
    var rowDisplayName = String(row[displayNameIndex] || rowUsername).trim();

    if (rowUsername === username && rowPassword === password) {
      if (rowStatus !== "active") {
        return {
          ok: false,
          message: "此帳號目前不是啟用狀態。"
        };
      }

      return {
        ok: true,
        user: {
          username: rowUsername,
          role: rowRole,
          displayName: rowDisplayName
        }
      };
    }
  }

  return {
    ok: false,
    message: "帳號或密碼錯誤。"
  };
}

function appendLog_(entry) {
  try {
    var sheet = getSheetByName_(LOGS_SHEET_NAME);
    if (!sheet) {
      return {
        ok: false,
        message: "找不到 logs 工作表。"
      };
    }

    var nextId = "LOG-" + zeroPad_(Math.max(1, sheet.getLastRow()), 4);
    sheet.appendRow([
      nextId,
      entry.timestampTaipei || "",
      entry.username || "guest",
      entry.action || "chat",
      entry.userInput || "",
      entry.systemOutput || "",
      entry.aiModel || "gemini",
      entry.keySource || "server",
      entry.success ? "TRUE" : "FALSE",
      "",
      "",
      entry.notes || ""
    ]);

    return {
      ok: true,
      message: "Log saved."
    };
  } catch (error) {
    return {
      ok: false,
      message: "Log save failed: " + error.message
    };
  }
}

function getSheetByName_(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return null;
  return ss.getSheetByName(sheetName);
}

function zeroPad_(num, size) {
  var text = String(num);
  while (text.length < size) {
    text = "0" + text;
  }
  return text;
}

function parseRequestPayload_(e) {
  try {
    if (e && e.postData && e.postData.contents) {
      return {
        ok: true,
        data: JSON.parse(e.postData.contents)
      };
    }

    if (e && e.parameter) {
      return {
        ok: true,
        data: e.parameter
      };
    }

    return {
      ok: false,
      error: "沒有收到任何請求內容。"
    };
  } catch (error) {
    return {
      ok: false,
      error: "JSON 解析失敗： " + error.message
    };
  }
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
