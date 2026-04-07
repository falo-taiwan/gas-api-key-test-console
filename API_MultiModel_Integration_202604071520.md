---
title: Multi-Model AI 助理專案：GAS 整合開發與踩坑指南
author: Force Cheng
framework: FALO (Formosa AI Life Outlook)
date: 2026-04-07
version: v1.02
copyright: © Force Cheng × FALO 2026/04/01
tags:
  - AI Agent
  - GAS
  - Multi Model
  - Gemini
  - Kimi
  - MiniMax
---

# Multi-Model AI 助理專案：GAS 整合開發與踩坑指南

> 更新時間：2026-04-07 15:20  
> 系統架構：Google Apps Script (GAS) + 前端 Vanilla JS  
> 支援模型：Google Gemini 2.5 Flash、Kimi (Moonshot)、MiniMax M2.7

這份文件整理了將單一 Gemini 助理升級為「多模型路由切換（Multi-Model Router）」架構的實作方式，並記錄在串接 Kimi 與 MiniMax 過程中，實際遇到的 API 問題、成因與對應解法。

## v1.02 教學版更新

本次版本控制重點是把教學專案升級成「角色分流」版本：

- `Use Server Key`：才套用角色分流
- `admin`：管理者帳號，可取得完整教學版回覆
- `user`：一般使用者帳號，在 Server Key 模式下進入受限教學回覆，不進行完整回答
- `Use My API Key`：不受 admin / user 限制，直接走完整回覆
- 專案主檔同步保留 `_v001` 複本，方便教學時對照版本差異

---

## 一、系統架構總覽

本節先用整體架構角度說明這份多模型整合方案，方便快速掌握前端、後端與模型路由之間的角色分工。

本專案採用以下架構：

- 後端：Google Apps Script（GAS）
- 前端：Vanilla JS
- 模型路由：依據前端傳入的 `aiModel`，由後端決定要呼叫哪一個模型 API

這樣的設計有兩個主要好處：

- 金鑰集中管理，不必暴露在前端
- 未來若要擴充更多模型，只需在後端 Router 增加對應分支

---

## 二、支援模型

本節整理目前文件涵蓋的模型範圍，幫助讀者先理解整個 Router 需要對接哪些供應商。

目前文件涵蓋的模型如下：

- Gemini：Google Gemini 2.5 Flash
- Kimi：Moonshot 系列模型
- MiniMax：MiniMax M2.7

---

## 三、三大踩坑與解法

本節聚焦實作時最容易出錯的三個關鍵點，包含端點、方案限制與輸出格式污染問題。

在串接過程中，最關鍵的問題主要集中在三個地方：端點混淆、方案限制，以及模型輸出格式污染。

### 1. Kimi 國際版與大陸版端點混淆

**狀況**

送出請求後，Kimi 持續回報 `401 Invalid Authentication`。即使反覆確認 API Key 正確，本機測試仍然失敗。

**原因**

API 說明文件中的網域資訊有誤。國際版金鑰若誤打到中國版網域 `api.moonshot.cn`，就會被伺服器視為非法憑證。

**解法**

請改用國際版端點：

```text
https://api.moonshot.ai/v1/chat/completions
```

### 2. MiniMax Token Plan 與模型綁定限制

**狀況**

- 打國際版網域時，出現 `401 - invalid api key (2049)`
- 改打中國版網域後，若呼叫其他模型，又會出現 `500 - your current token plan not support model (2061)`

**原因**

目前使用的 `sk-cp-*` 金鑰屬於中國版的 Token Plan 套餐。這類金鑰有兩個限制：

- 只能在中國機房網域使用
- 只能綁定指定模型，不能任意切換到其他型號

**解法**

後端需固定以下設定：

- 端點：`https://api.minimax.io/v1/chat/completions`
- 模型：`MiniMax-M2.7`

### 3. CoT `<think>` 汙染 JSON 結構

**狀況**

即使請求已成功返回 `200 OK`，GAS 仍可能拋出：

- `Cannot read properties of null`
- `Unexpected token '<'`

**原因**

某些具備思考鏈（Chain of Thought, CoT）能力的模型，會在實際 JSON 前插入類似 `<think>...</think>` 的內容，導致 `JSON.parse()` 在解析時直接失敗。

**解法**

在後端先清理思考鏈內容，再抽出真正的 JSON 主體：

```javascript
// 1. 先移除思考鏈標籤內容
aiResponseText = aiResponseText.replace(/<think>[\s\S]*?<\/think>/gi, '').trim();

// 2. 再抽出 JSON 主體
const jsonMatch = aiResponseText.match(/\{[\s\S]*\}/);
```

---

## 四、GAS 後端環境變數設定

本節說明金鑰管理方式，重點是把敏感資訊集中放在 GAS 指令碼屬性中，而不是直接寫死在程式碼。

為了兼顧安全性與可維護性，不建議把 API Key 寫死在程式碼中。請進入 GAS 的：

`專案設定 > 指令碼屬性`

新增以下屬性：

- `GEMINI_API_KEY`
- `KIMI_API_KEY`
- `MINIMAX_API_KEY`

補充說明：

- `KIMI_API_KEY` 若為國際版，需搭配 `.ai` 網域
- `MINIMAX_API_KEY` 需與對應的 Token Plan 與模型限制一致

---

## 五、GAS Router / Code.gs 範例

本節提供後端路由的最小可行範例，示範如何根據 `aiModel` 將請求分流到不同模型服務。

下面的範例示範如何根據前端傳來的 `aiModel` 進行路由：

```javascript
// ...接收前端 POST 請求...
const aiModel = postData.aiModel || "gemini";
const scriptProperties = PropertiesService.getScriptProperties();

if (aiModel === "kimi" || aiModel === "minimax") {
  // 1. 動態讀取 API Key，並用 trim() 避免隱藏空白造成驗證失敗
  let apiKey = aiModel === "kimi"
    ? scriptProperties.getProperty("KIMI_API_KEY")
    : scriptProperties.getProperty("MINIMAX_API_KEY");
  apiKey = apiKey.trim();

  // 2. 依模型決定端點與模型名稱
  const apiEndpoint = aiModel === "kimi"
    ? "https://api.moonshot.ai/v1/chat/completions"
    : "https://api.minimax.io/v1/chat/completions";

  const modelName = aiModel === "kimi"
    ? "moonshot-v1-8k"
    : "MiniMax-M2.7";

  // 3. 統一封裝 OpenAI 相容格式 Payload
  const payload_oa = {
    model: modelName,
    messages: [
      {
        role: "user",
        content: systemPrompt + "\n\n用戶要求：" + message
      }
    ],
    temperature: 0.1
  };

  // 4. MiniMax 有時會要求 max_tokens，建議明確帶入
  if (aiModel === "minimax") payload_oa.max_tokens = 2048;

  // ...接著使用 UrlFetchApp.fetch() 發送請求並處理結果...
} else {
  // 預設走原本的 Gemini 路線
}
```

---

## 六、前端 HTML：用 `<select>` 切換模型

本節示範如何在前端介面加入模型切換控制，讓使用者能直接選擇要呼叫的 AI 模型。

若要讓使用者在前端切換模型，可在表單中加入一個 `<select>`：

```html
<form id="chat-form" class="chat-input-wrapper">
  <select id="model-selector" class="model-select-btn">
    <option value="gemini">Gemini 2.5 Flash</option>
    <option value="kimi">Kimi k2.5</option>
    <option value="minimax">MiniMax M2.7</option>
  </select>

  <input type="text" id="chat-input" placeholder="Type your command...">
  <button type="submit" id="send-btn">Send</button>
</form>
```

這樣的做法可讓前端明確決定要把哪一個模型代號送到後端。

---

## 七、前端 JavaScript：`fetch` 時傳送 `aiModel`

本節對應前端送出流程，重點是把目前選定的模型一併傳給後端 Router。

在送出表單時，讀取使用者目前選擇的模型，並把它加入 API Payload：

```javascript
chatForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  const rawMessage = chatInput.value.trim();
  if (!rawMessage) return;

  // 取得目前選擇的模型
  const selectedModel = document.getElementById("model-selector").value;

  try {
    const response = await fetchAPI("chat", {
      message: rawMessage,
      aiModel: selectedModel
    });

    // 處理回傳...
  } catch (error) {
    console.error(error);
  }
});
```

前端負責傳值，後端負責路由，這樣模組邊界會更清楚，也更容易擴充。

---

## 八、結尾總結

本節用最精簡的方式回收全文重點，方便當成教學結語或文件摘要直接使用。

這份多模型整合的核心重點可以濃縮成三件事：

1. 前端只負責選擇模型並傳送 `aiModel`
2. 後端 Router 負責 API Key、端點與模型名稱的真正分流
3. 對於具備思考鏈能力的模型，要先處理 `<think>` 汙染，再進行 JSON 解析

只要把這三層做好，後續若要擴充 Claude、Groq 或其他模型，基本上只需要在 Router 中增加對應分支即可。
