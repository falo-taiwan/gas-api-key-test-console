# Multi-Model API Key Test Console v1.02

Force Teaching Edition  
Author: Force Cheng  
Framework: FALO (Formosa AI Life Outlook)  
Last Updated: 2026/4/7

這是一個教學型專案，用來示範如何用：

- Google Apps Script（GAS）
- Google Sheet
- 單檔前端 HTML
- 多模型 API（Gemini / Kimi / MiniMax）

組成一個可教學、可部署、可測試的 `API Key Test Console`。

它不是正式產品聊天台，而是偏向：

- 教學說明頁
- API Key 驗證頁
- 多模型整合測試頁
- GAS + Google Sheet 的最小可行範例

---

## 一、專案目標

這個專案的核心目標，是讓初學者能理解一個完整的 AI 系統最小迴路：

1. 前端表單怎麼組成 payload
2. 前端怎麼送到 GAS Web App
3. GAS 怎麼做登入驗證、角色分流、模型路由
4. Google Sheet 怎麼提供帳號、系統設定、固定輸出文字
5. 後端怎麼呼叫 AI，再把前言 / AI 回覆 / 後語組裝成可閱讀的結果

它也特別適合：

- AI 教學課程
- 給客戶做 API Key 測試
- 團隊內部做多模型串接示範
- 初學者理解 `前端 -> 後端 -> AI -> 回傳` 的實際流程

---

## 二、目前專案包含什麼

根目錄主要檔案：

- [index.html](/Users/force/AI-CodeX/Api-Key-Teach/index.html)
  說明頁，適合 GitHub Pages 首頁或教學講義頁。

- [playground.html](/Users/force/AI-CodeX/Api-Key-Teach/playground.html)
  執行頁，讓使用者實際測試登入、模型、API Key、system prompt、輸出限制等行為。

- [Code.gs](/Users/force/AI-CodeX/Api-Key-Teach/Code.gs)
  GAS 後端主程式，負責：
  - 處理 `GET / POST`
  - 驗證帳號
  - 檢查環境變數
  - 模型切換
  - 角色分流
  - AI 呼叫
  - logs 紀錄
  - 稽核計數

- [SheetSetup.gs](/Users/force/AI-CodeX/Api-Key-Teach/SheetSetup.gs)
  Google Sheet 初始化與修復腳本。

- [Api-Key-Teach-Template.xlsx](/Users/force/AI-CodeX/Api-Key-Teach/Api-Key-Teach-Template.xlsx)
  教學用 Excel 範本，可上傳後轉存成 Google Sheet。

- [API_MultiModel_Integration_202604071520.md](/Users/force/AI-CodeX/Api-Key-Teach/API_MultiModel_Integration_202604071520.md)
  原始技術整理文件。

其他資料夾：

- [backup](/Users/force/AI-CodeX/Api-Key-Teach/backup)
  版本封存與壓縮備份。

- [history](/Users/force/AI-CodeX/Api-Key-Teach/history)
  舊版歷史檔。

---

## 三、系統架構

這個專案可以用下面這條線來理解：

1. 使用者在 [playground.html](/Users/force/AI-CodeX/Api-Key-Teach/playground.html) 輸入資料
2. 前端以 `POST + text/plain + JSON.stringify(payload)` 送到 GAS
3. [Code.gs](/Users/force/AI-CodeX/Api-Key-Teach/Code.gs) 讀取：
   - Script Properties
   - `users`
   - `runtime_config`
   - `forced_output`
4. 後端決定：
   - 是否允許執行
   - 用哪個模型
   - 用 server key 還是 user key
   - 是否走完整回覆
5. 後端組裝回傳：
   - 前言
   - 系統限制提示
   - 使用者輸入
   - AI 回覆
   - 輸出限制
   - 回應時間
   - 結語
6. 前端把結果整理成教學可讀版本

簡化圖：

```text
playground.html
    ↓
GAS Web App (Code.gs)
    ↓
Google Sheet
  - users
  - runtime_config
  - forced_output
  - logs
    ↓
Gemini / Kimi / MiniMax
    ↓
回傳教學版混合輸出
```

---

## 四、角色與金鑰規則

這個專案有兩種金鑰來源：

- `Use Server Key`
- `Use My API Key`

角色規則只在 `Use Server Key` 生效。

### 1. Use Server Key

- `admin`
  可以取得完整教學版回覆。

- `user`
  只會進入受限教學回覆，不做完整 AI 回答。

### 2. Use My API Key

- 不套用 `admin / user` 限制
- 只要使用者自己填入 API Key，就可直接走完整回覆

這個設計很適合教學，因為學生可以清楚比較：

- 使用後端金鑰時，系統如何控管權限
- 使用個人金鑰時，系統如何放寬限制

---

## 五、支援模型

目前支援三個模型來源：

- `gemini`
- `kimi`
- `minimax`

後端會根據 `aiModel` 切換對應 API。

對應 Script Properties：

- `GEMINI_API_KEY`
- `KIMI_API_KEY`
- `MINIMAX_API_KEY`

如果選了 `Use Server Key`，但對應環境變數沒填，系統會明確回：

```text
未填寫環境變數：GEMINI_API_KEY
未填寫環境變數：KIMI_API_KEY
未填寫環境變數：MINIMAX_API_KEY
```

如果選了 `Use My API Key`，但欄位沒填，會直接擋下：

```text
未填寫使用者 API Key
```

---

## 六、Google Sheet 結構

這個專案目前使用 5 張工作表：

### 1. `README`

用途：

- 操作說明
- 環境變數說明
- 部署步驟
- 教學規則備註

### 2. `users`

用途：

- 存放帳號、密碼、角色

預設資料：

- `admin / 123`
- `user / 123`

### 3. `logs`

用途：

- 紀錄請求輸入
- 紀錄回應輸出
- 紀錄模型
- 紀錄 key source
- 紀錄 debug 與教學資訊

### 4. `forced_output`

用途：

- 存放前言與結語

預設項目：

- `preface`
- `conclusion`

### 5. `runtime_config`

用途：

- 控制後端即時 `systemPrompt`
- 控制後端即時 `outputCharLimit`

目前預設：

- `system_prompt`
- `output_char_limit = 500`

---

## 七、SheetSetup 的設計原則

[SheetSetup.gs](/Users/force/AI-CodeX/Api-Key-Teach/SheetSetup.gs) 不是單純建立 Sheet，它是「初始化 + 自我修復」腳本。

它的規則是：

- 如果結構不正確：
  - 缺 sheet
  - 表頭錯
  - 欄位順序錯
  - 必要預設資料缺失
  - 就清空並重建

- 如果結構正確：
  - 保留既有資料
  - 不主動覆寫內容

這一點很重要：

如果你的 `runtime_config.output_char_limit` 舊值是 `200`，但結構本身沒問題，重新執行 `setupTeachingSheets()` 不會自動洗成 `500`。

也就是說：

- `SheetSetup.gs` 的初始化預設是 `500`
- 但既有 Sheet 若資料已存在，腳本會選擇保留

---

## 八、環境變數（Script Properties）

請到：

`Apps Script -> 專案設定 -> 指令碼屬性`

### 主要 AI 金鑰

```text
GEMINI_API_KEY
KIMI_API_KEY
MINIMAX_API_KEY
```

### A / B 驗證機制

```text
SHARE_EXEC_CODE
SHARE_EXEC_URL
ADMIN_PASSWORD_CODE
```

用途如下：

#### A 機制：取得範例網址

- `SHARE_EXEC_CODE`
  第一組驗證碼

- `SHARE_EXEC_URL`
  驗證成功後要回傳的範例 `exec` 網址

前端行為：

- 使用者按「取得範例網址」
- 輸入第一組驗證碼
- 驗證成功後：
  - 自動回填 `GAS Web App URL`
  - 自動複製到剪貼簿

#### B 機制：載入管理者密碼

- `ADMIN_PASSWORD_CODE`
  第二組驗證碼

前端行為：

- 使用者按「填入管理者示範值」
- 輸入第二組驗證碼
- 驗證成功後：
  - 後端會去讀 `users` sheet 中 `admin` 的目前密碼
  - 自動回填到密碼欄位

### 稽核用隱藏變數

```text
AUDIT_TOTAL_EXECUTIONS
```

用途：

- 內部稽核用總執行次數
- 每次執行 `doGet` / `doPost` 會自動 `+1`

說明：

- 建議讓系統自動維護
- 不建議手動改值
- 前端右側摘要會顯示成：

```text
Audit 119
```

---

## 九、前端執行頁功能

[playground.html](/Users/force/AI-CodeX/Api-Key-Teach/playground.html) 目前支援：

- `GAS Web App URL` 手動輸入
- 透過第一組驗證碼取得範例網址
- 帳號與密碼輸入
- 管理者示範值載入
- 一般使用者示範值載入
- 模型切換
- `Use Server Key / Use My API Key`
- `systemPrompt` 覆蓋值
- `outputCharLimit`
- `沒限制` 按鈕
- 快速測試 prompt
- `Send`
- 詳細 JSON
- `Mix AI Response`

### 沒限制按鈕

`輸出字數限制` 底下有一顆：

- `沒限制`

按下後會填入：

```text
0
```

後端規則是：

- `0 = 不限制字數`

這比塞超大數字更適合教學，因為語意清楚。

---

## 十、後端回傳設計

後端現在不只是回一整段文字，還會回結構化資料：

- `reply`
- `coreReply`
- `coreReplyRaw`
- `replySections`

其中 `replySections` 會明確分出：

- `preface`
- `systemPromptLine`
- `userInputLine`
- `systemReplyLine`
- `outputLimitLine`
- `responseTimeLine`
- `conclusion`

這樣前端就不需要一直靠字串猜段落，可以更穩定顯示：

- 藍色：前言
- 紅色：AI 回覆
- 綠色：結語
- 黑色：限制、輸入、回應時間等補充資訊

---

## 十一、輸出限制邏輯

目前輸出限制不是把整段混合輸出一起亂砍，而是：

1. 先限制 AI 核心回覆
2. 再把前言 / AI 回覆 / 後語組裝回去

這樣好處是：

- 前言不會因為字數限制被吃掉
- 結語不會被不小心截斷
- 教學顯示更穩定

例如：

- `500`：最多 500 字
- `0`：不限制

---

## 十二、部署步驟

### 方法 A：從 Excel 開始

1. 開啟 [Api-Key-Teach-Template.xlsx](/Users/force/AI-CodeX/Api-Key-Teach/Api-Key-Teach-Template.xlsx)
2. 上傳到 Google Drive
3. 轉成 Google Sheet
4. 在該 Google Sheet 綁定 Apps Script
5. 貼上 [Code.gs](/Users/force/AI-CodeX/Api-Key-Teach/Code.gs)
6. 貼上 [SheetSetup.gs](/Users/force/AI-CodeX/Api-Key-Teach/SheetSetup.gs)
7. 執行 `setupTeachingSheets()`
8. 到 Script Properties 建立需要的環境變數
9. 部署成 Web App

### 方法 B：從空白 Google Sheet 開始

1. 建立空白 Google Sheet
2. 綁定 Apps Script
3. 貼上 `Code.gs` 與 `SheetSetup.gs`
4. 執行 `setupTeachingSheets()`
5. 補 Script Properties
6. 部署為 Web App

部署建議：

- 類型：`網頁應用程式`
- 執行身分：`自己`
- 權限：依教學需求選擇

---

## 十三、GitHub Pages 使用方式

如果你要放到 GitHub Pages：

- [index.html](/Users/force/AI-CodeX/Api-Key-Teach/index.html) 作為首頁
- [playground.html](/Users/force/AI-CodeX/Api-Key-Teach/playground.html) 作為執行頁

建議：

- `index.html`：當教學首頁
- `playground.html`：當實作操作頁

如果要公開給學生下載，建議注意：

- 不要直接把正式 `exec` 預設寫死在前端
- 帳號密碼不要使用真正正式環境的值
- `Use Server Key` 若會消耗你的配額，請先評估權限

更直白地說：

- 在這個專案裡，真正較高風險的通常不是說明文件本身，而是可直接打到後端的 live `exec` 網址
- 只要 `exec` 還能被外部呼叫，就可能被拿來測試、探測、寫入 logs，甚至消耗你的 server key 配額
- 因此公開版最重要的安全動作，就是不要把真實 `exec` 寫死在前端預設值或隱藏常數裡
- 本專案目前主檔已移除內建真實網址，只保留「手動貼上」與「受控發放」的設計

---

## 十四、常見問題

### Q1. 為什麼我改了 `SheetSetup.gs`，但 Google Sheet 還是舊值？

因為這支初始化腳本的策略是：

- 結構正確就保留資料
- 不主動覆寫內容

所以舊值會保留。

### Q2. 為什麼 `Use Server Key` 和 `Use My API Key` 行為不同？

因為設計上要讓學生理解：

- 後端金鑰模式可以做角色控管
- 個人金鑰模式可以放寬權限

### Q3. 為什麼前端要用 `text/plain`？

因為 GAS Web App 很容易在瀏覽器遇到 `application/json` 導致的 CORS 預檢問題，所以本專案固定用：

```text
POST + text/plain;charset=utf-8 + JSON.stringify(payload)
```

### Q4. 為什麼右上角會顯示 `Audit 119`？

那是稽核計數器，代表這支 GAS 累積已執行多少次。

---

## 十五、版本備註

目前版本：

- `v1.02`

此版重點包括：

- 多模型切換
- 管理者 / 一般使用者角色分流
- `Use Server Key / Use My API Key`
- `A / B` 驗證機制
- 結構化回傳 `replySections`
- 稽核計數器 `AUDIT_TOTAL_EXECUTIONS`
- `0 = 不限制字數`

---

## 十六、授權與備註

本專案為：

- Force Cheng × FALO
- 教學 / 內部知識資產用途

版權聲明：

```text
© Force Cheng × FALO 2026/04/01. All Rights Reserved.
```

如果要對外散佈、商業整合、移除署名或改造成正式服務，建議先重新做：

- 權限控管
- 密碼安全
- 錯誤處理
- 流量限制
- 稽核與日誌分級
