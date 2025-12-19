# 📊 Google Apps Script LINE Bot - 全能記帳與理財助手

這是一個基於 Google Apps Script (GAS) 開發的 LINE Bot，結合了 Google Sheets 作為資料庫。它不僅能幫助你輕鬆記帳，還能即時查詢台銀匯率與台股股價，並自動生成收支圓餅圖報表。

## 🌟 功能特色

  * **📝 互動式記帳**：透過按鈕引導，輕鬆記錄收入與支出（類別 -\> 品項 -\> 金額）。
  * **🗂️ 類別管理**：可自定義新增或刪除收支類別，靈活符合個人需求。
  * **📊 視覺化報表**：自動計算今日、本月或今年的收支，並生成圓餅圖（Pie Chart）回傳。
  * **💱 即時匯率**：爬取台灣銀行（Bot of Taiwan）即時現金與即期匯率。
  * **📈 台股查詢**：輸入股票代碼，即時抓取證交所盤中股價、漲跌幅與成交量。
  * **☁️ 雲端儲存**：所有資料儲存在 Google Sheets，圖表儲存在 Google Drive，安全且易於備份。

-----

## 🚀 部署教學 (Step-by-Step)

請按照以下步驟將程式碼部署到你的 Google 帳號中。

### 步驟 1：準備 Google Sheet

1.  建立一個新的 [Google Sheet](https://sheets.google.com)。
2.  將試算表命名為「記帳機器人資料庫」或是你喜歡的名字。
3.  **複製網址中的 ID**。
      * 網址格式為：`https://docs.google.com/spreadsheets/d/`**`這裡就是ID`**`/edit...`
      * 請記下這串 ID，稍後會用到。

### 步驟 2：建立 Google Apps Script 專案

1.  在剛剛建立的 Google Sheet 中，點擊上方選單的 **「擴充功能」 (Extensions)** \> **「Apps Script」**。
2.  這會開啟一個新的程式碼編輯器視窗。
3.  將原本編輯器內的 `function myFunction() {...}` 刪除。
4.  **複製並貼上** 你提供的完整程式碼到編輯器中。
5.  按 `Ctrl + S` 儲存專案，命名為「LINE Bot Backend」。

### 步驟 3：安裝 Cheerio 函式庫 (用於匯率爬蟲)

由於程式碼中使用了 `Cheerio` 來解析網頁，必須手動加入此函式庫：

1.  在 Apps Script 編輯器左側，找到 **「程式庫」 (Libraries)** 旁邊的 `+` 號。
2.  在「指令碼 ID」欄位輸入：
    ```text
    1ReeQ6WO8kKNxagqrJnOf29QhBOcrg5NQD_bA7XwnpiKqA9jL00g9vA2I
    ```
3.  點擊 **「查詢」**。
4.  版本選擇最新版，點擊 **「新增」**。

### 步驟 4：設定 LINE Developers Channel

1.  前往 [LINE Developers Console](https://developers.line.biz/) 並登入。
2.  建立一個新的 Provider 和 Channel (Messaging API)。
3.  在 **Messaging API** 分頁中：
      * 產生並複製 **Channel access token (long-lived)**。
      * 在 **Basic settings** 分頁中，複製 **Channel secret**。
4.  回到 Google Apps Script 編輯器，將程式碼最上方的變數填入：

<!-- end list -->

```javascript
const LINE_CHANNEL_ACCESS_TOKEN = '貼上你的 Channel Access Token';
const LINE_CHANNEL_SECRET = '貼上你的 Channel Secret';
const GOOGLE_SHEET_ID = '貼上步驟 1 取得的 Google Sheet ID';
```

### 步驟 5：部署為 Web App

這是最重要的一步，請務必設定正確：

1.  點擊編輯器右上角的 **「部署」 (Deploy)** \> **「新增部署作業」 (New deployment)**。
2.  點擊左側齒輪圖示，選擇 **「網頁應用程式」 (Web app)**。
3.  設定如下：
      * **說明**：LINE Bot V1 (隨意填寫)。
      * **執行身分 (Execute as)**：**我 (Me)**。
      * **誰可以存取 (Who has access)**：**所有人 (Anyone)**。\<small\>*(注意：這是為了讓 LINE 伺服器能夠呼叫你的腳本，請務必選「所有人」)*\</small\>
4.  點擊 **「部署」**。
5.  系統會要求 **授權存取** (Authorize access)，請登入你的 Google 帳號，並在警告畫面點選 **「Advanced」** \> **「Go to ... (unsafe)」** \> **「Allow」**。
6.  部署成功後，**複製「網頁應用程式網址」 (Web App URL)**。

### 步驟 6：設定 LINE Webhook

1.  回到 [LINE Developers Console](https://developers.line.biz/) 的 **Messaging API** 分頁。
2.  找到 **Webhook settings**。
3.  貼上剛剛複製的 **Web App URL**。
4.  點擊 **Verify**，如果出現 `Success` 代表連線成功。
5.  開啟 **Use webhook** 開關。
6.  (建議) 在下方的 **LINE Official Account features** 區塊，點擊 Edit，將 **Auto-reply messages (自動回應)** 設為 **Disabled**，**Greeting messages (加入好友歡迎詞)** 依喜好設定。

-----

## 📖 使用說明 (指令列表)

你可以直接在 LINE 聊天室輸入以下文字指令，或透過手機版選單操作。

| 觸發關鍵字 | 功能說明 |
| :--- | :--- |
| **記帳** | 開啟記帳流程。選擇「收入/支出」 \> 選擇類別 \> 輸入項目 \> 輸入金額。 |
| **收支類別** | 查看目前的分類，並提供新增或刪除類別的按鈕。 |
| **收支報表** | 選擇產生「今日」、「本月」或「今年」的統計報表與圓餅圖。 |
| **匯率** | 選擇查詢美金、日圓、歐元等幣別的即時匯率。 |
| **股票** | 機器人會提示輸入股票代碼（例如：2330），回傳即時股價資訊。 |

## Demo影片



https://github.com/user-attachments/assets/bcbf7254-72e4-4e1e-aa1f-7f79a93f946a


