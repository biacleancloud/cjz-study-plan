# 蔡佳蓁升學計畫學習網：Google 試算表 (Level 2) 手機與電腦雙向同步設定教學

> **目的**：徹底解決「手機打勾但電腦沒同步」或「換設備打卡進度被重置」的痛點。  
> **架構**：依據三師爸《AntiGravity 基本功 EP05》教學應用程式第二階段（Level 2：前端網頁 + Google Apps Script + Google 試算表）。  
> **正式網址**：`https://biacleancloud.github.io/cjz-study-plan/`  
> **預估耗時**：約 1 ~ 2 分鐘（完全免費、免架伺服器、免管資料庫，做一次永久生效）。

---

## 🛠️ 設定步驟（只需設定一次）

### 步驟 1：建立專屬 Google 試算表
1. 打開您的 [Google 雲端硬碟](https://drive.google.com/)。
2. 點擊左上角「**＋ 新增**」 $\rightarrow$ 「**Google 試算表**」 $\rightarrow$ 「**空白試算表**」。
3. 將試算表左上角檔名命名為：`蔡佳蓁學習網打卡紀錄表`。

---

### 步驟 2：開啟 Apps Script 編輯器
1. 在剛剛建立的 Google 試算表上方功能表中，點擊「**擴充功能**」 $\rightarrow$ 「**Apps Script**」。
2. 進入 Apps Script 程式碼編輯器畫面。

---

### 步驟 3：貼上專用同步程式碼
1. 將編輯器內原本預設的 `function myFunction() {}` 全部刪除。
2. 複製以下完整程式碼（或打開專案中的 [GoogleAppsScript_Code.gs](file:///c:/Users/USER/Desktop/國中生讀書計劃(蔡佳蓁)/GoogleAppsScript_Code.gs) 複製）：

```javascript
/**
 * 蔡佳蓁升學計畫學習網 - Google Apps Script 雲端雙向同步後端 (Level 2)
 *
 * 功能：
 * 1. doGet(e): 取得最新打卡、衝刺進度、星星與獎勵設定 (JSON)
 * 2. doPost(e): 接收來自手機或電腦網頁的更新，寫入 Google 試算表，並記錄歷史日誌
 */

const SHEET_STATE_NAME = '當前狀態';
const SHEET_LOG_NAME = '打卡歷史紀錄';

function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(ss, SHEET_STATE_NAME);
    
    // 支援 GET 保存（備用容錯）：若 URL 帶有 data 參數
    if (e && e.parameter && e.parameter.data) {
      const dataStr = e.parameter.data;
      sheet.getRange('A1').setValue(dataStr);
      sheet.getRange('B1').setValue(new Date());
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', action: 'saved_via_get' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // 正常讀取狀態
    const raw = sheet.getRange('A1').getValue();
    let data = {};
    if (raw && typeof raw === 'string') {
      try {
        data = JSON.parse(raw);
      } catch (err) {
        data = {};
      }
    }
    const updatedAt = sheet.getRange('B1').getValue();
    
    const output = {
      status: 'success',
      data: data,
      updatedAt: updatedAt ? new Date(updatedAt).toISOString() : null
    };
    
    return ContentService.createTextOutput(JSON.stringify(output))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(ss, SHEET_STATE_NAME);
    const logSheet = getOrCreateSheet(ss, SHEET_LOG_NAME);
    
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e && e.parameter && e.parameter.data) {
      payload = JSON.parse(e.parameter.data);
    }
    
    const now = new Date();
    // 寫入當前完整狀態至 A1, 時間寫入 B1
    sheet.getRange('A1').setValue(JSON.stringify(payload));
    sheet.getRange('B1').setValue(now);
    
    // 在日誌表記錄一筆打卡變更日誌
    if (logSheet.getLastRow() === 0) {
      logSheet.appendRow(['時間戳記', '完成項目數', '衝刺達成數', '總星星數', '來源備註']);
    }
    const checkedCount = payload.checks ? Object.values(payload.checks).filter(Boolean).length : 0;
    const sprintCount = payload.sprint ? Object.values(payload.sprint).filter(Boolean).length : 0;
    const stars = payload.stars || 0;
    logSheet.appendRow([now, checkedCount, sprintCount, stars, payload.device || 'Web App']);
    
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      updatedAt: now.toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
```

3. 貼上後，點擊編輯器上方「**儲存專案**」圖示（💾 磁碟片圖示，或按 `Ctrl + S`）。

---

### 步驟 4：發布為網頁應用程式 (重要)
1. 點擊編輯器右上角藍色按鈕「**部署**」 $\rightarrow$ 選擇「**新增部署作業**」。
2. 點擊左側齒輪「選取類型」 $\rightarrow$ 選擇「**網頁應用程式**」。
3. 填寫部署設定（**非常重要！**）：
   - **說明**：`蔡佳蓁學習網雙向同步後端`
   - **以何者身分執行**：選擇「**我 (您的 Gmail 帳號)**」
   - **誰可以存取**：務必改選為「**所有人 (Anyone)**」（這樣手機或電腦網頁不用登入即可直接同步打卡狀態）
4. 點擊右下角「**部署**」。
5. 若首次部署跳出授權提示：
   - 點擊「**審查權限**」 $\rightarrow$ 選擇您的 Google 帳號。
   - 若出現「Google 尚未驗證此應用程式」警告，點擊左下角「**進階 (Advanced)**」 $\rightarrow$ 點擊下方「**前往『未命名專案』(不安全)**」 $\rightarrow$ 點擊「**允許 (Allow)**」。
6. 部署完成後，畫面會顯示一串「**網頁應用程式網址 (Web App URL)**」（開頭為 `https://script.google.com/macros/s/AKfycb.../exec`）。
7. 點擊「**複製**」該網址。

---

### 步驟 5：啟用三端連線（電腦、爸爸手機、女兒手機）

#### 1. 電腦端啟用：
1. 打開本機的 [index.html](file:///c:/Users/USER/Desktop/國中生讀書計劃(蔡佳蓁)/index.html) 或線上網址 `https://biacleancloud.github.io/cjz-study-plan/`。
2. 點擊頂部狀態列的「**⚙️ 雲端同步設定**」按鈕。
3. 將剛剛複製的 Web App 網址貼入輸入框中，點擊「**💾 儲存並立即上傳**」。
4. 看到頂部狀態膠囊變為「**🟢 雲端已同步**」，即代表連線成功！

#### 2. 手機端無縫啟用（超級防呆，免手動打字！）：
將以下專屬網址傳到 LINE 給**您的手機**與**蔡佳蓁的手機**：

```text
https://biacleancloud.github.io/cjz-study-plan/?gas=您剛剛複製的AppsScript網址
```

- **手機點開瞬間自動綁定**：完全不用在手機上複製貼上長串網址，一開即同步！
- **將網頁加入手機主畫面（秒變專屬 App）**：
  - **iPhone (Safari)**：點擊底部的「分享」圖示 $\rightarrow$ 選擇「**加入主畫面**」。
  - **Android (Chrome)**：點擊右上角「⋮」選單 $\rightarrow$ 選擇「**加到主螢幕**」或「**安裝應用程式**」。
  - 手機桌面上就會產生一顆「蔡佳蓁學習網」專屬元氣頭像圖示，就像原生 App 一樣，隨開隨打卡！

---

## ⚡ 三端即時聯動效果

- **佳蓁在書桌前用手機打卡**：勾選完 1 秒內自動推送到 Google 試算表。
- **爸爸在電腦或手機查看**：打開網頁、切回視窗、或點擊頂部「🔄 **立即同步**」，打卡進度、星星存款、兌換記錄瞬間全部同步呈現！
- **歷史日誌存檔**：Google 試算表第二張分頁「`打卡歷史紀錄`」會忠實記錄每次打卡的時間戳記與累計星星數，努力歷程一目瞭然！
