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
