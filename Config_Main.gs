/**
 * SMCI Database to Google Contacts Sync System
 * Version: 27.0 (Security Update: Use ScriptProperties for IDs)
 * File: Config_Main.gs
 */

// ==========================================
// 【設定】秘密情報（SS-IDなど）はスクリプトプロパティから読み込みます
// ==========================================
const START_ROW = 4;
const SCRIPT_VERSION = "v27.0";
const BASE_DELIMITER = "SM://SMCI_AutoUpdater";
const SYSTEM_LABEL = "SMCI Auto Updater"; 
const TIME_LIMIT_MS = 270 * 1000; 

// ▼ プロパティストアからIDリストを取得する関数
function getTargetSpreadsheetIds() {
  const props = PropertiesService.getScriptProperties();
  const idsString = props.getProperty('TARGET_SS_IDS');
  
  if (!idsString) {
    console.error("❌ エラー: スクリプトプロパティ 'TARGET_SS_IDS' が設定されていません。_Setup_Secrets.gsを実行しましたか？");
    return [];
  }
  
  // カンマ区切りの文字列を配列に戻して返す
  return idsString.split(',').map(s => s.trim());
}

// グローバルキャッシュ
let groupMap = new Map();
let mapSMCI11 = new Map();
let mapSMCI9 = new Map();
let mapTMDb = new Map();
let mapEmail = new Map();
let mapName = new Map(); 

let executionStartTime = 0;

function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('SMCI連携')
      .addItem('【設定】自動更新トリガーをセット (初回のみ)', 'setupAutomatedTriggers')
      .addSeparator()
      .addItem('【手動】クイック更新 (最新/末尾のみ)', 'runQuickSync')
      .addItem('【手動】フル更新 (中断・再開対応)', 'mainSyncProcess')
      .addSeparator()
      .addItem('進行状況をリセット', 'resetSyncStatus')
      .addSeparator()
      .addItem('☁️ GitHubから最新コードを取得', 'pullFromGitHub') // ★メニューに追加
      .addToUi();
  } catch (e) {
    console.log("Running in standalone/trigger mode. Menu not created.");
  }
}

function setupAutomatedTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  
  ScriptApp.newTrigger('runQuickSync').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('mainSyncProcess').timeBased().everyHours(1).create();
    
  console.log("✅ Triggers set successfully.");
}

function runQuickSync() {
  console.log("🚀 Quick Sync Started...");
  
  // ★重要変更：IDを直接書かず、関数経由で取得
  let fileIds = getTargetSpreadsheetIds();
  
  if (!fileIds || fileIds.length === 0) {
    // ID設定がない、かつ手動実行ならActiveSheetを使う
    try { fileIds = [SpreadsheetApp.getActiveSpreadsheet().getId()]; } 
    catch (e) { console.warn("No Target IDs and no active sheet. Skipped."); return; }
  }

  initAllMaps();

  fileIds.forEach((fileId, idx) => {
    try {
      const ss = SpreadsheetApp.openById(fileId);
      const sheet = ss.getSheets()[0];
      const maxRow = sheet.getLastRow();
      
      if (maxRow < START_ROW) return;

      const colAValues = sheet.getRange(START_ROW, 1, maxRow - START_ROW + 1, 1).getValues();
      let trueLastRow = -1;
      for (let i = colAValues.length - 1; i >= 0; i--) {
        const val = String(colAValues[i][0]).trim();
        if (val && val !== "#N/A") { 
          trueLastRow = START_ROW + i;
          break;
        }
      }

      if (trueLastRow === -1) return;

      const processStart = Math.max(START_ROW, trueLastRow - 4);
      console.log(`[QuickSync] File ${idx+1}: Processing rows ${processStart} to ${trueLastRow}`);

      for (let r = processStart; r <= trueLastRow; r++) {
        processSingleRow(sheet, r);
      }
    } catch (e) {
      console.error(`QuickSync Failed for ID ${fileId}: ${e.message}`);
    }
  });
  console.log("✅ Quick Sync Completed.");
}

function mainSyncProcess() {
  executionStartTime = new Date().getTime();
  const props = PropertiesService.getScriptProperties();
  
  // ★重要変更：IDを直接書かず、関数経由で取得
  let fileIds = getTargetSpreadsheetIds();

  if (!fileIds || fileIds.length === 0) {
    try { fileIds = [SpreadsheetApp.getActiveSpreadsheet().getId()]; }
    catch (e) { console.error("No Target IDs."); return; }
  }

  let currentFileIndex = parseInt(props.getProperty('SYNC_FILE_INDEX') || '0');
  let currentRowIndex = parseInt(props.getProperty('SYNC_ROW_INDEX') || START_ROW.toString());
  
  if (currentFileIndex >= fileIds.length) {
    currentFileIndex = 0; currentRowIndex = START_ROW;
  }

  console.log(`🔄 Full Sync Resumed... FileIndex: ${currentFileIndex}, Row: ${currentRowIndex}`);
  initAllMaps();

  for (let i = currentFileIndex; i < fileIds.length; i++) {
    const fileId = fileIds[i];
    let sheet;
    try {
      sheet = SpreadsheetApp.openById(fileId).getSheets()[0];
      console.log(`Processing File [${i + 1}/${fileIds.length}]: ${sheet.getParent().getName()}`);
    } catch (e) {
      console.error(`Skipping ID: ${fileId}`);
      continue;
    }

    const lastRow = sheet.getLastRow();
    const loopStart = (i === currentFileIndex) ? currentRowIndex : START_ROW;
    
    for (let r = loopStart; r <= lastRow; r++) {
      if (isTimeUp()) {
        props.setProperty('SYNC_FILE_INDEX', i.toString());
        props.setProperty('SYNC_ROW_INDEX', r.toString());
        console.warn(`⏳ Time Limit. Paused at File[${i}], Row[${r}].`);
        return; 
      }
      processSingleRow(sheet, r);
    }
  }

  resetSyncStatus();
  console.log("✅ Full Sync Completed.");
}

function resetSyncStatus() {
  PropertiesService.getScriptProperties().deleteProperty('SYNC_FILE_INDEX');
  PropertiesService.getScriptProperties().deleteProperty('SYNC_ROW_INDEX');
}

function isTimeUp() {
  return (new Date().getTime() - executionStartTime) > TIME_LIMIT_MS;
}
