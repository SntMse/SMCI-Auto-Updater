/**
 * SMCI Database to Google Contacts Sync System
 * Version: 24.0 (Modularized)
 * File: Config_Main.gs
 */

// ==========================================
// 【設定】対象のスプレッドシートIDリスト
const TARGET_SPREADSHEET_IDS = [
  "1GQQy2hAvZ8afRxySWcpayQLaOIN7MkrYmFcKrauKOiw", //SMCI-PKJ管理表
  "1610hT2vzaY-7d9IXktwv-eTWeVyBh5p_cxiHHYwI8g8", //SMCI-PNY管理表
  "1mUnPSb1vzkAgLkv4FA_8VLI-o3JDzN01eg3U3XV3LBk", //SMCI-PFA管理表
  "10TF7E4yLpYzy7RXJwWpoMbnfY-oc2Kfm_W2AJR1xhHU", //SMCI-PYT管理表
  "1GbNdY1UdQNDCQ5feVY65jnqIGsjxTTegla_vk0s6MN0"  //SMCI管理表
]; 
// ==========================================

const START_ROW = 4;
const SCRIPT_VERSION = "v24.0";
const BASE_DELIMITER = "SM://SMCI_AutoUpdater";
const SYSTEM_LABEL = "SMCI Auto Updater"; 
const TIME_LIMIT_MS = 270 * 1000; 

// グローバルキャッシュ (API_Managerで使用)
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
    
  console.log("✅ Triggers set successfully: QuickSync(5min), FullSync(1hour).");
}

/**
 * ■ クイック更新 (ロジックはLogic_Process.gsへ委譲)
 */
function runQuickSync() {
  console.log("🚀 Quick Sync Started...");
  let fileIds = TARGET_SPREADSHEET_IDS;
  if (!fileIds || fileIds.length === 0) {
    try { fileIds = [SpreadsheetApp.getActiveSpreadsheet().getId()]; } 
    catch (e) { console.warn("No Target IDs. Skipped."); return; }
  }

  initAllMaps(); // API_Manager.gs

  fileIds.forEach((fileId, idx) => {
    try {
      const ss = SpreadsheetApp.openById(fileId);
      const sheet = ss.getSheets()[0];
      const maxRow = sheet.getLastRow();
      
      if (maxRow < START_ROW) return;

      // A列基準で最終行を特定
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
        processSingleRow(sheet, r); // Logic_Process.gs
      }
    } catch (e) {
      console.error(`QuickSync Failed for ID ${fileId}: ${e.message}`);
    }
  });
  console.log("✅ Quick Sync Completed.");
}

/**
 * ■ フル更新
 */
function mainSyncProcess() {
  executionStartTime = new Date().getTime();
  const props = PropertiesService.getScriptProperties();
  
  let fileIds = TARGET_SPREADSHEET_IDS;
  if (!fileIds || fileIds.length === 0) {
    try { fileIds = [SpreadsheetApp.getActiveSpreadsheet().getId()]; }
    catch (e) { console.error("No Target Spreadsheet IDs."); return; }
  }

  let currentFileIndex = parseInt(props.getProperty('SYNC_FILE_INDEX') || '0');
  let currentRowIndex = parseInt(props.getProperty('SYNC_ROW_INDEX') || START_ROW.toString());
  
  if (currentFileIndex >= fileIds.length) {
    currentFileIndex = 0; currentRowIndex = START_ROW;
  }

  console.log(`🔄 Full Sync Resumed... FileIndex: ${currentFileIndex}, Row: ${currentRowIndex}`);
  initAllMaps(); // API_Manager.gs

  for (let i = currentFileIndex; i < fileIds.length; i++) {
    const fileId = fileIds[i];
    let sheet;
    try {
      sheet = SpreadsheetApp.openById(fileId).getSheets()[0];
      console.log(`Processing File [${i + 1}/${fileIds.length}]: ${sheet.getParent().getName()}`);
    } catch (e) {
      console.error(`Can't open spreadsheet ID: ${fileId}. Skipping.`);
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
      processSingleRow(sheet, r); // Logic_Process.gs
    }
  }

  resetSyncStatus();
  console.log("✅ Full Sync Completed.");
}

function resetSyncStatus() {
  PropertiesService.getScriptProperties().deleteProperty('SYNC_FILE_INDEX');
  PropertiesService.getScriptProperties().deleteProperty('SYNC_ROW_INDEX');
  console.log("Sync status reset.");
}

function isTimeUp() {
  return (new Date().getTime() - executionStartTime) > TIME_LIMIT_MS;
}
