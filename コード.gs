/**
 * SMCI Database to Google Contacts Sync System
 * Version: 23.0 (Update: School->CF, SchoolFax Label, Kata->Hira, New Timestamp)
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
const SCRIPT_VERSION = "v23.0";
const BASE_DELIMITER = "SM://SMCI_AutoUpdater";
const SYSTEM_LABEL = "SMCI Auto Updater"; 
const TIME_LIMIT_MS = 270 * 1000; 

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
      .addToUi();
  } catch (e) {
    console.log("Running in standalone/trigger mode. Menu not created.");
  }
}

function setupAutomatedTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  
  // 1. クイック更新 (5分ごと)
  ScriptApp.newTrigger('runQuickSync')
    .timeBased()
    .everyMinutes(5)
    .create();

  // 2. フル更新 (1時間ごと)
  ScriptApp.newTrigger('mainSyncProcess')
    .timeBased()
    .everyHours(1) 
    .create();
    
  console.log("✅ Triggers set successfully: QuickSync(5min), FullSync(1hour).");
}

/**
 * ■ クイック更新
 */
function runQuickSync() {
  console.log("🚀 Quick Sync Started...");
  let fileIds = TARGET_SPREADSHEET_IDS;
  if (!fileIds || fileIds.length === 0) {
    try { fileIds = [SpreadsheetApp.getActiveSpreadsheet().getId()]; } 
    catch (e) { console.warn("No Target IDs. Skipped."); return; }
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
  initAllMaps();

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
      processSingleRow(sheet, r);
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

// --- Initialization ---
function initAllMaps() {
  console.log("Initializing Index...");
  initGroupMap(); 
  let pageToken = null;
  try {
    do {
      const res = People.People.Connections.list('people/me', {
        personFields: 'names,emailAddresses,userDefined,urls,memberships', 
        pageSize: 1000,
        pageToken: pageToken
      });
      if (res.connections) {
        res.connections.forEach(person => {
          if (person.userDefined) {
            person.userDefined.forEach(ud => {
              if (ud.key === 'SMCI11' && ud.value) mapSMCI11.set(ud.value, person);
              if (ud.key === 'SMCI9' && ud.value) mapSMCI9.set(ud.value, person);
            });
          }
          if (person.urls) {
            person.urls.forEach(urlObj => {
              if (urlObj.value && urlObj.value.includes('themoviedb.org')) mapTMDb.set(urlObj.value, person);
            });
          }
          if (person.emailAddresses) {
            person.emailAddresses.forEach(email => {
              if (email.value) mapEmail.set(email.value.toLowerCase(), person);
            });
          }
          if (person.names) {
             const disp = person.names[0].displayName || "";
             if (disp) mapName.set(disp, person);
          }
        });
      }
      pageToken = res.nextPageToken;
    } while (pageToken);
  } catch (e) {
    console.error("Critical Error during initialization: " + e.message);
    throw e;
  }
}

// --- Helper Functions ---
function initGroupMap() {
  try {
    let pageToken = null;
    do {
      const resp = People.ContactGroups.list({ pageSize: 1000, pageToken: pageToken });
      if (resp.contactGroups) {
        resp.contactGroups.forEach(g => {
          if (g.formattedName && g.resourceName) groupMap.set(g.formattedName, g.resourceName);
        });
      }
      pageToken = resp.nextPageToken;
    } while (pageToken);
  } catch (e) { console.warn("Failed to init groups: " + e.message); }
}

function getOrCreateGroupResourceName(labelName) {
  if (!labelName) return null;
  const name = String(labelName).trim();
  if (groupMap.has(name)) return groupMap.get(name);
  try {
    const newGroup = People.ContactGroups.create({ contactGroup: { name: name } });
    const resName = newGroup.resourceName;
    groupMap.set(name, resName);
    return resName;
  } catch (e) { return null; }
}

function cleanData(val) {
  if (val === null || val === undefined) return "";
  if (val instanceof Date) return val;
  if (val.toString().includes("CellImage") || val.toString() === "Obj") return "";
  const str = String(val).trim();
  const ignoreList = ["NIL", "UNK", "取得中", "#N/A", "N/A"];
  if (ignoreList.includes(str)) return "";
  return str;
}
function convertKokiToDate(val) {
  if (!val) return null;
  let date;
  if (val instanceof Date) date = new Date(val);
  else {
    const str = String(val).trim();
    if (!str.match(/[\d]{2,4}[\/\-]/)) return null;
    date = new Date(str);
  }
  if (isNaN(date.getTime())) return null;
  const year = date.getFullYear();
  if (year > 2600) date.setFullYear(year - 660);
  return { year: date.getFullYear(), month: date.getMonth() + 1, day: date.getDate() };
}
function isDateString(val) {
  if (!val) return false;
  if (val instanceof Date) return true;
  return /[\d]{4}[\/\-][\d]{1,2}/.test(String(val));
}
function extractImageUrl(formula) {
  if (!formula) return null;
  const match = formula.match(/IMAGE\s*\(\s*["']([^"']+)["']/i);
  return match ? match[1] : null;
}

// カタカナ→ひらがな変換関数
function kataToHira(str) {
  if (!str) return "";
  return str.replace(/[\u30a1-\u30f6]/g, function(match) {
    var chr = match.charCodeAt(0) - 0x60;
    return String.fromCharCode(chr);
  });
}

// タイムスタンプ生成 (皇紀 + JST時刻)
// 例: SM-ﾄ26860120075505
function getTimestampString() {
  const now = new Date(); // GASのタイムゾーン(JST)で取得される前提
  const kokiYear = now.getFullYear() + 660;
  
  const pad = (n) => String(n).padStart(2, '0');
  const mm = pad(now.getMonth() + 1);
  const dd = pad(now.getDate());
  const hh = pad(now.getHours());
  const mi = pad(now.getMinutes());
  const ss = pad(now.getSeconds());

  return `SM-ﾄ${kokiYear}${mm}${dd}${hh}${mi}${ss}`;
}

// --- Single Row Process ---
function processSingleRow(sheet, rowNumber) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rowData = sheet.getRange(rowNumber, 1, 1, lastCol).getValues()[0];
  const rowFormulas = sheet.getRange(rowNumber, 1, 1, lastCol).getFormulas()[0];

  let d = {}; 
  let f = {};
  headers.forEach((id, i) => { d[id] = rowData[i]; f[id] = rowFormulas[i]; });

  const valSMCI11 = cleanData(d['SMCI-XX48']); 
  const valSMCI9  = cleanData(d['SMCI-XX01']); 
  const valTMDb   = cleanData(d['SMCI-XX43']); 
  
  if (!valSMCI11 && !valSMCI9) return;

  try {
    let existingPerson = null;
    let foundBy = "";

    if (valSMCI11 && mapSMCI11.has(valSMCI11)) { existingPerson = mapSMCI11.get(valSMCI11); foundBy = "SMCI11(Map)"; }
    if (!existingPerson && valSMCI9 && mapSMCI9.has(valSMCI9)) { existingPerson = mapSMCI9.get(valSMCI9); foundBy = "SMCI9(Map)"; }
    if (!existingPerson && valTMDb && mapTMDb.has(valTMDb)) { existingPerson = mapTMDb.get(valTMDb); foundBy = "TMDb(Map)"; }
    if (!existingPerson) {
      const emailsToCheck = [];
      if (d['SMCI-XX22']) cleanData(d['SMCI-XX22']).split(';').forEach(e => emailsToCheck.push(e.trim()));
      if (d['SMCI-XX23']) cleanData(d['SMCI-XX23']).split(';').forEach(e => emailsToCheck.push(e.trim()));
      for (let email of emailsToCheck) {
        const lower = email.toLowerCase();
        if (mapEmail.has(lower)) { existingPerson = mapEmail.get(lower); foundBy = `Email(${email})`; break; }
      }
    }

    let labelIds = [];
    const labelName = cleanData(d['SMCI-XX61']);
    if (labelName) labelName.split(';').forEach(ln => { const id = getOrCreateGroupResourceName(ln); if (id) labelIds.push(id); });
    const sysLabelId = getOrCreateGroupResourceName(SYSTEM_LABEL);
    if (sysLabelId) labelIds.push(sysLabelId);

    const buildResult = buildPersonPayload(d, existingPerson, valSMCI11, valSMCI9, labelIds);
    const personData = buildResult.payload;

    let resourceName = "";
    if (existingPerson) {
      resourceName = existingPerson.resourceName;
      if (valSMCI11) mapSMCI11.set(valSMCI11, existingPerson);
      People.People.updateContact(personData, resourceName, {
        updatePersonFields: 'names,nicknames,organizations,emailAddresses,phoneNumbers,addresses,birthdays,events,urls,userDefined,biographies,memberships,relations'
      });
      console.log(`[Updated] File: ${sheet.getParent().getName()} Row: ${rowNumber} (${valSMCI11 || valSMCI9}): Found by ${foundBy}.`);
    } else {
      const created = People.People.createContact(personData);
      resourceName = created.resourceName;
      if (valSMCI11) mapSMCI11.set(valSMCI11, created);
      console.log(`[Created] File: ${sheet.getParent().getName()} Row: ${rowNumber} (${valSMCI11 || valSMCI9}): New Contact.`);
    }

    const photoUrl = extractImageUrl(f['SMCI-XX02']) || (String(d['SMCI-XX02']).startsWith("http") ? String(d['SMCI-XX02']) : null);
    if (photoUrl && resourceName) updateContactPhoto(resourceName, photoUrl);

  } catch (e) {
    console.error(`Error File: ${sheet.getParent().getName()} Row: ${rowNumber}: ` + e.message);
  }
}

function updateContactPhoto(resourceName, url) {
  try {
    const response = UrlFetchApp.fetch(url);
    const photoData = Utilities.base64Encode(response.getBlob().getBytes());
    People.People.updateContactPhoto({ photoBytes: photoData }, resourceName);
  } catch (e) { console.warn('Photo fail: ' + e.message); }
}

// --- Payload Builder (v23.0) ---
function buildPersonPayload(d, existing, valSMCI11, valSMCI9, newLabelIds) {
  let updates = []; 

  // 1. Name & Furigana (Kata->Hira)
  let fName = cleanData(d['SMCI-XX05']); 
  let gName = cleanData(d['SMCI-XX07']); 
  let mName = cleanData(d['SMCI-XX06']); 

  if (!fName && !gName) {
    const enLast  = cleanData(d['SMCI-XX13']);
    const enFirst = cleanData(d['SMCI-XX11']);
    const enMid   = cleanData(d['SMCI-XX12']);
    if (enLast || enFirst) {
      fName = enLast; gName = enFirst; mName = enMid;
      updates.push("Name(English)");
    } else {
      const dispName = cleanData(d['SMCI-XX03']);
      if (dispName) { gName = dispName; updates.push("Name(Display)"); }
    }
  } else {
    updates.push("Name(Japanese)");
  }

  // フリガナをひらがなに変換してセット
  const nameObj = {
    familyName: fName, 
    givenName: gName, 
    middleName: mName,
    honorificPrefix: cleanData(d['SMCI-XX04']), 
    honorificSuffix: cleanData(d['SMCI-XX14']),
    phoneticFamilyName: kataToHira(cleanData(d['SMCI-XX08'])), 
    phoneticGivenName: kataToHira(cleanData(d['SMCI-XX10']))
  };

  const nick = cleanData(d['SMCI-XX15']);

  const payload = {
    names: [nameObj],
    nicknames: nick ? [{ value: nick }] : [],
    organizations: [],
    emailAddresses: existing ? [...(existing.emailAddresses || [])] : [],
    phoneNumbers: existing ? [...(existing.phoneNumbers || [])] : [],
    addresses: existing ? [...(existing.addresses || [])] : [],
    memberships: existing ? [...(existing.memberships || [])] : [], 
    relations: existing ? [...(existing.relations || [])] : [], 
    birthdays: [], events: [], urls: [], userDefined: [], biographies: []
  };

  // 2. Organization (Work only) - School moved to CF
  const company = cleanData(d['SMCI-XX19']);
  if (company) {
    payload.organizations.push({ name: company, title: cleanData(d['SMCI-XX21']), department: cleanData(d['SMCI-XX20']), type: 'work' });
  }

  // 3. Labels
  if (newLabelIds && newLabelIds.length > 0) {
    const existingGroupIds = payload.memberships.map(m => m.contactGroupMembership.contactGroupResourceName);
    newLabelIds.forEach(id => {
      if (!existingGroupIds.includes(id)) payload.memberships.push({ contactGroupMembership: { contactGroupResourceName: id
