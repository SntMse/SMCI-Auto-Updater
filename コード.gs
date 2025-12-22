/**
 * SMCI Database to Google Contacts Sync System
 * Version: 21.2 (Fix: QuickSync targets strictly last 5 rows based on Col A)
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
const SCRIPT_VERSION = "v21.2";
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
      .addItem('【設定】自動更新トリガーをセット (初回のみ実行)', 'setupAutomatedTriggers')
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
  console.log("このログが表示されていれば設定完了です。");
}

/**
 * ■ クイック更新 (5分に1回実行) - v21.2 改修版
 * 全ファイルの「A列(SMCI-XX48)が空欄でない最後の行」を特定し、
 * その行を含む直近5行のみを更新します。
 */
function runQuickSync() {
  console.log("🚀 Quick Sync Started (Targeting strict bottom rows)...");
  
  let fileIds = TARGET_SPREADSHEET_IDS;
  if (!fileIds || fileIds.length === 0) {
    try {
      fileIds = [SpreadsheetApp.getActiveSpreadsheet().getId()];
    } catch (e) {
      console.warn("No Target IDs and no Active Sheet. Quick sync skipped.");
      return;
    }
  }

  initAllMaps();

  fileIds.forEach((fileId, idx) => {
    try {
      const ss = SpreadsheetApp.openById(fileId);
      const sheet = ss.getSheets()[0];
      const maxRow = sheet.getLastRow();
      
      if (maxRow < START_ROW) {
        console.log(`[QuickSync] File ${idx+1}: Empty or header only. Skipped.`);
        return;
      }

      // 1. A列(1列目)の値を一括取得して、真の最終行を探す
      // getRange(row, col, numRows, numCols)
      const colAValues = sheet.getRange(START_ROW, 1, maxRow - START_ROW + 1, 1).getValues();
      
      let trueLastRow = -1;

      // 下から上にループして、値がある最初の行を探す
      for (let i = colAValues.length - 1; i >= 0; i--) {
        const val = String(colAValues[i][0]).trim();
        if (val && val !== "#N/A") { 
          // 配列インデックス i に START_ROW を足すと実在する行番号になる
          trueLastRow = START_ROW + i;
          break;
        }
      }

      // データが見つからなかった場合
      if (trueLastRow === -1) {
        console.log(`[QuickSync] File ${idx+1}: No valid data in Col A. Skipped.`);
        return;
      }

      // 2. 処理範囲を決定 (真の最終行から4行戻る)
      const processStart = Math.max(START_ROW, trueLastRow - 4);
      
      console.log(`[QuickSync] File ${idx+1}: ${ss.getName()} - Processing rows ${processStart} to ${trueLastRow} (Based on Col A)`);

      // 3. 該当範囲を実行
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
  
  let fileIds = TARGET_SPREADSHEET_IDS;
  if (!fileIds || fileIds.length === 0) {
    try {
      fileIds = [SpreadsheetApp.getActiveSpreadsheet().getId()];
    } catch (e) {
      console.error("No Target Spreadsheet IDs set. Cannot run Full Sync via trigger.");
      return;
    }
  }

  let currentFileIndex = parseInt(props.getProperty('SYNC_FILE_INDEX') || '0');
  let currentRowIndex = parseInt(props.getProperty('SYNC_ROW_INDEX') || START_ROW.toString());
  
  if (currentFileIndex >= fileIds.length) {
    currentFileIndex = 0; currentRowIndex = START_ROW;
  }

  console.log(`🔄 Full Sync Started/Resumed... FileIndex: ${currentFileIndex}, Row: ${currentRowIndex}`);

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
        console.warn(`⏳ Time Limit. Paused at File[${i}], Row[${r}]. Will resume in next cycle.`);
        return; 
      }
      processSingleRow(sheet, r);
    }
  }

  resetSyncStatus();
  console.log("✅ Full Sync Completed for all files.");
}

function resetSyncStatus() {
  PropertiesService.getScriptProperties().deleteProperty('SYNC_FILE_INDEX');
  PropertiesService.getScriptProperties().deleteProperty('SYNC_ROW_INDEX');
  console.log("Sync status reset.");
}

function isTimeUp() {
  return (new Date().getTime() - executionStartTime) > TIME_LIMIT_MS;
}

function testSyncFirstFive() { runQuickSync(); }


// --- Initialization ---
function initAllMaps() {
  console.log("Initializing: Fetching ALL contacts from Google to build index...");
  initGroupMap(); 
  let pageToken = null;
  let count = 0;
  try {
    do {
      const res = People.People.Connections.list('people/me', {
        personFields: 'names,emailAddresses,userDefined,urls,memberships', 
        pageSize: 1000,
        pageToken: pageToken
      });
      const connections = res.connections;
      if (connections) {
        connections.forEach(person => {
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
        count += connections.length;
      }
      pageToken = res.nextPageToken;
    } while (pageToken);
    console.log(`Index Built: Loaded ${count} contacts.`);
  } catch (e) {
    console.error("Critical Error during initialization: " + e.message);
    throw e;
  }
}

// --- Group (Label) Helpers ---
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

// --- Helper Functions ---
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
function getKokiDateString() {
  const now = new Date();
  return `${now.getFullYear() + 660}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
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

    // Identification
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

    // Label
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

// --- Payload Builder ---
function buildPersonPayload(d, existing, valSMCI11, valSMCI9, newLabelIds) {
  let updates = []; 

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

  const nameObj = {
    familyName: fName, 
    givenName: gName, 
    middleName: mName,
    honorificPrefix: cleanData(d['SMCI-XX04']), 
    honorificSuffix: cleanData(d['SMCI-XX14']),
    phoneticFamilyName: cleanData(d['SMCI-XX08']), 
    phoneticGivenName: cleanData(d['SMCI-XX10'])
  };

  const nick = cleanData(d['SMCI-XX15']);

  const payload = {
    names: [nameObj],
    nicknames: nick ? [{ value: nick }] : [],
    organizations: [{ name: cleanData(d['SMCI-XX19']), title: cleanData(d['SMCI-XX21']), department: cleanData(d['SMCI-XX20']) }],
    emailAddresses: existing ? [...(existing.emailAddresses || [])] : [],
    phoneNumbers: existing ? [...(existing.phoneNumbers || [])] : [],
    addresses: existing ? [...(existing.addresses || [])] : [],
    memberships: existing ? [...(existing.memberships || [])] : [], 
    relations: existing ? [...(existing.relations || [])] : [], 
    birthdays: [], events: [], urls: [], userDefined: [], biographies: []
  };

  if (newLabelIds && newLabelIds.length > 0) {
    const existingGroupIds = payload.memberships.map(m => m.contactGroupMembership.contactGroupResourceName);
    newLabelIds.forEach(id => {
      if (!existingGroupIds.includes(id)) payload.memberships.push({ contactGroupMembership: { contactGroupResourceName: id } });
    });
  }

  const mergeEmail = (val, type) => {
    const v = cleanData(val);
    if (v) {
      payload.emailAddresses = payload.emailAddresses.filter(e => e.type !== type);
      v.split(';').forEach(emailStr => {
        const clean = emailStr.trim();
        if (clean) payload.emailAddresses.push({ value: clean, type: type });
      });
    }
  };
  mergeEmail(d['SMCI-XX22'], 'home');
  mergeEmail(d['SMCI-XX23'], 'work');

  const mergePhone = (val, type) => {
    if (cleanData(val)) {
      payload.phoneNumbers = payload.phoneNumbers.filter(p => p.type !== type);
      payload.phoneNumbers.push({ value: String(val), type: type });
    }
  };
  mergePhone(d['SMCI-XX25'], 'mobile'); 
  mergePhone(d['SMCI-XX24'], 'home'); 
  mergePhone(d['SMCI-XX26'], 'work');
  mergePhone(d['SMCI-XX27'], 'homeFax'); 
  mergePhone(d['SMCI-XX28'], 'workFax');

  const mergeAddress = (val, type) => {
    const v = cleanData(val);
    if (v && !isDateString(v)) {
      payload.addresses = payload.addresses.filter(a => a.type !== type);
      payload.addresses.push({ formattedValue: String(v), type: type });
    }
  };
  mergeAddress(d['SMCI-XX29'], 'home');
  mergeAddress(d['SMCI-XX30'], 'work');

  const mergeRelation = (personName, type) => {
    const pName = cleanData(personName);
    if (pName) {
      payload.relations = payload.relations.filter(r => r.type !== type);
      payload.relations.push({ person: pName, type: type });
    }
  };
  mergeRelation(d['SMCI-XX62'], 'spouse'); 
  mergeRelation(d['SMCI-XX63'], 'father'); 
  mergeRelation(d['SMCI-XX64'], 'mother'); 
  const customRelType = cleanData(d['SMCI-XX65']); 
  const customRelName = cleanData(d['SMCI-XX66']); 
  if (customRelType && customRelName) {
    payload.relations = payload.relations.filter(r => !(r.type === customRelType && r.person === customRelName));
    payload.relations.push({ person: customRelName, type: customRelType });
  }

  const bday = convertKokiToDate(d['SMCI-XX32']); 
  if (bday) payload.birthdays.push({ date: bday });
  
  const addEvent = (v, typeStr) => {
    const dt = convertKokiToDate(v);
    if (dt) payload.events.push({ date: dt, type: typeStr });
  };
  addEvent(d['SMCI-XX31'], '人物把握日時');
  addEvent(d['SMCI-XX33'], '最終面会年月日');
  addEvent(d['SMCI-XX34'], '死去日時');

  const pushUrl = (v, type) => { 
    if (cleanData(v)) String(v).split(';').forEach(url => payload.urls.push({ value: url.trim(), type: type }));
  };
  pushUrl(d['SMCI-XX36'], 'homePage'); 
  pushUrl(d['SMCI-XX37'], 'profile');
  pushUrl(d['SMCI-XX38'], 'profile');
  pushUrl(d['SMCI-XX39'], 'profile');
  pushUrl(d['SMCI-PNY02'], 'profile');
  pushUrl(d['SMCI-XX40'], 'homePage'); 
  pushUrl(d['SMCI-XX41'], 'homePage'); 
  pushUrl(d['SMCI-XX42'], 'homePage'); 
  pushUrl(d['SMCI-XX43'], 'homePage'); 
  pushUrl(d['SMCI-XX45'], 'homePage'); 
  pushUrl(d['SMCI-XX47'], 'homePage'); 
  pushUrl(d['SMCI-XX49'], 'homePage'); 

  const setCF = (l, v) => { if (cleanData(v)) payload.userDefined.push({ key: l, value: String(v) }); };
  setCF("SMCI11", valSMCI11);
  setCF("SMCI9", valSMCI9);
  setCF("SM人物等級™️", d['SMCI-XX74']);
  const engName = `${cleanData(d['SMCI-XX11'])} ${cleanData(d['SMCI-XX12'])} ${cleanData(d['SMCI-XX13'])}`.trim();
  setCF("英語名", engName);
  ["SMCI-XX50","SMCI-XX51","SMCI-XX52","SMCI-XX53","SMCI-XX54","SMCI-XX55","SMCI-XX56","SMCI-XX57","SMCI-XX58","SMCI-XX59","SMCI-XX60"]
    .forEach(id => setCF(id, d[id]));
  setCF("支払金額(日本円)", d['SMCI-XX71']);
  setCF("支払金額(米ドル)", d['SMCI-XX72']);
  setCF("SM通貨", d['SMCI-XX73']);

  let userNotes = "";
  if (existing && existing.biographies) {
    userNotes = existing.biographies[0].value.split(BASE_DELIMITER)[0];
    userNotes = userNotes.replace(/----\s*$/, "").trim(); 
  }
  let footer = `\n\n----\n${BASE_DELIMITER}\n${SCRIPT_VERSION} sync${getKokiDateString()} ↓\n\n`;
  footer += `SMCI11: ${valSMCI11}\n`;
  footer += `SMCI9: ${valSMCI9}\n`;
  footer += `英語名: ${engName}\n`;
  footer += `\n備考: ${cleanData(d['SMCI-XX75'])}`;
  
  payload.biographies.push({ value: (userNotes + footer).trim() });
  if (existing) payload.etag = existing.etag;
  updates.push("Processed"); 
  return { payload: payload, updates: updates };
}
