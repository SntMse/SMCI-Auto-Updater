/**
 * File: Logic_Process.gs
 * Row processing and Payload construction
 * Version: 25.1 (Full Support for Semicolon-separated Multi-values & Paired Fields)
 */

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
      ['SMCI-XX22', 'SMCI-XX23', 'SMCI-XX85'].forEach(k => {
        if(d[k]) cleanData(d[k]).split(';').forEach(e => emailsToCheck.push(e.trim()));
      });
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

// --- Helper for splitting and zipping arrays ---
function getSplitValues(str) {
  if (!str) return [];
  return String(str).split(';').map(s => s.trim());
}

// --- Payload Builder ---
function buildPersonPayload(d, existing, valSMCI11, valSMCI9, newLabelIds) {
  
  // 1. Name (Structured)
  let fName = cleanData(d['SMCI-XX05']); 
  let gName = cleanData(d['SMCI-XX07']); 
  let mName = cleanData(d['SMCI-XX06']); 

  // Name Fallback Logic
  if (!fName && !gName) {
    const enLast  = cleanData(d['SMCI-XX13']);
    const enFirst = cleanData(d['SMCI-XX11']);
    const enMid   = cleanData(d['SMCI-XX12']);
    if (enLast || enFirst) {
      fName = enLast; gName = enFirst; mName = enMid;
    } else {
      const dispName = cleanData(d['SMCI-XX03']);
      if (dispName) gName = dispName;
    }
  }

  const nameObj = {
    familyName: fName, 
    givenName: gName, 
    middleName: mName,
    honorificPrefix: cleanData(d['SMCI-XX04']), 
    honorificSuffix: cleanData(d['SMCI-XX14']),
    phoneticFamilyName: kataToHira(cleanData(d['SMCI-XX08'])), 
    phoneticGivenName: kataToHira(cleanData(d['SMCI-XX10']))   
  };

  const payload = {
    names: [nameObj],
    nicknames: [],
    organizations: [],
    emailAddresses: existing ? [...(existing.emailAddresses || [])] : [],
    phoneNumbers: existing ? [...(existing.phoneNumbers || [])] : [],
    addresses: existing ? [...(existing.addresses || [])] : [],
    memberships: existing ? [...(existing.memberships || [])] : [], 
    relations: existing ? [...(existing.relations || [])] : [], 
    birthdays: [], events: [], urls: [], userDefined: [], biographies: []
  };

  // --- 1.1 Nicknames (XX15) ---
  getSplitValues(d['SMCI-XX15']).forEach(v => {
    if(v) payload.nicknames.push({ value: v });
  });

  // --- 2. Organization (Company/Dept/Title Paired) ---
  const companies = getSplitValues(d['SMCI-XX19']); // XX19 会社名
  const depts     = getSplitValues(d['SMCI-XX20']); // XX20 部門
  const titles    = getSplitValues(d['SMCI-XX21']); // XX21 役職
  
  const maxOrg = Math.max(companies.length, depts.length, titles.length);
  for(let i=0; i<maxOrg; i++) {
    const c = companies[i] || "";
    const dp = depts[i] || "";
    const t = titles[i] || "";
    if (c || dp || t) {
      payload.organizations.push({ name: c, department: dp, title: t, type: 'work' });
    }
  }

  // --- 3. Labels ---
  if (newLabelIds && newLabelIds.length > 0) {
    const existingGroupIds = payload.memberships.map(m => m.contactGroupMembership.contactGroupResourceName);
    newLabelIds.forEach(id => {
      if (!existingGroupIds.includes(id)) payload.memberships.push({ contactGroupMembership: { contactGroupResourceName: id } });
    });
  }

  // --- 4. Emails (Split) ---
  const mergeEmail = (val, type) => {
    const vList = getSplitValues(val);
    if (vList.length > 0) {
      payload.emailAddresses = payload.emailAddresses.filter(e => e.type !== type);
      vList.forEach(email => {
        if (email) payload.emailAddresses.push({ value: email, type: type });
      });
    }
  };
  mergeEmail(d['SMCI-XX22'], 'home');
  mergeEmail(d['SMCI-XX23'], 'work');
  mergeEmail(d['SMCI-XX85'], 'school');

  // --- 5. Phones (Split & Custom Label) ---
  const mergePhone = (val, type, label) => {
    const vList = getSplitValues(val);
    if (vList.length > 0) {
      payload.phoneNumbers = payload.phoneNumbers.filter(p => !(p.type === type && p.formattedType === label));
      vList.forEach(ph => {
        if(ph) {
          let obj = { value: ph, type: type };
          if(label) obj.formattedType = label;
          payload.phoneNumbers.push(obj);
        }
      });
    }
  };
  mergePhone(d['SMCI-XX24'], 'home'); 
  mergePhone(d['SMCI-XX25'], 'mobile'); 
  mergePhone(d['SMCI-XX26'], 'work');
  mergePhone(d['SMCI-XX27'], 'homeFax'); 
  mergePhone(d['SMCI-XX28'], 'workFax');
  mergePhone(d['SMCI-XX86'], 'other', 'FAX（学校）');

  // --- 6. Address (Split & Custom Label) ---
  const mergeAddress = (val, type, label) => {
    const vList = getSplitValues(val);
    if (vList.length > 0) {
      payload.addresses = payload.addresses.filter(a => !(a.type === type && a.formattedType === label));
      vList.forEach(addr => {
        if(addr && !isDateString(addr)) {
          let obj = { formattedValue: addr, type: type };
          if(label) obj.formattedType = label;
          payload.addresses.push(obj);
        }
      });
    }
  };
  mergeAddress(d['SMCI-XX29'], 'home');
  mergeAddress(d['SMCI-XX30'], 'work');
  mergeAddress(d['SMCI-XX87'], 'school'); // Standard type 'school' now supported by API or mapped to other
  mergeAddress(d['SMCI-XX80'], 'other', '実家');

  // --- 7. Relations (Paired XX65 & XX66, plus others) ---
  // A. Individual Relations (Parent/Spouse) - also allow split
  const mergeSimpleRel = (val, type) => {
    getSplitValues(val).forEach(p => {
      if(p) payload.relations.push({ person: p, type: type });
    });
  };
  mergeSimpleRel(d['SMCI-XX62'], 'spouse');
  mergeSimpleRel(d['SMCI-XX63'], 'father');
  mergeSimpleRel(d['SMCI-XX64'], 'mother');

  // B. Custom Relations (Paired Type & Name)
  const relTypes = getSplitValues(d['SMCI-XX65']); // 関係
  const relNames = getSplitValues(d['SMCI-XX66']); // 名前
  const maxRel = Math.max(relTypes.length, relNames.length);
  for(let i=0; i<maxRel; i++) {
    const rt = relTypes[i] || "";
    const rn = relNames[i] || "";
    if (rt && rn) {
      // 既存の重複削除は複雑になるため、追加のみ行う（API側でマージされるか、重複時は複数登録される）
      payload.relations.push({ person: rn, type: rt });
    }
  }

  // --- 8. URLs (Social & Related) ---
  const pushUrl = (val, type) => { 
    getSplitValues(val).forEach(u => {
      if(u) payload.urls.push({ value: u, type: type });
    });
  };
  pushUrl(d['SMCI-XX36'], 'homePage'); 
  pushUrl(d['SMCI-XX37'], 'profile'); // Facebook
  pushUrl(d['SMCI-XX38'], 'profile'); // X
  pushUrl(d['SMCI-XX39'], 'profile'); // Instagram
  pushUrl(d['SMCI-PNY02'], 'profile'); // LinkedIn (Old)
  pushUrl(d['SMCI-XX83'], 'profile');  // LinkedIn (New)
  pushUrl(d['SMCI-XX40'], 'homePage'); // TikTok
  pushUrl(d['SMCI-XX41'], 'homePage'); // Skype
  pushUrl(d['SMCI-XX42'], 'homePage'); // Snapchat
  pushUrl(d['SMCI-XX43'], 'homePage'); // TMDb

  // Related Links (Paired XX46 & XX47)
  const linkNames = getSplitValues(d['SMCI-XX46']);
  const linkUrls  = getSplitValues(d['SMCI-XX47']);
  const maxLink = Math.max(linkNames.length, linkUrls.length);
  for(let i=0; i<maxLink; i++) {
    const ln = linkNames[i] || ""; // 名前がない場合は空文字（APIによってはラベルが必要）
    const lu = linkUrls[i] || "";
    if (lu) {
      let urlObj = { value: lu };
      // ラベルがある場合、またはGoogle標準以外のタイプとして扱う場合
      if (ln) urlObj.formattedType = ln; 
      else urlObj.type = 'other';
      payload.urls.push(urlObj);
    }
  }

  // --- 9. Custom Fields (All Multi-value) ---
  const setCF = (key, val) => {
    getSplitValues(val).forEach(v => {
      if(v) payload.userDefined.push({ key: key, value: v });
    });
  };

  setCF("SMCI11", valSMCI11);
  setCF("SMCI9", valSMCI9);
  setCF("SM人物等級™️", d['SMCI-XX74']);
  setCF("SMCI3", d['SMCI-XX78']);
  setCF("食物制限", d['SMCI-XX76']);
  setCF("英語表示名", d['SMCI-XX77']);
  setCF("愛称", d['SMCI-XX16']); // XX16
  setCF("別名", d['SMCI-XX79']); // XX79
  setCF("性別", d['SMCI-XX17']);
  setCF("代名詞", d['SMCI-XX18']);
  setCF("出身地", d['SMCI-XX81']);
  setCF("出生地", d['SMCI-XX82']);
  setCF("学校名", d['SMCI-XX84']); 

  // English Name CF
  const engName = `${cleanData(d['SMCI-XX11'])} ${cleanData(d['SMCI-XX12'])} ${cleanData(d['SMCI-XX13'])}`.trim();
  setCF("英語名", engName);
  
  // Other ID/Codes
  ["SMCI-XX50","SMCI-XX51","SMCI-XX52","SMCI-XX53","SMCI-XX54","SMCI-XX55","SMCI-XX56","SMCI-XX57","SMCI-XX58","SMCI-XX59","SMCI-XX60"]
    .forEach(id => setCF(id, d[id]));
  
  setCF("支払金額(日本円)", d['SMCI-XX71']);
  setCF("支払金額(米ドル)", d['SMCI-XX72']);
  setCF("SM通貨", d['SMCI-XX73']);

  // Dates (Single values usually, but logic kept simple)
  const bday = convertKokiToDate(d['SMCI-XX32']); 
  if (bday) payload.birthdays.push({ date: bday });
  
  const addEvent = (v, typeStr) => {
    const dt = convertKokiToDate(v);
    if (dt) payload.events.push({ date: dt, type: typeStr });
  };
  addEvent(d['SMCI-XX31'], '人物把握日時');
  addEvent(d['SMCI-XX33'], '最終面会年月日');
  addEvent(d['SMCI-XX34'], '死去日時');

  // --- Biography / Footer ---
  let userNotes = "";
  if (existing && existing.biographies) {
    userNotes = existing.biographies[0].value.split(BASE_DELIMITER)[0];
    userNotes = userNotes.replace(/----\s*$/, "").trim(); 
  }
  let footer = `\n\n----\n${BASE_DELIMITER}\n${getTimestampString()}\n\n`;
  footer += `SMCI11: ${valSMCI11}\n`;
  footer += `SMCI9: ${valSMCI9}\n`;
  footer += `SMCI3: ${cleanData(d['SMCI-XX78'])}\n`; 
  footer += `英語名: ${engName}\n`;
  footer += `\n備考: ${cleanData(d['SMCI-XX75'])}`;
  
  payload.biographies.push({ value: (userNotes + footer).trim() });
  
  if (existing) payload.etag = existing.etag;
  
  return { payload: payload, updates: ["Processed"] };
}
