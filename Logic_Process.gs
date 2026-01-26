/**
 * File: Logic_Process.gs
 * Row processing and Payload construction
 * Version: 26.1 (Fix: Force Overwrite/Delete for Empty Fields)
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

function getSplitValues(str) {
  const cleaned = cleanData(str);
  if (!cleaned) return [];
  return cleaned.split(';').map(s => s.trim()).filter(s => s !== "");
}

function buildPersonPayload(d, existing, valSMCI11, valSMCI9, newLabelIds) {
  
  // 1. Name
  let fName = cleanData(d['SMCI-XX05']); 
  let gName = cleanData(d['SMCI-XX07']); 
  let mName = cleanData(d['SMCI-XX06']); 

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

  // 1.1 Nicknames
  getSplitValues(d['SMCI-XX15']).forEach(v => {
    payload.nicknames.push({ value: v });
  });

  // 2. Organization
  const companies = getSplitValues(d['SMCI-XX19']); 
  const depts     = getSplitValues(d['SMCI-XX20']); 
  const titles    = getSplitValues(d['SMCI-XX21']); 
  
  const maxOrg = Math.max(companies.length, depts.length, titles.length);
  for(let i=0; i<maxOrg; i++) {
    const c = companies[i] || "";
    const dp = depts[i] || "";
    const t = titles[i] || "";
    if (c || dp || t) {
      payload.organizations.push({ name: c, department: dp, title: t, type: 'work' });
    }
  }

  // 3. Labels
  if (newLabelIds && newLabelIds.length > 0) {
    const existingGroupIds = payload.memberships.map(m => m.contactGroupMembership.contactGroupResourceName);
    newLabelIds.forEach(id => {
      if (!existingGroupIds.includes(id)) payload.memberships.push({ contactGroupMembership: { contactGroupResourceName: id } });
    });
  }

  // 4. Emails (Force Overwrite)
  const mergeEmail = (val, type) => {
    // 【重要】まず既存の同じタイプを全て削除する（空欄であっても削除して整合性を保つ）
    payload.emailAddresses = payload.emailAddresses.filter(e => e.type !== type);
    
    // 新しい値があれば追加する
    const vList = getSplitValues(val);
    vList.forEach(email => {
      payload.emailAddresses.push({ value: email, type: type });
    });
  };
  mergeEmail(d['SMCI-XX22'], 'home');
  mergeEmail(d['SMCI-XX23'], 'work');
  mergeEmail(d['SMCI-XX85'], 'school');

  // 5. Phones (Force Overwrite)
  const mergePhone = (val, type, label) => {
    // 【重要】まず既存を削除
    payload.phoneNumbers = payload.phoneNumbers.filter(p => !(p.type === type && p.formattedType === label));
    
    const vList = getSplitValues(val);
    vList.forEach(ph => {
      let obj = { value: ph, type: type };
      if(label) obj.formattedType = label; 
      payload.phoneNumbers.push(obj);
    });
  };
  mergePhone(d['SMCI-XX24'], 'home'); 
  mergePhone(d['SMCI-XX25'], 'mobile'); 
  mergePhone(d['SMCI-XX26'], 'work');
  mergePhone(d['SMCI-XX27'], 'homeFax'); 
  mergePhone(d['SMCI-XX28'], 'workFax');
  mergePhone(d['SMCI-XX86'], 'other', 'FAX（学校）');

  // 6. Address (Force Overwrite)
  const mergeAddress = (val, type, label) => {
    // 【重要】まず既存を削除
    payload.addresses = payload.addresses.filter(a => !(a.type === type && a.formattedType === label));
    
    const vList = getSplitValues(val);
    vList.forEach(addr => {
      if(!isDateString(addr)) {
        let obj = { formattedValue: addr, type: type };
        if(label) obj.formattedType = label;
        payload.addresses.push(obj);
      }
    });
  };
  mergeAddress(d['SMCI-XX29'], 'home');
  mergeAddress(d['SMCI-XX30'], 'work');
  mergeAddress(d['SMCI-XX87'], 'school'); 
  mergeAddress(d['SMCI-XX80'], 'other', '実家');

  // 7. Relations (Force Overwrite)
  // Simple Relations
  const mergeSimpleRel = (val, type) => {
    // 【重要】まず既存の同じタイプを削除
    payload.relations = payload.relations.filter(r => r.type !== type);
    
    getSplitValues(val).forEach(p => {
      payload.relations.push({ person: p, type: type });
    });
  };
  mergeSimpleRel(d['SMCI-XX62'], 'spouse');
  mergeSimpleRel(d['SMCI-XX63'], 'father');
  mergeSimpleRel(d['SMCI-XX64'], 'mother');

  // Custom Relations
  const relTypes = getSplitValues(d['SMCI-XX65']); 
  const relNames = getSplitValues(d['SMCI-XX66']); 
  const maxRel = Math.max(relTypes.length, relNames.length);
  for(let i=0; i<maxRel; i++) {
    const rt = relTypes[i] || "";
    const rn = relNames[i] || "";
    if (rt && rn) {
      // カスタム関係は種類が多岐にわたるため、ここでは追加のみ行う（重複チェックは複雑になるため省略）
      // 必要であれば完全一致削除を追加可能
      payload.relations.push({ person: rn, type: rt });
    }
  }

  // 8. URLs
  const pushUrl = (val, type) => { 
    getSplitValues(val).forEach(u => {
      payload.urls.push({ value: u, type: type });
    });
  };
  pushUrl(d['SMCI-XX36'], 'homePage'); 
  pushUrl(d['SMCI-XX37'], 'profile'); 
  pushUrl(d['SMCI-XX38'], 'profile'); 
  pushUrl(d['SMCI-XX39'], 'profile'); 
  pushUrl(d['SMCI-PNY02'], 'profile'); 
  pushUrl(d['SMCI-XX83'], 'profile');  
  pushUrl(d['SMCI-XX40'], 'homePage');
  pushUrl(d['SMCI-XX41'], 'homePage');
  pushUrl(d['SMCI-XX42'], 'homePage');
  pushUrl(d['SMCI-XX43'], 'homePage');

  const linkNames = getSplitValues(d['SMCI-XX46']);
  const linkUrls  = getSplitValues(d['SMCI-XX47']);
  const maxLink = Math.max(linkNames.length, linkUrls.length);
  for(let i=0; i<maxLink; i++) {
    const ln = linkNames[i] || ""; 
    const lu = linkUrls[i] || "";
    if (lu) {
      let urlObj = { value: lu };
      if (ln) urlObj.formattedType = ln; 
      else urlObj.type = 'other';
      payload.urls.push(urlObj);
    }
  }

  // 9. Custom Fields
  const setCF = (key, val) => {
    getSplitValues(val).forEach(v => {
      payload.userDefined.push({ key: key, value: v });
    });
  };

  setCF("SMCI11", valSMCI11);
  setCF("SMCI9", valSMCI9);
  setCF("SM人物等級™️", d['SMCI-XX74']);
  setCF("SMCI3", d['SMCI-XX78']);
  setCF("食物制限", d['SMCI-XX76']);
  setCF("英語表示名", d['SMCI-XX77']);
  setCF("愛称", d['SMCI-XX16']);
  setCF("別名", d['SMCI-XX79']);
  setCF("性別", d['SMCI-XX17']);
  setCF("代名詞", d['SMCI-XX18']);
  setCF("出身地", d['SMCI-XX81']);
  setCF("出生地", d['SMCI-XX82']);
  setCF("学校名", d['SMCI-XX84']); 
  setCF("学部学科等", d['SMCI-XX88']); 

  const engName = `${cleanData(d['SMCI-XX11'])} ${cleanData(d['SMCI-XX12'])} ${cleanData(d['SMCI-XX13'])}`.trim();
  setCF("英語名", engName);
  
  ["SMCI-XX50","SMCI-XX51","SMCI-XX52","SMCI-XX53","SMCI-XX54","SMCI-XX55","SMCI-XX56","SMCI-XX57","SMCI-XX58","SMCI-XX59","SMCI-XX60"]
    .forEach(id => setCF(id, d[id]));
  
  setCF("支払金額(日本円)", d['SMCI-XX71']);
  setCF("支払金額(米ドル)", d['SMCI-XX72']);
  setCF("SM通貨", d['SMCI-XX73']);

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
  
  // v26.1 指定フォーマット
  let footer = `\n\n----\n${BASE_DELIMITER}\n${SCRIPT_VERSION} (${getTimestampString()})\n\n`;
  footer += `SMCI11: ${valSMCI11}\n`;
  footer += `SMCI9: ${valSMCI9}\n`;
  footer += `SMCI3: ${cleanData(d['SMCI-XX78'])}\n`; 
  footer += `英語名: ${engName}\n`;
  footer += `\n備考: ${cleanData(d['SMCI-XX75'])}`;
  
  payload.biographies.push({ value: (userNotes + footer).trim() });
  
  if (existing) payload.etag = existing.etag;
  
  return { payload: payload, updates: ["Processed"] };
}
