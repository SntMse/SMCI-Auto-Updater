/**
 * File: Logic_Process.gs
 * Row processing and Payload construction
 * v24.1 Fix: Changed 'label' property to 'formattedType' for Custom Labels
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

// --- Payload Builder ---
function buildPersonPayload(d, existing, valSMCI11, valSMCI9, newLabelIds) {
  let updates = []; 

  // 1. Name & Furigana
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

  // 2. Organization
  const company = cleanData(d['SMCI-XX19']);
  if (company) {
    payload.organizations.push({ name: company, title: cleanData(d['SMCI-XX21']), department: cleanData(d['SMCI-XX20']), type: 'work' });
  }

  // 3. Labels
  if (newLabelIds && newLabelIds.length > 0) {
    const existingGroupIds = payload.memberships.map(m => m.contactGroupMembership.contactGroupResourceName);
    newLabelIds.forEach(id => {
      if (!existingGroupIds.includes(id)) payload.memberships.push({ contactGroupMembership: { contactGroupResourceName: id } });
    });
  }

  // 4. Email Merge
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
  mergeEmail(d['SMCI-XX85'], 'school');

  // 5. Phone Merge (Fix: Use 'formattedType' instead of 'label')
  const mergePhone = (val, type, label) => {
    const v = cleanData(val);
    if (v) {
      // 既存の同タイプ・同ラベルを除去
      payload.phoneNumbers = payload.phoneNumbers.filter(p => !(p.type === type && p.formattedType === label));
      
      let phoneObj = { value: String(v), type: type };
      
      // 【重要修正】label ではなく formattedType を使う
      if (label) phoneObj.formattedType = label; 
      
      payload.phoneNumbers.push(phoneObj);
    }
  };
  mergePhone(d['SMCI-XX25'], 'mobile'); 
  mergePhone(d['SMCI-XX24'], 'home'); 
  mergePhone(d['SMCI-XX26'], 'work');
  mergePhone(d['SMCI-XX27'], 'homeFax'); 
  mergePhone(d['SMCI-XX28'], 'workFax');
  mergePhone(d['SMCI-XX86'], 'other', 'FAX（学校）');

  // 6. Address Merge (Fix: Use 'formattedType' instead of 'label')
  const mergeAddress = (val, type, label) => {
    const v = cleanData(val);
    if (v && !isDateString(v)) {
      payload.addresses = payload.addresses.filter(a => !(a.type === type && a.formattedType === label));
      
      let addrObj = { formattedValue: String(v), type: type };
      
      // 【重要修正】label ではなく formattedType を使う
      if (label) addrObj.formattedType = label;
      
      payload.addresses.push(addrObj);
    }
  };
  mergeAddress(d['SMCI-XX29'], 'home');
  mergeAddress(d['SMCI-XX30'], 'work');
  mergeAddress(d['SMCI-XX87'], 'school');
  mergeAddress(d['SMCI-XX80'], 'other', '実家');

  // 7. Relations
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

  // 8. Dates & URLs
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
  pushUrl(d['SMCI-XX83'],  'profile');
  pushUrl(d['SMCI-XX40'], 'homePage'); 
  pushUrl(d['SMCI-XX41'], 'homePage'); 
  pushUrl(d['SMCI-XX42'], 'homePage'); 
  pushUrl(d['SMCI-XX43'], 'homePage'); 
  pushUrl(d['SMCI-XX45'], 'homePage'); 
  pushUrl(d['SMCI-XX47'], 'homePage'); 
  pushUrl(d['SMCI-XX49'], 'homePage'); 

  // 9. Custom Fields
  const setCF = (l, v) => { if (cleanData(v)) payload.userDefined.push({ key: l, value: String(v) }); };
  setCF("SMCI11", valSMCI11);
  setCF("SMCI9", valSMCI9);
  setCF("SM人物等級™️", d['SMCI-XX74']);
  setCF("SMCI3", d['SMCI-XX78']); 
  setCF("食物制限", d['SMCI-XX76']); 
  setCF("英語表示名", d['SMCI-XX77']); 
  setCF("別名", d['SMCI-XX79']); 
  setCF("出身地", d['SMCI-XX81']); 
  setCF("出生地", d['SMCI-XX82']); 
  setCF("学校名", d['SMCI-XX84']); 
  
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
  let footer = `\n\n----\n${BASE_DELIMITER}\n${getTimestampString()}\n\n`;
  footer += `SMCI11: ${valSMCI11}\n`;
  footer += `SMCI9: ${valSMCI9}\n`;
  footer += `SMCI3: ${cleanData(d['SMCI-XX78'])}\n`; 
  footer += `英語名: ${engName}\n`;
  footer += `\n備考: ${cleanData(d['SMCI-XX75'])}`;
  
  payload.biographies.push({ value: (userNotes + footer).trim() });
  if (existing) payload.etag = existing.etag;
  updates.push("Processed"); 
  return { payload: payload, updates: updates };
}
