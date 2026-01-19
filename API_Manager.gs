/**
 * File: API_Manager.gs
 * Interactions with Google People API
 */

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

function updateContactPhoto(resourceName, url) {
  try {
    const response = UrlFetchApp.fetch(url);
    const photoData = Utilities.base64Encode(response.getBlob().getBytes());
    People.People.updateContactPhoto({ photoBytes: photoData }, resourceName);
  } catch (e) { console.warn('Photo fail: ' + e.message); }
}
