/**
 * GitHub to GAS Sync Script (Direct API Mode)
 * 秘密情報は ScriptProperties から読み込みます
 */

const SYNC_FILES = [
  "Config_Main",
  "Logic_Process",
  "API_Manager",
  "Utils_Helpers",
  "_GitHub_Sync"
];

function pullFromGitHub_New() {
  console.log("🚀 GitHub Sync Started (Direct API Mode)...");

  const props = PropertiesService.getScriptProperties();
  const USER = props.getProperty('GITHUB_USER');
  const REPO = props.getProperty('GITHUB_REPO');
  const TOKEN = props.getProperty('GITHUB_TOKEN');
  const BRANCH = "main";
  const SCRIPT_ID = ScriptApp.getScriptId();

  if (!USER || !REPO || !TOKEN) {
    console.error("❌ エラー: スクリプトプロパティ(GITHUB_USER, GITHUB_REPO, GITHUB_TOKEN)を確認してください。");
    return;
  }

  // 1. GitHubから最新コードを取得
  const baseUrl = `https://api.github.com/repos/${USER}/${REPO}/contents/`;
  const ghHeaders = {
    "Authorization": `token ${TOKEN}`,
    "Accept": "application/vnd.github.v3+json"
  };

  let newFiles = [];
  let successCount = 0;

  SYNC_FILES.forEach(fileName => {
    const fileUrl = `${baseUrl}${fileName}.gs?ref=${BRANCH}`;
    try {
      const response = UrlFetchApp.fetch(fileUrl, { headers: ghHeaders, muteHttpExceptions: true });
      if (response.getResponseCode() === 200) {
        const json = JSON.parse(response.getContentText());
        const decodedContent = Utilities.newBlob(Utilities.base64Decode(json.content)).getDataAsString();
        
        newFiles.push({
          name: fileName,
          type: "SERVER_JS",
          source: decodedContent
        });
        console.log(`✅ Loaded from GitHub: ${fileName}.gs`);
        successCount++;
      } else {
        console.warn(`⚠️ GitHub Fetch Failed: ${fileName}.gs (Code: ${response.getResponseCode()})`);
      }
    } catch (e) {
      console.error(`❌ Error fetching ${fileName}: ${e.message}`);
    }
  });

  if (successCount !== SYNC_FILES.length) {
    console.error("❌ 全ファイルの取得に失敗したため、更新を中止します。");
    return;
  }

  // 2. GAS APIを直接叩いて自分自身を更新
  const gasApiUrl = `https://script.googleapis.com/v1/projects/${SCRIPT_ID}/content`;
  const gasToken = ScriptApp.getOAuthToken();
  
  try {
    const currentRes = UrlFetchApp.fetch(gasApiUrl, {
      headers: { "Authorization": `Bearer ${gasToken}` }
    });
    const currentData = JSON.parse(currentRes.getContentText());
    const manifest = currentData.files.find(f => f.name === "appsscript");
    
    if (manifest) {
      newFiles.push(manifest); 
    } else {
      console.error("❌ appsscript.json が取得できませんでした。");
      return;
    }

    UrlFetchApp.fetch(gasApiUrl, {
      method: "put",
      contentType: "application/json",
      headers: { "Authorization": `Bearer ${gasToken}` },
      payload: JSON.stringify({ files: newFiles })
    });

    console.log(`🎉 Sync Complete! Updated ${successCount} files.`);
    console.log("⚠️ 設定反映のため、ブラウザをリロードしてください。");

  } catch (e) {
    console.error("❌ Project Update Failed: " + e.message);
  }
}
