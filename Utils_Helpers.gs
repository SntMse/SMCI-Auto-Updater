/**
 * File: Utils_Helpers.gs
 * Helper functions
 */

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

// 【新規】カタカナ→ひらがな変換
function kataToHira(str) {
  if (!str) return "";
  return str.replace(/[\u30a1-\u30f6]/g, function(match) {
    var chr = match.charCodeAt(0) - 0x60;
    return String.fromCharCode(chr);
  });
}

// 【新規】タイムスタンプ生成 (SM-ﾄ + 皇紀 + JST時刻)
// 例: SM-ﾄ26860120075505
function getTimestampString() {
  const now = new Date();
  const kokiYear = now.getFullYear() + 660;
  
  // GASのタイムゾーン設定(JST)に合わせて時刻文字列を生成
  const mm = Utilities.formatDate(now, "JST", "MM");
  const dd = Utilities.formatDate(now, "JST", "dd");
  const hh = Utilities.formatDate(now, "JST", "HH");
  const mi = Utilities.formatDate(now, "JST", "mm");
  const ss = Utilities.formatDate(now, "JST", "ss");

  return `SM-ﾄ${kokiYear}${mm}${dd}${hh}${mi}${ss}`;
}
