/**
 * File: Utils_Helpers.gs
 * Helper functions
 * Version: 27.1 (Fix: Apostrophe Date Support)
 */

function cleanData(val) {
  if (val === null || val === undefined) return "";
  if (val instanceof Date) return val;
  let str = String(val);
  if (str.includes("CellImage") || str === "Obj") return "";
  str = str.trim();
  const upperStr = str.toUpperCase();
  const ignoreList = ["NIL", "UNK", "取得中", "#N/A", "N/A", "=NA()", "0", ""];
  if (ignoreList.includes(upperStr)) return "";
  return str;
}

/**
 * 日付・皇紀・年なし形式をGoogleコンタクト形式に変換
 * '10/15 などの形式は年なしとして処理します
 */
function convertKokiToDate(val) {
  if (!val) return null;
  
  let dateObj = { year: null, month: null, day: null };
  let strVal = String(val).trim();
  
  // 1. 年なし形式の検出 ('10/15, 10-15 など)
  // 4桁の年が含まれていない、かつ M/D 形式の場合
  const noYearMatch = strVal.match(/^(\d{1,2})[\/\-](\d{1,2})$/);
  if (noYearMatch && !(val instanceof Date)) {
    dateObj.month = parseInt(noYearMatch[1], 10);
    dateObj.day = parseInt(noYearMatch[2], 10);
    delete dateObj.year; // 年をプロパティごと削除（Googleコンタクトで「年不明」になる）
    return dateObj;
  }
  
  // 2. 通常の日付オブジェクトまたは日付文字列の処理
  let date;
  if (val instanceof Date) {
    date = new Date(val);
  } else {
    if (!strVal.match(/[\d]{2,4}[\/\-]/)) return null;
    date = new Date(strVal);
  }
  
  if (isNaN(date.getTime())) return null;
  
  let year = date.getFullYear();
  
  // 皇紀対応 (2600年以上なら西暦に変換)
  if (year > 2600) year = year - 660;
  
  dateObj.year = year;
  dateObj.month = date.getMonth() + 1;
  dateObj.day = date.getDate();
  
  return dateObj;
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

function kataToHira(str) {
  if (!str) return "";
  return str.replace(/[\u30a1-\u30f6]/g, function(match) {
    var chr = match.charCodeAt(0) - 0x60;
    return String.fromCharCode(chr);
  });
}

function getTimestampString() {
  const now = new Date();
  const kokiYear = now.getFullYear() + 660;
  const mm = Utilities.formatDate(now, "JST", "MM");
  const dd = Utilities.formatDate(now, "JST", "dd");
  const hh = Utilities.formatDate(now, "JST", "HH");
  const mi = Utilities.formatDate(now, "JST", "mm");
  const ss = Utilities.formatDate(now, "JST", "ss");
  return `SM-ﾄ${kokiYear}${mm}${dd}${hh}${mi}${ss}`;
}
