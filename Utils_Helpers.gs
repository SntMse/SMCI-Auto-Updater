/**
 * File: Utils_Helpers.gs
 * Helper functions
 * v26.0 Fix: Enhanced cleanData to strictly remove UNK, NIL, #N/A, etc.
 */

function cleanData(val) {
  // null, undefined check
  if (val === null || val === undefined) return "";
  
  // Date object check
  if (val instanceof Date) return val;
  
  // Convert to string
  let str = String(val);
  
  // remove CellImage objects
  if (str.includes("CellImage") || str === "Obj") return "";
  
  // Trim spaces (full-width & half-width)
  str = str.trim();
  
  // Case-insensitive check for ignore list
  const upperStr = str.toUpperCase();
  const ignoreList = [
    "NIL", 
    "UNK", 
    "取得中", 
    "#N/A", 
    "N/A", 
    "=NA()", 
    "0", // 0も場合によっては無視したいケースがあれば追加、今回は基本リストのみ
    ""
  ];
  
  if (ignoreList.includes(upperStr)) return "";
  
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
