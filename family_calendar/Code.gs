/**
 * לוח משפחתי לסבתא — Google Apps Script
 * ---------------------------------------
 * קובץ זה מודבק ב-Code.gs. הממשק נמצא ב-Index.html.
 * הפעלה ראשונה: תפריט "לוח משפחתי" > "הכנה ראשונית" (או הרץ setup פעם אחת).
 */

var SHOW_AGE = false;          // true = לכתוב בלוח גם את הגיל, למשל "יום הולדת משה (70)"
var ADAR_IN_LEAP = 'אדר ב';    // בשנה מעוברת: מי שנולד ב"אדר" רגיל מופיע באדר ב (או שנה ל-'אדר א')
var SHEET_PEOPLE = 'אנשים';
var SHEET_FAMILIES = 'משפחות';
var SHEET_ARCHIVE = 'ארכיון';

var PEOPLE_HEADERS = ['זמן שליחה','משפחה','הורים','שם','סוג אירוע','חודש','יום','שנה עברית','תאריך לועזי','הערה'];
var FAMILY_HEADERS = ['משפחה','הורים','טלפון','עדכון אחרון','מספר אירועים'];

// ---------- תפריט והכנה ----------
function onOpen(){
  SpreadsheetApp.getUi().createMenu('לוח משפחתי')
    .addItem('הכנה ראשונית','setup')
    .addItem('צור גיליון לוח לשנה...','createCalendarSheet')
    .addToUi();
}

function setup(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEET_PEOPLE, PEOPLE_HEADERS);
  ensureSheet_(ss, SHEET_FAMILIES, FAMILY_HEADERS);
  ensureSheet_(ss, SHEET_ARCHIVE, PEOPLE_HEADERS);
  SpreadsheetApp.getUi().alert('מוכן. עכשיו: פרוס > פריסה חדשה > אפליקציית אינטרנט > "כל אחד" ושלח את הקישור למשפחה.');
}

function ensureSheet_(ss, name, headers){
  var sh = ss.getSheetByName(name);
  if (!sh){ sh = ss.insertSheet(name); }
  if (sh.getLastRow()===0){
    sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1); sh.setRightToLeft(true);
  }
  return sh;
}

// ---------- אפליקציית אינטרנט ----------
function doGet(){
  var t = HtmlService.createTemplateFromFile('Index');
  t.hebcalJs = hebcalSource_();
  return t.evaluate().setTitle('הלוח המשפחתי')
    .addMetaTag('viewport','width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// שולח את פונקציות הלוח העברי גם לדפדפן, כדי לא לשכפל קוד
function hebcalSource_(){
  return [HEB_EPOCH_SRC(), hebIsLeap, hebElapsedDays, hebYearDelay, hebNewYear, hebDaysInYear,
    hebMonthLength, hebMonthsInYear, hebToFixed, hebFromFixed, gregToFixed, hebMonthName,
    hebYearMonthOrder, hebDayLetter, hebYearName].map(function(f){ return typeof f==='string'?f:f.toString(); }).join('\n');
}
function HEB_EPOCH_SRC(){ return 'var HEB_EPOCH=' + HEB_EPOCH + ';var HEB_MONTHS=' + JSON.stringify(HEB_MONTHS) + ';'; }

function normKey_(s){ return String(s||'').trim().replace(/^משפחת\s+/,'').replace(/\s+/g,' '); }

/** מחזיר נתונים קיימים של משפחה (אם מילאה כבר) */
function loadFamily(family, parents){
  family = normKey_(family); parents = normKey_(parents);
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PEOPLE);
  if (!sh || sh.getLastRow()<2) return {found:false, people:[]};
  var rows = sh.getRange(2,1,sh.getLastRow()-1,PEOPLE_HEADERS.length).getValues();
  var people = [], phone = '';
  rows.forEach(function(r){
    if (normKey_(r[1])===family && normKey_(r[2])===parents){
      people.push({name:r[3], type:r[4], month:r[5], day:r[6], hebYear:r[7]||'', greg:r[8]||'', note:r[9]||''});
    }
  });
  var fs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_FAMILIES);
  if (fs && fs.getLastRow()>1){
    fs.getRange(2,1,fs.getLastRow()-1,FAMILY_HEADERS.length).getValues().forEach(function(r){
      if (normKey_(r[0])===family && normKey_(r[1])===parents) phone = r[2];
    });
  }
  return {found: people.length>0, people: people, phone: phone};
}

/** שומר משפחה שלמה. שליחה חוזרת מחליפה את הקודמת (הישנה עוברת לארכיון). */
function saveFamily(payload){
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try{
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ensureSheet_(ss, SHEET_PEOPLE, PEOPLE_HEADERS);
    var ar = ensureSheet_(ss, SHEET_ARCHIVE, PEOPLE_HEADERS);
    var fs = ensureSheet_(ss, SHEET_FAMILIES, FAMILY_HEADERS);
    var family = normKey_(payload.family), parents = normKey_(payload.parents);
    if (!family) throw new Error('חסר שם משפחה');
    var now = new Date();

    // העברת שורות ישנות של המשפחה לארכיון
    if (sh.getLastRow()>1){
      var rows = sh.getRange(2,1,sh.getLastRow()-1,PEOPLE_HEADERS.length).getValues();
      var keep=[], old=[];
      rows.forEach(function(r){ (normKey_(r[1])===family && normKey_(r[2])===parents ? old : keep).push(r); });
      if (old.length){
        ar.getRange(ar.getLastRow()+1,1,old.length,PEOPLE_HEADERS.length).setValues(old);
        sh.getRange(2,1,rows.length,PEOPLE_HEADERS.length).clearContent();
        if (keep.length) sh.getRange(2,1,keep.length,PEOPLE_HEADERS.length).setValues(keep);
      }
    }
    // כתיבת השורות החדשות
    var out = (payload.people||[]).filter(function(p){ return p.name && p.month && p.day; }).map(function(p){
      return [now, family, parents, String(p.name).trim(), p.type||'יום הולדת', p.month, Number(p.day),
              p.hebYear?Number(p.hebYear):'', p.greg||'', p.note||''];
    });
    if (!out.length) throw new Error('לא הוזן אף אדם');
    sh.getRange(sh.getLastRow()+1,1,out.length,PEOPLE_HEADERS.length).setValues(out);

    // עדכון רשימת המשפחות
    var frow = 0;
    if (fs.getLastRow()>1){
      fs.getRange(2,1,fs.getLastRow()-1,2).getValues().forEach(function(r,i){
        if (normKey_(r[0])===family && normKey_(r[1])===parents) frow = i+2;
      });
    }
    var frec = [family, parents, payload.phone||'', now, out.length];
    if (frow) fs.getRange(frow,1,1,frec.length).setValues([frec]);
    else fs.appendRow(frec);
    return {ok:true, count:out.length};
  } finally { lock.releaseLock(); }
}

// ---------- יצירת גיליון בפורמט "עתים לבינה" ----------
function createCalendarSheet(){
  var ui = SpreadsheetApp.getUi();
  var today = hebFromFixed(gregToFixed(new Date().getFullYear(), new Date().getMonth()+1, new Date().getDate()));
  var def = today.year + 1;
  var res = ui.prompt('לאיזו שנה עברית ליצור את הלוח? (מספר, למשל ' + def + ' = ' + hebYearName(def) + ')', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton()!==ui.Button.OK) return;
  var year = parseInt(res.getResponseText(),10);
  if (!(year>5000 && year<6000)) { ui.alert('שנה לא תקינה'); return; }
  buildCalendarSheet(year);
  ui.alert('נוצר הגיליון "לוח ' + hebYearName(year) + '". להורדה כאקסל: קובץ > הורדה > Microsoft Excel.');
}

function buildCalendarSheet(year){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = 'לוח ' + hebYearName(year);
  var sh = ss.getSheetByName(name); if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(name); sh.setRightToLeft(true);

  // איסוף אירועים לפי (חודש, יום) של שנת היעד
  var events = {};
  var ps = ss.getSheetByName(SHEET_PEOPLE);
  if (ps && ps.getLastRow()>1){
    ps.getRange(2,1,ps.getLastRow()-1,PEOPLE_HEADERS.length).getValues().forEach(function(r){
      if (!r[3]) return;
      var key = mapToYear_(String(r[5]), Number(r[6]), year);
      var text = eventText_(r[4], r[3], r[7], year);
      (events[key] = events[key] || []).push(text);
    });
  }

  var rows = [];
  rows.push(['', 'לוח עתים לבינה מותאם אישית לשנת ' + hebYearName(year), '', '', '', '', '']);
  rows.push(['', '', '', 'אנא מלאו נתונים אלו בתשומת לב מרובה:', '', 'מצורף כריכה', "כן/מס' כריכה/קולאז/סטנדרט"]);
  rows.push(['', '', '', '', '', '', '']);
  rows.push(['', 'מספר הזמנה:', '', 'לוח עבור משפחת', '', 'טלפון לבירורים:', '']);
  rows.push(['', '', '', '', '', '', '']);
  rows.push(['', 'אופק הלוח', '', 'הערות:', '', 'סה"כ תאריכים', '']);
  rows.push(['', 'הוראות: יש למלא את האירוע הרצוי בעמודת האירוע\nאם קיימים מספר אירועים באותו תאריך - יש להפריד בין האירועים על ידי כוכבית', '', '', '', '', '']);
  rows.push(['', '', '', '', '', '', '']);
  rows.push(['', 'חודש', 'תאריך', 'יום', 'אירוע', 'מס תמונה (אופציונלי)', '']);
  var total = 0;
  hebYearMonthOrder(year).forEach(function(m){
    var mname = hebMonthName(m, year), len = hebMonthLength(m, year);
    for (var d=1; d<=len; d++){
      var list = events[mname+'|'+d] || [];
      total += list.length;
      rows.push(['', mname, holidayName_(m, d, year), hebDayLetter(d), list.join(' * '), '', '']);
    }
  });
  sh.getRange(1,1,rows.length,7).setValues(rows);
  sh.getRange('G6').setValue(total);
  sh.getRange('B9:F9').setFontWeight('bold');
  sh.setColumnWidth(5, 380);
  sh.setFrozenRows(9);
}

function eventText_(type, name, hebYear, year){
  var t = String(type||'יום הולדת');
  var s = (t==='אחר') ? String(name) : t + ' ' + name;
  if (SHOW_AGE && t==='יום הולדת' && hebYear) s += ' (' + (year - Number(hebYear)) + ')';
  return s;
}

/** ממפה חודש+יום שנשמרו לשורה הנכונה בשנת היעד (אדר/אדר א/אדר ב, וחודשים בני 29 יום) */
function mapToYear_(monthName, day, year){
  var leap = hebIsLeap(year), m = monthName;
  if (leap){ if (m==='אדר') m = ADAR_IN_LEAP; }
  else { if (m==='אדר א' || m==='אדר ב') m = 'אדר'; }
  var num = hebMonthNumber_(m, year);
  var len = hebMonthLength(num, year);
  if (day>len) day = len;
  return m + '|' + day;
}
function hebMonthNumber_(name, year){
  for (var i=1;i<=hebMonthsInYear(year);i++) if (hebMonthName(i,year)===name) return i;
  return 7;
}

function holidayName_(m, d, year){
  var chan = [];
  if (m===7){ if(d<=2)return 'ר"ה'; if(d===10)return 'יו"כ'; if(d>=15&&d<=21)return 'סוכות'; if(d===22)return 'שמחת תורה'; }
  if (m===8 && d===1) return 'ר"ח';
  if (m===9){ if(d===1)return 'ר"ח'; if(d>=25)return 'חנוכה'; }
  if (m===10){ var k=hebMonthLength(9,year); if(d<=8-(k-24))return 'חנוכה'; if(d===10)return 'עשרה בטבת'; }
  if (m===11 && d===15) return 'טו בשבט';
  if ((m===12 && !hebIsLeap(year) || m===13) && d===14) return 'פורים';
  if (m===1){ if(d===1)return 'ראש חדש'; if(d>=15&&d<=21)return 'פסח'; if(d===22)return 'אסרו חג'; }
  if (m===2 && d===18) return 'ל"ג בעומר';
  if (m===3){ if(d===6)return 'שבועות'; if(d===7)return 'איסרו חג'; }
  if (m===4 && d===17) return 'יז בתמוז';
  if (m===5 && d===9) return 'תשעה באב';
  return '';
}

// ===== לוח עברי: אלגוריתם Reingold–Dershowitz (משותף לשרת וללקוח) =====
var HEB_EPOCH = -1373427;
var HEB_MONTHS = ['תשרי','חשון','כסליו','טבת','שבט','אדר','אדר א','אדר ב','ניסן','אייר','סיון','תמוז','אב','אלול'];
function hebIsLeap(y){ return ((7*y+1)%19) < 7; }
function hebElapsedDays(y){
  var months = Math.floor((235*y-234)/19);
  var parts = 12084 + 13753*months;
  var day = months*29 + Math.floor(parts/25920);
  if (((3*(day+1))%7) < 3) day++;
  return day;
}
function hebYearDelay(y){
  var ny0=hebElapsedDays(y-1), ny1=hebElapsedDays(y), ny2=hebElapsedDays(y+1);
  if (ny2-ny1===356) return 2;
  if (ny1-ny0===382) return 1;
  return 0;
}
function hebNewYear(y){ return HEB_EPOCH + hebElapsedDays(y) + hebYearDelay(y); }
function hebDaysInYear(y){ return hebNewYear(y+1) - hebNewYear(y); }
// m: 1=ניסן ... 6=אלול 7=תשרי 8=חשון 9=כסליו 10=טבת 11=שבט 12=אדר/אדר א 13=אדר ב
function hebMonthLength(m,y){
  if (m===2||m===4||m===6||m===10||m===13) return 29;
  if (m===12 && !hebIsLeap(y)) return 29;
  if (m===8 && hebDaysInYear(y)%10!==5) return 29;
  if (m===9 && hebDaysInYear(y)%10===3) return 29;
  return 30;
}
function hebMonthsInYear(y){ return hebIsLeap(y)?13:12; }
function hebToFixed(y,m,d){
  var f = hebNewYear(y) + d - 1, i;
  if (m<7){ for(i=7;i<=hebMonthsInYear(y);i++) f+=hebMonthLength(i,y); for(i=1;i<m;i++) f+=hebMonthLength(i,y); }
  else { for(i=7;i<m;i++) f+=hebMonthLength(i,y); }
  return f;
}
function hebFromFixed(f){
  var y = Math.floor((f-HEB_EPOCH)*98496/35975351) - 1;
  while (hebNewYear(y+1) <= f) y++;
  var m = (f < hebToFixed(y,1,1)) ? 7 : 1;
  while (f > hebToFixed(y,m,hebMonthLength(m,y))) m++;
  var d = f - hebToFixed(y,m,1) + 1;
  return {year:y, month:m, day:d};
}
function gregToFixed(y,m,d){
  var leap = (y%4===0 && y%100!==0) || y%400===0;
  return 365*(y-1) + Math.floor((y-1)/4) - Math.floor((y-1)/100) + Math.floor((y-1)/400)
       + Math.floor((367*m-362)/12) + (m<=2?0:(leap?-1:-2)) + d;
}
// שם חודש (לפי הכתיב של הקובץ) ממספר חודש ושנה
function hebMonthName(m,y){
  if (m===12) return hebIsLeap(y)?'אדר א':'אדר';
  if (m===13) return 'אדר ב';
  return [null,'ניסן','אייר','סיון','תמוז','אב','אלול','תשרי','חשון','כסליו','טבת','שבט'][m];
}
// סדר החודשים בלוח (מתשרי) לשנה נתונה, כמספרי חודש
function hebYearMonthOrder(y){
  var o=[7,8,9,10,11,12]; if (hebIsLeap(y)) o.push(13); return o.concat([1,2,3,4,5,6]);
}
function hebDayLetter(d){
  var t=['','א','ב','ג','ד','ה','ו','ז','ח','ט'];
  if (d<10) return t[d];
  if (d===10) return 'י'; if (d===15) return 'טו'; if (d===16) return 'טז';
  if (d<20) return 'י'+t[d-10]; if (d===20) return 'כ'; if (d<30) return 'כ'+t[d-20];
  return 'ל';
}
function hebYearName(y){
  var n=y%1000, s='';
  var hund=['','ק','ר','ש','ת','תק','תר','תש','תת','תתק'];
  s+=hund[Math.floor(n/100)]; n%=100;
  var tens=['','י','כ','ל','מ','נ','ס','ע','פ','צ'];
  if (n===15) s+='טו'; else if (n===16) s+='טז'; else { s+=tens[Math.floor(n/10)]; s+=['','א','ב','ג','ד','ה','ו','ז','ח','ט'][n%10]; }
  return s.length>1 ? s.slice(0,-1)+'"'+s.slice(-1) : s+"'";
}
