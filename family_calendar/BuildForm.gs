/**
 * יוצר טופס Google "הלוח המשפחתי של סבתא" עם עד MAX_PEOPLE בני משפחה בטופס אחד.
 * הרצה: script.google.com > פרויקט חדש > הדבק > הרץ buildForm > קרא את הקישורים ביומן הביצוע.
 */
var MAX_PEOPLE = 12;
var MONTHS = ['תשרי','חשון','כסליו','טבת','שבט','אדר','אדר א','אדר ב','ניסן','אייר','סיון','תמוז','אב','אלול'];
var TYPES = ['יום הולדת','יום נישואין','יארצייט'];

function buildForm(){
  var form = FormApp.create('הלוח המשפחתי של סבתא');
  form.setDescription('ממלאים פעם אחת את כל בני המשפחה. לכל אחד: שם, סוג האירוע, חודש ויום עבריים. לוקח כ-2 דקות.')
      .setConfirmationMessage('נשמר, תודה! רוצים לתקן? פשוט מלאו שוב, הגרסה האחרונה היא הקובעת.')
      .setAllowResponseEdits(true)
      .setProgressBar(true);

  form.addTextItem().setTitle('שם המשפחה').setHelpText('למשל: כהן').setRequired(true);
  form.addTextItem().setTitle('שמות ההורים').setHelpText('למשל: דוד ורחל (כדי להבדיל בין משפחות עם אותו שם)').setRequired(true);
  form.addTextItem().setTitle('טלפון (לא חובה, למקרה של שאלה)');

  var days = [];
  for (var d = 1; d <= 30; d++) days.push(hebDay(d) + '  (' + d + ')');

  var prevNav = null;
  for (var i = 1; i <= MAX_PEOPLE; i++){
    var pb = form.addPageBreakItem().setTitle('בן משפחה ' + i);
    if (prevNav){
      prevNav.setChoices([
        prevNav.createChoice('כן, יש עוד', pb),
        prevNav.createChoice('לא, זה הכל', FormApp.PageNavigationType.SUBMIT)
      ]);
    }
    var req = (i === 1);
    form.addMultipleChoiceItem().setTitle('סוג האירוע').setChoiceValues(TYPES).setRequired(req);
    form.addTextItem().setTitle('שם').setHelpText('ליום הולדת: שם פרטי. ליום נישואין: שני השמות. ליארצייט: שם הנפטר/ת').setRequired(req);
    form.addListItem().setTitle('חודש עברי').setChoiceValues(MONTHS).setRequired(req);
    form.addListItem().setTitle('יום בחודש').setChoiceValues(days).setRequired(req);
    form.addTextItem().setTitle('שנה (לא חובה)').setHelpText('שנת לידה או שנת האירוע, עברית או לועזית. למשל: תשמ"ה או 1985');
    if (i < MAX_PEOPLE){
      prevNav = form.addMultipleChoiceItem().setTitle('יש עוד בן משפחה להוסיף?').setRequired(true);
    } else {
      prevNav = null;
    }
  }

  var ss = SpreadsheetApp.create('הלוח המשפחתי - תשובות');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  Logger.log('קישור לשליחה למשפחה:\n' + form.getPublishedUrl());
  Logger.log('קישור לעריכת הטופס:\n' + form.getEditUrl());
  Logger.log('גיליון התשובות:\n' + ss.getUrl());
}

function hebDay(d){
  var t = ['','א','ב','ג','ד','ה','ו','ז','ח','ט'];
  if (d < 10) return t[d];
  if (d === 10) return 'י'; if (d === 15) return 'טו'; if (d === 16) return 'טז';
  if (d < 20) return 'י' + t[d-10]; if (d === 20) return 'כ'; if (d < 30) return 'כ' + t[d-20];
  return 'ל';
}
