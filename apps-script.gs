// 한자 몬스터 헌트 - 결과 수집용 Google Apps Script
// 구글 시트에서 [확장 프로그램 > Apps Script]에 붙여넣고 웹 앱으로 배포하세요.

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var d = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('결과') || ss.insertSheet('결과');
    if (sh.getLastRow() === 0) {
      sh.appendRow(['시각', '이름/번호', '잡은 한자 수', '전체', '잡은 한자', '이번 판']);
    }
    var name = clean(d.name, 30);
    var row = [new Date(), name, Number(d.caught) || 0, Number(d.total) || 0, clean(d.chars, 200), Number(d.sess) || 0];
    var data = sh.getDataRange().getValues();
    var found = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]) === name) { found = i + 1; break; }
    }
    if (found) sh.getRange(found, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('한자 몬스터 헌트 결과 수집 중');
}

// 수식으로 해석될 수 있는 글자로 시작하면 앞에 작은따옴표를 붙여요.
function clean(v, max) {
  var t = String(v || '').slice(0, max);
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}
