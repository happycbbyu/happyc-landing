// 분유/수유 기록 API (Google Apps Script)
// 시트 열: A 시각 | B 종류 | C 양 | D 메모
// 배포: 배포 > 새 배포 > 웹 앱 > 실행: 나 / 액세스: 모든 사용자
var TZ = 'Asia/Seoul';

function doPost(e) {
  var p = e.parameter;
  var sheet = SpreadsheetApp.getActiveSheet();
  if (p.action === 'delete') {
    if (sheet.getLastRow() > 0) sheet.deleteRow(sheet.getLastRow());
    return json({ ok: true, deleted: true });
  }
  var amount = Number(p.amount) || 0;
  var type = p.type || '분유';
  var when = p.time ? new Date(p.time) : new Date();
  if (isNaN(when)) when = new Date();
  var ts = Utilities.formatDate(when, TZ, 'yyyy-MM-dd HH:mm');
  sheet.appendRow([ts, type, amount, p.note || '']);
  return json({ ok: true, time: ts, type: type, amount: amount });
}

function doGet() {
  // shortcut: 최근 300줄만 반환, 1년 넘게 쓰면 날짜 필터로 바꿀 것
  var rows = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
  var out = rows.filter(function (r) { return r[0]; }).slice(-300).map(function (r) {
    return { time: fmt(r[0]), type: String(r[1] || ''), amount: Number(r[2]) || 0, note: String(r[3] || '') };
  });
  return json(out);
}

function fmt(v) {
  return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm') : String(v);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
