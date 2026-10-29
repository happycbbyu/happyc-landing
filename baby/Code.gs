// 분유/수유 기록 API (Google Apps Script)
// 시트 열: A 시각 | B 종류 | C 양 | D 메모
// 배포: 배포 > 새 배포 > 웹 앱 > 실행: 나 / 액세스: 모든 사용자
// 시트가 문자열을 날짜로 해석할 때 쓰는 시간대와 맞춰야 시각이 어긋나지 않음
var TZ = SpreadsheetApp.getActive().getSpreadsheetTimeZone();

function doPost(e) {
  var p = e.parameter;
  // 단축어에서 요청 본문을 JSON으로 보낸 경우
  if (!p.type && !p.action && e.postData && e.postData.contents) { try { p = JSON.parse(e.postData.contents); } catch (err) {} }
  var sheet = SpreadsheetApp.getActiveSheet();
  if (p.action === 'delete') {
    if (sheet.getLastRow() > 0) sheet.deleteRow(sheet.getLastRow());
    return json({ ok: true, deleted: true });
  }
  if (p.type === '일기') return json(writeDiary(p.note));
  if (p.text) p = parseText(p.text);
  if (!p) return json({ ok: false, say: '숫자를 못 알아들었어요. 다시 말해 주세요.' });
  var amount = Number(p.amount) || 0;
  var type = p.type || '분유';
  var when = p.time ? new Date(p.time) : new Date();
  if (isNaN(when)) when = new Date();
  sheet.appendRow([when, type, amount, p.note || '']);
  return json({ ok: true, time: fmt(when), type: type, amount: amount, say: type + ' ' + amount + (type === '모유' ? '분' : 'ml') + ' 기록했어요.' });
}

// 말로 한 문장 해석: "20분 먹였어" → 모유 20, "240ml 먹었어" → 분유 240, "유축 90" → 유축 90
function parseText(text) {
  var t = String(text);
  var num = t.match(/\d+/);
  if (!num) return null;
  var type = '분유';
  if (/유축/.test(t)) type = '유축';
  else if (/모유|직수/.test(t) || (/\d+\s*분(?!유)/.test(t) && !/ml|㎖|밀리|미리|cc/i.test(t))) type = '모유';
  return { type: type, amount: num[0], note: t };
}

function doGet(e) {
  // shortcut: 최근 300줄만 반환, 1년 넘게 쓰면 날짜 필터로 바꿀 것
  var rows = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
  var out = rows.filter(function (r) { return r[0] instanceof Date || !isNaN(new Date(r[0])); }).slice(-300).map(function (r) {
    return { time: fmt(r[0]), type: String(r[1] || ''), amount: Number(r[2]) || 0, note: String(r[3] || '') };
  });
  // ?summary=1 → 시리가 읽어 줄 한 문장
  if (e && e.parameter && e.parameter.summary) return ContentService.createTextOutput(summary(out));
  return json(out);
}

function summary(rows) {
  if (!rows.length) return '아직 기록이 없어요.';
  var last = rows[rows.length - 1];
  var mins = Math.max(0, Math.round((Date.now() - new Date(last.time.replace(' ', 'T'))) / 60000));
  var ago = mins < 60 ? mins + '분 전' : Math.floor(mins / 60) + '시간 ' + (mins % 60) + '분 전';
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var todays = rows.filter(function (r) { return r.time.indexOf(today) === 0; });
  var ml = 0;
  todays.forEach(function (r) { if (r.type !== '모유') ml += r.amount; });
  return '마지막 수유는 ' + ago + ', ' + last.type + ' ' + last.amount + (last.type === '모유' ? '분' : 'ml') + '이에요. 오늘은 ' + todays.length + '번, 총 ' + ml + 'ml 먹었어요.';
}

// 육아일기: 구글 문서 하나에 날짜 제목 + "HH:mm 내용"으로 이어 씀. 문서는 첫 기록 때 자동 생성
function writeDiary(text) {
  text = String(text || '').trim();
  if (!text) return { ok: false, error: 'empty' };
  var props = PropertiesService.getScriptProperties();
  var doc;
  try { doc = DocumentApp.openById(props.getProperty('diaryDocId')); }
  catch (err) { doc = DocumentApp.create('육아일기'); props.setProperty('diaryDocId', doc.getId()); }
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if (props.getProperty('diaryLastDay') !== today) {
    doc.getBody().appendParagraph(today).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    props.setProperty('diaryLastDay', today);
  }
  doc.getBody().appendParagraph(Utilities.formatDate(new Date(), TZ, 'HH:mm') + '  ' + text);
  return { ok: true, url: doc.getUrl() };
}

// 편집기에서 한 번 실행: 권한 승인 + 문서 생성 확인용
function testDiary() { Logger.log(writeDiary('테스트 일기')); }

// 편집기에서 실행: 이 스크립트가 어느 시트를 읽는지, 줄이 몇 개인지 확인용
function testRead() {
  var sh = SpreadsheetApp.getActiveSheet();
  Logger.log('파일: ' + SpreadsheetApp.getActive().getName() + ' / 탭: ' + sh.getName() + ' / 줄 수: ' + sh.getLastRow() + ' / 첫 칸: ' + sh.getRange(1, 1).getValue());
}

// 시간대 오프셋을 붙여 보내면 폰이 어디서 열어도 같은 시각으로 계산됨
function fmt(v) {
  return v instanceof Date ? Utilities.formatDate(v, TZ, "yyyy-MM-dd'T'HH:mm:ssXXX") : String(v);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
