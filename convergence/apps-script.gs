/**
 * 융합전공 이수 현황판 — 데이터 저장 서버
 *
 * 하는 일: 학생과 관리자가 앱에서 입력한 내용을 구글 스프레드시트에 저장하고 돌려준다.
 * 학생은 구글 로그인이 필요 없다. 이 코드가 관리자 계정 권한으로 대신 기록하기 때문.
 *
 * ── 설치 방법 ──────────────────────────────────────
 *  1. script.google.com 에 접속해 "새 프로젝트"를 만든다.
 *  2. 기본으로 들어있는 코드를 모두 지우고 이 파일 내용을 붙여넣는다.
 *  3. 오른쪽 위 "배포" → "새 배포" 를 누른다.
 *  4. 톱니바퀴 → "웹 앱" 을 고른다.
 *  5. 다음 값으로 맞춘다.
 *        실행 계정        : 나
 *        액세스 권한이 있는 사용자 : 모든 사용자     ← 반드시 이것으로!
 *  6. "배포"를 누르고 권한을 허용한다.
 *     (처음엔 "확인되지 않은 앱" 경고가 뜬다. 고급 → 이동 을 눌러 진행)
 *  7. 나오는 "웹 앱 URL"(.../exec 로 끝남)을 복사한다.
 *  8. index.html 맨 위쪽의  var API_URL = '';  안에 그 주소를 붙여넣는다.
 *
 *  스프레드시트는 처음 실행될 때 자동으로 만들어진다.
 *  만들어진 시트는 구글 드라이브에서 "융합전공 이수 현황 데이터" 라는 이름으로 찾을 수 있다.
 *
 *  주의: K-STAR 앱과 같은 Apps Script 배포를 재사용해도 되지만, 이 앱은
 *  컬렉션 이름 앞에 cvg_ 를 붙여 시트 이름이 서로 겹치지 않게 해두었다.
 * ──────────────────────────────────────────────
 */

var SHEET_NAME = '융합전공 이수 현황 데이터';

function getBook() {
  var id = PropertiesService.getScriptProperties().getProperty('bookId');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* 지워졌으면 새로 만든다 */ }
  }
  var book = SpreadsheetApp.create(SHEET_NAME);
  PropertiesService.getScriptProperties().setProperty('bookId', book.getId());
  return book;
}

function getSheet(book, name) {
  var sh = book.getSheetByName(name);
  if (!sh) {
    sh = book.insertSheet(name);
    sh.getRange(1, 1, 1, 2).setValues([['id', 'data']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAll(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, 2).getValues();
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var id = String(rows[i][0]);
    if (!id) continue;
    var obj;
    try { obj = JSON.parse(rows[i][1] || '{}'); } catch (e) { obj = {}; }
    obj.id = id;
    out.push(obj);
  }
  return out;
}

function findRow(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return 0;
}

function handle(req) {
  var action = req.action;
  var book = getBook();

  if (action === 'ping') {
    return { ok: true, sheet: book.getUrl() };
  }

  var sh = getSheet(book, req.collection);

  if (action === 'all') {
    return { docs: readAll(sh) };
  }

  if (action === 'get') {
    var row = findRow(sh, req.id);
    if (!row) return { doc: null };
    var obj;
    try { obj = JSON.parse(sh.getRange(row, 2).getValue() || '{}'); } catch (e) { obj = {}; }
    obj.id = String(req.id);
    return { doc: obj };
  }

  if (action === 'set' || action === 'update' || action === 'delete') {
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);                    // 동시에 저장할 때 서로 덮어쓰지 않도록
    try {
      var r = findRow(sh, req.id);

      if (action === 'delete') {
        if (r) sh.deleteRow(r);
        return { ok: true };
      }

      var data = req.data || {};
      if (action === 'update' && r) {
        var cur;
        try { cur = JSON.parse(sh.getRange(r, 2).getValue() || '{}'); } catch (e) { cur = {}; }
        for (var k in data) cur[k] = data[k];
        data = cur;
      }
      delete data.id;
      var json = JSON.stringify(data);

      if (r) {
        sh.getRange(r, 2).setValue(json);
      } else {
        sh.appendRow([String(req.id), json]);
      }
      return { ok: true };
    } finally {
      lock.releaseLock();
    }
  }

  return { error: '알 수 없는 요청: ' + action };
}

function doPost(e) {
  var out;
  try {
    out = handle(JSON.parse(e.postData.contents));
  } catch (err) {
    out = { error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  // 브라우저에서 주소를 직접 열었을 때 살아있는지 확인용
  var book = getBook();
  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    message: '융합전공 이수 현황 데이터 서버가 정상 동작 중입니다.',
    sheet: book.getUrl()
  })).setMimeType(ContentService.MimeType.JSON);
}
