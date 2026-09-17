/**
 * K-STAR 현황판 — 데이터 저장 서버
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
 *  만들어진 시트는 구글 드라이브에서 "K-STAR 데이터" 라는 이름으로 찾을 수 있다.
 * ──────────────────────────────────────────────
 */

var SHEET_NAME = 'K-STAR 데이터';

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

  if (action === 'analyze') {
    return analyzeShot(req);
  }

  if (!req.collection) {
    return { error: '어느 자료인지 알 수 없습니다.' };
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

/* ══════════════════════════════════════════════════
 *  캡쳐 이미지 읽기 (analyze)
 *
 *  학생이 STORY+ "역량 대시보드" 화면을 캡쳐해서 올리면
 *   1) 구글 드라이브 "K-STAR 캡쳐" 폴더에 원본을 보관하고
 *   2) 구글 문자인식(OCR)으로 글자를 뽑아
 *   3) F·A·C·E 점수를 찾아 돌려준다.
 *  못 읽으면 사진만 보관하고 학생이 직접 입력하게 한다.
 * ══════════════════════════════════════════════════ */

var SHOT_FOLDER = 'K-STAR 캡쳐';

// STORY+ 화면의 "최소인증 기준점수" — 이 숫자를 기준점 삼아 바로 뒤의 "나의 점수"를 읽는다.
var THRESHOLDS = [
  { area: 'F', base: 350 },
  { area: 'A', base: 200 },
  { area: 'C', base: 280 },
  { area: 'E', base: 720 }
];

function getShotFolder() {
  var it = DriveApp.getFoldersByName(SHOT_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(SHOT_FOLDER);
}

function stamp() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd_HHmmss');
}

function analyzeShot(req) {
  var m = String(req.dataUrl || '').match(/^data:([^;]+);base64,(.*)$/);
  if (!m) return { error: '이미지를 알아보지 못했습니다. 다시 한 번 올려 주세요.' };

  var who = String(req.name || 'capture').replace(/[\\/:*?"<>|]/g, '');
  var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], who + '_' + stamp() + '.jpg');

  var file;
  try {
    file = getShotFolder().createFile(blob);
  } catch (e) {
    return { error: '캡쳐를 저장하지 못했습니다: ' + (e && e.message ? e.message : e) };
  }

  var out = { fileId: file.getId(), fileUrl: file.getUrl() };

  try {
    var values = parseDashboard(ocrText(blob));
    if (values) out.values = values;
    else out.readError = '캡쳐에서 점수를 찾지 못했습니다.';
  } catch (e) {
    out.readError = '자동 읽기에 실패했습니다. (' + (e && e.message ? e.message : e) + ')';
  }
  return out;
}

// 이미지를 구글 문서로 변환하면서 OCR을 돌리고, 글자만 뽑은 뒤 임시 문서는 지운다.
function ocrText(blob) {
  var boundary = '----kstar' + Date.now();
  var head = Utilities.newBlob(
    '--' + boundary + '\r\n' +
    'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify({ name: 'kstar-ocr-tmp', mimeType: 'application/vnd.google-apps.document' }) +
    '\r\n--' + boundary + '\r\n' +
    'Content-Type: ' + blob.getContentType() + '\r\n\r\n'
  ).getBytes();
  var tail = Utilities.newBlob('\r\n--' + boundary + '--\r\n').getBytes();

  var res = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files' +
    '?uploadType=multipart&ocrLanguage=ko&fields=id',
    {
      method: 'post',
      contentType: 'multipart/related; boundary=' + boundary,
      payload: head.concat(blob.getBytes()).concat(tail),
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    }
  );
  if (res.getResponseCode() >= 300) throw new Error('OCR 요청 실패 ' + res.getResponseCode());

  var id = JSON.parse(res.getContentText()).id;
  try {
    return DocumentApp.openById(id).getBody().getText();
  } finally {
    try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { /* 임시파일 정리 실패는 무시 */ }
  }
}

// 글자 속 숫자를 순서대로 뽑는다. ("1,234" 같은 쉼표와 +/- 부호 포함)
function numbersIn(text) {
  var out = [];
  var re = /[+-]?\d[\d,]*/g, m;
  while ((m = re.exec(String(text || '')))) {
    var v = parseInt(m[0].replace(/,/g, ''), 10);
    if (!isNaN(v)) out.push(v);
  }
  return out;
}

/* STORY+ 화면은 영역마다 [기준점수, 나의 점수, 점수차] 가 이 순서로 나오고
   점수차 = 나의 점수 − 기준점수 다. 이 관계가 맞는 묶음만 믿는다. */
function parseDashboard(text) {
  var nums = numbersIn(text);
  var got = {}, found = 0;

  // 1순위: 아는 기준점수(350/200/280/720)를 찾아 그 뒤 숫자를 읽는다.
  for (var t = 0; t < THRESHOLDS.length; t++) {
    var base = THRESHOLDS[t].base;
    for (var i = 0; i + 2 < nums.length; i++) {
      if (nums[i] === base && nums[i + 1] - base === nums[i + 2]) {
        got[THRESHOLDS[t].area] = nums[i + 1];
        found++;
        break;
      }
    }
  }
  if (found === 4) return got;

  // 2순위: 기준점수가 학과마다 다를 수 있으니, 관계가 맞는 묶음을 순서대로 F·A·C·E 에 넣는다.
  var trips = [];
  for (var j = 0; j + 2 < nums.length; j++) {
    if (nums[j] >= 50 && nums[j + 1] >= 0 && nums[j + 1] - nums[j] === nums[j + 2]) {
      trips.push(nums[j + 1]);
      j += 2;
    }
  }
  if (trips.length === 4) return { F: trips[0], A: trips[1], C: trips[2], E: trips[3] };

  return found ? got : null;   // 4개를 다 못 찾으면 찾은 것만이라도 돌려준다
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
    message: 'K-STAR 데이터 서버가 정상 동작 중입니다.',
    sheet: book.getUrl()
  })).setMimeType(ContentService.MimeType.JSON);
}
