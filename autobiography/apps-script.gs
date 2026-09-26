/**
 * 나의 자서전 — 구글 시트 자동 저장 서버
 *
 * 하는 일: 어머니 휴대폰의 자서전 앱이 쓰신 이야기를 이 구글 계정의 스프레드시트
 * ("나의 자서전")에, 사진은 구글 드라이브 폴더("나의 자서전 사진")에 저장한다.
 * 스프레드시트와 사진은 이 계정 주인만 볼 수 있다.
 *
 * 앱 코드와 이 서버 주소는 공개 저장소에 있지 않고, 비밀 열쇠(key)는 이 계정의
 * 스크립트 속성에만 있다. 앱은 링크에 담긴 열쇠를 가진 경우에만 읽고 쓸 수 있다.
 *
 * ── 설치 방법 (처음 한 번) ─────────────────────────────
 *  1. script.google.com 에 접속해 "새 프로젝트"를 만든다. 이름은 "나의 자서전 서버" 정도.
 *  2. 기본으로 들어있는 코드를 모두 지우고 이 파일 내용을 붙여넣은 뒤 저장(💾)한다.
 *  3. 위쪽 함수 고르는 칸에서 "링크만들기"를 고르고 ▶ 실행.
 *     권한 허용 창이 뜨면 계정을 고르고 "고급" → "(안전하지 않음)으로 이동" → "허용".
 *     (내가 만든 스크립트라서 구글이 확인을 거치지 않았다는 뜻이다.)
 *     → 스프레드시트·사진 폴더·비밀 열쇠가 만들어진다. 아래 "실행 로그"에 안내가 나온다.
 *  4. 오른쪽 위 "배포" → "새 배포" → 톱니바퀴에서 "웹 앱"
 *     · 실행 계정: 나   · 액세스 권한이 있는 사용자: 모든 사용자  → "배포"
 *  5. 다시 "링크만들기"를 ▶ 실행하면 실행 로그에 링크 두 개가 나온다.
 *     · 어머니용 링크 — 카톡으로 어머니께 보낸다.
 *     · 보기용 링크   — 내 휴대폰에서 연다. 어머니가 쓰신 이야기를 책으로 본다.
 *
 * ── 코드만 새로 바꾸는 경우 ─────────────────────────────
 *  코드를 붙여넣고 저장 → "배포" → "배포 관리" → 연필(수정) → 버전 "새 버전" → "배포".
 *  ※ "새 배포"를 하면 주소가 바뀌어 어머니 링크가 더는 동작하지 않는다.
 * ──────────────────────────────────────────────────
 */

var APP_URL = 'https://kjstkd-autobiography.vercel.app/';
var BOOK_NAME = '나의 자서전';
var FOLDER_NAME = '나의 자서전 사진';
var STORY_SHEET = '이야기';
var INFO_SHEET = '정보';
var PHOTO_SHEET = '_사진';
var STORY_HEAD = ['순서', '장', '질문', '이야기', '사진', '고친 때', 'qid', 'photoIds'];
var COL = { order: 1, chapter: 2, question: 3, text: 4, photos: 5, updated: 6, qid: 7, photoIds: 8 };

/* ============ 저장소 ============ */

function props() { return PropertiesService.getScriptProperties(); }

function getBook() {
  var id = props().getProperty('bookId');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* 지워졌으면 새로 만든다 */ }
  }
  var book = SpreadsheetApp.create(BOOK_NAME);
  props().setProperty('bookId', book.getId());
  var first = book.getSheets()[0];
  first.setName(STORY_SHEET);
  setupStorySheet(first);
  return book;
}

function setupStorySheet(sh) {
  sh.getRange(1, 1, 1, STORY_HEAD.length).setValues([STORY_HEAD]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setColumnWidth(COL.order, 50);
  sh.setColumnWidth(COL.chapter, 130);
  sh.setColumnWidth(COL.question, 260);
  sh.setColumnWidth(COL.text, 480);
  sh.setColumnWidth(COL.photos, 200);
  sh.setColumnWidth(COL.updated, 140);
  sh.getRange('C:E').setWrap(true).setVerticalAlignment('top');
  sh.hideColumns(COL.qid, 2);  // qid, photoIds — 앱이 쓰는 칸
}

function getSheet(book, name) {
  var sh = book.getSheetByName(name);
  if (sh) return sh;
  sh = book.insertSheet(name);
  if (name === STORY_SHEET) setupStorySheet(sh);
  else if (name === INFO_SHEET) { sh.getRange(1, 1, 1, 2).setValues([['항목', '내용']]).setFontWeight('bold'); sh.setColumnWidth(2, 300); }
  else if (name === PHOTO_SHEET) { sh.getRange(1, 1, 1, 4).setValues([['photoId', 'qid', 'fileId', 'url']]); sh.hideSheet(); }
  return sh;
}

function getFolder() {
  var id = props().getProperty('folderId');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* 지워졌으면 새로 만든다 */ }
  }
  var folder = DriveApp.createFolder(FOLDER_NAME);
  props().setProperty('folderId', folder.getId());
  return folder;
}

function getKey() {
  var k = props().getProperty('key');
  if (!k) {
    k = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
    props().setProperty('key', k);
  }
  return k;
}

// 첫 번째 열(또는 col 열)에서 값이 같은 줄 번호. 없으면 0.
function findRow(sh, col, value) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var vals = sh.getRange(2, col, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) if (String(vals[i][0]) === String(value)) return i + 2;
  return 0;
}

function withLock(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// 시트는 =, +, -, @ 로 시작하는 글을 수식으로 읽으므로 앞에 ' 를 붙여 글로 두고,
// 한 칸 한도(5만 자)를 넘지 않게 자른다. 전체 글은 어머니 휴대폰에 그대로 남아 있다.
function asText(v) {
  var t = String(v == null ? '' : v);
  if (t.length > 49000) t = t.slice(0, 49000) + ' …(너무 길어 잘렸어요)';
  return /^[=+\-@]/.test(t) ? "'" + t : t;
}

function fail(msg) { var e = new Error(msg); e.isAppError = true; throw e; }

/* ============ 요청 처리 ============ */

function handle(req) {
  var key = props().getProperty('key');
  if (!key) fail('서버 준비가 아직 안 됐어요. Apps Script 에서 "링크만들기"를 먼저 실행해 주세요.');
  if (!req || req.k !== key) fail('열쇠가 맞지 않아요. 새 링크로 다시 열어 주세요.');
  switch (req.action) {
    case 'ping': return { ok: true };
    case 'photo': return withLock(function () { return savePhoto(req); });
    case 'save': return withLock(function () { return saveStories(req); });
    case 'load': return loadAll();
  }
  fail('알 수 없는 요청: ' + req.action);
}

function photoMap(book) {
  var sh = getSheet(book, PHOTO_SHEET), last = sh.getLastRow(), map = {};
  if (last < 2) return map;
  sh.getRange(2, 1, last - 1, 4).getValues().forEach(function (r) {
    if (r[0]) map[String(r[0])] = { qid: String(r[1]), fileId: String(r[2]), url: String(r[3]) };
  });
  return map;
}

function savePhoto(req) {
  if (!/^p[0-9a-z]+$/.test(String(req.id || ''))) fail('사진 번호가 이상해요.');
  var book = getBook(), sh = getSheet(book, PHOTO_SHEET);
  if (findRow(sh, 1, req.id)) return { ok: true, already: true };
  var bytes = Utilities.base64Decode(String(req.data || ''));
  if (!bytes.length) fail('사진 내용이 비어 있어요.');
  var file = getFolder().createFile(Utilities.newBlob(bytes, 'image/jpeg', String(req.qid).replace(/[^\w-]/g, '_') + '-' + req.id + '.jpg'));
  sh.appendRow([req.id, req.qid, file.getId(), file.getUrl()]);
  return { ok: true };
}

function saveStories(req) {
  var book = getBook(), sh = getSheet(book, STORY_SHEET), photos = photoMap(book);
  var tz = Session.getScriptTimeZone();
  (req.answers || []).forEach(function (a) {
    var urls = (a.photos || []).map(function (id) { return photos[id] ? photos[id].url : ''; }).filter(String);
    var row = [
      Number(a.order) || 0, asText(a.chapter), asText(a.question), asText(a.text),
      urls.join('\n'), Utilities.formatDate(new Date(Number(a.updated) || Date.now()), tz, 'yyyy-MM-dd HH:mm'),
      String(a.qid), (a.photos || []).join(',')
    ];
    var at = findRow(sh, COL.qid, a.qid);
    if (at) sh.getRange(at, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
  });
  (req.removed || []).forEach(function (qid) {
    var at = findRow(sh, COL.qid, qid);
    if (at) sh.deleteRow(at);
  });
  if (sh.getLastRow() > 2) sh.getRange(2, 1, sh.getLastRow() - 1, STORY_HEAD.length).sort(COL.order);

  if (req.profile) {
    var info = getSheet(book, INFO_SHEET), p = req.profile;
    info.getRange(2, 1, 3, 2).setValues([['성함', p.name || ''], ['태어난 해', p.birthYear || ''], ['고향', p.hometown || '']]);
  }
  var infoSh = getSheet(book, INFO_SHEET);
  infoSh.getRange(5, 1, 1, 2).setValues([['마지막 저장', Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm')]]);
  return { ok: true };
}

function loadAll() {
  var book = getBook(), sh = getSheet(book, STORY_SHEET), photos = photoMap(book);
  var info = getSheet(book, INFO_SHEET).getRange(2, 2, 3, 1).getValues();
  var out = { ok: true, profile: { name: String(info[0][0] || ''), birthYear: String(info[1][0] || ''), hometown: String(info[2][0] || '') }, answers: [], photos: {} };
  var last = sh.getLastRow();
  if (last < 2) return out;
  var rows = sh.getRange(2, 1, last - 1, STORY_HEAD.length).getValues();
  rows.forEach(function (r) {
    var ids = String(r[COL.photoIds - 1] || '').split(',').filter(String);
    ids.forEach(function (id) {
      if (out.photos[id] || !photos[id]) return;
      try { out.photos[id] = Utilities.base64Encode(DriveApp.getFileById(photos[id].fileId).getBlob().getBytes()); } catch (e) {}
    });
    out.answers.push({ qid: String(r[COL.qid - 1]), question: String(r[COL.question - 1]), text: String(r[COL.text - 1]), photos: ids });
  });
  return out;
}

/* ============ 진입점 ============ */

function doPost(e) {
  var out;
  try {
    out = handle(JSON.parse(e.postData.contents));
  } catch (err) {
    out = { error: (err && err.isAppError) ? err.message : String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  // 브라우저에서 주소를 직접 열었을 때 살아있는지 확인용 (자료는 아무것도 내주지 않는다)
  return ContentService.createTextOutput(JSON.stringify({ ok: true, message: '나의 자서전 저장 서버가 정상 동작 중입니다.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============ 편집기에서 직접 실행하는 함수 ============ */

// 스프레드시트·사진 폴더·열쇠를 만들고, 배포가 끝났으면 어머니용/보기용 링크를 알려준다.
// 링크의 성함·보내는 사람을 바꾸려면 아래 두 줄을 고친다.
var MOM_NAME = '정영순';
var FROM = '아들 종수';

function 링크만들기() {
  var book = getBook(); getSheet(book, INFO_SHEET); getSheet(book, PHOTO_SHEET); getFolder();
  var key = getKey();
  Logger.log('스프레드시트: ' + book.getUrl());
  Logger.log('사진 폴더: ' + getFolder().getUrl());
  var url = ScriptApp.getService().getUrl();
  var m = url && url.match(/\/macros\/s\/([^/]+)\/(exec|dev)/);
  if (!m || m[2] !== 'exec') {
    Logger.log('아직 웹 앱으로 배포하지 않았어요. "배포" → "새 배포" → 웹 앱(실행 계정: 나, 액세스: 모든 사용자)을 한 뒤 다시 실행해 주세요.');
    return;
  }
  var base = APP_URL + '?s=' + encodeURIComponent(m[1]) + '&k=' + encodeURIComponent(key);
  Logger.log('───── 어머니용 링크 (카톡으로 보내세요) ─────');
  Logger.log(base + '&name=' + encodeURIComponent(MOM_NAME) + '&from=' + encodeURIComponent(FROM));
  Logger.log('───── 보기용 링크 (내 휴대폰에서 여세요) ─────');
  Logger.log(base + '&view=1');
  Logger.log('두 링크 모두 다른 사람에게 보여주지 마세요. 열쇠가 들어 있어요.');
}
