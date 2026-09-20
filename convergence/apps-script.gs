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
 *
 * ── 이미 설치돼 있는데 코드만 새로 바꾸는 경우(업데이트) ──────
 *  이미 index.html의 API_URL이 채워져 있다면(=이미 배포돼 있다면), 주소를 새로
 *  받을 필요 없이 아래처럼 "같은 배포"에 코드만 새로 반영하면 된다.
 *  1. script.google.com 에서 이 앱에 쓰는 프로젝트를 연다.
 *  2. 코드를 전부 지우고 이 파일 내용으로 통째로 바꿔서 저장한다.
 *  3. 오른쪽 위 "배포" → "배포 관리"를 누른다.
 *  4. 이미 있는 배포 옆 연필(수정) 아이콘을 누른다.
 *  5. "버전"을 "새 버전"으로 바꾸고 "배포"를 누른다.
 *     (이렇게 해야 웹 앱 주소가 안 바뀌고, index.html의 API_URL도 그대로 써도 된다.
 *      "새 배포"를 누르면 주소가 바뀌어서 index.html도 같이 고쳐야 하니 주의.)
 *  6. DEFAULT_MAJOR_ID 값이 index.html의 DEFAULT_MAJOR_ID와 같은 문자열인지 확인한다
 *     (지금은 둘 다 'global-k').
 *
 * ── 이 버전에서 바뀐 것(개인정보 보호) ──────────────
 *  예전 버전은 "학번+비밀번호가 맞는지"를 브라우저(클라이언트) 쪽에서만 확인하고,
 *  서버는 누가 요청했는지 전혀 확인하지 않았다. 그래서 이 웹앱 주소만 알면 누구나
 *  전체 학생 이름·학번·학과·이수과목·비밀번호 해시까지 다 읽고 고치고 지울 수 있었다.
 *  지금은 로그인/회원가입만 비밀번호를 다루고, 그 결과로 받은 세션 토큰이 있어야만
 *  학생·관리자 데이터를 읽거나 고칠 수 있다. 관리자는 자기 융합전공(majorId) 학생만
 *  볼 수 있고, 학생은 자기 자신의 기록만 볼 수 있다. 비밀번호도 사용자마다 다른
 *  임의의 salt를 붙여 여러 번 반복해서 해시하므로(PBKDF2 방식), 해시값이 유출돼도
 *  원래 비밀번호를 알아내기 훨씬 어렵다. 예전 방식(salt 없는 SHA-256 1회)으로 만든
 *  계정도 다음 로그인 때 자동으로 새 방식으로 바뀐다.
 * ──────────────────────────────────────────────
 */

var SHEET_NAME = '융합전공 이수 현황 데이터';
var SESSION_HOURS = 12;          // 로그인 세션 유지 시간. 지나면 다시 로그인해야 한다.
var PBKDF2_ITERATIONS = 8000;    // 비밀번호 해시 반복 횟수(Apps Script 실행시간 한도 안에서 최대한 크게).
var DEFAULT_MAJOR_ID = 'global-k';

/* ============ 스프레드시트 기본 도구 ============ */

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

function readRow(sh, row) {
  if (!row) return null;
  try { return JSON.parse(sh.getRange(row, 2).getValue() || '{}'); } catch (e) { return {}; }
}

function writeRow(sh, row, id, data) {
  var clean = {};
  for (var k in data) if (k !== 'id') clean[k] = data[k];
  var json = JSON.stringify(clean);
  if (row) sh.getRange(row, 2).setValue(json);
  else sh.appendRow([String(id), json]);
}

function fail(code) {
  var e = new Error(code);
  e.isAppError = true;
  throw e;
}

/* ============ 비밀번호 해시(PBKDF2-HMAC-SHA256, 사용자별 임의 salt) ============ */

function bytesToHex(bytes) {
  var hex = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i]; if (b < 0) b += 256;
    hex += (b < 16 ? '0' : '') + b.toString(16);
  }
  return hex;
}

// RFC 2898 PBKDF2 한 블록(dkLen=32바이트)짜리 계산. salt는 사용자별로 매번 새로 만든다.
function pbkdf2Hex(password, saltStr, iterations) {
  var passwordBytes = Utilities.newBlob(String(password)).getBytes();
  var saltBytes = Utilities.newBlob(String(saltStr)).getBytes();
  var input = saltBytes.concat([0, 0, 0, 1]); // INT_BE_32(1)
  var u = Utilities.computeHmacSha256Signature(input, passwordBytes);
  var t = u.slice();
  for (var i = 1; i < iterations; i++) {
    u = Utilities.computeHmacSha256Signature(u, passwordBytes);
    for (var j = 0; j < t.length; j++) t[j] = t[j] ^ u[j];
  }
  return bytesToHex(t);
}

function legacySha256Hex(str) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(str));
  return bytesToHex(digest);
}

function hashNewPassword(password) {
  var salt = Utilities.getUuid();
  return { passwordSalt: salt, passwordHash: pbkdf2Hex(password, salt, PBKDF2_ITERATIONS) };
}

// 맞으면 true를 돌려주고, upgrade가 필요하면(예전 방식 계정) doc에 새 해시를 채워 넣는다
// (호출한 쪽에서 그 doc을 저장해야 실제로 반영된다).
function verifyPassword(password, doc) {
  if (doc.passwordSalt) {
    return pbkdf2Hex(password, doc.passwordSalt, PBKDF2_ITERATIONS) === doc.passwordHash;
  }
  if (doc.passwordHash && legacySha256Hex(password) === doc.passwordHash) {
    var upgraded = hashNewPassword(password);
    doc.passwordSalt = upgraded.passwordSalt;
    doc.passwordHash = upgraded.passwordHash;
    return true;
  }
  return false;
}

function randomTempPassword() {
  // 사람이 옮겨 적기 쉬운 8자리 숫자.
  var n = Math.floor(10000000 + Math.random() * 89999999);
  return String(n);
}

/* ============ 로그인 세션 ============ */

function randomToken() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function createSession(book, role, id, majorId) {
  var sh = getSheet(book, 'cvg_sessions');
  var token = randomToken();
  var expiresAt = new Date(Date.now() + SESSION_HOURS * 3600 * 1000).toISOString();
  sh.appendRow([token, JSON.stringify({ role: role, id: String(id), majorId: majorId, expiresAt: expiresAt })]);
  return token;
}

// 토큰이 없거나, 없는 토큰이거나, 만료됐으면 null. 만료된 세션은 지운다.
function getSession(book, token) {
  if (!token) return null;
  var sh = getSheet(book, 'cvg_sessions');
  var row = findRow(sh, token);
  if (!row) return null;
  var obj = readRow(sh, row);
  if (!obj || !obj.expiresAt || new Date(obj.expiresAt).getTime() < Date.now()) {
    sh.deleteRow(row);
    return null;
  }
  return obj;
}

function destroySession(book, token) {
  if (!token) return;
  var sh = getSheet(book, 'cvg_sessions');
  var row = findRow(sh, token);
  if (row) sh.deleteRow(row);
}

function requireSession(book, token) {
  var session = getSession(book, token);
  if (!session) fail('SESSION_EXPIRED');
  return session;
}

/* ============ 개인정보가 담긴 필드를 응답에서 제거 ============ */

function stripSecrets(doc) {
  if (!doc) return doc;
  var out = {};
  for (var k in doc) {
    if (k === 'passwordHash' || k === 'passwordSalt') continue;
    out[k] = doc[k];
  }
  return out;
}

/* ============ cvg_config 문서 id ============ */

function configDocId(majorId) {
  return majorId === DEFAULT_MAJOR_ID ? 'app' : ('major_' + majorId);
}

/* ============ 요청 처리 ============ */

function handle(req) {
  var action = req.action;
  var book = getBook();

  if (action === 'ping') {
    return { ok: true, sheet: book.getUrl() };
  }

  // ---- 회원가입/로그인/비밀번호(비밀번호만 다루는 전용 요청) ----
  if (action === 'registerStudent') return doRegisterStudent(book, req);
  if (action === 'registerAdmin') return doRegisterAdmin(book, req);
  if (action === 'login') return doLogin(book, req);
  if (action === 'logout') { destroySession(book, req.token); return { ok: true }; }
  if (action === 'changePassword') return doChangePassword(book, req);
  if (action === 'resetPassword') return doResetPassword(book, req);

  // ---- 일반 데이터(과목/이수정보/설정 등) — 여기서부터는 collection별 권한 검사 ----
  var col = req.collection;

  if (action === 'get') return handleGetAction(book, col, req);
  if (action === 'all') return doAll(book, col, req);
  if (action === 'update') return doUpdate(book, col, req);
  if (action === 'delete') return doDelete(book, col, req);

  return { error: 'UNKNOWN_ACTION' };
}

/* ---- 회원가입 ---- */

function doRegisterStudent(book, req) {
  if (!req.id || !req.name) fail('MISSING_FIELDS');
  if (!req.dept) fail('MISSING_DEPT');
  if (!req.password || req.password.length < 6) fail('WEAK_PASSWORD');
  var sh = getSheet(book, 'cvg_students');
  if (findRow(sh, req.id)) fail('DUPLICATE_ID');
  var hashed = hashNewPassword(req.password);
  var doc = Object.assign({
    name: req.name, dept: req.dept || '', year: req.year || '',
    courses: [], recognizedSemesters: '', majorId: req.majorId || DEFAULT_MAJOR_ID,
    microdegrees: [], thesisRequired: false, thesisDone: false,
    createdAt: new Date().toISOString()
  }, hashed);
  writeRow(sh, 0, req.id, doc);
  var token = createSession(book, 'student', req.id, doc.majorId);
  var out = stripSecrets(doc); out.id = String(req.id);
  return { token: token, doc: out };
}

function doRegisterAdmin(book, req) {
  if (!req.id || !req.name) fail('MISSING_FIELDS');
  if (!req.password || req.password.length < 6) fail('WEAK_PASSWORD');
  var majorId = req.majorId || DEFAULT_MAJOR_ID;
  var adminSh = getSheet(book, 'cvg_admins');
  var allAdmins = readAll(adminSh);
  var sameMajorAdmins = allAdmins.filter(function (a) { return (a.majorId || DEFAULT_MAJOR_ID) === majorId; });
  if (sameMajorAdmins.length > 0) {
    var cfgSh = getSheet(book, 'cvg_config');
    var cfgRow = findRow(cfgSh, configDocId(majorId));
    var cfg = readRow(cfgSh, cfgRow) || {};
    var required = cfg.adminCode || '';
    if (!required) fail('ADMIN_CODE_LOCKED');
    if (req.code !== required) fail('ADMIN_CODE_INVALID');
  }
  if (findRow(adminSh, req.id)) fail('DUPLICATE_ID');
  var hashed = hashNewPassword(req.password);
  var doc = Object.assign({ name: req.name, majorId: majorId, createdAt: new Date().toISOString() }, hashed);
  writeRow(adminSh, 0, req.id, doc);
  var token = createSession(book, 'admin', req.id, majorId);
  var out = stripSecrets(doc); out.id = String(req.id);
  return { token: token, doc: out };
}

/* ---- 로그인 ---- */

function doLogin(book, req) {
  var col = req.collection === 'cvg_admins' ? 'cvg_admins' : 'cvg_students';
  var role = col === 'cvg_admins' ? 'admin' : 'student';
  var sh = getSheet(book, col);
  var row = findRow(sh, req.id);
  if (!row) fail('NOT_REGISTERED');
  var doc = readRow(sh, row);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (!verifyPassword(req.password, doc)) fail('INVALID_CREDENTIALS');
    // verifyPassword가 예전 방식 계정을 새 방식으로 갱신해뒀을 수 있으니 다시 저장한다.
    writeRow(sh, row, req.id, doc);
  } finally {
    lock.releaseLock();
  }
  var majorId = doc.majorId || DEFAULT_MAJOR_ID;
  var token = createSession(book, role, req.id, majorId);
  var out = stripSecrets(doc); out.id = String(req.id);
  return { token: token, doc: out };
}

/* ---- 비밀번호 변경(본인) / 초기화(관리자가 학생 대신) ---- */

function doChangePassword(book, req) {
  var session = requireSession(book, req.token);
  if (!req.newPassword || req.newPassword.length < 6) fail('WEAK_PASSWORD');
  var col = session.role === 'admin' ? 'cvg_admins' : 'cvg_students';
  var sh = getSheet(book, col);
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var row = findRow(sh, session.id);
    var doc = readRow(sh, row);
    if (!doc) fail('NOT_FOUND');
    if (!verifyPassword(req.oldPassword, doc)) fail('INVALID_CREDENTIALS');
    var hashed = hashNewPassword(req.newPassword);
    doc.passwordSalt = hashed.passwordSalt;
    doc.passwordHash = hashed.passwordHash;
    doc.mustChangePassword = false;
    writeRow(sh, row, session.id, doc);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

function doResetPassword(book, req) {
  var session = requireSession(book, req.token);
  if (session.role !== 'admin') fail('FORBIDDEN');
  var sh = getSheet(book, 'cvg_students');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var row = findRow(sh, req.studentId);
    var doc = readRow(sh, row);
    if (!doc) fail('NOT_FOUND');
    if ((doc.majorId || DEFAULT_MAJOR_ID) !== session.majorId) fail('FORBIDDEN');
    var tempPw = randomTempPassword();
    var hashed = hashNewPassword(tempPw);
    doc.passwordSalt = hashed.passwordSalt;
    doc.passwordHash = hashed.passwordHash;
    doc.mustChangePassword = true;
    writeRow(sh, row, req.studentId, doc);
    return { ok: true, tempPassword: tempPw };
  } finally {
    lock.releaseLock();
  }
}

/* ---- 일반 데이터 읽기/쓰기: collection별 권한 검사 ---- */

function handleGetAction(book, col, req) {
  if (col === 'cvg_config') {
    var sh = getSheet(book, col);
    var doc = readRow(sh, findRow(sh, req.id)) || null;
    if (doc) {
      doc = Object.assign({}, doc);
      var session = getSession(book, req.token);
      var isOwnerAdmin = session && session.role === 'admin' && configDocId(session.majorId) === req.id;
      if (!isOwnerAdmin) delete doc.adminCode;
    }
    return { doc: doc };
  }
  if (col === 'cvg_students') {
    var session = requireSession(book, req.token);
    var sh2 = getSheet(book, col);
    var row = findRow(sh2, req.id);
    var doc2 = readRow(sh2, row);
    if (!doc2) return { doc: null };
    var majorId = doc2.majorId || DEFAULT_MAJOR_ID;
    var allowed = (session.role === 'student' && String(session.id) === String(req.id)) ||
      (session.role === 'admin' && session.majorId === majorId);
    if (!allowed) fail('FORBIDDEN');
    var out = stripSecrets(doc2); out.id = String(req.id);
    return { doc: out };
  }
  fail('FORBIDDEN');
}

function doAll(book, col, req) {
  if (col !== 'cvg_students') fail('FORBIDDEN'); // cvg_admins 목록은 절대 통째로 내려주지 않는다
  var session = requireSession(book, req.token);
  if (session.role !== 'admin') fail('FORBIDDEN');
  var sh = getSheet(book, col);
  var all = readAll(sh);
  var mine = all.filter(function (s) { return (s.majorId || DEFAULT_MAJOR_ID) === session.majorId; });
  return { docs: mine.map(stripSecrets).map(function (d, i) { d.id = mine[i].id; return d; }) };
}

function doUpdate(book, col, req) {
  if (col === 'cvg_config') {
    var session = requireSession(book, req.token);
    if (session.role !== 'admin') fail('FORBIDDEN');
    if (configDocId(session.majorId) !== req.id) fail('FORBIDDEN');
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      var sh = getSheet(book, col);
      var row = findRow(sh, req.id);
      var cur = readRow(sh, row) || {};
      var data = req.data || {};
      for (var k in data) cur[k] = data[k];
      writeRow(sh, row, req.id, cur);
    } finally {
      lock.releaseLock();
    }
    return { ok: true };
  }
  if (col === 'cvg_students') {
    var session2 = requireSession(book, req.token);
    var lock2 = LockService.getScriptLock();
    lock2.waitLock(20000);
    try {
      var sh2 = getSheet(book, col);
      var row2 = findRow(sh2, req.id);
      if (!row2) fail('NOT_FOUND'); // 새 학생은 registerStudent로만 만들 수 있다
      var cur2 = readRow(sh2, row2);
      var majorId = cur2.majorId || DEFAULT_MAJOR_ID;
      var allowed = (session2.role === 'student' && String(session2.id) === String(req.id)) ||
        (session2.role === 'admin' && session2.majorId === majorId);
      if (!allowed) fail('FORBIDDEN');
      var data2 = Object.assign({}, req.data || {});
      // 비밀번호·소속 융합전공은 이 통로로 못 바꾸게 막는다(전용 요청으로만 변경).
      delete data2.passwordHash; delete data2.passwordSalt; delete data2.majorId; delete data2.mustChangePassword;
      for (var k2 in data2) cur2[k2] = data2[k2];
      writeRow(sh2, row2, req.id, cur2);
    } finally {
      lock2.releaseLock();
    }
    return { ok: true };
  }
  fail('FORBIDDEN');
}

function doDelete(book, col, req) {
  if (col !== 'cvg_students') fail('FORBIDDEN');
  var session = requireSession(book, req.token);
  if (session.role !== 'admin') fail('FORBIDDEN');
  var sh = getSheet(book, col);
  var row = findRow(sh, req.id);
  if (!row) return { ok: true };
  var doc = readRow(sh, row);
  if ((doc.majorId || DEFAULT_MAJOR_ID) !== session.majorId) fail('FORBIDDEN');
  sh.deleteRow(row);
  return { ok: true };
}

/* ============ 진입점 ============ */

function doPost(e) {
  var out;
  try {
    out = handle(JSON.parse(e.postData.contents));
  } catch (err) {
    out = { error: (err && err.isAppError) ? err.message : String(err && err.message ? err.message : err) };
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
