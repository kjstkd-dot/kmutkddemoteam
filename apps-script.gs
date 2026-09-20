/**
 * K-STAR 현황판 — 데이터 저장 서버
 *
 * 하는 일: 학생과 관리자가 앱에서 입력한 내용을 구글 스프레드시트에 저장하고 돌려준다.
 * 학생은 구글 로그인이 필요 없다. 이 코드가 관리자 계정 권한으로 대신 기록하기 때문.
 *
 * 로그인한 사람만, 자기에게 허용된 자료만 읽고 쓸 수 있다.
 *   · 학생  — 자기 기록 하나만
 *   · 관리자 — 전체 명단
 * 주소(웹 앱 URL)를 안다고 해서 남의 기록을 볼 수 있는 것이 아니다.
 *
 * ── 설치 방법 ──────────────────────────────────────
 *  1. script.google.com 에 접속해 기존 K-STAR 프로젝트를 연다(처음이면 "새 프로젝트").
 *  2. 기본으로 들어있는 코드를 모두 지우고 이 파일 내용을 붙여넣는다.
 *  3. 오른쪽 위 "배포" → "배포 관리" → 연필(수정) → 버전 "새 버전" → "배포".
 *     ※ "새 배포"를 하면 주소가 바뀌어 앱이 서버를 못 찾는다. 반드시 기존 배포를 수정할 것.
 *     (처음 설치라면 "새 배포" → 웹 앱 → 실행 계정: 나 / 액세스: 모든 사용자)
 *  4. 권한 허용 창이 뜨면 허용한다.
 *  5. 처음 설치라면 나오는 "웹 앱 URL"(.../exec)을 index.html 의 API_URL 에 넣는다.
 *
 *  스프레드시트는 처음 실행될 때 자동으로 만들어진다("K-STAR 데이터").
 * ──────────────────────────────────────────────
 */

var SHEET_NAME = 'K-STAR 데이터';
var SESSION_DAYS = 30;           // 로그인 유지 기간. 지나면 다시 로그인해야 한다.
var PBKDF2_ITERATIONS = 8000;    // 비밀번호 해시 반복 횟수(Apps Script 실행시간 한도 안에서 최대한 크게)
var MAX_SESSION_ROWS = 400;      // 이보다 많아지면 만료된 세션을 청소한다

/* ============ 스프레드시트 기본 ============ */

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

function nowIso() { return new Date().toISOString(); }

function withLock(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ============ 비밀번호 (PBKDF2-HMAC-SHA256, 사용자별 임의 salt) ============ */

function bytesToHex(bytes) {
  var hex = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i]; if (b < 0) b += 256;
    hex += (b < 16 ? '0' : '') + b.toString(16);
  }
  return hex;
}

// RFC 2898 PBKDF2 한 블록(dkLen=32바이트). salt는 사용자마다 새로 만든다.
function pbkdf2Hex(password, saltStr, iterations) {
  var passwordBytes = Utilities.newBlob(String(password)).getBytes();
  var saltBytes = Utilities.newBlob(String(saltStr)).getBytes();
  var input = saltBytes.concat([0, 0, 0, 1]);            // INT_BE_32(1)
  var u = Utilities.computeHmacSha256Signature(input, passwordBytes);
  var t = u.slice();
  for (var i = 1; i < iterations; i++) {
    u = Utilities.computeHmacSha256Signature(u, passwordBytes);
    for (var j = 0; j < t.length; j++) t[j] = t[j] ^ u[j];
  }
  return bytesToHex(t);
}

// 예전 방식: 앱이 브라우저에서 SHA-256 한 번 돌려 보내던 값
function legacySha256Hex(str) {
  return bytesToHex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(str)));
}

function hashNewPassword(password) {
  var salt = Utilities.getUuid();
  return { passwordSalt: salt, passwordHash: pbkdf2Hex(password, salt, PBKDF2_ITERATIONS) };
}

// 맞으면 true. 예전 방식 계정이면 doc 안의 해시를 새 방식으로 바꿔 놓는다
// (부른 쪽에서 그 doc을 저장해야 실제로 반영된다).
function verifyPassword(password, doc) {
  if (!doc) return false;
  if (doc.passwordSalt) {
    return pbkdf2Hex(password, doc.passwordSalt, PBKDF2_ITERATIONS) === doc.passwordHash;
  }
  if (doc.passwordHash && legacySha256Hex(password) === doc.passwordHash) {
    var up = hashNewPassword(password);
    doc.passwordSalt = up.passwordSalt;
    doc.passwordHash = up.passwordHash;
    return true;
  }
  return false;
}

/* ============ 로그인 세션 ============ */

function randomToken() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function purgeExpiredSessions(sh) {
  var last = sh.getLastRow();
  if (last - 1 <= MAX_SESSION_ROWS) return;
  var rows = sh.getRange(2, 1, last - 1, 2).getValues();
  var now = Date.now();
  for (var i = rows.length - 1; i >= 0; i--) {
    var obj;
    try { obj = JSON.parse(rows[i][1] || '{}'); } catch (e) { obj = {}; }
    if (!obj.expiresAt || new Date(obj.expiresAt).getTime() < now) sh.deleteRow(i + 2);
  }
}

function createSession(book, role, id, rosterId) {
  var sh = getSheet(book, 'sessions');
  purgeExpiredSessions(sh);
  var token = randomToken();
  sh.appendRow([token, JSON.stringify({
    role: role, id: String(id), rosterId: rosterId || '',
    expiresAt: new Date(Date.now() + SESSION_DAYS * 86400 * 1000).toISOString()
  })]);
  return token;
}

function getSession(book, token) {
  if (!token) return null;
  var sh = getSheet(book, 'sessions');
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
  var sh = getSheet(book, 'sessions');
  var row = findRow(sh, token);
  if (row) sh.deleteRow(row);
}

function requireSession(book, token) {
  var s = getSession(book, token);
  if (!s) fail('SESSION_EXPIRED');
  return s;
}

function requireAdmin(book, token) {
  var s = requireSession(book, token);
  if (s.role !== 'admin') fail('FORBIDDEN');
  return s;
}

/* ============ 응답에서 비밀번호 관련 항목 제거 ============ */

function stripSecrets(doc) {
  if (!doc) return doc;
  var out = {};
  for (var k in doc) {
    if (k === 'passwordHash' || k === 'passwordSalt') continue;
    out[k] = doc[k];
  }
  return out;
}

/* ============ 명단(roster) 공통 ============ */

function sameName(a, b) {
  return String(a == null ? '' : a).replace(/\s+/g, '') === String(b == null ? '' : b).replace(/\s+/g, '');
}

function newRosterId(id) {
  return 'm_' + (id || String(Date.now()).slice(-8)) + '_' +
         Math.random().toString(36).slice(2, 6);
}

// 학생이 자기 기록인지 확인
function ownRoster(session, doc, id) {
  if (session.role === 'admin') return true;
  if (!doc) return false;
  return String(session.rosterId) === String(id) || String(doc.claimedBy) === String(session.id);
}

/* ============ 요청 처리 ============ */

function handle(req) {
  var action = req.action;
  var book = getBook();

  if (action === 'ping') return { ok: true };

  // ---- 로그인·가입·비밀번호 ----
  if (action === 'login')           return doLogin(book, req);
  if (action === 'logout')          { destroySession(book, req.token); return { ok: true }; }
  if (action === 'registerStudent') return doRegisterStudent(book, req);
  if (action === 'registerAdmin')   return doRegisterAdmin(book, req);
  if (action === 'saveProfile')     return doSaveProfile(book, req);

  // ---- 관리자 전용 동작 ----
  if (action === 'deleteMember')    return doDeleteMember(book, req);
  if (action === 'resetStudent')    return doResetStudent(book, req);

  // ---- 캡쳐 보관 ----
  if (action === 'analyze')         { requireSession(book, req.token); return analyzeShot(req); }

  // ---- 명단·설정 읽고 쓰기 ----
  var col = req.collection;
  if (col !== 'roster' && col !== 'config') fail('FORBIDDEN');

  if (action === 'get')    return doGetDoc(book, col, req);
  if (action === 'all')    return doAllDocs(book, col, req);
  if (action === 'set' || action === 'update') return doWriteDoc(book, col, req, action);
  if (action === 'delete') fail('FORBIDDEN');   // 삭제는 deleteMember 로만

  return { error: 'UNKNOWN_ACTION' };
}

/* ---- 로그인 ---- */

function doLogin(book, req) {
  var isAdmin = req.role === 'admin';
  var sh = getSheet(book, isAdmin ? 'admins' : 'students');
  var row = findRow(sh, req.id);
  if (!row) fail('NOT_REGISTERED');

  var doc = withLock(function () {
    var d = readRow(sh, row);
    if (!verifyPassword(req.password, d)) fail('INVALID_CREDENTIALS');
    writeRow(sh, row, req.id, d);     // 예전 방식 계정이었다면 새 방식으로 갱신돼 있다
    return d;
  });

  if (isAdmin) {
    return { token: createSession(book, 'admin', req.id, ''), name: doc.name };
  }

  // 가리키던 명단 기록이 사라졌거나 잘못 연결됐으면 학번·이름으로 다시 찾아 이어붙인다
  var rs = getSheet(book, 'roster');
  var rosterId = doc.rosterId;
  if (!rosterId || !findRow(rs, rosterId)) {
    var all = readAll(rs);
    var fix = null;
    for (var i = 0; i < all.length; i++) {
      if (String(all[i].claimedBy) === String(req.id)) { fix = all[i]; break; }
    }
    if (!fix) {
      for (var j = 0; j < all.length; j++) {
        if (sameName(all[j].name, doc.name) && !all[j].claimedBy) { fix = all[j]; break; }
      }
    }
    if (fix) {
      rosterId = fix.id;
      var fixRow = findRow(rs, rosterId);
      var fixDoc = readRow(rs, fixRow);
      if (!fixDoc.claimedBy) { fixDoc.claimedBy = String(req.id); writeRow(rs, fixRow, rosterId, fixDoc); }
      doc.rosterId = rosterId;
      writeRow(sh, row, req.id, doc);
    }
  }
  return { token: createSession(book, 'student', req.id, rosterId), name: doc.name, rosterId: rosterId };
}

/* ---- 가입 ---- */

function doRegisterStudent(book, req) {
  if (!req.id || !req.name) fail('MISSING_FIELDS');
  if (!req.password || String(req.password).length < 6) fail('WEAK_PASSWORD');

  return withLock(function () {
    var sh = getSheet(book, 'students');
    if (findRow(sh, req.id)) fail('DUPLICATE_ID');

    var rs = getSheet(book, 'roster');
    var all = readAll(rs);

    // 이 학번이 이미 잡아둔 기록이 있으면 그것을 쓴다(단추를 여러 번 눌러도 하나만 생기도록)
    var match = null;
    for (var i = 0; i < all.length; i++) {
      if (String(all[i].claimedBy) === String(req.id)) { match = all[i]; break; }
    }
    if (!match) {
      for (var j = 0; j < all.length; j++) {
        if (sameName(all[j].name, req.name) && !all[j].claimedBy) { match = all[j]; break; }
      }
    }

    var rosterId;
    if (match) {
      rosterId = match.id;
      var row = findRow(rs, rosterId);
      var cur = readRow(rs, row);
      cur.claimedBy = String(req.id);
      cur.major = req.dept || cur.major || '';
      cur.year = Number(req.year) || cur.year || 1;
      cur.updatedAt = nowIso();
      writeRow(rs, row, rosterId, cur);
    } else {
      rosterId = newRosterId(req.id);
      writeRow(rs, 0, rosterId, {
        name: req.name, major: req.dept || '', year: Number(req.year) || 1, role: '',
        claimedBy: String(req.id),
        r1: false, r2: false, r3: false, certA: false, certT: false, certS: false,
        total: 0, F: 0, A: 0, C: 0, E: 0, note: '', updatedAt: nowIso()
      });
    }

    var hashed = hashNewPassword(req.password);
    writeRow(sh, 0, req.id, {
      name: req.name, rosterId: rosterId,
      passwordSalt: hashed.passwordSalt, passwordHash: hashed.passwordHash,
      createdAt: nowIso()
    });
    return { token: createSession(book, 'student', req.id, rosterId), name: req.name, rosterId: rosterId };
  });
}

function doRegisterAdmin(book, req) {
  if (!req.id || !req.name) fail('MISSING_FIELDS');
  if (!req.password || String(req.password).length < 6) fail('WEAK_PASSWORD');

  return withLock(function () {
    var sh = getSheet(book, 'admins');
    var admins = readAll(sh);
    if (admins.length > 0) {
      var cfgSh = getSheet(book, 'config');
      var cfg = readRow(cfgSh, findRow(cfgSh, 'app')) || {};
      if (!cfg.adminCode) fail('ADMIN_CODE_LOCKED');
      if (req.code !== cfg.adminCode) fail('ADMIN_CODE_INVALID');
    }
    if (findRow(sh, req.id)) fail('DUPLICATE_ID');
    var hashed = hashNewPassword(req.password);
    writeRow(sh, 0, req.id, {
      name: req.name,
      passwordSalt: hashed.passwordSalt, passwordHash: hashed.passwordHash,
      createdAt: nowIso()
    });
    return { token: createSession(book, 'admin', req.id, ''), name: req.name };
  });
}

/* ---- 내 정보·비밀번호 ---- */

function doSaveProfile(book, req) {
  var session = requireSession(book, req.token);
  var isAdmin = session.role === 'admin';
  var sh = getSheet(book, isAdmin ? 'admins' : 'students');

  return withLock(function () {
    var row = findRow(sh, session.id);
    var doc = readRow(sh, row);
    if (!doc) fail('NOT_FOUND');
    if (!verifyPassword(req.curPassword, doc)) fail('INVALID_CREDENTIALS');

    if (req.newPassword) {
      if (String(req.newPassword).length < 6) fail('WEAK_PASSWORD');
      var hashed = hashNewPassword(req.newPassword);
      doc.passwordSalt = hashed.passwordSalt;
      doc.passwordHash = hashed.passwordHash;
    }
    if (req.name) doc.name = req.name;
    writeRow(sh, row, session.id, doc);

    if (!isAdmin && session.rosterId) {
      var rs = getSheet(book, 'roster');
      var rRow = findRow(rs, session.rosterId);
      var rDoc = readRow(rs, rRow);
      if (rDoc) {
        if (req.name) rDoc.name = req.name;
        if (req.major) rDoc.major = req.major;
        if (req.year) rDoc.year = Number(req.year);
        rDoc.updatedAt = nowIso();
        writeRow(rs, rRow, session.rosterId, rDoc);
      }
    }
    return { ok: true };
  });
}

/* ---- 관리자 전용 ---- */

function doDeleteMember(book, req) {
  requireAdmin(book, req.token);
  return withLock(function () {
    var rs = getSheet(book, 'roster');
    var row = findRow(rs, req.id);
    if (!row) return { ok: true };
    var doc = readRow(rs, row);
    if (doc && doc.claimedBy) {
      var ss = getSheet(book, 'students');
      var sRow = findRow(ss, doc.claimedBy);
      if (sRow) ss.deleteRow(sRow);
    }
    rs.deleteRow(row);
    return { ok: true };
  });
}

function doResetStudent(book, req) {
  requireAdmin(book, req.token);
  return withLock(function () {
    var rs = getSheet(book, 'roster');
    var row = findRow(rs, req.id);
    if (!row) fail('NOT_FOUND');
    var doc = readRow(rs, row);
    if (doc && doc.claimedBy) {
      var ss = getSheet(book, 'students');
      var sRow = findRow(ss, doc.claimedBy);
      if (sRow) ss.deleteRow(sRow);
      doc.claimedBy = null;
      doc.updatedAt = nowIso();
      writeRow(rs, row, req.id, doc);
    }
    return { ok: true };
  });
}

/* ---- 명단·설정 읽고 쓰기 ---- */

function doGetDoc(book, col, req) {
  var session = requireSession(book, req.token);
  var sh = getSheet(book, col);
  var doc = readRow(sh, findRow(sh, req.id));

  if (col === 'config') {
    if (session.role !== 'admin') fail('FORBIDDEN');   // 관리자 등록코드가 들어 있다
    if (!doc) return { doc: null };
    doc.id = String(req.id);
    return { doc: doc };
  }
  if (!doc) return { doc: null };
  if (!ownRoster(session, doc, req.id)) fail('FORBIDDEN');
  doc.id = String(req.id);
  return { doc: doc };
}

function doAllDocs(book, col, req) {
  requireAdmin(book, req.token);          // 전체 명단은 관리자만
  return { docs: readAll(getSheet(book, col)) };
}

function doWriteDoc(book, col, req, action) {
  var session = requireSession(book, req.token);

  if (col === 'config') {
    if (session.role !== 'admin') fail('FORBIDDEN');
  }

  return withLock(function () {
    var sh = getSheet(book, col);
    var row = findRow(sh, req.id);
    var cur = readRow(sh, row);

    if (col === 'roster') {
      if (!row) {
        // 명단에 없는 사람을 새로 만드는 것은 관리자만(학생 가입은 registerStudent 로)
        if (session.role !== 'admin') fail('FORBIDDEN');
      } else if (!ownRoster(session, cur, req.id)) {
        fail('FORBIDDEN');
      }
    }

    var data = {};
    var src = req.data || {};
    for (var k in src) data[k] = src[k];
    // 이 통로로는 못 바꾸게 막는다 — 로그인 연결은 가입·초기화에서만 다룬다
    delete data.claimedBy;
    delete data.passwordHash;
    delete data.passwordSalt;

    if (action === 'update' && row && cur) {
      for (var k2 in cur) if (!(k2 in data) && k2 !== 'id') data[k2] = cur[k2];
    }
    // 통째로 덮어쓸 때도 로그인 연결은 잃지 않도록 한다
    if (cur && cur.claimedBy && !data.claimedBy) data.claimedBy = cur.claimedBy;
    writeRow(sh, row, req.id, data);
    return { ok: true };
  });
}

/* ══════════════════════════════════════════════════
 *  캡쳐 이미지 보관 (analyze)
 *
 *  점수 읽기는 학생 휴대폰 안에서 끝난다. 여기서는 원본만 받아
 *  구글 드라이브 "K-STAR 캡쳐" 폴더에 보관한다(관리자 확인용).
 * ══════════════════════════════════════════════════ */

var SHOT_FOLDER = 'K-STAR 캡쳐';

function getShotFolder() {
  var it = DriveApp.getFoldersByName(SHOT_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(SHOT_FOLDER);
}

function analyzeShot(req) {
  var m = String(req.dataUrl || '').match(/^data:([^;]+);base64,(.*)$/);
  if (!m) fail('BAD_IMAGE');
  var who = String(req.name || 'capture').replace(/[\\\/:*?"<>|]/g, '');
  var stamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd_HHmmss');
  var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], who + '_' + stamp + '.jpg');
  var file = getShotFolder().createFile(blob);
  return { fileId: file.getId(), fileUrl: file.getUrl() };
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
  // 브라우저에서 주소를 직접 열었을 때 살아있는지 확인용 (자료는 아무것도 내주지 않는다)
  return ContentService.createTextOutput(JSON.stringify({
    ok: true,
    message: 'K-STAR 데이터 서버가 정상 동작 중입니다.'
  })).setMimeType(ContentService.MimeType.JSON);
}
