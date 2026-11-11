/**
 * 우리집 가계부 — Google Apps Script 서버
 *
 * 이 스크립트가 붙어 있는 구글 시트에 가계부를 저장해요.
 *   users   시트: 계정 (이름, PIN 암호값, 잠금 상태, 기억한 분류)
 *   entries 시트: 모든 사람의 내역 (user 열로 구분)
 * PIN은 그대로 저장하지 않고 되돌릴 수 없는 암호값(SHA-256)으로만 저장해요.
 */

const MAX_USERS = 4;          // 만들 수 있는 계정 수 (모르는 사람이 계정을 계속 만들지 못하게)
const SESSION_HOURS = 6;      // 로그인 유지 시간 (최대 6시간)
const MAX_FAILS = 5;          // PIN을 이만큼 틀리면 잠금

const USERS = 'users';
const ENTRIES = 'entries';
const USER_COLS = ['name', 'salt', 'pinHash', 'failCount', 'lockUntil', 'lockCount', 'rules', 'createdAt'];
const ENTRY_COLS = ['user', 'id', 'date', 'type', 'category', 'amount', 'memo', 'source', 'origAmount', 'batch'];
const TYPES = ['income', 'expense', 'saving'];

// ---------- 화면 ----------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('우리집 가계부')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

// ---------- 시트 읽고 쓰기 ----------
function sheet_(name, cols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAll_(sh, cols) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, cols.length).getValues();
}

// 시트 내용을 rows로 통째로 바꿔요 (잠금 안에서만 호출)
function writeAll_(sh, cols, rows, numberCols) {
  const old = Math.max(sh.getLastRow() - 1, 0);
  const need = rows.length + 1;
  if (need > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  if (rows.length) {
    const range = sh.getRange(2, 1, rows.length, cols.length);
    range.setNumberFormat('@'); // 날짜·숫자처럼 보이는 글자가 바뀌지 않게
    (numberCols || []).forEach(c => sh.getRange(2, c + 1, rows.length, 1).setNumberFormat('#,##0'));
    range.setValues(rows);
  }
  if (old > rows.length) sh.getRange(rows.length + 2, 1, old - rows.length, cols.length).clearContent();
}

// 수식으로 해석되지 않게
function text_(v, max) {
  let s = String(v == null ? '' : v).slice(0, max || 200);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function dateText_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '');
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

const norm_ = s => String(s || '').trim().toLowerCase();

function hash_(salt, pin) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pin, Utilities.Charset.UTF_8);
  return bytes.map(b => ((b + 256) % 256).toString(16).padStart(2, '0')).join('');
}

function checkInput_(name, pin) {
  name = String(name || '').trim();
  if (!name || name.length > 20) throw new Error('이름은 1~20자로 입력해 주세요.');
  if (!/^\d{4}$/.test(String(pin || ''))) throw new Error('PIN은 숫자 4자리예요.');
  return name;
}

// ---------- 로그인 ----------
function session_(name) {
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('s:' + token, name, SESSION_HOURS * 3600);
  return { token: token, name: name };
}

function auth_(token) {
  const name = token && CacheService.getScriptCache().get('s:' + token);
  if (!name) throw new Error('AUTH');
  return name;
}

function signup(name, pin) {
  name = checkInput_(name, pin);
  return withLock_(() => {
    const sh = sheet_(USERS, USER_COLS);
    const rows = readAll_(sh, USER_COLS);
    const i = rows.findIndex(r => norm_(r[0]) === norm_(name));
    // PIN 칸을 비워 둔 계정은 같은 이름으로 새 PIN을 정할 수 있어요 (PIN을 잊었을 때)
    if (i >= 0 && rows[i][2]) throw new Error('이미 있는 이름이에요. 로그인하거나 다른 이름을 써 주세요.');
    if (i < 0 && rows.length >= MAX_USERS) throw new Error('더 이상 계정을 만들 수 없어요.');
    const salt = Utilities.getUuid();
    const row = [text_(name, 20), salt, hash_(salt, pin), '0', '', '0',
      i >= 0 ? rows[i][6] : '{}', i >= 0 ? rows[i][7] : new Date().toISOString()];
    if (i >= 0) { row[0] = rows[i][0]; rows[i] = row; } else rows.push(row);
    writeAll_(sh, USER_COLS, rows);
    return session_(String(row[0]));
  });
}

function login(name, pin) {
  name = checkInput_(name, pin);
  // 잠금 안에서 기록하고, 틀린 경우 잠금 밖에서 알려 줘요
  const result = withLock_(() => {
    const sh = sheet_(USERS, USER_COLS);
    const rows = readAll_(sh, USER_COLS);
    const i = rows.findIndex(r => norm_(r[0]) === norm_(name));
    if (i < 0) return { error: '이름 또는 PIN이 맞지 않아요.' };
    const r = rows[i];
    if (!r[2]) return { error: 'PIN을 새로 정해야 해요. "처음이에요"를 눌러 같은 이름으로 PIN을 정해 주세요.' };
    const now = Date.now();
    const lockUntil = Number(r[4]) || 0;
    if (now < lockUntil) {
      return { error: `PIN을 여러 번 틀려서 잠겨 있어요. ${Math.ceil((lockUntil - now) / 60000)}분 후에 다시 시도해 주세요.` };
    }
    if (hash_(r[1], pin) === r[2]) {
      r[3] = '0'; r[4] = ''; r[5] = '0';
      writeAll_(sh, USER_COLS, rows);
      return { ok: String(r[0]) };
    }
    const fails = (Number(r[3]) || 0) + 1;
    let error;
    if (fails >= MAX_FAILS) {
      // 잠글 때마다 잠금 시간이 두 배로 (10분 → 20분 → 40분 … 최대 하루)
      const lockCount = (Number(r[5]) || 0) + 1;
      const minutes = Math.min(10 * Math.pow(2, lockCount - 1), 1440);
      r[3] = '0'; r[4] = String(now + minutes * 60000); r[5] = String(lockCount);
      error = `PIN을 ${MAX_FAILS}번 틀려서 ${minutes}분 동안 잠겼어요.`;
    } else {
      r[3] = String(fails);
      error = `PIN이 맞지 않아요. (${fails}/${MAX_FAILS})`;
    }
    writeAll_(sh, USER_COLS, rows);
    return { error: error };
  });
  if (result.error) throw new Error(result.error);
  return session_(result.ok);
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('s:' + token);
  return true;
}

// ---------- 가계부 데이터 ----------
function getData(token) {
  const name = auth_(token);
  const users = readAll_(sheet_(USERS, USER_COLS), USER_COLS);
  const u = users.find(r => norm_(r[0]) === norm_(name));
  let rules = {};
  try { rules = JSON.parse((u && u[6]) || '{}') || {}; } catch (e) { rules = {}; }

  const entries = readAll_(sheet_(ENTRIES, ENTRY_COLS), ENTRY_COLS)
    .filter(r => norm_(r[0]) === norm_(name))
    .map(r => {
      const e = {
        id: String(r[1]), date: dateText_(r[2]), type: String(r[3]), category: String(r[4]),
        amount: Number(r[5]), memo: String(r[6] || ''),
      };
      if (r[7]) e.source = String(r[7]);
      if (Number(r[8]) > 0) e.origAmount = Number(r[8]);
      if (r[9]) e.batch = String(r[9]);
      return e;
    });
  return { name: name, entries: entries, rules: rules };
}

function entryRow_(name, e) {
  if (!e || TYPES.indexOf(e.type) < 0) throw new Error('잘못된 내역이 있어요.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(e.date))) throw new Error('잘못된 날짜가 있어요.');
  const amount = Math.round(Number(e.amount));
  if (!(amount > 0)) throw new Error('잘못된 금액이 있어요.');
  return [
    text_(name, 20), text_(e.id, 40), String(e.date), e.type, text_(e.category, 20), amount,
    text_(e.memo, 200), text_(e.source || '', 10), Number(e.origAmount) > 0 ? Math.round(Number(e.origAmount)) : '',
    text_(e.batch || '', 40),
  ];
}

// upserts: 새로 생기거나 바뀐 내역, deleteIds: 지운 내역의 id
function saveEntries(token, upserts, deleteIds) {
  const name = auth_(token);
  const newRows = (upserts || []).map(e => entryRow_(name, e));
  const del = new Set((deleteIds || []).map(String));
  return withLock_(() => {
    const sh = sheet_(ENTRIES, ENTRY_COLS);
    const rows = readAll_(sh, ENTRY_COLS);
    const mine = new Map();
    rows.forEach((r, i) => { if (norm_(r[0]) === norm_(name)) mine.set(String(r[1]), i); });
    for (const row of newRows) {
      const i = mine.get(String(row[1]));
      if (i === undefined) { mine.set(String(row[1]), rows.length); rows.push(row); }
      else rows[i] = row;
    }
    const kept = rows.filter(r => !(norm_(r[0]) === norm_(name) && del.has(String(r[1]))));
    // 날짜순으로 정리해 두면 시트에서 보기 편해요
    kept.sort((a, b) => String(a[2]).localeCompare(String(b[2])) || norm_(a[0]).localeCompare(norm_(b[0])));
    writeAll_(sh, ENTRY_COLS, kept, [5, 8]);
    return true;
  });
}

function saveRules(token, rules) {
  const name = auth_(token);
  const json = JSON.stringify(rules && typeof rules === 'object' ? rules : {});
  if (json.length > 45000) throw new Error('기억한 분류가 너무 많아요. 필요 없는 것을 지워 주세요.');
  return withLock_(() => {
    const sh = sheet_(USERS, USER_COLS);
    const rows = readAll_(sh, USER_COLS);
    const i = rows.findIndex(r => norm_(r[0]) === norm_(name));
    if (i < 0) throw new Error('AUTH');
    rows[i][6] = json;
    writeAll_(sh, USER_COLS, rows);
    return true;
  });
}
