/**
 * SYMC 마케터 활동일지 - 구글 시트 백엔드
 *
 * [배포] 손으로 하지 않는다.
 *   이 파일을 고쳐 main 브랜치에 push 하면 GitHub Actions 가
 *   (.github/workflows/deploy-apps-script.yml)
 *     1) clasp push            — 이 코드를 Apps Script 로 올리고
 *     2) clasp update-deployment — 기존 웹 앱 배포를 새 버전으로 바꾼다(주소 그대로)
 *     3) 상태 확인              — 새 코드가 실제로 도는지 build 값으로 확인
 *   시트 구조(열·드롭다운)는 배포 후 첫 요청에서 자동으로 맞춰진다 → SCHEMA_VERSION 참고.
 *   ※ Apps Script 편집기에서 직접 고치지 말 것. 다음 배포 때 이 파일로 덮어써진다.
 *
 * [시트] 활동일지 / 단원명단 / 설정
 *   단원명단: 이름 | 코드 | 역할 | 학령 | 성별   (역할·학령·성별은 설정 시트의 드롭다운)
 *   정신우 | 1234 | 선생님 |      |        <- 전체 열람
 *   오수민 | 5821 | 단원   | 고   | 여     <- 본인 것만
 *
 * [처음부터 새로 만들 때만] 스프레드시트 > 확장 프로그램 > Apps Script 에 붙여넣고 setup 실행,
 *   배포 > 새 배포 > 웹 앱(실행 자격: 나 / 액세스: 모든 사용자). 나온 주소를 index.html 의
 *   API_ENDPOINT 에 넣는다. 그 뒤로는 위의 자동 배포를 쓴다.
 */

const SHEET_DATA    = '활동일지';
const SHEET_MEMBERS = '단원명단';
const SHEET_CONFIG  = '설정';

/* 예전 영문 시트명. 시트를 찾을 때 예전 이름이 있으면 이름만 바꿔 쓴다(데이터 보존) */
const OLD_SHEET_NAMES = { '활동일지': 'activities', '단원명단': 'members' };

const HEADERS = ['id','member','kind','date','start','end','place','types','offline','etc',
                 'topic','duty','make','up','part','content','evaluation','next','savedAt'];

/* 단원명단 열 구성. 학령/성별은 사람의 속성이라 활동일지가 아니라 여기에 둔다. */
const MEMBER_HEADERS = ['이름','코드','역할','학령','성별'];

/* 설정 시트 열 구성. 여기 적힌 값이 단원명단의 드롭다운 목록이 된다.
   값을 더 적으면 드롭다운에 바로 반영된다(코드를 다시 실행할 필요 없음). */
const CONFIG_COLUMNS = [
  { title: '역할', values: ['선생님','단원'] },
  { title: '학령', values: ['초','중','고','대','학교밖'] },
  { title: '성별', values: ['남','여'] }
];
/* 단원명단에서 각 설정 열이 걸릴 위치 (1부터). 역할=C, 학령=D, 성별=E */
const VALIDATION_TARGET = { '역할': 3, '학령': 4, '성별': 5 };

/* 역할은 시트에 한글로 적고, 코드 안에서는 admin/student 로 다룬다.
   예전에 적어둔 영문 값(admin / student)도 그대로 알아듣는다. */
const ROLE_ADMIN_LABEL   = '선생님';
const ROLE_STUDENT_LABEL = '단원';

function normRole(v) {
  const t = String(v || '').trim();
  if (t === ROLE_ADMIN_LABEL || t === '관리자' || t === '교사') return 'admin';
  if (t.toLowerCase() === 'admin' || t.toLowerCase() === 'administrator') return 'admin';
  return 'student';
}

/**
 * 설정 시트의 역할 목록이 예전 영문 기본값 그대로면 한글 표기로 바꿔준다.
 * 선생님이 직접 손본 흔적이 있으면 건드리지 않는다.
 */
function upgradeRoleOptions() {
  const cfg = getConfigSheet();
  const n = Math.max(cfg.getLastRow() - 1, 1);
  const cur = cfg.getRange(2, 1, n, 1).getValues()
                 .map(function (r) { return String(r[0] || '').trim(); })
                 .filter(function (v) { return v; });
  const isOldDefault = cur.length === 2 &&
                       cur.indexOf('admin') >= 0 && cur.indexOf('student') >= 0;
  if (!isOldDefault) return;
  cfg.getRange(2, 1, n, 1).clearContent();
  cfg.getRange(2, 1, 2, 1).setValues([[ROLE_ADMIN_LABEL], [ROLE_STUDENT_LABEL]]);
}

/** 단원명단 역할 칸에 남아 있는 영문 값을 한글 표기로 맞춘다 */
function normalizeRoleColumn() {
  const sheet = getMemberSheet();
  const n = sheet.getLastRow() - 1;
  if (n < 1) return;
  const range = sheet.getRange(2, VALIDATION_TARGET['역할'], n, 1);
  const vals = range.getValues();
  let changed = false;
  for (let i = 0; i < vals.length; i++) {
    const cur = String(vals[i][0] || '').trim();
    if (!cur) continue;
    const want = normRole(cur) === 'admin' ? ROLE_ADMIN_LABEL : ROLE_STUDENT_LABEL;
    if (cur !== want) { vals[i][0] = want; changed = true; }
  }
  if (changed) range.setValues(vals);
}

/**
 * 시트 구조 버전. 시트의 열·드롭다운 구성을 바꿀 때마다 이 숫자를 1 올린다.
 * 배포 후 첫 요청에서 저장된 버전과 다르면 setupSheets() 를 한 번 자동으로 실행하므로,
 * 더 이상 Apps Script 편집기에서 setup 을 손으로 돌릴 필요가 없다.
 */
const SCHEMA_VERSION = 1;

/* 배포 표시. GitHub Actions 가 배포할 때 커밋 번호로 바꿔 넣는다(저장소에는 'local' 로 남는다). */
const BUILD = 'local';

/** 매 요청 맨 앞에서 부른다. 평소에는 속성 하나 읽고 끝나서 거의 비용이 없다. */
function ensureSchema() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('schemaVersion') === String(SCHEMA_VERSION)) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    /* 기다리는 사이 다른 요청이 먼저 끝냈을 수 있다 */
    if (props.getProperty('schemaVersion') === String(SCHEMA_VERSION)) return;
    setupSheets();
    props.setProperty('schemaVersion', String(SCHEMA_VERSION));
  } finally {
    lock.releaseLock();
  }
}

/**
 * 시트를 만들거나 최신 상태로 맞춘다. 몇 번을 실행해도 안전하다.
 * 지우는 동작은 하나도 없다.
 */
function setupSheets() {
  getDataSheet();
  ensureMemberColumns();
  getConfigSheet();
  upgradeRoleOptions();
  normalizeRoleColumn();
  applyValidation();
}

/** 편집기에서 손으로 돌리고 싶을 때용. 자동 갱신과 같은 일을 하고 알림을 띄운다. */
function setup() {
  setupSheets();
  PropertiesService.getScriptProperties().setProperty('schemaVersion', String(SCHEMA_VERSION));
  SpreadsheetApp.getActiveSpreadsheet()
    .toast('설정 완료. 단원명단 시트의 역할·학령·성별 칸이 드롭다운으로 바뀌었습니다.');
}

/** 이름으로 시트를 찾되, 예전 영문 이름만 남아 있으면 새 이름으로 바꿔서 돌려준다 */
function findSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  const old = OLD_SHEET_NAMES[name] && ss.getSheetByName(OLD_SHEET_NAMES[name]);
  if (old) { old.setName(name); return old; }   // 이름만 교체. 내용은 건드리지 않는다.
  return null;
}

function getDataSheet() {
  let sheet = findSheet(SHEET_DATA);
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_DATA);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getMemberSheet() {
  let sheet = findSheet(SHEET_MEMBERS);
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_MEMBERS);
    sheet.appendRow(MEMBER_HEADERS);
    sheet.appendRow(['선생님','1234','admin','','']);
    sheet.setFrozenRows(1);
    sheet.getRange('B:B').setNumberFormat('@');   // 0으로 시작하는 코드 보호
  }
  return sheet;
}

/**
 * 이미 쓰던 시트에 모자란 머리글을 채우고 서식을 맞춘다.
 * setup 에서만 부른다 — 로그인마다 시트에 쓰면 느려지고 잠금이 걸리기 때문.
 */
function ensureMemberColumns() {
  const sheet = getMemberSheet();
  const head = sheet.getRange(1, 1, 1, MEMBER_HEADERS.length).getValues()[0];
  for (let i = 0; i < MEMBER_HEADERS.length; i++) {
    if (String(head[i] || '').trim() !== MEMBER_HEADERS[i]) {
      sheet.getRange(1, i + 1).setValue(MEMBER_HEADERS[i]);
    }
  }
  sheet.setFrozenRows(1);
  sheet.getRange('B:B').setNumberFormat('@');
  return sheet;
}

/** 드롭다운에 쓸 값을 모아두는 시트 */
function getConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CONFIG);
    CONFIG_COLUMNS.forEach(function (col, i) {
      sheet.getRange(1, i + 1).setValue(col.title);
      if (col.values.length) {
        sheet.getRange(2, i + 1, col.values.length, 1)
             .setValues(col.values.map(function (v) { return [v]; }));
      }
      sheet.setColumnWidth(i + 1, 110);
    });
    sheet.getRange(1, 1, 1, CONFIG_COLUMNS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange('E1').setValue('이 시트의 값이 단원명단의 드롭다운 목록이 됩니다.');
    sheet.getRange('E2').setValue('값을 더 적으면 드롭다운에 바로 반영됩니다. 코드를 다시 실행할 필요는 없습니다.');
    sheet.getRange('E3').setValue('단, 열을 새로 추가했다면 setup 을 한 번 더 실행하세요.');
    sheet.setColumnWidth(5, 420);
  }
  return sheet;
}

/**
 * 단원명단의 역할·학령·성별 칸에 드롭다운을 건다.
 * 목록을 '설정' 시트의 열 범위로 참조하므로, 설정 값을 고치면 드롭다운도 같이 바뀐다.
 */
function applyValidation() {
  const cfg = getConfigSheet();
  const mem = getMemberSheet();
  const lastRow = Math.max(mem.getMaxRows() - 1, 1);

  CONFIG_COLUMNS.forEach(function (col, i) {
    const target = VALIDATION_TARGET[col.title];
    if (!target) return;
    const letter = String.fromCharCode(65 + i);          // A, B, C ...
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(cfg.getRange(letter + '2:' + letter), true)
      .setAllowInvalid(false)
      .setHelpText('설정 시트의 ' + col.title + ' 목록에서 골라주세요.')
      .build();
    mem.getRange(2, target, lastRow, 1).setDataValidation(rule);
  });
}
/**
 * 단원 명단. 선생님(admin)에게만 내려보낸다.
 * 코드(로그인 비밀번호) 열은 절대 포함하지 않는다 — 프론트는 공개 소스라서
 * 한 번이라도 내보내면 그대로 노출된다.
 */
function rosterFor(auth) {
  if (auth.role !== 'admin') return [];
  const rows = getMemberSheet().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const name = String(rows[i][0] || '').trim();
    if (!name) continue;
    out.push({
      name: name,
      role: normRole(rows[i][2]),
      grade: String(rows[i][3] || '').trim(),
      gender: String(rows[i][4] || '').trim()
    });
  }
  return out;
}

/** 이름+코드 확인. 실패 시 null */
function authenticate(name, code) {
  if (!name || !code) return null;
  const rows = getMemberSheet().getDataRange().getValues();
  const n = String(name).trim();
  const c = String(code).trim();
  for (let i = 1; i < rows.length; i++) {
    const rn = String(rows[i][0]).trim();
    const rc = String(rows[i][1]).trim();
    if (!rn) continue;
    if (rn === n && rc === c) {
      return { name: rn, role: normRole(rows[i][2]) };
    }
  }
  return null;
}

function doPost(e) {
  try {
    ensureSchema();      // 배포 직후 첫 요청이면 시트 구조를 맞춘다
    const body = JSON.parse(e.postData.contents);
    const auth = authenticate(body.name, body.code);
    if (!auth) return json({ ok: false, error: 'auth', message: '이름 또는 코드가 올바르지 않습니다.' });

    if (body.action === 'login' || body.action === 'list') {
      return json({ ok: true, role: auth.role, name: auth.name, entries: entriesFor(auth), members: rosterFor(auth) });
    }

    if (body.action === 'save') {
      const target = auth.role === 'admin' ? String(body.member || auth.name).trim() : auth.name;
      const sheet = getDataSheet();
      const lock = LockService.getScriptLock();
      lock.waitLock(20000);
      try {
        removeRowsOf(sheet, target);
        const entries = (body.entries || []).filter(function(en){
          return auth.role === 'admin' ? true : String(en.member).trim() === auth.name;
        });
        entries.forEach(function(en){
          en.member = target;
          sheet.appendRow(entryToRow(en));
        });
      } finally { lock.releaseLock(); }
      return json({ ok: true, role: auth.role, entries: entriesFor(auth), members: rosterFor(auth) });
    }

    return json({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json({ ok: false, error: 'server', message: String(err) });
  }
}

/**
 * 상태 확인용. GitHub Actions 가 배포 직후 이 주소를 불러
 *  (1) 새 코드가 실제로 돌고 있는지(build), (2) 시트 구조가 맞춰졌는지(schema) 확인한다.
 * 배포 직후 시트 구조 갱신도 여기서 미리 끝내 두어, 첫 학생 요청이 느려지지 않게 한다.
 */
function doGet(e) {
  try {
    ensureSchema();
    return json({ ok: true, message: 'SYMC 활동일지 서버가 동작 중입니다.', build: BUILD, schema: SCHEMA_VERSION });
  } catch (err) {
    return json({ ok: false, error: 'server', message: String(err), build: BUILD });
  }
}

function entriesFor(auth) {
  const rows = getDataSheet().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    const entry = rowToEntry(rows[i]);
    if (auth.role === 'admin' || String(entry.member).trim() === auth.name) out.push(entry);
  }
  return out;
}

function removeRowsOf(sheet, member) {
  const rows = sheet.getDataRange().getValues();
  const col = HEADERS.indexOf('member');
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][col]).trim() === String(member).trim()) sheet.deleteRow(i + 1);
  }
}

function rowToEntry(row) {
  function g(name) { return row[HEADERS.indexOf(name)]; }
  return {
    id: String(g('id')), member: String(g('member')), kind: String(g('kind')),
    date: fmtDate(g('date')), start: fmtTime(g('start')), end: fmtTime(g('end')),
    place: String(g('place')),
    types: parse(g('types'), []),
    offline: String(g('offline')), etc: String(g('etc')),
    topic: String(g('topic')), duty: String(g('duty')),
    make: parse(g('make'), {tiktok:0,insta:0,youtube:0,blog:0}),
    up: parse(g('up'), {tiktok:0,insta:0,youtube:0,blog:0}),
    part: parse(g('part'), {infant:0,teen:0,adult:0}),
    content: String(g('content')), evaluation: String(g('evaluation')),
    next: String(g('next')), savedAt: String(g('savedAt'))
  };
}

function entryToRow(en) {
  const map = {
    id: en.id, member: en.member, kind: en.kind, date: "'" + en.date,
    start: "'" + (en.start || ''), end: "'" + (en.end || ''), place: en.place || '',
    types: JSON.stringify(en.types || []), offline: en.offline || '', etc: en.etc || '',
    topic: en.topic || '', duty: en.duty || '',
    make: JSON.stringify(en.make || {}), up: JSON.stringify(en.up || {}), part: JSON.stringify(en.part || {}),
    content: en.content || '', evaluation: en.evaluation || '', next: en.next || '',
    savedAt: en.savedAt || new Date().toISOString()
  };
  return HEADERS.map(function(h){ return map[h] !== undefined ? map[h] : ''; });
}

function fmtDate(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '').replace(/^'/, '');
}
function fmtTime(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  return String(v || '').replace(/^'/, '');
}
function parse(str, fallback) {
  try { return JSON.parse(str); } catch (e) { return fallback; }
}
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
