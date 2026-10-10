/**
 * SYMC 마케터 활동일지 - 구글 시트 백엔드
 *
 * [최초 설정]
 * 1. 이 코드를 붙여넣고 저장합니다.
 * 2. 상단 함수 목록에서 setup 을 선택하고 실행(▶)합니다.
 *    -> activities / members 두 개의 시트가 자동으로 만들어집니다.
 * 3. members 시트에 팀원 정보를 입력합니다.
 *    이름   | 코드   | 역할
 *    정신우 | 1234   | admin      <- 선생님(전체 열람)
 *    오수민 | 5821   | student    <- 학생(본인 것만)
 * 4. 배포(Deploy) > 새 배포 > 웹 앱
 *    - 실행 자격: 나(Me)
 *    - 액세스 권한: 모든 사용자(Anyone)
 * 5. 나온 주소를 index.html 의 API_ENDPOINT 에 붙여넣습니다.
 *
 * [주의] 코드를 수정한 뒤에는 배포 > 배포 관리 > 수정 > 새 버전으로 다시 배포해야
 *        변경 내용이 반영됩니다.
 */

const SHEET_DATA = 'activities';
const SHEET_MEMBERS = 'members';
const HEADERS = ['id','member','kind','date','start','end','place','types','offline','etc',
                 'topic','duty','make','up','part','content','evaluation','next','savedAt'];

/* members 시트 열 구성. 학령/성별은 사람의 속성이라 활동일지가 아니라 여기에 둔다.
   학령 값: 초 / 중 / 고 / 대 / 학교밖  (학교를 안 다니는 청소년) */
const MEMBER_HEADERS = ['이름','코드','역할','학령','성별'];

function setup() {
  getDataSheet();
  const m = getMemberSheet();
  SpreadsheetApp.getActiveSpreadsheet().toast('설정 완료. members 시트에 팀원을 등록하세요.');
}

function getDataSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_DATA);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DATA);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getMemberSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_MEMBERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_MEMBERS);
    sheet.appendRow(MEMBER_HEADERS);
    sheet.appendRow(['선생님','1234','admin','','']);
    sheet.setFrozenRows(1);
    sheet.getRange('B:B').setNumberFormat('@');
  }
  return sheet;
}


/**
 * 이미 쓰고 있는 members 시트에 '학령'/'성별' 열을 붙인다.
 * 상단 함수 목록에서 한 번만 실행하면 된다. 기존 이름/코드/역할은 건드리지 않는다.
 */
function migrateMembers() {
  const sheet = getMemberSheet();
  const head = sheet.getRange(1, 1, 1, MEMBER_HEADERS.length).getValues()[0];
  for (let i = 0; i < MEMBER_HEADERS.length; i++) {
    if (String(head[i] || '').trim() !== MEMBER_HEADERS[i]) {
      sheet.getRange(1, i + 1).setValue(MEMBER_HEADERS[i]);
    }
  }
  sheet.setFrozenRows(1);
  sheet.getRange('B:B').setNumberFormat('@');
  SpreadsheetApp.getActiveSpreadsheet()
    .toast('members 시트에 학령/성별 열을 준비했습니다. 학령은 초/중/고/대/학교밖 중에서 적어주세요.');
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
    const role = String(rows[i][2] || 'student').trim().toLowerCase();
    out.push({
      name: name,
      role: role === 'admin' ? 'admin' : 'student',
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
      const role = String(rows[i][2] || 'student').trim().toLowerCase();
      return { name: rn, role: role === 'admin' ? 'admin' : 'student' };
    }
  }
  return null;
}

function doPost(e) {
  try {
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

function doGet(e) {
  return json({ ok: true, message: 'SYMC 활동일지 서버가 동작 중입니다.' });
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
