/**
 * 경기지역본부 이력서 선정 접수 Web App (Google Apps Script)
 * 첫 실행 시 setup()을 실행하여 접수 저장용 Google Sheet 생성.
 * 배포: 웹 앱 / 실행 사용자 본인 / 액세스 권한 모든 사용자.
 * 주의: 이름+지부만으로는 본인인증이 불가능하므로 시험 운영용.
 */
const REPO_RAW = 'https://raw.githubusercontent.com/hktower226-lang/waiting-order-dashboard/main/';
const SHEET_PROP = 'SELECTION_SHEET_ID';
const AUTH_PROP = 'AUTH_SHEET_ID'; // 비공개 참석자 인증명단 스프레드시트 ID (스크립트 속성에서만 설정)
const DEFAULT_NOTICE = '2026-10-29-umi36bl';

function setup() {
  const props = PropertiesService.getScriptProperties();
  const oldId = props.getProperty(SHEET_PROP);
  if (oldId) return '접수 저장 시트: https://docs.google.com/spreadsheets/d/' + oldId + '/edit';
  const book = SpreadsheetApp.create('경기지역본부 이력서 선정 희망자 접수');
  const sheet = book.getSheets()[0];
  sheet.setName('희망접수');
  sheet.appendRow(['접수시각', '공고ID', '성명', '지부', '접수당시 대기순번', '상태']);
  sheet.setFrozenRows(1);
  props.setProperty(SHEET_PROP, book.getId());
  return '설정 완료: ' + book.getUrl();
}

function doGet(e) {
  const supplied = (e && e.parameter && e.parameter.id) || DEFAULT_NOTICE;
  const id = assertId_(supplied);
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.noticeId = id;
  return tpl.evaluate()
    .setTitle('경기지역본부 이력서 선정')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function assertId_(id) {
  if (typeof id !== 'string' || !/^[a-zA-Z0-9_-]{5,100}$/.test(id)) {
    throw new Error('올바르지 않은 공고 주소입니다.');
  }
  return id;
}
function getJson_(path) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'v1:' + path;
  const hit = cache.get(cacheKey);
  if (hit) return JSON.parse(hit);
  const response = UrlFetchApp.fetch(REPO_RAW + path, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) throw new Error('공고 또는 명단을 읽을 수 없습니다.');
  const obj = JSON.parse(response.getContentText('UTF-8'));
  cache.put(cacheKey, JSON.stringify(obj), 30);
  return obj;
}
function notice_(id) {
  return getJson_('selection-notices/' + assertId_(id) + '.json');
}
function roster_() {
  return getJson_('data.json');
}
function sheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(SHEET_PROP);
  if (!id) throw new Error('관리자가 접수 저장 시트를 아직 초기화하지 않았습니다.');
  return SpreadsheetApp.openById(id).getSheetByName('희망접수');
}
function normalizeBranch_(v) {
  return String(v || '').replace(/\s+/g, '').replace(/지부$/, '');
}
function member_(name, branch) {
  const matches = roster_().filter(r =>
    String(r.name || '').trim() === name &&
    normalizeBranch_(r.branch) === normalizeBranch_(branch)
  );
  if (matches.length !== 1) return null;
  return matches[0];
}
/**
 * 참석자명부_인증용_비공개 > 인증명단: A 성명 / B 소속지부 / C 생년월일
 * 인증명단은 비공개 구글 시트로 유지하고, 개인정보는 브라우저·GitHub·접수기록에 내려주지 않는다.
 */
function birth6_(value) {
  const digits = String(value == null ? '' : value).replace(/\D/g, '');
  if (digits.length === 8) return digits.slice(2);
  if (digits.length === 6) return digits;
  return '';
}
function authenticationSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(AUTH_PROP);
  if (!id) throw new Error('관리자가 비공개 조합원 인증명단을 아직 연결하지 않았습니다.');
  const sheet = SpreadsheetApp.openById(id).getSheetByName('인증명단');
  if (!sheet) throw new Error('비공개 인증명단의 시트 이름을 확인해주세요.');
  const headers = sheet.getRange(1, 1, 1, 3).getDisplayValues()[0];
  if (headers[0] !== '성명' || headers[1] !== '소속지부' || headers[2] !== '생년월일') {
    throw new Error('비공개 인증명단의 A~C열 제목을 확인해주세요.');
  }
  return sheet;
}
function birthMatches_(name, branch, birth6) {
  const sheet = authenticationSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const matches = sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues()
    .filter(row => String(row[0]).trim() === name &&
      normalizeBranch_(row[1]) === normalizeBranch_(branch));
  // 같은 이름·지부에 서로 다른 생년월일이 있으면 관리자 확인 필요
  const birthdays = [...new Set(matches.map(row => birth6_(row[2])).filter(Boolean))];
  return birthdays.length === 1 && birthdays[0] === birth6;
}
function failedAttemptKey_(name,branch) {
  const salt = PropertiesService.getScriptProperties().getProperty('AUTH_SALT') || 'selection-rate-v1';
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    salt + '|' + normalizeBranch_(branch) + '|' + name, Utilities.Charset.UTF_8);
  return 'authfail:' + Utilities.base64EncodeWebSafe(digest).slice(0,40);
}
function requireAttemptAllowed_(key) {
  if (Number(CacheService.getScriptCache().get(key) || 0) >= 5) {
    throw new Error('인증 시도 횟수를 초과했습니다. 30분 후 다시 시도하거나 관리자에게 문의해주세요.');
  }
}
function failedAttempt_(key) {
  const cache = CacheService.getScriptCache();
  const count = Number(cache.get(key) || 0);
  cache.put(key, String(count + 1), 1800);
}
function successfulAttempt_(key) {
  CacheService.getScriptCache().remove(key);
}
function getState(id) {
  const n = notice_(id);
  const rows = sheet_().getDataRange().getValues().slice(1)
    .filter(r => String(r[1]) === id && String(r[5]) === '접수');
  const roster = roster_();
  const used = new Set();
  const people = [];
  rows.forEach(r => {
    const name = String(r[2]).trim(), branch = String(r[3]).trim();
    const key = branch + '|' + name;
    if (used.has(key)) return;
    used.add(key);
    const current = roster.find(x => String(x.name).trim() === name &&
      String(x.branch).trim() === branch);
    const rank = current && Number.isFinite(Number(current.seq)) ?
      Number(current.seq) : Number(r[4]);
    people.push({name:name, branch:branch, seq:rank, time:new Date(r[0]).toISOString()});
  });
  people.sort((a,b) => a.seq-b.seq || a.time.localeCompare(b.time));
  const now = Date.now(), deadline = new Date(n.deadline).getTime();
  const isOpen = n.status === 'open' && Number.isFinite(deadline) && now < deadline;
  return {
    notice: {
      id:n.id, project:n.project, site:n.site, rental:n.rental,
      wage:n.wage, floors:n.floors, machine:n.machine,
      plannedInstallation:n.plannedInstallation, startCondition:n.startCondition,
      duration:n.duration, notes:n.notes || [], deadline:n.deadline
    },
    accepting: isOpen, configured: n.status === 'open',
    now: new Date(now).toISOString(), participants:people,
    total:people.length
  };
}
function submitApplicant(id, name, branch, birth6, consent) {
  id = assertId_(id);
  name = String(name || '').trim();
  branch = normalizeBranch_(branch);
  birth6 = String(birth6 || '').trim();
  if (!/^[0-9]{6}$/.test(birth6)) {
    throw new Error('생년월일 6자리(YYMMDD)를 정확히 입력해주세요.');
  }
  if (!consent) throw new Error('본인 신청 확인에 동의해주세요.');
  if (!name || !branch || name.length > 40 || branch.length > 30) {
    throw new Error('성명과 소속 지부를 정확히 입력해주세요.');
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('접수가 많습니다. 다시 시도해주세요.');
  try {
    const n = notice_(id);
    if (n.status !== 'open' || Date.now() >= new Date(n.deadline).getTime()) {
      throw new Error('접수 기간이 아니거나 마감되었습니다.');
    }
    const key = failedAttemptKey_(name,branch);
    requireAttemptAllowed_(key);
    const person = member_(name,branch);
    if (!person || !birthMatches_(name,branch,birth6)) {
      failedAttempt_(key);
      throw new Error('조합원 인증정보가 일치하지 않습니다. 성명·지부·생년월일을 확인해주세요.');
    }
    successfulAttempt_(key);
    const sheet = sheet_();
    const rows = sheet.getDataRange().getValues().slice(1);
    const duplicate = rows.some(r => String(r[1]) === id &&
      String(r[2]).trim() === name && String(r[3]).trim() === branch &&
      String(r[5]) === '접수');
    if (duplicate) throw new Error('이미 접수된 이름입니다.');
    if (Date.now() >= new Date(n.deadline).getTime()) {
      throw new Error('접수 마감시간이 지났습니다.');
    }
    sheet.appendRow([new Date(), id, name, branch, Number(person.seq), '접수']);
    return {ok:true, message:'희망 신청이 접수되었습니다.'};
  } finally {
    lock.releaseLock();
  }
}
