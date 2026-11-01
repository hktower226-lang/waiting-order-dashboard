/**
 * 경기지역본부 이력서 선정 접수 Web App (Google Apps Script)
 * 첫 실행 시 setup()을 실행하여 접수 저장용 Google Sheet 생성.
 * 배포: 웹 앱 / 실행 사용자 본인 / 액세스 권한 모든 사용자.
 * 주의: 성명 및 대기순번 선택만으로는 강한 본인인증이 불가능합니다.
 */
const REPO_RAW = 'https://raw.githubusercontent.com/hktower226-lang/waiting-order-dashboard/main/';
const SHEET_PROP = 'SELECTION_SHEET_ID';
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
  // 관리자 모드는 별도 주소에서만 열리고, 서버에서도 비밀번호를 검사한다.
  tpl.view = (e && e.parameter && e.parameter.view==='admin') ? 'admin' : 'apply';
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
  return String(v || '').replace(/\s+/g,'').replace(/지부$/,'');
}
function noticeType_(n) {
  return String(n.type||n.towerType||'large').toLowerCase()==='small'?'small':'large';
}
function memberKey_(p) {
  // 항상 이름과 지부까지 함께 검사하므로 순번 변동 시 타인으로 잘못 연결되지 않는다.
  return JSON.stringify([String(p.seq),String(p.name).trim(),
    normalizeBranch_(p.branch),String(p.endDate||'').trim()]);
}
function findMember_(key) {
  let tuple;
  try {tuple=JSON.parse(String(key||''));} catch(e) {return null;}
  if(!Array.isArray(tuple)||tuple.length!==4)return null;
  const found=roster_().filter(r=>memberKey_(r)===JSON.stringify(tuple));
  return found.length===1?found[0]:null;
}
function hangulInitials_(text) {
  const CHO='ㄱㄲㄴㄷㄸㄹㅁㅂㅃㅅㅆㅇㅈㅉㅊㅋㅌㅍㅎ';
  return [...String(text||'')].map(c=>{
    const n=c.charCodeAt(0)-0xAC00;
    return n>=0&&n<11172?CHO[Math.floor(n/588)]:c;
  }).join('');
}
function searchMembers(query, noticeId) {
  const notice=notice_(assertId_(noticeId));
  const q=String(query||'').trim().replace(/\s+/g,'');
  if(q.length<1||q.length>30)return [];
  const matches=roster_().filter(r=>{
    const name=String(r.name||'').replace(/\s+/g,'');
    return name.startsWith(q)||hangulInitials_(name).startsWith(q);
  });
  return matches.map(r=>{
    let allowed=true, reason='';
    try{checkApplicant_(r,notice);}catch(e){allowed=false;reason=e.message||'지원 불가';}
    return {key:memberKey_(r),name:String(r.name),branch:String(r.branch),
      seq:r.seq,allowed:allowed,reason:reason};
  });
}
function hashCancelCode_(code) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    String(code || ''), Utilities.Charset.UTF_8);
  return digest.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function resolveContact_(n) {
  const contact = n.contact || null;
  if (!contact) return null;
  if (contact.phone) return contact;
  const requested = String(contact.role || contact.name || '').replace(/\s+/g, '');
  let targetRole = '';
  if (/남부지부장|남부지부/.test(requested)) targetRole = '남부지부장';
  else if (/본부사무장|사무장/.test(requested)) targetRole = '사무장';
  if (!targetRole) return contact;
  try {
    const cache = CacheService.getScriptCache();
    let txt = cache.get('union-contact-data-v1');
    if (!txt) {
      const url = 'https://raw.githubusercontent.com/hktower226-lang/gyeonggi-info-dashboard/main/data.json';
      const res = UrlFetchApp.fetch(url, {muteHttpExceptions:true});
      if (res.getResponseCode()!==200) return contact;
      txt = res.getContentText('UTF-8');
      cache.put('union-contact-data-v1', txt, 600);
    }
    const matches = (JSON.parse(txt).contacts || []).filter(x => x.role===targetRole);
    if (matches.length===1) {
      return {branch:targetRole==='남부지부장'?'남부지부':'본부',
        role:matches[0].role, name:matches[0].name, phone:matches[0].phone};
    }
  } catch (e) {}
  return contact;
}
function smallNoteFlags_(person) {
  // 비고에서 명확히 표시된 내용만 활용한다. 날짜/문맥이 불명확하면 판단하지 않는다.
  const note=String(person.note||'').replace(/\s+/g,'');
  const completed=/소형(?:현장)?근무(?:완료|종료)|소형[^/]{0,20}근무완료/.test(note);
  const working=/소형(?:현장)?근무중|소형근무진행중/.test(note);
  const used=/소형이력서선정|이력서선정\(소형\)/.test(note);
  return {completed,working,used};
}
function checkApplicant_(person,notice) {
  const flags=smallNoteFlags_(person);
  if(noticeType_(notice)==='small') {
    if(person.isSub!=='O')throw new Error('소형근무 희망자(O)만 신청 가능합니다. 활동 미희망자는 신청할 수 없습니다.');
    if(flags.completed||flags.working)throw new Error('기존 비고상 소형근무 중 또는 완료로 표시되어 신청할 수 없습니다.');
    if(flags.used)throw new Error('기존 비고에 소형 이력서 선정 이력이 있어 중복 신청할 수 없습니다.');
  } else if(flags.working) {
    throw new Error('기존 비고상 소형근무 중이므로 대형 이력서 선정에 지원할 수 없습니다.');
  }
}
function alreadyAppliedForSmall_(sheet,person) {
  // 이 접수 시스템에 기록된 소형 지원 이력만 검사; 과거 외부 지원은 기존 비고 확인.
  const rows=sheet.getDataRange().getValues().slice(1);
  const currentKey=memberKey_(person);
  for(const r of rows){
    // 소형 이력서 신청 1회 제한: 취소한 접수도 해당 대기기간의 지원 이력으로 유지.
    if(!['접수','취소','반려'].includes(String(r[5])))continue;
    if(String(r[2]).trim()!==String(person.name).trim() ||
       normalizeBranch_(r[3])!==normalizeBranch_(person.branch))continue;
    if(r[7]&&String(r[7]).trim()!==String(person.endDate||'').trim())continue;
    let small=String(r[6]).toLowerCase()==='small';
    if(!small&&!r[6]) {
      try {small=noticeType_(notice_(String(r[1])))==='small';}catch(e) {}
    }
    if(small)return true;
  }
  return false;
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
      normalizeBranch_(x.branch) === normalizeBranch_(branch));
    const rank = current && Number.isFinite(Number(current.seq)) ?
      Number(current.seq) : Number(r[4]);
    people.push({name:name, branch:branch, seq:rank, time:new Date(r[0]).toISOString()});
  });
  people.sort((a,b) => a.seq-b.seq || a.time.localeCompare(b.time));
  const now = Date.now(), deadline = new Date(n.deadline).getTime();
  const isOpen = n.status === 'open' && Number.isFinite(deadline) && now < deadline;
  return {
    notice: {
      id:n.id, project:n.project, deadline:n.deadline,
      type:noticeType_(n),contact:resolveContact_(n),
      rankingUrl:n.signupUrl?(n.signupUrl+(n.signupUrl.indexOf('?')>=0?'&':'?')+'view=ranking'):''
    },
    accepting: isOpen, configured: n.status === 'open',
    now: new Date(now).toISOString(), participants:people,
    total:people.length
  };
}
function submitApplicant(id, memberKey, consent) {
  id=assertId_(id);
  if(!consent)throw new Error('본인 신청 여부를 확인해주세요.');
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(15000))throw new Error('접수가 많습니다. 다시 시도해주세요.');
  try {
    const n=notice_(id);
    if(n.status!=='open'||Date.now()>=new Date(n.deadline).getTime())
      throw new Error('접수기간이 아니거나 마감되었습니다.');
    const person=findMember_(memberKey);
    if(!person)throw new Error('선택한 조합원이 현재 대기명단과 일치하지 않습니다. 다시 검색해주세요.');
    checkApplicant_(person,n);
    const sheet=sheet_();
    const rows=sheet.getDataRange().getValues().slice(1);
    const duplicate=rows.some(r=>String(r[1])===id &&
      String(r[2]).trim()===String(person.name).trim() &&
      normalizeBranch_(r[3])===normalizeBranch_(person.branch) &&
      String(r[5])==='접수');
    if(duplicate)throw new Error('이미 접수된 조합원입니다.');
    if(noticeType_(n)==='small'&&alreadyAppliedForSmall_(sheet,person))
      throw new Error('이번 대기기간 소형 이력서 지원 기록이 있어 재지원할 수 없습니다.');
    if(Date.now()>=new Date(n.deadline).getTime())
      throw new Error('접수 마감시간이 지났습니다.');
    // 기존 시트의 A~F열을 유지하고, 소형 지원 이력용 G/H열을 사용.
    if(!sheet.getRange('G1').getValue())sheet.getRange('G1').setValue('모집유형');
    if(!sheet.getRange('H1').getValue())sheet.getRange('H1').setValue('대기회차 종료일');
    if(!sheet.getRange('I1').getValue())sheet.getRange('I1').setValue('취소코드 해시');
    if(!sheet.getRange('J1').getValue())sheet.getRange('J1').setValue('취소시각');
    const cancelCode = Utilities.getUuid() + '-' + Utilities.getUuid();
    sheet.appendRow([new Date(),id,String(person.name).trim(),
      normalizeBranch_(person.branch),Number(person.seq),'접수',
      noticeType_(n),String(person.endDate||''),hashCancelCode_(cancelCode),'']);
    return {ok:true,message:'희망 신청이 접수되었습니다.',cancelCode:cancelCode};
  } finally {
    lock.releaseLock();
  }
}

/**
 * 본인 접수 취소: 신청자에게만 발급된 비밀 취소코드를 대조.
 * 마감 전까지만 허용. 취소 행은 삭제하지 않고 감사 이력으로 남김.
 * 관리자 반려는 별도의 비밀번호 검사 함수에서만 수행.
 */
function cancelApplicant(id, cancelCode) {
  id = assertId_(id);
  cancelCode = String(cancelCode || '').trim();
  if (cancelCode.length < 50 || cancelCode.length > 100 ||
      !/^[0-9a-fA-F-]+$/.test(cancelCode))
    throw new Error('희망 신청 후 발급받은 취소코드를 정확히 입력해주세요.');
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(15000))throw new Error('접수가 많습니다. 다시 시도해주세요.');
  try {
    const notice=notice_(id);
    const deadline=new Date(notice.deadline).getTime();
    if (notice.status!=='open'||!Number.isFinite(deadline)||Date.now()>=deadline)
      throw new Error('접수 마감 후에는 취소할 수 없습니다.');
    const sheet=sheet_();
    const lastRow=sheet.getLastRow();
    if(lastRow<2)throw new Error('해당 신청 기록을 찾지 못했습니다.');
    const rows=sheet.getRange(2,1,lastRow-1,9).getDisplayValues();
    const hash=hashCancelCode_(cancelCode);
    const matches=[];
    for(let i=0;i<rows.length;i++){
      if(String(rows[i][1]).trim()===id && String(rows[i][8]).trim()===hash)
        matches.push({row:i+2,status:String(rows[i][5]).trim()});
    }
    if(matches.length!==1)throw new Error('취소코드가 일치하는 신청 기록이 없습니다.');
    if(matches[0].status==='취소')throw new Error('이미 취소된 신청입니다.');
    if(matches[0].status!=='접수')throw new Error('현재 취소 가능한 신청 상태가 아닙니다.');
    if(Date.now()>=deadline)throw new Error('접수 마감 후에는 취소할 수 없습니다.');
    sheet.getRange(matches[0].row,6).setValue('취소');
    sheet.getRange(matches[0].row,10).setValue(new Date());
    return {ok:true,message:'희망 신청이 취소되었습니다. 선순위 명단에서도 제외됩니다.'};
  } finally {
    lock.releaseLock();
  }
}


/**
 * 관리자 반려 비밀번호(12자 이상)를 Apps Script 프로젝트 설정 > 스크립트 속성에
 * ADMIN_REVIEW_PASSWORD 로 한 번 등록한다. 코드는 공개되는 GitHub에 암호를 넣지 않는다.
 * 공개 Web App 환경에서는 Google 계정 정보로 관리자 권한을 판별할 수 없으므로
 * 모든 관리자 작업 호출에서 서버 측 비밀번호 검사가 필수다.
 */
function assertAdmin_(password) {
  const configured = PropertiesService.getScriptProperties()
    .getProperty('ADMIN_REVIEW_PASSWORD');
  if (!configured || configured.length < 12)
    throw new Error('관리자 비밀번호가 아직 설정되지 않았습니다. Apps Script 스크립트 속성을 확인해주세요.');
  const given=String(password||'');
  if(given.length<12||given.length>256||
      hashCancelCode_(given)!==hashCancelCode_(configured))
    throw new Error('관리자 비밀번호가 일치하지 않습니다.');
}
function adminGetApplications(id, password) {
  assertAdmin_(password);
  id=assertId_(id);
  const rows=sheet_().getDataRange().getValues().slice(1);
  return rows.map((r,i)=>({
    sheetRow:i+2, noticeId:String(r[1]), name:String(r[2]||''),
    branch:String(r[3]||''), seq:r[4],
    status:String(r[5]||''), reason:String(r[10]||''),
    submittedAt: r[0] instanceof Date
      ? Utilities.formatDate(r[0],'Asia/Seoul','yyyy-MM-dd HH:mm:ss')
      : String(r[0]||'')
  })).filter(r=>r.noticeId===id);
}
function adminRejectApplicant(id, sheetRow, password, reason) {
  assertAdmin_(password);
  id=assertId_(id);
  const row=Number(sheetRow);
  reason=String(reason||'').trim();
  if(!Number.isInteger(row)||row<2)
    throw new Error('올바르지 않은 접수번호입니다.');
  if(reason.length>200)
    throw new Error('반려 사유는 200자 이내로 입력해주세요.');
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(15000))throw new Error('처리 요청이 많습니다. 다시 시도해주세요.');
  try {
    const sheet=sheet_();
    if(row>sheet.getLastRow())throw new Error('접수 기록이 존재하지 않습니다.');
    const record=sheet.getRange(row,1,1,6).getValues()[0];
    if(String(record[1])!==id)
      throw new Error('해당 공고의 접수 기록이 아닙니다.');
    if(String(record[5])!=='접수')
      throw new Error('이미 취소·반려되었거나 접수 상태가 아닙니다.');
    if(!sheet.getRange('K1').getValue())sheet.getRange('K1').setValue('반려사유');
    if(!sheet.getRange('L1').getValue())sheet.getRange('L1').setValue('반려시각');
    sheet.getRange(row,6).setValue('반려');
    sheet.getRange(row,11,1,2).setValues([[reason,new Date()]]);
    return {ok:true,message:String(record[2])+' 조합원 신청을 반려했습니다.'};
  } finally {
    lock.releaseLock();
  }
}
