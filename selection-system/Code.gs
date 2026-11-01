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
  tpl.view = (e && e.parameter && e.parameter.view==='ranking') ? 'ranking' : 'apply';
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
    if(String(r[5])!=='접수')continue;
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
      id:n.id, project:n.project, site:n.site, rental:n.rental,
      wage:n.wage, floors:n.floors, machine:n.machine,
      plannedInstallation:n.plannedInstallation, startCondition:n.startCondition,
      duration:n.duration, notes:n.notes || [], deadline:n.deadline,
      type:noticeType_(n),contact:n.contact||null,
      rankingUrl:n.signupUrl?(n.signupUrl+(n.signupUrl.indexOf('?')>=0?'&':'?')+'view=ranking'):''
    },
    accepting: isOpen, configured: n.status === 'open',
    now: new Date(now).toISOString(), participants:people,
    waitingRanks:roster.map(r=>({seq:r.seq,name:r.name,branch:r.branch}))
      .filter(r=>Number.isFinite(Number(r.seq)))
      .sort((a,b)=>Number(a.seq)-Number(b.seq)),
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
    sheet.appendRow([new Date(),id,String(person.name).trim(),
      normalizeBranch_(person.branch),Number(person.seq),'접수',
      noticeType_(n),String(person.endDate||'')]);
    return {ok:true,message:'희망 신청이 접수되었습니다.'};
  } finally {
    lock.releaseLock();
  }
}
