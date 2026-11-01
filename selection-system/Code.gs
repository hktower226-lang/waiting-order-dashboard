/**
 * 경기지역본부 이력서 선정 접수 Web App (Google Apps Script)
 * 첫 실행 시 setup()을 실행하여 접수 저장용 Google Sheet 생성.
 * 배포: 웹 앱 / 실행 사용자 본인 / 액세스 권한 모든 사용자.
 * 주의: 성명 및 대기순번 선택만으로는 강한 본인인증이 불가능합니다.
 */
const REPO_RAW = 'https://raw.githubusercontent.com/hktower226-lang/waiting-order-dashboard/main/';
const SHEET_PROP = 'SELECTION_SHEET_ID';
const AUTH_PROP = 'AUTH_SHEET_ID'; // 비공개 인증명단 스프레드시트 ID, 스크립트 속성에서 설정
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
function sameApplicant_(r,p) {
  if(String(r[2]||'').trim()!==String(p.name||'').trim() ||
     normalizeBranch_(r[3])!==normalizeBranch_(p.branch)) return false;
  // 동명이인 오인 방지: 대기기간 종료일이 양쪽에 있을 때는 일치해야 함.
  return !(r[7] && p.endDate && String(r[7]).trim()!==String(p.endDate).trim());
}
function selectedRecord_(rows,p,excludeRow) {
  return rows.find((r,i)=>i+2!==excludeRow &&
    String(r[5]).trim()==='선정' && sameApplicant_(r,p));
}
function searchMembers(query, noticeId) {
  const notice=notice_(assertId_(noticeId));
  const q=String(query||'').trim().replace(/\s+/g,'');
  if(q.length<1||q.length>30)return [];
  const matches=roster_().filter(r=>{
    const name=String(r.name||'').replace(/\s+/g,'');
    return name.startsWith(q)||hangulInitials_(name).startsWith(q);
  });
  const selectedRows=sheet_().getDataRange().getValues().slice(1);
  return matches.map(r=>{
    let allowed=true, reason='';
    try{
      checkApplicant_(r,notice);
      if(selectedRecord_(selectedRows,r))
        throw new Error('다른 공고에서 이력서 선정된 상태입니다. 선정 취소 후 신청할 수 있습니다.');
    }catch(e){allowed=false;reason=e.message||'지원 불가';}
    return {key:memberKey_(r),name:String(r.name),branch:String(r.branch),
      seq:r.seq,allowed:allowed,reason:reason};
  });
}
function hashCancelCode_(code) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    String(code || ''), Utilities.Charset.UTF_8);
  return digest.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}

/**
 * 비공개 인증명단: A 성명 / B 소속지부 / C 생년월일.
 * 생년월일은 공개 GitHub·신청 응답·접수 저장 시트에 기록하지 않음.
 */
function birth6_(value) {
  const digits=String(value==null?'':value).replace(/\D/g,'');
  if(digits.length===8)return digits.slice(2);
  if(digits.length===6)return digits;
  return '';
}
function authenticationSheet_() {
  const id=PropertiesService.getScriptProperties().getProperty(AUTH_PROP);
  if(!id)throw new Error('관리자가 비공개 조합원 인증명단을 아직 연결하지 않았습니다.');
  const sheet=SpreadsheetApp.openById(id).getSheetByName('인증명단');
  if(!sheet)throw new Error('비공개 인증명단의 시트 이름을 확인해주세요.');
  const headers=sheet.getRange(1,1,1,3).getDisplayValues()[0];
  if(headers[0]!=='성명'||headers[1]!=='소속지부'||headers[2]!=='생년월일')
    throw new Error('비공개 인증명단 A~C열 제목이 일치하지 않습니다.');
  return sheet;
}
function birthMatches_(authRows,name,branch,birth6) {
  const matching=authRows.filter(r=>String(r[0]||'').trim()===name &&
    normalizeBranch_(r[1])===normalizeBranch_(branch));
  const births=[...new Set(matching.map(r=>birth6_(r[2])).filter(Boolean))];
  // 같은 성명·지부에 서로 다른 생년월일이 있으면 안전을 위해 직접 취소 보류.
  return births.length===1&&births[0]===birth6;
}
function cancelAttemptKey_(id,name) {
  const salt=PropertiesService.getScriptProperties().getProperty('AUTH_SALT')||'selection-cancel-rate-v1';
  const digest=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    salt+'|'+id+'|'+name,Utilities.Charset.UTF_8);
  return 'cancelfail:'+Utilities.base64EncodeWebSafe(digest).slice(0,40);
}
function cancelAttemptAllowed_(key) {
  if(Number(CacheService.getScriptCache().get(key)||0)>=5)
    throw new Error('인증 시도 횟수를 초과했습니다. 30분 후 다시 시도하거나 담당자에게 문의해주세요.');
}
function cancelAttemptFailed_(key) {
  const cache=CacheService.getScriptCache();
  cache.put(key,String(Number(cache.get(key)||0)+1),1800);
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
  // '신청'과 '선정'은 근무 이력이 아니다. 비고에 실제 근무 중/완료가 명확할 때만 제한한다.
  const note=String(person.note||'').replace(/\s+/g,'');
  const completed=/소형(?:현장)?근무(?:완료|종료)|소형근무[^/]{0,45}(?:완료|종료)/.test(note);
  const working=/소형(?:현장)?근무중|소형근무진행중/.test(note);
  return {completed,working};
}
function checkApplicant_(person,notice) {
  const flags=smallNoteFlags_(person);
  if(noticeType_(notice)==='small') {
    if(person.isSub!=='O')throw new Error('소형근무 희망자(O)만 신청 가능합니다. 활동 미희망자는 신청할 수 없습니다.');
    if(flags.completed||flags.working)throw new Error('기존 비고상 소형근무 중 또는 완료로 표시되어 신청할 수 없습니다.');
  } else if(flags.working) {
    throw new Error('기존 비고상 소형근무 중이므로 대형 이력서 선정에 지원할 수 없습니다.');
  }
}
function getState(id) {
  id=assertId_(id);
  const n=notice_(id);
  const rows=sheet_().getDataRange().getValues().slice(1);
  const active=rows.filter(r=>String(r[1])===id &&
     ['접수','선정'].includes(String(r[5])));
  const roster=roster_();
  const used=new Set(), people=[];
  active.forEach(r=>{
    const p={name:String(r[2]).trim(),branch:String(r[3]).trim(),endDate:r[7]||''};
    // 다른 공고에 이미 선정된 사람은 이 공고의 기존 희망접수 선순위에서 제외.
    if(String(r[5])==='접수' &&
       rows.some(x=>String(x[1])!==id && String(x[5])==='선정' && sameApplicant_(x,p)))
      return;
    const key=normalizeBranch_(p.branch)+'|'+p.name+'|'+p.endDate;
    if(used.has(key))return;
    used.add(key);
    const current=roster.find(x=>sameApplicant_(r,x));
    const rank=current && Number.isFinite(Number(current.seq))?Number(current.seq):Number(r[4]);
    people.push({name:p.name,branch:p.branch,seq:rank,
      status:String(r[5]),time:r[0] instanceof Date?r[0].toISOString():new Date(r[0]).toISOString()});
  });
  people.sort((a,b)=>a.seq-b.seq||a.time.localeCompare(b.time));
  const now=Date.now(),deadline=new Date(n.deadline).getTime();
  const isOpen=n.status==='open' && Number.isFinite(deadline) && now<deadline;
  return {
    notice:{
      id:n.id,project:n.project,deadline:n.deadline,
      type:noticeType_(n),contact:resolveContact_(n)
    },
    accepting:isOpen,configured:n.status==='open',
    now:new Date(now).toISOString(),
    participants:people.slice(0,3),
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
    if(selectedRecord_(rows,person))
      throw new Error('이미 이력서가 선정된 상태입니다. 선정 취소 후 다른 공고에 신청할 수 있습니다.');
    const duplicate=rows.some(r=>String(r[1])===id &&
      sameApplicant_(r,person) && ['접수','선정'].includes(String(r[5])));
    if(duplicate)throw new Error('이미 접수된 조합원입니다.');
    if(Date.now()>=new Date(n.deadline).getTime())
      throw new Error('접수 마감시간이 지났습니다.');
    // 기존 시트 A~F열 유지. G/H열은 모집유형/대기회차 기록용이며 신청 횟수 제한에 사용하지 않는다.
    if(!sheet.getRange('G1').getValue())sheet.getRange('G1').setValue('모집유형');
    if(!sheet.getRange('H1').getValue())sheet.getRange('H1').setValue('대기회차 종료일');
    // 취소용 생년월일은 기록하지 않는다. 과거 I열의 취소코드 해시는 보존만 함.
    sheet.appendRow([new Date(),id,String(person.name).trim(),
      normalizeBranch_(person.branch),Number(person.seq),'접수',
      noticeType_(n),String(person.endDate||'')]);
    return {ok:true,message:'희망 신청이 접수되었습니다.'};
  } finally {
    lock.releaseLock();
  }
}

/**
 * 본인 신청 취소: 성명과 생년월일(YYMMDD)을 비공개 명단에 서버에서 대조.
 * 일반 신청 취소는 접수 마감 전, 이력서 선정 취소는 마감 후에도 가능.
 * 개인정보를 공개 저장소, 접수 시트 및 브라우저 저장소에 기록하지 않는다.
 */
function cancelApplicant(id, name, birth6) {
  id=assertId_(id);
  name=String(name||'').trim().replace(/\s+/g,'');
  birth6=String(birth6||'').trim();
  if(name.length<2||name.length>40||!/^[0-9]{6}$/.test(birth6))
    throw new Error('신청자 성명과 생년월일 6자리(YYMMDD)를 정확히 입력해주세요.');
  const key=cancelAttemptKey_(id,name);
  cancelAttemptAllowed_(key);
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(15000))throw new Error('요청이 많습니다. 다시 시도해주세요.');
  try {
    const sheet=sheet_(), last=sheet.getLastRow();
    if(last<2)throw new Error('인증정보 또는 취소 가능한 신청 기록이 일치하지 않습니다.');
    const rows=sheet.getRange(2,1,last-1,8).getValues();
    const candidates=[];
    rows.forEach((r,i)=>{
      if(String(r[1])===id && String(r[2]||'').trim()===name &&
         ['접수','선정'].includes(String(r[5])))
        candidates.push({row:i+2,branch:String(r[3]||''),status:String(r[5])});
    });
    const authSheet=authenticationSheet_();
    const authLast=authSheet.getLastRow();
    const authRows=authLast>=2
      ?authSheet.getRange(2,1,authLast-1,3).getDisplayValues():[];
    const verified=candidates.filter(c=>birthMatches_(authRows,name,c.branch,birth6));
    if(verified.length!==1){
      cancelAttemptFailed_(key);
      throw new Error('인증정보 또는 취소 가능한 신청 기록이 일치하지 않습니다. 동명이인은 담당자에게 문의해주세요.');
    }
    const target=verified[0],notice=notice_(id);
    const deadline=new Date(notice.deadline).getTime();
    // 일반 취소는 마감 전까지. 선정이 확정된 경우 마감 후에도 취소할 수 있음.
    if(target.status==='접수'){
      if(notice.status!=='open'||!Number.isFinite(deadline)||Date.now()>=deadline)
        throw new Error('일반 희망 신청은 접수 마감 전까지만 취소할 수 있습니다.');
      if(!sheet.getRange('J1').getValue())sheet.getRange('J1').setValue('취소시각');
      sheet.getRange(target.row,6).setValue('취소');
      sheet.getRange(target.row,10).setValue(new Date());
    }else{
      if(!sheet.getRange('N1').getValue())sheet.getRange('N1').setValue('선정취소시각');
      sheet.getRange(target.row,6).setValue('선정취소');
      sheet.getRange(target.row,14).setValue(new Date());
    }
    CacheService.getScriptCache().remove(key);
    return {ok:true,message:'취소 처리되었습니다. 선순위 및 다른 공고 신청 제한에 반영됩니다.'};
  }finally{lock.releaseLock();}
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

function adminSetSelection(id,sheetRow,password,action) {
  assertAdmin_(password);
  id=assertId_(id);
  const row=Number(sheetRow);
  if(!Number.isInteger(row)||row<2)throw new Error('올바르지 않은 접수번호입니다.');
  if(action!=='select'&&action!=='release')throw new Error('올바르지 않은 처리 요청입니다.');
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(15000))throw new Error('요청이 많습니다. 다시 시도해주세요.');
  try{
    const sheet=sheet_(),last=sheet.getLastRow();
    if(row>last)throw new Error('접수 기록을 찾을 수 없습니다.');
    const data=sheet.getRange(2,1,last-1,14).getValues();
    const r=data[row-2];
    if(String(r[1])!==id)throw new Error('다른 공고의 접수 기록입니다.');
    if(action==='select'){
      if(String(r[5])!=='접수')throw new Error('접수 상태인 희망자만 선정할 수 있습니다.');
      const person={name:r[2],branch:r[3],endDate:r[7]};
      if(selectedRecord_(data,person,row))
        throw new Error('이미 다른 공고에서 선정된 조합원입니다. 중복 선정할 수 없습니다.');
      if(!sheet.getRange('M1').getValue())sheet.getRange('M1').setValue('선정시각');
      sheet.getRange(row,6).setValue('선정');
      sheet.getRange(row,13).setValue(new Date());
      return {ok:true,message:String(r[2])+' 조합원을 이력서 선정 처리했습니다.'};
    }
    if(String(r[5])!=='선정')throw new Error('선정 상태인 조합원만 선정 취소할 수 있습니다.');
    if(!sheet.getRange('N1').getValue())sheet.getRange('N1').setValue('선정취소시각');
    sheet.getRange(row,6).setValue('선정취소');
    sheet.getRange(row,14).setValue(new Date());
    return {ok:true,message:String(r[2])+' 조합원의 선정을 취소했습니다. 신청 제한이 해제됩니다.'};
  }finally{lock.releaseLock();}
}
