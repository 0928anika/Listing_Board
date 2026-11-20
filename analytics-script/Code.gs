/** Listing Analytics A004. Bound to a NEW spreadsheet, never the source. */
const LA_TABS=['월별 비교','주별 비교','작업자 월별','일별 기록','외부 지원','Node별','내부·외부 비교'];
function onOpen(){SpreadsheetApp.getUi().createMenu('누적 분석').addItem('초기 설정','setupListingAnalytics').addItem('전체 새로고침','refreshListingAnalytics').addItem('매일 자동 갱신 설치','installListingAnalyticsDaily').addToUi();}
function setupListingAnalytics(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();let s=ss.getSheetByName('CONFIG');
  if(!s){s=ss.insertSheet('CONFIG');s.getRange(1,1,6,3).setValues([
    ['설정','값','설명'],['SOURCE_SPREADSHEET_ID','','운영 스프레드시트 ID 또는 URL'],
    ['MIN_RANK_HOURS',1,'단계별 최소 누적 시간. 미만은 순위 제외'],
    ['EXTERNAL_WORKER','','B열에 외부 지원 이름. 같은 키로 행 추가 가능'],
    ['EXTERNAL_WORKER','','공백 제거 후 대소문자 구분 없이 정확히 일치'],
    ['정책','전체 이력 재계산','외부 목록 변경은 과거 전체 기간에도 반영']]);s.setFrozenRows(1);s.autoResizeColumns(1,3);}
  LA_TABS.forEach(n=>{if(!ss.getSheetByName(n))ss.insertSheet(n);});
  laPrepareWeights_(ss);
}
function laConfig_(ss){
  const s=ss.getSheetByName('CONFIG');if(!s)throw Error('초기 설정을 먼저 실행하세요.');
  const rows=s.getDataRange().getValues(), external=new Set();let id='',min=1;
  rows.slice(1).forEach(r=>{if(r[0]==='SOURCE_SPREADSHEET_ID')id=String(r[1]||'').trim();if(r[0]==='MIN_RANK_HOURS'){min=Number(r[1]);if(!Number.isFinite(min)||min<0)throw Error('MIN_RANK_HOURS는 0 이상의 숫자여야 합니다.');}if(r[0]==='EXTERNAL_WORKER'&&String(r[1]||'').trim())external.add(laName_(r[1]));});
  const match=id.match(/\/d\/([\w-]+)/);if(match)id=match[1];
  if(!/^[\w-]{20,}$/.test(id)||id===ss.getId())throw Error('별도 운영 파일의 올바른 ID를 입력하세요. 분석 파일 자체는 원본으로 사용할 수 없습니다.');
  const weights={};rows.slice(1).forEach(r=>{const name=laName_(r[6]);if(!name)return;const ratio=r[7]===''||r[7]===undefined?1:Number(r[7]);if(!Number.isFinite(ratio)||ratio<=0||ratio>1)throw Error('근무비율은 0 초과 1 이하: '+r[6]);if(Object.prototype.hasOwnProperty.call(weights,name))throw Error('근무비율 작업자 중복: '+r[6]);weights[name]=ratio;});
  return {id:id,min:min,external:external,weights:weights};
}
function laPrepareWeights_(ss){const s=ss.getSheetByName('CONFIG');if(!s)return;if(s.getMaxColumns()<9)s.insertColumnsAfter(s.getMaxColumns(),9-s.getMaxColumns());const header=s.getRange('G1:I1').getValues()[0];if(header.every(v=>v===''))s.getRange('G1:I2').setValues([['작업자','근무비율','설명'],['',1,'기본 1 / 절반 0.5. 이름은 Worker와 동일하게 입력']]);else if(header[0]!=='작업자'||header[1]!=='근무비율')throw Error('CONFIG G:I에 다른 데이터가 있습니다. 빈 영역으로 옮긴 후 실행하세요.');}
function laName_(v){return String(v||'').trim().toLowerCase();}
function laWorkerDays_(sessions,weights){
  const groups={};sessions.forEach(s=>{const month=s.day.slice(0,7),key=JSON.stringify([month,s.name]),g=groups[key]||(groups[key]={month:month,worker:s.worker,external:s.external,ratio:Object.prototype.hasOwnProperty.call(weights,s.name)?weights[s.name]:1,v2:0,v3:0,dates:{}});g[s.stage==='2P'?'v2':'v3']+=s.delta;g.dates[s.day]=(g.dates[s.day]||0)+s.hours;});
  return Object.values(groups).map(g=>{g.days=Object.values(g.dates).filter(h=>h>0).length;g.equivalentDays=g.days*g.ratio;return g;});
}
function laGet_(url){const r=UrlFetchApp.fetch(url,{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true});if(r.getResponseCode()!==200)throw Error('운영 파일 읽기 실패: HTTP '+r.getResponseCode()+'. 파일 ID와 조회 권한을 확인하세요.');return JSON.parse(r.getContentText());}
function laRead_(id){
  const base='https://sheets.googleapis.com/v4/spreadsheets/'+encodeURIComponent(id), meta=laGet_(base+'?fields=sheets.properties.title'), names=meta.sheets.map(s=>s.properties.title);
  ['REQUESTS','WORK_SESSIONS'].forEach(n=>{if(names.indexOf(n)<0)throw Error('필수 원본 시트 누락: '+n);});
  const wanted=['REQUESTS','WORK_SESSIONS','LONG_TERM_SESSIONS'].filter(n=>names.indexOf(n)>=0);
  // One read request for all source values. No source writes, repairs, or triggers.
  const result=laGet_(base+'/values:batchGet?valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=SERIAL_NUMBER&'+wanted.map(n=>'ranges='+encodeURIComponent("'"+n+"'")).join('&'));
  const out={};wanted.forEach((n,i)=>out[n]=result.valueRanges[i].values||[]);return out;
}
function laTable_(rows,required,label){
  if(!rows||!rows.length)throw Error(label+' 헤더 없음');const h={};rows[0].forEach((n,i)=>{if(n&&h[n]!==undefined)throw Error(label+' 중복 헤더: '+n);h[n]=i;});required.forEach(n=>{if(h[n]===undefined)throw Error(label+' 헤더 누락: '+n);});return {h:h,rows:rows.slice(1).filter(r=>r.some(v=>v!==''&&v!==null))};
}
// Sheets numeric date serial already represents local wall time. Text dates are rejected.
function laDay_(v){return typeof v==='number'&&Number.isFinite(v)&&v>0?new Date(Math.floor(v-25569)*86400000).toISOString().slice(0,10):'';}
function laWeek_(day){const d=new Date(day+'T00:00:00Z');d.setUTCDate(d.getUTCDate()-(d.getUTCDay()+6)%7);return d.toISOString().slice(0,10);}
function laNumber_(v){return typeof v==='number'&&Number.isFinite(v);}
function laBuild_(data,c,today){
  const issues=[],sessions=[],req={},seen=new Set(),aggSeen=new Set(),rawGroups=new Set();
  const warn=(type,id,detail)=>issues.push([type,id,detail]);
  const rt=laTable_(data.REQUESTS,['Request ID','Node','Status','Requested At','Completed At'],'REQUESTS');
  rt.rows.forEach(r=>{const v=n=>r[rt.h[n]],id=String(v('Request ID')||'').trim();if(!id)return;if(req[id])throw Error('중복 Request ID: '+id);req[id]={id:id,node:String(v('Node')||'미지정'),status:String(v('Status')||''),requested:laDay_(v('Requested At')),completed:laDay_(v('Completed At')),start:v('Requested At'),end:v('Completed At')};});
  function add(r,h,aggregate){
    const v=n=>r[h[n]],id=String(v('Request ID')||''),worker=String(v('Worker')||'').trim(),stage=String(v('Stage')||''),status=String(v('Status')||'');
    if(aggregate&&status!=='COMMITTED')return;if(!aggregate&&status==='ACTIVE'){warn('ACTIVE 제외',id,worker);return;}if(stage!=='3P'&&stage!=='2P')return;
    const day=aggregate?String(v('Work Date')||''):laDay_(v('Ended At'));
    if(!/^\d{4}-\d{2}-\d{2}$/.test(day)||!worker||!id){warn('식별값/종료일 누락',id,worker);return;}
    const delta=v('Completed VI Delta'),minutes=v('Duration Minutes');
    if(!laNumber_(delta)||!laNumber_(minutes)||minutes<0){warn('VI/시간 누락 또는 오류',id,worker+' '+day);return;}
    if(day>today){warn('미래 기록 제외',id,day);return;}
    const group=JSON.stringify([id,laName_(worker),stage,day]);
    if(aggregate){const key=String(v('Aggregate ID')||'');if(!key||aggSeen.has(key))throw Error('장기 집계 ID 누락/중복: '+key);aggSeen.add(key);if(rawGroups.has(group))throw Error('원본 세션과 장기 집계가 겹칩니다. 백업 완료 후 재조회: '+id+' '+day);}
    else {const key=String(v('Session ID')||'');if(key){if(seen.has(key))throw Error('중복 Session ID: '+key);seen.add(key);}else warn('수기 세션 포함',id,worker+' '+day);rawGroups.add(group);}
    if(!req[id])warn('요청 없는 세션 포함',id,worker);
    if(minutes===0&&delta!==0)warn('시간 0 실적',id,worker+' '+day+' VI='+delta);
    sessions.push({id:id,worker:worker,name:laName_(worker),stage:stage,day:day,delta:delta,hours:minutes/60,external:c.external.has(laName_(worker)),node:req[id]?req[id].node:'미지정'});
  }
  const st=laTable_(data.WORK_SESSIONS,['Session ID','Worker','Request ID','Stage','Status','Ended At','Duration Minutes','Completed VI Delta'],'WORK_SESSIONS');st.rows.forEach(r=>add(r,st.h,false));
  if(data.LONG_TERM_SESSIONS){const lt=laTable_(data.LONG_TERM_SESSIONS,['Aggregate ID','Status','Request ID','Worker','Stage','Work Date','Duration Minutes','Completed VI Delta'],'LONG_TERM_SESSIONS');lt.rows.forEach(r=>add(r,lt.h,true));}else warn('장기 집계 없음','','과거 백업 세션 포함 여부를 확인하세요. 상세 아카이브는 합산하지 않습니다.');
  function summarize(list,kind,worker){
    const groups={};list.forEach(s=>{const period=kind==='월'?s.day.slice(0,7):kind==='주'?laWeek_(s.day):s.day,k=JSON.stringify([period,worker?s.name:'']);const g=groups[k]||(groups[k]={period:period,worker:worker?s.worker:'전체',v3:0,v2:0,h3:0,h2:0,days:new Set(),requests:new Set()});g[s.stage==='3P'?'v3':'v2']+=s.delta;g[s.stage==='3P'?'h3':'h2']+=s.hours;g.days.add(s.day);g.requests.add(s.id);});return Object.values(groups).sort((a,b)=>a.period.localeCompare(b.period)||a.worker.localeCompare(b.worker));
  }
  const headers=['기간','작업자','3P 완료 VI','2P 완료 VI','완료 VI 합계','3P 작업시간(h)','2P 작업시간(h)','작업시간 합계(h)','3P VI/h','2P VI/h','기록일 수','작업 REQ 수','3P 순위','2P 순위'];
  function rows(gs){const ranks={};['3','2'].forEach(stage=>{const periods={};gs.forEach(g=>{if(g['h'+stage]>0&&g['h'+stage]>=c.min)(periods[g.period]||(periods[g.period]=[])).push(g);});Object.keys(periods).forEach(p=>{let prev=null,rank=0;periods[p].sort((a,b)=>b['v'+stage]/b['h'+stage]-a['v'+stage]/a['h'+stage]).forEach((g,i)=>{const rate=g['v'+stage]/g['h'+stage];if(rate!==prev)rank=i+1;prev=rate;ranks[JSON.stringify([p,g.worker,stage])]=rank;});});});return gs.map(g=>[g.period,g.worker,g.v3,g.v2,g.v3+g.v2,g.h3,g.h2,g.h3+g.h2,g.h3>0?g.v3/g.h3:'',g.h2>0?g.v2/g.h2:'',g.days.size,g.requests.size,ranks[JSON.stringify([g.period,g.worker,'3'])]||'',ranks[JSON.stringify([g.period,g.worker,'2'])]||'']);}
  const internal=sessions.filter(s=>!s.external),external=sessions.filter(s=>s.external),tabs={};
  tabs._workerDays=laWorkerDays_(sessions,c.weights||{});
  tabs._nodeCompletion={};Object.values(req).forEach(r=>{if(r.status!=='DONE'||!r.completed||r.completed>today)return;const key=JSON.stringify([r.completed.slice(0,7),r.node]),g=tabs._nodeCompletion[key]||(tabs._nodeCompletion[key]={current:0,prior:0,unknown:0});if(!r.requested||!laNumber_(r.start)||!laNumber_(r.end)||r.start>r.end)g.unknown++;else if(r.requested.slice(0,7)===r.completed.slice(0,7))g.current++;else g.prior++;});
  [['월別','월'],['주別','주']].forEach(pair=>{const kind=pair[1],gs=summarize(internal,kind,false),rs=rows(gs);tabs[kind==='월'?'월별 비교':'주별 비교']=[headers.slice(0,12).concat(['직전 달력기간 VI','VI 증감률','직전 달력기간 시간','시간 증감률','집계 상태']),rs.map((r,i)=>{const d=new Date((kind==='월'?r[0]+'-01':r[0])+'T00:00:00Z');if(kind==='월')d.setUTCMonth(d.getUTCMonth()-1);else d.setUTCDate(d.getUTCDate()-7);const prevKey=d.toISOString().slice(0,kind==='월'?7:10),prev=rs.find(x=>x[0]===prevKey),open=kind==='월'?r[0]===today.slice(0,7):r[0]===laWeek_(today);return r.slice(0,12).concat([prev?prev[4]:'',prev&&prev[4]!==0?(r[4]-prev[4])/prev[4]:'',prev?prev[7]:'',prev&&prev[7]!==0?(r[7]-prev[7])/prev[7]:'',open?'진행 중: 이전 전체기간과 비교':'종료 기간: 재계산 가능']);})];});
  tabs['작업자 월별']=[headers,rows(summarize(internal,'월',true))];tabs['작업자 주별']=[headers,rows(summarize(internal,'주',true))];tabs['일별 기록']=[headers.slice(0,12),rows(summarize(internal,'일',true)).map(r=>r.slice(0,12))];
  tabs['외부 지원']=[['집계 단위'].concat(headers),['월','주','일'].flatMap(k=>rows(summarize(external,k,true)).map(r=>[k].concat(r)))];
  const nodes={};function node(period,node){const key=JSON.stringify([period,node]);return nodes[key]||(nodes[key]={period:period,node:node,incoming:0,completed:0,leads:[],v3:0,v2:0,h3:0,h2:0});}
  Object.values(req).forEach(r=>{if(r.requested)node(r.requested.slice(0,7),r.node).incoming++;if(r.completed&&r.completed<=today&&r.status==='DONE'){const g=node(r.completed.slice(0,7),r.node);g.completed++;if(laNumber_(r.start)&&laNumber_(r.end)&&r.end>=r.start)g.leads.push((r.end-r.start)*24);}});
  internal.forEach(s=>{const g=node(s.day.slice(0,7),s.node);g[s.stage==='3P'?'v3':'v2']+=s.delta;g[s.stage==='3P'?'h3':'h2']+=s.hours;});
  tabs['Node별']=[['월','Node','접수 요청(전체)','완료 요청(전체)','리드타임 중앙값(h,전체)','내부 3P VI','내부 2P VI','내부 3P 시간(h)','내부 2P 시간(h)'],Object.values(nodes).sort((a,b)=>a.period.localeCompare(b.period)||a.node.localeCompare(b.node)).map(g=>{g.leads.sort((a,b)=>a-b);const n=g.leads.length;return [g.period,g.node,g.incoming,g.completed,n?(g.leads[Math.floor((n-1)/2)]+g.leads[Math.floor(n/2)])/2:'',g.v3,g.v2,g.h3,g.h2];})];
  // Keep request-level samples so cross-month medians are never averaged.
  tabs._monthlyLeadTime=Object.values(req).filter(r=>r.status==='DONE'&&r.completed&&r.completed<=today&&laNumber_(r.start)&&laNumber_(r.end)&&r.end>=r.start).map(r=>({month:r.completed.slice(0,7),hours:(r.end-r.start)*24}));
  c.external.forEach(n=>{if(!sessions.some(s=>s.name===n))warn('외부 이름 매칭 없음',n,'기간/이름 표기를 확인하세요.');});
  tabs['데이터 점검']=[['항목','Request/작업자','설명'],[['원본 범위','','REQUESTS + WORK_SESSIONS + COMMITTED LONG_TERM_SESSIONS. 상세 아카이브 직접 합산 안 함'],['외부 제외 정책','','외부 목록은 과거 전체에 소급. 내부 VI/시간/순위에서 제외. 요청 접수/완료는 전체 업무량'],['기간 정책','','Asia/Seoul, 종료일 귀속, 월요일 시작, 월력 기준. ACTIVE 제외'],['기록일 정책','','세션 종료일의 고유 일수. 출근일 아님'],['순위 정책','','단계별 합산 VI/합산 시간, 최소 '+c.min+'시간, 동률 공동순위. 전체 재계산'],['원본 한계','','현재 Request Node 기준. 과거 월말 상태 및 당시 Node는 복원하지 않음']].concat(issues)];return tabs;
}
function refreshListingAnalytics(){
  const lock=LockService.getDocumentLock();if(!lock.tryLock(1000))throw Error('다른 분석 갱신이 실행 중입니다.');
  const ss=SpreadsheetApp.getActiveSpreadsheet();try{
    laPrepareWeights_(ss);
    const c=laConfig_(ss),now=new Date(),today=Utilities.formatDate(now,'Asia/Seoul','yyyy-MM-dd'),tabs=laBuild_(laRead_(c.id),c,today);
    // Validate all input before changing reports. Last successful timestamp written LAST.
    laRenderReports_(ss,laLayouts_(tabs),Utilities.formatDate(now,'Asia/Seoul','yyyy-MM-dd HH:mm:ss'));
    // User-requested removal, only in this validated analysis destination.
    ['작업자 주별','데이터 점검'].forEach(n=>{const s=ss.getSheetByName(n);if(s)ss.deleteSheet(s);});
    const notices=tabs['데이터 점검'][1].slice(6);console.log(JSON.stringify({analyticsNotices:notices}));
    ss.getSheetByName('CONFIG').getRange('E5:F5').setValues([['집계 참고사항 수',notices.length]]);
    const cfg=ss.getSheetByName('CONFIG');cfg.getRange('E1:F3').setValues([['갱신 상태','성공'],['최종 성공(KST)',Utilities.formatDate(now,'Asia/Seoul','yyyy-MM-dd HH:mm:ss')],['버전','A004']]);
  }catch(e){const cfg=ss.getSheetByName('CONFIG');if(cfg)cfg.getRange('E1:F1').setValues([['갱신 상태','실패: '+e.message]]);throw e;}finally{lock.releaseLock();}
}
function installListingAnalyticsDaily(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();laConfig_(ss);
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='refreshListingAnalytics').forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('refreshListingAnalytics').timeBased().atHour(6).everyDays(1).inTimezone('Asia/Seoul').create();
}



