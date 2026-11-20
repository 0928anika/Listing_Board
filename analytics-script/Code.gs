/** Listing Analytics A006. Bound to the analysis spreadsheet, never the source. */
const LA_TABS=['월별 비교','주별 비교','작업자 월별','일별 기록','외부 지원','Node별','내부·외부 비교','크롤링 효과','투입 계획'];
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
  return Object.values(groups).map(g=>{g.days=Object.keys(g.dates).length;g.equivalentDays=g.days*g.ratio;return g;});
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
  const issues=[],sessions=[],req={},seen=new Set(),aggSeen=new Set(),rawGroups=new Set(),crawlMeta=laCrawlMeta_(data.REQUESTS);
  const warn=(type,id,detail)=>issues.push([type,id,detail]);
  const rt=laTable_(data.REQUESTS,['Request ID','Node','Status','Requested At','Completed At'],'REQUESTS');
  rt.rows.forEach(r=>{const v=n=>r[rt.h[n]],id=String(v('Request ID')||'').trim();if(!id)return;if(req[id])throw Error('중복 Request ID: '+id);req[id]={id:id,node:String(v('Node')||'미지정'),status:String(v('Status')||''),requested:laDay_(v('Requested At')),completed:laDay_(v('Completed At')),start:v('Requested At'),end:v('Completed At'),blocked:v('Blocked Time (Hours)')};});
  function add(r,h,aggregate){
    const v=n=>r[h[n]],id=String(v('Request ID')||''),worker=String(v('Worker')||'').trim(),stage=String(v('Stage')||''),status=String(v('Status')||'');
    if(aggregate&&status!=='COMMITTED')return;if(!aggregate&&status==='ACTIVE'){warn('ACTIVE 제외',id,worker);return;}if(stage!=='3P'&&stage!=='2P')return;
    const day=aggregate?String(v('Work Date')||''):laDay_(v('Ended At'));
    if(!/^\d{4}-\d{2}-\d{2}$/.test(day)||!worker||!id){warn('식별값/종료일 누락',id,worker);return;}
    const delta=v('Completed VI Delta'),rawMinutes=v('Duration Minutes');
    if(!laNumber_(delta)){warn('VI/시간 누락 또는 오류',id,worker+' '+day);return;}
    const minutes=laNumber_(rawMinutes)&&rawMinutes>=0?rawMinutes:0;
    if(!laNumber_(rawMinutes)||rawMinutes<0)warn('VI/시간 누락 또는 오류',id,worker+' '+day+' (일 기준 실적은 포함, 시간은 제외)');
    if(day>today){warn('미래 기록 제외',id,day);return;}
    const group=JSON.stringify([id,laName_(worker),stage,day]);
    if(aggregate){const key=String(v('Aggregate ID')||'');if(!key||aggSeen.has(key))throw Error('장기 집계 ID 누락/중복: '+key);aggSeen.add(key);if(rawGroups.has(group))throw Error('원본 세션과 장기 집계가 겹칩니다. 백업 완료 후 재조회: '+id+' '+day);}
    else {const key=String(v('Session ID')||'');if(key){if(seen.has(key))throw Error('중복 Session ID: '+key);seen.add(key);}else warn('수기 세션 포함',id,worker+' '+day);rawGroups.add(group);}
    if(!req[id])warn('요청 없는 세션 포함',id,worker);
    if(minutes===0&&delta!==0)warn('시간 0 실적',id,worker+' '+day+' VI='+delta);
    sessions.push({id:id,worker:worker,name:laName_(worker),stage:stage,day:day,delta:delta,hours:minutes/60,ratio:Object.prototype.hasOwnProperty.call(c.weights||{},laName_(worker))?c.weights[laName_(worker)]:1,external:c.external.has(laName_(worker)),node:req[id]?req[id].node:'미지정',crawl:laCrawlClass_(crawlMeta[id])});
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
  tabs._dayProductivity=laDayProductivity_(sessions);
  tabs._crawl=laCrawlAnalysis_(data,sessions,today);
  tabs._crawl.preparation=laCrawlPreparation_(data,today);
  tabs._crawl.monthlyPreparation=laCrawlMonthly_(data,today,c.weights||{});
  tabs._plan=laPlan_(data,tabs._crawl,c.plan||{},today);
  tabs._nodeCompletion={};Object.values(req).forEach(r=>{if(r.status!=='DONE'||!r.completed||r.completed>today)return;const key=JSON.stringify([r.completed.slice(0,7),r.node]),g=tabs._nodeCompletion[key]||(tabs._nodeCompletion[key]={current:0,prior:0,unknown:0});if(!r.requested||!laNumber_(r.start)||!laNumber_(r.end)||r.start>r.end)g.unknown++;else if(r.requested.slice(0,7)===r.completed.slice(0,7))g.current++;else g.prior++;});
  [['월別','월'],['주別','주']].forEach(pair=>{const kind=pair[1],gs=summarize(internal,kind,false),rs=rows(gs);tabs[kind==='월'?'월별 비교':'주별 비교']=[headers.slice(0,12).concat(['직전 달력기간 VI','VI 증감률','직전 달력기간 시간','시간 증감률','집계 상태']),rs.map((r,i)=>{const d=new Date((kind==='월'?r[0]+'-01':r[0])+'T00:00:00Z');if(kind==='월')d.setUTCMonth(d.getUTCMonth()-1);else d.setUTCDate(d.getUTCDate()-7);const prevKey=d.toISOString().slice(0,kind==='월'?7:10),prev=rs.find(x=>x[0]===prevKey),open=kind==='월'?r[0]===today.slice(0,7):r[0]===laWeek_(today);return r.slice(0,12).concat([prev?prev[4]:'',prev&&prev[4]!==0?(r[4]-prev[4])/prev[4]:'',prev?prev[7]:'',prev&&prev[7]!==0?(r[7]-prev[7])/prev[7]:'',open?'진행 중: 이전 전체기간과 비교':'종료 기간: 재계산 가능']);})];});
  tabs['작업자 월별']=[headers,rows(summarize(internal,'월',true))];tabs['작업자 주별']=[headers,rows(summarize(internal,'주',true))];tabs['일별 기록']=[headers.slice(0,12),rows(summarize(internal,'일',true)).map(r=>r.slice(0,12))];
  tabs['외부 지원']=[['집계 단위'].concat(headers),['월','주','일'].flatMap(k=>rows(summarize(external,k,true)).map(r=>[k].concat(r)))];
  const nodes={};function node(period,node){const key=JSON.stringify([period,node]);return nodes[key]||(nodes[key]={period:period,node:node,incoming:0,completed:0,leads:[],v3:0,v2:0,h3:0,h2:0});}
  Object.values(req).forEach(r=>{if(r.requested)node(r.requested.slice(0,7),r.node).incoming++;if(r.completed&&r.completed<=today&&r.status==='DONE'){const g=node(r.completed.slice(0,7),r.node);g.completed++;if(laNumber_(r.start)&&laNumber_(r.end)&&r.end>=r.start)g.leads.push((r.end-r.start)*24);}});
  internal.forEach(s=>{const g=node(s.day.slice(0,7),s.node);g[s.stage==='3P'?'v3':'v2']+=s.delta;g[s.stage==='3P'?'h3':'h2']+=s.hours;});
  tabs['Node별']=[['월','Node','접수 요청(전체)','완료 요청(전체)','리드타임 중앙값(h,전체)','내부 3P VI','내부 2P VI','내부 3P 시간(h)','내부 2P 시간(h)'],Object.values(nodes).sort((a,b)=>a.period.localeCompare(b.period)||a.node.localeCompare(b.node)).map(g=>{g.leads.sort((a,b)=>a-b);const n=g.leads.length;return [g.period,g.node,g.incoming,g.completed,n?(g.leads[Math.floor((n-1)/2)]+g.leads[Math.floor(n/2)])/2:'',g.v3,g.v2,g.h3,g.h2];})];
  // Keep request-level samples so cross-month medians are never averaged.
  tabs._monthlyLeadTime=Object.values(req).filter(r=>r.status==='DONE'&&r.completed&&r.completed<=today&&laNumber_(r.start)&&laNumber_(r.end)&&r.end>=r.start).map(r=>{const hours=(r.end-r.start)*24,valid=laNumber_(r.blocked)&&r.blocked>=0&&r.blocked<=hours;return {month:r.completed.slice(0,7),node:r.node,hours:hours,blocked:valid?r.blocked:null,adjusted:valid?hours-r.blocked:null};});
  c.external.forEach(n=>{if(!sessions.some(s=>s.name===n))warn('외부 이름 매칭 없음',n,'기간/이름 표기를 확인하세요.');});
  tabs['데이터 점검']=[['항목','Request/작업자','설명'],[['원본 범위','','REQUESTS + WORK_SESSIONS + COMMITTED LONG_TERM_SESSIONS. 상세 아카이브 직접 합산 안 함'],['외부 제외 정책','','외부 목록은 과거 전체에 소급. 내부 VI/시간/순위에서 제외. 요청 접수/완료는 전체 업무량'],['기간 정책','','Asia/Seoul, 종료일 귀속, 월요일 시작, 월력 기준. ACTIVE 제외'],['기록일 정책','','세션 종료일의 고유 일수. 출근일 아님'],['순위 정책','','단계별 합산 VI/합산 시간, 최소 '+c.min+'시간, 동률 공동순위. 전체 재계산'],['원본 한계','','현재 Request Node 기준. 과거 월말 상태 및 당시 Node는 복원하지 않음']].concat(issues)];return tabs;
}
// Classification intentionally uses current stage only, per the reporting contract.
function laCrawlMeta_(rows){
  const result=Object.create(null);if(!rows||!rows.length)return result;
  const h={};rows[0].forEach((n,i)=>h[n]=i);
  rows.slice(1).forEach(r=>{const id=String(r[h['Request ID']]||'').trim();if(id)result[id]={stage:String(r[h['크롤링 단계']]||'').trim()};});return result;
}
function laCrawlClass_(m){
  if(!m)return '요청 메타데이터 없음';
  const stage=String(m.stage||'').trim().toUpperCase();
  if(stage==='CRAWL_COMPLETED')return '크롤링';
  if(stage==='N/A')return '비크롤링';
  return '단계 미입력/미완료/기타';
}
function laDayProductivity_(sessions){
  const maps={};['월별 비교','주별 비교','작업자 월별','일별 기록','외부 지원','Node별','내부·외부 비교'].forEach(n=>maps[n]=new Map());
  function add(name,period,label,s){const key=JSON.stringify([period,label]),map=maps[name];if(!map.has(key))map.set(key,{period:period,label:label,v2:0,v3:0,days:new Set(),equiv:0});const g=map.get(key),d=JSON.stringify([s.name,s.day]);if(!g.days.has(d)){g.days.add(d);g.equiv+=s.ratio||1;}g[s.stage==='2P'?'v2':'v3']+=s.delta;}
  sessions.forEach(s=>{const month=s.day.slice(0,7);add('내부·외부 비교',month,s.external?'외부 지원':'내부',s);if(s.external){add('외부 지원',month,s.worker,s);return;}add('월별 비교',month,'내부 전체',s);add('주별 비교',laWeek_(s.day),'내부 전체',s);add('작업자 월별',month,s.worker,s);add('일별 기록',s.day,s.worker,s);add('Node별',month,s.node,s);});
  const result={};Object.keys(maps).forEach(n=>result[n]=Array.from(maps[n].values()).sort((a,b)=>b.period.localeCompare(a.period)||a.label.localeCompare(b.label)).map(g=>[g.period,g.label,g.v2,g.v3,g.v2+g.v3,g.days.size,g.equiv,g.equiv?g.v2/g.equiv:'',g.equiv?g.v3/g.equiv:'',g.equiv?(g.v2+g.v3)/g.equiv:'']));return result;
}
function laCrawlBucket_(){return {vi:0,hours:0,timedVi:0,untimedVi:0,requests:new Set(),records:0,dayKeys:new Set(),equivalentDays:0};}
function laCrawlAdd_(g,s){const dayKey=JSON.stringify([s.name,s.day]);if(!g.dayKeys.has(dayKey)){g.dayKeys.add(dayKey);g.equivalentDays+=s.ratio||1;}g.vi+=s.delta;g.records++;g.requests.add(s.id);if(s.hours>0){g.hours+=s.hours;g.timedVi+=s.delta;}else g.untimedVi+=s.delta;}
function laCrawlRate_(g){return g.hours>0?g.timedVi/g.hours:'';}
function laCrawlRatio_(a,b){return typeof a==='number'&&typeof b==='number'&&a>0&&b>0?a/b:'';}
function laCrawlAnalysis_(data,sessions,today){
  const monthly=new Map(),daily=new Map(),workers=new Map(),nodes=new Map(),reasons=new Map();
  function groups(map,key){if(!map.has(key))map.set(key,{crawl:laCrawlBucket_(),manual:laCrawlBucket_(),unknown:laCrawlBucket_()});return map.get(key);}
  const samples=sessions.filter(s=>s.stage==='3P'&&!s.external);
  samples.forEach(s=>{
    const type=s.crawl==='크롤링'?'crawl':s.crawl==='비크롤링'?'manual':'unknown',month=s.day.slice(0,7);
    [groups(monthly,month),groups(daily,s.day),groups(workers,JSON.stringify([month,s.name])),groups(nodes,JSON.stringify([month,s.node]))].forEach(g=>laCrawlAdd_(g[type],s));
    if(type==='unknown'){const key=JSON.stringify([month,s.crawl]);if(!reasons.has(key))reasons.set(key,laCrawlBucket_());laCrawlAdd_(reasons.get(key),s);}
  });
  const labels={crawl:'크롤링',manual:'비크롤링',unknown:'확인 불가'};
  const monthlyRows=[];Array.from(monthly.keys()).sort().reverse().forEach(month=>Object.keys(labels).forEach(k=>{const g=monthly.get(month)[k];monthlyRows.push([month,labels[k],g.vi,g.hours,g.timedVi,laCrawlRate_(g),g.requests.size,g.records,g.untimedVi]);}));
  function rates(g){const total=g.crawl.vi+g.manual.vi+g.unknown.vi,known=g.crawl.vi+g.manual.vi,positive=Object.values(g).every(b=>b.vi>=0),cr=laCrawlRate_(g.crawl),mr=laCrawlRate_(g.manual);return [total,positive&&total>0?g.crawl.vi/total:'',positive&&known>0?g.crawl.vi/known:'',cr,mr,laCrawlRatio_(cr,mr),positive&&total>0?known/total:''];}
  const dates=Array.from(daily.keys()).sort(),dailyRows=[],dailyDayRows=[];let left=0;
  dates.forEach((day,i)=>{
    const g=daily.get(day),r=rates(g),cutoff=new Date(Date.parse(day+'T00:00:00Z')-6*86400000).toISOString().slice(0,10);
    while(dates[left]<cutoff)left++;
    const rolling={crawl:laCrawlBucket_(),manual:laCrawlBucket_(),unknown:laCrawlBucket_()};
    for(let j=left;j<=i;j++)Object.keys(rolling).forEach(k=>{const source=daily.get(dates[j])[k];['vi','hours','timedVi','equivalentDays'].forEach(f=>rolling[k][f]+=source[f]);});
    const rr=rates(rolling);dailyRows.push([day,g.crawl.vi,g.manual.vi,g.unknown.vi,r[1],r[2],r[6],r[3],r[4],r[5],r[0],rr[1],rr[3],rr[4],rr[5]]);
    const dr=b=>b.equivalentDays>0?b.vi/b.equivalentDays:'';dailyDayRows.push([day,g.crawl.vi,g.manual.vi,g.unknown.vi,dr(g.crawl),dr(g.manual),laCrawlRatio_(dr(g.crawl),dr(g.manual)),dr(rolling.crawl),dr(rolling.manual),laCrawlRatio_(dr(rolling.crawl),dr(rolling.manual))]);
  });
  function comparisons(map){return Array.from(map.keys()).sort().reverse().map(key=>{const [month,label]=JSON.parse(key),g=map.get(key),a=g.crawl,b=g.manual,cr=laCrawlRate_(a),mr=laCrawlRate_(b);return [month,label,a.timedVi,a.hours,cr,b.timedVi,b.hours,mr,laCrawlRatio_(cr,mr),a.requests.size,b.requests.size,g.unknown.vi];});}
  const trends=Array.from(monthly.keys()).sort().reverse().map(month=>{const g=monthly.get(month),r=rates(g);return [month,r[0],g.crawl.vi,g.manual.vi,g.unknown.vi,r[1],r[2],r[6],r[3],r[4],r[5]];});
  const dayRate=g=>g.equivalentDays>0?g.vi/g.equivalentDays:'';
  const dayMonthly=[];Array.from(monthly.keys()).sort().reverse().forEach(month=>Object.keys(labels).forEach(k=>{const g=monthly.get(month)[k];dayMonthly.push([month,labels[k],g.vi,g.dayKeys.size,g.equivalentDays,dayRate(g),g.requests.size]);}));
  function dayCompare(map,pair){return Array.from(map.keys()).sort().reverse().map(key=>{const parts=pair?JSON.parse(key):[key],g=map.get(key),a=g.crawl,b=g.manual;return parts.concat([a.vi,a.equivalentDays,dayRate(a),b.vi,b.equivalentDays,dayRate(b),laCrawlRatio_(dayRate(a),dayRate(b)),g.unknown.vi]);});}
  return {dayDaily:dailyDayRows,dayMonthly:dayMonthly,dayTrends:dayCompare(monthly,false),dayWorkers:dayCompare(workers,true),dayNodes:dayCompare(nodes,true),monthly:monthlyRows,trends:trends,daily:dailyRows,workers:comparisons(workers),nodes:comparisons(nodes),reasons:Array.from(reasons.entries()).map(([key,g])=>JSON.parse(key).concat([g.vi,g.hours,g.requests.size,g.records])),today:today};
}
function laCrawlPreparation_(data,today){
  const workers=new Map(),requestDays=new Map(),audit=[];
  function worker(name){const key=laName_(name);if(!workers.has(key))workers.set(key,{name:String(name).trim(),days:new Set(),vi:0,requests:0,linkedVi:0,linked:0,invalid:0});return workers.get(key);}
  const rows=data.WORK_SESSIONS||[],h={};(rows[0]||[]).forEach((n,i)=>h[n]=i);
  rows.slice(1).forEach(r=>{
    if(String(r[h.Stage]||'').trim().toUpperCase()!=='CRAWL')return;
    const name=String(r[h.Worker]||'').trim(),id=String(r[h['Request ID']]||'').trim();
    const day=laDay_(r[h['Ended At']])||laDay_(r[h['Started At']]);
    if(!name||!day||day>today){audit.push([id,name,'CRAWL 작업자/날짜 누락 또는 미래 날짜','']);return;}
    const g=worker(name);g.days.add(day);const key=JSON.stringify([id,laName_(name)]);if(!requestDays.has(key))requestDays.set(key,new Set());requestDays.get(key).add(day);
  });
  const requests=data.REQUESTS||[],rh={};(requests[0]||[]).forEach((n,i)=>rh[n]=i);
  requests.slice(1).forEach(r=>{
    if(String(r[rh['크롤링 단계']]||'').trim().toUpperCase()!=='CRAWL_COMPLETED')return;
    const id=String(r[rh['Request ID']]||'').trim(),name=String(r[rh['크롤링 작업자']]||'').trim(),vi=r[rh['Total VI']];
    if(!name){audit.push([id,'','완료 요청의 크롤링 작업자 없음',laNumber_(vi)?vi:'']);return;}
    const g=worker(name);g.requests++;
    if(!laNumber_(vi)||vi<0){g.invalid++;audit.push([id,name,'Total VI 누락/오류','']);return;}
    g.vi+=vi;
    if(requestDays.has(JSON.stringify([id,laName_(name)]))){g.linked++;g.linkedVi+=vi;}else audit.push([id,name,'해당 작업자의 CRAWL 기록일 없음 (백업/누락 가능)',vi]);
  });
  return {rows:Array.from(workers.values()).sort((a,b)=>a.name.localeCompare(b.name)).map(g=>{
    const dates=Array.from(g.days).sort(),n=dates.length;
    return [g.name,g.vi,g.requests,n,n&&g.requests>0&&g.linked===g.requests&&!g.invalid?g.vi/n:'',g.linkedVi,g.linked,g.requests?g.linked/g.requests:'',n&&g.linked>0?g.linkedVi/n:'',dates[0]||'',dates[dates.length-1]||''];
  }),audit:audit};
}
function laCrawlMonthly_(data,today,weights){
  weights=weights||{};const groups=new Map(),latest=new Map(),rh={},sh={};
  const requests=data.REQUESTS||[],sessions=data.WORK_SESSIONS||[];
  (requests[0]||[]).forEach((n,i)=>rh[n]=i);(sessions[0]||[]).forEach((n,i)=>sh[n]=i);
  function group(month,name){const key=JSON.stringify([month,laName_(name)]);if(!groups.has(key))groups.set(key,{month:month,name:String(name).trim(),days:new Set(),vi:0,count:0,registered:0,invalid:0});return groups.get(key);}
  sessions.slice(1).forEach(r=>{
    if(String(r[sh.Stage]||'').trim().toUpperCase()!=='CRAWL')return;
    const name=String(r[sh.Worker]||'').trim(),id=String(r[sh['Request ID']]||'').trim(),day=laDay_(r[sh['Ended At']])||laDay_(r[sh['Started At']]);
    if(!name||!day||day>today)return;
    group(day.slice(0,7),name).days.add(day);const key=JSON.stringify([id,laName_(name)]);if(!latest.has(key)||latest.get(key)<day)latest.set(key,day);
  });
  requests.slice(1).forEach(r=>{
    if(String(r[rh['크롤링 단계']]||'').trim().toUpperCase()!=='CRAWL_COMPLETED')return;
    const id=String(r[rh['Request ID']]||'').trim(),name=String(r[rh['크롤링 작업자']]||'').trim()||'작업자 미확인',day=latest.get(JSON.stringify([id,laName_(name)]));
    const g=group(day?day.slice(0,7):'시점 미확인',name),vi=r[rh['Total VI']],done=r[rh['3P Completed VI']];g.count++;
    if(laNumber_(vi)&&vi>=0)g.vi+=vi;else g.invalid++;
    if(laNumber_(done)&&done>=0&&laNumber_(vi)&&done<=vi)g.registered+=done;else g.invalid++;
  });
  const entries=Array.from(groups.values()).sort((a,b)=>b.month.localeCompare(a.month)||a.name.localeCompare(b.name));
  return {entries:entries,rows:entries.map(g=>[g.month,g.name,g.vi,g.count,g.days.size,g.days.size&&!g.invalid?g.vi/g.days.size:'',g.registered,g.vi>0&&!g.invalid?g.registered/g.vi:'',g.invalid,g.days.size*(weights[laName_(g.name)]||1),g.days.size&&!g.invalid?g.vi/(g.days.size*(weights[laName_(g.name)]||1)):''])};
}
function laPlanConfig_(ss,today){
  const name='투입 계획 설정';let sheet=ss.getSheetByName(name);
  if(!sheet){sheet=ss.insertSheet(name);sheet.getRange(1,1,9,3).setValues([
    ['계획 설정(A006)','값','설명'],['기준 월',"'"+today.slice(0,7),'YYYY-MM. 등록·크롤링 속도의 표본 월'],
    ['크롤링 환산 인원',2,'예시 시작값. 1명×0.5 근무비율은 0.5명. 전담 역할별 유효 인원 입력'],
    ['등록 환산 인원','','3P 등록에 배정할 유효 인원 합계. 크롤링 인원과 역할 중복 배정 주의'],
    ['목표 작업일',5,'현재 잔량만 처리할 목표. 신규 유입·휴일·대기시간 제외'],
    ['미확정 요청당 예상 VI','','빈칸이면 기준 월 완료 크롤링 요청당 평균 사용'],
    ['크롤링 VI/환산인일 보정값','','빈칸이면 기준 월 물량/작업자별 환산 기록일 합계. 누락 시 보정'],
    ['등록 VI/환산인일 보정값','','빈칸이면 기준 월 크롤링 건의 내부 3P 기록 기준 속도'],
    ['설명','변경 후 전체 새로고침','환산인일=작업자 고유 기록일×CONFIG G:H 근무비율. 시간은 사용하지 않음']]);sheet.setFrozenRows(1);sheet.setColumnWidth(1,255);sheet.setColumnWidth(2,160);sheet.setColumnWidth(3,700);sheet.getRange('A1:C9').setWrap(true);sheet.getRange('B3:B8').setNumberFormat('0.00');}
  const rows=sheet.getDataRange().getValues();if(!rows[0]||rows[0][0]!=='계획 설정(A006)')throw Error('투입 계획 설정 탭에 기존 내용이 있습니다. 이름을 변경해 보존한 뒤 다시 실행하세요.');
  const values={};rows.slice(1).forEach(r=>values[r[0]]=r[1]);
  const month=String(values['기준 월']||'').replace(/^'/,'').trim();if(!/^\d{4}-(0[1-9]|1[0-2])$/.test(month))throw Error('투입 계획 설정: 기준 월을 YYYY-MM으로 입력하세요.');
  function number(key,integer){const value=values[key];if(value===''||value===undefined)return null;const n=Number(value);if(!Number.isFinite(n)||n<=0||(integer&&!Number.isInteger(n)))throw Error('투입 계획 설정의 양수 입력 확인: '+key);return n;}
  return {month:month,crawlers:number('크롤링 환산 인원'),listers:number('등록 환산 인원'),targetDays:number('목표 작업일',true),requestVi:number('미확정 요청당 예상 VI'),crawlRate:number('크롤링 VI/환산인일 보정값'),listingRate:number('등록 VI/환산인일 보정값')};
}
function laPlan_(data,crawl,input,today){
  const month=input.month||today.slice(0,7),monthly=crawl.monthlyPreparation.rows.filter(r=>r[0]===month);
  const days=monthly.reduce((n,r)=>n+r[9],0),vi=monthly.reduce((n,r)=>n+r[2],0),requests=monthly.reduce((n,r)=>n+r[3],0),bad=monthly.some(r=>r[8]>0);
  const crawlRate=input.crawlRate>0?input.crawlRate:days>0&&vi>0&&!bad?vi/days:null;
  const averageVi=input.requestVi>0?input.requestVi:requests>0&&vi>0&&!bad?vi/requests:null;
  const cr=crawl.dayMonthly.find(r=>r[0]===month&&r[1]==='크롤링'),mr=crawl.dayMonthly.find(r=>r[0]===month&&r[1]==='비크롤링');
  const listingRate=input.listingRate>0?input.listingRate:cr&&cr[5]>0?cr[5]:null,manualRate=mr&&mr[5]>0?mr[5]:null;
  const rh={},rows=data.REQUESTS||[];(rows[0]||[]).forEach((n,i)=>rh[n]=i);
  const backlog={ready:{count:0,vi:0},manual:{count:0,vi:0},pending:{count:0,vi:0,unknown:0},unknown:{count:0,vi:0},held:{count:0,vi:0,unknown:0}};
  rows.slice(1).forEach(r=>{
    const status=String(r[rh.Status]||'').trim().toUpperCase(),type=String(r[rh['Request Type']]||''),stage=String(r[rh['크롤링 단계']]||'').trim().toUpperCase();
    if(!String(r[rh['Request ID']]||'').trim()||['DONE','DROP','CANCELLED','CANCELED'].includes(status)||type==='2P Lite 전환')return;
    const total=r[rh['Total VI']],done=r[rh['3P Completed VI']],known=laNumber_(total)&&total>=0&&laNumber_(done)&&done>=0&&done<=total;
    if(known&&total>0&&done===total)return;
    if(status==='ON HOLD'){backlog.held.count++;if(known&&total>0)backlog.held.vi+=total-done;else backlog.held.unknown++;return;}
    if(stage.startsWith('CRAWL_')&&stage!=='CRAWL_COMPLETED'){
      if(known&&total>0&&done===total)return;
      const g=backlog.pending;g.count++;if(known&&total>0)g.vi+=total-done;else g.unknown++;return;
    }
    if(known&&done===total&&total>0)return;
    const key=stage==='CRAWL_COMPLETED'?'ready':stage==='N/A'?'manual':'unknown',g=backlog[key];
    if(known&&total>0){g.count++;g.vi+=total-done;}else backlog.unknown.count++;
  });
  const estimated=backlog.pending.unknown===0?0:averageVi?backlog.pending.unknown*averageVi:null,pendingVi=estimated!==null?backlog.pending.vi+estimated:null;
  function need(q,rate){return q===0?0:q!==null&&rate>0?q/rate:null;}
  const crawlDays=need(pendingVi,crawlRate),listingCrawlVi=pendingVi!==null?backlog.ready.vi+pendingVi:null;
  const listingCrawlDays=need(listingCrawlVi,listingRate),manualDays=need(backlog.manual.vi,manualRate),listingDays=listingCrawlDays!==null&&manualDays!==null?listingCrawlDays+manualDays:null;
  const allDays=crawlDays!==null&&listingDays!==null?crawlDays+listingDays:null;
  const crawlerElapsed=crawlDays===0?0:crawlDays!==null&&input.crawlers>0?crawlDays/input.crawlers:null;
  const listingElapsed=listingDays===0?0:listingDays!==null&&input.listers>0?listingDays/input.listers:null;
  const parallel=crawlerElapsed!==null&&listingElapsed!==null?Math.ceil(Math.max(crawlerElapsed,listingElapsed)):null,sequential=crawlerElapsed!==null&&listingElapsed!==null?Math.ceil(crawlerElapsed+listingElapsed):null;
  const crawlerFte=crawlDays===0?0:crawlDays!==null&&input.targetDays>0?crawlDays/input.targetDays:null,listerFte=listingDays===0?0:listingDays!==null&&input.targetDays>0?listingDays/input.targetDays:null;
  const old1000=need(1000,manualRate),register1000=need(1000,listingRate),prep1000=need(1000,crawlRate),new1000=register1000!==null&&prep1000!==null?register1000+prep1000:null;
  const show=x=>x===null||x===undefined?'':x;
  return {assumptions:[['기준 월',month],['크롤링 VI/환산인일',show(crawlRate)],['크롤링 건 3P VI/환산인일',show(listingRate)],['비크롤링 건 3P VI/환산인일',show(manualRate)],['요청당 예상 VI',show(averageVi)],['크롤링 환산 인원',show(input.crawlers)],['등록 환산 인원',show(input.listers)],['목표 작업일',show(input.targetDays)]],
    backlog:[['크롤링 완료·3P 잔량',backlog.ready.count,backlog.ready.vi,0],['비크롤링·3P 잔량',backlog.manual.count,backlog.manual.vi,0],['크롤링 대기',backlog.pending.count,backlog.pending.vi,backlog.pending.unknown],['분류/물량 확인 필요',backlog.unknown.count,backlog.unknown.vi,backlog.unknown.count],['ON HOLD (재개 전 예측 제외)',backlog.held.count,backlog.held.vi,backlog.held.unknown]],
    impact:[['비크롤링 등록 1,000 VI(환산인일)',show(old1000)],['크롤링 준비 1,000 VI(환산인일)',show(prep1000)],['크롤링 후 등록 1,000 VI(환산인일)',show(register1000)],['준비+등록 합산 1,000 VI(환산인일)',show(new1000)],['준비+등록 합산 VI/환산인일',new1000>0?1000/new1000:''],['1,000 VI당 절감 환산인일(추정)',old1000!==null&&new1000!==null?old1000-new1000:''],['총 투입인일 기준 추정 효율 배수',old1000>0&&new1000>0?old1000/new1000:''],['크롤링 단계 1작업일 공급 VI',crawlRate&&input.crawlers>0?crawlRate*input.crawlers:''],['등록 단계 1작업일 처리 VI',listingRate&&input.listers>0?listingRate*input.listers:''],['직렬 공정의 1작업일 처리용량(병목)',crawlRate&&listingRate&&input.crawlers>0&&input.listers>0?Math.min(crawlRate*input.crawlers,listingRate*input.listers):'']],
    forecast:[['미확정 크롤링 VI 추정분',show(estimated)],['크롤링 대기 VI(확정+추정)',show(pendingVi)],['필요 크롤링 환산인일',show(crawlDays)],['필요 3P 등록 환산인일',show(listingDays)],['준비+3P 필요 환산인일 합계',show(allDays)],['크롤링 단계 소요 작업일',show(crawlerElapsed)],['등록 단계 소요 작업일',show(listingElapsed)],['이상적 병행 처리 하한(작업일)',show(parallel)],['순차 처리 참고(작업일)',show(sequential)],['목표 작업일 내 필요 크롤링 환산 인원',show(crawlerFte)],['목표 작업일 내 필요 등록 환산 인원',show(listerFte)],['전일 전담 기준 크롤링 인원(올림)',crawlerFte!==null?Math.ceil(crawlerFte):''],['전일 전담 기준 등록 인원(올림)',listerFte!==null?Math.ceil(listerFte):''],['예측에서 제외된 요청 수',backlog.unknown.count+backlog.held.count]],
    raw:{pendingVi:pendingVi,crawlDays:crawlDays,listingDays:listingDays,allDays:allDays,parallel:parallel,sequential:sequential,crawlerFte:crawlerFte,listerFte:listerFte,old1000:old1000,new1000:new1000}};
}

function refreshListingAnalytics(){
  const lock=LockService.getDocumentLock();if(!lock.tryLock(1000))throw Error('다른 분석 갱신이 실행 중입니다.');
  const ss=SpreadsheetApp.getActiveSpreadsheet();try{
    laPrepareWeights_(ss);
    const now=new Date(),today=Utilities.formatDate(now,'Asia/Seoul','yyyy-MM-dd'),c=laConfig_(ss);c.plan=laPlanConfig_(ss,today);
    const tabs=laBuild_(laRead_(c.id),c,today);
    // Validate all input before changing reports. Last successful timestamp written LAST.
    laRenderReports_(ss,laLayouts_(tabs),Utilities.formatDate(now,'Asia/Seoul','yyyy-MM-dd HH:mm:ss'));
    // User-requested removal, only in this validated analysis destination.
    ['작업자 주별','데이터 점검'].forEach(n=>{const s=ss.getSheetByName(n);if(s)ss.deleteSheet(s);});
    const notices=tabs['데이터 점검'][1].slice(6);console.log(JSON.stringify({analyticsNotices:notices}));
    ss.getSheetByName('CONFIG').getRange('E5:F5').setValues([['집계 참고사항 수',notices.length]]);
    const cfg=ss.getSheetByName('CONFIG');cfg.getRange('E1:F3').setValues([['갱신 상태','성공'],['최종 성공(KST)',Utilities.formatDate(now,'Asia/Seoul','yyyy-MM-dd HH:mm:ss')],['버전','A006']]);
  }catch(e){const cfg=ss.getSheetByName('CONFIG');if(cfg)cfg.getRange('E1:F1').setValues([['갱신 상태','실패: '+e.message]]);throw e;}finally{lock.releaseLock();}
}
function installListingAnalyticsDaily(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();laConfig_(ss);
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='refreshListingAnalytics').forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('refreshListingAnalytics').timeBased().atHour(6).everyDays(1).inTimezone('Asia/Seoul').create();
}



