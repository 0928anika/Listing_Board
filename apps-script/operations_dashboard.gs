// Read-only dashboard. Never calls Task readers, writers, repair or view rebuild.
function showOperationsDashboard(){
  assertDashboardAccess_();
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('OperationsDashboard').setWidth(1250).setHeight(850),'Operations Dashboard');
}
function openDashboardAssignmentV1(){
  assertDashboardAdmin_();
  return showPersonalKanbanDialog();
}
function normalizeDashboardBarcodeProvided_(value){
  const normalized=String(value||'').trim().toUpperCase();
  return normalized==='YES'||normalized==='NO'?normalized:'UNKNOWN';
}
const DASHBOARD_LONG_TERM_SESSION_COLUMNS_V1=['Aggregate ID','Status','Request ID','Worker','Stage','Work Date','Duration Minutes','Completed VI Delta','Session Count','First Started At','Last Ended At','Active Elapsed Minutes','Request Active Elapsed Minutes','Request Completed At','Aggregated At','Archive Batch ID'];
function getOperationsDashboard(payload){
  const access=assertDashboardAccess_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = 'Asia/Seoul';
  const now = new Date();
  const key=function(d){return d instanceof Date&&Number.isFinite(d.getTime())?Utilities.formatDate(d,tz,'yyyy-MM-dd'):'';};
  const today=key(now),p=payload||{},from=p.from||key(new Date(now.getTime()-29*86400000)),to=p.to||today;
  function valid(d){return /^\d{4}-\d{2}-\d{2}$/.test(d)&&!isNaN(Date.parse(d))&&new Date(d).toISOString().slice(0,10)===d;}
  if(!valid(from)||!valid(to)||from>to||Date.parse(to)-Date.parse(from)>366*86400000)throw Error('조회 기간은 시작일 이후 최대 366일로 선택하세요.');
  const rt=readOpsTable_(APP_CONFIG.SHEETS.REQUESTS),h=rt.headers;
  ['Request ID','Node','Status','Requested At','Completed At','Request Type','Total VI','3P Completed VI','2P Completed VI'].forEach(function(n){if(!h[n])throw Error('REQUESTS 필수 헤더 누락: '+n);});
  const barcodeProvidedReady=!!h['Barcode Provided'];
  const requests=Object.keys(rt.byId).map(function(id){if(rt.byId[id].length!==1)throw Error('중복 Request ID: '+id);const r=rt.byId[id][0].values,v=function(n){return opsValue_(r,h,n);};
    const start=v('Requested At'),end=v('Completed At');
    return {id:id,node:String(v('Node')||'미지정'),brand:String(v('Brand')||''),type:String(v('Request Type')),status:normalizeRequestStatusV2_(v('Status'))||APP_CONFIG.STATUS.TO_DO,
      allocationStatus:String(v('Allocation Status')||''),total:Number(v('Total VI'))||0,done3:Number(v('3P Completed VI'))||0,done2:Number(v('2P Completed VI'))||0,
      barcodeProvided:normalizeDashboardBarcodeProvided_(v('Barcode Provided')),requested:key(start),completed:key(end),
      lead:start instanceof Date&&end instanceof Date&&end>=start?(end-start)/3600000:null};});
  const sheet=ss.getSheetByName('WORK_SESSIONS'),sessions=[];let sessionReady=false,longTermReady=true;
  if(sheet){const values=sheet.getRange(1,1,Math.max(1,sheet.getLastRow()),13).getValues();sessionReady=WORK_SESSION_COLUMNS_.every(function(n,i){return values[0][i]===n;});
    if(sessionReady)values.slice(1).forEach(function(r){if(!r[0]&&!r[1])return;sessions.push({worker:String(r[1]||''),requestId:String(r[2]||''),stage:String(r[3]||''),status:String(r[4]||''),started:r[5] instanceof Date?r[5].getTime():null,endedAt:r[6] instanceof Date?r[6].getTime():null,ended:key(r[6]),minutes:typeof r[7]==='number'&&Number.isFinite(r[7])?Math.max(0,r[7]):null,delta:typeof r[12]==='number'&&Number.isFinite(r[12])?r[12]:null,sessionCount:1,aggregate:false});});}
  const longTerm=ss.getSheetByName('LONG_TERM_SESSIONS');
  if(longTerm){const values=longTerm.getRange(1,1,Math.max(1,longTerm.getLastRow()),DASHBOARD_LONG_TERM_SESSION_COLUMNS_V1.length).getValues();longTermReady=DASHBOARD_LONG_TERM_SESSION_COLUMNS_V1.every(function(n,i){return values[0][i]===n;});
    if(longTermReady)values.slice(1).forEach(function(r){if(r[1]!=='COMMITTED')return;sessions.push({worker:String(r[3]||''),requestId:String(r[2]||''),stage:String(r[4]||''),status:'PAUSED',started:r[9] instanceof Date?r[9].getTime():null,endedAt:r[10] instanceof Date?r[10].getTime():null,ended:String(r[5]||''),minutes:typeof r[6]==='number'&&Number.isFinite(r[6])?Math.max(0,r[6]):null,delta:typeof r[7]==='number'&&Number.isFinite(r[7])?r[7]:null,sessionCount:Math.max(1,Number(r[8])||1),aggregate:true,activeMinutes:typeof r[11]==='number'?r[11]:null,requestActiveMinutes:typeof r[12]==='number'?r[12]:null});});}
  return buildOperationsDashboard_(requests,sessions,{from:from,to:to,node:String(p.node||''),worker:String(p.worker||''),today:today,timezone:tz,updatedAt:Utilities.formatDate(now,tz,'yyyy-MM-dd HH:mm:ss'),sessionReady:sessionReady,longTermReady:longTermReady,barcodeProvidedReady:barcodeProvidedReady,canManageAssignments:!!(access&&typeof access==='object'&&access.isAdmin===true)});
}
// Pure aggregation, independent of Spreadsheet services.
function buildOperationsDashboard_(requests,sessions,o){
  requests=requests.map(r=>Object.assign({},r,{status:normalizeRequestStatusV2_(r.status)||APP_CONFIG.STATUS.TO_DO}));
  const inPeriod=d=>!!d&&d>=o.from&&d<=o.to,byId=Object.create(null),nodes=new Set(),workers=new Set(),weekly=Object.create(null),groups=Object.create(null),daily=Object.create(null);
  const warnings={missingRequestDate:0,missingLead:0,unknownDelta:0,orphanSessions:0,activeSessions:0},globalWarnings={missingRequestDate:0,missingLead:0,unknownDelta:0,orphanSessions:0,activeSessions:0};
  requests.forEach(r=>{byId[r.id]=r;nodes.add(r.node);});sessions.forEach(s=>{if(s.worker)workers.add(s.worker);});
  const selected=requests.filter(r=>r.status!==APP_CONFIG.STATUS.DROP&&(!o.node||r.node===o.node));
  function week(d){const date=new Date(d+'T00:00:00Z');date.setUTCDate(date.getUTCDate()-date.getUTCDay());return date.toISOString().slice(0,10);}
  function stats(a){a=a.slice().sort((x,y)=>x-y);return {count:a.length,average:a.length?a.reduce((s,n)=>s+n,0)/a.length:null,median:a.length?(a[Math.floor((a.length-1)/2)]+a[Math.ceil((a.length-1)/2)])/2:null};}
  const leads=[],backlog=[];let received=0,completed=0,blocked=0,waiting=0,inProgress=0,unassigned=0;
  selected.forEach(r=>{
    if(!groups[r.node])groups[r.node]={node:r.node,received:0,periodCompleted:0,currentTotal:0,currentDone:0,waiting:0,inProgress:0,unassigned:0,open:0,blocked:0,total3:0,done3:0,total2:0,done2:0,leads:[]};const g=groups[r.node];
    g.currentTotal++;if(r.status===APP_CONFIG.STATUS.DONE)g.currentDone++;if(r.status===APP_CONFIG.STATUS.TO_DO){waiting++;g.waiting++;}if(r.status===APP_CONFIG.STATUS.ONGOING){inProgress++;g.inProgress++;}if(r.status!==APP_CONFIG.STATUS.DONE&&r.allocationStatus==='분배 필요'){unassigned++;g.unassigned++;}
    if(!r.requested){warnings.missingRequestDate++;globalWarnings.missingRequestDate++;}
    if(inPeriod(r.requested)){received++;g.received++;const w=week(r.requested),k=w+'\n'+r.node;if(!weekly[k])weekly[k]={week:w,node:r.node,count:0};weekly[k].count++;}
    if(r.status===APP_CONFIG.STATUS.DONE&&inPeriod(r.completed)){completed++;g.periodCompleted++;if(r.lead!==null&&Number.isFinite(r.lead)){leads.push(r.lead);g.leads.push(r.lead);}else{warnings.missingLead++;globalWarnings.missingLead++;}}
    if(r.status===APP_CONFIG.STATUS.DONE&&!r.completed){warnings.missingLead++;globalWarnings.missingLead++;}
    if(r.status!==APP_CONFIG.STATUS.DONE){
      g.open++;if(r.status===APP_CONFIG.STATUS.ON_HOLD){blocked++;g.blocked++;}
      if(r.type!=='2P Lite 전환'){g.total3+=r.total;g.done3+=r.done3;}
      if(r.type!=='3P Listing'){g.total2+=r.total;g.done2+=r.done2;}
      backlog.push({id:r.id,node:r.node,brand:r.brand,status:r.status,age:r.age===null||r.age===undefined?(r.requested?Math.max(0,(Date.parse(o.today)-Date.parse(r.requested))/86400000):null):r.age});
    }
  });
  sessions.forEach(s=>{
    if(s.stage!=='3P'&&s.stage!=='2P')return;
    const r=byId[s.requestId];if(o.node&&(!r||r.node!==o.node))return;
    if(s.status==='ACTIVE')globalWarnings.activeSessions++;
    else if(inPeriod(s.ended)){
      if(!r)globalWarnings.orphanSessions++;
      if(s.delta===null||!s.worker)globalWarnings.unknownDelta++;
    }
    if(o.worker&&s.worker!==o.worker)return;
    if(s.status==='ACTIVE'){warnings.activeSessions++;return;}
    if(!inPeriod(s.ended))return;
    if(!r)warnings.orphanSessions++;
    if(s.delta===null||!s.worker){warnings.unknownDelta++;return;}
    const k=s.ended+'\n'+s.worker;if(!daily[k])daily[k]={day:s.ended,worker:s.worker,threeP:0,twoP:0,hours3P:0,hours2P:0,sessions:0};const sk=s.stage==='3P'?'threeP':'twoP';daily[k][sk]+=s.delta;if(Number.isFinite(s.minutes))daily[k][s.stage==='3P'?'hours3P':'hours2P']+=s.minutes/60;daily[k].sessions+=s.sessionCount||1;
  });
  const base={options:o,nodes:Array.from(nodes).sort(),workers:Array.from(workers).sort(),warnings:warnings,globalWarnings:globalWarnings,
    kpi:{received:received,completed:completed,open:backlog.length,blocked:blocked,waiting:waiting,inProgress:inProgress,unassigned:unassigned,lead:stats(leads)},
    daily:Object.keys(daily).sort().map(k=>daily[k]),weekly:Object.keys(weekly).sort().map(k=>weekly[k]),
    nodeRows:Object.keys(groups).sort().map(k=>{const g=groups[k];g.lead=stats(g.leads);delete g.leads;return g;}),backlog:backlog.sort((a,b)=>(b.age===null?-1:b.age)-(a.age===null?-1:a.age)).slice(0,20)};
  base.analytics=buildOperationalAnalytics_(requests,sessions,o,stats);
  base.analyticsOverall=o.worker?buildOperationalAnalytics_(requests,sessions,Object.assign({},o,{worker:''}),stats):base.analytics;
  return base;
}

// Analytics use only REQUESTS + finalized WORK_SESSIONS already read above.
function buildOperationalAnalytics_(requests,sessions,o,stats){
  const byId=Object.create(null);requests.forEach(r=>byId[r.id]=r);
  const sizeBuckets=['1–50','51–200','201–500','501–1,000','1,001–2,000','2,001–4,000','4,001+'];
  const bucket=v=>v<=50?sizeBuckets[0]:v<=200?sizeBuckets[1]:v<=500?sizeBuckets[2]:v<=1000?sizeBuckets[3]:v<=2000?sizeBuckets[4]:v<=4000?sizeBuckets[5]:sizeBuckets[6],inPeriod=d=>!!d&&d>=o.from&&d<=o.to,brandName=r=>{const value=String(r&&r.brand||'').trim();return value||'미지정';};
  const usable=sessions.filter(s=>(s.stage==='3P'||s.stage==='2P')&&s.status!=='ACTIVE'&&inPeriod(s.ended)&&s.delta!==null&&Number.isFinite(s.minutes)&&s.worker&&(!o.worker||s.worker===o.worker)&&(!o.node||byId[s.requestId]&&byId[s.requestId].node===o.node));
  const worker=Object.create(null),condition=Object.create(null),node=Object.create(null),brandMap=Object.create(null),difficultyCohorts=Object.create(null),stage={threeP:{vi:0,hours:0},twoP:{vi:0,hours:0}};
  function brandRow(name){return brandMap[name]||(brandMap[name]={brand:name,completedVI:0,workerHours:0,threeP:{vi:0,hours:0},twoP:{vi:0,hours:0},requestIds:Object.create(null),sessionCount:0,workers:Object.create(null),leads:[],completedRequestCount:0,difficultyWeighted:0,comparisonCohorts:0,comparisonVI:0});}
  usable.forEach(s=>{const r=byId[s.requestId],sk=s.stage==='3P'?'threeP':'twoP',h=s.minutes/60,w=worker[s.worker]||(worker[s.worker]={worker:s.worker,threeP:{vi:0,hours:0},twoP:{vi:0,hours:0},requests:Object.create(null)});w[sk].vi+=s.delta;w[sk].hours+=h;if(r)w.requests[r.id]=r.total;
    stage[sk].vi+=s.delta;stage[sk].hours+=h;if(r){const nk=r.node,nodeRow=node[nk]||(node[nk]={node:nk,vi:0,hours:0,sessions:0});nodeRow.vi+=s.delta;nodeRow.hours+=h;nodeRow.sessions++;
      const brand=String(r.brand||''),ck=s.stage+'\n'+brand+'\n'+bucket(r.total),c=condition[ck]||(condition[ck]={stage:s.stage,brand:brand,bucket:bucket(r.total),workers:Object.create(null)}),cw=c.workers[s.worker]||(c.workers[s.worker]={vi:0,hours:0});cw.vi+=s.delta;cw.hours+=h;
      const bn=brandName(r),br=brandRow(bn);br.completedVI+=s.delta;br.workerHours+=h;br[sk].vi+=s.delta;br[sk].hours+=h;br.requestIds[r.id]=true;br.sessionCount+=s.sessionCount||1;br.workers[s.worker]=true;
      const dk=JSON.stringify([s.worker,s.stage,bucket(r.total)]),dc=difficultyCohorts[dk]||(difficultyCohorts[dk]={brands:Object.create(null)}),db=dc.brands[bn]||(dc.brands[bn]={vi:0,hours:0});db.vi+=s.delta;db.hours+=h;}});
  const rate=x=>x.hours>0?x.vi/x.hours:null,median=a=>stats(a).median,cv=a=>{if(a.length<2)return null;const mean=a.reduce((s,n)=>s+n,0)/a.length;if(mean<=0)return null;return Math.sqrt(a.reduce((s,n)=>s+(n-mean)*(n-mean),0)/a.length)/mean;};
  Object.keys(difficultyCohorts).forEach(function(key){
    const cohort=difficultyCohorts[key],comparable=Object.keys(cohort.brands).map(function(name){const value=cohort.brands[name];return {brand:name,vi:value.vi,rate:rate(value)};}).filter(function(row){return row.rate!==null&&row.rate>0;});
    if(comparable.length<2)return;
    const cohortMedian=median(comparable.map(function(row){return row.rate;}));
    comparable.forEach(function(row){const br=brandRow(row.brand),ratio=cohortMedian/row.rate*100;br.difficultyWeighted+=ratio*row.vi;br.comparisonCohorts++;br.comparisonVI+=row.vi;});
  });
  requests.forEach(function(r){
    if(r.status!==APP_CONFIG.STATUS.DONE||r.status===APP_CONFIG.STATUS.DROP||!inPeriod(r.completed)||(o.node&&r.node!==o.node))return;
    const br=brandRow(brandName(r));br.completedRequestCount++;if(r.lead!==null&&r.lead!==undefined&&Number.isFinite(r.lead))br.leads.push(r.lead);
  });
  const barcodeGroups={YES:{key:'YES',vi:0,hours:0,sessionCount:0,requestIds:Object.create(null)},NO:{key:'NO',vi:0,hours:0,sessionCount:0,requestIds:Object.create(null)}},barcodeUnknown={sessionCount:0,requestIds:Object.create(null)};
  usable.forEach(function(s){
    if(s.stage!=='2P')return;
    const r=byId[s.requestId],provided=normalizeDashboardBarcodeProvided_(r&&r.barcodeProvided);
    if(provided==='UNKNOWN'){
      barcodeUnknown.sessionCount+=s.sessionCount||1;barcodeUnknown.requestIds[s.requestId]=true;return;
    }
    const g=barcodeGroups[provided];g.vi+=s.delta;g.hours+=s.minutes/60;g.sessionCount+=s.sessionCount||1;g.requestIds[s.requestId]=true;
  });
  const barcodeRows=['YES','NO'].map(function(key){const g=barcodeGroups[key];return {key:key,vi:g.vi,hours:g.hours,rate:rate(g),requestCount:Object.keys(g.requestIds).length,sessionCount:g.sessionCount};});
  const barcodeYes=barcodeRows[0],barcodeNo=barcodeRows[1],barcodeAbsolute=barcodeYes.rate!==null&&barcodeNo.rate!==null?barcodeYes.rate-barcodeNo.rate:null;
  const barcodeProductivity={groups:barcodeRows,difference:{absoluteViPerHour:barcodeAbsolute,relative:barcodeAbsolute!==null&&barcodeNo.rate>0?barcodeYes.rate/barcodeNo.rate-1:null},unknownExcluded:{requestCount:Object.keys(barcodeUnknown.requestIds).length,sessionCount:barcodeUnknown.sessionCount}};
  const workers=Object.keys(worker).map(name=>{const w=worker[name],ids=Object.keys(w.requests),r3=rate(w.threeP),r2=rate(w.twoP);return {worker:name,completedVI:w.threeP.vi+w.twoP.vi,workHours:w.threeP.hours+w.twoP.hours,viPerHour:rate({vi:w.threeP.vi+w.twoP.vi,hours:w.threeP.hours+w.twoP.hours}),threeP:Object.assign({},w.threeP,{rate:r3}),twoP:Object.assign({},w.twoP,{rate:r2}),requestCount:ids.length,averageRequestVI:ids.length?ids.reduce((s,id)=>s+w.requests[id],0)/ids.length:null};});
  ['threeP','twoP'].forEach(sk=>{const rates=workers.map(w=>w[sk].rate).filter(v=>v!==null),med=median(rates);workers.forEach(w=>w[sk].medianDiff=w[sk].rate!==null&&med>0?(w[sk].rate/med-1):null);stage[sk].rate=rate(stage[sk]);stage[sk].median=med;stage[sk].cv=cv(rates);stage[sk].workerCount=rates.length;});
  const requestSessions=Object.create(null),periodRequestIds=Object.create(null);usable.forEach(s=>periodRequestIds[s.requestId]=true);sessions.filter(s=>(s.stage==='3P'||s.stage==='2P')&&s.status!=='ACTIVE'&&s.worker&&Number.isFinite(s.minutes)).forEach(s=>{const q=requestSessions[s.requestId]||(requestSessions[s.requestId]={workers:Object.create(null),hours:0,vi:0,intervals:[],aggregateActiveMinutes:0});q.workers[s.worker]=true;q.hours+=s.minutes/60;if(s.delta!==null)q.vi+=s.delta;if(s.aggregate&&Number.isFinite(s.requestActiveMinutes))q.aggregateActiveMinutes+=s.requestActiveMinutes;else if(s.started!==null&&s.endedAt!==null&&s.endedAt>=s.started)q.intervals.push([s.started,s.endedAt]);});
  function unionHours(a){a=a.slice().sort((x,y)=>x[0]-y[0]);let total=0,start=null,end=null;a.forEach(x=>{if(start===null){start=x[0];end=x[1];}else if(x[0]<=end)end=Math.max(end,x[1]);else{total+=end-start;start=x[0];end=x[1];}});return start===null?null:(total+end-start)/3600000;}
  const reqs=requests.filter(r=>r.status!==APP_CONFIG.STATUS.DROP&&(!o.node||r.node===o.node)&&(inPeriod(r.completed)||inPeriod(r.requested)||periodRequestIds[r.id])).map(r=>{const s=requestSessions[r.id]||{workers:{},hours:0,vi:0,intervals:[],aggregateActiveMinutes:0},wc=Object.keys(s.workers).length,rawActive=unionHours(s.intervals),active=(rawActive===null&&!s.aggregateActiveMinutes)?null:(rawActive||0)+s.aggregateActiveMinutes/60;return {id:r.id,node:r.node,type:r.type,bucket:bucket(r.total),totalVI:r.total,leadHours:r.lead,workerCount:wc,workerGroup:wc>=3?'3+':String(wc),workerHours:s.hours||null,activeWorkHours:active,completedVI:s.vi,viPerWorkerHour:s.hours>0?s.vi/s.hours:null};});
  function groupRows(field,predicate,order){const m=Object.create(null);reqs.filter(predicate||function(){return true;}).forEach(r=>{const k=r[field],g=m[k]||(m[k]={group:k,count:0,leads:[],hours:0,vi:0});g.count++;if(r.leadHours!==null)g.leads.push(r.leadHours);if(r.workerHours!==null){g.hours+=r.workerHours;g.vi+=r.completedVI;}});return Object.keys(m).sort(function(a,b){return order?order.indexOf(a)-order.indexOf(b):a.localeCompare(b);}).map(k=>{const g=m[k],z=stats(g.leads);return {group:g.group,count:g.count,medianLead:z.median,averageLead:z.average,viPerWorkerHour:g.hours>0?g.vi/g.hours:null,workerHours:g.hours};});}
  const benchmarks=[];Object.keys(condition).forEach(k=>{const c=condition[k],rows=Object.keys(c.workers).map(w=>({worker:w,rate:rate(c.workers[w]),hours:c.workers[w].hours})).filter(x=>x.rate!==null),med=median(rows.map(x=>x.rate));rows.forEach(x=>benchmarks.push({worker:x.worker,stage:c.stage,brand:c.brand,bucket:c.bucket,rate:x.rate,median:med,diff:med>0?x.rate/med-1:null,hours:x.hours,sampleWorkers:rows.length}));});
  const brands=Object.keys(brandMap).sort().map(function(name){const br=brandMap[name],lead=stats(br.leads);return {brand:br.brand,completedVI:br.completedVI,workerHours:br.workerHours,viPerWorkerHour:rate({vi:br.completedVI,hours:br.workerHours}),threeP:{vi:br.threeP.vi,hours:br.threeP.hours,rate:rate(br.threeP)},twoP:{vi:br.twoP.vi,hours:br.twoP.hours,rate:rate(br.twoP)},requestCount:Object.keys(br.requestIds).length,sessionCount:br.sessionCount,workerCount:Object.keys(br.workers).length,medianLeadHours:lead.median,completedRequestCount:br.completedRequestCount,difficultyIndex:br.comparisonVI>0?br.difficultyWeighted/br.comparisonVI:null,comparisonCohorts:br.comparisonCohorts,comparisonVI:br.comparisonVI};});
  return {definitions:{currentWorkloadAvailable:false,currentWorkloadReason:'REQUESTS에는 작업자별 stage 배정 수량이 없습니다. Task 전체 scan이나 균등 추정 없이 계산할 수 없습니다.',productivity:'확정 Completed VI Delta 합계 ÷ 동일 확정 세션 Duration 합계',barcodeProductivity:'Barcode Provided가 YES/NO인 종료 2P 세션의 순 Completed VI Delta ÷ Worker-hours',requestEfficiency:'확정 세션 순 VI ÷ Worker-hours',brandDifficulty:'난이도 지수(100=동일 조건 기준)는 같은 작업자·Stage·REQ 규모 구간에서 브랜드별 확정 VI/hour를 비교해 구간 중앙값 ÷ 브랜드 생산성 × 100을 해당 브랜드 비교 VI로 가중 평균합니다. 100보다 높을수록 동일 조건 대비 생산성이 낮아 더 어려운 것으로 해석합니다. 비교 가능한 양의 생산성 브랜드가 2개 미만이면 제외하며, 비교 구간이 없으면 null입니다. 현재 REQUESTS 메타데이터를 기간 내 세션에 결합한 관찰 지표이므로 소표본·메타데이터 변경의 영향을 받으며 인과적 난이도를 뜻하지 않습니다.',causalWarning:'작업자 수별 결과는 관찰 비교이며 인력 추가의 인과 효과가 아닙니다.',barcodeCausalWarning:'바코드 제공 여부별 차이는 관찰 비교이며 Node·난이도·작업자 구성의 영향을 받을 수 있습니다.'},team:{vi:stage.threeP.vi+stage.twoP.vi,hours:stage.threeP.hours+stage.twoP.hours,rate:rate({vi:stage.threeP.vi+stage.twoP.vi,hours:stage.threeP.hours+stage.twoP.hours}),stages:stage},workers:workers.sort((a,b)=>(b.viPerHour||-Infinity)-(a.viPerHour||-Infinity)),requests:reqs,byWorkerCount:groupRows('workerGroup',function(r){return r.workerCount>0;},['1','2','3+']),bySize:groupRows('bucket',null,sizeBuckets),byNode:groupRows('node'),byType:groupRows('type'),nodes:Object.keys(node).sort().map(k=>Object.assign(node[k],{rate:rate(node[k])})),brands:brands,benchmarks:benchmarks,barcodeProductivity:barcodeProductivity};
}

