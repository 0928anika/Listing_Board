// Rebuild derived values. A PAUSED row with a blank end may recover its end
// from a valid start and stored Duration; ACTIVE rows and existing ends stay intact.
const RECONCILE_QUEUE_PREFIX_='REQ_RECONCILE_V1_';
const RECONCILE_REPORT_PREFIX_='REQ_RECONCILE_REPORT_V1_';
function reconcileKey_(id){return RECONCILE_QUEUE_PREFIX_+encodeURIComponent(String(id));}
function reconcileQueueEntry_(id){const raw=PropertiesService.getScriptProperties().getProperty(reconcileKey_(id));return raw?JSON.parse(raw):null;}
function markRequestReconcileNeeded_(ids,domains,reason){
  return withOpsDocumentLock_(function(){
  const props=PropertiesService.getScriptProperties(),updates={},tickets={},now=Date.now();
  Array.from(new Set(ids.map(String))).forEach(function(id){
    const old=reconcileQueueEntry_(id)||{requestId:id,firstAt:now,attempts:0,domains:{}};
    old.domains=Object.assign({},old.domains,domains);old.version=Utilities.getUuid();old.updatedAt=now;
    old.reason=String(reason||'event');old.state='PENDING';old.nextAt=0;
    updates[reconcileKey_(id)]=JSON.stringify(old);tickets[id]=old.version;
  });
  if(Object.keys(updates).length)props.setProperties(updates,false);
  return tickets;
  });
}
function settleRequestReconcile_(id,version,domains,report,error){
  const props=PropertiesService.getScriptProperties(),entry=reconcileQueueEntry_(id);
  if(!entry||entry.version!==version)return; // Do not acknowledge a newer mutation.
  entry.lastAttemptAt=Date.now();entry.attempts++;
  if(error){entry.state='ERROR';entry.lastError=String(error.message||error).slice(0,800);entry.nextAt=Date.now()+Math.min(3600000,15000*Math.pow(2,Math.min(entry.attempts-1,8)));props.setProperty(reconcileKey_(id),JSON.stringify(entry));return;}
  const blocked=report&&(report.workTimeBlocked||report.archived);
  Object.keys(domains).forEach(function(domain){if(domains[domain]&&!(blocked&&(domain==='sessions'||report.archived)))delete entry.domains[domain];});
  if(blocked){entry.state='REVIEW';entry.lastError=report.archived?'백업 완료 REQ: 운영 원본으로 덮어쓰지 않았습니다.':'세션 원본 이상: Work Time 갱신 보류';entry.nextAt=0;}
  else {entry.state='PENDING';entry.lastError='';entry.nextAt=0;}
  if(Object.keys(entry.domains).length)props.setProperty(reconcileKey_(id),JSON.stringify(entry));else props.deleteProperty(reconcileKey_(id));
}
function recordReconcileReport_(id,report){
  const key=RECONCILE_REPORT_PREFIX_+encodeURIComponent(id),props=PropertiesService.getScriptProperties();
  if(!report.warnings.length&&!report.archived){props.deleteProperty(key);return;}
  props.setProperty(key,JSON.stringify({requestId:id,at:Date.now(),archived:!!report.archived,workTimeBlocked:!!report.workTimeBlocked,
    warningCount:report.warnings.length,warnings:report.warnings.slice(0,15),manualSessions:report.manualSessions||0}));
}
function readReconcileSessionRecords_(ids,providedRows){
  const sheet=getRequiredSheet_('WORK_SESSIONS');assertWorkSessionViSchema_(sheet);
  const wanted=new Set(ids),records=[];
  if(providedRows){providedRows.forEach(function(values,i){if(wanted.has(String(values[2])))records.push({row:i+2,values:values.slice()});});return {sheet:sheet,records:records};}
  const last=sheet.getLastRow();if(last<2)return {sheet:sheet,records:records};
  const positions=[];
  ids.forEach(function(id){sheet.getRange(2,3,last-1,1).createTextFinder(id).matchEntireCell(true).matchCase(true).useRegularExpression(false).findAll().forEach(function(cell){positions.push(cell.getRow());});});
  const groups=[];Array.from(new Set(positions)).sort(function(a,b){return a-b;}).forEach(function(row){const last=groups[groups.length-1];if(last&&last.start+last.count===row)last.count++;else groups.push({start:row,count:1});});
  groups.forEach(function(group){sheet.getRange(group.start,1,group.count,13).getValues().forEach(function(values,i){
    if(!wanted.has(String(values[2])))throw Error('세션 행 위치가 변경되었습니다. 다시 계산합니다.');records.push({row:group.start+i,values:values});
  });});return {sheet:sheet,records:records};
}
function planReconcileSessions_(id,records,now){
  const result={requestId:id,sessions:0,endedSessions:0,activeSessions:0,manualSessions:0,durationChecked:0,durationChanged:0,endedAtRestored:0,deltaChanged:0,
    minutes:0,workTimeBlocked:false,warnings:[],patches:[]};
  function warning(record,code,message){result.warnings.push({row:record.row,sessionId:String(record.values[0]||''),code:code,message:message});}
  const date=v=>v instanceof Date&&Number.isFinite(v.getTime()),count=v=>typeof v==='number'&&Number.isFinite(v)&&v>=0&&Math.floor(v)===v;
  records.forEach(function(record){const r=record.values;if(String(r[2])!==id)return;result.sessions++;
    if(!String(r[0]||'').trim())result.manualSessions++;
    const stage=String(r[3]||''),status=String(r[4]||'');
    if(!String(r[1]||'').trim()||['3P','2P','CRAWL','BARCODE'].indexOf(stage)<0){warning(record,'IDENTITY','Worker 또는 Stage 확인 필요');result.workTimeBlocked=true;return;}
    if(status==='ACTIVE'){
      result.activeSessions++;
      if(!date(r[5])||r[5]>now||r[6]!==''&&r[6]!=null){warning(record,'ACTIVE_TIME','ACTIVE 시작 시각 또는 종료 시각 확인 필요');result.workTimeBlocked=true;return;}
      result.minutes+=(now-r[5])/60000;return; // Never write ACTIVE Duration/Delta.
    }
    if(status!=='PAUSED'){warning(record,'STATUS','확인되지 않은 종료 Status: '+status);result.workTimeBlocked=true;return;}
    result.endedSessions++;
    if((r[6]===''||r[6]==null)&&date(r[5])&&typeof r[7]==='number'&&Number.isFinite(r[7])&&r[7]>=0){
      const restored=new Date(r[5].getTime()+r[7]*60000);
      if(date(restored)&&restored>=r[5]){r[6]=restored;result.patches.push({row:record.row,col:7,value:restored});result.endedAtRestored++;}
    }
    if(!date(r[5])||!date(r[6])||r[6]<r[5]){warning(record,'TIMESTAMP',!date(r[5])?'Started At 누락/형식 오류':!date(r[6])?'Ended At 누락/형식 오류':'Ended At이 Started At보다 빠름');result.workTimeBlocked=true;return;}
    const minutes=(r[6]-r[5])/60000;result.durationChecked++;
    if(r[7]!==minutes){result.patches.push({row:record.row,col:8,value:minutes});result.durationChanged++;}
    // Sum the newly derived value, never the old stored Duration.
    result.minutes+=minutes;r[7]=minutes;
    if(stage==='3P'||stage==='2P'){
      if(count(r[10])&&count(r[11])&&!/VI 종료값 확인 불가|VI 기준값 없음/.test(String(r[9]||''))){const delta=r[11]-r[10];if(r[12]!==delta){result.patches.push({row:record.row,col:13,value:delta});result.deltaChanged++;r[12]=delta;}}
      else warning(record,'VI_UNKNOWN','Start/End Completed VI를 확정할 수 없어 기존 Delta 유지');
    }
  });return result;
}
function reconcileRequestSummary_(requestId,options){return reconcileRequestSummaries_([String(requestId)],options);}
function reconcileRequestSummaries_(requestIds,options){
  options=options||{};
  return withOpsDocumentLock_(function(){
    const rt=options.requestTable||(requestIds===null?readOpsTable_(APP_CONFIG.SHEETS.REQUESTS):readSessionRequestRecords_(requestIds)),h=rt.headers;
    const ids=Array.from(new Set((requestIds===null?Object.keys(rt.byId):requestIds).map(String))),domains=options.domains||{tasks:true,sessions:true};
    const tickets=options.tickets||markRequestReconcileNeeded_(ids,domains,options.reason||'reconcile'),reports={},eligible=[];
    const now=options.now||new Date();
    try{
      ids.forEach(function(id){const records=rt.byId[id]||[];if(records.length!==1)throw Error('요청 ID가 없거나 중복되었습니다: '+id);
        if(String(opsValue_(records[0].values,h,'Archive Status')).toUpperCase()==='ARCHIVED')reports[id]={requestId:id,archived:true,warnings:[]};else eligible.push(id);
      });
      if(domains.sessions&&eligible.length){
        if(!h['Work Time (Hours)'])throw Error('Work Time (Hours) 헤더가 없습니다.');
        const sessions=readReconcileSessionRecords_(eligible,options.sessionRows),patches=[];
        eligible.forEach(function(id){reports[id]=planReconcileSessions_(id,sessions.records,now);Array.prototype.push.apply(patches,reports[id].patches);});
        // Derived Duration is persisted before writing REQUESTS Work Time.
        if(patches.length){
          // A restored timestamp must not depend on the prior blank cell format.
          patches.filter(p=>p.col===7&&p.value instanceof Date).forEach(function(p){sessions.sheet.getRange(p.row,p.col).setNumberFormat('yyyy-mm-dd hh:mm:ss');});
          writeOpsPatches_(sessions.sheet,patches);
        }
      }
      const taskOptions=Object.assign({},options,{requestTable:rt});
      if(!taskOptions.executionContext&&!taskOptions.taskSources)taskOptions.executionContext={sidebarTaskLookup:true};
      const result=domains.tasks&&eligible.length?calculateAndWriteRequestSummaries_(eligible,taskOptions):{requests:eligible.length,sources:0,patches:0,summaries:{}};
      ids.filter(id=>reports[id]&&reports[id].archived).forEach(function(id){const values={};Object.keys(h).forEach(k=>values[k]=opsValue_(rt.byId[id][0].values,h,k));result.summaries[id]=values;});
      eligible.forEach(function(id){const record=rt.byId[id][0],report=reports[id]||(reports[id]={requestId:id,warnings:[]});
        if(domains.sessions&&!report.workTimeBlocked){const patches=[];addOpsPatch_(patches,record,h,'Work Time (Hours)',report.minutes/60);writeOpsPatches_(rt.sheet,patches);}
        report.summary=result.summaries[id]||{};report.workTimeHours=opsValue_(record.values,h,'Work Time (Hours)');
      });
      if(options.notify!==false)bumpMyWorkSidebarRevisionV2_([]);
      ids.forEach(function(id){if(domains.sessions||reports[id].archived)recordReconcileReport_(id,reports[id]);settleRequestReconcile_(id,tickets[id],domains,reports[id]);delete reports[id].patches;});
      result.reports=reports;return result;
    }catch(error){ids.forEach(function(id){try{settleRequestReconcile_(id,tickets[id],domains,null,error);}catch(ignored){}});throw error;}
  });
}
// Existing summary callers reuse their loaded Task indexes; no extra session scan.
function syncRequestSummaries_(requestIds,options){
  if(options&&options.participantsOnly)return withOpsDocumentLock_(function(){return calculateAndWriteRequestSummaries_(requestIds,options);});
  options=Object.assign({},options||{},{domains:{tasks:true},notify:false});
  return reconcileRequestSummaries_(requestIds,options);
}
function recoverPendingRequestReconciliations_(){
  const lock=LockService.getScriptLock();if(lock.hasLock()||!lock.tryLock(100))return {busy:true};
  try{
    const props=PropertiesService.getScriptProperties().getProperties(),now=Date.now();
    const entries=Object.keys(props).filter(k=>k.indexOf(RECONCILE_QUEUE_PREFIX_)===0).map(k=>JSON.parse(props[k]))
      .filter(e=>e.state!=='REVIEW'&&Number(e.nextAt||0)<=now).sort((a,b)=>(a.lastAttemptAt||0)-(b.lastAttemptAt||0));
    if(!entries.length)return {processed:0};const entry=entries[0],tickets={};tickets[entry.requestId]=entry.version;
    try{reconcileRequestSummary_(entry.requestId,{domains:entry.domains,tickets:tickets,reason:'recovery'});return {processed:1};}
    catch(error){const latest=reconcileQueueEntry_(entry.requestId);if(latest&&latest.lastAttemptAt===entry.lastAttemptAt)settleRequestReconcile_(entry.requestId,entry.version,entry.domains,null,error);console.error('REQ reconcile recovery: '+String(error));return {processed:1,pending:true};}
  }finally{lock.releaseLock();}
}
function recalculateSelectedRequest(){
  assertDashboardAdmin_();const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getActiveSheet(),range=ss.getActiveRange();
  if(sheet.getName()!==APP_CONFIG.SHEETS.REQUESTS||!range||range.getRow()<2||range.getNumRows()!==1)throw Error('REQUESTS에서 재계산할 REQ 한 행을 선택하세요.');
  const h=getHeaderMap_(sheet);if(!h['Request ID'])throw Error('Request ID 헤더가 없습니다.');
  const id=String(sheet.getRange(range.getRow(),h['Request ID']).getValue()||'').trim();if(!id)throw Error('선택 행에 Request ID가 없습니다.');
  const result=reconcileRequestSummary_(id,{reason:'admin'}),r=result.reports[id],s=r.summary||{};
  const lines=[id+(r.archived?' 재계산 보류 (백업 완료)':r.workTimeBlocked?' 재계산 완료 · Work Time 갱신 보류':' 재계산 완료')];
  if(!r.archived){lines.push('종료 세션 '+r.endedSessions+'건 / ACTIVE '+r.activeSessions+'건','Duration '+r.durationChecked+'건 확인 · '+r.durationChanged+'건 수정',
    'Ended At 복원 '+r.endedAtRestored+'건','Session ID 없는 세션 '+r.manualSessions+'건 (다른 세션과 동일한 검증 적용)','Work Time: '+(r.workTimeBlocked?'기존 값 유지':r.minutes+'분'),
    'Total VI: '+s['Total VI'],'3P Completed VI: '+s['3P Completed VI'],'2P Completed VI: '+s['2P Completed VI'],'경고: '+r.warnings.length+'건');}
  r.warnings.slice(0,20).forEach(function(w){lines.push('WORK_SESSIONS '+w.row+'행 / '+(w.sessionId||'Session ID 없음')+' : '+w.message);});
  if(r.warnings.length>20)lines.push('외 '+(r.warnings.length-20)+'건. 반환 결과 또는 재계산 대기·오류 확인에서 확인하세요.');
  SpreadsheetApp.getUi().alert(lines.join('\n'));return result;
}
function getRequestReconcileDiagnostics(){
  assertDashboardAdmin_();const props=PropertiesService.getScriptProperties().getProperties(),pending=[],warnings=[];
  Object.keys(props).forEach(function(key){if(key.indexOf(RECONCILE_QUEUE_PREFIX_)===0)pending.push(JSON.parse(props[key]));else if(key.indexOf(RECONCILE_REPORT_PREFIX_)===0)warnings.push(JSON.parse(props[key]));});
  pending.sort((a,b)=>a.firstAt-b.firstAt);return {pending:pending,warnings:warnings};
}
function showRequestReconcileDiagnostics(){
  const data=getRequestReconcileDiagnostics(),now=Date.now(),lines=['재계산 대기/오류 '+data.pending.length+'건 · 경고 기록 '+data.warnings.length+'건'];
  data.pending.slice(0,30).forEach(function(e){lines.push(e.requestId+' | '+e.state+' | '+Math.floor((now-e.firstAt)/60000)+'분 경과 | 시도 '+e.attempts+'회 | '+(e.lastError||e.reason));});
  data.warnings.slice(0,20).forEach(function(r){lines.push(r.requestId+' | 경고 '+r.warningCount+'건'+(r.archived?' | 백업 완료':''));r.warnings.slice(0,3).forEach(w=>lines.push('  '+w.row+'행: '+w.message));});
  SpreadsheetApp.getUi().alert(lines.join('\n'));return data;
}
