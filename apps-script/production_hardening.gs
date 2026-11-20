// Production helpers. Raw Tasks remain authoritative; no background repair.
function withOpsDocumentLock_(fn) {
  const lock=LockService.getScriptLock();
  if(typeof lock.hasLock==='function'&&lock.hasLock())return fn();
  if(!lock.tryLock(5000))throw Error('다른 변경 처리 중입니다. 잠시 후 새로고침하거나 요청 집계를 다시 실행하세요.');
  try{return fn();}finally{lock.releaseLock();}
}
function readOpsTable_(name) {
  const sheet=getRequiredSheet_(name),values=sheet.getDataRange().getValues(),headers={},byId=Object.create(null);
  (values[0]||[]).forEach(function(v,i){if(v)headers[String(v)]=i+1;});
  values.slice(1).forEach(function(row,i){const id=String(row[(headers['Request ID']||1)-1]||'');if(id){if(!byId[id])byId[id]=[];byId[id].push({row:i+2,values:row});}});
  return {sheet:sheet,headers:headers,rows:values.slice(1),byId:byId};
}
function opsValue_(row,h,name){return h[name]?row[h[name]-1]:'';}
function taskAdapter_(h) {
  if(!h['Request ID']||!h['2P Completed']&&!h['3P Completed'])throw Error('지원하지 않는 Task 헤더입니다. 원본을 수정하지 않았습니다.');
  return {id:h['Request ID']-1,a3:(h['3P Assignee']||0)-1,d3:(h['3P Completed']||0)-1,a2:(h['2P Assignee']||0)-1,d2:(h['2P Completed']||0)-1,work:(h['Work Status']||0)-1};
}
function readTaskSources_(names) {
  const sources={};
  names.filter(function(n,i){return names.indexOf(n)===i;}).forEach(function(n){sources[n]=readOpsTable_(n);sources[n].adapter=taskAdapter_(sources[n].headers);});
  return sources;
}
function emptyTaskAggregate_(){return {rawTotal:0,rawMissing3:false,rawMissing2:false,total:0,done3:0,done2:0,both:0,missing3:false,missing2:false,participants:[],workers:Object.create(null),valid3:[],valid2:[],pending3:[],pending2:[]};}
// One pass per source, shared INVALID rule for summaries and worker metrics.
function buildTaskIndex_(sources) {
  const index={};
  Object.keys(sources).forEach(function(name){
    const t=sources[name],a=t.adapter,map=Object.create(null);index[name]=map;
    t.rows.forEach(function(row){
      const id=String(row[a.id]||'');if(!id)return;
      const m=map[id]||(map[id]=emptyTaskAggregate_()),d3=a.d3>=0&&row[a.d3]===true,d2=a.d2>=0&&row[a.d2]===true;
      m.rawTotal++;
      [[a.a3,'rawMissing3'],[a.a2,'rawMissing2']].forEach(function(stage){if(stage[0]>=0&&!splitWorkersV2_(row[stage[0]]).length)m[stage[1]]=true;});
      if(!isViValidV3_(row,t.headers)){[a.a3,a.a2].forEach(function(col){if(col<0)return;splitWorkersV2_(row[col]).forEach(function(worker){if(m.participants.indexOf(worker)<0)m.participants.push(worker);});});return;}
      m.total++;if(d3)m.done3++;if(d2)m.done2++;if(d3&&d2)m.both++;
      [[a.a3,d3,'3P'],[a.a2,d2,'2P']].forEach(function(stage){
        if(stage[0]<0)return;const names=splitWorkersV2_(row[stage[0]]);
        const key=stage[2]==='3P'?'3':'2';names.forEach(function(worker){if(m['valid'+key].indexOf(worker)<0)m['valid'+key].push(worker);if(!stage[1]&&m['pending'+key].indexOf(worker)<0)m['pending'+key].push(worker);});
        if(!stage[1]&&!names.length)m[stage[2]==='3P'?'missing3':'missing2']=true;
        names.forEach(function(worker){
          const w=m.workers[worker]||(m.workers[worker]={assigned3P:0,done3P:0,assigned2P:0,done2P:0});
          w['assigned'+stage[2]]++;if(stage[1]){w['done'+stage[2]]++;if(m.participants.indexOf(worker)<0)m.participants.push(worker);}
        });
      });
    });
  });return index;
}
function opsSourceForRequest_(row,h) {
  return opsValue_(row,h,'Request Type')==='2P Lite 전환'?APP_CONFIG.SHEETS.TASKS_2P:
    opsValue_(row,h,'크롤링 필요')===true?APP_CONFIG.SHEETS.CRAWL_TASKS:APP_CONFIG.SHEETS.TASKS;
}
function aggregateRequestMetrics_(row,h,m) {
  const type=String(opsValue_(row,h,'Request Type')||'3P Listing'),old=normalizeRequestStatusV2_(opsValue_(row,h,'Status'))||APP_CONFIG.STATUS.TO_DO;
  const has3=type!=='2P Lite 전환',has2=type!=='3P Listing';
  const done=has3&&has2?m.both:has3?m.done3:m.done2;
  const finished=m.total>0&&done===m.total;
  const status=old===APP_CONFIG.STATUS.DROP||old===APP_CONFIG.STATUS.ON_HOLD?old:old===APP_CONFIG.STATUS.DONE&&finished?APP_CONFIG.STATUS.DONE:
    (m.done3||m.done2||opsValue_(row,h,'Started At')||old===APP_CONFIG.STATUS.ONGOING?APP_CONFIG.STATUS.ONGOING:APP_CONFIG.STATUS.TO_DO);
  const participants=splitWorkersV2_(opsValue_(row,h,'Participants'));m.participants.forEach(function(worker){if(participants.indexOf(worker)<0)participants.push(worker);});
  const missing=(has3&&m.rawMissing3)||(has2&&m.rawMissing2),currentActive=splitWorkersV2_(opsValue_(row,h,'Active Assignees'));
  let allocation=String(opsValue_(row,h,'Allocation Status')||'');
  // Allocation is Task-authoritative. Active Assignees is only a derived summary,
  // so cancellation or a cleared summary cell must not imply missing owners.
  if(missing)allocation='분배 필요';
  else if(m.rawTotal>0&&['자동분배됨','수동분배됨'].indexOf(allocation)<0)allocation='수동분배됨';
  const activeWorkers=[];((status===APP_CONFIG.STATUS.DONE||finished)?(old===APP_CONFIG.STATUS.DONE?[currentActive]:[has3?m.valid3:[],has2?m.valid2:[]]):status===APP_CONFIG.STATUS.DROP?[[]]:[has3?m.pending3:[],has2?m.pending2:[]]).forEach(function(list){list.forEach(function(worker){if(activeWorkers.indexOf(worker)<0)activeWorkers.push(worker);});});
  return {'Total VI':m.total,'3P Completed VI':has3?m.done3:0,'2P Completed VI':has2?m.done2:0,'Completed VI':done,'Active Assignees':activeWorkers.join(', '),
    'Progress':m.total?done/m.total:0,'Status':status,'Participants':participants.join(', '),'Allocation Status':allocation};
}
// Patch only owned fields, coalescing contiguous rows. Never rewrite whole records.
function writeOpsPatches_(sheet,patches) {
  const statusColumn=sheet.getName()===APP_CONFIG.SHEETS.REQUESTS&&patches.some(p=>p.value===APP_CONFIG.STATUS.DONE)?getHeaderMap_(sheet).Status:0;
  const columns={};patches.forEach(function(p){(columns[p.col]||(columns[p.col]=[])).push(p);});
  Object.keys(columns).forEach(function(col){
    const list=columns[col].sort(function(a,b){return a.row-b.row;});let start=0;
    while(start<list.length){let end=start+1;while(end<list.length&&list[end].row===list[end-1].row+1)end++;
      const values=list.slice(start,end).map(p=>p.value);
      if(Number(col)===statusColumn)writeRequestStatuses_(sheet,list[start].row,Number(col),values);
      else sheet.getRange(list[start].row,Number(col),end-start,1).setValues(values.map(v=>[v]));
      start=end;
    }
  });
}
function addOpsPatch_(patches,record,h,name,value) {
  if(!h[name])return;const before=record.values[h[name]-1];
  if(before===value||before instanceof Date&&value instanceof Date&&before.getTime()===value.getTime())return;
  patches.push({row:record.row,col:h[name],value:value});record.values[h[name]-1]=value;
}
// Existing work session columns are retained; one history read for all closures.
function closeOpsRequestSessions_(ids,now,sessionRows,viContext) {
  const totals={},sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('WORK_SESSIONS');
  if(!ids.length||!sheet||sheet.getLastRow()<2)return totals;
  const wanted=new Set(ids),rows=sessionRows||workSessionRows_(),changes=[];
  const context=viContext||createWorkSessionViContext_();
  if(rows.some(function(r){return wanted.has(String(r[2]))&&r[4]==='ACTIVE';}))assertWorkSessionViSchema_(sheet,context);
  rows.forEach(function(r,i){const id=String(r[2]||'');if(!wanted.has(id))return;
    let minutes=Math.max(0,Number(r[7])||0);
    if(r[4]==='ACTIVE'){
      const next=finalizeWorkSessionViMetrics_(r,now,'요청 완료로 자동 종료',context);
      changes.push({row:i+2,values:next.slice(4,13),record:r,next:next});minutes=next[7];
    }totals[id]=(totals[id]||0)+minutes/60;
  });writeFinalizedWorkSessions_(sheet,changes);return totals;
}
function calculateAndWriteRequestSummaries_(requestIds,options) {
  options=options||{};
  const started=Date.now(),rt=options.requestTable||readOpsTable_(APP_CONFIG.SHEETS.REQUESTS),h=rt.headers;
  const ids=requestIds==null?Object.keys(rt.byId):Array.from(new Set(requestIds.map(String)));
  const records=ids.map(function(id){const found=rt.byId[id]||[];if(found.length!==1)throw Error('요청 ID가 없거나 중복되었습니다: '+id);return found[0];});
  const names=records.map(function(r){return opsSourceForRequest_(r.values,h);});
  const sources=options.executionContext?scopedSourcesForRequests_(rt,ids,options.executionContext):{};
  if(!options.executionContext)names.forEach(function(name){if(!sources[name])sources[name]=options.taskSources&&options.taskSources[name]||readTaskSources_([name])[name];});
  records.forEach(function(record){const id=String(opsValue_(record.values,h,'Request ID')),t=sources[opsSourceForRequest_(record.values,h)];
    if(t.scoped&&!t.loadedRequestIds.has(id))throw Error('미조회 Task 요청입니다: '+id);
  });
  const index=options.executionContext?Object.keys(sources).reduce(function(out,name){out[name]=sources[name].index;return out;},{}):buildTaskIndex_(sources),patches=[],complete=[],logs=[],summaries=Object.create(null),now=new Date();
  records.forEach(function(record){
    const row=record.values,id=String(opsValue_(row,h,'Request ID')),old=normalizeRequestStatusV2_(opsValue_(row,h,'Status'))||APP_CONFIG.STATUS.TO_DO;
    const metrics=(index[opsSourceForRequest_(row,h)]||{})[id]||emptyTaskAggregate_();
    const values=aggregateRequestMetrics_(row,h,metrics);
    if(options.preserveStatus)values.Status=old;
    summaries[id]=values;
    if(options.participantsOnly){addOpsPatch_(patches,record,h,'Participants',values.Participants);return;}
    // Pending crawl has no VI yet; keep its request/session workflow state.
    if(opsValue_(row,h,'크롤링 필요')===true&&opsValue_(row,h,'크롤링 단계')!=='CRAWL_COMPLETED'&&!metrics.total){values.Status=old;}
    if(['자동분배됨','수동분배됨'].indexOf(options.preferredAllocation)>=0&&values['Allocation Status']!=='분배 필요')values['Allocation Status']=options.preferredAllocation;
    Object.keys(values).forEach(function(k){addOpsPatch_(patches,record,h,k,values[k]);});
    if(values.Status==='DONE'){
      const end=opsValue_(row,h,'Completed At') instanceof Date?opsValue_(row,h,'Completed At'):now;
      addOpsPatch_(patches,record,h,'Completed At',end);
      const requested=opsValue_(row,h,'Requested At');if(requested instanceof Date)addOpsPatch_(patches,record,h,'Lead Time (Hours)',Math.max(0,(end-requested)/3600000));
      if(old!=='DONE')complete.push(id);
    }else if(old==='DONE'){
      addOpsPatch_(patches,record,h,'Completed At','');addOpsPatch_(patches,record,h,'Lead Time (Hours)','');
    }
    if(old!==values.Status)logs.push({id:id,before:old,after:values.Status});
  });
  const hours=closeOpsRequestSessions_(complete,now,options.sessionRows,createWorkSessionViContext_(rt,sources,index));
  complete.forEach(function(id){if(hours[id]!==undefined)addOpsPatch_(patches,rt.byId[id][0],h,'Work Time (Hours)',hours[id]);});
  writeOpsPatches_(rt.sheet,patches);
  logs.forEach(function(l){logActivity_(l.id,l.after==='DONE'?'REQUEST_COMPLETED':'REQUEST_STATUS_CHANGED',l.before,l.after);});
  opsTiming_('request_aggregate',started,{requests:ids.length,sources:Object.keys(sources).length});
  return {requests:ids.length,sources:Object.keys(sources).length,patches:patches.length,summaries:summaries};
}
// Explicit maintenance; never attached to selection or a clock trigger.
function rebuildRequestSummaries(){return withOpsDocumentLock_(function(){const result=syncRequestSummaries_(null);bumpMyWorkSidebarRevisionV2_([]);return result;});}
var OPS_PROFILE_ENABLED_;
function opsTiming_(operation,started,counts){
  // Opt-in execution logging, no UI messages and no request/worker identifiers.
  if(OPS_PROFILE_ENABLED_===undefined)OPS_PROFILE_ENABLED_=PropertiesService.getScriptProperties().getProperty('OPS_PROFILE')==='true';
  if(OPS_PROFILE_ENABLED_)console.log(JSON.stringify({operation:operation,elapsedMs:Date.now()-started,counts:counts||{}}));
}
// Known retired handlers only, and only triggers visible to the calling account.
function configureProductionHardening(){
  const retired=['repairMissingStageTimestamps','installedTaskStageEdit','onSelectionChange'];let removed=0;
  ScriptApp.getProjectTriggers().forEach(function(t){if(retired.indexOf(t.getHandlerFunction())>=0){ScriptApp.deleteTrigger(t);removed++;}});
  return {removed:removed,editHandler:'onEdit',automaticRepair:false};
}
function nativeTaskBlock_(table,id){
  const records=table.byId[id]||[];return records.length?{startRow:2,rows:records.map(function(r){return r.values;}),rowNumbers:records.map(function(r){return r.row;})}:null;
}
// Direct REQUESTS multi-row assignment: one read per relevant Task source.
function applyRequestAssignmentEdit_(e){
  const requestSheet=e.range.getSheet(),directHeaders=getHeaderMap_(requestSheet);
  if(e.range.getNumRows()===1){
    const row=e.range.getRow(),id=String(requestSheet.getRange(row,directHeaders['Request ID']).getValue()||'');
    const workers=splitWorkersV2_(requestSheet.getRange(row,directHeaders['Active Assignees']).getValue());
    const status=normalizeRequestStatusV2_(requestSheet.getRange(row,directHeaders['Status']).getValue());
    if(id&&workers.length&&status!==APP_CONFIG.STATUS.DROP)applyRemainingDistributionV2_(id,workers);
    else if(id)syncRequestSummaries_([id]);
    bumpMyWorkSidebarRevisionV2_(splitWorkersV2_(e.oldValue));
    return;
  }
  const rt=readOpsTable_(APP_CONFIG.SHEETS.REQUESTS),h=rt.headers,start=e.range.getRow(),end=e.range.getLastRow();
  const rows=Object.keys(rt.byId).flatMap(function(id){return rt.byId[id];}).filter(function(r){return r.row>=start&&r.row<=end;});
  const eligible=rows.filter(function(record){return splitWorkersV2_(opsValue_(record.values,h,'Active Assignees')).length&&normalizeRequestStatusV2_(opsValue_(record.values,h,'Status'))!==APP_CONFIG.STATUS.DROP;});
  const sources=eligible.length?readTaskSources_(eligible.map(function(r){return opsSourceForRequest_(r.values,h);})):{};
  const context={rt:rt,sources:sources,workers:new Set()},reconcileIds=[];
  rows.forEach(function(record){const id=String(opsValue_(record.values,h,'Request ID')),workers=splitWorkersV2_(opsValue_(record.values,h,'Active Assignees'));
    if(rt.byId[id].length!==1)throw Error('중복 요청 ID: '+id);
    if(workers.length&&normalizeRequestStatusV2_(opsValue_(record.values,h,'Status'))!==APP_CONFIG.STATUS.DROP)applyRemainingDistributionV2_(id,workers,context);
    else reconcileIds.push(id);
  });
  if(reconcileIds.length)syncRequestSummaries_(reconcileIds,{requestTable:rt});
  splitWorkersV2_(e.oldValue).forEach(function(w){context.workers.add(w);});
  bumpMyWorkSidebarRevisionV2_(Array.from(context.workers));
}
// Request-scoped data only. Never cache an edit context across executions/dialogs.
function readOperationalEditContext_(e) {
  const sheet=e.range.getSheet(),h=getHeaderMap_(sheet),start=Math.max(e.range.getRow(),2),count=e.range.getLastRow()-start+1;
  return {sheet:sheet,h:h,start:start,count:count,rows:count>0?sheet.getRange(start,1,count,sheet.getLastColumn()).getValues():[],patches:[],rejected:0,executionContext:e.sessionTaskContext||{}};
}
function handleOperationalTaskEdit_(e,context) {
  if(!e||!e.range||e.range.getLastRow()<2)return;
  if(!context)return withOpsDocumentLock_(function(){return handleOperationalTaskEdit_(e,readOperationalEditContext_(e));});
  const c=context||readOperationalEditContext_(e),sheet=c.sheet,h=c.h,a=taskAdapter_(h),first=e.range.getColumn(),last=e.range.getLastColumn();
  const touches=function(col){return col>0&&first<=col&&last>=col&&(!e.editedColumns||e.editedColumns.indexOf(col)>=0);};
  if(!['3P Assignee','3P Completed','2P Assignee','2P Completed','VI Status','Invalid Reason'].some(function(n){return touches(h[n]);}))return;
  return withOpsDocumentLock_(function(){
    const started=Date.now(),start=c.start,count=c.count;
    const rows=c.rows,rt=c.requestTable||readOpsTable_(APP_CONFIG.SHEETS.REQUESTS),ids=new Set(),patches=c.patches,workers=new Set();
    rows.forEach(function(row,i){
      const id=String(row[a.id]||'');if(!id)return;const request=rt.byId[id];if(!request||request.length!==1)throw Error('요청 ID가 없거나 중복되었습니다: '+id);
      const r=request[0].values,rh=rt.headers,type=String(opsValue_(r,rh,'Request Type')||'3P Listing');
      function put(col,value){if(col>=0&&row[col]!==value){patches.push({row:start+i,col:col+1,value:value});row[col]=value;}}
      if(touches(h['VI Status'])||touches(h['Invalid Reason'])){
        const reason=String(opsValue_(row,h,'Invalid Reason')||'').trim();
        put((h['VI Status']||0)-1,reason?'INVALID':'VALID');
      }
      [[a.a3,a.d3,type!=='2P Lite 전환','3P'],[a.a2,a.d2,type!=='3P Listing','2P']].forEach(function(stage){
        if(stage[0]<0||stage[1]<0)return;splitWorkersV2_(row[stage[0]]).forEach(function(w){workers.add(w);});
        // Owner-only edits must not rewrite existing completion history.
        if(!touches(stage[1]+1))return;
        const priorStageOK=stage[3]!=='2P'||type!=='3P Listing 후 2P Lite 전환'||a.d3>=0&&row[a.d3]===true;
        const barcodeProvidedOK=stage[3]!=='2P'||!!normalizeBarcodeProvidedV1_(opsValue_(r,rh,'Barcode Provided'));
        if(row[stage[1]]===true&&(!stage[2]||!isViValidV3_(row,h)||!splitWorkersV2_(row[stage[0]]).length||!priorStageOK||!barcodeProvidedOK)){
          put(stage[1],false);
        }
      });
      const d3=type==='2P Lite 전환'||a.d3>=0&&row[a.d3]===true,d2=type==='3P Listing'||a.d2>=0&&row[a.d2]===true;
      put(a.work,d3&&d2?'작업 완료':type!=='2P Lite 전환'&&d3?'3P Listing 완료':type!=='3P Listing'&&d2?'2P Lite 완료':'작업 대기');
      ids.add(id);
    });
    // Checkbox/Work Status only; VI dates are no longer written or cleared.
    markRequestReconcileNeeded_(Array.from(ids),{tasks:true},'task-edit');
    writeOpsPatches_(sheet,patches);
    // Native auto-START already read this source for its baseline. Apply this
    // edit's validated rows to that execution-local snapshot before summarizing.
    try{
    const executionContext=c.executionContext||(c.executionContext={});
    scopedSourcesForRequests_(rt,Array.from(ids),executionContext);
    overlayScopedTaskEdit_(executionContext,sheet.getName(),start,rows);
    const ownerEdit=touches(h['3P Assignee'])||touches(h['2P Assignee']);
    const result=syncRequestSummaries_(Array.from(ids),{requestTable:rt,sessionRows:c.sessions,executionContext:executionContext,preferredAllocation:ownerEdit?'수동분배됨':''});
    if(ownerEdit)splitWorkersV2_(e.oldValue).forEach(function(w){workers.add(w);});
    if(touches(h['VI Status'])||touches(h['Invalid Reason']))ids.forEach(function(id){logActivity_(id,'VI_VALIDITY_CHANGED','',count+' edited rows');});
    bumpMyWorkSidebarRevisionV2_(Array.from(workers));opsTiming_('task_edit',started,{rows:count,requests:ids.size});return result;
    }catch(error){
      // Task checks are already persisted. Summary/revision failures are repair
      // work, not a reason to roll back a check or gate the COMPLETE button.
      const pending=Array.from(ids),tickets=markRequestReconcileNeeded_(pending,{tasks:true},'task-edit-post');
      pending.forEach(function(id){settleRequestReconcile_(id,tickets[id],{tasks:true},null,error);});
      console.error('Task edit post-processing pending: '+String(error));return {summaryPending:true};
    }
  });
}
