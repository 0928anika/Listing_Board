// DROP must never leave a live work session behind.
//
// The normal DROP path closes sessions before Task rows are made
// INVALID, so 3P/2P End Completed VI and Delta still reflect the last valid
// Task state.  The repair path is deliberately different: once a historical
// request is already DROP/INVALID, its pre-DROP end count cannot be
// proven.  Repair therefore preserves every existing VI cell and records the
// elapsed time and an explicit "not counted" note instead of inventing zero.

function appendWorkSessionNoteV1_(existing, addition) {
  const before=String(existing||'').trim(),after=String(addition||'').trim();
  if(!after||before.indexOf(after)>=0)return before;
  return before?before+' | '+after:after;
}

// Close every active worker/stage session while the Task validity is intact.
function finalizeRequestSessionsForCancellationV1_(requestId, now, providedRequestTable, providedRows, preparedValidityContext) {
  const id=String(requestId||'').trim();
  if(!id)throw Error('DROP할 Request ID를 확인해 주세요.');
  const spreadsheet=SpreadsheetApp.getActiveSpreadsheet(),sheet=spreadsheet.getSheetByName('WORK_SESSIONS');
  if(!sheet||sheet.getLastRow()<2)return {requestId:id,closedSessions:0,workers:[],rows:providedRows||[]};
  const rows=providedRows||workSessionRows_(),targets=[];
  rows.forEach(function(row,index){
    if(String(row[2]||'')===id&&String(row[4]||'')==='ACTIVE')targets.push({record:row,row:index+2});
  });
  if(!targets.length)return {requestId:id,closedSessions:0,workers:[],rows:rows};

  // Prepare every VI end value before the first session write.
  const requestTable=preparedValidityContext?preparedValidityContext.requestTable:readSessionRequestRecords_([id],providedRequestTable),found=requestTable.byId[id]||[];
  if(found.length!==1)throw Error('요청 ID가 없거나 중복되었습니다: '+id);
  const context=createWorkSessionViContext_(requestTable),viIds=targets.filter(function(target){
    return ['3P','2P'].indexOf(String(target.record[3]||''))>=0;
  }).map(function(){return id;});
  assertWorkSessionViSchema_(sheet,context);
  if(viIds.length){
    if(preparedValidityContext){
      context.sources[preparedValidityContext.sourceName]=preparedValidityContext.table;
      context.index[preparedValidityContext.sourceName]=preparedValidityContext.table.index;
    }else{
      context.sources=scopedSourcesForRequests_(requestTable,viIds,context);
      Object.keys(context.sources).forEach(function(name){context.index[name]=context.sources[name].index;});
    }
  }
  now=now||new Date();
  const workers=[],changes=targets.map(function(target){
    const record=target.record,stage=String(record[3]||'');
    const note=appendWorkSessionNoteV1_(record[9],'요청 DROP으로 자동 종료');
    const next=finalizeWorkSessionViMetrics_(record,now,note,context);
    // Retired/unknown stages have no current VI adapter. Preserve any legacy
    // VI cells rather than clearing historical values during DROP.
    if(['3P','2P','CRAWL'].indexOf(stage)<0){next[10]=record[10];next[11]=record[11];next[12]=record[12];}
    const worker=String(record[1]||'');if(worker&&workers.indexOf(worker)<0)workers.push(worker);
    return {row:target.row,values:next.slice(4,13),record:record,next:next};
  });
  writeFinalizedWorkSessions_(sheet,changes);
  updateRequestWorkTimeFromSessions_(id,now,rows,requestTable);
  return {requestId:id,closedSessions:changes.length,workers:workers,rows:rows};
}

// One scoped Task read supplies both the session VI close and invalidation.
function cancelRequestTasksAndSessionsV1_(requestId, requestReason, now) {
  const context=requestTaskValidityContextV5_(requestId);
  const sessions=finalizeRequestSessionsForCancellationV1_(context.id,now,context.requestTable,null,context);
  const validity=invalidateRequestTasksV3_(context.id,requestReason,context);
  return {requestId:context.id,closedSessions:sessions.closedSessions,changedTasks:validity.changed,source:context.sourceName};
}

function cancelledActiveSessionPlanV1_() {
  const requestTable=readOpsTable_(APP_CONFIG.SHEETS.REQUESTS),h=requestTable.headers;
  if(!h['Request ID']||!h['Status'])throw Error('REQUESTS의 Request ID/Status 스키마를 확인해 주세요.');
  const cancelled=Object.create(null),ambiguous=Object.create(null);
  Object.keys(requestTable.byId).forEach(function(id){
    const records=requestTable.byId[id]||[];
    if(records.length!==1){ambiguous[id]=true;return;}
    if(normalizeRequestStatusV2_(opsValue_(records[0].values,h,'Status'))===APP_CONFIG.STATUS.DROP)cancelled[id]=true;
  });
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('WORK_SESSIONS');
  const rows=!sheet||sheet.getLastRow()<2?[]:workSessionRows_(),targets=[],workers=[],ids=[];
  rows.forEach(function(row,index){
    if(String(row[4]||'')!=='ACTIVE')return;
    const id=String(row[2]||'').trim();
    if(ambiguous[id])throw Error('중복 Request ID의 ACTIVE 세션은 자동 복구할 수 없습니다: '+id);
    if(!cancelled[id])return;
    const worker=String(row[1]||'').trim();
    targets.push({record:row,row:index+2});
    if(id&&ids.indexOf(id)<0)ids.push(id);
    if(worker&&workers.indexOf(worker)<0)workers.push(worker);
  });
  return {requestTable:requestTable,sheet:sheet,rows:rows,targets:targets,requestIds:ids,workers:workers};
}

// Read-only preflight for an administrator before the one-time repair.
function previewCancelledActiveWorkSessionsV1() {
  const plan=cancelledActiveSessionPlanV1_();
  return {ok:true,activeSessions:plan.targets.length,requestIds:plan.requestIds,workers:plan.workers};
}

function finalizeCancelledOrphanSessionV1_(record, now, context) {
  const next=Array.from({length:13},function(_,i){return record[i]===undefined?'':record[i];});
  if(String(record[4]||'')!=='ACTIVE')return next;
  const existingEnd=record[6] instanceof Date&&!isNaN(record[6].getTime())?record[6]:null,end=existingEnd||now;
  next[4]='PAUSED';next[6]=end;
  next[7]=record[5] instanceof Date&&!isNaN(record[5].getTime())?Math.max(0,(end-record[5])/60000):Math.max(0,Number(record[7])||0);
  if(context.user===undefined)context.user=workSessionUser_();
  next[8]=context.user;
  next[9]=appendWorkSessionNoteV1_(record[9],'DROP 후 남은 세션 복구 · VI 종료값 확인 불가: 생산량 미집계');
  // K:M are intentionally preserved. The post-DROP Task state cannot prove
  // an End Completed VI, so blank remains UNKNOWN rather than becoming zero.
  return next;
}

// One-time, idempotent repair for historical DROP + ACTIVE rows.
// No Task rows are changed. Existing VI cells and notes are never discarded.
function repairCancelledActiveWorkSessionsV1_() {
  const plan=cancelledActiveSessionPlanV1_();
  if(!plan.targets.length)return {repairedSessions:0,requestIds:[],workers:[],viPolicy:'PRESERVE_EXISTING_OR_UNKNOWN',endedAtPolicy:'DROP_AT_OR_REPAIR_TIME'};
  const context=createWorkSessionViContext_(plan.requestTable),now=new Date();
  assertWorkSessionViSchema_(plan.sheet,context);
  const changes=plan.targets.map(function(target){
    const id=String(target.record[2]||''),request=plan.requestTable.byId[id][0];
    const cancelledAt=opsValue_(request.values,plan.requestTable.headers,'Cancelled At'),startedAt=target.record[5];
    const safeCancelledAt=cancelledAt instanceof Date&&!isNaN(cancelledAt.getTime())&&
      (!(startedAt instanceof Date)||isNaN(startedAt.getTime())||cancelledAt.getTime()>=startedAt.getTime())?cancelledAt:now;
    const next=finalizeCancelledOrphanSessionV1_(target.record,safeCancelledAt,context);
    return {row:target.row,values:next.slice(4,13),record:target.record,next:next};
  });
  writeFinalizedWorkSessions_(plan.sheet,changes);
  plan.requestIds.forEach(function(id){
    updateRequestWorkTimeFromSessions_(id,now,plan.rows,plan.requestTable);
    logActivity_(id,'DROP_ACTIVE_SESSIONS_REPAIRED','ACTIVE','PAUSED ('+changes.filter(function(change){return String(change.record[2])===id;}).length+')');
  });
  bumpMyWorkSidebarRevisionV2_(plan.workers);
  return {repairedSessions:changes.length,requestIds:plan.requestIds,workers:plan.workers,viPolicy:'PRESERVE_EXISTING_OR_UNKNOWN',endedAtPolicy:'DROP_AT_OR_REPAIR_TIME'};
}

function repairCancelledActiveWorkSessionsV1() {
  return withOpsDocumentLock_(repairCancelledActiveWorkSessionsV1_);
}

// Fail loudly if any DROP request still owns an ACTIVE session.
function verifyCancelledActiveWorkSessionsV1() {
  const plan=cancelledActiveSessionPlanV1_();
  if(plan.targets.length)throw Error('DROP 요청에 ACTIVE 세션이 남아 있습니다: '+plan.targets.length+'건 / '+plan.requestIds.join(', '));
  return {ok:true,activeSessions:0};
}
