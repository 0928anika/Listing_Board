// START/PAUSE protocol. Receipts are private recovery metadata, not session history.
// A busy lock is a retryable result; only the locked section may decide to write.
const SESSION_RECEIPT_PREFIX_ = 'SESSION_COMMIT_V1_';
function sessionReceiptKey_(id){return SESSION_RECEIPT_PREFIX_+id;}
function sessionReceipt_(id){const raw=PropertiesService.getScriptProperties().getProperty(sessionReceiptKey_(id));return raw?JSON.parse(raw):null;}
function saveSessionReceipt_(receipt){PropertiesService.getScriptProperties().setProperty(sessionReceiptKey_(receipt.operationId),JSON.stringify(receipt));}
function sessionCommand_(payload){
  const p=Object.assign({},payload),id=String(p.operationId||'');
  if(!/^[a-zA-Z0-9-]{16,80}$/.test(id))throw Error('세션 처리 ID가 올바르지 않습니다.');
  p.worker=String(p.worker||'').trim();p.requestId=String(p.requestId||'').trim();p.stage=String(p.stage||'').trim();p.action=String(p.action||'').toUpperCase();
  if(!p.worker||!p.requestId||['START','PAUSE','COMPLETE'].indexOf(p.action)<0)throw Error('세션 처리 항목을 확인해 주세요.');
  if(p.action!=='START'&&!p.sessionId)throw Error('중지할 Session ID가 없습니다. 작업함을 다시 열어 주세요.');
  return p;
}
function sessionSameCommand_(receipt,p){
  return receipt.worker===p.worker&&receipt.requestId===p.requestId&&receipt.stage===p.stage&&receipt.action===p.action&&
    (p.action==='START'||receipt.sessionId===p.sessionId);
}
function sessionValidEnd_(row){return row[4]==='PAUSED'&&row[6] instanceof Date&&Number.isFinite(row[6].getTime())&&typeof row[7]==='number'&&Number.isFinite(row[7])&&row[7]>=0;}
function sessionRetry_(state){return {ok:false,state:state||'WAITING',retryAfterMs:700};}
function sessionResult_(p,receipt,rows){
  if(p.action==='COMPLETE')return completeReadCommitted_(p,receipt,rows);
  const matches=rows.filter(function(r){return String(r[0])===receipt.sessionId;});
  if(matches.length!==1)return null;
  const row=matches[0];
  if(String(row[1])!==p.worker||String(row[2])!==p.requestId||String(row[3])!==p.stage)throw Error('Session ID의 작업자/요청/단계가 일치하지 않습니다.');
  const active=row[4]==='ACTIVE'&&row[5] instanceof Date&&!row[6]&&row[7]==='';
  if(p.action==='PAUSE'?!sessionValidEnd_(row):!active&&!sessionValidEnd_(row))return null;
  const duplicates=rows.filter(function(r){return String(r[1])===p.worker&&r[4]==='ACTIVE';});
  if(duplicates.length>1)throw Error('작업자에게 ACTIVE 세션이 중복되어 있습니다. 관리자 확인이 필요합니다.');
  return {ok:true,state:'COMMITTED',action:p.action,operationId:p.operationId,sessionId:receipt.sessionId,
    sessionStatus:String(row[4]),needsRefresh:true,postProcessing:receipt.postDone?'DONE':'PENDING',
    snapshot:getWorkSessionSnapshot_(true,rows.filter(function(r){return String(r[1])===p.worker;}))};
}
function withSessionCommitLock_(fn,waitMs){
  const lock=LockService.getScriptLock(),nested=lock.hasLock();
  if(!nested&&!lock.tryLock(waitMs||1200))return sessionRetry_('WAITING');
  try{return fn();}finally{if(!nested)lock.releaseLock();}
}
function setWorkSessionState(payload,nativeEdit){
  if(payload&&String(payload.action||'').toUpperCase()==='COMPLETE')return completeWorkSessionCommand_(payload);
  // COMPLETE uses its atomic command path above. Native checkbox
  // START callers share the same atomic core, but have no browser retry runner.
  if(!payload||['START','PAUSE'].indexOf(String(payload.action||'').toUpperCase())<0)return setWorkSessionStateLegacy_(payload,nativeEdit);
  const legacy=!payload.operationId,context={},began=Date.now();
  let p=Object.assign({},payload);
  if(legacy){p.operationId=Utilities.getUuid();p.issuedAt=Date.now();}
  const result=withSessionCommitLock_(function(){
    context.lockWaitMs=Date.now()-began;
    const sheet=ensureWorkSessionsSheet_(),rows=workSessionRows_();
    if(legacy&&String(p.action).toUpperCase()==='PAUSE'&&!p.sessionId){
      const matches=rows.filter(function(r){return String(r[1])===String(p.worker)&&String(r[2])===String(p.requestId)&&String(r[3])===String(p.stage);});
      const target=matches.find(function(r){return r[4]==='ACTIVE';})||matches.filter(sessionValidEnd_).pop();
      if(target)p.sessionId=String(target[0]);
    }
    p=sessionCommand_(p);
    let receipt=sessionReceipt_(p.operationId);
    if(receipt&&!sessionSameCommand_(receipt,p))return {ok:false,state:'REJECTED',message:'다른 세션 명령에 사용된 처리 ID입니다.'};
    if(receipt&&receipt.rejected)return {ok:false,state:'REJECTED',message:receipt.rejected};
    if(receipt){const saved=sessionResult_(p,receipt,rows);if(saved)return saved;}
    // Completed receipts may be compacted; the deterministic Session ID remains
    // the durable START deduplication key for as long as its history is online.
    if(!receipt&&p.action==='START'){
      const saved=sessionResult_(p,{sessionId:'WS-'+p.operationId,postDone:true},rows);if(saved)return saved;
    }
    // Do not replay an arbitrarily old command after its receipt/history was archived.
    if(!receipt&&(!Number.isFinite(Number(p.issuedAt))||Math.abs(Date.now()-Number(p.issuedAt))>86400000))return {ok:false,state:'REJECTED',message:'오래된 미처리 요청입니다. 최신 작업 상태를 확인한 뒤 다시 눌러 주세요.',snapshot:getWorkSessionSnapshot_(true,rows)};
    if(p.action==='PAUSE'){
      const exact=rows.filter(function(r){return String(r[0])===String(p.sessionId);});
      if(exact.length!==1||String(exact[0][1])!==p.worker||String(exact[0][2])!==p.requestId||String(exact[0][3])!==p.stage)return {ok:false,state:'REJECTED',message:'중지할 세션을 정확히 확인할 수 없습니다.',snapshot:getWorkSessionSnapshot_(true,rows)};
      if(sessionValidEnd_(exact[0]))return sessionResult_(p,{sessionId:p.sessionId,postDone:true},rows);
      if(exact[0][4]!=='ACTIVE')return {ok:false,state:'REJECTED',message:'세션 종료 기록이 불완전합니다. 관리자 확인이 필요합니다.'};
    }else{
      const active=rows.filter(function(r){return String(r[1])===p.worker&&r[4]==='ACTIVE';});
      const actual=active.length===1?String(active[0][0]):'';
      if(active.length>1)return {ok:false,state:'REJECTED',message:'ACTIVE 세션이 중복되어 있습니다. 관리자 확인이 필요합니다.'};
      if(p.expectedActiveSessionId!==undefined&&String(p.expectedActiveSessionId)!==actual)return {ok:false,state:'REJECTED',message:'다른 화면에서 작업 상태가 변경되었습니다. 최신 상태를 반영했습니다.',snapshot:getWorkSessionSnapshot_(true,rows)};
    }
    context.sheet=sheet;context.rows=rows;context.receipt=receipt;context.payload=p;
    try{
      setWorkSessionStateLegacy_(p,null,context);
      context.coreSaved=true;
      // Read the persisted row, not just the in-memory write plan, before success.
      const confirmed=sessionReadCommitted_(p,context.receipt);
      return confirmed||sessionRetry_('VERIFYING');
    }catch(error){
      console.error('Session '+p.operationId+': '+String(error&&error.stack||error));
      if(context.writeAttempted||context.coreSaved||context.reused)return sessionRetry_('VERIFYING');
      if(context.receipt)try{context.receipt.rejected=String(error.message||error);context.receipt.postDone=true;saveSessionReceipt_(context.receipt);}catch(ignored){}
      return {ok:false,state:'REJECTED',message:String(error.message||error),snapshot:getWorkSessionSnapshot_(true,rows)};
    }
  },legacy?30000:1200);
  result.timing={serverMs:Date.now()-began,lockWaitMs:context.lockWaitMs===undefined?Date.now()-began:context.lockWaitMs};
  if(legacy){
    if(!result.ok)throw Error(result.message||'세션 저장 상태를 확인 중입니다. 작업함에서 최신 세션을 확인하세요.');
    if(nativeEdit&&nativeEdit.range)nativeEdit.sessionTaskContext=context.viContext;
    if(!context.reused&&!context.deferHistoricalRepair)try{runSessionPostProcessing_(p.operationId,context.viContext);}catch(error){console.error('Session post-processing: '+String(error));}
    if(p.action==='START'&&!context.reused)try{navigateCommittedSession({operationId:p.operationId});}catch(error){console.error('Session navigation: '+String(error));}
  }
  return result;
}
function sessionPrepareCommit_(context,sessionId,affected){
  const p=context.payload;
  const receipt={operationId:p.operationId,worker:p.worker,requestId:p.requestId,stage:p.stage,action:p.action,
    sessionId:sessionId,affected:Array.from(new Set(affected)),createdAt:Number(p.issuedAt),postDone:false};
  // Persist BEFORE attempting the atomic write. A crash after the write can then
  // be resolved from the real Session ID and retried without repeating append.
  saveSessionReceipt_(receipt);context.receipt=receipt;
  markRequestReconcileNeeded_(receipt.affected,{tasks:true,sessions:true},'session-commit');
  PropertiesService.getScriptProperties().setProperty('SESSION_POST_PENDING','1');
}
function sessionReadCommitted_(p,receipt){
  if(p.action==='COMPLETE')return completeReadCommitted_(p,receipt);
  if(!receipt)return null;
  const sheet=getRequiredSheet_('WORK_SESSIONS'),last=sheet.getLastRow();
  const matches=last<2?[]:sheet.getRange(2,1,last-1,1).createTextFinder(receipt.sessionId).matchEntireCell(true).matchCase(true).useRegularExpression(false).findAll();
  if(matches.length!==1)return null;
  const target=sheet.getRange(matches[0].getRow(),1,1,13).getValues()[0];
  // Active-only snapshot keeps recovery independent of historical row count.
  const activeRows=sidebarWorkerSessionRows_(p.worker),snapshot=getWorkSessionSnapshot_(true,activeRows);
  const valid=p.action==='PAUSE'?sessionValidEnd_(target):(target[4]==='ACTIVE'&&target[5] instanceof Date&&!target[6]&&target[7]===''||sessionValidEnd_(target));
  if(!valid||String(target[0])!==receipt.sessionId||String(target[1])!==p.worker||String(target[2])!==p.requestId||String(target[3])!==p.stage)return null;
  if(activeRows.length>1)return null;
  // Include the paused target so the client can unambiguously clear its active card.
  if(target[4]!=='ACTIVE'&&!snapshot.sessions.some(function(s){return s.requestId===p.requestId&&s.stage===p.stage;}))snapshot.sessions.push(getWorkSessionSnapshot_(true,[target]).sessions[0]);
  return {ok:true,state:'COMMITTED',action:p.action,operationId:p.operationId,sessionId:receipt.sessionId,sessionStatus:target[4],
    snapshot:snapshot,needsRefresh:true,postProcessing:receipt.postDone?'DONE':'PENDING'};
}
function getWorkSessionCommandStatus(payload){
  const p=sessionCommand_(payload);
  return withSessionCommitLock_(function(){
    const receipt=sessionReceipt_(p.operationId);
    if(receipt&&!sessionSameCommand_(receipt,p))return {ok:false,state:'REJECTED',message:'세션 처리 ID가 일치하지 않습니다.'};
    if(receipt&&receipt.rejected)return {ok:false,state:'REJECTED',message:receipt.rejected};
    return sessionReadCommitted_(p,receipt)||{ok:false,state:'RETRY',retryAfterMs:700};
  });
}

// BatchUpdate applies every old-session close and the new-session append atomically.
// Explicit date formats retain Date values for the existing Dashboard/VI readers.
function sessionCell_(value,tz){
  if(value instanceof Date){
    const offset=Utilities.formatDate(value,tz,'Z'),match=/^([+-])(\d{2})(\d{2})$/.exec(offset);
    if(!match||!Number.isFinite(value.getTime()))throw Error('세션 시간대 형식을 확인할 수 없습니다.');
    const minutes=(Number(match[2])*60+Number(match[3]))*(match[1]==='-'?-1:1);
    return {userEnteredValue:{numberValue:25569+(value.getTime()+minutes*60000)/86400000},userEnteredFormat:{numberFormat:{type:'DATE_TIME',pattern:'yyyy-mm-dd hh:mm:ss'}}};
  }
  if(value===''||value==null)return {};
  return {userEnteredValue:typeof value==='number'?{numberValue:value}:typeof value==='boolean'?{boolValue:value}:{stringValue:String(value)}};
}
function commitSessionRows_(sheet,changes,newRow,commitContext){
  if(!changes.length&&!newRow&&!(commitContext&&commitContext.extraRequests&&commitContext.extraRequests.length))return;
  const ss=SpreadsheetApp.getActiveSpreadsheet(),tz=ss.getSpreadsheetTimeZone(),requests=[];
  const fields='userEnteredValue';
  function cells(values){return values.map(function(v){return sessionCell_(v,tz);});}
  changes.forEach(function(c){
    if(!sessionValidEnd_(c.next))throw Error('종료 시각/소요시간/상태를 함께 기록해야 합니다.');
    requests.push({updateCells:{start:{sheetId:sheet.getSheetId(),rowIndex:c.row-1,columnIndex:4},rows:[{values:cells(c.next.slice(4,13))}],fields:fields}});
    requests.push({repeatCell:{range:{sheetId:sheet.getSheetId(),startRowIndex:c.row-1,endRowIndex:c.row,startColumnIndex:6,endColumnIndex:7},cell:{userEnteredFormat:{numberFormat:{type:'DATE_TIME',pattern:'yyyy-mm-dd hh:mm:ss'}}},fields:'userEnteredFormat.numberFormat'}});
  });
  if(newRow){
    const appended=cells(newRow);
    // Empty Ended At must retain a date format for later writes/reconciliation.
    [5,6].forEach(function(i){appended[i].userEnteredFormat={numberFormat:{type:'DATE_TIME',pattern:'yyyy-mm-dd hh:mm:ss'}};});
    requests.push({appendCells:{sheetId:sheet.getSheetId(),rows:[{values:appended}],fields:'userEnteredValue,userEnteredFormat.numberFormat'}});
  }
  if(commitContext&&commitContext.extraRequests)Array.prototype.push.apply(requests,commitContext.extraRequests);
  SpreadsheetApp.flush();
  if(commitContext)commitContext.writeAttempted=true;
  Sheets.Spreadsheets.batchUpdate({requests:requests},ss.getId());
  changes.forEach(function(c){c.next.forEach(function(v,i){c.record[i]=v;});});
}

// Post-processing has its own short lock attempt and its own result. Failure
// never changes a committed session result. Subsequent Sidebar polling retries it.
function finishWorkSessionCommand(payload){
  const id=String(payload&&payload.operationId||'');
  if(!/^[a-zA-Z0-9-]{16,80}$/.test(id))return {ok:false,pending:true};
  return runSessionPostProcessing_(id);
}
function runSessionPostProcessing_(id,executionContext){
  let postReceipt=null;
  const lock=LockService.getScriptLock();if(lock.hasLock()||!lock.tryLock(100))return {ok:false,pending:true};
  try{
    const receipt=sessionReceipt_(id);postReceipt=receipt;if(!receipt||receipt.postDone)return {ok:true};
    const rows=workSessionRows_(),result=sessionResult_(receipt,receipt,rows);
    if(!result)return {ok:false,pending:true};
    const now=new Date(),table=readSessionRequestRecords_(receipt.affected);
    const record=table.byId[receipt.requestId][0],h=table.headers,current=normalizeRequestStatusV2_(opsValue_(record.values,h,'Status'));
    const session=rows.find(function(row){return String(row[0])===receipt.sessionId;});
    if(receipt.action==='START'){
      const firstStart=rows.filter(function(row){return String(row[2])===receipt.requestId&&row[5] instanceof Date;}).reduce(function(first,row){return row[5]<first?row[5]:first;},session[5]);
      if(h['Started At']&&!opsValue_(record.values,h,'Started At')){table.sheet.getRange(record.row,h['Started At']).setValue(firstStart).setNumberFormat('yyyy-mm-dd hh:mm');record.values[h['Started At']-1]=firstStart;}
      if([APP_CONFIG.STATUS.DONE,APP_CONFIG.STATUS.DROP,APP_CONFIG.STATUS.ON_HOLD,APP_CONFIG.STATUS.ONGOING].indexOf(current)<0){table.sheet.getRange(record.row,h.Status).setValue(APP_CONFIG.STATUS.ONGOING);record.values[h.Status-1]=APP_CONFIG.STATUS.ONGOING;}
    }
    reconcileRequestSummaries_(receipt.affected,{requestTable:table,executionContext:executionContext,sessionRows:rows,now:now,notify:false,reason:'session-post'});
    bumpMyWorkSidebarRevisionV2_([receipt.worker]);
    receipt.postDone=true;saveSessionReceipt_(receipt);
    return {ok:true};
   }catch(error){
    if(postReceipt)try{const domains={tasks:true,sessions:true},tickets=markRequestReconcileNeeded_(postReceipt.affected,domains,'session-post-failed');postReceipt.affected.forEach(requestId=>settleRequestReconcile_(requestId,tickets[requestId],domains,null,error));}catch(queueError){console.error('Session repair queue: '+String(queueError));}
    console.error('Session post-processing: '+String(error&&error.stack||error));return {ok:false,pending:true};
  }
  finally{lock.releaseLock();}
}
function retrySessionPostProcessing_(){
  const props=PropertiesService.getScriptProperties();if(!props.getProperty('SESSION_POST_PENDING'))return;
  const all=props.getProperties(),keys=Object.keys(all).filter(function(key){return key.indexOf(SESSION_RECEIPT_PREFIX_)===0;}),now=Date.now();
  const pending=[];
  keys.forEach(function(key){const r=JSON.parse(all[key]);if(!r.postDone)pending.push(r);else if(now-r.createdAt>(r.action==='PAUSE'||r.sessionId==='WS-'+r.operationId?900000:172800000))props.deleteProperty(key);});
  // Round-robin prevents one invalid historical request from starving other jobs.
  if(pending.length){const cursor=Number(props.getProperty('SESSION_POST_CURSOR')||0),r=pending[cursor%pending.length];props.setProperty('SESSION_POST_CURSOR',String(cursor+1));runSessionPostProcessing_(r.operationId);}
  // Keep the flag: clearing it outside the commit lock could lose a new job.
}
function maintainWorkSessionCommits(){
  let ok=true;
  try{retrySessionPostProcessing_();}catch(error){ok=false;console.error(String(error));}
  try{recoverPendingRequestReconciliations_();}catch(error){ok=false;console.error(String(error));}
  return {ok:ok,pending:!ok};
}
function navigateCommittedSession(payload){
  const receipt=sessionReceipt_(String(payload&&payload.operationId||''));
  if(!receipt||receipt.action!=='START')return {ok:false};
  const result=sessionReadCommitted_(receipt,receipt);if(!result||result.sessionStatus!=='ACTIVE')return {ok:false};
  const table=readSessionRequestRecords_([receipt.requestId]),record=table.byId[receipt.requestId][0];
  activateWorkRows_(receipt.requestId,receipt.stage,String(opsValue_(record.values,table.headers,'Request Type')||'3P Listing'));
  return {ok:true};
}

// Admin repair: formats only; never rewrites timestamps, duration, VI or status.
function applyWorkSessionDateFormatsV1(){
  assertDashboardAdmin_();
  const lock=LockService.getScriptLock();if(!lock.tryLock(1000))throw Error('세션 저장 중입니다. 잠시 후 다시 실행하세요.');
  try{
    const sheet=getRequiredSheet_('WORK_SESSIONS');assertWorkSessionViSchema_(sheet);
    const h=getHeaderMap_(sheet),rows=Math.max(0,sheet.getLastRow()-1);
    if(!h['Started At']||!h['Ended At'])throw Error('세션 날짜 헤더가 없습니다.');
    if(rows)['Started At','Ended At'].forEach(function(name){sheet.getRange(2,h[name],rows,1).setNumberFormat('yyyy-mm-dd hh:mm:ss');});
    SpreadsheetApp.flush();console.log('WORK_SESSIONS 날짜 서식 적용: '+rows+'행. 값은 변경하지 않았습니다.');
    return {rows:rows,valuesChanged:false};
  }finally{lock.releaseLock();}
}
