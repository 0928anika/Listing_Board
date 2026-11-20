// Session net VALID completion counts. No per-VI writes, ledger or background work.
const WORK_SESSION_COLUMNS_ = Object.freeze([
  'Session ID','Worker','Request ID','Stage','Status','Started At','Ended At',
  'Duration Minutes','Updated By','Note','Start Completed VI','End Completed VI','Completed VI Delta'
]);

// Existing sheets must be prepared explicitly by an administrator; never migrate here.
function assertWorkSessionViSchema_(sheet,context) {
  if(context&&context.schemaChecked)return;
  const headers=sheet.getRange(1,1,1,Math.min(sheet.getLastColumn(),13)).getDisplayValues()[0];
  if(WORK_SESSION_COLUMNS_.some(function(name,i){return headers[i]!==name;})){
    throw Error('WORK_SESSIONS 스키마를 확인하세요. 기존 A:J를 유지하고 K:M에 Start Completed VI / End Completed VI / Completed VI Delta 헤더를 수동으로 준비해야 합니다. 자동 변경하지 않았습니다.');
  }
  if(context)context.schemaChecked=true;
}
function createWorkSessionViContext_(requestTable,sources,index) {
  return {requestTable:requestTable||null,sources:sources||Object.create(null),index:index||Object.create(null)};
}
function workSessionViSource_(requestId,context) {
  if(!context.requestTable)context.requestTable=readOpsTable_(APP_CONFIG.SHEETS.REQUESTS);
  const rt=context.requestTable,records=rt.byId[String(requestId)]||[];
  if(records.length!==1)throw Error('세션의 요청 ID가 없거나 중복되었습니다: '+requestId);
  const source=opsSourceForRequest_(records[0].values,rt.headers);
  if(!context.sources[source]||context.sources[source].scoped&&!context.sources[source].loadedRequestIds.has(String(requestId))){
    context.sources[source]=readTasksForRequests_(source,[String(requestId)],context);
    context.index[source]=context.sources[source].index;
  }
  if(!context.index[source]){
    const one={};one[source]=context.sources[source];context.index[source]=buildTaskIndex_(one)[source];
  }
  return source;
}
function completedViForWorkSession_(worker,requestId,stage,context) {
  // CRAWL preparation measures time, not completed listing VI.
  // Stage 3P on CRAWL_3P_TASKS follows the usual request source adapter.
  if(stage!=='3P'&&stage!=='2P')return null;
  const source=workSessionViSource_(requestId,context),m=context.index[source][String(requestId)],w=m&&m.workers[String(worker)];
  return w?Number(w['done'+stage])||0:0;
}

// Reuse the session history already loaded by START; never recount Tasks for the baseline.
// End Completed VI is worker-specific, so histories must not cross workers or stages.
function previousWorkSessionEndViV2_(rows,worker,requestId,stage,commitContext) {
  if(stage!=='3P'&&stage!=='2P')return '';
  const validDate=v=>v instanceof Date&&Number.isFinite(v.getTime());
  const validCount=v=>typeof v==='number'&&Number.isFinite(v)&&v>=0&&Math.floor(v)===v;
  const candidates=[];
  rows.forEach(function(row,index){
    if(String(row[1])!==worker||String(row[2])!==requestId||String(row[3])!==stage||row[4]==='ACTIVE')return;
    // Read-only ordering: never repair history or invoke reconciliation in commit.
    const complete=validDate(row[6]),duration=typeof row[7]==='number'&&Number.isFinite(row[7])&&row[7]>=0;
    if(commitContext&&(!validDate(row[5])||!complete||!duration||!validCount(row[10])||!validCount(row[11])||typeof row[12]!=='number'||!Number.isFinite(row[12])))commitContext.deferHistoricalRepair=true;
    let time=complete?row[6].getTime():validDate(row[5])?row[5].getTime():null;
    if(!complete&&(row[6]===''||row[6]==null)&&validDate(row[5])&&duration){const end=new Date(row[5].getTime()+row[7]*60000);if(validDate(end))time=end.getTime();}
    candidates.push({row:row,index:index,time:time});
  });
  if(!candidates.length)return 0;
  // Unorderable history must not fabricate a VI baseline or prevent START.
  if(candidates.length>1&&candidates.some(c=>c.time===null))return '';
  candidates.sort((a,b)=>(a.time-b.time)||(a.index-b.index));
  const end=candidates[candidates.length-1].row[11];
  return validCount(end)?end:'';
}
// Pure row finalization; callers persist this together with the existing pause fields.
function finalizeWorkSessionViMetrics_(record,now,note,context) {
  if(record[4]!=='ACTIVE')return record.slice();
  const next=Array.from({length:13},function(_,i){return record[i]===undefined?'':record[i];});
  const end=completedViForWorkSession_(String(record[1]),String(record[2]),String(record[3]),context);
  next[4]='PAUSED';next[6]=now;
  next[7]=record[5] instanceof Date?Math.max(0,(now-record[5])/60000):0;
  if(context.user===undefined)context.user=workSessionUser_();next[8]=context.user;next[9]=note||'';
  if(end!==null){
    next[11]=end;
    // Missing historical baseline is unknown, never inferred as zero/current end.
    if(typeof record[10]==='number'&&Number.isFinite(record[10]))next[12]=end-record[10];
    else {next[12]='';next[9]+=(next[9]?' | ':'')+'VI 기준값 없음: 생산량 미집계';}
  }else {next[10]='';next[11]='';next[12]='';}
  return next;
}
// Consecutive session rows E:M in one call; prepare all rows before writing any.
function writeFinalizedWorkSessions_(sheet,changes,commitContext) {
  commitSessionRows_(sheet,changes,null,commitContext);
}
