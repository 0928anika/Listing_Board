// Resumable archive. One durable job, bounded Task rows, recurring recovery.
const ARCHIVE_RESUME_V2=Object.freeze({RUN:'ARCHIVE_RESUME_RUN_V2',INSTALL:'ARCHIVE_RESUME_INSTALL_V2',HANDLER:'continueArchiveV2',PROGRESS:'ARCHIVE_PROGRESS',ROWS:100,BUDGET:45000});
function archiveResumePropsV2_(){return PropertiesService.getDocumentProperties();}
function archiveResumeReadV2_(key){return JSON.parse(archiveResumePropsV2_().getProperty(key)||'null');}
function archiveResumeSaveV2_(run){run.updatedAt=new Date().toISOString();archiveResumePropsV2_().setProperty(ARCHIVE_RESUME_V2.RUN,JSON.stringify(run));}
function archiveResumeInstallV2_(){
  const source=SpreadsheetApp.getActiveSpreadsheet();if(!source)throw Error('연결된 Apps Script 편집기에서 실행하세요.');
  const owner=String(Session.getEffectiveUser().getEmail()||'').toLowerCase(),recorded=archiveResumePropsV2_().getProperty(ARCHIVE_CONFIG_V1.TRIGGER_OWNER_KEY),old=archiveResumeReadV2_(ARCHIVE_RESUME_V2.INSTALL);
  if(!owner||recorded&&recorded.toLowerCase()!==owner||old&&old.owner!==owner)throw Error('일일 백업을 설치한 관리자 계정으로 실행하세요.');
  const matches=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===ARCHIVE_RESUME_V2.HANDLER),trigger=matches.shift()||ScriptApp.newTrigger(ARCHIVE_RESUME_V2.HANDLER).timeBased().everyMinutes(1).create();matches.forEach(t=>ScriptApp.deleteTrigger(t));
  archiveResumePropsV2_().setProperty(ARCHIVE_RESUME_V2.INSTALL,JSON.stringify({owner:owner,sourceId:source.getId(),triggerId:trigger.getUniqueId()}));
}
function startArchiveResumeV2_(options){
  archiveResumeInstallV2_();
  const lock=LockService.getDocumentLock();if(!lock.tryLock(100))return {ok:true,state:'WAITING',message:'기존 백업 실행이 끝난 뒤 다시 시작하세요.'};
  try{
    let run=archiveResumeReadV2_(ARCHIVE_RESUME_V2.RUN);
    if(typeof archiveLinkResetRetriesV1_==='function')archiveLinkResetRetriesV1_();
    if(run&&(run.active||run.status==='NEEDS_ATTENTION')){run.active=true;run.status='QUEUED';run.failures=0;run.inFlight=false;run.lastError='';run.nextAt=0;archiveResumeSaveV2_(run);archiveResumeProgressV2_(run);return {ok:true,state:'QUEUED',runId:run.id,resumed:true};}
    const source=SpreadsheetApp.getActiveSpreadsheet(),targetId=archiveSpreadsheetIdV1_(true);if(targetId===source.getId())throw Error('백업 대상이 운영파일과 같습니다.');
    run={id:'ARCH2-'+Utilities.getUuid(),active:true,sourceId:source.getId(),targetId:targetId,cutoff:(options&&options.now instanceof Date?options.now:new Date()).toISOString(),onlyRequestIds:options&&options.onlyRequestIds||null,completed:0,taskRows:0,status:'QUEUED',job:null,nextAt:0,lastError:''};
    archiveResumeSaveV2_(run);archiveResumeProgressV2_(run);return {ok:true,state:'QUEUED',runId:run.id,resumed:false};
  }finally{lock.releaseLock();}
}
function continueArchiveV2(event){
  const install=archiveResumeReadV2_(ARCHIVE_RESUME_V2.INSTALL),source=SpreadsheetApp.getActiveSpreadsheet();
  if(!install||!source||install.sourceId!==source.getId()||!event||String(event.triggerUid)!==String(install.triggerId)||String(Session.getEffectiveUser().getEmail()||'').toLowerCase()!==install.owner)throw Error('설치된 백업 이어하기 트리거에서만 실행할 수 있습니다.');
  return processArchiveResumeV2_();
}
function getArchiveProgressV2(){if(!SpreadsheetApp.getActiveSpreadsheet())throw Error('연결된 편집기에서 실행하세요.');assertDashboardAdmin_();const run=archiveResumeReadV2_(ARCHIVE_RESUME_V2.RUN);console.log(JSON.stringify(run||{status:'NOT_STARTED'}));return run;}
function archiveResumeProgressV2_(run){
  try{
    const ss=SpreadsheetApp.getActiveSpreadsheet();let sheet=ss.getSheetByName(ARCHIVE_RESUME_V2.PROGRESS);if(!sheet)sheet=ss.insertSheet(ARCHIVE_RESUME_V2.PROGRESS);
    const headers=['Run ID','Status','Current Request','Phase','Source Sheet','Completed Requests','Archived Task Rows','Updated At','Last Error','Remaining Requests','Copy Cursor'];
    if(!sheet.getLastRow()){sheet.getRange(1,1,1,headers.length).setValues([headers]);sheet.setFrozenRows(1);}
    if(JSON.stringify(sheet.getRange(1,1,1,headers.length).getValues()[0])!==JSON.stringify(headers))throw Error('ARCHIVE_PROGRESS 헤더가 달라 표시를 보류했습니다.');
    const job=run.job||{},sourceName=job.phase==='DELETE_TASKS'?job.taskNames[job.index]:job.phase==='FINALIZE'?'WORK_SESSIONS':job.names&&job.names[job.index]||'';sheet.getRange(2,1,1,headers.length).setValues([[run.id,run.status,job.id||'',job.phase||'',sourceName,run.completed,run.taskRows,run.updatedAt,run.lastError||'',run.remaining===undefined?'':run.remaining,job.cursor||0]]);
  }catch(error){console.error('Archive progress display: '+String(error));}
}
function archiveResumeFindV2_(sheet,column,id){const n=sheet.getLastRow()-1;if(n<1)return [];return sheet.getRange(2,column,n,1).createTextFinder(String(id)).matchEntireCell(true).matchCase(true).useRegularExpression(false).findAll().map(c=>c.getRow()).sort((a,b)=>a-b);}
let archiveTaskCopyCacheV4_=null; // Invocation-local only.
function archiveResumeReadRowsV2_(sheet,rows,startColumn,width){
  const result=[],sorted=rows.slice().sort((a,b)=>a-b);let i=0;
  if(width===1){
    while(i<sorted.length){const start=sorted[i];let end=i+1;while(end<sorted.length&&sorted[end]-start<5000)end++;
      const values=sheet.getRange(start,startColumn,sorted[end-1]-start+1,1).getValues();
      for(let j=i;j<end;j++)result.push({row:sorted[j],values:values[sorted[j]-start]});i=end;
    }return result;
  }
  while(i<sorted.length){const start=sorted[i];let end=i+1;while(end<sorted.length&&sorted[end]===sorted[end-1]+1&&end-i<100)end++;
    const values=sheet.getRange(start,startColumn,end-i,width).getValues();values.forEach((v,j)=>result.push({row:start+j,values:v}));i=end;
  }return result;
}
function archiveResumeSourceV2_(name,id){
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);if(!sheet){if(name==='WORK_SESSIONS')return null;throw Error('필수 원본 시트 누락: '+name);}
  const info=archiveHeaderInfoV1_(sheet,['Request ID']),rows=archiveResumeFindV2_(sheet,info.map['Request ID']+1,id);return {sheet:sheet,headers:info.headers,map:info.map,rows:rows};
}
function archiveResumeRequestV2_(job){
  const t=archiveResumeSourceV2_(APP_CONFIG.SHEETS.REQUESTS,job.id);if(t.rows.length!==1)throw Error('Request ID가 없거나 중복되었습니다.');
  const record={rowNumber:t.rows[0],requestId:job.id,map:t.map,values:t.sheet.getRange(t.rows[0],1,1,t.headers.length).getValues()[0]},v=n=>record.values[t.map[n]];
  const guard=JSON.stringify([normalizeRequestStatusV2_(v('Status')),v('Completed At'),v('Cancelled At'),v('Request Type'),v('크롤링 필요')]);
  if(job.guard&&guard!==job.guard)throw Error('백업 중 요청 상태/완료일/종류가 변경되었습니다. 원본 삭제를 보류했습니다.');
  return {table:t,record:record,guard:guard};
}
function archiveResumeDestinationV2_(target,name,headers,id,sourceId){
  const sheet=ensureArchiveDestinationV1_(target,ARCHIVE_CONFIG_V1.DESTINATIONS[name],headers),info=archiveHeaderInfoV1_(sheet,['Request ID','Archive Key','Source Spreadsheet ID']),rows=archiveResumeFindV2_(sheet,info.map['Request ID']+1,id),keys={},owned=[];
  const owners=archiveResumeReadRowsV2_(sheet,rows,info.map['Source Spreadsheet ID']+1,1),keyValues=archiveResumeReadRowsV2_(sheet,rows,info.map['Archive Key']+1,1);
  rows.forEach(function(row,i){if(String(owners[i].values[0])!==sourceId)return;const key=String(keyValues[i].values[0]||'');if(!key||keys[key])throw Error('백업 고유키가 비었거나 중복되었습니다.');keys[key]=row;owned.push(row);});
  return {sheet:sheet,headers:info.headers,map:info.map,keys:keys,rows:owned};
}
function archiveResumeSameV2_(source,destination,sourceHeaders,destMap){
  return sourceHeaders.every(function(name,i){if(!name||ARCHIVE_METADATA_HEADERS_V1.indexOf(name)>=0)return true;if(!Object.prototype.hasOwnProperty.call(destMap,name))return false;return JSON.stringify(source[i]===undefined?'':source[i])===JSON.stringify(destination[destMap[name]]===undefined?'':destination[destMap[name]]);});
}
function archiveResumeRecordsV2_(table,rows,id){return archiveResumeReadRowsV2_(table.sheet,rows,1,table.headers.length).map(r=>({rowNumber:r.row,requestId:id,map:table.map,values:r.values}));}
function archiveResumeCopyV2_(run,job,target){
  const name=job.names[job.index],token=run.id+'|'+job.id+'|'+name;
  const cached=archiveTaskCopyCacheV4_&&archiveTaskCopyCacheV4_.token===token?archiveTaskCopyCacheV4_:null;
  if(cached&&job.cursor>=cached.keys.length)archiveTaskCopyCacheV4_=null;
  const table=archiveTaskCopyCacheV4_&&archiveTaskCopyCacheV4_.token===token?archiveTaskCopyCacheV4_.table:archiveResumeSourceV2_(name,job.id);
  if(!table){job.counts[name]=0;job.index++;job.cursor=0;return;}
  if(job.taskNames.indexOf(name)>=0)return archiveResumeCopyTasksV3_(run,job,target,name,table);
  if(name==='WORK_SESSIONS'&&table.rows.length){
    const status=table.map['Status'];if(status===undefined)throw Error('WORK_SESSIONS Status 헤더 누락');
    if(archiveResumeReadRowsV2_(table.sheet,table.rows,status+1,1).some(r=>r.values[0]==='ACTIVE'))throw Error('ACTIVE 세션이 남아 있어 백업을 보류합니다.');
  }
  const dest=archiveResumeDestinationV2_(target,name,table.headers,job.id,run.sourceId);
  if(job.cursor===0)job.counts[name]=table.rows.length;
  if(job.counts[name]!==table.rows.length){job.cursor=0;job.counts[name]=table.rows.length;return;}
  const records=archiveResumeRecordsV2_(table,table.rows.slice(job.cursor,job.cursor+job.chunk),job.id),pending=[],seen={};
  const metadata={'Archive Batch ID':job.batchId,'Archived At':new Date(job.createdAt),'Source Spreadsheet ID':run.sourceId,'Source Sheet':name,'Schema Version':archiveSchemaVersionV1_(table.headers),'Extra Data JSON':''};
  records.forEach(function(record){
    const key=archiveKeyV1_(name,job.id,record.values,record.rowNumber,table.map);if(seen[key])throw Error('원본 Task/세션 고유키 중복');seen[key]=true;
    let existing=null;if(dest.keys[key])existing=dest.sheet.getRange(dest.keys[key],1,1,dest.headers.length).getValues()[0];
    if(existing&&archiveResumeSameV2_(record.values,existing,table.headers,dest.map))return;
    const values=Array(dest.headers.length).fill('');table.headers.forEach(function(h,i){if(h&&Object.prototype.hasOwnProperty.call(dest.map,h))values[dest.map[h]]=record.values[i];});
    Object.keys(metadata).forEach(h=>values[dest.map[h]]=existing?existing[dest.map[h]]:metadata[h]);values[dest.map['Archive Key']]=key;values[dest.map['Source Row Number']]=record.rowNumber;
    if(existing)archiveResumeWriteValuesV5_(target,dest.sheet,dest.keys[key],[values]);else pending.push(values);
  });
  if(pending.length){const start=dest.sheet.getLastRow()+1,required=start+pending.length-1;if(required>dest.sheet.getMaxRows())dest.sheet.insertRowsAfter(dest.sheet.getMaxRows(),required-dest.sheet.getMaxRows());archiveResumeWriteValuesV5_(target,dest.sheet,start,pending);}
  // Cursor advances only after writes return. If a response is lost, key + full
  // source-field equality makes repeating this same chunk safe.
  job.cursor+=records.length;
  if(job.cursor>=table.rows.length){const actual=Object.keys(dest.keys).length+pending.length;if(actual!==table.rows.length)throw Error('원본/백업 고유 행 수가 다릅니다: '+name+' (원본 '+table.rows.length+', 백업 '+actual+')');job.index++;job.cursor=0;}
}
function archiveResumeDeleteV2_(run,job,target){
  const name=job.taskNames[job.index],table=archiveResumeSourceV2_(name,job.id),dest=archiveResumeDestinationV2_(target,name,table.headers,job.id,run.sourceId);
  if(dest.rows.length!==job.counts[name])throw Error('백업 행 수가 변경되었습니다: '+name);
  if(!table.rows.length){job.index++;return;}
  const records=archiveResumeRecordsV2_(table,table.rows.slice(0,job.chunk),job.id);
  const verifyRows=records.map(function(r){const key=archiveKeyV1_(name,job.id,r.values,r.rowNumber,r.map);if(!dest.keys[key])throw Error('백업 Task 누락: 삭제 보류 '+name);return dest.keys[key];}),verified={};
  archiveResumeReadRowsV2_(dest.sheet,verifyRows,1,dest.headers.length).forEach(r=>verified[r.row]=r.values);
  records.forEach(function(r,i){if(!archiveResumeSameV2_(r.values,verified[verifyRows[i]],table.headers,dest.map))throw Error('복사 후 원본/백업 내용이 달라 삭제를 보류했습니다: '+name);});
  const lock=LockService.getScriptLock();if(!lock.tryLock(100))throw archiveResumeBusyV2_('세션 저장 중: 다음 실행에서 삭제를 재시도합니다.');
  try{
    archiveResumeRequestV2_(job);
    const fresh=archiveResumeRecordsV2_(table,records.map(r=>r.rowNumber),job.id);records.forEach(function(r,i){if(JSON.stringify(r.values)!==JSON.stringify(fresh[i].values))throw Error('원본 행이 이동/변경되어 삭제를 보류했습니다.');});
    const requests=archiveResumeDeleteRequestsV2_(table.sheet,records.map(r=>r.rowNumber));
    Sheets.Spreadsheets.batchUpdate({requests:requests},run.sourceId);
    // No deletion cursor: always re-query remaining source rows. A lost delete
    // response cannot delete the next request's rows or lose archived counts.
  }finally{lock.releaseLock();}
}
function archiveResumeDeleteRequestsV2_(sheet,rows){const sorted=rows.slice().sort((a,b)=>b-a),requests=[];let i=0;while(i<sorted.length){const high=sorted[i];let low=high;i++;while(i<sorted.length&&sorted[i]===low-1)low=sorted[i++];requests.push({deleteDimension:{range:{sheetId:sheet.getSheetId(),dimension:'ROWS',startIndex:low-1,endIndex:high}}});}return requests;}
function archiveResumeCellV2_(sheet,row,column,value){return {updateCells:{start:{sheetId:sheet.getSheetId(),rowIndex:row-1,columnIndex:column-1},rows:[{values:[sessionCell_(value,ARCHIVE_CONFIG_V1.TIMEZONE)]}],fields:value instanceof Date?'userEnteredValue,userEnteredFormat.numberFormat':'userEnteredValue'}};}
function archiveResumeFinalizeV2_(run,job,target){
  const request=archiveResumeRequestV2_(job),sessions=archiveResumeSourceV2_('WORK_SESSIONS',job.id);
  const originalInfo=sessions||{headers:['Request ID','Session ID','Worker','Stage','Status','Started At','Ended At','Duration Minutes','Completed VI Delta']};
  const dest=archiveResumeDestinationV2_(target,'WORK_SESSIONS',originalInfo.headers,job.id,run.sourceId);
  if(dest.rows.length!==(job.counts.WORK_SESSIONS||0))throw Error('백업 세션 수가 변경되었습니다.');
  const archiveRecords=archiveResumeRecordsV2_({sheet:dest.sheet,headers:dest.headers,map:dest.map},dest.rows,job.id);
  const current=sessions?archiveResumeRecordsV2_(sessions,sessions.rows,job.id):[];
  current.forEach(function(r){const key=archiveKeyV1_('WORK_SESSIONS',job.id,r.values,r.rowNumber,r.map),saved=archiveRecords.find(a=>a.rowNumber===dest.keys[key]);if(!saved||!archiveResumeSameV2_(r.values,saved.values,sessions.headers,dest.map)||r.values[r.map.Status]==='ACTIVE')throw Error('세션 기록이 변경되어 완료를 보류합니다.');});
  const longRows=buildLongTermSessionRowsV1_({sheet:dest.sheet,records:archiveRecords},{records:[request.record]},new Date(job.createdAt),job.batchId,ARCHIVE_CONFIG_V1.TIMEZONE);
  const aggregate=ensureLongTermSessionSheetV1_(),ah=archiveHeaderInfoV1_(aggregate,LONG_TERM_SESSION_COLUMNS_V1),existing={};
  const matches=archiveResumeFindV2_(aggregate,ah.map['Request ID']+1,job.id);archiveResumeReadRowsV2_(aggregate,matches,1,aggregate.getLastColumn()).forEach(r=>{const id=String(r.values[ah.map['Aggregate ID']]);if(existing[id])throw Error('장기 집계 ID 중복');existing[id]=r.row;});
  const requests=[],tz=ARCHIVE_CONFIG_V1.TIMEZONE;
  // All summaries, source session deletions and Request metadata are one source
  // spreadsheet atomic batch. No PENDING gap or double-counted Dashboard sessions.
  let nextRow=aggregate.getLastRow()+1;
  longRows.forEach(function(row){row[1]='COMMITTED';const at=existing[String(row[0])]||nextRow++;requests.push({updateCells:{start:{sheetId:aggregate.getSheetId(),rowIndex:at-1,columnIndex:0},rows:[{values:row.map(v=>sessionCell_(v,tz))}],fields:'userEnteredValue,userEnteredFormat.numberFormat'}});});
  if(nextRow-1>aggregate.getMaxRows())requests.unshift({appendDimension:{sheetId:aggregate.getSheetId(),dimension:'ROWS',length:nextRow-1-aggregate.getMaxRows()}});
  const fields={'Archive Status':'ARCHIVED','Archived At':new Date(job.createdAt),'Archive Batch ID':job.batchId,'Archived Task Rows':job.taskNames.reduce((n,name)=>n+job.counts[name],0),'Archive File ID':run.targetId};
  Object.keys(fields).forEach(function(name){if(request.table.map[name]===undefined)throw Error('REQUESTS 백업 헤더 누락: '+name);requests.push(archiveResumeCellV2_(request.table.sheet,request.record.rowNumber,request.table.map[name]+1,fields[name]));});
  const lock=LockService.getScriptLock();if(!lock.tryLock(100))throw archiveResumeBusyV2_('세션 저장 중: 최종 확정을 다음 실행으로 보류합니다.');
  try{
    const freshRequest=archiveResumeRequestV2_(job);if(freshRequest.record.rowNumber!==request.record.rowNumber||JSON.stringify(freshRequest.record.values)!==JSON.stringify(request.record.values))throw Error('요청 행이 변경되어 완료를 보류합니다.');
    for(const name of job.taskNames)if(archiveResumeSourceV2_(name,job.id).rows.length)throw Error('남은 Task가 있어 완료를 보류합니다.');
    const latest=archiveResumeSourceV2_('WORK_SESSIONS',job.id);if(JSON.stringify(latest?latest.rows:[])!==JSON.stringify(sessions?sessions.rows:[]))throw Error('세션 행 위치가 변경되었습니다.');
    if(latest){const records=archiveResumeRecordsV2_(latest,latest.rows,job.id);if(JSON.stringify(records.map(r=>r.values))!==JSON.stringify(current.map(r=>r.values)))throw Error('세션 내용이 변경되었습니다.');requests.push(...archiveResumeDeleteRequestsV2_(latest.sheet,latest.rows));}
    Sheets.Spreadsheets.batchUpdate({requests:requests},run.sourceId);
  }finally{lock.releaseLock();}
  job.phase='DONE';
}
function archiveResumeCandidatesV2_(cutoff,allowed){
  const sheet=getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS),h=getHeaderMap_(sheet),n=sheet.getLastRow()-1,fields=['Request ID','Status','Completed At','Cancelled At','Archive Status'],columns={};
  fields.forEach(function(name){if(!h[name])throw Error('REQUESTS 필수 헤더 누락: '+name);columns[name]=n>0?sheet.getRange(2,h[name],n,1).getValues():[];});
  const result=[],seen={};for(let i=0;i<n;i++){const id=String(columns['Request ID'][i][0]||'').trim();if(!id)continue;if(seen[id])throw Error('중복 Request ID: 백업을 보류합니다.');seen[id]=true;
    if((!allowed||allowed.indexOf(id)>=0)&&archiveEligibleV1_(columns.Status[i][0],columns['Completed At'][i][0],columns['Cancelled At'][i][0],columns['Archive Status'][i][0],cutoff))result.push({id:id});
  }return result;
}
function archiveResumeStepV2_(run){
  let job=run.job;
  if(!job){
    const candidates=archiveResumeCandidatesV2_(new Date(run.cutoff),run.onlyRequestIds);
    run.remaining=candidates.length;
    if(!candidates.length){run.active=false;run.status='COMPLETE';return;}
    const id=candidates[0].id;job=run.job={id:id,batchId:run.id+'-'+id,createdAt:new Date().toISOString(),phase:'COPY',copyEncodingV5:true,index:0,cursor:0,chunk:ARCHIVE_RESUME_V2.ROWS,counts:{},names:[APP_CONFIG.SHEETS.REQUESTS,APP_CONFIG.SHEETS.TASKS,APP_CONFIG.SHEETS.CRAWL_TASKS,APP_CONFIG.SHEETS.TASKS_2P,'WORK_SESSIONS'],taskNames:[APP_CONFIG.SHEETS.TASKS,APP_CONFIG.SHEETS.CRAWL_TASKS,APP_CONFIG.SHEETS.TASKS_2P]};job.guard=archiveResumeRequestV2_(job).guard;return;
  }
  archiveResumeRequestV2_(job);
  const target=SpreadsheetApp.openById(run.targetId);
  if(job.phase==='COPY'){
    archiveResumeCopyV2_(run,job,target);if(job.index>=job.names.length){job.phase='DELETE_TASKS';job.index=0;}
  }else if(job.phase==='DELETE_TASKS'){
    archiveResumeDeleteV2_(run,job,target);if(job.index>=job.taskNames.length){job.phase='FINALIZE';job.index=0;}
  }else if(job.phase==='FINALIZE')archiveResumeFinalizeV2_(run,job,target);
  else if(job.phase==='DONE'){
    // Best-effort log; durable completion is the atomic REQUESTS/aggregate commit.
    try{appendArchiveLogV1_(target,ARCHIVE_CONFIG_V1.RUN_LOG,['Timestamp','Batch ID','Status','Candidate Requests','Archived Task Rows','Message'],[new Date(),job.batchId,'SUCCESS',1,job.taskNames.reduce((n,name)=>n+job.counts[name],0),job.id+' 이어하기 완료']);}catch(e){console.error(String(e));}
    run.completed++;run.remaining=Math.max(0,(run.remaining||1)-1);run.taskRows+=job.taskNames.reduce((n,name)=>n+job.counts[name],0);run.job=null;
  }else throw Error('알 수 없는 백업 단계');
}
function processArchiveResumeV2_(){
  archiveTaskCopyCacheV4_=null;
  let run=archiveResumeReadV2_(ARCHIVE_RESUME_V2.RUN);if(run&&!run.active)archiveResumeFinishTimersV2_(run);if(!run||!run.active||run.nextAt>Date.now())return {ok:true,state:run?run.status:'IDLE'};
  const lock=LockService.getDocumentLock();if(!lock.tryLock(100))return {ok:true,state:'WAITING'};
  const started=Date.now();
  try{
    run=archiveResumeReadV2_(ARCHIVE_RESUME_V2.RUN);if(!run||!run.active)return {ok:true,state:'IDLE'};
    if(run.inFlight){run.failures=(run.failures||0)+1;if(run.job)run.job.chunk=Math.max(1,Math.floor(run.job.chunk/2));if(run.failures>=5){run.lastError='동일 단계에서 실행 중단이 5회 연속 발생했습니다. 다음 일일 백업에서 재시도합니다.';archiveResumePauseV2_(run);return {ok:false,state:run.status};}}
    run.status='RUNNING';run.lastError='';
    while(run.active&&Date.now()-started<ARCHIVE_RESUME_V2.BUDGET){
      if(run.job&&run.job.phase==='COPY'&&!run.job.copyEncodingV5){run.job.cursor=0;run.job.taskCopyV3='';run.job.chunk=ARCHIVE_RESUME_V2.ROWS;run.job.copyEncodingV5=true;}
      if(run.job&&!run.job.copySizingV3){if(!run.inFlight)run.job.chunk=ARCHIVE_RESUME_V2.ROWS;run.job.copySizingV3=true;}
      run.inFlight=true;archiveResumeSaveV2_(run); // Durable before each external effect.
      archiveResumeStepV2_(run);
      run.inFlight=false;run.failures=0;run.nextAt=0;archiveResumeSaveV2_(run);
    }
    if(run.active)run.status='WAITING';archiveResumeSaveV2_(run);
  }catch(error){
    // Reload the last checkpoint: effects may have succeeded, but uncommitted
    // in-memory cursors/phase changes must never skip verification on retry.
    run=archiveResumeReadV2_(ARCHIVE_RESUME_V2.RUN)||run;
    run.status='RETRYING';run.lastError=String(error&&error.message||error).slice(0,500);run.failures=(run.failures||0)+(error.archiveLockBusy?0:1);run.nextAt=Date.now()+Math.min(300000,60000*run.failures);if(run.job)run.job.chunk=Math.max(1,Math.floor(run.job.chunk/2));run.inFlight=false;if(run.failures>=5)archiveResumePauseV2_(run);else archiveResumeSaveV2_(run);console.error('Archive retry: '+run.lastError);
  }finally{archiveTaskCopyCacheV4_=null;archiveResumeProgressV2_(run);lock.releaseLock();}
  if(!run.active)archiveResumeFinishTimersV2_(run);
  return {ok:run.status!=='RETRYING'&&run.status!=='NEEDS_ATTENTION',state:run.status,completed:run.completed,error:run.lastError||''};
}

// Schedule links before removing recovery: interrupted handoff retries safely.
function archiveResumeFinishTimersV2_(run){
  if(run.status==='NEEDS_ATTENTION'){archiveResumeStopTimerV2_();return;}
  if(run.status!=='COMPLETE')return;
  try{if(typeof archiveLinkScheduleV1_==='function')archiveLinkScheduleV1_();}
  catch(error){run.handoffFailures=(run.handoffFailures||0)+1;run.lastError='백업 링크 예약: '+String(error);archiveResumeSaveV2_(run);archiveResumeProgressV2_(run);if(run.handoffFailures>=5)archiveResumeStopTimerV2_();return;}
  archiveResumeStopTimerV2_();
}

function archiveResumeBusyV2_(message){const error=Error(message);error.archiveLockBusy=true;return error;}
function archiveResumeStopTimerV2_(){ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()===ARCHIVE_RESUME_V2.HANDLER)ScriptApp.deleteTrigger(t);});}
function archiveResumePauseV2_(run){run.active=false;run.status='NEEDS_ATTENTION';run.inFlight=false;archiveResumeSaveV2_(run);archiveResumeStopTimerV2_();}

// Task copy reconciles stable IDs before trusting a saved positional cursor.
function archiveResumeCopyTasksV3_(run,job,target,name,table){
  const started=Date.now(),idColumn=table.map['Internal Task ID'];
  if(idColumn===undefined)throw Error(name+' Internal Task ID 헤더 누락');
  const token=run.id+'|'+job.id+'|'+name;let cache=archiveTaskCopyCacheV4_;
  const reused=!!(cache&&cache.token===token);
  if(!reused){
    const ids=archiveResumeReadRowsV2_(table.sheet,table.rows,idColumn+1,1),keys=[],seen={};
    ids.forEach(function(r){const id=String(r.values[0]||'').trim(),key=name+'|'+job.id+'|'+id;if(!id||seen[key])throw Error(name+' Internal Task ID 누락/중복: 원본 삭제를 보류합니다.');seen[key]=true;keys.push(key);});
    cache={token:token,table:table,ids:ids,keys:keys,seen:seen,dest:archiveResumeDestinationV2_(target,name,table.headers,job.id,run.sourceId)};
    archiveTaskCopyCacheV4_=cache;
  }
  const ids=cache.ids,keys=cache.keys,seen=cache.seen,dest=cache.dest;
  Object.keys(dest.keys).forEach(function(key){if(!seen[key])throw Error(name+' 원본에 없는 백업 Task ID: 확인이 필요합니다.');});
  let previous=0,missing=-1;
  keys.forEach(function(key,i){const row=dest.keys[key];if(!row){if(missing<0)missing=i;return;}if(row<=previous)throw Error(name+' 기존 백업 Task 순서가 원본과 다릅니다.');previous=row;});
  if(job.taskCopyV3!==name||job.counts[name]!==keys.length){job.cursor=0;job.taskCopyV3=name;}
  job.counts[name]=keys.length;
  if(missing>=0&&missing<job.cursor)job.cursor=missing;
  if(job.cursor>=keys.length){if(missing>=0)job.cursor=missing;else{job.index++;job.cursor=0;job.taskCopyV3='';archiveTaskCopyCacheV4_=null;return;}}
  const begin=job.cursor,isNew=!dest.keys[keys[begin]];let end=begin;
  while(end<keys.length&&end-begin<job.chunk&&(!dest.keys[keys[end]])===isNew)end++;
  const transferStarted=Date.now();
  const records=archiveResumeRecordsV2_(table,table.rows.slice(begin,end),job.id);
  const existingRows=isNew?[]:keys.slice(begin,end).map(k=>dest.keys[k]);
  const existing=archiveResumeReadRowsV2_(dest.sheet,existingRows,1,dest.headers.length);
  existing.forEach(function(r,i){if(r.values[dest.map['Archive Key']]!==keys[begin+i]||String(r.values[dest.map['Source Spreadsheet ID']])!==run.sourceId)throw Error('백업 Task 위치가 변경되어 쓰기를 보류합니다.');});
  const values=records.map(function(r,i){
    if(String(r.values[idColumn]||'').trim()!==String(ids[begin+i].values[0]||'').trim()||String(r.values[table.map['Request ID']]||'').trim()!==job.id)throw Error('원본 Task 위치가 변경되어 복사를 보류합니다.');
    const old=isNew?null:existing[i].values,row=Array(dest.headers.length).fill('');
    table.headers.forEach((h,c)=>{if(h&&Object.prototype.hasOwnProperty.call(dest.map,h))row[dest.map[h]]=r.values[c];});
    const meta={'Archive Batch ID':job.batchId,'Archived At':new Date(job.createdAt),'Source Spreadsheet ID':run.sourceId,'Source Sheet':name,'Schema Version':archiveSchemaVersionV1_(table.headers),'Extra Data JSON':''};
    Object.keys(meta).forEach(h=>row[dest.map[h]]=old?old[dest.map[h]]:meta[h]);row[dest.map['Archive Key']]=keys[begin+i];row[dest.map['Source Row Number']]=r.rowNumber;return row;
  });
  let writtenRows,insertAt=0;
  if(isNew){
    const next=keys.slice(end).find(k=>dest.keys[k]),at=next?dest.keys[next]:dest.sheet.getLastRow()+1;
    if(next){
      if(String(dest.sheet.getRange(at,dest.map['Archive Key']+1).getValue())!==next||String(dest.sheet.getRange(at,dest.map['Source Spreadsheet ID']+1).getValue())!==run.sourceId)throw Error('백업 삽입 위치가 변경되었습니다.');
      insertAt=at;
      // Insert + fill atomically. A lost response is reconciled by key next time.
      const tz=target.getSpreadsheetTimeZone();
      Sheets.Spreadsheets.batchUpdate({requests:[{insertDimension:{range:{sheetId:dest.sheet.getSheetId(),dimension:'ROWS',startIndex:at-1,endIndex:at-1+values.length},inheritFromBefore:false}},{updateCells:{start:{sheetId:dest.sheet.getSheetId(),rowIndex:at-1,columnIndex:0},rows:values.map(row=>({values:row.map(v=>archiveResumeTypedCellV5_(v,tz))})),fields:'userEnteredValue,userEnteredFormat.numberFormat'}}]},target.getId());
    }else{
      const required=at+values.length-1;if(required>dest.sheet.getMaxRows())dest.sheet.insertRowsAfter(dest.sheet.getMaxRows(),required-dest.sheet.getMaxRows());archiveResumeWriteValuesV5_(target,dest.sheet,at,values);
    }
    writtenRows=values.map((r,i)=>at+i);
  }else{
    writtenRows=existingRows;
    // Batch contiguous updates instead of one write/read per existing Task.
    let i=0;while(i<values.length){let j=i+1;while(j<values.length&&writtenRows[j]===writtenRows[j-1]+1)j++;
      if(records.slice(i,j).some((r,n)=>!archiveResumeSameV2_(r.values,existing[i+n].values,table.headers,dest.map)))archiveResumeWriteValuesV5_(target,dest.sheet,writtenRows[i],values.slice(i,j));i=j;}
  }
  SpreadsheetApp.flush();
  const saved=archiveResumeReadRowsV2_(dest.sheet,writtenRows,1,dest.headers.length);
  records.forEach(function(r,i){if(saved[i].values[dest.map['Archive Key']]!==keys[begin+i]||String(saved[i].values[dest.map['Source Spreadsheet ID']])!==run.sourceId||!archiveResumeSameV2_(r.values,saved[i].values,table.headers,dest.map))throw Error(name+' 복사 후 내용 검증 실패: 진행 위치를 유지합니다.');});
  if(isNew){
    if(insertAt){Object.keys(dest.keys).forEach(k=>{if(dest.keys[k]>=insertAt)dest.keys[k]+=values.length;});dest.rows=dest.rows.map(r=>r>=insertAt?r+values.length:r);}
    writtenRows.forEach((r,i)=>{dest.keys[keys[begin+i]]=r;dest.rows.push(r);});dest.rows.sort((a,b)=>a-b);
  }
  job.cursor=end;
  const elapsed=Date.now()-started,transferMs=Date.now()-transferStarted;if(transferMs>15000)job.chunk=Math.max(10,Math.floor(job.chunk/2));else if(transferMs<5000)job.chunk=Math.min(100,job.chunk+20);
  console.log('archive_task_copy '+JSON.stringify({sheet:name,rows:records.length,inserted:isNew,elapsedMs:elapsed,transferMs:transferMs,indexReused:reused,cursor:job.cursor,total:keys.length,nextChunk:job.chunk}));
}

// Explicit ExtendedValue avoids parsing numeric strings, leading zeroes or '=' text.
function archiveResumeTypedCellV5_(value,tz){
  const cell=sessionCell_(value,tz);
  if(typeof value==='string')cell.userEnteredFormat={numberFormat:{type:'TEXT',pattern:'@'}};
  else if(typeof value==='number')cell.userEnteredFormat={numberFormat:{type:'NUMBER',pattern:'0.###############'}};
  return cell;
}
function archiveResumeWriteValuesV5_(target,sheet,row,values){
  if(!values.length)return;
  const tz=target.getSpreadsheetTimeZone();
  Sheets.Spreadsheets.batchUpdate({requests:[{updateCells:{start:{sheetId:sheet.getSheetId(),rowIndex:row-1,columnIndex:0},rows:values.map(r=>({values:r.map(v=>archiveResumeTypedCellV5_(v,tz))})),fields:'userEnteredValue,userEnteredFormat.numberFormat'}}]},target.getId());
}
