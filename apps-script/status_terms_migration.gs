// 요청 상태 용어 전환과 이전 값 호환을 담당함.
const REQUEST_STATUS_LEGACY_MAP_V2_ = Object.freeze({
  'REQUESTED':'TO DO',
  'IN PROGRESS':'ONGOING',
  'BLOCKED':'ON HOLD',
  'CANCELLED':'DROP'
});

const REQUEST_DROP_REASON_PREFIX_V2_='REQ DROP';
const REQUEST_DROP_REASON_LEGACY_PREFIX_V2_='REQ CANCELLED';

function requestStatusValuesV2_(){
  return [
    APP_CONFIG.STATUS.TO_DO,
    APP_CONFIG.STATUS.ONGOING,
    APP_CONFIG.STATUS.ON_HOLD,
    APP_CONFIG.STATUS.DONE,
    APP_CONFIG.STATUS.DROP
  ];
}

// 공백과 대소문자를 정리한 뒤 이전 상태는 새 상태로 변환함.
function normalizeRequestStatusV2_(value){
  const key=String(value==null?'':value).trim().toUpperCase().replace(/\s+/g,' ');
  return Object.prototype.hasOwnProperty.call(REQUEST_STATUS_LEGACY_MAP_V2_,key)
    ?REQUEST_STATUS_LEGACY_MAP_V2_[key]:key;
}

function isRequestDropReasonV2_(value){
  const reason=String(value==null?'':value);
  return reason.indexOf(REQUEST_DROP_REASON_PREFIX_V2_)===0||reason.indexOf(REQUEST_DROP_REASON_LEGACY_PREFIX_V2_)===0;
}

function normalizeRequestDropReasonV2_(value){
  const raw=value==null?'':String(value);
  return raw.indexOf(REQUEST_DROP_REASON_LEGACY_PREFIX_V2_)===0
    ?REQUEST_DROP_REASON_PREFIX_V2_+raw.slice(REQUEST_DROP_REASON_LEGACY_PREFIX_V2_.length):raw;
}

function requestStatusMigrationSpreadsheetV2_(){
  const spreadsheet=SpreadsheetApp.getActiveSpreadsheet();
  if(!spreadsheet)throw Error('대상 Spreadsheet를 찾을 수 없습니다. 원본 시트의 Apps Script 편집기에서 실행하세요.');
  return spreadsheet;
}

function requestStatusMigrationHeaderV2_(sheet,required){
  const lastColumn=Math.max(sheet.getLastColumn(),1);
  const headers=sheet.getRange(1,1,1,lastColumn).getDisplayValues()[0].map(function(value){return String(value||'').trim();});
  const counts=Object.create(null),map=Object.create(null);
  headers.forEach(function(name,index){if(!name)return;counts[name]=(counts[name]||0)+1;if(!map[name])map[name]=index+1;});
  const duplicate=Object.keys(counts).filter(function(name){return counts[name]>1;});
  if(duplicate.length)throw Error(sheet.getName()+' 중복 헤더: '+duplicate.join(', '));
  const missing=(required||[]).filter(function(name){return !map[name];});
  if(missing.length)throw Error(sheet.getName()+' 필수 헤더 누락: '+missing.join(', '));
  return {headers:headers,map:map};
}

function requestStatusMigrationColumnPlanV2_(sheet,column,convert){
  const rowCount=Math.max(sheet.getLastRow()-1,0),changes=[],unknown=[];
  if(!rowCount)return {sheet:sheet,column:column,changes:changes,unknown:unknown,rowCount:0};
  const values=sheet.getRange(2,column,rowCount,1).getValues();
  values.forEach(function(row,index){
    const before=row[0],after=convert(before,index+2);
    if(String(before==null?'':before)!==String(after==null?'':after))changes.push({row:index+2,before:before,after:after});
  });
  return {sheet:sheet,column:column,changes:changes,unknown:unknown,rowCount:rowCount};
}

function requestStatusMigrationStatusPlanV2_(sheet,column){
  const canonical=requestStatusValuesV2_(),allowed=Object.create(null),plan=requestStatusMigrationColumnPlanV2_(sheet,column,normalizeRequestStatusV2_);
  canonical.forEach(function(value){allowed[value]=true;});
  if(plan.rowCount){
    const values=sheet.getRange(2,column,plan.rowCount,1).getDisplayValues();
    values.forEach(function(row,index){
      const raw=String(row[0]||''),normalized=normalizeRequestStatusV2_(raw),value=String(normalized||'').trim();
      if(value&&!allowed[value])plan.unknown.push({row:index+2,value:raw});
    });
  }
  return plan;
}

function requestStatusMigrationConfigPlanV2_(sheet,column){
  const canonical=requestStatusValuesV2_(),rowCount=Math.max(sheet.getLastRow()-1,canonical.length),current=rowCount?sheet.getRange(2,column,rowCount,1).getValues():[],changes=[];
  for(let index=0;index<rowCount;index+=1){
    const before=current[index]?current[index][0]:'',after=index<canonical.length?canonical[index]:'';
    if(String(before==null?'':before)!==String(after))changes.push({row:index+2,before:before,after:after});
  }
  return {sheet:sheet,column:column,changes:changes,unknown:[],rowCount:rowCount};
}

function requestStatusMigrationReasonPlanV2_(sheet,column){
  return requestStatusMigrationColumnPlanV2_(sheet,column,normalizeRequestDropReasonV2_);
}

function requestStatusMigrationArchivePlanV2_(){
  const id=typeof archiveSpreadsheetIdV1_==='function'?archiveSpreadsheetIdV1_():'';
  if(!id)return {configured:false,spreadsheetId:'',sheet:null,status:null,taskReasons:[]};
  const spreadsheet=SpreadsheetApp.openById(id),destinations=typeof ARCHIVE_CONFIG_V1!=='undefined'&&ARCHIVE_CONFIG_V1.DESTINATIONS?ARCHIVE_CONFIG_V1.DESTINATIONS:{},requestName=destinations.REQUESTS||'ARCHIVE_REQUESTS',sheet=spreadsheet.getSheetByName(requestName),taskReasons=[];
  [APP_CONFIG.SHEETS.TASKS,APP_CONFIG.SHEETS.CRAWL_TASKS,APP_CONFIG.SHEETS.TASKS_2P].forEach(function(sourceName){
    const archiveName=destinations[sourceName]||('ARCHIVE_'+sourceName),taskSheet=spreadsheet.getSheetByName(archiveName);
    if(!taskSheet)return;
    const header=requestStatusMigrationHeaderV2_(taskSheet,['Invalid Reason']);
    taskReasons.push({name:archiveName,plan:requestStatusMigrationReasonPlanV2_(taskSheet,header.map['Invalid Reason'])});
  });
  if(!sheet)return {configured:true,spreadsheetId:id,sheet:null,status:null,taskReasons:taskReasons};
  const header=requestStatusMigrationHeaderV2_(sheet,['Status']);
  return {configured:true,spreadsheetId:id,sheet:sheet,status:requestStatusMigrationStatusPlanV2_(sheet,header.map.Status),taskReasons:taskReasons};
}

function requestStatusMigrationPlanV2_(){
  const spreadsheet=requestStatusMigrationSpreadsheetV2_(),requests=spreadsheet.getSheetByName(APP_CONFIG.SHEETS.REQUESTS),config=spreadsheet.getSheetByName(APP_CONFIG.SHEETS.CONFIG);
  if(!requests)throw Error('필수 시트를 찾을 수 없습니다: '+APP_CONFIG.SHEETS.REQUESTS);
  if(!config)throw Error('필수 시트를 찾을 수 없습니다: '+APP_CONFIG.SHEETS.CONFIG);
  const requestHeader=requestStatusMigrationHeaderV2_(requests,['Status']),configHeader=requestStatusMigrationHeaderV2_(config,['Status']),taskReasons=[];
  [APP_CONFIG.SHEETS.TASKS,APP_CONFIG.SHEETS.CRAWL_TASKS,APP_CONFIG.SHEETS.TASKS_2P].forEach(function(name){
    const sheet=spreadsheet.getSheetByName(name);
    if(!sheet)throw Error('필수 시트를 찾을 수 없습니다: '+name);
    const header=requestStatusMigrationHeaderV2_(sheet,['Invalid Reason']);
    taskReasons.push({name:name,plan:requestStatusMigrationReasonPlanV2_(sheet,header.map['Invalid Reason'])});
  });
  const plan={
    spreadsheet:spreadsheet,
    requests:requestStatusMigrationStatusPlanV2_(requests,requestHeader.map.Status),
    config:requestStatusMigrationConfigPlanV2_(config,configHeader.map.Status),
    taskReasons:taskReasons,
    archive:requestStatusMigrationArchivePlanV2_()
  };
  const unknown=plan.requests.unknown.slice();
  if(plan.archive.status)plan.archive.status.unknown.forEach(function(item){unknown.push({row:item.row,value:item.value,source:'ARCHIVE_REQUESTS'});});
  plan.unknown=unknown;
  return plan;
}

function requestStatusMigrationSummaryV2_(plan){
  const taskSheets={},archiveTaskSheets={},taskReasonChanges=plan.taskReasons.reduce(function(total,item){taskSheets[item.name]=item.plan.changes.length;return total+item.plan.changes.length;},0),archiveTaskReasonChanges=plan.archive.taskReasons.reduce(function(total,item){archiveTaskSheets[item.name]=item.plan.changes.length;return total+item.plan.changes.length;},0);
  return {
    ok:plan.unknown.length===0,
    requestStatusChanges:plan.requests.changes.length,
    configChanges:plan.config.changes.length,
    taskReasonChanges:taskReasonChanges,
    taskSheets:taskSheets,
    archiveConfigured:plan.archive.configured,
    archiveSheetFound:!!plan.archive.sheet,
    archiveStatusChanges:plan.archive.status?plan.archive.status.changes.length:0,
    archiveTaskReasonChanges:archiveTaskReasonChanges,
    archiveTaskSheets:archiveTaskSheets,
    unknownStatuses:plan.unknown.map(function(item){return {source:item.source||APP_CONFIG.SHEETS.REQUESTS,row:item.row,value:item.value};})
  };
}

function requestStatusMigrationApplyV2_(plan){
  if(!plan.changes.length)return 0;
  const changes=plan.changes.slice().sort(function(a,b){return a.row-b.row;}),groups=[];
  changes.forEach(function(change){
    const last=groups[groups.length-1];
    if(last&&change.row===last.start+last.values.length)last.values.push([change.after]);
    else groups.push({start:change.row,values:[[change.after]]});
  });
  groups.forEach(function(group){
    if(plan.sheet===requestStatusMigrationSpreadsheetV2_().getSheetByName(APP_CONFIG.SHEETS.REQUESTS)&&group.values.some(v=>v[0]===APP_CONFIG.STATUS.DONE))writeRequestStatuses_(plan.sheet,group.start,plan.column,group.values.map(v=>v[0]));
    else plan.sheet.getRange(group.start,plan.column,group.values.length,1).setValues(group.values);
  });
  return changes.length;
}

function requestStatusMigrationAssertKnownV2_(plan){
  if(!plan.unknown.length)return;
  const sample=plan.unknown.slice(0,10).map(function(item){return (item.source||APP_CONFIG.SHEETS.REQUESTS)+'!'+item.row+'='+item.value;});
  throw Error('알 수 없는 요청 상태가 있어 전환하지 않았습니다: '+sample.join(', '));
}

// 쓰기 없이 상태 전환 대상과 스키마를 점검함.
function previewRequestStatusTermsMigrationV2(){
  const summary=requestStatusMigrationSummaryV2_(requestStatusMigrationPlanV2_());
  console.log(JSON.stringify(summary));
  return summary;
}

// 이전 상태값과 취소 마커를 새 용어로 한 번에 전환함. 반복 실행해도 결과가 달라지지 않음.
function migrateRequestStatusTermsV2(){
  const scriptLock=LockService.getScriptLock();
  if(!scriptLock||!scriptLock.tryLock(30000))throw Error('다른 시트 변경 작업이 진행 중입니다. 잠시 후 다시 실행하세요.');
  let documentLock=null,documentLocked=false;
  try{
    documentLock=typeof LockService.getDocumentLock==='function'?LockService.getDocumentLock():null;
    if(documentLock){
      if(!documentLock.tryLock(30000))throw Error('백업 작업이 진행 중입니다. 잠시 후 다시 실행하세요.');
      documentLocked=true;
    }
    const plan=requestStatusMigrationPlanV2_();
    requestStatusMigrationAssertKnownV2_(plan);
    // 기존 strict dropdown이 새 값을 거부하지 않도록 canonical validation을 먼저 적용함.
    configureRequestsSheet_();
    requestStatusMigrationApplyV2_(plan.requests);
    requestStatusMigrationApplyV2_(plan.config);
    plan.taskReasons.forEach(function(item){requestStatusMigrationApplyV2_(item.plan);});
    if(plan.archive.status)requestStatusMigrationApplyV2_(plan.archive.status);
    plan.archive.taskReasons.forEach(function(item){requestStatusMigrationApplyV2_(item.plan);});
    if(typeof applySemanticStatusColorsV4_==='function')applySemanticStatusColorsV4_();
    if(typeof bumpMyWorkSidebarRevisionV2_==='function')bumpMyWorkSidebarRevisionV2_([]);
    if(typeof refreshAllViews==='function')refreshAllViews();
    const result=requestStatusMigrationSummaryV2_(plan);
    result.ok=true;
    console.log(JSON.stringify(result));
    return result;
  }finally{
    if(documentLocked)documentLock.releaseLock();
    scriptLock.releaseLock();
  }
}

// 원본, CONFIG, Task 취소 마커와 설정된 아카이브에 이전 용어가 남지 않았는지 확인함.
function verifyRequestStatusTermsV2(){
  const plan=requestStatusMigrationPlanV2_(),summary=requestStatusMigrationSummaryV2_(plan),remaining=summary.requestStatusChanges+summary.configChanges+summary.taskReasonChanges+summary.archiveStatusChanges+summary.archiveTaskReasonChanges;
  requestStatusMigrationAssertKnownV2_(plan);
  if(remaining)throw Error('상태 용어 전환이 끝나지 않았습니다. 남은 변경 '+remaining+'건');
  summary.ok=true;
  console.log(JSON.stringify(summary));
  return summary;
}
