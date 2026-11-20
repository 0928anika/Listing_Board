// DONE 요청의 대용량 상세행을 외부 스프레드시트로 안전하게 이관함.
const ARCHIVE_CONFIG_V1 = Object.freeze({
  PROPERTY_KEY:'LISTING_OPS_ARCHIVE_SPREADSHEET_ID',
  TRIGGER_OWNER_KEY:'LISTING_OPS_ARCHIVE_TRIGGER_OWNER',
  TRIGGER_ID_KEY:'LISTING_OPS_ARCHIVE_TRIGGER_ID',
  TRIGGER_INSTALLED_AT_KEY:'LISTING_OPS_ARCHIVE_TRIGGER_INSTALLED_AT',
  DEFAULT_SPREADSHEET_ID:'',
  RETENTION_MS:7*24*60*60*1000,
  TIMEZONE:'Asia/Seoul',
  TRIGGER_HANDLER:'dailyArchiveCompletedRequestsV1',
  REQUEST_COLUMNS:Object.freeze(['Archive Status','Archived At','Archive Batch ID','Archived Task Rows','Archive File ID']),
  DESTINATIONS:Object.freeze({
    REQUESTS:'ARCHIVE_REQUESTS',
    '3P_TASKS':'ARCHIVE_3P_TASKS',
    CRAWL_3P_TASKS:'ARCHIVE_CRAWL_3P_TASKS',
    '2P_TASKS':'ARCHIVE_2P_TASKS',
    WORK_SESSIONS:'ARCHIVE_WORK_SESSIONS'
  }),
  MANIFEST:'ARCHIVE_MANIFEST',
  RUN_LOG:'BACKUP_RUN_LOG'
});

const ARCHIVE_METADATA_HEADERS_V1 = Object.freeze([
  'Archive Key','Archive Batch ID','Archived At','Source Spreadsheet ID','Source Sheet','Schema Version','Source Row Number','Extra Data JSON'
]);

const LONG_TERM_SESSION_SHEET_V1='LONG_TERM_SESSIONS';
const LONG_TERM_SESSION_COLUMNS_V1=Object.freeze([
  'Aggregate ID','Status','Request ID','Worker','Stage','Work Date','Duration Minutes','Completed VI Delta','Session Count',
  'First Started At','Last Ended At','Active Elapsed Minutes','Request Active Elapsed Minutes','Request Completed At','Aggregated At','Archive Batch ID'
]);

function extractSpreadsheetIdV1_(value){
  const text=String(value||'').trim();
  const match=text.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/)||text.match(/^([a-zA-Z0-9_-]{20,})$/);
  return match?match[1]:'';
}

function archiveSpreadsheetIdV1_(required){
  const id=PropertiesService.getDocumentProperties().getProperty(ARCHIVE_CONFIG_V1.PROPERTY_KEY)||ARCHIVE_CONFIG_V1.DEFAULT_SPREADSHEET_ID;
  if(required&&!id)throw Error('백업 스프레드시트가 설정되지 않았습니다. 관리자 메뉴의 아카이브 시트 설정을 먼저 실행해 주세요.');
  return id;
}

function ensureArchiveRequestColumnsV1_(){
  const sheet=getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS),headers=getHeaderMap_(sheet);
  ARCHIVE_CONFIG_V1.REQUEST_COLUMNS.forEach(function(name){
    if(!headers[name]){
      const column=sheet.getLastColumn()+1;
      sheet.getRange(1,column).setValue(name).setBackground('#263650').setFontColor('#FFFFFF').setFontWeight('bold');
      headers[name]=column;
    }
    sheet.showColumns(headers[name]);
  });
  const rows=Math.max(sheet.getMaxRows()-1,1);
  sheet.getRange(2,headers['Archived At'],rows,1).setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange(2,headers['Archived Task Rows'],rows,1).setNumberFormat('0');
  return headers;
}

function ensureLongTermSessionSheetV1_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();let sheet=ss.getSheetByName(LONG_TERM_SESSION_SHEET_V1);
  if(!sheet)sheet=ss.insertSheet(LONG_TERM_SESSION_SHEET_V1);
  if(sheet.getLastRow()===0){sheet.getRange(1,1,1,LONG_TERM_SESSION_COLUMNS_V1.length).setValues([LONG_TERM_SESSION_COLUMNS_V1.slice()]);styleHeader_(sheet.getRange(1,1,1,LONG_TERM_SESSION_COLUMNS_V1.length));sheet.setFrozenRows(1);}
  archiveHeaderInfoV1_(sheet,LONG_TERM_SESSION_COLUMNS_V1);
  return sheet;
}

function configureArchiveSpreadsheetV1(){
  assertDashboardAdmin_();
  const ui=SpreadsheetApp.getUi(),currentId=archiveSpreadsheetIdV1_();
  let currentName=currentId?'접근 확인 불가':'설정되지 않음',currentUrl=currentId?'\nhttps://docs.google.com/spreadsheets/d/'+currentId+'/edit':'';
  if(currentId)try{currentName=SpreadsheetApp.openById(currentId).getName();}catch(ignored){}
  const response=ui.prompt('백업 아카이브 설정','현재: '+currentName+currentUrl+'\n\n변경할 스프레드시트 URL 또는 ID를 입력하세요.',ui.ButtonSet.OK_CANCEL);
  if(response.getSelectedButton()!==ui.Button.OK)return;
  const id=extractSpreadsheetIdV1_(response.getResponseText());
  if(!id)throw Error('Google 스프레드시트 URL 또는 ID를 확인해 주세요.');
  const target=SpreadsheetApp.openById(id);
  if(target.getId()===SpreadsheetApp.getActiveSpreadsheet().getId())throw Error('운영 스프레드시트와 다른 백업 스프레드시트를 선택해 주세요.');
  target.getSheets();
  PropertiesService.getDocumentProperties().setProperty(ARCHIVE_CONFIG_V1.PROPERTY_KEY,id);
  ensureArchiveRequestColumnsV1_();
  ui.alert('백업 스프레드시트가 변경되었습니다.\n'+target.getName()+'\n'+target.getUrl());
}

function archiveHeaderInfoV1_(sheet,required){
  const last=Math.max(sheet.getLastColumn(),1),headers=sheet.getRange(1,1,1,last).getDisplayValues()[0].map(function(v){return String(v||'').trim();}),seen={};
  headers.forEach(function(name){if(name)seen[name]=(seen[name]||0)+1;});
  const duplicate=Object.keys(seen).filter(function(name){return seen[name]>1;});
  if(duplicate.length)throw Error(sheet.getName()+' 중복 헤더: '+duplicate.join(', '));
  const missing=(required||[]).filter(function(name){return !seen[name];});
  if(missing.length)throw Error(sheet.getName()+' 필수 헤더 누락: '+missing.join(', '));
  const map={};headers.forEach(function(name,i){if(name)map[name]=i;});
  return {headers:headers,map:map};
}

function archiveEligibleV1_(status,completedAt,cancelledAt,archiveStatus,now){
  const normalized=normalizeRequestStatusV2_(status),terminalAt=normalized===APP_CONFIG.STATUS.DONE?completedAt:normalized===APP_CONFIG.STATUS.DROP?cancelledAt:null;
  return terminalAt instanceof Date&&Number.isFinite(terminalAt.getTime())&&now.getTime()-terminalAt.getTime()>=ARCHIVE_CONFIG_V1.RETENTION_MS&&String(archiveStatus||'').trim()!=='ARCHIVED';
}

function collectArchiveCandidatesV1_(now){
  const sheet=getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS),h=ensureArchiveRequestColumnsV1_(),info=archiveHeaderInfoV1_(sheet,['Request ID','Status','Completed At','Cancelled At','Archive Status']);
  const count=Math.max(sheet.getLastRow()-1,0);if(!count)return [];
  return sheet.getRange(2,1,count,sheet.getLastColumn()).getValues().map(function(row,i){
    return {id:String(row[info.map['Request ID']]||'').trim(),row:i+2,values:row,completedAt:row[info.map['Completed At']],cancelledAt:row[info.map['Cancelled At']],archiveStatus:row[info.map['Archive Status']]};
  }).filter(function(record){return record.id&&archiveEligibleV1_(record.values[info.map['Status']],record.completedAt,record.cancelledAt,record.archiveStatus,now);});
}

function previewArchiveCandidatesV1(){
  assertDashboardAdmin_();
  const rows=collectArchiveCandidatesV1_(new Date()),ui=SpreadsheetApp.getUi();
  ui.alert(rows.length?'백업 대상 '+rows.length+'건\n\n'+rows.slice(0,30).map(function(r){return r.id;}).join('\n')+(rows.length>30?'\n외 '+(rows.length-30)+'건':''):'현재 DONE/DROP 후 168시간이 지난 백업 대상이 없습니다.');
  return {count:rows.length,requestIds:rows.map(function(r){return r.id;})};
}

function archiveSchemaVersionV1_(headers){
  const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,headers.join('\u001f'),Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/,'').slice(0,16);
}

function archiveKeyV1_(sheetName,requestId,row,rowNumber,map){
  let identity='';
  if(map['Internal Task ID']!==undefined)identity=String(row[map['Internal Task ID']]||'').trim();
  if(sheetName==='WORK_SESSIONS'&&map['Session ID']!==undefined)identity=String(row[map['Session ID']]||'').trim();
  if(sheetName==='REQUESTS')identity=String(requestId);
  if(!identity){
    const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(row),Utilities.Charset.UTF_8);
    identity=Utilities.base64EncodeWebSafe(bytes).replace(/=+$/,'').slice(0,24);
  }
  return sheetName+'|'+requestId+'|'+identity;
}

function ensureArchiveDestinationV1_(target,name,sourceHeaders){
  let sheet=target.getSheetByName(name);
  if(!sheet)sheet=target.insertSheet(name);
  const desired=ARCHIVE_METADATA_HEADERS_V1.concat(sourceHeaders.filter(function(h){return h&&ARCHIVE_METADATA_HEADERS_V1.indexOf(h)<0;}));
  const existingHeaders=sheet.getLastRow()>0?sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0]:[];
  const neededColumns=existingHeaders.length+desired.filter(function(h){return existingHeaders.indexOf(h)<0;}).length;
  if(neededColumns>sheet.getMaxColumns())sheet.insertColumnsAfter(sheet.getMaxColumns(),neededColumns-sheet.getMaxColumns());
  if(sheet.getLastRow()===0){sheet.getRange(1,1,1,desired.length).setValues([desired]);styleHeader_(sheet.getRange(1,1,1,desired.length));sheet.setFrozenRows(1);}
  else{
    const current=archiveHeaderInfoV1_(sheet,[]),missing=desired.filter(function(h){return !current.map.hasOwnProperty(h);});
    missing.forEach(function(header){sheet.getRange(1,sheet.getLastColumn()+1).setValue(header);});
  }
  return sheet;
}

function appendArchiveRowsV1_(destination,sourceName,sourceHeaders,records,context){
  if(!records.length)return 0;
  const destInfo=archiveHeaderInfoV1_(destination,ARCHIVE_METADATA_HEADERS_V1),existing={};
  if(destination.getLastRow()>1){destination.getRange(2,destInfo.map['Archive Key']+1,destination.getLastRow()-1,1).getDisplayValues().forEach(function(r){if(r[0])existing[r[0]]=true;});}
  const schema=archiveSchemaVersionV1_(sourceHeaders),rows=[],expected={},batchKeys={};
  records.forEach(function(record){
    const key=archiveKeyV1_(sourceName,record.requestId,record.values,record.rowNumber,record.map);
    if(batchKeys[key])throw Error(sourceName+' 백업 고유키 중복: '+key);
    batchKeys[key]=true;expected[key]=true;
    if(existing[key])return;
    const output=Array(destination.getLastColumn()).fill('');
    const metadata={'Archive Key':key,'Archive Batch ID':context.batchId,'Archived At':context.now,'Source Spreadsheet ID':context.sourceId,'Source Sheet':sourceName,'Schema Version':schema,'Source Row Number':record.rowNumber,'Extra Data JSON':''};
    Object.keys(metadata).forEach(function(name){output[destInfo.map[name]]=metadata[name];});
    sourceHeaders.forEach(function(name,i){
      if(!name||!destInfo.map.hasOwnProperty(name)||metadata.hasOwnProperty(name))return;
      let value=record.values[i];
      if(sourceName===APP_CONFIG.SHEETS.REQUESTS&&name==='Status')value=normalizeRequestStatusV2_(value);
      else if(name==='Invalid Reason'&&[APP_CONFIG.SHEETS.TASKS,APP_CONFIG.SHEETS.CRAWL_TASKS,APP_CONFIG.SHEETS.TASKS_2P].indexOf(sourceName)>=0)value=normalizeRequestDropReasonV2_(value);
      output[destInfo.map[name]]=value;
    });
    rows.push(output);existing[key]=true;
  });
  if(rows.length)destination.getRange(destination.getLastRow()+1,1,rows.length,destination.getLastColumn()).setValues(rows);
  const verified={};
  if(destination.getLastRow()>1)destination.getRange(2,destInfo.map['Archive Key']+1,destination.getLastRow()-1,1).getDisplayValues().forEach(function(r){if(r[0])verified[r[0]]=true;});
  const missing=Object.keys(expected).filter(function(key){return !verified[key];});
  if(missing.length)throw Error(sourceName+' 백업 복사 검증 실패: '+missing.length+'행');
  return rows.length;
}

function sourceArchiveRecordsV1_(sheet,requestIds,required){
  const info=archiveHeaderInfoV1_(sheet,required),count=Math.max(sheet.getLastRow()-1,0),wanted={};requestIds.forEach(function(id){wanted[id]=true;});
  if(!count)return {headers:info.headers,records:[]};
  const values=sheet.getRange(2,1,count,sheet.getLastColumn()).getValues(),idIndex=info.map['Request ID'];
  return {headers:info.headers,records:values.map(function(row,i){return {requestId:String(row[idIndex]||'').trim(),rowNumber:i+2,values:row,map:info.map};}).filter(function(r){return wanted[r.requestId];})};
}

function deleteArchivedTaskRowsV1_(sheet,rowNumbers){
  const sorted=rowNumbers.slice().sort(function(a,b){return b-a;});let i=0;
  while(i<sorted.length){let high=sorted[i],low=high;i++;while(i<sorted.length&&sorted[i]===low-1){low=sorted[i];i++;}sheet.deleteRows(low,high-low+1);}
}

function intervalUnionMinutesV1_(intervals){
  const rows=intervals.filter(function(x){return x[0] instanceof Date&&x[1] instanceof Date&&x[1]>=x[0];}).map(function(x){return [x[0].getTime(),x[1].getTime()];}).sort(function(a,b){return a[0]-b[0];});
  if(!rows.length)return 0;let total=0,start=rows[0][0],end=rows[0][1];
  rows.slice(1).forEach(function(x){if(x[0]<=end)end=Math.max(end,x[1]);else{total+=end-start;start=x[0];end=x[1];}});
  return (total+end-start)/60000;
}

function buildLongTermSessionRowsV1_(sessionTable,requestTable,now,batchId,timezone){
  if(!sessionTable)return [];
  const sh=sessionTable.records.length?sessionTable.records[0].map:archiveHeaderInfoV1_(sessionTable.sheet,['Session ID','Worker','Request ID','Stage','Status','Started At','Ended At','Duration Minutes','Completed VI Delta']).map;
  const requestCompleted={};requestTable.records.forEach(function(r){requestCompleted[r.requestId]=r.values[r.map['Completed At']];});
  const groups={},requestIntervals={};
  sessionTable.records.forEach(function(record){
    const r=record.values,status=String(r[sh['Status']]||''),ended=r[sh['Ended At']],stage=String(r[sh['Stage']]||''),worker=String(r[sh['Worker']]||'').trim();
    if(status==='ACTIVE'||!(ended instanceof Date)||!worker||(stage!=='3P'&&stage!=='2P'))return;
    const day=Utilities.formatDate(ended,timezone,'yyyy-MM-dd'),completed=requestCompleted[record.requestId],completionKey=completed instanceof Date?String(completed.getTime()):'NO_COMPLETION',key=[record.requestId,completionKey,worker,stage,day].join('|');
    const g=groups[key]||(groups[key]={id:key,requestId:record.requestId,worker:worker,stage:stage,day:day,minutes:0,delta:0,deltaKnown:true,count:0,starts:[],ends:[],intervals:[]});
    g.minutes+=typeof r[sh['Duration Minutes']]==='number'&&Number.isFinite(r[sh['Duration Minutes']])?Math.max(0,r[sh['Duration Minutes']]):0;
    if(typeof r[sh['Completed VI Delta']]==='number'&&Number.isFinite(r[sh['Completed VI Delta']]))g.delta+=r[sh['Completed VI Delta']];else g.deltaKnown=false;
    g.count++;if(r[sh['Started At']] instanceof Date)g.starts.push(r[sh['Started At']]);g.ends.push(ended);g.intervals.push([r[sh['Started At']],ended]);
    (requestIntervals[record.requestId]||(requestIntervals[record.requestId]=[])).push([r[sh['Started At']],ended]);
  });
  const requestActive={};Object.keys(requestIntervals).forEach(function(k){requestActive[k]=intervalUnionMinutesV1_(requestIntervals[k]);});
  return Object.keys(groups).sort().map(function(key){const g=groups[key],active=Object.prototype.hasOwnProperty.call(requestActive,g.requestId)?requestActive[g.requestId]:'';delete requestActive[g.requestId];
    return [g.id,'PENDING',g.requestId,g.worker,g.stage,g.day,g.minutes,g.deltaKnown?g.delta:'',g.count,
      g.starts.length?new Date(Math.min.apply(null,g.starts.map(function(d){return d.getTime();}))):'',new Date(Math.max.apply(null,g.ends.map(function(d){return d.getTime();}))),
      intervalUnionMinutesV1_(g.intervals),active,requestCompleted[g.requestId]||'',now,batchId];
  });
}

function upsertPendingLongTermSessionsV1_(rows){
  if(!rows.length)return [];
  const sheet=ensureLongTermSessionSheetV1_(),h=archiveHeaderInfoV1_(sheet,LONG_TERM_SESSION_COLUMNS_V1),existing={};
  if(sheet.getLastRow()>1)sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues().forEach(function(r,i){if(r[h.map['Aggregate ID']])existing[String(r[h.map['Aggregate ID']])]=i+2;});
  const touched=[],newRows=[];rows.forEach(function(row){const id=String(row[0]);if(existing[id]){touched.push(existing[id]);return;}newRows.push(row);});
  if(newRows.length){const start=sheet.getLastRow()+1;sheet.getRange(start,1,newRows.length,LONG_TERM_SESSION_COLUMNS_V1.length).setValues(newRows);newRows.forEach(function(_,i){touched.push(start+i);});}
  return {sheet:sheet,rows:touched};
}

function commitLongTermSessionsV1_(pending){
  if(!pending||!pending.rows.length)return;
  const h=archiveHeaderInfoV1_(pending.sheet,LONG_TERM_SESSION_COLUMNS_V1),statusColumn=h.map['Status']+1;
  pending.rows.forEach(function(row){pending.sheet.getRange(row,statusColumn).setValue('COMMITTED');});
}

function appendArchiveLogV1_(target,name,headers,row){
  let sheet=target.getSheetByName(name);if(!sheet)sheet=target.insertSheet(name);
  if(sheet.getLastRow()===0){sheet.getRange(1,1,1,headers.length).setValues([headers]);styleHeader_(sheet.getRange(1,1,1,headers.length));sheet.setFrozenRows(1);}
  sheet.appendRow(row);
}

// Both manual and daily runs enqueue the same resumable worker.
function runArchiveBatchV1_(options){return startArchiveResumeV2_(options||{});}
function dailyArchiveCompletedRequestsV1(){return runArchiveBatchV1_({});}
function runArchiveNowV1(){
  assertDashboardAdmin_();const result=runArchiveBatchV1_({});
  console.log(JSON.stringify(result));
  SpreadsheetApp.getActiveSpreadsheet().toast(result.state==='WAITING'?'백업 실행 중입니다. ARCHIVE_PROGRESS에서 확인하세요.':'백업 이어하기 예약 완료. ARCHIVE_PROGRESS에서 진행 상태를 확인하세요.','백업',10);
  return result;
}

function installDailyArchiveTriggerV1(){
  const currentAdmin=assertDashboardAdmin_(),primaryAdmin=dashboardPrimaryAdminEmail_();
  if(currentAdmin!==primaryAdmin)throw Error('매일 3시 백업 트리거는 기본 관리자만 설치할 수 있습니다.');
  archiveSpreadsheetIdV1_(true);
  const props=PropertiesService.getDocumentProperties(),installer=String(Session.getEffectiveUser().getEmail()||Session.getActiveUser().getEmail()||'').trim().toLowerCase(),owner=String(props.getProperty(ARCHIVE_CONFIG_V1.TRIGGER_OWNER_KEY)||'').trim().toLowerCase();
  if(!installer)throw Error('트리거 설치자 이메일을 확인할 수 없습니다. 관리자 계정으로 다시 실행해 주세요.');
  if(owner&&owner!==installer)throw Error('백업 트리거는 이미 '+owner+' 계정이 설치했습니다. 중복 설치하지 않았습니다. 기존 설치자에게 재설치를 요청해 주세요.');
  let removed=0;ScriptApp.getProjectTriggers().forEach(function(trigger){if(trigger.getHandlerFunction()===ARCHIVE_CONFIG_V1.TRIGGER_HANDLER){ScriptApp.deleteTrigger(trigger);removed++;}});
  const trigger=ScriptApp.newTrigger(ARCHIVE_CONFIG_V1.TRIGGER_HANDLER).timeBased().atHour(3).nearMinute(0).everyDays(1).inTimezone(ARCHIVE_CONFIG_V1.TIMEZONE).create();
  props.setProperties((function(){const values={};values[ARCHIVE_CONFIG_V1.TRIGGER_OWNER_KEY]=installer;values[ARCHIVE_CONFIG_V1.TRIGGER_ID_KEY]=trigger.getUniqueId();values[ARCHIVE_CONFIG_V1.TRIGGER_INSTALLED_AT_KEY]=new Date().toISOString();return values;})(),false);
  ensureArchiveRequestColumnsV1_();
  ensureLongTermSessionSheetV1_();
  return {ok:true,removed:removed,timezone:ARCHIVE_CONFIG_V1.TIMEZONE,owner:installer,triggerId:trigger.getUniqueId()};
}
