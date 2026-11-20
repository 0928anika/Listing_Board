// Archive navigation is a separate, bounded worker. No archive data copy/delete calls.
const ARCHIVE_LINKS_V1 = Object.freeze({STATE:'ARCHIVE_LINK_MAINTENANCE_V1',QUEUE:'ARCHIVE_LINK_QUEUE_V1',BATCH:10,MAX_QUEUE:10,BUDGET_MS:20000,HANDLER:'archiveLinkMaintenanceV1'});
function archiveLinkPropsV1_(){return PropertiesService.getDocumentProperties();}
function archiveLinkStateV1_(){return JSON.parse(archiveLinkPropsV1_().getProperty(ARCHIVE_LINKS_V1.STATE)||'{}');}
function archiveLinkQueueV1_(){const q=JSON.parse(archiveLinkPropsV1_().getProperty(ARCHIVE_LINKS_V1.QUEUE)||'[]');if(!Array.isArray(q))throw Error('백업 링크 대기 목록이 잘못되었습니다.');return q;}
function archiveLinkSaveQueueV1_(queue){archiveLinkPropsV1_().setProperty(ARCHIVE_LINKS_V1.QUEUE,JSON.stringify(queue));}
function archiveLinkUrlV1_(fileId,sheetId,viewId){return 'https://docs.google.com/spreadsheets/d/'+fileId+'/edit#gid='+sheetId+'&fvid='+viewId;}
function archiveLinkReadyV1_(url,fileId){return !!fileId&&String(url||'').indexOf('https://docs.google.com/spreadsheets/d/'+fileId+'/edit#gid=')===0&&/\&fvid=\d+$/.test(String(url));}
function archiveLinkArchivedV1_(row,h){return !!h['Archive Status']&&String(row[h['Archive Status']-1]||'').toUpperCase()==='ARCHIVED';}
// The ARCHIVED rows with an unconverted link are the durable discovery backlog.
// Only in-flight/retry items occupy Properties, bounded independently of history size.
function archiveLinkSnapshotV1_(source){
  const sheet=source.getSheetByName(APP_CONFIG.SHEETS.REQUESTS);if(!sheet)throw Error('REQUESTS 시트가 없습니다.');
  const h=getHeaderMap_(sheet),names=['Request ID','Request Type','Archive Status','Archive File ID'],n=Math.max(0,sheet.getLastRow()-1),columns={};
  names.forEach(function(name){if(!h[name])throw Error('REQUESTS 필수 헤더 누락: '+name);});
  names.concat(h['크롤링 필요']?['크롤링 필요']:[]).forEach(function(name){columns[name]=n?sheet.getRange(2,h[name],n,1).getValues():[];});
  const rich=n?sheet.getRange(2,h['Request ID'],n,1).getRichTextValues():[],records=[],seen={};
  for(let i=0;i<n;i++){
    const id=String(columns['Request ID'][i][0]||'').trim();if(!id)continue;
    if(seen[id])throw Error('중복 Request ID가 있어 백업 링크 변경을 중단했습니다.');seen[id]=true;
    if(String(columns['Archive Status'][i][0]||'').toUpperCase()!=='ARCHIVED')continue;
    const fileId=String(columns['Archive File ID'][i][0]||'').trim(),type=String(columns['Request Type'][i][0]||''),crawl=columns['크롤링 필요']&&columns['크롤링 필요'][i][0]===true;
    const sourceName=type==='2P Lite 전환'?APP_CONFIG.SHEETS.TASKS_2P:crawl?APP_CONFIG.SHEETS.CRAWL_TASKS:APP_CONFIG.SHEETS.TASKS;
    records.push({id:id,row:i+2,fileId:fileId,sourceName:sourceName,url:rich[i][0]&&rich[i][0].getLinkUrl()||''});
  }
  return {sheet:sheet,headers:h,records:records};
}
function installArchiveLinkMaintenanceV1(){
  if(!SpreadsheetApp.getActiveSpreadsheet())throw Error('연결된 스프레드시트의 Apps Script 편집기에서 실행하세요.');
  const admin=assertDashboardAdmin_();if(admin!==dashboardPrimaryAdminEmail_())throw Error('기본 관리자만 설치할 수 있습니다.');
  const state=archiveLinkStateV1_(),owner=String(Session.getEffectiveUser().getEmail()||'').toLowerCase();
  if(!owner||owner!==String(admin).toLowerCase())throw Error('실행 계정을 확인하세요.');
  if(state.owner&&state.owner!==owner)throw Error('기존 설치자가 백업 링크 트리거를 제거한 뒤 담당자를 변경해 주세요.');
  archiveLinkSnapshotV1_(SpreadsheetApp.getActiveSpreadsheet());
  const existing=ScriptApp.getProjectTriggers().filter(function(t){return t.getHandlerFunction()===ARCHIVE_LINKS_V1.HANDLER;});
  const trigger=existing.shift()||ScriptApp.newTrigger(ARCHIVE_LINKS_V1.HANDLER).timeBased().everyMinutes(5).create();
  existing.forEach(function(t){ScriptApp.deleteTrigger(t);});
  archiveLinkPropsV1_().setProperty(ARCHIVE_LINKS_V1.STATE,JSON.stringify({owner:owner,triggerId:trigger.getUniqueId(),sourceId:SpreadsheetApp.getActiveSpreadsheet().getId(),installedAt:new Date().toISOString()}));
  const run=JSON.parse(archiveLinkPropsV1_().getProperty('ARCHIVE_RESUME_RUN_V2')||'null');
  if(run&&(run.active||run.status==='NEEDS_ATTENTION'))archiveLinkStopTimerV1_();
  console.log('백업 링크 후처리 설정 완료. 백업 중에는 대기하고, 링크 작업이 끝나면 예약을 제거합니다.');
}
function uninstallArchiveLinkMaintenanceV1(){
  archiveLinkAuthorizedV1_(null,true);
  ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()===ARCHIVE_LINKS_V1.HANDLER)ScriptApp.deleteTrigger(t);});
  archiveLinkPropsV1_().deleteProperty(ARCHIVE_LINKS_V1.STATE);
  console.log('백업 링크 후처리 트리거만 제거했습니다. 대기 목록과 데이터는 보존했습니다.');
}
function archiveLinkAuthorizedV1_(event,manual){
  const state=archiveLinkStateV1_(),source=SpreadsheetApp.getActiveSpreadsheet(),owner=String(Session.getEffectiveUser().getEmail()||'').toLowerCase();
  if(!source||!state.sourceId||source.getId()!==state.sourceId||owner!==state.owner)throw Error('백업 링크 후처리 설치 계정/문서를 확인하세요.');
  if(manual)assertDashboardAdmin_();else if(!event||String(event.triggerUid)!==String(state.triggerId))throw Error('설치된 백업 링크 트리거에서만 실행할 수 있습니다.');
  return source;
}
function archiveLinkStopTimerV1_(){
  ScriptApp.getProjectTriggers().forEach(function(t){if(t.getHandlerFunction()===ARCHIVE_LINKS_V1.HANDLER)ScriptApp.deleteTrigger(t);});
  const state=archiveLinkStateV1_();delete state.triggerId;archiveLinkPropsV1_().setProperty(ARCHIVE_LINKS_V1.STATE,JSON.stringify(state));
}
function archiveLinkScheduleV1_(){
  const state=archiveLinkStateV1_();if(!state.owner)return;
  if(state.owner!==String(Session.getEffectiveUser().getEmail()||'').toLowerCase())throw Error('백업과 링크 후처리 설치 계정이 다릅니다.');
  const lock=LockService.getUserLock();if(!lock.tryLock(100))throw Error('링크 예약 중: 다음 실행에서 재시도합니다.');
  try{
    const matches=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===ARCHIVE_LINKS_V1.HANDLER);
    const trigger=matches.shift()||ScriptApp.newTrigger(ARCHIVE_LINKS_V1.HANDLER).timeBased().everyMinutes(5).create();
    matches.forEach(t=>ScriptApp.deleteTrigger(t));state.triggerId=trigger.getUniqueId();archiveLinkPropsV1_().setProperty(ARCHIVE_LINKS_V1.STATE,JSON.stringify(state));
  }finally{lock.releaseLock();}
}
function archiveLinkMaintenanceV1(event){
  const source=archiveLinkAuthorizedV1_(event,false);
  try{const result=processArchiveLinksV1_(source);const state=archiveLinkStateV1_();state.failures=0;archiveLinkPropsV1_().setProperty(ARCHIVE_LINKS_V1.STATE,JSON.stringify(state));return result;}
  catch(error){const state=archiveLinkStateV1_();state.failures=(state.failures||0)+1;state.lastError=String(error).slice(0,500);archiveLinkPropsV1_().setProperty(ARCHIVE_LINKS_V1.STATE,JSON.stringify(state));if(state.failures>=5)archiveLinkStopTimerV1_();throw error;}
}
function archiveLinkResetRetriesV1_(){
  const lock=LockService.getUserLock();if(!lock.tryLock(100))throw Error('링크 작업 중입니다. 잠시 후 백업을 시작하세요.');
  try{const queue=archiveLinkQueueV1_();queue.forEach(q=>{q.attempts=0;q.nextAt=0;});archiveLinkSaveQueueV1_(queue);const state=archiveLinkStateV1_();state.failures=0;state.lastError='';archiveLinkPropsV1_().setProperty(ARCHIVE_LINKS_V1.STATE,JSON.stringify(state));}finally{lock.releaseLock();}
}
function repairArchivedRequestLinksV1(){archiveLinkAuthorizedV1_(null,true);archiveLinkScheduleV1_();const result=processArchiveLinksV1_(archiveLinkAuthorizedV1_(null,true));console.log(JSON.stringify(result));return result;}
function inspectArchivedRequestLinksV1(){
  if(!SpreadsheetApp.getActiveSpreadsheet())throw Error('연결된 편집기에서 실행하세요.');assertDashboardAdmin_();
  const snapshot=archiveLinkSnapshotV1_(SpreadsheetApp.getActiveSpreadsheet()),queue=archiveLinkQueueV1_(),state=archiveLinkStateV1_();
  const result={installed:!!state.owner,scheduled:!!state.triggerId,failures:state.failures||0,lastError:state.lastError||'',owner:state.owner||'',unconverted:snapshot.records.filter(function(r){return !archiveLinkReadyV1_(r.url,r.fileId);}).length,inFlight:queue.length,retries:queue.map(function(q){return {requestId:q.id,createdAt:q.createdAt,attempts:q.attempts||0,lastError:q.error||'',nextAt:q.nextAt||0};})};
  console.log(JSON.stringify(result));return result;
}
function archiveLinkViewsV1_(sourceId,target,items){
  const meta=Sheets.Spreadsheets.get(target.getId(),{fields:'sheets(properties(sheetId),filterViews(filterViewId,title))'});
  if(!meta||!Array.isArray(meta.sheets))throw Error('백업 필터뷰 조회 응답이 불완전합니다.');
  const requests=[],plans=[],headers={};
  items.forEach(function(item){
    const name=ARCHIVE_CONFIG_V1.DESTINATIONS[item.sourceName],sheet=target.getSheetByName(name);if(!sheet)throw Error('백업 Task 시트가 없습니다: '+name);
    const h=headers[name]||(headers[name]=getHeaderMap_(sheet));
    if(!h['Request ID']||!h['Source Spreadsheet ID'])throw Error('백업 Task 필수 헤더가 없습니다: '+name);
    const entry=meta.sheets.find(function(s){return s.properties&&s.properties.sheetId===sheet.getSheetId();});if(!entry)throw Error('백업 시트 필터뷰 메타데이터가 누락되었습니다.');
    const title='[ARCHIVE:'+sourceId+'] '+item.id,matches=(entry.filterViews||[]).filter(function(v){return v.title===title;});
    if(matches.length>1)throw Error('동일한 자동 백업 필터뷰가 중복되었습니다.');
    const criteria={};criteria[h['Request ID']-1]={condition:{type:'TEXT_EQ',values:[{userEnteredValue:item.id}]}};criteria[h['Source Spreadsheet ID']-1]={condition:{type:'TEXT_EQ',values:[{userEnteredValue:sourceId}]}};
    const filter={title:title,range:{sheetId:sheet.getSheetId(),startRowIndex:0,startColumnIndex:0,endColumnIndex:sheet.getLastColumn()},criteria:criteria,filterSpecs:Object.keys(criteria).map(function(c){return {columnIndex:Number(c),filterCriteria:criteria[c]};})};
    if(matches.length){filter.filterViewId=matches[0].filterViewId;requests.push({updateFilterView:{filter:filter,fields:'title,range,criteria,filterSpecs'}});}else requests.push({addFilterView:{filter:filter}});
    plans.push({item:item,sheetId:sheet.getSheetId(),viewId:matches.length?matches[0].filterViewId:null});
  });
  const response=Sheets.Spreadsheets.batchUpdate({requests:requests},target.getId());
  if(!response||!Array.isArray(response.replies)||response.replies.length!==plans.length)throw Error('백업 필터뷰 저장 응답이 불완전합니다.');
  plans.forEach(function(p,i){if(p.viewId===null){const created=response.replies[i]&&response.replies[i].addFilterView&&response.replies[i].addFilterView.filter;if(!created||created.filterViewId==null)throw Error('백업 필터뷰 ID를 확인할 수 없습니다.');p.viewId=created.filterViewId;}p.url=archiveLinkUrlV1_(target.getId(),p.sheetId,p.viewId);});
  return plans;
}
function archiveLinkApplyV1_(source,plans){
  // Only source metadata/link updates briefly share the archive DocumentLock.
  // Never hold it while creating views in another spreadsheet.
  const lock=LockService.getDocumentLock();if(!lock.tryLock(100))throw Error('백업 실행 중: 링크 변경을 다음 실행으로 보류합니다.');
  try{
    const snapshot=archiveLinkSnapshotV1_(source),byId={};snapshot.records.forEach(function(r){byId[r.id]=r;});
    const meta=Sheets.Spreadsheets.get(source.getId(),{fields:'sheets(properties(sheetId),filterViews(filterViewId,title))'});if(!meta||!Array.isArray(meta.sheets))throw Error('운영 필터뷰 조회가 불완전합니다.');
    const taskIds=[APP_CONFIG.SHEETS.TASKS,APP_CONFIG.SHEETS.CRAWL_TASKS,APP_CONFIG.SHEETS.TASKS_2P].map(function(name){const sheet=source.getSheetByName(name);return sheet?sheet.getSheetId():-1;}),deletes=[];
    plans.forEach(function(p){const r=byId[p.item.id];if(!r||r.fileId!==p.item.fileId||r.sourceName!==p.item.sourceName)throw Error('백업 대상 정보가 변경되어 링크 처리를 보류합니다.');
      const cell=snapshot.sheet.getRange(r.row,snapshot.headers['Request ID']),old=cell.getRichTextValue();
      if(cell.getDisplayValue()!==r.id)throw Error('Request ID 행 위치가 변경되었습니다.');
      cell.setRichTextValue((old?old.copy():SpreadsheetApp.newRichTextValue().setText(r.id)).setLinkUrl(p.url).build());
      meta.sheets.forEach(function(s){if(!s.properties||taskIds.indexOf(s.properties.sheetId)<0)return;(s.filterViews||[]).forEach(function(v){if(v.title==='[AUTO] '+r.id)deletes.push({deleteFilterView:{filterId:v.filterViewId}});});});
    });
    SpreadsheetApp.flush();
    plans.forEach(function(p){const cell=snapshot.sheet.getRange(byId[p.item.id].row,snapshot.headers['Request ID']),rich=cell.getRichTextValue();if(cell.getDisplayValue()!==p.item.id||!rich||rich.getLinkUrl()!==p.url)throw Error('백업 링크 저장 확인 실패');});
    if(deletes.length)Sheets.Spreadsheets.batchUpdate({requests:deletes},source.getId());
    return plans.length;
  }finally{lock.releaseLock();}
}
function processArchiveLinksV1_(source){
  // All entry points enforce the installer account, so UserLock serializes this
  // worker without occupying the session ScriptLock or archive DocumentLock.
  const lock=LockService.getUserLock();if(!lock.tryLock(100))return {skipped:true,reason:'worker busy'};
  const started=Date.now();let processed=0,failed=0;
  try{
    const run=JSON.parse(archiveLinkPropsV1_().getProperty('ARCHIVE_RESUME_RUN_V2')||'null');
    if(run&&(run.active||run.status==='NEEDS_ATTENTION')){archiveLinkStopTimerV1_();return {skipped:true,reason:'archive active'};}
    const snapshot=archiveLinkSnapshotV1_(source),byId={};snapshot.records.forEach(function(r){byId[r.id]=r;});
    let queue=archiveLinkQueueV1_().filter(function(q){const r=byId[q.id];return r&&r.fileId===q.fileId&&r.sourceName===q.sourceName;});
    // Re-discover obsolete managed views too, including after an interrupted
    // old refresh recreates a source view or an earlier manual link conversion.
    const sourceMeta=Sheets.Spreadsheets.get(source.getId(),{fields:'sheets(properties(sheetId),filterViews(filterViewId,title))'});
    if(!sourceMeta||!Array.isArray(sourceMeta.sheets))throw Error('운영 필터뷰 조회가 불완전합니다.');
    const taskIds=[APP_CONFIG.SHEETS.TASKS,APP_CONFIG.SHEETS.CRAWL_TASKS,APP_CONFIG.SHEETS.TASKS_2P].map(function(name){const sheet=source.getSheetByName(name);return sheet?sheet.getSheetId():-1;}),oldViews={};
    sourceMeta.sheets.forEach(function(sheet){if(!sheet.properties||taskIds.indexOf(sheet.properties.sheetId)<0)return;(sheet.filterViews||[]).forEach(function(v){if(String(v.title||'').indexOf('[AUTO] ')===0)oldViews[v.title.substring(7)]=true;});});
    const backlog=snapshot.records.filter(r=>!archiveLinkReadyV1_(r.url,r.fileId)||oldViews[r.id]||queue.some(q=>q.id===r.id)).length;
    const known={};queue.forEach(function(q){known[q.id]=true;});
    snapshot.records.forEach(function(r){if(queue.length>=ARCHIVE_LINKS_V1.MAX_QUEUE||known[r.id]||archiveLinkReadyV1_(r.url,r.fileId)&&!oldViews[r.id])return;queue.push({id:r.id,fileId:r.fileId,sourceName:r.sourceName,createdAt:new Date().toISOString(),attempts:0,nextAt:0});known[r.id]=true;});
    archiveLinkSaveQueueV1_(queue); // Before any view/link mutation, including response loss.
    const due=queue.filter(function(q){return (q.attempts||0)<5&&(!q.nextAt||q.nextAt<=Date.now());}).slice(0,ARCHIVE_LINKS_V1.BATCH),groups={};
    due.forEach(function(q){(groups[q.fileId]||(groups[q.fileId]=[])).push(q);});
    function attempt(fileId,items,splitOnError){
      if(Date.now()-started>=ARCHIVE_LINKS_V1.BUDGET_MS)return;
      if(splitOnError){items.forEach(q=>{q.attempts=(q.attempts||0)+1;});archiveLinkSaveQueueV1_(queue);}
      try{
        if(!fileId||fileId===source.getId())throw Error('유효한 백업 File ID가 없습니다.');
        const target=SpreadsheetApp.openById(fileId),plans=archiveLinkViewsV1_(source.getId(),target,items);
        if(Date.now()-started>=ARCHIVE_LINKS_V1.BUDGET_MS)return; // Next run reuses created views.
        processed+=archiveLinkApplyV1_(source,plans);
        queue=queue.filter(function(q){return items.indexOf(q)<0;});
        archiveLinkSaveQueueV1_(queue);
      }catch(error){
        if(splitOnError&&items.length>1){items.forEach(function(item){attempt(fileId,[item],false);});return;}
        failed+=items.length;items.forEach(function(q){q.error=String(error&&error.message||error).slice(0,80);q.nextAt=Date.now()+Math.min(3600000,300000*Math.pow(2,Math.min(q.attempts-1,4)));});
        archiveLinkSaveQueueV1_(queue);
      }
    }
    Object.keys(groups).forEach(function(fileId){attempt(fileId,groups[fileId],true);});
    const result={processed:processed,failed:failed,pending:Math.max(queue.length,backlog-processed),elapsedMs:Date.now()-started};result.blocked=queue.filter(q=>(q.attempts||0)>=5).length;if(!result.pending||queue.length&&queue.every(q=>(q.attempts||0)>=5))archiveLinkStopTimerV1_();console.log('archive_links '+JSON.stringify(result));return result;
  }finally{lock.releaseLock();}
}
