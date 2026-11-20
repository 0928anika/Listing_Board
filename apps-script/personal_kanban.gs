// 시트 위에서 넓은 작업자 배정 팝업을 열기.
// 이 화면은 REQUESTS를 읽기만 하며 원본 데이터를 수정하지 않음.
function requireMyWorkTaskHeadersV3_(headers,names,sheetName){
  const missing=names.filter(function(name){return !headers[name];});
  if(missing.length)throw new Error((sheetName?sheetName+' ':'')+'필수 Task 열 누락: '+missing.join(', '));
  return headers;
}

function showPersonalKanbanDialog() {
  const html = HtmlService.createHtmlOutputFromFile('PersonalKanban')
    .setWidth(1240)
    .setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, '작업자 배정');
}

// 작업자 배정 팝업에 필요한 REQUESTS 데이터와 필터 목록을 한 번에 반환함.
// 팝업 안에서 필터링하므로 필터 선택이 다른 사용자 화면에 영향을 주지 않음.
function getPersonalKanbanData(fast) {
  const sheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  const headers = getHeaderMap_(sheet);
  // Read-only startup. Assignment writes already clear legacy validation when needed.
  const rows = getDataRows_(sheet);
  const tz = Session.getScriptTimeZone() || 'Asia/Seoul';
  const now = new Date();

  function value_(row, name) {
    const column = headers[name];
    return column ? row[column - 1] : '';
  }

  function dateText_(value) {
    return value instanceof Date
      ? Utilities.formatDate(value, tz, 'yyyy-MM-dd HH:mm')
      : String(value || '');
  }

  function elapsedHours_(value) {
    const millis = value instanceof Date ? value.getTime() : Date.parse(String(value || ''));
    return isFinite(millis) ? Math.max(0, (now.getTime() - millis) / 3600000) : null;
  }

  const requests = rows.map(function(row) {
    const total = Number(value_(row, 'Total VI')) || 0;
    const done3P = Number(value_(row, '3P Completed VI')) || 0;
    const done2P = Number(value_(row, '2P Completed VI')) || 0;
    const completed = Math.round(total * (Number(value_(row, 'Progress')) || 0));
    const type = String(value_(row, 'Request Type') || '3P Listing');
    const stageCount = type === '3P Listing 후 2P Lite 전환' ? 2 : 1;
    const completedStages = type === '2P Lite 전환' ? done2P
      : type === '3P Listing 후 2P Lite 전환' ? done3P + done2P : done3P;
    const progress = total > 0 ? Math.min(1, Math.max(0, completedStages / (total * stageCount))) : null;
    return {
      requestId: String(value_(row, 'Request ID')),
      requestedAt: dateText_(value_(row, 'Requested At')),
      elapsedHours: elapsedHours_(value_(row, 'Requested At')),
      requester: String(value_(row, 'Requestor') || ''),
      brand: String(value_(row, 'Brand') || ''),
      node: String(value_(row, 'Node') || ''),
      priority: String(value_(row, 'Priority') || 'NORMAL'),
      assignee: String(value_(row, 'Active Assignees') || ''),
      crawlAssignee: String(value_(row, '크롤링 작업자') || ''),
      crawlRequired: value_(row, '크롤링 필요') === true,
      crawlStage: String(value_(row, '크롤링 단계') || ''),
      preparationStatus: String(value_(row, '자료 준비 상태') || ''),
      status: normalizeRequestStatusV2_(value_(row, 'Status')) || APP_CONFIG.STATUS.TO_DO,
      requestType: type,
      totalVi: total,
      completedVi: completed,
      completed3P: done3P,
      completed2P: done2P,
      progress: progress,
      progress3P: total ? Math.min(done3P / total, 1) : 0,
      progress2P: total ? Math.min(done2P / total, 1) : 0,
      barcodeProvided: normalizeBarcodeProvidedV1_(value_(row, 'Barcode Provided')),
      notes: String(value_(row, 'Requestor Notes') || '')
    };
  }).filter(function(item) {
    return item.requestId;
  });

  if(fast===true)requests.forEach(function(r){r.assignmentSummaryVersion=assignmentVersionV3_(r,[]);});
  else attachAssignmentVersionsV3_(requests);

  function unique_(key) {
    return requests.map(function(item) { return item[key]; })
      .filter(function(value) { return value; })
      .filter(function(value, index, values) { return values.indexOf(value) === index; })
      .sort();
  }

  // 남은 3P/2P VI를 요청 참여자 수로 자동 균등분배함.
  const autoAllocations = [];
  requests.forEach(function(request) {
    const workers = String(request.assignee || '').split(',').map(function(value){return value.trim();}).filter(Boolean);
    if (!workers.length) return;
    const has3P = request.requestType === '3P Listing' || request.requestType === '3P Listing 후 2P Lite 전환';
    const has2P = request.requestType === '2P Lite 전환' || request.requestType === '3P Listing 후 2P Lite 전환';
    const remaining3P = has3P ? Math.max(request.totalVi - request.completed3P, 0) : 0;
    const remaining2P = has2P ? Math.max(request.totalVi - request.completed2P, 0) : 0;
    workers.forEach(function(worker, index) {
      autoAllocations.push({
        requestId: request.requestId,
        worker: worker,
        vi3p: Math.floor(remaining3P / workers.length) + (index < remaining3P % workers.length ? 1 : 0),
        vi2p: Math.floor(remaining2P / workers.length) + (index < remaining2P % workers.length ? 1 : 0)
      });
    });
  });
  const workloadMap = {};
  autoAllocations.forEach(function(item) {
    if (!workloadMap[item.worker]) workloadMap[item.worker] = {worker:item.worker, vi3p:0, vi2p:0, requests:0};
    workloadMap[item.worker].vi3p += item.vi3p;
    workloadMap[item.worker].vi2p += item.vi2p;
    workloadMap[item.worker].requests += 1;
  });

  return {
    requests: requests,
    spreadsheetId: SpreadsheetApp.getActiveSpreadsheet().getId(),
    preparationRequired: fast===true,
    assignmentDraftKey: SpreadsheetApp.getActiveSpreadsheet().getId() + ':' + getCurrentUser_(),
    allocations: autoAllocations,
    workload: Object.keys(workloadMap).sort().map(function(key){return workloadMap[key];}),
    options: {
      assignees: getWebConfig_().workers,
      brands: unique_('brand'),
      nodes: unique_('node')
    },
    workSessions: fast===true ? {active:[],sessions:[]} : getWorkSessionSnapshot_(true),
    updatedAt: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm')
  };
}

// 예전에 만든 표 형태 MY_REQUESTS는 숨기고 개인 Kanban 팝업을 기본 방식으로 전환함.
// 데이터는 삭제하지 않으므로 필요할 때 다시 표시할 수 있음.


// 요청 전체 담당자를 변경함.
// 이미 완료된 Task는 작업 이력 보호를 위해 건드리지 않고 미완료 Task만 변경함.
function assignRequestFromKanban(requestId, assignees) {
  requestId = String(requestId || '').trim();
  const selected = (Array.isArray(assignees) ? assignees : [assignees])
    .map(function(value) { return String(value || '').trim(); })
    .filter(function(value, index, values) { return value && values.indexOf(value) === index; });
  if (!requestId) throw new Error('Request ID가 없습니다.');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('다른 사용자가 배정 중입니다. 잠시 후 다시 시도해 주세요.');
  try {
    const requestSheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
    const requestHeaders = getHeaderMap_(requestSheet);
    const requestRow = findRowByRequestId_(requestSheet, requestId);
    if (requestHeaders['크롤링 필요'] && requestSheet.getRange(requestRow,requestHeaders['크롤링 필요']).getValue()===true &&
        String(requestSheet.getRange(requestRow,requestHeaders['크롤링 단계']).getValue())!=='CRAWL_COMPLETED') {
      throw new Error('크롤링 결과 제출 전에는 3P 작업자를 배정할 수 없습니다.');
    }
    const before = String(requestSheet.getRange(requestRow, requestHeaders['Active Assignees']).getValue() || '');
    let after = selected.join(', ');
    requestSheet.getRange(requestRow, requestHeaders['Active Assignees'])
      .clearDataValidations()
      .setValue(after);
    if (selected.length) {
      applyRemainingDistributionV2_(requestId, selected);
    } else {
      const synced=syncRequestSummaries_([requestId]);
      after=String(synced.summaries[String(requestId)]['Active Assignees']||'');
    }
    logActivity_(requestId, 'REQUEST_PARTICIPANTS_CHANGED', before, after);
    return {ok:true, requestId:requestId, assignees:splitWorkersV2_(after)};
  } finally {
    lock.releaseLock();
  }
}

function assignCrawlWorkerFromKanban(requestId,worker) {
return withOpsDocumentLock_(function(){

  requestId=String(requestId||'').trim(); worker=String(worker||'').trim();
  if(!requestId||!worker)throw new Error('크롤링 요청과 작업자를 선택해 주세요.');
  if((getWebConfig_().workers||[]).indexOf(worker)<0)throw new Error('CONFIG에 등록된 작업자를 선택해 주세요.');
  const sheet=getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS),h=ensureRequestIntakeMetadataColumnsV2_();
  const row=findRowByRequestId_(sheet,requestId);
  if(sheet.getRange(row,h['크롤링 필요']).getValue()!==true)throw new Error('크롤링 필요 요청이 아닙니다.');
  if(String(sheet.getRange(row,h['크롤링 단계']).getValue())==='CRAWL_COMPLETED')throw new Error('이미 크롤링이 완료된 요청입니다.');
  const before=String(sheet.getRange(row,h['크롤링 작업자']).getValue()||'');
  sheet.getRange(row,h['크롤링 작업자']).setValue(worker);
  sheet.getRange(row,h['크롤링 단계']).setValue('CRAWL_ASSIGNED');
  sheet.getRange(row,h['자료 준비 상태']).setValue('크롤링 배정됨');
  logActivity_(requestId,'CRAWL_WORKER_ASSIGNED',before,worker);
  bumpMyWorkSidebarRevisionV2_([before,worker]);
  return {ok:true,requestId:requestId,worker:worker,preparationStatus:'크롤링 배정됨'};

});
}

// 특정 Task 한 건의 담당자만 변경함.
// 완료된 Task는 과거 작업자의 실적이므로 담당자 변경을 막음.
function assignTaskFromKanban(taskId,assignee){return withOpsDocumentLock_(function(){
 const t=resolveWebTask_(taskId),a=t.adapter,worker=String(assignee||'').trim();
 const rt=readOpsTable_(APP_CONFIG.SHEETS.REQUESTS),records=rt.byId[t.id]||[];
 if(records.length!==1||[APP_CONFIG.STATUS.DONE,APP_CONFIG.STATUS.DROP].indexOf(normalizeRequestStatusV2_(opsValue_(records[0].values,rt.headers,'Status')))>=0)throw Error('DONE·DROP 상태이거나 유효하지 않은 요청입니다.');
 if(opsSourceForRequest_(records[0].values,rt.headers)!==t.sheet.getName())throw Error('요청의 작업 원본이 변경됐습니다. 새로고침하세요.');
 if((getWebConfig_().workers||[]).indexOf(worker)<0)throw Error('등록된 작업자를 선택하세요.');
 const patches=[];[[a.a3,a.d3],[a.a2,a.d2]].forEach(function(stage){if(stage[0]>=0&&stage[1]>=0&&t.values[stage[1]]!==true&&String(t.values[stage[0]])!=='N/A')patches.push({row:t.row,col:stage[0]+1,value:worker});});
 if(!patches.length)throw Error('완료된 Task의 담당자는 변경할 수 없습니다.');
 const cols=patches.map(function(p){return p.col;});
 writeOpsPatches_(t.sheet,patches);handleOperationalTaskEdit_({range:t.sheet.getRange(t.row,Math.min.apply(null,cols),1,Math.max.apply(null,cols)-Math.min.apply(null,cols)+1),editedColumns:cols});
 logActivity_(t.id,'TASK_ASSIGNEE_CHANGED','',worker);return {ok:true,requestId:t.id,taskId:taskId,assignee:worker};
});}

// 선택한 요청의 Task 목록을 개인 Kanban 상세 배정창에 반환함.
function getPersonalKanbanTasks(requestId) {
  return getRequestTasksFromWeb(String(requestId || '').trim());
}
// REQUESTS의 Assignee 컬럼에 남아 있는 단일선택 검증만 제거함.
// 복수 담당자 저장을 허용하기 위한 1회성 마이그레이션임.
function repairRequestAssigneeForMultipleWorkers() {
  const sheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  const headers = getHeaderMap_(sheet);
  const column = headers['Active Assignees'];
  if (!column) throw new Error('REQUESTS에서 Assignee 컬럼을 찾을 수 없습니다.');
  sheet.getRange(2, column, Math.max(sheet.getMaxRows() - 1, 1), 1).clearDataValidations();
  SpreadsheetApp.getActiveSpreadsheet().toast('요청 담당자 복수선택 저장 준비가 완료되었습니다.', 'Listing Ops', 4);
}


// 실제 착수/중지 구간을 요청·단계·작업자별로 기록함.
function ensureWorkSessionsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = 'WORK_SESSIONS';
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, 13).setValues([WORK_SESSION_COLUMNS_.slice()]).setBackground('#263650').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function workSessionUser_() {
  return Session.getActiveUser().getEmail() || Session.getTemporaryActiveUserKey();
}

function workSessionRows_() {
  const sheet = ensureWorkSessionsSheet_();
  return sheet.getLastRow() < 2 ? [] : sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.min(sheet.getLastColumn(),13)).getValues();
}

function findActiveWorkerSessionV2_(worker) {
  const rows = workSessionRows_();
  for (let i = rows.length - 1; i >= 0; i -= 1) {
    if (String(rows[i][1] || '') === String(worker) && String(rows[i][4] || '') === 'ACTIVE') {
      return {requestId:String(rows[i][2] || ''),stage:String(rows[i][3] || '')};
    }
  }
  return null;
}

// 미완료 체크 전에 동일 작업자의 실제 작업 세션이 시작되었는지 확인함.
function guardTaskCheckboxStartV2_(e,context) {
  if(!context)return guardTaskCheckboxStartLegacy_(e);
  const c=context,h=c.h,first=e.range.getColumn(),last=e.range.getLastColumn(),candidates=[];
  c.rows.forEach(function(row,i){['3P','2P'].forEach(function(stage){
    const col=h[stage+' Completed'];if(!col||col<first||col>last||row[col-1]!==true)return;
    const worker=String(opsValue_(row,h,stage+' Assignee')||'').trim(),id=String(opsValue_(row,h,'Request ID')||'');
    candidates.push({row:row,index:i,col:col,stage:stage,worker:worker,id:id,valid:isViValidV3_(row,h)});
  });});
  const active=Object.create(null);
  if(candidates.some(function(x){return x.valid&&x.worker&&x.id;})){
    c.sessions=workSessionRows_();c.sessions.forEach(function(row){if(String(row[4])==='ACTIVE')active[String(row[1])]={id:String(row[2]),stage:String(row[3])};});
  }
  if(candidates.some(function(x){return x.valid&&x.id&&x.stage==='2P';})){
    c.requestTable=readSessionRequestRecords_(candidates.filter(function(x){return x.valid&&x.id&&x.stage==='2P';}).map(function(x){return x.id;}),c.requestTable);
    candidates.forEach(function(x){
      if(!x.valid||!x.id||x.stage!=='2P')return;
      const records=c.requestTable.byId[x.id]||[],record=records.length===1?records[0]:null;
      x.needsBarcodeProvided=!record||!normalizeBarcodeProvidedV1_(opsValue_(record.values,c.requestTable.headers,'Barcode Provided'));
    });
  }
  const rejected=candidates.filter(function(x){const s=active[x.worker];return !x.valid||!x.worker||!x.id||x.needsBarcodeProvided||!s||s.id!==x.id||s.stage!==x.stage;});
  if(!c.noPrompt&&rejected.length&&e.range.getNumRows()===1&&e.range.getNumColumns()===1&&rejected[0].valid&&rejected[0].worker&&rejected[0].id){c.needsStart=true;return false;}
  rejected.forEach(function(x){x.row[x.col-1]=false;c.patches.push({row:c.start+x.index,col:x.col,value:false});});
  c.rejected=rejected.length;return true;
}
// Single-cell start confirmation only; called outside the script lock.
function guardTaskCheckboxStartLegacy_(e) {
  if (!e || !e.range) return true;
  const range = e.range, sheet = range.getSheet(), name = sheet.getName(), column = range.getColumn();
  const validityHeaders=getHeaderMap_(sheet);
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) {
    throw Error('다중 편집은 handleTaskRouterV7_ 경로로 처리해야 합니다.');
  }
  let stage = '', assigneeColumn = 0;
  if(column===validityHeaders['3P Completed']){stage='3P';assigneeColumn=validityHeaders['3P Assignee'];}
  if(column===validityHeaders['2P Completed']){stage='2P';assigneeColumn=validityHeaders['2P Assignee'];}
  if (!stage || range.getValue() !== true) return true;
  requireMyWorkTaskHeadersV3_(validityHeaders,['Request ID',stage+' Assignee',stage+' Completed'],name);
  if(validityHeaders['VI Status'] && String(sheet.getRange(range.getRow(),validityHeaders['VI Status']).getValue()||'VALID').toUpperCase()==='INVALID'){
    range.setValue(false);
    SpreadsheetApp.getActive().toast('INVALID VI는 완료 처리할 수 없습니다.','무효 VI',5);
    return false;
  }
  const requestId = String(sheet.getRange(range.getRow(),validityHeaders['Request ID']).getValue() || '').trim();
  const worker = String(sheet.getRange(range.getRow(),assigneeColumn).getValue() || '').trim();
  if (!requestId || !worker) return true;
  const active = findActiveWorkerSessionV2_(worker),ui = SpreadsheetApp.getUi();
  const sameActive=active&&active.requestId===requestId&&active.stage===stage;
  if (!sameActive) {
    let message = worker+'님의 '+requestId+' / '+stage+' 작업을 시작하시겠습니까?';
    if (active) message += '\n\n현재 진행 중인 '+active.requestId+' / '+active.stage+' 작업은 자동으로 일시중지됩니다.';
    const answer = ui.alert('작업 시작 확인',message,ui.ButtonSet.YES_NO);
    if (answer !== ui.Button.YES) { range.setValue(false); return false; }
  }
  const command={worker:worker,requestId:requestId,stage:stage,action:'START'};
  if(stage==='2P'){
    const requestTable=readSessionRequestRecords_([requestId]),record=requestTable.byId[requestId][0];
    if(!normalizeBarcodeProvidedV1_(opsValue_(record.values,requestTable.headers,'Barcode Provided'))){
      const buttons=ui.ButtonSet.YES_NO_CANCEL||ui.ButtonSet.YES_NO;
      const provided=ui.alert('바코드 제공 여부','이 2P 작업은 요청자가 바코드를 제공했나요?\n\n예: 제공됨 · 아니요: 제공되지 않음 · 취소: 완료 체크 취소',buttons);
      if(provided===ui.Button.YES)command.barcodeProvided='YES';
      else if(provided===ui.Button.NO)command.barcodeProvided='NO';
      else {range.setValue(false);return false;}
    }
  }
  try {
    setWorkSessionState(command,e);
    return true;
  } catch (error) {
    range.setValue(false);
    ui.alert('작업을 시작하지 못했습니다.\n'+(error.message || String(error)));
    return false;
  }
}

function getWorkSessionSnapshot_(readOnly,providedRows) {
  let rows;
  if(providedRows)rows=providedRows;
  else if(readOnly===true){
    const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('WORK_SESSIONS');
    const last=sheet?sheet.getLastRow():0;
    rows=last<2?[]:sheet.getRange(2,1,last-1,10).getValues();
  }else rows=workSessionRows_();
  const now = new Date();
  const map = {};
  rows.forEach(function(row) {
    const worker = String(row[1] || '');
    const requestId = String(row[2] || '');
    const stage = String(row[3] || '');
    if (!worker || !requestId || !stage) return;
    const key = worker + '\n' + requestId + '\n' + stage;
    if (!map[key]) map[key] = {
      worker: worker, requestId: requestId, stage: stage,
      status: 'PAUSED', startedAt: '', elapsedMinutes: 0
    };
    const start = row[5] instanceof Date ? row[5] : null;
    const end = row[6] instanceof Date ? row[6] : null;
    let minutes = Number(row[7] || 0);
    if (String(row[4]) === 'ACTIVE' && start) {
      minutes = Math.max(0, (now - start) / 60000);
      map[key].status = 'ACTIVE';
      map[key].sessionId = String(row[0]);
      map[key].startedAt = Utilities.formatDate(start, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
    }
    if (end || String(row[4]) !== 'ACTIVE') map[key].elapsedMinutes += Math.max(0, minutes);
    else map[key].elapsedMinutes += minutes;
  });
  const sessions = Object.keys(map).map(function(key) {
    map[key].elapsedMinutes = Math.round(map[key].elapsedMinutes * 10) / 10;
    return map[key];
  });
  const active = sessions.filter(function(item) { return item.status === 'ACTIVE'; });
  return {sessions: sessions, active: active};
}

function pauseSessionRow_(sheet, rowNumber, now, note, record,viContext) {
  return withOpsDocumentLock_(function(){
    const context=viContext||createWorkSessionViContext_(),r=record||sheet.getRange(rowNumber,1,1,Math.min(sheet.getLastColumn(),13)).getValues()[0];
    if(r[4]!=='ACTIVE')return false;
    assertWorkSessionViSchema_(sheet,context);
    const next=finalizeWorkSessionViMetrics_(r,now,note,context);
    writeFinalizedWorkSessions_(sheet,[{row:rowNumber,values:next.slice(4,13),record:r,next:next}]);return true;
  });
}

function validateWorkStage_(requestType, stage) {
  if (stage==='CRAWL') return;
  if (['3P', '2P'].indexOf(stage) < 0) throw new Error('작업 단계를 확인해 주세요.');
  if (requestType === '3P Listing' && stage !== '3P') throw new Error('이 요청은 3P 단계만 진행합니다.');
  if (requestType === '2P Lite 전환' && stage !== '2P') throw new Error('이 요청은 2P 단계만 진행합니다.');
  if (requestType === '3P Listing 후 2P Lite 전환' && ['3P','2P'].indexOf(stage) < 0) {
    throw new Error('이 요청의 작업 단계를 확인해 주세요.');
  }
}

// 작업자별 활성 세션은 하나만 허용하며 시작 시 기존 세션을 자동 중지함.
function setWorkSessionStateLegacy_(payload,nativeEdit,commitContext) {
  payload = payload || {};
  const worker = String(payload.worker || '').trim();
  const requestId = String(payload.requestId || '').trim();
  const stage = String(payload.stage || '').trim();
  const action = String(payload.action || '').toUpperCase();
  if (!worker || !requestId) throw new Error('작업자와 Request ID를 확인해 주세요.');
  if (['START', 'PAUSE', 'COMPLETE'].indexOf(action) < 0) throw new Error('지원하지 않는 작업 상태입니다.');

  return withOpsDocumentLock_(function(){
    const sidebarHadChanges=!commitContext&&!!payload.sidebarResult&&payload.expectedRevision!==getMyWorkSidebarRevision();
    const requestTable=readSessionRequestRecords_([requestId]),found=requestTable.byId[requestId]||[];
    if(found.length!==1)throw Error('요청 ID가 없거나 중복되었습니다.');
    const requestSheet=requestTable.sheet,headers=requestTable.headers,requestRow=found[0].row,request=found[0].values;
    const val=function(name){return opsValue_(request,headers,name);},requestType=String(val('Request Type')||'3P Listing');
    if(action==='START'&&[APP_CONFIG.STATUS.DONE,APP_CONFIG.STATUS.DROP,APP_CONFIG.STATUS.ON_HOLD].indexOf(normalizeRequestStatusV2_(val('Status')))>=0)throw Error('DONE·DROP·ON HOLD 요청은 시작할 수 없습니다.');
    const participants = stage==='CRAWL'
      ? [String(val('크롤링 작업자')||'').trim()].filter(Boolean)
      : String(val('Active Assignees')||'').split(',').map(function(value) { return value.trim(); }).filter(Boolean);
    if (action==='START'&&participants.indexOf(worker) < 0) throw new Error(worker + ' 작업자는 이 요청에 배정되어 있지 않습니다.');
    validateWorkStage_(requestType, stage);
    if(action==='START'&&stage==='CRAWL' && (!headers['크롤링 단계'] || String(val('크롤링 단계'))==='CRAWL_COMPLETED')) throw new Error('크롤링 작업을 시작할 수 없는 요청입니다.');
    const sheet = commitContext?commitContext.sheet:ensureWorkSessionsSheet_();
    const rows = commitContext?commitContext.rows:workSessionRows_();
    const now = new Date();
    const matchingActiveIndex=action==='START'?rows.findIndex(function(r){return String(r[1])===worker&&String(r[2])===requestId&&String(r[3])===stage&&r[4]==='ACTIVE';}):-1;
    // Duplicate START never changes the baseline of an existing active session.
    if(matchingActiveIndex>=0){
      if(stage==='2P')captureBarcodeProvidedV1_(requestSheet,found[0],headers,payload.barcodeProvided);
      if(commitContext){commitContext.reused=true;sessionPrepareCommit_(commitContext,String(rows[matchingActiveIndex][0]),[requestId]);}
      return {ok:true,action:action,snapshot:getWorkSessionSnapshot_(true,rows)};
    }
    const targets=rows.map(function(row,index){return {row:row,index:index};}).filter(function(x){return String(x.row[1])===worker&&x.row[4]==='ACTIVE'&&(action==='START'||String(x.row[2])===requestId&&String(x.row[3])===stage&&(!payload.sessionId||String(x.row[0])===String(payload.sessionId)));});
    if((action==='PAUSE'||action==='COMPLETE')&&!targets.length){
      if(action==='PAUSE'&&rows.some(function(r){return String(r[1])===worker&&String(r[2])===requestId&&String(r[3])===stage&&r[4]==='PAUSED';}))return {ok:true,action:action,alreadyPaused:true,snapshot:getWorkSessionSnapshot_(true,rows)};
      throw Error('진행 중인 작업 세션을 찾지 못했습니다.');
    }
    const viContext=createWorkSessionViContext_(requestTable);
    if(commitContext){commitContext.viContext=viContext;viContext.sidebarTaskLookup=true;}
    readSessionRequestRecords_(targets.map(function(x){return String(x.row[2]);}),requestTable);
    assertWorkSessionViSchema_(sheet,viContext);
    const startBaseline=action==='START'?previousWorkSessionEndViV2_(rows,worker,requestId,stage,commitContext):'';
    const viIds=targets.filter(function(x){return ['3P','2P'].indexOf(String(x.row[3]))>=0;}).map(function(x){return String(x.row[2]);});
    // START baseline comes from prior End Completed VI. Only combined 2P has a
    // task-count eligibility gate; navigation is no longer part of this lock.
    if(action==='START'&&(!commitContext&&['3P','2P'].indexOf(stage)>=0||stage==='2P'&&requestType==='3P Listing 후 2P Lite 전환'))viIds.push(requestId);
    viContext.sources=scopedSourcesForRequests_(requestTable,viIds,viContext);
    Object.keys(viContext.sources).forEach(function(name){viContext.index[name]=viContext.sources[name].index;});
    let completesRequest=false;
    if(action==='COMPLETE'){
      if(['3P','2P'].indexOf(stage)<0)throw Error('이 단계의 완료는 전용 완료 버튼을 사용해 주세요.');
      const source=workSessionViSource_(requestId,viContext),metrics=viContext.index[source]&&viContext.index[source][requestId];
      const mine=metrics&&metrics.workers&&metrics.workers[worker],assigned=Number(mine&&mine['assigned'+stage])||0,done=Number(mine&&mine['done'+stage])||0;
      if(!(assigned>0&&done>=assigned))throw Error('내게 배정된 VALID VI를 모두 체크한 뒤 완료할 수 있습니다.');
      const total=Number(metrics.total)||0;
      completesRequest=total>0&&(requestType==='3P Listing'?Number(metrics.done3)>=total:requestType==='2P Lite 전환'?Number(metrics.done2)>=total:Number(metrics.both)>=total);
    }
    if(stage==='2P'&&(action==='START'||action==='COMPLETE')){
      if(requestType==='3P Listing 후 2P Lite 전환'){
        const gateSource=workSessionViSource_(requestId,viContext),gateMetrics=viContext.index[gateSource]&&viContext.index[gateSource][requestId],gateMine=gateMetrics&&gateMetrics.workers&&gateMetrics.workers[worker];
        if(!(Number(gateMine&&gateMine.assigned3P)>0&&Number(gateMine&&gateMine.done3P)>=Number(gateMine&&gateMine.assigned3P)))throw Error('내게 배정된 VALID 3P VI를 모두 완료한 뒤 2P Lite를 진행할 수 있습니다.');
      }
      captureBarcodeProvidedV1_(requestSheet,found[0],headers,payload.barcodeProvided);
    }
    // Calculate all old-request ends and the new baseline before any session writes.
    if(!commitContext)markRequestReconcileNeeded_([requestId].concat(targets.map(x=>String(x.row[2]))),{tasks:true,sessions:true},'session-complete');
    const changes=targets.map(function(x){const note=action==='START'?'다른 작업 시작으로 자동 일시중지':action==='COMPLETE'?'사용자 완료':'사용자 일시중지';const next=finalizeWorkSessionViMetrics_(x.row,now,note,viContext);return {row:x.index+2,values:next.slice(4,13),record:x.row,next:next};});
    if (action === 'START') {
      const baseline=startBaseline;
      const sessionId = 'WS-' + (commitContext?payload.operationId:Utilities.getUuid());
      const next=[sessionId,worker,requestId,stage,'ACTIVE',now,'','',workSessionUser_(),'',baseline,'',''];
      if(commitContext)sessionPrepareCommit_(commitContext,sessionId,[requestId].concat(targets.map(function(x){return String(x.row[2]);})));
      commitSessionRows_(sheet,changes,next,commitContext);rows.push(next);
      if(commitContext)return {ok:true};
      if (headers['Started At'] && !val('Started At')) {
        requestSheet.getRange(requestRow, headers['Started At']).setValue(now).setNumberFormat('yyyy-mm-dd hh:mm');
      }
      const currentStatus = normalizeRequestStatusV2_(val('Status'));
      if (currentStatus !== APP_CONFIG.STATUS.DONE && currentStatus !== APP_CONFIG.STATUS.ON_HOLD && currentStatus !== APP_CONFIG.STATUS.ONGOING) {
        requestSheet.getRange(requestRow, headers['Status']).setValue(APP_CONFIG.STATUS.ONGOING);
        request[headers['Status']-1]=APP_CONFIG.STATUS.ONGOING;
      }
      activateWorkRows_(requestId, stage, requestType,viContext);
    } else {
      if(commitContext)sessionPrepareCommit_(commitContext,String(payload.sessionId),[requestId]);
      writeFinalizedWorkSessions_(sheet,changes,commitContext);
      if(commitContext)return {ok:true};
    }
    const revisionWorkers=new Set([worker]);
    if(action==='COMPLETE'&&completesRequest){
      rows.forEach(function(row){if(String(row[2])===requestId&&row[4]==='ACTIVE')revisionWorkers.add(String(row[1]));});
      const hours=closeOpsRequestSessions_([requestId],now,rows,viContext)[requestId],patches=[],beforeStatus=normalizeRequestStatusV2_(val('Status'))||APP_CONFIG.STATUS.TO_DO;
      addOpsPatch_(patches,found[0],headers,'Status',APP_CONFIG.STATUS.DONE);
      addOpsPatch_(patches,found[0],headers,'Completed At',now);
      const requestedAt=val('Requested At');if(requestedAt instanceof Date)addOpsPatch_(patches,found[0],headers,'Lead Time (Hours)',Math.max(0,(now-requestedAt)/3600000));
      writeOpsPatches_(requestSheet,patches);
      if(beforeStatus!==APP_CONFIG.STATUS.DONE)try{logActivity_(requestId,'REQUEST_COMPLETED',beforeStatus,APP_CONFIG.STATUS.DONE);}catch(error){console.error('Completion log: '+String(error));}
    }
    const affected=new Set([requestId]);targets.forEach(function(x){affected.add(String(x.row[2]));});
    let summaryPending=false;
    try{
      reconcileRequestSummaries_(Array.from(affected),{requestTable:requestTable,sessionRows:rows,taskSources:viContext.sources,now:now,notify:false,reason:'session-complete'});
      bumpMyWorkSidebarRevisionV2_(Array.from(revisionWorkers));
    }catch(error){
      summaryPending=true;
      const pending=Array.from(affected),tickets=markRequestReconcileNeeded_(pending,{tasks:true,sessions:true},'complete-post');
      pending.forEach(id=>settleRequestReconcile_(id,tickets[id],{tasks:true,sessions:true},null,error));
      console.error('Completion post-processing pending: '+String(error));
    }
    if(nativeEdit&&nativeEdit.range)nativeEdit.sessionTaskContext=viContext;
    const result={ok:true,action:action,summaryPending:summaryPending,needsRefresh:sidebarHadChanges||summaryPending,snapshot:getWorkSessionSnapshot_(true,payload.sidebarResult?rows.filter(function(r){return String(r[1])===worker;}):rows)};
    try{return sidebarSessionResult_(result,payload,requestTable,viContext);}catch(error){result.needsRefresh=true;return result;}
  });
}

// 시작한 요청의 실제 작업 시트 첫 행으로 이동함. 공유 필터는 건드리지 않음.
function activateWorkRows_(requestId, stage, requestType,viContext) {
  if (stage === 'CRAWL') return;
  if(viContext&&viContext.requestTable){
    const rt=viContext.requestTable,record=rt.byId[String(requestId)][0],name=opsSourceForRequest_(record.values,rt.headers),table=viContext.sources[name];
    if(table&&table.scoped&&table.loadedRequestIds.has(String(requestId))){
      const records=table.byId[String(requestId)]||[];table.sheet.activate();
      if(records.length)table.sheet.getRange(records[0].row,1).activate();return;
    }
  }
  const use2PSheet = stage === '2P' && requestType === '2P Lite 전환';
  const rs=getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS),rh=getHeaderMap_(rs),rr=findRowByRequestId_(rs,requestId);
  const isCrawl=rh['크롤링 필요']&&rs.getRange(rr,rh['크롤링 필요']).getValue()===true;
  const sheet = getRequiredSheet_(use2PSheet ? APP_CONFIG.SHEETS.TASKS_2P : (isCrawl?APP_CONFIG.SHEETS.CRAWL_TASKS:APP_CONFIG.SHEETS.TASKS));
  if (sheet.getLastRow() < 2) return;
  const taskHeaders=getHeaderMap_(sheet);
  requireMyWorkTaskHeadersV3_(taskHeaders,['Request ID'],sheet.getName());
  const idColumn=taskHeaders['Request ID'];
  const ids = sheet.getRange(2, idColumn, sheet.getLastRow() - 1, 1).getDisplayValues();
  const index = ids.findIndex(function(row) { return String(row[0]) === requestId; });
  sheet.activate();
  if (index >= 0) sheet.getRange(index + 2, idColumn).activate();
}

function pauseActiveStageSessions_(requestId, stage, now, note,viContext) {
 return withOpsDocumentLock_(function(){
  const sheet = ensureWorkSessionsSheet_();
  const rows = workSessionRows_();
  const context=viContext||createWorkSessionViContext_(),changes=[];now=now||new Date();
  rows.forEach(function(row, index) {
    if (String(row[2]) === String(requestId) && String(row[3]) === String(stage) && String(row[4]) === 'ACTIVE') {
      assertWorkSessionViSchema_(sheet,context);
      const next=finalizeWorkSessionViMetrics_(row,now,note||'단계 완료',context);changes.push({row:index+2,values:next.slice(4,13),record:row,next:next});
    }
  });
  writeFinalizedWorkSessions_(sheet,changes);return rows;
 });
}

function updateRequestWorkTimeFromSessions_(requestId, now, providedRows,requestTable) {
  const result=reconcileRequestSummary_(String(requestId),{domains:{sessions:true},now:now||new Date(),sessionRows:providedRows,requestTable:requestTable,notify:false,reason:'work-time'});
  const report=result.reports[String(requestId)];return report.workTimeBlocked||report.archived?null:report.minutes;
}

// 요청 완료 시 남아 있는 모든 활성 구간을 완료 시각으로 닫음.
function closeCompletedWorkSessions_(requestId, now) {
 return withOpsDocumentLock_(function(){
  const rows=workSessionRows_(),context=createWorkSessionViContext_();now=now||new Date();
  closeOpsRequestSessions_([String(requestId)],now,rows,context);
  updateRequestWorkTimeFromSessions_(requestId,now,rows,context.requestTable);
 });
}


// 작업자용 개인 작업함을 우측 고정 사이드바로 열기.
function showMyWorkSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('MyWorkSidebar')
    .setTitle('내 작업함');
  SpreadsheetApp.getUi().showSidebar(html);
}

// 마지막으로 선택한 작업자는 Google 계정별 사용자 속성에 저장됨.
function saveMyWorkSidebarWorker(worker,progressive) {
  worker = String(worker || '').trim();
  // Progressive bootstrap validates and persists the worker itself; avoid a duplicate CONFIG read.
  if(progressive)return getMyWorkSidebarInitial(worker);
  const workers = getWebConfig_().workers || [];
  if (worker && workers.indexOf(worker) < 0) throw new Error('등록되지 않은 작업자입니다.');
  PropertiesService.getUserProperties().setProperty('LISTING_OPS_MY_WORKER', worker);
  return getMyWorkSidebarData(true);
}

function getMyWorkSidebarRevision() {
  const worker = PropertiesService.getUserProperties().getProperty('LISTING_OPS_MY_WORKER') || '';
  const props = PropertiesService.getDocumentProperties();
  return (props.getProperty('MY_WORK_SIDEBAR_REVISION') || '0') + ':' +
    (props.getProperty('MY_WORK_WORKER_REV_' + encodeURIComponent(worker)) || '0');
}

function bumpMyWorkSidebarRevisionV2_(workers) {
  const revision = String(new Date().getTime()) + '-' + Utilities.getUuid().slice(0, 8);
  const props = PropertiesService.getDocumentProperties();
  const names = Array.isArray(workers) ? workers.filter(Boolean) : [];
  if (!names.length) props.setProperty('MY_WORK_SIDEBAR_REVISION', revision);
  else {
    const updates = {};
    names.forEach(function(worker) {
      updates['MY_WORK_WORKER_REV_' + encodeURIComponent(String(worker).trim())] = revision;
    });
    props.setProperties(updates, false);
  }
  return revision;
}

// 이미 생성된 요청별 필터 보기 URL을 3P/2P 단계별로 읽음.
function getMyWorkFilterViewUrlsV2_(requestTypes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const task3P = getRequiredSheet_(APP_CONFIG.SHEETS.TASKS);
  const task2P = getRequiredSheet_(APP_CONFIG.SHEETS.TASKS_2P);
  const sheetsById = {};
  sheetsById[task3P.getSheetId()] = {sheet:task3P, ids:{}};
  sheetsById[task2P.getSheetId()] = {sheet:task2P, ids:{}};
  try {
    const metadata = Sheets.Spreadsheets.get(ss.getId(), {
      fields:'sheets(properties(sheetId),filterViews(filterViewId,title))'
    });
    (metadata.sheets || []).forEach(function(item) {
      const target = item.properties && sheetsById[item.properties.sheetId];
      if (!target) return;
      (item.filterViews || []).forEach(function(view) {
        const title = String(view.title || '');
        if (title.indexOf('[AUTO] ') === 0) target.ids[title.substring(7)] = view.filterViewId;
      });
    });
  } catch (ignored) {}
  const urls = {}, base = ss.getUrl();
  Object.keys(requestTypes || {}).forEach(function(requestId) {
    const type = requestTypes[requestId];
    const makeUrl_ = function(target) {
      const fvid = target.ids[requestId];
      const gid = target.sheet.getSheetId();
      return base + '#gid=' + gid + (fvid ? '&fvid=' + fvid : '');
    };
    urls[requestId] = {
      threeP: type === '2P Lite 전환' ? '' : makeUrl_(sheetsById[task3P.getSheetId()]),
      twoP: type === '3P Listing' ? '' : makeUrl_(sheetsById[task2P.getSheetId()])
    };
  });
  return urls;
}

function getMyWorkSidebarData(forceRefresh) {
  const workers = getWebConfig_().workers || [];
  let worker = PropertiesService.getUserProperties().getProperty('LISTING_OPS_MY_WORKER') || '';
  if (workers.indexOf(worker) < 0) worker = '';
  const revision = getMyWorkSidebarRevision();
  const cache = CacheService.getDocumentCache();
  const cacheKey = 'MY_WORK_V8_' + (worker || 'NONE') + '_' + revision;
  const cached = forceRefresh ? null : cache.get(cacheKey);
  if (cached) return JSON.parse(cached);
  const requestSheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  const headers = getHeaderMap_(requestSheet);
  const rows = getDataRows_(requestSheet);
  function value_(row, name) { return headers[name] ? row[headers[name] - 1] : ''; }
  const originalFileUrls = {};
  const crawlResultFileUrls = {};
  if (headers['원본 요청 파일'] && requestSheet.getLastRow() >= 2) {
    const count = requestSheet.getLastRow() - 1;
    const ids = requestSheet.getRange(2, headers['Request ID'], count, 1).getDisplayValues();
    const fileRange = requestSheet.getRange(2, headers['원본 요청 파일'], count, 1);
    const richValues = fileRange.getRichTextValues();
    const displayValues = fileRange.getDisplayValues();
    const notes = fileRange.getNotes();
    richValues.forEach(function(row, index) {
      const rich = row[0], requestId = String(ids[index][0] || '');
      if (!rich || !requestId) return;
      let url = rich.getLinkUrl() || '';
      if (!url) {
        const linkedRun = rich.getRuns().filter(function(run){return run.getLinkUrl();})[0];
        if (linkedRun) url = linkedRun.getLinkUrl();
      }
      if (!url && /^https?:\/\//i.test(String(notes[index][0] || ''))) url = String(notes[index][0]);
      const display = String(displayValues[index][0] || '');
      if (!url && /^https?:\/\//i.test(display)) url = display;
      if (!url && display) {
        try {
          const files = getOriginalRequestFolderV2_().getFilesByName(display);
          if (files.hasNext()) url = files.next().getUrl();
        } catch (ignored) {}
      }
      if (url) originalFileUrls[requestId] = url;
    });
  }
  if (headers['크롤링 결과 파일']) rows.forEach(function(row){
    const id=String(value_(row,'Request ID')||''), url=String(value_(row,'크롤링 결과 파일')||'');
    if(id&&url) crawlResultFileUrls[id]=url;
  });
  const assignments = {};
  function bucket_(requestId) {
    if (!assignments[requestId]) assignments[requestId] = {assigned3P:0,done3P:0,assigned2P:0,done2P:0};
    return assignments[requestId];
  }
  if(worker){
    const sourceNames=rows.map(function(row){return opsSourceForRequest_(row,headers);});
    const index=buildTaskIndex_(readTaskSources_(sourceNames));
    rows.forEach(function(row){const id=String(value_(row,'Request ID')||''),m=(index[opsSourceForRequest_(row,headers)]||{})[id];
      if(m&&m.workers[worker])assignments[id]=m.workers[worker];
    });
  }
  const requestTypesForViews = {};
  rows.forEach(function(row) {
    const requestId = String(value_(row, 'Request ID') || '');
    if (requestId) requestTypesForViews[requestId] = String(value_(row, 'Request Type') || '3P Listing');
  });
  const taskViewUrls = worker ? getMyWorkFilterViewUrlsV2_(requestTypesForViews) : {};
  const requests = worker ? rows.filter(function(row) {
    const id = String(value_(row, 'Request ID') || ''), b = assignments[id];
    const status=normalizeRequestStatusV2_(value_(row,'Status'));
    const crawlMine=value_(row,'크롤링 필요')===true&&String(value_(row,'크롤링 작업자')||'')===worker&&String(value_(row,'크롤링 단계')||'')!=='CRAWL_COMPLETED';
    return (crawlMine || (b && (b.assigned3P > 0 || b.assigned2P > 0))) && status !== APP_CONFIG.STATUS.DONE && status !== APP_CONFIG.STATUS.DROP;
  }).map(function(row) {
    const requestId = String(value_(row, 'Request ID') || ''), personal = assignments[requestId] || {};
    const total = Number(value_(row, 'Total VI')) || 0;
    const done3 = Number(value_(row, '3P Completed VI')) || 0;
    const done2 = Number(value_(row, '2P Completed VI')) || 0;
    return {
      requestId:requestId, requestor:String(value_(row, 'Requestor') || ''),
      brand:String(value_(row, 'Brand') || ''), node:String(value_(row, 'Node') || ''),
      aCode:String(value_(row, 'A코드') || ''), requestType:String(value_(row, 'Request Type') || '3P Listing'),
      source:String(value_(row, 'Source') || ''), notes:String(value_(row, 'Requestor Notes') || ''),
      priority:String(value_(row, 'Priority') || 'NORMAL'), status:normalizeRequestStatusV2_(value_(row, 'Status')) || APP_CONFIG.STATUS.TO_DO,
      assignee:String(value_(row, 'Active Assignees') || ''), totalVi:total, completed3P:done3, completed2P:done2,
      assigned3P:Number(personal.assigned3P)||0, done3P:Number(personal.done3P)||0,
      assigned2P:Number(personal.assigned2P)||0, done2P:Number(personal.done2P)||0,
      barcodeProvided:normalizeBarcodeProvidedV1_(value_(row, 'Barcode Provided')),
      taskViewUrls:taskViewUrls[requestId] || {threeP:'', twoP:''},
      taskViewUrl:String((taskViewUrls[requestId] || {}).threeP || (taskViewUrls[requestId] || {}).twoP || ''),
      originalFileUrl:String(originalFileUrls[requestId] || value_(row, '원본 요청 파일') || ''),
      crawlResultFileUrl:String(crawlResultFileUrls[requestId] || ''),
      crawlRequired:value_(row,'크롤링 필요')===true,
      crawlStage:String(value_(row,'크롤링 단계')||''),
      crawlAssignee:String(value_(row,'크롤링 작업자')||''),
      preparationStatus:String(value_(row,'자료 준비 상태')||'')
    };
  }) : [];
  const snapshot = worker ? getWorkSessionSnapshot_(true) : {sessions:[],active:[]};
  const result = {
    revision: revision,
    worker: worker,
    workers: workers,
    requests: requests,
    workSessions: {
      sessions:snapshot.sessions.filter(function(item){return item.worker === worker;}),
      active:snapshot.active.filter(function(item){return item.worker === worker;})
    },
    updatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd HH:mm')
  };
  const serialized = JSON.stringify(result);
  try {
    if (Utilities.newBlob(serialized).getBytes().length < 95000) cache.put(cacheKey, serialized, 120);
  } catch (cacheError) {
    // 캐시 한도를 넘더라도 작업함 자체는 정상 반환함.
  }
  return result;
}

function getMyWorkRequest_(requestId) {
  const data = getMyWorkSidebarData();
  return data.requests.filter(function(item){return item.requestId === requestId;})[0] || null;
}

// 타이머 상태를 바꾸지 않고 관련 작업 행으로만 이동함.
function goToMyWorkRows(requestId, stage, worker) {
  requestId = String(requestId || '').trim();
  stage = String(stage || '').trim();
  worker = String(worker || PropertiesService.getUserProperties().getProperty('LISTING_OPS_MY_WORKER') || '').trim();
  const requestSheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  const headers = getHeaderMap_(requestSheet);
  const row = findRowByRequestId_(requestSheet, requestId);
  const requestType = String(requestSheet.getRange(row, headers['Request Type']).getValue() || '3P Listing');
  validateWorkStage_(requestType, stage);
  const use2PSheet = stage === '2P' && requestType === '2P Lite 전환';
  const sheet = getRequiredSheet_(use2PSheet ? APP_CONFIG.SHEETS.TASKS_2P : APP_CONFIG.SHEETS.TASKS);
  if (sheet.getLastRow() < 2) throw new Error('이동할 작업 행이 없습니다.');
  const h = getHeaderMap_(sheet), values = getDataRows_(sheet);
  const assigneeHeader=stage+' Assignee',completedHeader=stage+' Completed';
  requireMyWorkTaskHeadersV3_(h,['Request ID',assigneeHeader,completedHeader],sheet.getName());
  const idIndex = h['Request ID'] - 1;
  const assigneeIndex=h[assigneeHeader]-1,completedIndex=h[completedHeader]-1;
  let index = values.findIndex(function(r){return String(r[idIndex])===requestId&&String(r[assigneeIndex]||'').trim()===worker&&r[completedIndex]!==true;});
  if (index < 0) index = values.findIndex(function(r){return String(r[idIndex])===requestId&&String(r[assigneeIndex]||'').trim()===worker;});
  if (index < 0) index = values.findIndex(function(r){return String(r[idIndex])===requestId;});
  if (index < 0) throw new Error('이동할 작업 행을 찾지 못했습니다.');
  sheet.activate(); sheet.getRange(index + 2, assigneeIndex + 1).activate();
  return {ok:true,row:index+2};
}
