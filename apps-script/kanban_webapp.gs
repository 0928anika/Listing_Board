// Listing Operations 웹앱 화면을 반환함.
// Google Sheets는 데이터 저장소로만 사용하고 사용자는 이 HTML 화면에서 작업함.
function doGet() {
  return showReadOnlyDashboardWeb_();
}

// 시트 메뉴에서 조회 전용 Kanban 팝업을 열기.
// 작업자 배정 콘솔과는 별도 화면이며 데이터 변경 기능을 제공하지 않음.
function showKanbanDialog() {
  var html = HtmlService.createHtmlOutputFromFile('KanbanDialog')
    .setWidth(1400)
    .setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, 'Kanban');
}

// 조회 전용 Kanban에 필요한 요청만 반환함.
// DROP 요청은 제외하고 DONE 요청은 KANBAN 시트와 동일하게 최근 14일만 표시함.
function getKanbanDialogData() {
  var data = getOperationsWebData();
  var timezone = Session.getScriptTimeZone() || 'Asia/Seoul';
  var doneRetentionDays = 14;
  var doneCutoff = new Date();
  doneCutoff.setDate(doneCutoff.getDate() - doneRetentionDays);
  var cutoffKey = Utilities.formatDate(doneCutoff, timezone, 'yyyy-MM-dd');
  var oldDoneCount = 0;

  var requests = data.requests.filter(function(request) {
    if (request.status === 'DROP') return false;
    if (request.status !== 'DONE') return true;
    var completedKey = String(request.completedAt || '').slice(0, 10);
    if (!completedKey || completedKey >= cutoffKey) return true;
    oldDoneCount += 1;
    return false;
  });

  return {
    generatedAt: data.generatedAt,
    doneRetentionDays: doneRetentionDays,
    oldDoneCount: oldDoneCount,
    requests: requests,
    config: data.config
  };
}

// 웹앱 첫 로딩에 필요한 요청, Task, CONFIG 선택 목록을 한 번에 반환함.
function getOperationsWebData() {
  var requestSheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  var requestHeaders = getHeaderMap_(requestSheet);
  var requestRows = getDataRows_(requestSheet);
  var timezone = Session.getScriptTimeZone() || 'Asia/Seoul';

  var requests = requestRows.filter(function(row) {
    return String(row[requestHeaders['Request ID'] - 1] || '').trim();
  }).map(function(row) {
    var total = Number(row[requestHeaders['Total VI'] - 1]) || 0;
    var done = Math.round(total * (Number(row[requestHeaders['Progress'] - 1]) || 0));
    return {
      requestId: String(row[requestHeaders['Request ID'] - 1] || ''),
      requestedAt: formatWebDate_(row[requestHeaders['Requested At'] - 1], timezone),
      requester: String(row[requestHeaders['Requestor'] - 1] || ''),
      brand: String(row[requestHeaders['Brand'] - 1] || ''),
      node: String(row[requestHeaders['Node'] - 1] || ''),
      priority: String(row[requestHeaders['Priority'] - 1] || ''),
      requestType: String(row[requestHeaders['Request Type'] - 1] || ''),
      assignee: String(row[requestHeaders['Active Assignees'] - 1] || ''),
      status: normalizeWebStatus_(row[requestHeaders['Status'] - 1]),
      startedAt: formatWebDate_(row[requestHeaders['Started At'] - 1], timezone),
      completedAt: formatWebDate_(row[requestHeaders['Completed At'] - 1], timezone),
      totalVi: total,
      completedVi: done,
      completed3P: Number(row[requestHeaders['3P Completed VI'] - 1]) || 0,
      completed2P: Number(row[requestHeaders['2P Completed VI'] - 1]) || 0,
      barcodeProvided: normalizeBarcodeProvidedV1_(row[(requestHeaders['Barcode Provided'] || 0) - 1]),
      progress: total ? Math.round(done / total * 100) : 0,
      notes: String(row[requestHeaders['Notes'] - 1] || '')
    };
  }).sort(function(a,b){ return b.requestedAt.localeCompare(a.requestedAt); });

  var config = getWebConfig_();
  config.nodes = requests.map(function(item) { return item.node; }).filter(function(value, index, array) { return value && array.indexOf(value) === index; }).sort();
  return {
    generatedAt: Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd HH:mm:ss'),
    requests: requests, tasks: [], config: config
  };
}


// 상세 화면을 열 때 선택한 Request ID의 Task만 지연 조회함.
function getRequestTasksFromWeb(requestId){return readRequestWebTasks_(String(requestId));}

// 웹앱에서 요청과 여러 상품 Task를 동시에 등록함.
// 기존 검증·Document Lock 로직을 그대로 재사용함.
function createListingRequestFromWeb(payload) {
  if (!payload || !String(payload.node || '').trim()) throw new Error('Node를 입력해 주세요.');
  if (!payload || ['3P Listing','2P Lite 전환','3P Listing 후 2P Lite 전환'].indexOf(String(payload.requestType || '')) < 0) {
    throw new Error('요청 유형을 선택해 주세요.');
  }
  var result = submitCombinedListingRequest(payload);
  if (!result || result.ok !== true) return result;
  var sheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  var headers = getHeaderMap_(sheet);
  var row = findRowByRequestId_(sheet, result.requestId);
  sheet.getRange(row, headers['Request Type']).setValue(payload.requestType);
  sheet.getRange(row, headers['Node']).setValue(String(payload.node).trim());
  propagateNodeToTasksFromWeb_(result.requestId, String(payload.node).trim());
  if (String(payload.requestType) === '2P Lite 전환') {
    sync2PRequestV7_(result.requestId);
  } else {
    syncRequestLayoutV6_(result.requestId);
  }
  logActivity_(result.requestId, 'REQUEST_TYPE_SET', '', payload.requestType);
  return result;
}


// Request의 Node를 연결된 모든 Task에 복사함.
function propagateNodeToTasksFromWeb_(requestId, node) {
  // Node는 REQUESTS에서만 관리하며 TASKS에는 중복 저장하지 않음.
}

// 웹앱에서 요청의 담당자와 상태를 변경함.
// 담당자 변경은 연결된 모든 Task에도 한 번만 전파함.
function updateRequestFromWeb(payload) {
  if (!payload || !payload.requestId) throw new Error('Request ID가 없습니다.');
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('다른 작업을 처리 중입니다. 잠시 후 다시 시도해 주세요.');
  try {
    var requestId = String(payload.requestId).trim();
    var sheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
    var headers = getHeaderMap_(sheet);
    var row = findRowByRequestId_(sheet, requestId);
    var nextAssignee = String(payload.assignee || '').trim();
    var nextStatus = String(payload.status || '').trim();
    var nextRequestType = String(payload.requestType || '').trim();
    var beforeRequestType = String(sheet.getRange(row, headers['Request Type']).getValue() || '');
    var beforeAssignee = String(sheet.getRange(row, headers['Active Assignees']).getValue() || '');
    var beforeStatus = normalizeRequestStatusV2_(sheet.getRange(row, headers['Status']).getValue());

    if (nextRequestType && ['3P Listing','2P Lite 전환','3P Listing 후 2P Lite 전환'].indexOf(nextRequestType) < 0) throw new Error('허용되지 않은 요청 유형입니다.');
    if (nextStatus && normalizeRequestStatusV2_(nextStatus)!==beforeStatus) assertManualRequestStatus_(nextStatus);
    if (nextRequestType && beforeRequestType !== nextRequestType) {
      const sources=readTaskSources_([APP_CONFIG.SHEETS.TASKS,APP_CONFIG.SHEETS.CRAWL_TASKS,APP_CONFIG.SHEETS.TASKS_2P]);
      if(Object.keys(sources).some(function(name){return (sources[name].byId[requestId]||[]).length>0;}))throw new Error('Task가 있는 요청은 유형만 변경할 수 없습니다. 원본 작업의 단계와 시트 이전을 먼저 검토하세요.');
    }
    if (nextRequestType && beforeRequestType !== nextRequestType) {
      sheet.getRange(row, headers['Request Type']).setValue(nextRequestType);
      logActivity_(requestId, 'REQUEST_TYPE_CHANGED_FROM_WEB', beforeRequestType, nextRequestType);
      recalculateTaskCompletionFromWeb_(requestId, nextRequestType);
    }
    if (beforeAssignee !== nextAssignee) {
      sheet.getRange(row, headers['Active Assignees']).setValue(nextAssignee);
      propagateAssigneeToTasksFromWeb_(requestId, nextAssignee);
      logActivity_(requestId, 'ASSIGNEE_CHANGED_FROM_WEB', beforeAssignee, nextAssignee);
    }
    if (nextStatus && beforeStatus !== nextStatus) changeRequestStatus(requestId, nextStatus);
    refreshAllViews();
    return {ok:true, requestId:requestId};
  } finally {
    lock.releaseLock();
  }
}

// 웹앱 체크박스를 현재 Task 헤더로 반영함. VI timestamp는 변경하지 않음.
function setTaskCompletedFromWeb(taskId,completed){return updateWebTaskCompletion_(taskId,completed,false);}

// 연결된 Task의 담당자를 일괄 변경함.
function propagateAssigneeToTasksFromWeb_(requestId, assignee) {
  // 요청 참여자는 REQUESTS에만 저장함. 단계별 담당자는 TASKS에서 직접 지정함.
}

// CONFIG 시트에서 상태, 우선순위, 작업자, 브랜드 목록을 읽음.
function getWebConfig_() {
  var sheet = getRequiredSheet_(APP_CONFIG.SHEETS.CONFIG);
  var lastRow = Math.max(sheet.getLastRow(), 2);
  var values = sheet.getRange(2, 1, lastRow - 1, 4).getDisplayValues();
  function unique(column) {
    var seen={};
    return values.map(function(r){return String(r[column]||'').trim();}).filter(function(v){
      if(!v || v.indexOf('TODO:')===0 || seen[v]) return false; seen[v]=true; return true;
    });
  }
  return {statuses:requestStatusValuesV2_(), priorities:unique(1), workers:unique(2), brands:unique(3),
    requestTypes:['3P Listing','2P Lite 전환','3P Listing 후 2P Lite 전환'],
    workStatuses:['아무것도 안한 상태','3P Listing 완료','2P Lite 전환 완료']};
}


// Task의 세분화된 작업 상태를 기록하고 기존 완료 지표를 자동 동기화함.
function setTaskWorkStatusFromWeb(taskId,workStatus,barcodeProvided){return updateWebTaskCompletion_(taskId,workStatus,true,barcodeProvided);}

// 요청 유형에 따라 Task가 최종 완료인지 계산하고 기존 지표 컬럼을 갱신함.
function applyDerivedTaskCompletionFromWeb_(sheet,rowNumber,requestId,workStatus){const row=sheet.getRange(rowNumber,1,1,sheet.getLastColumn()).getValues()[0];return updateWebTaskCompletion_(webTaskToken_(sheet.getName(),{row:rowNumber,values:row},String(requestId)),workStatus,true);}

// 새로 생성한 Task에 기본 상태를 기록함.
function initializeTaskWorkStatusFromWeb_(requestId){return syncWorkTypeProgressFromWeb_(requestId);}

// 요청 유형 변경 시 연결된 모든 Task의 완료 여부를 다시 계산함.
function recalculateTaskCompletionFromWeb_(requestId){return withOpsDocumentLock_(function(){return syncRequestSummaries_([requestId]);});}

// V1.5용 컬럼을 안전하게 추가하고 기존 Task 상태를 초기화함.
function setupWorkTypeV15() {
  var requestSheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  var requestHeaders = requestSheet.getRange(1,1,1,requestSheet.getLastColumn()).getDisplayValues()[0];
  if (requestHeaders.indexOf('Request Type') < 0) requestSheet.getRange(1,requestSheet.getLastColumn()+1).setValue('Request Type');
  var taskSheet = getRequiredSheet_(APP_CONFIG.SHEETS.TASKS);
  var taskHeaders = getHeaderMap_(taskSheet);
  if (!taskHeaders['Work Status']) {
    taskSheet.getRange(1,taskSheet.getLastColumn()+1).setValue('Work Status');
    taskHeaders = getHeaderMap_(taskSheet);
  }
  if (taskSheet.getLastRow() > 1) {
    var range = taskSheet.getRange(2,taskHeaders['Work Status'],taskSheet.getLastRow()-1,1);
    var values = range.getDisplayValues().map(function(row){ return [row[0] || '아무것도 안한 상태']; });
    range.setValues(values);
  }
  return {ok:true};
}


// Task 상태를 집계해 REQUESTS의 3P/2P 완료 수치를 저장함.
// 2P 완료는 선행 3P도 끝난 것으로 간주함.
function syncWorkTypeProgressFromWeb_(requestId){return withOpsDocumentLock_(function(){return syncRequestSummaries_([requestId]);});}

// 유형별 진행률 컬럼을 추가하고 기존 Task 상태를 다시 집계함.
function setupWorkProgressV16() {
  var sheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0];
  ['3P Completed VI','2P Completed VI'].forEach(function(name){
    if (headers.indexOf(name) < 0) { sheet.getRange(1,sheet.getLastColumn()+1).setValue(name); headers.push(name); }
  });
  if (sheet.getLastRow() > 1) {
    var ids = sheet.getRange(2,1,sheet.getLastRow()-1,1).getDisplayValues();
    ids.forEach(function(row){ if(String(row[0]).trim()) syncWorkTypeProgressFromWeb_(row[0]); });
  }
  return {ok:true};
}

// Date 값을 브라우저에 전달 가능한 문자열로 변환함.
function formatWebDate_(value, timezone) {
  if (!value) return '';
  if (Object.prototype.toString.call(value)==='[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, timezone, 'yyyy-MM-dd HH:mm');
  }
  return String(value);
}

// 예상하지 못한 상태값이 Kanban을 깨뜨리지 않도록 정규화함.
function normalizeWebStatus_(value) {
  var status=normalizeRequestStatusV2_(value);
  return requestStatusValuesV2_().indexOf(status)>-1 ? status : APP_CONFIG.STATUS.TO_DO;
}
