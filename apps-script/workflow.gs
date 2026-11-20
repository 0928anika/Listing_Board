// 새 요청을 REQUESTS에 한 번만 기록함.
// requester, brand 등 입력 방식은 아직 결정하지 않았으므로 단순 객체를 받는 골격만 제공함.
function createRequest(requestData) {
return withOpsDocumentLock_(function(){

  const sheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  const headers = getHeaderMap_(sheet);
  const requestedAt = new Date();
  const requestId = createRequestId_(requestedAt);
  const row = new Array(sheet.getLastColumn()).fill('');

  function put_(header, value) {
    if (headers[header]) row[headers[header] - 1] = value;
  }

  put_('Request ID', requestId);
  put_('Node', requestData.node || '');
  put_('Brand', requestData.brand || '');
  put_('Requestor', requestData.requester || requestData.requestor || '');
  put_('Request Type', requestData.requestType || '3P Listing');
  put_('Priority', requestData.priority || 'NORMAL');
  put_('Active Assignees', requestData.assignee || '');
  put_('Status', APP_CONFIG.STATUS.TO_DO);
  put_('Requestor Notes', requestData.notes || '');
  put_('Total VI', Number(requestData.totalVi || 0));
  put_('3P Completed VI', 0);
  put_('2P Completed VI', 0);
  put_('Requested At', requestedAt);
  put_('Completed VI', 0);
  put_('Progress', 0);
  put_('Submission ID', requestData.submissionId || '');
  put_('Submission File Name', requestData.submissionFileName || '');
  put_('Source', requestData.source || '');
  put_('A코드', requestData.aCode || '');
  put_('자료 준비 상태', requestData.preparationStatus || '');
  put_('크롤링 필요', requestData.crawlRequired === true);
  put_('크롤링 단계', requestData.crawlStage || '');

  const targetRow = findFirstEmptyRequestRow_(sheet);
  if(!headers.Status)throw Error("REQUESTS Status 헤더가 없습니다.");
  if(targetRow>sheet.getMaxRows())sheet.insertRowsAfter(sheet.getMaxRows(),targetRow-sheet.getMaxRows());
  applyRequestStatusRange_(sheet,targetRow,headers.Status,1);
  sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
  logActivity_(requestId, 'REQUEST_CREATED', '', JSON.stringify(row));
  if (!requestData.skipRefresh) refreshAllViews();
  return requestData.returnDetails ? {requestId:requestId,row:targetRow} : requestId;

});
}

// Request ID로 요청을 찾아 상태를 변경함.
// ONGOING 최초 진입 시 Started At, DONE 최초 진입 시 Completed At을 값으로 고정 기록함.
function changeRequestStatus(requestId, nextStatus) {
return withOpsDocumentLock_(function(){

  nextStatus = assertManualRequestStatus_(nextStatus);
  const sheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  const headers = getHeaderMap_(sheet);
  const rowNumber = findRowByRequestId_(sheet, requestId);
  const before = normalizeRequestStatusV2_(sheet.getRange(rowNumber, headers['Status']).getValue());
  const now = new Date();
  if(nextStatus===APP_CONFIG.STATUS.DROP){
    const reason=headers['Assignee Notes']?String(sheet.getRange(rowNumber,headers['Assignee Notes']).getValue()||'').trim():'';
    if(!reason)throw Error('Assignee Notes에 DROP 사유를 먼저 입력해주세요.');
    sheet.getRange(rowNumber,headers['Status']).setValue(nextStatus);
    if(headers['Cancelled At']&&!sheet.getRange(rowNumber,headers['Cancelled At']).getValue())sheet.getRange(rowNumber,headers['Cancelled At']).setValue(now);
    cancelRequestTasksAndSessionsV1_(requestId,reason,now);
    if(headers['Active Assignees'])sheet.getRange(rowNumber,headers['Active Assignees']).clearContent();
    bumpMyWorkSidebarRevisionV2_([]);
  }else if(before===APP_CONFIG.STATUS.DROP||(headers['Cancelled At']&&sheet.getRange(rowNumber,headers['Cancelled At']).getValue())){
    sheet.getRange(rowNumber,headers['Status']).setValue(nextStatus);
    reactivateCancelledRequestTasksV3_(requestId);
  }
  markRequestReconcileNeeded_([String(requestId)],{tasks:true,sessions:true},'request-status');
  applyRequestTimingV12_(sheet, rowNumber, headers, nextStatus, before, now);
  reconcileRequestSummary_(String(requestId),{preserveStatus:true,reason:'request-status'});
  logActivity_(requestId, 'STATUS_CHANGED', before, nextStatus);
  refreshAllViews();

});
}

// TASK 완료 여부를 변경하고 TASKS의 시작/완료 시각을 고정값으로 기록하는 확장 지점임.
function setTaskCompleted(taskId, completed) {
  // TODO: Task/VI ID 중복 정책과 작업 시작 트리거가 결정되면 구현함.
  throw new Error('TODO: Task 완료 workflow 정책이 아직 결정되지 않았습니다.');
}

// REQUESTS의 VI 수량, 진행률, 소요시간을 TASKS와 timestamp 값으로부터 다시 계산함.
function updateRequestMetrics_(sheet, rowNumber, headers) {
  const id=String(sheet.getRange(rowNumber,headers['Request ID']).getValue()||'');
  return syncRequestSummaries_([id],{preserveStatus:true});
}

// 허용되지 않은 상태값이 원본 데이터에 기록되지 않도록 방어함.
function assertAllowedStatus_(status) {
  const normalized = normalizeRequestStatusV2_(status);
  const allowed = Object.keys(APP_CONFIG.STATUS).map(function(key) { return APP_CONFIG.STATUS[key]; });
  if (allowed.indexOf(normalized) === -1) throw new Error('허용되지 않은 상태입니다: ' + status);
  return normalized;
}



// 미리 채워진 수식이나 서식 때문에 마지막 행이 커져 있어도
// Request ID가 비어 있는 첫 번째 데이터 행을 찾아 반환함.
function findFirstEmptyRequestRow_(sheet) {
  const firstDataRow = 2;
  const lastRowToCheck = Math.max(sheet.getLastRow(), firstDataRow);
  const requestIds = sheet
    .getRange(firstDataRow, 1, lastRowToCheck - firstDataRow + 1, 1)
    .getDisplayValues();

  for (let index = 0; index < requestIds.length; index += 1) {
    if (!String(requestIds[index][0]).trim()) {
      return firstDataRow + index;
    }
  }

  return lastRowToCheck + 1;
}


// 상태 전환 시 ON HOLD 구간을 누적하고 Lead/Work Time을 고정값으로 계산함.
// Lead Time은 요청 생성~완료, Work Time은 배정~완료에서 ON HOLD 누적시간을 뺀 값임.
function applyRequestTimingV12_(sheet, rowNumber, headers, nextStatus, previousStatus, now) {
  now = now || new Date();
  nextStatus = assertAllowedStatus_(nextStatus);
  previousStatus = normalizeRequestStatusV2_(previousStatus);
  const blockedAtCell = sheet.getRange(rowNumber, headers['Blocked At']);
  const resumedAtCell = sheet.getRange(rowNumber, headers['Resumed At']);
  const completedAtCell = sheet.getRange(rowNumber, headers['Completed At']);
  const blockedHoursCell = sheet.getRange(rowNumber, headers['Blocked Time (Hours)']);
  const leadHoursCell = sheet.getRange(rowNumber, headers['Lead Time (Hours)']);
  const workHoursCell = sheet.getRange(rowNumber, headers['Work Time (Hours)']);
  let blockedHours = Number(blockedHoursCell.getValue() || 0);
  const blockedAt = blockedAtCell.getValue();

  if (nextStatus === APP_CONFIG.STATUS.ON_HOLD && previousStatus !== APP_CONFIG.STATUS.ON_HOLD) {
    blockedAtCell.setValue(now);
  }

  if (previousStatus === APP_CONFIG.STATUS.ON_HOLD && nextStatus !== APP_CONFIG.STATUS.ON_HOLD) {
    if (blockedAt instanceof Date && !isNaN(blockedAt.getTime())) {
      blockedHours += Math.max(0, (now - blockedAt) / 3600000);
      blockedHoursCell.setValue(blockedHours);
    }
    resumedAtCell.setValue(now);
  }

  if (nextStatus === APP_CONFIG.STATUS.DONE) {
    const timingRequestId = headers['Request ID'] ? String(sheet.getRange(rowNumber, headers['Request ID']).getValue() || '') : '';
    if (timingRequestId && typeof closeCompletedWorkSessions_ === 'function') closeCompletedWorkSessions_(timingRequestId, now);
    let completedAt = completedAtCell.getValue();
    if (!(completedAt instanceof Date) || isNaN(completedAt.getTime())) {
      completedAt = now;
      completedAtCell.setValue(completedAt);
    }
    const requestedAt = headers['Requested At'] ? sheet.getRange(rowNumber, headers['Requested At']).getValue() : '';
    const startedAt = headers['Started At'] ? sheet.getRange(rowNumber, headers['Started At']).getValue() : '';
    if (requestedAt instanceof Date && !isNaN(requestedAt.getTime())) {
      leadHoursCell.setValue(Math.max(0, (completedAt - requestedAt) / 3600000));
    }
    if (!timingRequestId && startedAt instanceof Date && !isNaN(startedAt.getTime())) {
      workHoursCell.setValue(Math.max(0, (completedAt - startedAt) / 3600000 - blockedHours));
    }
  } else if (previousStatus === APP_CONFIG.STATUS.DONE) {
    completedAtCell.clearContent();
    leadHoursCell.clearContent();
    if (typeof getWorkSessionSnapshot_ !== 'function') workHoursCell.clearContent();
  }

  writeRequestStatuses_(sheet,rowNumber,headers['Status'],[nextStatus]);
  sheet.getRange(rowNumber, headers['Started At'], 1, 4).setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange(rowNumber, headers['Blocked Time (Hours)'], 1, 3).setNumberFormat('0.00 "h"');
}
