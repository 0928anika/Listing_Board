// 통합 요청 등록과 Task 생성에서 공통으로 사용하는 최소 유틸리티임.
function syncRequestMetricsFromTasks_(requestId) {
  // V2에서는 두 단계 체크박스를 기준으로 요청 지표를 계산함.
  syncRequestLayoutV6_(requestId);
}

function findFirstEmptyTaskRow_(sheet) {
  const firstDataRow = 2;
  const lastRow = Math.max(sheet.getLastRow(), firstDataRow);
  const ids = sheet.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, 1).getDisplayValues();
  for (let index = 0; index < ids.length; index += 1) {
    if (!String(ids[index][0]).trim()) return firstDataRow + index;
  }
  return lastRow + 1;
}

// 실제 생성된 Task 행에만 드롭다운과 체크박스를 일괄 적용함.
function applyNewTaskControlsV2_(sheet, startRow, rowCount, requestType) {
  if (!rowCount) return;
  const headers = getHeaderMap_(sheet);
  const configSheet = getRequiredSheet_(APP_CONFIG.SHEETS.CONFIG);
  const configHeaders = getHeaderMap_(configSheet);
  const workerColumn = configHeaders['Worker'] || 3;
  const workerRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(configSheet.getRange(2,workerColumn,Math.max(configSheet.getMaxRows()-1,1),1), true)
    .setAllowInvalid(false).build();
  const viRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['VALID','INVALID'], true).setAllowInvalid(false).build();
  if (headers['VI Status']) sheet.getRange(startRow,headers['VI Status'],rowCount,1).setDataValidation(viRule);
  if (sheet.getName() === APP_CONFIG.SHEETS.TASKS_2P) {
    if (headers['2P Assignee']) sheet.getRange(startRow,headers['2P Assignee'],rowCount,1).setDataValidation(workerRule);
    if (headers['2P Completed']) sheet.getRange(startRow,headers['2P Completed'],rowCount,1).insertCheckboxes().setValue(false);
    return;
  }
  if (requestType !== '2P Lite 전환') {
    if (headers['3P Assignee']) sheet.getRange(startRow,headers['3P Assignee'],rowCount,1).setDataValidation(workerRule);
    if (headers['3P Completed']) sheet.getRange(startRow,headers['3P Completed'],rowCount,1).insertCheckboxes().setValue(false);
  }
  if (requestType !== '3P Listing') {
    if (headers['2P Assignee']) sheet.getRange(startRow,headers['2P Assignee'],rowCount,1).setDataValidation(workerRule);
    if (headers['2P Completed']) sheet.getRange(startRow,headers['2P Completed'],rowCount,1).insertCheckboxes().setValue(false);
  }
}
