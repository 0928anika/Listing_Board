// 이전 TASKS 컬럼 구조에서 남은 데이터 검증 규칙을 제거하고
// V0.2 컬럼 위치에 맞는 검증만 다시 적용함.
// 셀 값과 Task 데이터는 삭제하지 않음.
function repairTaskValidationsV02() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [APP_CONFIG.SHEETS.TASKS,APP_CONFIG.SHEETS.CRAWL_TASKS,APP_CONFIG.SHEETS.TASKS_2P].forEach(function(name) {
    const sheet = getRequiredSheet_(name);
    const blankStart = Math.max(sheet.getLastRow() + 1, 2);
    const blankCount = sheet.getMaxRows() - blankStart + 1;
    if (blankCount > 0) sheet.getRange(blankStart,1,blankCount,sheet.getLastColumn()).clearDataValidations();
  });
  if (typeof ensureValidityManagementV3_ === 'function') ensureValidityManagementV3_();

  ss.toast('빈 Task 행의 드롭다운·체크박스 규칙을 정리했습니다.', 'Listing Ops', 5);
}
