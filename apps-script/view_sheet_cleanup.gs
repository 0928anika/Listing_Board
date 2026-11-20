// Popup으로 대체되어 더 이상 필요하지 않은 파생 View 시트임.
const OBSOLETE_VIEW_SHEET_NAMES_V1 = Object.freeze(['KANBAN', 'DASHBOARD']);

function inspectObsoleteViewSheetsV1_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = OBSOLETE_VIEW_SHEET_NAMES_V1.map(function(name) {
    const sheet = ss.getSheetByName(name);
    return sheet ? {name:name, sheetId:sheet.getSheetId()} : null;
  }).filter(Boolean);
  const existingNames = existing.map(function(item){return item.name;});
  return {
    spreadsheetId:ss.getId(),
    existing:existing,
    missing:OBSOLETE_VIEW_SHEET_NAMES_V1.filter(function(name){return existingNames.indexOf(name)<0;})
  };
}

// 삭제 전 exact-name 대상과 Sheet ID만 확인하는 읽기 전용 함수임.
function previewObsoleteViewSheetRemovalV1() {
  const result = inspectObsoleteViewSheetsV1_();
  console.log(JSON.stringify(result));
  return result;
}

// KANBAN/DASHBOARD 두 탭만 삭제함. 이미 없으면 변경 없이 성공함.
function removeObsoleteViewSheetsV1() {
  return withOpsDocumentLock_(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const before = inspectObsoleteViewSheetsV1_();
    if (ss.getSheets().length - before.existing.length < 1) {
      throw new Error('마지막 시트는 삭제할 수 없습니다. 원본을 수정하지 않았습니다.');
    }
    before.existing.forEach(function(item) {
      const sheet = ss.getSheetByName(item.name);
      if (!sheet || sheet.getSheetId() !== item.sheetId) {
        throw new Error(item.name + ' 시트가 사전점검 후 변경되었습니다. 원본을 수정하지 않았습니다.');
      }
    });
    before.existing.forEach(function(item) {
      ss.deleteSheet(ss.getSheetByName(item.name));
    });
    const after = inspectObsoleteViewSheetsV1_();
    if (after.existing.length) {
      throw new Error('View 시트 삭제 검증 실패: ' + after.existing.map(function(item){return item.name;}).join(', '));
    }
    const result = {
      ok:true,
      deleted:before.existing.map(function(item){return item.name;}),
      alreadyMissing:before.missing,
      verifiedMissing:after.missing
    };
    console.log(JSON.stringify(result));
    return result;
  });
}

// 삭제 후 두 탭이 다시 생성되지 않았는지 확인하는 읽기 전용 함수임.
function verifyObsoleteViewSheetsRemovedV1() {
  const result = inspectObsoleteViewSheetsV1_();
  if (result.existing.length) {
    throw new Error('아직 남아 있는 View 시트: ' + result.existing.map(function(item){return item.name;}).join(', '));
  }
  result.ok = true;
  console.log(JSON.stringify(result));
  return result;
}
