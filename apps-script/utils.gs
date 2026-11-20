// 이름으로 시트를 찾고 없으면 명확한 오류를 발생시킴.
function getRequiredSheet_(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error('필수 시트를 찾을 수 없습니다: ' + sheetName);
  return sheet;
}

// 시트가 없으면 만들고, 헤더가 비어 있을 때만 지정된 헤더를 기록함.
function ensureSheetWithHeaders_(ss, sheetName, headers) {
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  if (headers.length && sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader_(sheet.getRange(1, 1, 1, headers.length));
  }
  return sheet;
}

// 첫 행의 헤더명을 1부터 시작하는 컬럼 번호로 변환함.
function getHeaderMap_(sheet) {
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  return headers.reduce(function(map, header, index) {
    if (header) map[String(header)] = index + 1;
    return map;
  }, {});
}

// 헤더를 제외한 실제 데이터 행만 반환함.
function getDataRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues()
    .filter(function(row) { return row[0] !== ''; });
}

// Request ID가 있는 행 번호를 찾음. 찾지 못하면 오류를 발생시킴.
function findRowByRequestId_(sheet, requestId) {
  const finder = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1)
    .createTextFinder(String(requestId)).matchEntireCell(true).findNext();
  if (!finder) throw new Error('Request ID를 찾을 수 없습니다: ' + requestId);
  return finder.getRow();
}

// 현재 REQUESTS의 가장 큰 순번 다음 번호로 REQ-001 형식의 ID를 생성함.
function createRequestId_() {
  const sheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  if (sheet.getLastRow() < 2) return APP_CONFIG.REQUEST_ID_PREFIX + '-001';
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  let maxNumber = 0;
  ids.forEach(function(row) {
    const match = String(row[0] || '').match(/^REQ-(\d+)$/);
    if (match) maxNumber = Math.max(maxNumber, Number(match[1]));
  });
  return APP_CONFIG.REQUEST_ID_PREFIX + '-' + String(maxNumber + 1).padStart(3, '0');
}

// 기존 긴 Request ID를 REQ-001 형식으로 일괄 변환함.
// REQUESTS, TASKS, Activity Log와 Task/VI ID를 같은 매핑으로 함께 갱신함.
function migrateRequestIdsToSimpleFormat() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('다른 작업을 처리 중입니다. 잠시 후 다시 실행해 주세요.');
  try {
    const requestSheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
    const lastRequestRow = requestSheet.getLastRow();
    if (lastRequestRow < 2) return { ok: true, converted: 0 };

    const requestIds = requestSheet.getRange(2, 1, lastRequestRow - 1, 1).getDisplayValues();
    const mapping = {};
    let maxNumber = 0;
    requestIds.forEach(function(row) {
      const id = String(row[0] || '').trim();
      const match = id.match(/^REQ-(\d+)$/);
      if (match) maxNumber = Math.max(maxNumber, Number(match[1]));
    });
    let converted = 0;
    const newRequestIds = requestIds.map(function(row) {
      const oldId = String(row[0] || '').trim();
      if (!oldId) return [''];
      if (/^REQ-\d+$/.test(oldId)) { mapping[oldId] = oldId; return [oldId]; }
      maxNumber += 1;
      const newId = 'REQ-' + String(maxNumber).padStart(3, '0');
      mapping[oldId] = newId;
      converted += 1;
      return [newId];
    });
    requestSheet.getRange(2, 1, newRequestIds.length, 1).setValues(newRequestIds);

    const taskSheet = getRequiredSheet_(APP_CONFIG.SHEETS.TASKS);
    if (taskSheet.getLastRow() > 1) {
      const taskRange = taskSheet.getRange(2, 1, taskSheet.getLastRow() - 1, 2);
      const taskValues = taskRange.getDisplayValues();
      const sequences = {};
      taskValues.forEach(function(row) {
        const oldRequestId = String(row[0] || '').trim();
        const newRequestId = mapping[oldRequestId] || oldRequestId;
        if (!newRequestId) return;
        sequences[newRequestId] = (sequences[newRequestId] || 0) + 1;
        row[0] = newRequestId;
        row[1] = newRequestId + '-VI-' + String(sequences[newRequestId]).padStart(3, '0');
      });
      taskRange.setValues(taskValues);
    }

    const logSheet = getRequiredSheet_(APP_CONFIG.SHEETS.ACTIVITY_LOG);
    if (logSheet.getLastRow() > 1) {
      const logRange = logSheet.getRange(2, 3, logSheet.getLastRow() - 1, 1);
      const logValues = logRange.getDisplayValues().map(function(row) {
        const oldId = String(row[0] || '').trim();
        return [mapping[oldId] || oldId];
      });
      logRange.setValues(logValues);
    }

    refreshAllViews();
    return { ok: true, converted: converted, mapping: mapping };
  } finally {
    lock.releaseLock();
  }
}

// 현재 실행 사용자를 기록함. 도메인 정책에 따라 이메일이 비어 있을 수 있음.
function getCurrentUser_() {
  return Session.getActiveUser().getEmail() || 'UNKNOWN_USER';
}

// CONFIG 범위를 참조하는 드롭다운 검증 규칙을 적용함.
function applyListValidation_(targetRange, sourceRange) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(false)
    .build();
  targetRange.setDataValidation(rule);
}

// 공통 헤더 스타일을 적용함.
function styleHeader_(range) {
  range.setBackground('#16324F')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
}



// 기존 요청 유형 '둘 다'를 새 명칭으로 일괄 변경하고 View를 갱신함.
function migrateRequestTypeLabelV18() {
  const requestSheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  const headers = getHeaderMap_(requestSheet);
  if (requestSheet.getLastRow() > 1) {
    const range = requestSheet.getRange(2, headers['Request Type'], requestSheet.getLastRow() - 1, 1);
    const values = range.getDisplayValues().map(function(row) {
      return [String(row[0]) === '둘 다' ? '3P Listing 후 2P Lite 전환' : row[0]];
    });
    range.setValues(values);
  }
  refreshAllViews();
  return { ok: true };
}


// REQUESTS와 TASKS에 Node 컬럼을 추가함.
// 기존 데이터는 변경하지 않고 빈 Node로 유지함.
function setupNodeColumnsV19() {
  const requestSheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  let requestHeaders = requestSheet.getRange(1, 1, 1, requestSheet.getLastColumn()).getDisplayValues()[0];
  if (requestHeaders.indexOf('Node') < 0) {
    requestSheet.getRange(1, requestSheet.getLastColumn() + 1).setValue('Node');
  }

  const taskSheet = getRequiredSheet_(APP_CONFIG.SHEETS.TASKS);
  const taskHeaders = getHeaderMap_(taskSheet);
  if (!taskHeaders['Node']) taskSheet.getRange(1, taskSheet.getLastColumn() + 1).setValue('Node');
  return { ok: true };
}
