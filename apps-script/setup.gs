// Listing Operations System의 기본 시트와 UI를 준비함.
function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetWithHeaders_(ss, APP_CONFIG.SHEETS.REQUESTS, APP_CONFIG.REQUEST_COLUMNS);
  const task3P = ensureSheetWithHeaders_(ss, APP_CONFIG.SHEETS.TASKS, TASK_HEADERS_LAYOUT_V6);
  const crawl3P = ensureSheetWithHeaders_(ss, APP_CONFIG.SHEETS.CRAWL_TASKS, TASK_HEADERS_LAYOUT_V6);
  const task2P = ensureSheetWithHeaders_(ss, APP_CONFIG.SHEETS.TASKS_2P, TASK_HEADERS_2P_V8);
  [task3P,crawl3P,task2P].forEach(ensureTaskSetupCapacityV2_);
  ensureSheetWithHeaders_(ss, APP_CONFIG.SHEETS.ACTIVITY_LOG, APP_CONFIG.ACTIVITY_COLUMNS);
  ensureSheetWithHeaders_(ss, APP_CONFIG.SHEETS.CONFIG, ['Status', 'Priority', 'Worker', 'Brand']);
  ensureConfigNodeACodeColumns_();
  installBarcodeProvidedWorkflowV1();
  ensureArchiveRequestColumnsV1_();
  ensureLongTermSessionSheetV1_();
  configureRequestsSheet_();
  configure3PTasksSheet_();
  configureCrawl3PTasksSheet_();
  configure2PTasksSheet_();
  configureActivityLogSheet_();
  if (typeof applySemanticStatusColorsV4_ === 'function') applySemanticStatusColorsV4_();
  refreshAllViews();
  ss.toast('Listing Operations System 기본 설정이 완료되었습니다.', 'Listing Ops', 5);
}

const TASK_SETUP_DATA_ROWS_V2 = 50000;

function ensureTaskSetupCapacityV2_(sheet) {
  const requiredRows = TASK_SETUP_DATA_ROWS_V2 + 1;
  if (sheet.getMaxRows() < requiredRows) sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
}

// 기존 운영 문서에 빈 행 검증 제거·전체 조건부 서식 정책만 안전하게 반영함.
function applyTaskSetupPolicyV2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [APP_CONFIG.SHEETS.TASKS,APP_CONFIG.SHEETS.CRAWL_TASKS,APP_CONFIG.SHEETS.TASKS_2P]
    .map(function(name){return getRequiredSheet_(name);})
    .forEach(ensureTaskSetupCapacityV2_);
  configure3PTasksSheet_();
  configureCrawl3PTasksSheet_();
  configure2PTasksSheet_();
  repairTaskValidationsV02();
  if (typeof applySemanticStatusColorsV4_ === 'function') applySemanticStatusColorsV4_();
  ss.toast('빈 Task 행 검증 제거 및 전체 조건부 서식 적용 완료','Listing Ops',5);
}

// 기존 운영 문서에 크롤링 워크플로만 안전하게 추가함.
function applyCrawlingWorkflowUpdate() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetWithHeaders_(ss,APP_CONFIG.SHEETS.CRAWL_TASKS,TASK_HEADERS_LAYOUT_V6);
  ensureRequestIntakeMetadataColumnsV2_();
  configureCrawl3PTasksSheet_();
  configureCrawlRequestVisibility_();
  refreshAllViews();
  ss.toast('크롤링 접수·결과 등록 워크플로가 반영되었습니다.','Listing Ops',5);
}

function configureCrawlRequestVisibility_() {
  const sheet=getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS); let h=ensureRequestIntakeMetadataColumnsV2_();
  if(h['Request Type']&&h['자료 준비 상태']&&h['자료 준비 상태']!==h['Request Type']+1) {sheet.moveColumns(sheet.getRange(1,h['자료 준비 상태'],sheet.getMaxRows(),1),h['Request Type']+1);h=getHeaderMap_(sheet);}
  if(h['자료 준비 상태']&&h['크롤링 작업자']&&h['크롤링 작업자']!==h['자료 준비 상태']+1) {sheet.moveColumns(sheet.getRange(1,h['크롤링 작업자'],sheet.getMaxRows(),1),h['자료 준비 상태']+1);h=getHeaderMap_(sheet);}
  const count=Math.max(sheet.getLastRow()-1,0); if(!count)return;
  const rows=sheet.getRange(2,1,count,sheet.getLastColumn()).getValues();
  const values=rows.map(function(row){const crawl=row[h['크롤링 필요']-1]===true,stage=String(row[h['크롤링 단계']-1]||'');return [crawl?(stage==='CRAWL_COMPLETED'?'크롤링 완료 · 3P 리스팅 가능':(String(row[h['크롤링 작업자']-1]||'')?'크롤링 배정됨':'크롤링 담당자 배정 필요')):'바로 리스팅 가능'];});
  const statusRange=sheet.getRange(2,h['자료 준비 상태'],count,1);statusRange.setValues(values).setFontWeight('bold');sheet.setColumnWidth(h['자료 준비 상태'],190);sheet.setColumnWidth(h['크롤링 작업자'],120);
  const rules=sheet.getConditionalFormatRules().filter(function(rule){return rule.getRanges().every(function(r){return r.getColumn()!==h['자료 준비 상태'];});});
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('바로 리스팅 가능').setBackground('#DCFCE7').setFontColor('#166534').setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('크롤링 담당자').setBackground('#FFF1D6').setFontColor('#9A5B0A').setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('크롤링 배정됨').setBackground('#F3E8FF').setFontColor('#7E22CE').setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('크롤링 완료').setBackground('#DBEAFE').setFontColor('#315B8A').setRanges([statusRange]).build());sheet.setConditionalFormatRules(rules);
  ['크롤링 필요','크롤링 단계'].forEach(function(name){h=getHeaderMap_(sheet);if(h[name])sheet.showColumns(h[name]);});
}

function configureCrawl3PTasksSheet_() {
  const layout = ensureTaskMemoColumnForSheetV1_(APP_CONFIG.SHEETS.CRAWL_TASKS);
  const sheet = layout.sheet, headers = layout.headers;
  sheet.setFrozenRows(1); sheet.setFrozenColumns(3);
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setBackground('#6B4F8A').setFontColor('#FFFFFF').setFontWeight('bold');
  const rows = Math.max(sheet.getMaxRows() - 1, 1);
  ['판매가','정상가'].forEach(function(name){if(headers[name])sheet.getRange(2,headers[name],rows,1).setNumberFormat('#,##0');});
  styleTaskMemoColumnV1_(sheet,headers);
  if (headers['Internal Task ID']) sheet.hideColumns(headers['Internal Task ID']);
  applyTaskConditionalFormattingV9_(sheet);
}

// REQUESTS의 표시 형식과 현재 위치에 맞는 검증 규칙을 적용함.
function configureRequestsSheet_() {
  const sheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  const rows = Math.max(sheet.getMaxRows() - 1, 1), headers = getHeaderMap_(sheet);
  sheet.setFrozenRows(1); sheet.setFrozenColumns(1);
  sheet.getRange(1,1,1,sheet.getLastColumn()).setBackground('#263650').setFontColor('#FFFFFF').setFontWeight('bold');
  if(headers['Request Type'])sheet.getRange(2,headers['Request Type'],rows,1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['3P Listing','2P Lite 전환','3P Listing 후 2P Lite 전환'],true).setAllowInvalid(false).build());
  if(headers['Priority'])sheet.getRange(2,headers['Priority'],rows,1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(APP_CONFIG.PRIORITY,true).setAllowInvalid(false).build());
  if(headers['Status'])sheet.getRange(2,headers['Status'],rows,1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(manualRequestStatuses_(),true).setAllowInvalid(false).build());
  ['Total VI','3P Completed VI','2P Completed VI'].forEach(function(name){if(headers[name])sheet.getRange(2,headers[name],rows,1).setNumberFormat('0');});
  ['Requested At','Started At','Blocked At','Resumed At','Completed At','Cancelled At','Archived At'].forEach(function(name){if(headers[name])sheet.getRange(2,headers[name],rows,1).setNumberFormat('yyyy-mm-dd hh:mm');});
  ['Blocked Time (Hours)','Lead Time (Hours)','Work Time (Hours)'].forEach(function(name){if(headers[name])sheet.getRange(2,headers[name],rows,1).setNumberFormat('0.00 \"h\"');});
  if(headers['Progress'])sheet.getRange(2,headers['Progress'],rows,1).setNumberFormat('0%');
  ['Requested At','Progress'].forEach(function(name){if(headers[name])sheet.hideColumns(headers[name]);});
  if(headers['Barcode Provided']){
    sheet.showColumns(headers['Barcode Provided']);
    applyBarcodeProvidedValidationV1_(sheet,headers['Barcode Provided']);
  }
  ['Archive Status','Archived At','Archive Batch ID','Archived Task Rows','Archive File ID'].forEach(function(name){
    if(headers[name])sheet.showColumns(headers[name]);
  });
  applyRequestFormattingV10_(sheet);
}

// 3P_TASKS의 체크박스와 표시 형식을 적용함.
function configure3PTasksSheet_() {
  const layout = ensureTaskMemoColumnForSheetV1_(APP_CONFIG.SHEETS.TASKS);
  const sheet = layout.sheet, headers = layout.headers;
  sheet.setFrozenRows(1); sheet.setFrozenColumns(3);
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setBackground('#263650').setFontColor('#FFFFFF').setFontWeight('bold');
  const rows = Math.max(sheet.getMaxRows() - 1, 1);
  ['판매가','정상가'].forEach(function(name){if(headers[name])sheet.getRange(2,headers[name],rows,1).setNumberFormat('#,##0');});
  styleTaskMemoColumnV1_(sheet,headers);
  if (headers['Internal Task ID']) sheet.hideColumns(headers['Internal Task ID']);
  applyTaskConditionalFormattingV9_(sheet);
}

const TASK_HEADERS_2P_V8 = [
  'Request ID','업체상품ID','Product ID','3P 옵션 ID','바코드','업체상품코드',
  '쿠팡노출상품명','업체 등록 상품명','등록 옵션명','Memo','2P Assignee','2P Completed',
  'VI Status','Invalid Reason','Work Status','Internal Task ID'
];

const TASK_MEMO_HEADER_V1 = 'Memo';
const TASK_MEMO_LEGACY_HEADER_V1 = '메모';

function taskMemoExpectedHeadersV1_(sheetName) {
  const source = sheetName === APP_CONFIG.SHEETS.TASKS_2P ? TASK_HEADERS_2P_V8 : TASK_HEADERS_LAYOUT_V6;
  const anchor = sheetName === APP_CONFIG.SHEETS.TASKS_2P ? '등록 옵션명' : 'Brand';
  const headers = source.slice(), anchorIndex = headers.indexOf(anchor), memoIndex = headers.indexOf(TASK_MEMO_HEADER_V1);
  if (anchorIndex < 0) throw new Error(sheetName + '의 ' + anchor + ' 헤더를 확인해 주세요. 원본을 수정하지 않았습니다.');
  if (memoIndex < 0) headers.splice(anchorIndex + 1,0,TASK_MEMO_HEADER_V1);
  else if (memoIndex !== anchorIndex + 1) throw new Error(sheetName + '의 코드 헤더 순서에서 Memo 위치를 확인해 주세요.');
  return {
    sheetName:sheetName,
    anchor:anchor,
    headers:headers,
    korean:headers.map(function(name){return name===TASK_MEMO_HEADER_V1?TASK_MEMO_LEGACY_HEADER_V1:name;}),
    legacy:headers.filter(function(name){return name!==TASK_MEMO_HEADER_V1;})
  };
}

function taskMemoSheetPlanV1_(sheet,spec,allowLegacy) {
  const lastColumn=sheet.getLastColumn();
  const current=lastColumn ? sheet.getRange(1,1,1,lastColumn).getDisplayValues()[0].map(function(value){return String(value||'').trim();}) : [];
  function same_(left,right){return left.length===right.length&&left.every(function(value,index){return value===right[index];});}
  if(same_(current,spec.headers))return {sheet:sheet,spec:spec,action:'READY'};
  if(allowLegacy&&same_(current,spec.korean))return {sheet:sheet,spec:spec,action:'RENAME',memoColumn:spec.korean.indexOf(TASK_MEMO_LEGACY_HEADER_V1)+1};
  if(allowLegacy&&same_(current,spec.legacy))return {sheet:sheet,spec:spec,action:'INSERT',anchorColumn:spec.legacy.indexOf(spec.anchor)+1};
  throw new Error(spec.sheetName+'의 헤더 구성이 예상과 다릅니다. Memo 열을 변경하지 않았습니다.');
}

function taskMemoSpecsV1_() {
  return [APP_CONFIG.SHEETS.TASKS,APP_CONFIG.SHEETS.CRAWL_TASKS,APP_CONFIG.SHEETS.TASKS_2P].map(taskMemoExpectedHeadersV1_);
}

function verifyTaskMemoColumnsV1_() {
  const sheets=taskMemoSpecsV1_().map(function(spec){
    const plan=taskMemoSheetPlanV1_(getRequiredSheet_(spec.sheetName),spec,false),h=getHeaderMap_(plan.sheet);
    if(h[TASK_MEMO_HEADER_V1]!==h[spec.anchor]+1)throw new Error(spec.sheetName+'의 Memo 열 위치가 잘못되었습니다.');
    return {sheet:spec.sheetName,memoColumn:h[TASK_MEMO_HEADER_V1],lastColumn:plan.sheet.getLastColumn()};
  });
  return {ok:true,sheets:sheets};
}

// 읽기 전용 공개 검증. 헤더/순서가 다르면 원본을 수정하지 않고 실패함.
function verifyTaskMemoColumnsV1() {
  const result=verifyTaskMemoColumnsV1_();
  console.log(JSON.stringify(result));
  return result;
}

// 쓰기 전 공개 사전점검. 세 시트가 기존 또는 신규 정식 스키마인지 한 번에 확인함.
function previewTaskMemoColumnsV1() {
  const plans=taskMemoSpecsV1_().map(function(spec){
    const plan=taskMemoSheetPlanV1_(getRequiredSheet_(spec.sheetName),spec,true);
    return {sheet:spec.sheetName,action:plan.action,anchor:spec.anchor,targetMemoColumn:spec.headers.indexOf(TASK_MEMO_HEADER_V1)+1};
  });
  const result={ok:true,sheets:plans};
  console.log(JSON.stringify(result));
  return result;
}

function applyTaskMemoColumnPlanV1_(plan) {
  if(plan.action==='INSERT'){
    plan.sheet.insertColumnsAfter(plan.anchorColumn,1);
    plan.sheet.getRange(1,plan.anchorColumn+1).setValue(TASK_MEMO_HEADER_V1);
    return 'INSERT';
  }
  if(plan.action==='RENAME'){
    plan.sheet.getRange(1,plan.memoColumn).setValue(TASK_MEMO_HEADER_V1);
    return 'RENAME';
  }
  return 'READY';
}

function styleTaskMemoColumnV1_(sheet,headers) {
  const memoColumn=headers[TASK_MEMO_HEADER_V1];if(!memoColumn)return;
  sheet.setColumnWidth(memoColumn,220);
  sheet.getRange(1,memoColumn).setWrap(true);
  const bodyRows=sheet.getLastRow()-1;
  if(bodyRows>0)sheet.getRange(2,memoColumn,bodyRows,1).setWrap(true).setVerticalAlignment('middle');
}

// 세 Task 시트를 모두 먼저 검증한 뒤 실제 열을 삽입함.
// 재실행해도 열을 중복 추가하지 않으며 Task 행 데이터를 다시 쓰지 않음.
function installTaskMemoColumnsV1() {
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))throw new Error('다른 변경 처리 중입니다. 잠시 후 다시 실행해 주세요.');
  try{
    // Mutation before validation of every source would leave a partial schema.
    const plans=taskMemoSpecsV1_().map(function(spec){return taskMemoSheetPlanV1_(getRequiredSheet_(spec.sheetName),spec,true);});
    const inserted=[],renamed=[];
    plans.forEach(function(plan){
      const action=applyTaskMemoColumnPlanV1_(plan);
      if(action==='INSERT')inserted.push(plan.spec.sheetName);
      if(action==='RENAME')renamed.push(plan.spec.sheetName);
    });
    const verified=verifyTaskMemoColumnsV1_();
    verified.sheets.forEach(function(result){
      const sheet=getRequiredSheet_(result.sheet),h=getHeaderMap_(sheet);
      styleTaskMemoColumnV1_(sheet,h);
      if(h['Internal Task ID'])sheet.hideColumns(h['Internal Task ID']);
    });
    if(typeof bumpMyWorkSidebarRevisionV2_==='function')bumpMyWorkSidebarRevisionV2_([]);
    const result={ok:true,inserted:inserted,renamed:renamed,alreadyReady:plans.filter(function(plan){return plan.action==='READY';}).map(function(plan){return plan.spec.sheetName;}),sheets:verified.sheets};
    console.log(JSON.stringify(result));
    return result;
  }finally{lock.releaseLock();}
}

// 호환 진입점: 정상 설정에서는 행 데이터를 재작성하지 않고 필요한 열만 물리적으로 삽입함.
function ensureTwoPTaskLayoutV8_() {
  return ensureTaskMemoColumnForSheetV1_(APP_CONFIG.SHEETS.TASKS_2P);
}

function ensureTaskMemoColumnForSheetV1_(sheetName) {
  const spec=taskMemoExpectedHeadersV1_(sheetName),sheet=getRequiredSheet_(sheetName),plan=taskMemoSheetPlanV1_(sheet,spec,true);
  applyTaskMemoColumnPlanV1_(plan);
  const verified=taskMemoSheetPlanV1_(sheet,spec,false),headers=getHeaderMap_(sheet);
  if(headers[TASK_MEMO_HEADER_V1]!==headers[spec.anchor]+1)throw new Error(sheetName+'의 Memo 열 위치가 잘못되었습니다.');
  return {sheet:verified.sheet,headers:headers,inserted:plan.action==='INSERT',renamed:plan.action==='RENAME'};
}

// 정상 운영 중에는 헤더만 확인하고 기존 2P 데이터는 다시 쓰지 않음.
function getTwoPTaskLayoutV8_() {
  const sheet = getRequiredSheet_(APP_CONFIG.SHEETS.TASKS_2P);
  const headers = getHeaderMap_(sheet);
  const invalid = TASK_HEADERS_2P_V8.some(function(name,index){return headers[name] !== index+1;});
  if (invalid || sheet.getLastColumn() !== TASK_HEADERS_2P_V8.length) {
    throw new Error('2P_TASKS 열 구성이 변경되었습니다. 관리자에게 2P 양식 정리를 요청해 주세요.');
  }
  return {sheet:sheet,headers:headers};
}

// 2P_TASKS의 체크박스와 표시 형식을 적용함.
function configure2PTasksSheet_() {
  const layout = ensureTwoPTaskLayoutV8_();
  const sheet = layout.sheet, headers = layout.headers;
  sheet.getRange(1,1,1,sheet.getLastColumn()).setBackground('#263650').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.setFrozenRows(1); sheet.setFrozenColumns(3);
  const dataRows = Math.max(sheet.getLastRow()-1,0);
  if (dataRows) {
    const configSheet = getRequiredSheet_(APP_CONFIG.SHEETS.CONFIG), configHeaders = getHeaderMap_(configSheet);
    const workerColumn = configHeaders['Worker'] || 3;
    const workerRule = SpreadsheetApp.newDataValidation().requireValueInRange(
      configSheet.getRange(2,workerColumn,Math.max(configSheet.getMaxRows()-1,1),1),true
    ).setAllowInvalid(false).build();
    const viRule = SpreadsheetApp.newDataValidation().requireValueInList(['VALID','INVALID'],true).setAllowInvalid(false).build();
    sheet.getRange(2,headers['2P Assignee'],dataRows,1).setDataValidation(workerRule);
    const completed = sheet.getRange(2,headers['2P Completed'],dataRows,1), completedValues = completed.getValues();
    completed.insertCheckboxes().setValues(completedValues);
    sheet.getRange(2,headers['VI Status'],dataRows,1).setDataValidation(viRule);
  }
  styleTaskMemoColumnV1_(sheet,headers);
  if (headers['Internal Task ID']) sheet.hideColumns(headers['Internal Task ID']);
  applyTaskConditionalFormattingV9_(sheet);
}

// Activity Log의 헤더와 날짜 형식을 적용함.
function configureActivityLogSheet_() {
  const sheet = getRequiredSheet_(APP_CONFIG.SHEETS.ACTIVITY_LOG);
  sheet.getRange(1,1,1,APP_CONFIG.ACTIVITY_COLUMNS.length).setBackground('#263650').setFontColor('#FFFFFF').setFontWeight('bold');
  sheet.getRange(2,1,Math.max(sheet.getMaxRows()-1,1),1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.setFrozenRows(1);
}

// REQUESTS를 지정된 운영 순서로 재배치하고 시스템 컬럼을 숨김.
function setupRequestLayoutV5() {
  const sheet=getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  installBarcodeProvidedWorkflowV1();
  const visible=['Request ID','Node','Brand','Requestor','Request Type','자료 준비 상태','크롤링 작업자','Priority','Active Assignees','Status','Assignee Notes','Requestor Notes','Source','원본 요청 파일','Total VI','3P Completed VI','2P Completed VI','Barcode Provided','Started At','Blocked At','Resumed At','Completed At','Cancelled At','Blocked Time (Hours)','Lead Time (Hours)','Work Time (Hours)'];
  const system=['Requested At','Progress','Submission ID','Submission File Name'];
  const target=visible.concat(system);
  target.forEach(function(name,targetIndex){
    const headers=getHeaderMap_(sheet),current=headers[name]||(name==='Requestor'?headers.Requester:0);
    if(current&&current!==targetIndex+1)sheet.moveColumns(sheet.getRange(1,current,sheet.getMaxRows(),1),targetIndex+1);
  });
  configureRequestsSheet_();
  sheet.autoResizeColumns(1,Math.min(visible.length,sheet.getLastColumn()));
  const headers=getHeaderMap_(sheet);
  if(headers['Assignee Notes'])sheet.setColumnWidth(headers['Assignee Notes'],260);
  SpreadsheetApp.getActive().toast('REQUESTS 컬럼 순서를 정리했습니다.','Listing Ops',4);
}

// 기존 TASKS를 3P_TASKS로 변경하고 2P_TASKS를 준비함.
function setupSplitTaskSheetsV7() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let threeP=ss.getSheetByName('3P_TASKS');
  const legacy=ss.getSheetByName('TASKS');
  if(!threeP&&legacy){legacy.setName('3P_TASKS');threeP=legacy;}
  if(!threeP)throw new Error('기존 TASKS 또는 3P_TASKS 시트를 찾을 수 없습니다.');
  let twoP=ss.getSheetByName('2P_TASKS');
  if(!twoP)twoP=ss.insertSheet('2P_TASKS');
  configure3PTasksSheet_(); configure2PTasksSheet_();
  twoP.autoResizeColumns(1,twoP.getLastColumn());
  const twoPHeaders=getHeaderMap_(twoP);
  [['쿠팡노출상품명',220],['업체 등록 상품명',220],['등록 옵션명',190],['Memo',220]].forEach(function(entry){if(twoPHeaders[entry[0]])twoP.setColumnWidth(twoPHeaders[entry[0]],entry[1]);});
  if(twoPHeaders['Internal Task ID'])twoP.hideColumns(twoPHeaders['Internal Task ID']);
  SpreadsheetApp.getActive().toast('3P_TASKS와 2P_TASKS를 분리했습니다.','Listing Ops',5);
}


// Task의 Work Status에 따라 새 행까지 자동으로 색상을 적용함.
function applyTaskConditionalFormattingV9_(sheet) {
  const headers = getHeaderMap_(sheet);
  const rowCount = Math.max(sheet.getMaxRows() - 1, 1);
  const range = sheet.getRange(2, 1, rowCount, sheet.getLastColumn());
  const statusLetter = headers['Work Status'] ? columnLetterV4_(headers['Work Status']) : '';
  const viLetter = headers['VI Status'] ? columnLetterV4_(headers['VI Status']) : '';
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(statusLetter ? '=$' + statusLetter + '2="작업 대기"' : '=FALSE')
      .setBackground('#F4F6F9').setFontColor('#4B5563').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(statusLetter ? '=$' + statusLetter + '2="3P Listing 완료"' : '=FALSE')
      .setBackground('#E8F1FB').setFontColor('#244A73').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(statusLetter ? '=$' + statusLetter + '2="2P Lite 완료"' : '=FALSE')
      .setBackground('#E8F5F1').setFontColor('#285E52').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(statusLetter ? '=$' + statusLetter + '2="작업 완료"' : '=FALSE')
      .setBackground('#E3F3E8').setFontColor('#245B35').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(viLetter ? '=$' + viLetter + '2="INVALID"' : '=FALSE')
      .setBackground('#E7EAEE').setFontColor('#9299A3').setStrikethrough(true).setRanges([range]).build()
  ];
  sheet.setConditionalFormatRules(rules);
}

function applyTaskConditionalFormattingV8_(sheet) {
  applyTaskConditionalFormattingV9_(sheet);
}


// REQUESTS를 운영 현황표처럼 읽기 쉽게 꾸미고 신규 행까지 자동 색상화함.
function applyRequestFormattingV10_(sheet) {
  const dataRows=Math.max(sheet.getMaxRows()-1,1),lastColumn=sheet.getLastColumn(),headers=getHeaderMap_(sheet);
  const fullRange = sheet.getRange(2,1,dataRows,lastColumn);
  fullRange.setFontFamily('Roboto').setFontSize(10).setVerticalAlignment('middle')
    .setBorder(false, false, true, false, false, false, '#E7EBF0', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(1, 36);
  sheet.setRowHeights(2, dataRows, 28);
  sheet.getRange(1,1,1,lastColumn).setBackground('#24344D').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  ['Total VI','3P Completed VI','2P Completed VI','Progress','Blocked Time (Hours)','Lead Time (Hours)','Work Time (Hours)'].forEach(function(name){
    if(headers[name])sheet.getRange(2,headers[name],dataRows,1).setHorizontalAlignment('right');
  });
  ['Requested At','Started At','Blocked At','Resumed At','Completed At','Cancelled At','Barcode Provided'].forEach(function(name){
    if(headers[name])sheet.getRange(2,headers[name],dataRows,1).setHorizontalAlignment('center');
  });
  const widths={'Request ID':100,'Node':90,'Brand':110,'Requestor':120,'Request Type':205,'Priority':78,'Active Assignees':155,
    'Status':108,'Assignee Notes':240,'Total VI':82,'3P Completed VI':105,'2P Completed VI':105,'Barcode Provided':118,
    'Requested At':145,'Started At':145,'Blocked At':145,'Resumed At':145,'Completed At':145,'Cancelled At':145,
    'Blocked Time (Hours)':110,'Lead Time (Hours)':110,'Work Time (Hours)':110};
  Object.keys(widths).forEach(function(name){if(headers[name])sheet.setColumnWidth(headers[name],widths[name]);});

  const statusRange = headers['Status']?sheet.getRange(2,headers['Status'],dataRows,1):null;
  fullRange.setBackground('#FFFFFF').setFontColor('#313B49');
  const rules = statusRange?[
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('TO DO')
      .setBackground('#E7EDF7').setFontColor('#3E5F8A').setBold(true).setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ONGOING')
      .setBackground('#DCEFF7').setFontColor('#23627C').setBold(true).setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ON HOLD')
      .setBackground('#F8DDD8').setFontColor('#93453B').setBold(true).setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('DONE')
      .setBackground('#DDEFE3').setFontColor('#2F6841').setBold(true).setRanges([statusRange]).build()
  ]:[];
  sheet.setConditionalFormatRules(rules);
}


// REQUESTS의 색상과 레이아웃만 안전하게 다시 적용함.
function applyRequestVisualStyleV10() {
  const sheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  applyRequestFormattingV10_(sheet);
  refreshRequestIdLinks_();
  SpreadsheetApp.getActive().toast('REQUESTS 색상 서식을 적용했습니다.', 'Listing Ops', 4);
}
