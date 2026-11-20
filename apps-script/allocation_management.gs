const ALLOCATION_STATUS_V2 = Object.freeze({
  AUTO: '자동분배됨',
  REVIEW: '분배 필요',
  MANUAL: '수동분배됨'
});

// Task 스키마가 불완전하면 숫자 열로 추정하지 않고 쓰기를 중단함.
function requireAllocationTaskHeadersV4_(headers, names, sheetName) {
  const missing = names.filter(function(name){return !headers[name];});
  if (missing.length) throw new Error((sheetName ? sheetName + ' ' : '') + '필수 Task 열 누락: ' + missing.join(', '));
  return headers;
}

function allocationEditTouchesColumnV4_(firstColumn, lastColumn, column) {
  return !!column && firstColumn <= column && lastColumn >= column;
}

// 현재 운영 기능만 노출하는 스프레드시트 상단 메뉴임.
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('AM')
    .addItem('신규 요청 등록', 'showCombinedIntakeSidebar')
    .addItem('Kanban 열기', 'showKanbanDialog')
    .addToUi();
  ui.createMenu('작업자')
    .addItem('내 작업함 열기', 'showMyWorkSidebar')
    .addToUi();
  ui.createMenu('관리자')
    .addItem('작업자 배정', 'showPersonalKanbanDialog')
    .addItem('Kanban 열기', 'showKanbanDialog')
    .addItem('Dashboard 열기', 'showOperationsDashboard')
    .addItem('권한 관리', 'showDashboardAccessManagement')
    .addItem('선택 REQ 재계산', 'recalculateSelectedRequest')
    .addItem('재계산 대기·오류 확인', 'showRequestReconcileDiagnostics')
    .addSeparator()
    .addItem('원본 파일 저장 위치 설정', 'configureOriginalRequestFolderV2')
    .addSeparator()
    .addItem('백업 아카이브 설정', 'configureArchiveSpreadsheetV1')
    .addItem('백업 대상 미리보기', 'previewArchiveCandidatesV1')
    .addToUi();
}

function onInstall() {
  onOpen();
}

// 운영 시트의 열 순서와 표시 형식을 실무 흐름에 맞게 정리함.
function polishOperationsSheetsV3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function(sheet) {
    sheet.showSheet();
    if (sheet.getMaxColumns() > 0) sheet.showColumns(1, sheet.getMaxColumns());
  });
  installBarcodeProvidedWorkflowV1();
  const requests = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  ensureAllocationColumnsV2_();
  ensureValidityManagementV3_();
  cleanupActiveAssigneePlaceholdersV3_();
  const requestHeadersForFreeze = getHeaderMap_(requests);
  if (requestHeadersForFreeze['Active Assignees']) requests.setFrozenColumns(requestHeadersForFreeze['Active Assignees']);
  const threePForOrder = getRequiredSheet_(APP_CONFIG.SHEETS.TASKS);
  let threePHeadersForOrder = getHeaderMap_(threePForOrder);
  if (threePHeadersForOrder['2P Completed'] && threePHeadersForOrder['VI Status'] !== threePHeadersForOrder['2P Completed'] + 1) {
    threePForOrder.moveColumns(threePForOrder.getRange(1,threePHeadersForOrder['VI Status'],threePForOrder.getMaxRows(),1),threePHeadersForOrder['2P Completed']+1);
    threePHeadersForOrder = getHeaderMap_(threePForOrder);
  }
  if (threePHeadersForOrder['2P Completed'] && threePHeadersForOrder['Invalid Reason'] !== threePHeadersForOrder['2P Completed'] + 2) {
    threePForOrder.moveColumns(threePForOrder.getRange(1,threePHeadersForOrder['Invalid Reason'],threePForOrder.getMaxRows(),1),threePHeadersForOrder['2P Completed']+2);
  }
  const twoPForOrder = getRequiredSheet_(APP_CONFIG.SHEETS.TASKS_2P);
  let twoPHeadersForOrder = getHeaderMap_(twoPForOrder);
  if (twoPHeadersForOrder['2P Completed'] && twoPHeadersForOrder['VI Status'] !== twoPHeadersForOrder['2P Completed'] + 1) {
    twoPForOrder.moveColumns(twoPForOrder.getRange(1,twoPHeadersForOrder['VI Status'],twoPForOrder.getMaxRows(),1),twoPHeadersForOrder['2P Completed']+1);
    twoPHeadersForOrder = getHeaderMap_(twoPForOrder);
  }
  if (twoPHeadersForOrder['2P Completed'] && twoPHeadersForOrder['Invalid Reason'] !== twoPHeadersForOrder['2P Completed'] + 2) {
    twoPForOrder.moveColumns(twoPForOrder.getRange(1,twoPHeadersForOrder['Invalid Reason'],twoPForOrder.getMaxRows(),1),twoPHeadersForOrder['2P Completed']+2);
  }
  [APP_CONFIG.SHEETS.TASKS,APP_CONFIG.SHEETS.CRAWL_TASKS,APP_CONFIG.SHEETS.TASKS_2P].forEach(function(sheetName) {
    const taskSheet = getRequiredSheet_(sheetName), taskHeaders = getHeaderMap_(taskSheet);
    const idColumn = taskHeaders['Internal Task ID'];
    if (idColumn && idColumn !== taskSheet.getLastColumn()) {
      taskSheet.moveColumns(taskSheet.getRange(1,idColumn,taskSheet.getMaxRows(),1),taskSheet.getLastColumn()+1);
    }
  });
  if (typeof ensureRequestIntakeMetadataColumnsV2_ === 'function') ensureRequestIntakeMetadataColumnsV2_();

  const requestOrder = [
    'Request ID','Priority','Status','Assignee Notes','Allocation Status','Active Assignees',
    'Participants','Requestor','Brand','Node','A코드','Request Type','Source','Requestor Notes','원본 요청 파일',
    'Total VI','3P Completed VI','2P Completed VI','Progress','Barcode Provided',
    'Archive Status','Archived At','Archive Batch ID','Archived Task Rows','Archive File ID',
    'Requested At','Started At','Completed At',
    'Blocked At','Resumed At','Cancelled At','Blocked Time (Hours)','Lead Time (Hours)','Work Time (Hours)'
  ];
  requestOrder.forEach(function(name, targetIndex) {
    const headers = getHeaderMap_(requests);
    const current = headers[name];
    if (current && current !== targetIndex + 1) {
      requests.moveColumns(requests.getRange(1, current, requests.getMaxRows(), 1), targetIndex + 1);
    }
  });
  const requestHeadersAfterOrder = getHeaderMap_(requests);

  styleOperationsTableV3_(requests, {
    frozenColumns: requestHeadersAfterOrder['Active Assignees'] || 5,
    widths: {
      'Request ID':100,'Priority':78,'Status':105,'Assignee Notes':190,'Cancelled At':135,'Allocation Status':105,'Active Assignees':150,
      'Participants':140,'Node':125,'Brand':115,'Request Type':185,'Requestor':95,'A코드':100,
      'Source':180,'Total VI':76,'3P Completed VI':92,'2P Completed VI':92,'Progress':80,
      'Barcode Provided':118,'Requestor Notes':190,'원본 요청 파일':150
      ,'Archive Status':115,'Archived At':145,'Archive Batch ID':185,'Archived Task Rows':115,'Archive File ID':190
    }
  });
  styleOperationsTableV3_(getRequiredSheet_(APP_CONFIG.SHEETS.TASKS), {
    frozenColumns: 5,
    widths: {'Request ID':100,'품번':120,'등록상품명':210,'사이즈':75,'색상':95,'판매가':85,'정상가':85,
      'Brand':105,'Memo':220,'3P Assignee':105,'3P Completed':90,'2P Assignee':105,'2P Completed':90,
      '3P Completed At':125,'2P Completed At':125,'Work Status':115,'VI Status':80,'Invalid Reason':190},
    hidden: []
  });
  styleOperationsTableV3_(getRequiredSheet_(APP_CONFIG.SHEETS.CRAWL_TASKS), {
    frozenColumns: 5,
    widths: {'Request ID':100,'품번':120,'등록상품명':210,'사이즈':75,'색상':95,'판매가':85,'정상가':85,
      'Brand':105,'Memo':220,'3P Assignee':105,'3P Completed':90,'2P Assignee':105,'2P Completed':90,
      '3P Completed At':125,'2P Completed At':125,'Work Status':115,'VI Status':80,'Invalid Reason':190},
    hidden: []
  });
  styleOperationsTableV3_(getRequiredSheet_(APP_CONFIG.SHEETS.TASKS_2P), {
    frozenColumns: 3,
    widths: {'Request ID':100,'업체상품ID':115,'Product ID':105,'3P 옵션 ID':105,'바코드':120,
      '업체상품코드':120,'쿠팡노출상품명':200,'업체 등록 상품명':200,'등록 옵션명':170,
      'Memo':220,'2P Assignee':105,'2P Completed':90,'2P Completed At':125,'Work Status':115,'VI Status':80,'Invalid Reason':190},
    hidden: []
  });
  ['ACTIVITY_LOG','CONFIG'].forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    if (sheet) styleOperationsTableV3_(sheet, {frozenColumns:1,widths:{},hidden:[]});
  });
  optimizeOperationsUiV4_();
  ss.toast('시트 열 순서와 UI 정리가 완료되었습니다.','관리자',5);
}

// 운영 시트 전반의 시각 체계를 가볍고 일관되게 정리함.
function optimizeOperationsUiV4_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function(sheet){sheet.getBandings().forEach(function(banding){banding.remove();});});
  const themes = {};
  themes[APP_CONFIG.SHEETS.REQUESTS] = {tab:'#24364B',freeze:'Active Assignees',wrap:['Assignee Notes','Requestor Notes','Source']};
  themes[APP_CONFIG.SHEETS.TASKS] = {tab:'#5475B8',freeze:'색상',wrap:['등록상품명','Memo','Invalid Reason']};
  themes[APP_CONFIG.SHEETS.CRAWL_TASKS] = {tab:'#6B4F8A',freeze:'색상',wrap:['등록상품명','Memo','Invalid Reason']};
  themes[APP_CONFIG.SHEETS.TASKS_2P] = {tab:'#5BB5E5',freeze:'Product ID',wrap:['쿠팡노출상품명','업체 등록 상품명','등록 옵션명','Memo','Invalid Reason']};
  themes[APP_CONFIG.SHEETS.ACTIVITY_LOG] = {tab:'#7B8798',freeze:null,wrap:[]};
  themes[APP_CONFIG.SHEETS.CONFIG] = {tab:'#D69E3D',freeze:null,wrap:[]};

  Object.keys(themes).forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName); if (!sheet) return;
    const theme = themes[sheetName], h = getHeaderMap_(sheet), lastColumn = sheet.getLastColumn();
    const maxRows = sheet.getMaxRows();
    const lastDataRow = Math.max(sheet.getLastRow(),1);

    sheet.setTabColor(theme.tab).setHiddenGridlines(true).setFrozenRows(1);
    if (theme.freeze && h[theme.freeze]) sheet.setFrozenColumns(h[theme.freeze]);
    sheet.getRange(1,1,1,lastColumn).setBackground('#24364B').setFontColor('#FFFFFF')
      .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    const manualHeaders = sheetName === APP_CONFIG.SHEETS.TASKS
      ? ['3P Assignee','3P Completed','2P Assignee','2P Completed','VI Status','Invalid Reason']
      : (sheetName === APP_CONFIG.SHEETS.TASKS_2P
        ? ['2P Assignee','2P Completed','VI Status','Invalid Reason'] : []);
    manualHeaders.forEach(function(name){
      if(h[name])sheet.getRange(1,h[name]).setBackground('#C9822B').setFontColor('#FFFFFF').setFontWeight('bold');
    });
    sheet.setRowHeight(1,38);

    if (lastDataRow >= 2) {
      const body = sheet.getRange(2,1,lastDataRow-1,lastColumn);
      body.setFontFamily('Arial').setFontSize(10).setFontColor('#27364A')
        .setVerticalAlignment('middle').setBackground('#FFFFFF');
      sheet.setRowHeights(2,lastDataRow-1,27);
      theme.wrap.forEach(function(name){if(h[name])sheet.getRange(2,h[name],lastDataRow-1,1).setWrap(true).setVerticalAlignment('middle');});
      ['Priority','Status','Allocation Status','Request Type','Barcode Provided','VI Status','Work Status'].forEach(function(name){
        if(h[name])sheet.getRange(2,h[name],lastDataRow-1,1).setHorizontalAlignment('center').setFontWeight('bold');
      });
      ['3P Completed','2P Completed'].forEach(function(name){if(h[name])sheet.getRange(2,h[name],lastDataRow-1,1).setHorizontalAlignment('center');});
      ['판매가','정상가','판매가격','할인율기준가격','Total VI','3P Completed VI','2P Completed VI'].forEach(function(name){
        if(h[name])sheet.getRange(2,h[name],lastDataRow-1,1).setHorizontalAlignment('right').setNumberFormat('#,##0');
      });
      ['Requested At','Started At','Completed At','Blocked At','Resumed At','Cancelled At','3P Completed At','2P Completed At'].forEach(function(name){
        if(h[name])sheet.getRange(2,h[name],lastDataRow-1,1).setHorizontalAlignment('center').setNumberFormat('yyyy-mm-dd hh:mm');
      });
      if(h['Progress'])sheet.getRange(2,h['Progress'],lastDataRow-1,1).setHorizontalAlignment('right').setNumberFormat('0%');
    }
    if (lastDataRow < maxRows) {
      sheet.getRange(lastDataRow+1,1,maxRows-lastDataRow,lastColumn)
        .setBackground('#FFFFFF').setFontFamily('Arial').setFontSize(10).setFontColor('#27364A');
      sheet.setRowHeights(lastDataRow+1,maxRows-lastDataRow,24);
    }
  });

  const widths = {
    'Assignee Notes':220,'Requestor Notes':220,'Memo':220,'Invalid Reason':200,'VI Status':82,
    'Active Assignees':155,'Participants':155,'Request Type':180,'Source':175,
    '원본 요청 파일':150,'Internal Task ID':145,'3P Completed At':132,'2P Completed At':132,
    'Cancelled At':132,'등록상품명':210,'쿠팡노출상품명':205,'업체 등록 상품명':205,'등록 옵션명':175,
    'Priority':86,'Status':118,'Allocation Status':116,'3P Completed':100,'2P Completed':100
  };
  [APP_CONFIG.SHEETS.REQUESTS,APP_CONFIG.SHEETS.TASKS,APP_CONFIG.SHEETS.CRAWL_TASKS,APP_CONFIG.SHEETS.TASKS_2P].forEach(function(name){
    const sheet=ss.getSheetByName(name);if(!sheet)return;const h=getHeaderMap_(sheet);
    Object.keys(widths).forEach(function(header){if(h[header])sheet.setColumnWidth(h[header],widths[header]);});
  });
  applySemanticStatusColorsV4_();
}

function columnLetterV4_(column) {
  let value = Number(column), result = '';
  while (value > 0) { value -= 1; result = String.fromCharCode(65 + value % 26) + result; value = Math.floor(value / 26); }
  return result;
}

// 줄무늬 대신 실제 의미가 있는 상태값에만 색상을 적용함.
function applySemanticStatusColorsV4_() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  function rule_(range,formula,bg,fg,bold,strike) {
    let builder=SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula).setBackground(bg).setFontColor(fg);
    if(bold)builder=builder.setBold(true);
    if(strike)builder=builder.setStrikethrough(true);
    return builder.setRanges([range]).build();
  }
  function cellRules_(rules,sheet,h,lastRow,name,styles) {
    if(!h[name]||lastRow<2)return;
    const range=sheet.getRange(2,h[name],lastRow-1,1),letter=columnLetterV4_(h[name]);
    Object.keys(styles).forEach(function(value){const style=styles[value];rules.push(rule_(range,'=$'+letter+'2="'+value+'"',style[0],style[1],true,false));});
  }
  [APP_CONFIG.SHEETS.REQUESTS,APP_CONFIG.SHEETS.TASKS,APP_CONFIG.SHEETS.CRAWL_TASKS,APP_CONFIG.SHEETS.TASKS_2P].forEach(function(name){
    const sheet=ss.getSheetByName(name);if(!sheet)return;const h=getHeaderMap_(sheet),maxRows=sheet.getMaxRows();
    const lastRow=maxRows;
    const rules=[];
    if(lastRow>=2){
      const full=sheet.getRange(2,1,lastRow-1,sheet.getLastColumn());
      if(name===APP_CONFIG.SHEETS.REQUESTS&&h['Status']){
        const statusLetter=columnLetterV4_(h['Status']);
        // DROP 행은 모든 셀 색보다 우선하여 완전히 비활성화된 모습으로 표시함.
        rules.push(rule_(full,'=$'+statusLetter+'2="DROP"','#E7EAEE','#9299A3',false,true));
        cellRules_(rules,sheet,h,lastRow,'Priority',{
          HIGH:['#F9DAD6','#A93731'],NORMAL:['#E7EBF0','#505B6A'],LOW:['#DDEEE2','#326844']
        });
        cellRules_(rules,sheet,h,lastRow,'Status',{
          'TO DO':['#E3E8EE','#4F5F72'],ONGOING:['#D8EAFE','#1F5B91'],'ON HOLD':['#FFE8B8','#815600'],DONE:['#D8EDDE','#27683D'],DROP:['#D9DDE2','#737A84']
        });
        cellRules_(rules,sheet,h,lastRow,'Allocation Status',{
          '자동분배됨':['#D8EDDE','#27683D'],'분배 필요':['#FFE8B8','#815600'],'수동분배됨':['#DDE5F5','#3D5C91']
        });
        cellRules_(rules,sheet,h,lastRow,'Barcode Provided',{
          YES:['#D8EDDE','#27683D'],NO:['#F9DAD6','#A93731']
        });
        rules.push(rule_(full,'=$'+statusLetter+'2="TO DO"','#FFFFFF','#27364A',false,false));
        rules.push(rule_(full,'=$'+statusLetter+'2="ONGOING"','#EAF3FF','#244F78',false,false));
        rules.push(rule_(full,'=$'+statusLetter+'2="ON HOLD"','#FFF3D9','#6F5520',false,false));
        rules.push(rule_(full,'=$'+statusLetter+'2="DONE"','#E8F5EC','#356347',false,false));
      } else {
        const viLetter=h['VI Status']?columnLetterV4_(h['VI Status']):'';
        if(viLetter)rules.push(rule_(full,'=$'+viLetter+'2="INVALID"','#E7EAEE','#9299A3',false,true));
        const completed2=h['2P Completed']?columnLetterV4_(h['2P Completed']):'';
        const completed3=h['3P Completed']?columnLetterV4_(h['3P Completed']):'';
        if(completed2)rules.push(rule_(full,'=$'+completed2+'2=TRUE','#E5F3E9','#2E6541',false,false));
        if(completed3)rules.push(rule_(full,'=AND($'+completed3+'2=TRUE,$'+completed2+'2<>TRUE)','#E7F1FC','#285A86',false,false));
        if(completed3&&completed2)rules.push(rule_(full,'=AND($'+completed3+'2<>TRUE,$'+completed2+'2<>TRUE)','#FFFFFF','#27364A',false,false));
        else if(completed2)rules.push(rule_(full,'=$'+completed2+'2<>TRUE','#FFFFFF','#27364A',false,false));
      }
    }
    sheet.setConditionalFormatRules(rules);
  });
}

function styleOperationsTableV3_(sheet, options) {
  const lastColumn = sheet.getLastColumn();
  if (!lastColumn) return;
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const headers = getHeaderMap_(sheet);
  sheet.showSheet();
  sheet.showColumns(1, sheet.getMaxColumns());
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(Math.min(options.frozenColumns || 1, lastColumn));
  sheet.setHiddenGridlines(true);
  sheet.getRange(1,1,1,lastColumn)
    .setBackground('#24364B').setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle')
    .setHorizontalAlignment('center').setWrap(true);
  sheet.setRowHeight(1,34);
  if (lastRow > 1) {
    sheet.getRange(2,1,lastRow-1,lastColumn).setFontFamily('Arial').setFontSize(10)
      .setVerticalAlignment('middle');
    sheet.setRowHeights(2,lastRow-1,28);
  }
  Object.keys(options.widths || {}).forEach(function(name) {
    if (headers[name]) sheet.setColumnWidth(headers[name], options.widths[name]);
  });
  ['Priority','Status','Allocation Status','Work Status','Barcode Provided'].forEach(function(name) {
    if (!headers[name] || lastRow < 2) return;
    const range = sheet.getRange(2,headers[name],lastRow-1,1);
    range.setHorizontalAlignment('center').setFontWeight('bold');
  });
  ['Total VI','3P Completed VI','2P Completed VI','Progress','판매가','정상가','판매가격','할인율기준가격'].forEach(function(name) {
    if (headers[name] && lastRow >= 2) sheet.getRange(2,headers[name],lastRow-1,1).setHorizontalAlignment('right');
  });
  if (headers['Progress'] && lastRow >= 2) sheet.getRange(2,headers['Progress'],lastRow-1,1).setNumberFormat('0%');
  ['판매가','정상가','판매가격','할인율기준가격'].forEach(function(name) {
    if (headers[name] && lastRow >= 2) sheet.getRange(2,headers[name],lastRow-1,1).setNumberFormat('#,##0');
  });
  (options.hidden || []).forEach(function(name) {
    if (headers[name]) sheet.hideColumns(headers[name]);
  });
}

function ensureAllocationColumnsV2_() {
  const sheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  let headers = getHeaderMap_(sheet);
  ['Participants','Active Assignees','Allocation Status'].forEach(function(name) {
    if (headers[name]) return;
    const column = sheet.getLastColumn() + 1;
    sheet.getRange(1, column).setValue(name).setBackground('#263650')
      .setFontColor('#FFFFFF').setFontWeight('bold');
    headers[name] = column;
  });
  return headers;
}

// 요청 DROP과 VI 유효성 관리에 필요한 열·드롭다운을 준비함.
function ensureValidityManagementV3_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requestSheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  let rh = getHeaderMap_(requestSheet);
  if (rh['Cancellation Reason'] && !rh['Assignee Notes']) {
    requestSheet.getRange(1,rh['Cancellation Reason']).setValue('Assignee Notes');
    rh = getHeaderMap_(requestSheet);
  }
  if (rh['Notes'] && !rh['Requestor Notes']) {
    requestSheet.getRange(1,rh['Notes']).setValue('Requestor Notes');
    rh = getHeaderMap_(requestSheet);
  }
  if (rh['Cancelled By']) {
    requestSheet.deleteColumn(rh['Cancelled By']);
    rh = getHeaderMap_(requestSheet);
  }
  ['Assignee Notes','Cancelled At'].forEach(function(name) {
    if (rh[name]) return;
    const column = requestSheet.getLastColumn() + 1;
    requestSheet.getRange(1,column).setValue(name).setBackground('#24364B').setFontColor('#FFFFFF').setFontWeight('bold');
    rh = getHeaderMap_(requestSheet);
  });
  const requestStatusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(manualRequestStatuses_(), true).setAllowInvalid(false).build();
  requestSheet.getRange(2,rh['Status'],Math.max(requestSheet.getMaxRows()-1,1),1).setDataValidation(requestStatusRule);

  [APP_CONFIG.SHEETS.TASKS,APP_CONFIG.SHEETS.CRAWL_TASKS,APP_CONFIG.SHEETS.TASKS_2P].forEach(function(sheetName) {
    const sheet = getRequiredSheet_(sheetName);
    let h = getHeaderMap_(sheet);
    ['VI Status','Invalid Reason'].forEach(function(name) {
      if (h[name]) return;
      const column = sheet.getLastColumn() + 1;
      sheet.getRange(1,column).setValue(name).setBackground('#24364B').setFontColor('#FFFFFF').setFontWeight('bold');
      h = getHeaderMap_(sheet);
    });
    const maxDataRows = Math.max(sheet.getMaxRows()-1,1);
    const lastDataRow = Math.max(sheet.getLastRow(),1);
    sheet.getRange(2,h['VI Status'],maxDataRows,1).clearDataValidations();
    if (lastDataRow >= 2) sheet.getRange(2,h['VI Status'],lastDataRow-1,1)
      .setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(['VALID','INVALID'],true).setAllowInvalid(false).build());
    if (lastDataRow >= 2) {
      const current = sheet.getRange(2,h['VI Status'],lastDataRow-1,1).getValues();
      let changed = false;
      current.forEach(function(row){if(!String(row[0]||'').trim()){row[0]='VALID';changed=true;}});
      if (changed) sheet.getRange(2,h['VI Status'],current.length,1).setValues(current);
    }
  });
  return true;
}

function isViValidV3_(row,headers){return !headers['VI Status']||String(row[headers['VI Status']-1]||'VALID').trim().toUpperCase()!=='INVALID';}

// 유효한 VI만 REQUESTS 진행률 분모와 완료 수에 반영함.
function recalculateValidViMetricsV3_(requestId){return withOpsDocumentLock_(function(){return syncRequestSummaries_([requestId]);});}

// TextFinder 결과를 연속 행 단위로 묶어 DROP/복구의 Spreadsheet 왕복을 줄임.
function groupTaskMatchRowsV3_(matches) {
  const rows=matches.map(function(cell){return cell.getRow();}).sort(function(a,b){return a-b;}),groups=[];
  rows.forEach(function(row){const group=groups[groups.length-1];if(group&&group.start+group.count===row)group.count++;else groups.push({start:row,count:1});});
  return groups;
}

function writeTaskValidityBlockV3_(sheet,headers,group,statuses,reasons) {
  const statusCol=headers['VI Status'],reasonCol=headers['Invalid Reason'];
  if(Math.abs(statusCol-reasonCol)===1){
    const startCol=Math.min(statusCol,reasonCol),statusOffset=statusCol-startCol,reasonOffset=reasonCol-startCol;
    sheet.getRange(group.start,startCol,group.count,2).setValues(statuses.map(function(status,index){const row=['',''];row[statusOffset]=status;row[reasonOffset]=reasons[index];return row;}));
    return;
  }
  sheet.getRange(group.start,statusCol,group.count,1).setValues(statuses.map(function(value){return [value];}));
  sheet.getRange(group.start,reasonCol,group.count,1).setValues(reasons.map(function(value){return [value];}));
}

function cancelledInvalidReasonV4_(requestReason,previousStatus,previousReason) {
  const invalidReason=REQUEST_DROP_REASON_PREFIX_V2_+(String(requestReason||'').trim()?' · '+String(requestReason).trim():'');
  const reason=String(previousReason||'').trim(),marker=' || PREVIOUS INVALID: ';
  if(String(previousStatus||'').trim().toUpperCase()!=='INVALID'||!reason)return invalidReason;
  if(isRequestDropReasonV2_(reason)){
    const markerIndex=reason.indexOf(marker);
    return markerIndex>=0?invalidReason+marker+reason.slice(markerIndex+marker.length):invalidReason;
  }
  return invalidReason+marker+reason;
}

// 한 요청의 실제 원본 Task만 읽어 DROP/해제 왕복을 전체 시트 크기와 분리함.
function requestTaskValidityContextV5_(requestId){
  const id=String(requestId),requestTable=readOpsTable_(APP_CONFIG.SHEETS.REQUESTS),records=requestTable.byId[id]||[];
  if(records.length!==1)throw Error('요청 ID가 없거나 중복되었습니다: '+id);
  const sourceName=opsSourceForRequest_(records[0].values,requestTable.headers),executionContext={};
  const table=readTasksForRequests_(sourceName,[id],executionContext);
  return {id:id,requestTable:requestTable,requestRecord:records[0],sourceName:sourceName,executionContext:executionContext,table:table,taskRecords:table.byId[id]||[]};
}

function writeScopedTaskValidityV5_(table,records){
  markRequestReconcileNeeded_(records.map(function(r){return String(r.values[table.headers['Request ID']-1]);}),{tasks:true},'validity');
  const h=table.headers,sorted=records.slice().sort(function(a,b){return a.row-b.row;}),groups=[];
  sorted.forEach(function(record){const group=groups[groups.length-1];if(group&&group.start+group.records.length===record.row)group.records.push(record);else groups.push({start:record.row,records:[record]});});
  groups.forEach(function(group){
    writeTaskValidityBlockV3_(table.sheet,h,{start:group.start,count:group.records.length},
      group.records.map(function(record){return record.values[h['VI Status']-1];}),
      group.records.map(function(record){return record.values[h['Invalid Reason']-1];}));
  });
}

// 요청 DROP 시 연결된 모든 VI를 담당자 이력을 보존한 채 무효화함.
function invalidateRequestTasksV3_(requestId, requestReason, providedContext) {
  const context=providedContext||requestTaskValidityContextV5_(requestId),table=context.table,h=table.headers,changed=[];
  if(!h['VI Status']||!h['Invalid Reason'])throw Error('Task 유효성 열을 찾을 수 없습니다: '+context.sourceName);
  context.taskRecords.forEach(function(record){
    const row=record.values,previousStatus=String(row[h['VI Status']-1]||'VALID').trim().toUpperCase(),previousReason=String(row[h['Invalid Reason']-1]||'').trim();
    const nextReason=cancelledInvalidReasonV4_(requestReason,previousStatus,previousReason);
    if(previousStatus==='INVALID'&&previousReason===nextReason)return;
    row[h['VI Status']-1]='INVALID';row[h['Invalid Reason']-1]=nextReason;changed.push(record);
  });
  if(changed.length)writeScopedTaskValidityV5_(table,changed);
  rebuildScopedTaskRows_(table);
  syncRequestSummaries_([context.id],{requestTable:context.requestTable,executionContext:context.executionContext});
  return {tasks:context.taskRecords.length,changed:changed.length,source:context.sourceName};
}

// 설치형 트리거 전환 전에 일부 행만 무효화된 DROP 요청을 일괄 보정함.
// 모든 Task 원본을 한 번씩 읽고 변경 열만 연속 범위로 묶어 저장함.
function repairCancelledRequestValidityV4_() {
  const requestTable=readOpsTable_(APP_CONFIG.SHEETS.REQUESTS),states=Object.create(null),cancelled=Object.create(null),requestIds=[];
  Object.keys(requestTable.byId).forEach(function(requestId){
    const records=requestTable.byId[requestId];
    if(records.length!==1)return;
    const isCancelled=normalizeRequestStatusV2_(opsValue_(records[0].values,requestTable.headers,'Status'))===APP_CONFIG.STATUS.DROP;
    states[requestId]={cancelled:isCancelled,record:records[0]};
    if(isCancelled){
      cancelled[requestId]=String(opsValue_(records[0].values,requestTable.headers,'Assignee Notes')||'').trim();
      requestIds.push(requestId);
    }
  });
  const names=[APP_CONFIG.SHEETS.TASKS,APP_CONFIG.SHEETS.CRAWL_TASKS,APP_CONFIG.SHEETS.TASKS_2P];
  const sources=readTaskSources_(names),reactivated=Object.create(null);let matchedRows=0,repairedRows=0,reactivatedRows=0;
  names.forEach(function(name){
    const table=sources[name],h=table.headers,statusCol=h['VI Status'],reasonCol=h['Invalid Reason'];
    if(!statusCol||!reasonCol)return;
    const patches=[];
    table.rows.forEach(function(row,index){
      const requestId=String(opsValue_(row,h,'Request ID')||''),state=states[requestId];
      if(!state)return;
      const previousStatus=String(row[statusCol-1]||'VALID').trim().toUpperCase();
      const previousReason=String(row[reasonCol-1]||'').trim();
      let nextStatus=previousStatus,nextReason=previousReason;
      if(state.cancelled){
        matchedRows++;nextStatus='INVALID';nextReason=cancelledInvalidReasonV4_(cancelled[requestId],previousStatus,previousReason);
      }else if(isRequestDropReasonV2_(previousReason)){
        const marker=' || PREVIOUS INVALID: ',markerIndex=previousReason.indexOf(marker);
        if(markerIndex>=0){nextStatus='INVALID';nextReason=previousReason.slice(markerIndex+marker.length);}
        else{nextStatus='VALID';nextReason='';}
        reactivated[requestId]=true;reactivatedRows++;
      }else return;
      if(previousStatus!==nextStatus){
        patches.push({row:index+2,col:statusCol,value:nextStatus});
        row[statusCol-1]=nextStatus;
      }
      if(previousReason!==nextReason){
        patches.push({row:index+2,col:reasonCol,value:nextReason});
        row[reasonCol-1]=nextReason;
      }
      if(previousStatus!==nextStatus||previousReason!==nextReason)repairedRows++;
    });
    if(patches.length)writeOpsPatches_(table.sheet,patches);
  });
  const requestPatches=[],rh=requestTable.headers;
  Object.keys(reactivated).forEach(function(requestId){
    const record=states[requestId].record;
    addOpsPatch_(requestPatches,record,rh,'Cancelled At','');
  });
  if(requestPatches.length)writeOpsPatches_(requestTable.sheet,requestPatches);
  const syncIds=requestIds.concat(Object.keys(reactivated));
  if(syncIds.length)syncRequestSummaries_(syncIds,{requestTable:requestTable,taskSources:sources});
  bumpMyWorkSidebarRevisionV2_([]);
  return {requests:syncIds.length,cancelledRequests:requestIds.length,reactivatedRequests:Object.keys(reactivated).length,matchedRows:matchedRows,repairedRows:repairedRows,reactivatedRows:reactivatedRows};
}

function repairCancelledRequestValidityV4(){
  return withOpsDocumentLock_(repairCancelledRequestValidityV4_);
}

// 과거에 Active Assignees 공백만으로 분배 필요가 고정된 요청을 Task 기준으로 보정함.
// 전체 Task 본문 대신 후보 Request ID와 일치하는 행만 읽음.
function repairTaskDerivedAssignmentSummariesV6_(){
  const requestTable=readOpsTable_(APP_CONFIG.SHEETS.REQUESTS),h=requestTable.headers,candidates=[],before=Object.create(null),patches=[];
  Object.keys(requestTable.byId).forEach(function(requestId){
    const records=requestTable.byId[requestId];if(records.length!==1)return;
    const row=records[0].values,status=normalizeRequestStatusV2_(opsValue_(row,h,'Status'));
    const active=String(opsValue_(row,h,'Active Assignees')||''),allocation=String(opsValue_(row,h,'Allocation Status')||''),cancelledAt=opsValue_(row,h,'Cancelled At');
    const needsActive=status!==APP_CONFIG.STATUS.DROP&&status!==APP_CONFIG.STATUS.DONE&&!active;
    const needsAllocation=allocation===ALLOCATION_STATUS_V2.REVIEW;
    const needsCancelledAtClear=status!==APP_CONFIG.STATUS.DROP&&!!cancelledAt;
    if(!needsActive&&!needsAllocation&&!needsCancelledAtClear)return;
    candidates.push(requestId);before[requestId]={active:active,allocation:allocation,cancelledAt:cancelledAt};
    if(needsCancelledAtClear)addOpsPatch_(patches,records[0],h,'Cancelled At','');
  });
  if(!candidates.length)return {requests:0,activeRestored:0,allocationRestored:0};
  if(patches.length)writeOpsPatches_(requestTable.sheet,patches);
  const synced=syncRequestSummaries_(candidates,{requestTable:requestTable,executionContext:{}});
  let activeRestored=0,allocationRestored=0,cancelledAtCleared=0;
  candidates.forEach(function(requestId){
    const summary=synced.summaries[requestId];
    if(!before[requestId].active&&summary['Active Assignees'])activeRestored++;
    if(before[requestId].allocation===ALLOCATION_STATUS_V2.REVIEW&&summary['Allocation Status']!==ALLOCATION_STATUS_V2.REVIEW)allocationRestored++;
    if(before[requestId].cancelledAt&&normalizeRequestStatusV2_(opsValue_(requestTable.byId[requestId][0].values,h,'Status'))!==APP_CONFIG.STATUS.DROP)cancelledAtCleared++;
  });
  return {requests:candidates.length,activeRestored:activeRestored,allocationRestored:allocationRestored,cancelledAtCleared:cancelledAtCleared};
}

function repairTaskDerivedAssignmentSummariesV6(){
  return withOpsDocumentLock_(repairTaskDerivedAssignmentSummariesV6_);
}

// DROP을 되돌리면 DROP 때문에 무효화된 VI와 기존 활성 작업자를 복구함.
function reactivateCancelledRequestTasksV3_(requestId) {
  const context=requestTaskValidityContextV5_(requestId),table=context.table,h=table.headers,changed=[];
  if(!h['VI Status']||!h['Invalid Reason'])throw Error('Task 유효성 열을 찾을 수 없습니다: '+context.sourceName);
  context.taskRecords.forEach(function(record){
    const row=record.values;let status=String(row[h['VI Status']-1]||'VALID').trim().toUpperCase(),reason=String(row[h['Invalid Reason']-1]||'').trim();
    if(isRequestDropReasonV2_(reason)){
      const marker=' || PREVIOUS INVALID: ',markerIndex=reason.indexOf(marker);
      if(markerIndex>=0){status='INVALID';reason=reason.slice(markerIndex+marker.length);}
      else{status='VALID';reason='';}
      row[h['VI Status']-1]=status;row[h['Invalid Reason']-1]=reason;changed.push(record);
    }
  });
  if(changed.length)writeScopedTaskValidityV5_(table,changed);
  rebuildScopedTaskRows_(table);
  const requestSheet=context.requestTable.sheet,rh=context.requestTable.headers,requestRow=context.requestRecord.row;
  if(rh['Cancelled At']){requestSheet.getRange(requestRow,rh['Cancelled At']).clearContent();context.requestRecord.values[rh['Cancelled At']-1]='';}
  const synced=syncRequestSummaries_([context.id],{requestTable:context.requestTable,executionContext:context.executionContext});
  const workers=splitWorkersV2_(synced.summaries[context.id]['Active Assignees']);
  bumpMyWorkSidebarRevisionV2_(workers);
  return {tasks:context.taskRecords.length,changed:changed.length,workers:workers.length,source:context.sourceName};
}

function splitWorkersV2_(value) {
  return String(value || '').split(',').map(function(v){return v.trim();})
    .filter(function(v,i,a){return v && String(v).toUpperCase() !== 'N/A' && a.indexOf(v) === i;});
}

function cleanupActiveAssigneePlaceholdersV3_() {
  const sheet=getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS),h=getHeaderMap_(sheet);
  if(!h['Active Assignees']||sheet.getLastRow()<2)return;
  const range=sheet.getRange(2,h['Active Assignees'],sheet.getLastRow()-1,1);
  const values=range.getDisplayValues().map(function(row){return [splitWorkersV2_(row[0]).join(', ')];});
  range.setValues(values);
}

function syncParticipantsV2_(requestId) {
  return withOpsDocumentLock_(function(){const result=syncRequestSummaries_([requestId],{participantsOnly:true});return splitWorkersV2_(result.summaries[String(requestId)].Participants);});
}

function addParticipantV2_(requestId, worker) {
  worker = String(worker || '').trim();
  if (!worker) return;
  const requestSheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS), h = ensureAllocationColumnsV2_();
  const row = findRowByRequestId_(requestSheet, requestId);
  const names = splitWorkersV2_(requestSheet.getRange(row,h['Participants']).getValue());
  if (names.indexOf(worker) >= 0) return;
  names.push(worker);
  requestSheet.getRange(row,h['Participants']).setValue(names.join(', '));
}

// 기존 완료 이력을 기준으로 모든 요청의 Participants를 한 번에 재구성함.
function syncAllParticipantsV2() {
  return withOpsDocumentLock_(function(){return syncRequestSummaries_(null,{participantsOnly:true}).requests;});
}

// Explicit maintenance: removes obsolete sheets only when manually invoked.

function cleanupObsoleteOperationsV2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const obsoleteSheetKeys = ['taskinput','myrequests','requestform','운영가이드'];
  const deletedSheets = [];
  ss.getSheets().slice().forEach(function(sheet) {
    const key = String(sheet.getName() || '').toLowerCase().replace(/[\s_-]+/g,'');
    if (obsoleteSheetKeys.indexOf(key) < 0) return;
    deletedSheets.push(sheet.getName());
    ss.deleteSheet(sheet);
  });

  const requestSheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  const headers = getHeaderMap_(requestSheet);
  const columns = ['Completed VI','Assignee'].map(function(name){return headers[name] || 0;})
    .filter(Boolean).sort(function(a,b){return b-a;});
  columns.forEach(function(column){requestSheet.deleteColumn(column);});
  SpreadsheetApp.flush();
  ss.toast('구형 시트 '+deletedSheets.length+'개와 레거시 열 '+columns.length+'개를 삭제했습니다.','Listing Ops',6);
  return {deletedSheets:deletedSheets, deletedColumns:columns.length};
}

function hasUnassignedRemainingV2_(requestId) {
  const rt=readOpsTable_(APP_CONFIG.SHEETS.REQUESTS),records=rt.byId[String(requestId)]||[];
  if(records.length!==1)throw Error('요청 ID가 없거나 중복되었습니다: '+requestId);
  const r=records[0].values,type=String(opsValue_(r,rt.headers,'Request Type')||'3P Listing'),source=opsSourceForRequest_(r,rt.headers);
  const index=buildTaskIndex_(readTaskSources_([source])),metrics=index[source][String(requestId)]||emptyTaskAggregate_();
  return (type!=='2P Lite 전환'&&metrics.missing3>0)||(type!=='3P Listing'&&metrics.missing2>0);
}

function refreshAllocationStatusV2_(requestId, preferredStatus) {
  return withOpsDocumentLock_(function(){return syncRequestSummaries_([requestId],{preferredAllocation:preferredStatus}).summaries[String(requestId)]['Allocation Status'];});
}

function handleAllocationManagementEditV2_(e) {
  if (!e || !e.range || e.range.getRow() < 2) return;
  const sheet = e.range.getSheet(), name = sheet.getName();
  const editHeaders = getHeaderMap_(sheet);
  const isThreePTask = name === APP_CONFIG.SHEETS.TASKS || name === APP_CONFIG.SHEETS.CRAWL_TASKS;
  const isTwoPTask = name === APP_CONFIG.SHEETS.TASKS_2P;
  if (isThreePTask) {
    requireAllocationTaskHeadersV4_(editHeaders,
      ['Request ID','3P Assignee','3P Completed','2P Assignee','2P Completed','VI Status','Invalid Reason'],name);
  } else if (isTwoPTask) {
    requireAllocationTaskHeadersV4_(editHeaders,
      ['Request ID','2P Assignee','2P Completed','VI Status','Invalid Reason'],name);
  }
  const firstColumn = e.range.getColumn(), lastColumn = e.range.getLastColumn();
  if (name === APP_CONFIG.SHEETS.REQUESTS && editHeaders['Status'] &&
      e.range.getColumn() <= editHeaders['Status'] && e.range.getLastColumn() >= editHeaders['Status']) {
    for (let row=e.range.getRow(); row<=e.range.getLastRow(); row+=1) {
      const status=normalizeRequestStatusV2_(sheet.getRange(row,editHeaders['Status']).getValue());
      if(status!==APP_CONFIG.STATUS.DROP){
        const requestId=String(sheet.getRange(row,editHeaders['Request ID']).getValue()||'');
        const wasCancelled=(e.range.getNumRows()===1&&normalizeRequestStatusV2_(e.oldValue)===APP_CONFIG.STATUS.DROP) ||
          (editHeaders['Cancelled At']&&sheet.getRange(row,editHeaders['Cancelled At']).getValue());
        if(requestId&&wasCancelled)reactivateCancelledRequestTasksV3_(requestId);
        continue;
      }
      const reason=editHeaders['Assignee Notes'] ? String(sheet.getRange(row,editHeaders['Assignee Notes']).getValue()||'').trim() : '';
      if(!reason){
        writeRequestStatuses_(sheet,row,editHeaders['Status'],[normalizeRequestStatusV2_(e.oldValue)||APP_CONFIG.STATUS.TO_DO]);
        SpreadsheetApp.getActive().toast('Assignee Notes에 DROP 사유를 입력한 뒤 DROP으로 변경하세요.','DROP 사유 필요',6);
        continue;
      }
      const requestId=String(sheet.getRange(row,editHeaders['Request ID']).getValue()||'');
      const cancelledAt=editHeaders['Cancelled At']?sheet.getRange(row,editHeaders['Cancelled At']).getValue():null;
      const now=cancelledAt instanceof Date&&!isNaN(cancelledAt.getTime())?cancelledAt:new Date();
      if(editHeaders['Cancelled At']&&!cancelledAt)sheet.getRange(row,editHeaders['Cancelled At']).setValue(now);
      // Task validity is still authoritative here. Close every active session
      // first so its End Completed VI / Delta cannot collapse to zero after
      // DROP invalidates the Task rows.
      if(requestId)cancelRequestTasksAndSessionsV1_(requestId,reason,now);
      if(editHeaders['Active Assignees'])sheet.getRange(row,editHeaders['Active Assignees']).clearContent();
    }
    bumpMyWorkSidebarRevisionV2_([]);
  }
  if ((isThreePTask || isTwoPTask) &&
      (allocationEditTouchesColumnV4_(firstColumn,lastColumn,editHeaders['VI Status']) ||
       allocationEditTouchesColumnV4_(firstColumn,lastColumn,editHeaders['Invalid Reason']))) {
    const validityIds={};
    for (let row=e.range.getRow(); row<=e.range.getLastRow(); row+=1) {
      const requestId=String(sheet.getRange(row,editHeaders['Request ID']).getValue()||'');
      let status=String(sheet.getRange(row,editHeaders['VI Status']).getValue()||'VALID').toUpperCase();
      const reason=String(sheet.getRange(row,editHeaders['Invalid Reason']).getValue()||'').trim();
      if(reason && status!=='INVALID') { sheet.getRange(row,editHeaders['VI Status']).setValue('INVALID'); status='INVALID'; }
      if(!reason && status==='INVALID' && allocationEditTouchesColumnV4_(firstColumn,lastColumn,editHeaders['Invalid Reason'])) {
        sheet.getRange(row,editHeaders['VI Status']).setValue('VALID');
        if(requestId)validityIds[requestId]=true;
        continue;
      }
      if(status==='INVALID'&&!reason){
        sheet.getRange(row,editHeaders['VI Status']).setValue('VALID');
        SpreadsheetApp.getActive().toast('Invalid Reason을 먼저 입력한 뒤 INVALID로 변경하세요.','무효 사유 필요',6);
        continue;
      }
      if(requestId)validityIds[requestId]=true;
    }
    Object.keys(validityIds).forEach(function(id){recalculateValidViMetricsV3_(id);refreshAllocationStatusV2_(id);});
    bumpMyWorkSidebarRevisionV2_([]);
  }
  const affectedWorkers = [];
  function addAffected_(value) { splitWorkersV2_(value).forEach(function(worker){if(affectedWorkers.indexOf(worker)<0)affectedWorkers.push(worker);}); }
  if (isThreePTask) {
    const owner3Column=editHeaders['3P Assignee'],owner2Column=editHeaders['2P Assignee'];
    const ownerStart=Math.min(owner3Column,owner2Column),ownerWidth=Math.max(owner3Column,owner2Column)-ownerStart+1;
    sheet.getRange(e.range.getRow(),ownerStart,e.range.getNumRows(),ownerWidth).getDisplayValues()
      .forEach(function(r){addAffected_(r[owner3Column-ownerStart]);addAffected_(r[owner2Column-ownerStart]);});
    if(allocationEditTouchesColumnV4_(firstColumn,lastColumn,owner3Column)||allocationEditTouchesColumnV4_(firstColumn,lastColumn,owner2Column))addAffected_(e.oldValue);
    bumpMyWorkSidebarRevisionV2_(affectedWorkers);
  }
  if (isTwoPTask) {
    const assigneeColumn = editHeaders['2P Assignee'];
    sheet.getRange(e.range.getRow(),assigneeColumn,e.range.getNumRows(),1).getDisplayValues().forEach(function(r){addAffected_(r[0]);});
    if(allocationEditTouchesColumnV4_(firstColumn,lastColumn,assigneeColumn))addAffected_(e.oldValue);
    bumpMyWorkSidebarRevisionV2_(affectedWorkers);
  }
  if (name === APP_CONFIG.SHEETS.REQUESTS) {
    const h = ensureAllocationColumnsV2_();
    if (e.range.getColumn() <= h['Active Assignees'] && e.range.getLastColumn() >= h['Active Assignees']) {
      try {
        applyRequestAssignmentEdit_(e);
        SpreadsheetApp.getActive().toast('Active Assignees 변경에 따라 남은 VI를 자동분배했습니다.','Listing Ops',4);
      } catch (error) {
        SpreadsheetApp.getActive().toast('자동분배 실패: '+(error&&error.message?error.message:String(error)),'Listing Ops',8);
        throw error;
      }
    }
    return;
  }
  const ids = {}, participantIds = {};
  const taskStateColumns=isThreePTask
    ? ['Request ID','3P Assignee','3P Completed','2P Assignee','2P Completed','VI Status','Invalid Reason']
    : (isTwoPTask?['Request ID','2P Assignee','2P Completed','VI Status','Invalid Reason']:[]);
  const touchesTaskState=taskStateColumns.some(function(header){return allocationEditTouchesColumnV4_(firstColumn,lastColumn,editHeaders[header]);});
  if (touchesTaskState) {
    sheet.getRange(e.range.getRow(),editHeaders['Request ID'],e.range.getNumRows(),1).getDisplayValues()
      .forEach(function(r){if(r[0])ids[r[0]]=true;});
  }
  const touches3PCompleted=isThreePTask&&allocationEditTouchesColumnV4_(firstColumn,lastColumn,editHeaders['3P Completed']);
  const touchesCombined2PCompleted=isThreePTask&&allocationEditTouchesColumnV4_(firstColumn,lastColumn,editHeaders['2P Completed']);
  if (touches3PCompleted || touchesCombined2PCompleted) {
    sheet.getRange(e.range.getRow(),1,e.range.getNumRows(),sheet.getLastColumn()).getValues().forEach(function(r){
      const requestId = String(r[editHeaders['Request ID']-1] || ''); if (!requestId) return;
      if (touches3PCompleted) r[editHeaders['3P Completed']-1] === true ? addParticipantV2_(requestId,r[editHeaders['3P Assignee']-1]) : participantIds[requestId]=true;
      if (touchesCombined2PCompleted) r[editHeaders['2P Completed']-1] === true ? addParticipantV2_(requestId,r[editHeaders['2P Assignee']-1]) : participantIds[requestId]=true;
    });
  }
  const twoPAssigneeColumn = isTwoPTask ? editHeaders['2P Assignee'] : 0;
  const twoPCompletedColumn = isTwoPTask ? editHeaders['2P Completed'] : 0;
  if (isTwoPTask && allocationEditTouchesColumnV4_(firstColumn,lastColumn,twoPAssigneeColumn)) {
    sheet.getRange(e.range.getRow(),editHeaders['Request ID'],e.range.getNumRows(),1).getDisplayValues().forEach(function(r){if(r[0])ids[r[0]]=true;});
  }
  if (isTwoPTask && allocationEditTouchesColumnV4_(firstColumn,lastColumn,twoPCompletedColumn)) {
    sheet.getRange(e.range.getRow(),1,e.range.getNumRows(),sheet.getLastColumn()).getValues().forEach(function(r){
      const requestId = String(r[editHeaders['Request ID']-1] || ''); if (!requestId) return;
      r[twoPCompletedColumn-1] === true ? addParticipantV2_(requestId,r[twoPAssigneeColumn-1]) : participantIds[requestId]=true;
    });
  }
  Object.keys(ids).forEach(function(id){recalculateValidViMetricsV3_(id);refreshAllocationStatusV2_(id,ALLOCATION_STATUS_V2.MANUAL);});
  Object.keys(participantIds).forEach(function(id){syncParticipantsV2_(id);});
}

// 행 순서를 유지하면서 작업자별 연속 구간을 만듦.
function buildContiguousWorkerQueueV2_(count, workers) {
  const queue = [];
  workers.forEach(function(worker, index) {
    const size = Math.floor(count / workers.length) +
      (index < count % workers.length ? 1 : 0);
    for (let i = 0; i < size; i += 1) queue.push(worker);
  });
  return queue;
}

function normalizeProductIdentifierV2_(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toUpperCase();
}

function buildProductGroupingContextV2_(rows, targetIndices, skuIndex, nameIndex) {
  const skuToNames = {}, nameToSkus = {}, skuValues = {}, nameValues = {};
  function addPair_(map, key, value) {
    if (!key || !value) return;
    if (!map[key]) map[key] = {};
    map[key][value] = true;
  }
  targetIndices.forEach(function(rowIndex) {
    const sku = normalizeProductIdentifierV2_(rows[rowIndex][skuIndex]);
    const name = normalizeProductIdentifierV2_(rows[rowIndex][nameIndex]);
    if (sku) skuValues[sku] = true;
    if (name) nameValues[name] = true;
    addPair_(skuToNames, sku, name);
    addPair_(nameToSkus, name, sku);
  });
  return {
    skuToNames:skuToNames,
    nameToSkus:nameToSkus,
    skuDistinct:Object.keys(skuValues).length,
    nameDistinct:Object.keys(nameValues).length
  };
}

function getProductGroupKeyV2_(row, rowIndex, context, skuIndex, nameIndex) {
  const sku = normalizeProductIdentifierV2_(row[skuIndex]);
  const productName = normalizeProductIdentifierV2_(row[nameIndex]);
  if (!sku && productName) return 'NAME:' + productName;
  if (!productName && sku) return 'SKU:' + sku;
  if (sku && productName) {
    const namesForSku = Object.keys(context.skuToNames[sku] || {}).length;
    const skusForName = Object.keys(context.nameToSkus[productName] || {}).length;
    // 한 값이 여러 반대편 상품값과 연결되면 공통 안내/임의값으로 간주함.
    if (namesForSku > 1 && skusForName <= 1) return 'NAME:' + productName;
    if (skusForName > 1 && namesForSku <= 1) return 'SKU:' + sku;
    // 요청 전체에서 한쪽 값이 사실상 고정이고 다른 쪽만 상품별로 달라지는 경우임.
    if (context.skuDistinct <= 1 && context.nameDistinct > 1) return 'NAME:' + productName;
    if (context.nameDistinct <= 1 && context.skuDistinct > 1) return 'SKU:' + sku;
    // 두 값 모두 상품 식별력이 있으면 더 안정적인 품번을 우선함.
    return 'SKU:' + sku;
  }
  return 'ROW:' + rowIndex;
}

// 행 순서가 최우선임. 연속한 동일 상품만 묶고 경계에서 균등분배함.
// 뒤쪽에 재등장한 상품을 전역 그룹으로 합치면 담당자가 앞 작업자로 돌아감.
// 완료/무효 행은 호출자가 제외하므로 이 함수에서 해당 이력을 변경하지 않음.
function buildContiguousGroupAssignmentsV2_(rows, targetIndices, workers, skuIndex, nameIndex) {
  const groups = [], ordered=targetIndices.slice().sort(function(a,b){return a-b;});
  const context = buildProductGroupingContextV2_(rows, ordered, skuIndex, nameIndex);
  ordered.forEach(function(rowIndex) {
    const key = getProductGroupKeyV2_(rows[rowIndex], rowIndex, context, skuIndex, nameIndex);
    let group=groups[groups.length-1];
    if(!group||group.key!==key){group={key:key,rows:[]};groups.push(group);}
    group.rows.push(rowIndex);
  });
  const assignments = {};
  let workerIndex = 0, assignedToWorker = 0, remainingRows = targetIndices.length;
  groups.forEach(function(group, groupIndex) {
    const groupsLeft = groups.length - groupIndex;
    const workersLeft = workers.length - workerIndex;
    if (workerIndex < workers.length - 1 && assignedToWorker > 0) {
      const target = remainingRows / workersLeft;
      const stopDistance = Math.abs(target - assignedToWorker);
      const addDistance = Math.abs(target - (assignedToWorker + group.rows.length));
      const mustLeaveGroup = groupsLeft <= workersLeft - 1;
      if (mustLeaveGroup || stopDistance < addDistance) {
        remainingRows -= assignedToWorker;
        workerIndex += 1;
        assignedToWorker = 0;
      }
    }
    group.rows.forEach(function(rowIndex) { assignments[rowIndex] = workers[workerIndex]; });
    assignedToWorker += group.rows.length;
  });
  return assignments;
}

function getRequestTaskBlockV2_(sheet, requestId, width, providedHeaders) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const headers=providedHeaders||getHeaderMap_(sheet);
  requireAllocationTaskHeadersV4_(headers,['Request ID'],sheet.getName ? sheet.getName() : 'Task');
  const idColumn=headers['Request ID'];
  // Read IDs once instead of making one Range.getRow call per matching VI.
  // Preserve the enclosing block, including interleaved requests; callers filter IDs.
  const ids = sheet.getRange(2,idColumn,lastRow-1,1).getDisplayValues();
  const target = String(requestId).toLowerCase();
  let startRow = 0, endRow = 0;
  ids.forEach(function(row,index) {
    if (String(row[0]).toLowerCase() !== target) return;
    if (!startRow) startRow = index + 2;
    endRow = index + 2;
  });
  if (!startRow) return null;
  return {startRow:startRow, rows:sheet.getRange(startRow,1,endRow-startRow+1,width).getValues()};
}

// Normal intake appends every Request as one contiguous Task block.  For a
// direct REQUESTS edit, stop at that block instead of reading the whole ID
// column (simple onEdit has a hard 30 second limit).
function getContiguousRequestTaskBlockV3_(sheet,requestId,width,minimumRows,providedHeaders){
  const lastRow=sheet.getLastRow();
  if(lastRow<2)return null;
  const headers=providedHeaders||getHeaderMap_(sheet);
  requireAllocationTaskHeadersV4_(headers,['Request ID'],sheet.getName ? sheet.getName() : 'Task');
  const idColumn=headers['Request ID'];
  const id=String(requestId),first=sheet.getRange(2,idColumn,lastRow-1,1)
    .createTextFinder(id).matchEntireCell(true).matchCase(true).useRegularExpression(false).findNext();
  if(!first)return null;
  const startRow=first.getRow(),ids=[];
  let cursor=startRow,finished=false;
  while(cursor<=lastRow&&!finished){
    const remainingMinimum=Math.max(Number(minimumRows||0)-ids.length,0);
    const count=Math.min(Math.max(remainingMinimum+128,512),5000,lastRow-cursor+1);
    const values=sheet.getRange(cursor,idColumn,count,1).getDisplayValues();
    for(let i=0;i<values.length;i+=1){
      if(String(values[i][0])!==id){finished=true;break;}
      ids.push(id);
    }
    cursor+=count;
  }
  if(Number(minimumRows||0)>ids.length){
    throw Error('해당 REQ의 Task 행이 여러 구간으로 섞여 있어 수기 자동분배를 중단했습니다. 관리자 작업자 배정을 사용해주세요.');
  }
  return {startRow:startRow,rows:sheet.getRange(startRow,1,ids.length,width).getValues()};
}

function writeAssignmentTargetsV2_(sheet,start,column,rows,targets,assignments,rowNumbers){
  const physical=function(i){return rowNumbers?rowNumbers[i]:start+i;};
  let run=[];function flush(){if(run.length)sheet.getRange(physical(run[0]),column,run.length,1).setValues(run.map(function(i){return [assignments[i]];}));run=[];}
  targets.forEach(function(i){if(rows[i][column-1]===assignments[i])return;if(run.length&&physical(i)!==physical(run[run.length-1])+1)flush();run.push(i);});flush();
}
function applyRemainingDistributionV2_(requestId, workers, context) {
  const requestSheet = context?context.rt.sheet:getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS), h = context?context.rt.headers:getHeaderMap_(requestSheet);
  const requestRecord=context?context.rt.byId[requestId][0]:null;
  const requestRow = requestRecord?requestRecord.row:findRowByRequestId_(requestSheet,requestId);
  const type = String(requestRecord?opsValue_(requestRecord.values,h,'Request Type'):requestSheet.getRange(requestRow,h['Request Type']).getValue() || '');
  const minimumRows=Number(requestRecord?opsValue_(requestRecord.values,h,'Total VI'):(h['Total VI']?requestSheet.getRange(requestRow,h['Total VI']).getValue():0))||0;
  const affected = workers.slice();
  function remember_(value) {
    splitWorkersV2_(value).forEach(function(worker){if(affected.indexOf(worker)<0)affected.push(worker);});
  }
  if (type === '2P Lite 전환') {
    const sheet = getRequiredSheet_(APP_CONFIG.SHEETS.TASKS_2P);
    const headers = context?context.sources[APP_CONFIG.SHEETS.TASKS_2P].headers:getHeaderMap_(sheet);
    requireAllocationTaskHeadersV4_(headers,['Request ID','업체상품ID','Product ID','2P Assignee','2P Completed','VI Status'],APP_CONFIG.SHEETS.TASKS_2P);
    const block = context?nativeTaskBlock_(context.sources[APP_CONFIG.SHEETS.TASKS_2P],requestId):getContiguousRequestTaskBlockV3_(sheet,requestId,sheet.getLastColumn(),minimumRows,headers);
    if (block) {
      const skuIndex=headers['업체상품ID']-1,nameIndex=headers['Product ID']-1;
      const targets=[],owner=headers['2P Assignee']-1,done=headers['2P Completed']-1,idIndex=headers['Request ID']-1;
      block.rows.forEach(function(row,index){
        if(String(row[idIndex])!==requestId)return;
        remember_(row[owner]);
        if(isViValidV3_(row,headers) && row[done]!==true)targets.push(index);
      });
      const assignments=buildContiguousGroupAssignmentsV2_(block.rows,targets,workers,skuIndex,nameIndex);
      writeAssignmentTargetsV2_(sheet,block.startRow,owner+1,block.rows,targets,assignments,block.rowNumbers);
    }
  } else {
    const isCrawl=h['크롤링 필요']&&(requestRecord?opsValue_(requestRecord.values,h,'크롤링 필요'):requestSheet.getRange(requestRow,h['크롤링 필요']).getValue())===true;
    const sheet = getRequiredSheet_(isCrawl?APP_CONFIG.SHEETS.CRAWL_TASKS:APP_CONFIG.SHEETS.TASKS);
    const sourceName=isCrawl?APP_CONFIG.SHEETS.CRAWL_TASKS:APP_CONFIG.SHEETS.TASKS;
    const headers=context?context.sources[sourceName].headers:getHeaderMap_(sheet);
    const required=['Request ID','품번','등록상품명','3P Assignee','3P Completed','VI Status'];
    if(type!=='3P Listing')Array.prototype.push.apply(required,['2P Assignee','2P Completed']);
    requireAllocationTaskHeadersV4_(headers,required,sourceName);
    const block = context?nativeTaskBlock_(context.sources[sourceName],requestId):getContiguousRequestTaskBlockV3_(sheet,requestId,sheet.getLastColumn(),minimumRows,headers);
    if (block) {
      const skuIndex=headers['품번']-1,nameIndex=headers['등록상품명']-1,idIndex=headers['Request ID']-1;
      const targets3=[],targets2=[],owner3=headers['3P Assignee']-1,done3=headers['3P Completed']-1;
      const owner2=type!=='3P Listing'?headers['2P Assignee']-1:-1,done2=type!=='3P Listing'?headers['2P Completed']-1:-1;
      block.rows.forEach(function(row,index){
        if(String(row[idIndex])!==requestId)return;
        remember_(row[owner3]); if(owner2>=0)remember_(row[owner2]);
        if(!isViValidV3_(row,headers))return;
        if(row[done3]!==true)targets3.push(index);
        if(done2>=0&&row[done2]!==true)targets2.push(index);
      });
      const assignments3=buildContiguousGroupAssignmentsV2_(block.rows,targets3,workers,skuIndex,nameIndex);
      const assignments2=done2>=0?buildContiguousGroupAssignmentsV2_(block.rows,targets2,workers,skuIndex,nameIndex):[];
      writeAssignmentTargetsV2_(sheet,block.startRow,owner3+1,block.rows,targets3,assignments3,block.rowNumbers);
      if(type!=='3P Listing')writeAssignmentTargetsV2_(sheet,block.startRow,owner2+1,block.rows,targets2,assignments2,block.rowNumbers);
    }
  }
  if (h['Allocation Status']) requestSheet.getRange(requestRow,h['Allocation Status']).setValue(ALLOCATION_STATUS_V2.AUTO);
  if(context)affected.forEach(function(w){context.workers.add(w);});else bumpMyWorkSidebarRevisionV2_(affected);
  logActivity_(requestId,'REMAINING_VI_AUTO_DISTRIBUTED','',workers.join(', '));
}

function ensureOperationsGuideV2_() {
  const ss=SpreadsheetApp.getActiveSpreadsheet(),name='운영_가이드';let s=ss.getSheetByName(name);if(!s)s=ss.insertSheet(name);
  const values=[
    ['LISTING OPS 배정 규칙','관리자가 바로 확인할 수 있는 운영 기준'],
    ['항목','규칙'],
    ['Participants','완료된 3P·바코드·2P 이력에서 자동 집계되는 전체 참여자입니다. 직접 배정 기준으로 사용하지 않습니다.'],
    ['Active Assignees','앞으로 남은 VI를 수행할 작업자입니다. 시트 또는 팀 Kanban에서 수정해도 Task는 즉시 바뀌지 않습니다.'],
    ['자동분배됨','관리자가 미리보기 확인 후 남은 VI 균등분배를 실행한 상태입니다.'],
    ['분배 필요','Request의 모든 Task에서 실제로 수행하는 단계 중 담당자 없는 셀이 하나라도 있는 상태입니다. Active Assignees 공백만으로는 분배 필요가 되지 않습니다.'],
    ['수동분배됨','작업자가 TASKS에서 미완료 VI 담당자를 직접 수정한 상태입니다.'],
    ['완료 작업 보호','완료된 3P·2P 담당자와 완료 시각은 자동분배 시 변경하지 않습니다.'],
    ['3P·2P 동일 배정','두 단계가 모두 미완료면 같은 작업자를 우선 배정합니다.'],
    ['상품 묶음 배정','같은 상품의 모든 옵션은 한 작업자에게 배정합니다. 특정 문구나 언어를 기준으로 판단하지 않습니다.'],
    ['식별값 자동 판단','요청 안에서 품번과 등록상품명의 대응 관계를 분석합니다. 여러 상품에 반복되는 공통 안내값은 제외하고 실제로 상품을 구분하는 값을 선택합니다.'],
    ['연속 구간 배정','상품 묶음의 행 순서를 유지해 A 작업자 구간 다음 B 작업자 구간으로 배정하며, 묶음을 쪼개지 않습니다.'],
    ['재분배 실행','REQUESTS에서 요청 행 선택 → Listing Ops → 선택 요청 남은 VI 균등분배 → 미리보기 확인 후 실행합니다.'],
    ['개인 작업량','REQUESTS 명단이 아니라 TASKS에서 실제로 본인에게 배정된 미완료 VI만 계산합니다.']
  ];
  s.clear();s.getRange(1,1,values.length,2).setValues(values);s.getRange('A1:B1').setBackground('#16324F').setFontColor('#fff').setFontWeight('bold').setFontSize(14);s.getRange('A2:B2').setBackground('#263650').setFontColor('#fff').setFontWeight('bold');s.getRange(3,1,values.length-2,1).setFontWeight('bold').setBackground('#eef3f8');s.getRange(1,1,values.length,2).setWrap(true).setVerticalAlignment('top');s.setColumnWidth(1,150);s.setColumnWidth(2,620);s.setFrozenRows(2);return s;
}

function openOperationsGuideV2() {
  ensureOperationsGuideV2_().activate();
}

function setupAllocationManagementV2() {
  ensureValidityManagementV3_();
  const sheet=getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS),h=ensureAllocationColumnsV2_();
  getDataRows_(sheet).forEach(function(r){const id=String(r[h['Request ID']-1]||'');if(id)refreshAllocationStatusV2_(id,String(r[h['Allocation Status']-1]||''));});
  ensureOperationsGuideV2_();
}
