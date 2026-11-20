// 기존 호출 호환용 전체 갱신 진입점임.
// Kanban/Dashboard 팝업은 원본 데이터를 직접 읽으므로 Request ID 링크만 갱신함.
function refreshAllViews() {
  return withFilterViewRefreshContext_(function(){
    return refreshRequestIdLinks_();
  });
}

// 삭제 전 레거시 KANBAN 탭이 남아 있을 때만 다시 그림.
// 탭이 없으면 정상적으로 건너뛰며 팝업에는 영향을 주지 않음.
function refreshKanban() {
  const target = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('KANBAN');
  if (!target) return {skipped:true, reason:'KANBAN sheet removed'};
  const source = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  const taskSheet = getRequiredSheet_(APP_CONFIG.SHEETS.TASKS);
  const taskSheet2P = getRequiredSheet_(APP_CONFIG.SHEETS.TASKS_2P);
  const rows = getDataRows_(source);
  const headers = getHeaderMap_(source);
  const statuses = [
    APP_CONFIG.STATUS.TO_DO,
    APP_CONFIG.STATUS.ONGOING,
    APP_CONFIG.STATUS.ON_HOLD,
    APP_CONFIG.STATUS.DONE,
  ];
  const statusLabels = ['TO DO', 'ONGOING', 'ON HOLD', 'DONE'];
  const statusColors = ['#5475B8', '#5BB5E5', '#F17868', '#32B858'];
  const requestTypes = ['3P Listing', '2P Lite 전환', '3P Listing 후 2P Lite 전환'];
  const doneRetentionDays = 14;
  const doneCutoff = new Date();
  doneCutoff.setDate(doneCutoff.getDate() - doneRetentionDays);
  const isVisibleOnBoard = function(row) {
    const status = normalizeRequestStatusV2_(row[headers['Status'] - 1]);
    if (status !== APP_CONFIG.STATUS.DONE) return true;
    const completedAt = row[headers['Completed At'] - 1];
    return !(completedAt instanceof Date) || completedAt >= doneCutoff;
  };
  const visibleRows = rows.filter(isVisibleOnBoard);
  const oldDoneCount = rows.filter(function(row) {
    return normalizeRequestStatusV2_(row[headers['Status'] - 1]) === APP_CONFIG.STATUS.DONE && !isVisibleOnBoard(row);
  }).length;
  const laneLabels = ['3P LISTING', '2P LITE', '3P + 2P 모두'];
  const laneColors = ['#27385D', '#314A70', '#3E5578'];
  const archivedLinks = {};
  if(rows.some(function(row){return archiveLinkArchivedV1_(row,headers);})){
    archiveLinkSnapshotV1_(SpreadsheetApp.getActiveSpreadsheet()).records.forEach(function(r){archivedLinks[r.id]=archiveLinkReadyV1_(r.url,r.fileId)?r.url:r.fileId?'https://docs.google.com/spreadsheets/d/'+r.fileId+'/edit':'';});
  }
  const requestIds3P = rows.filter(function(row) {
    return !archiveLinkArchivedV1_(row,headers) && String(row[headers['Request Type'] - 1]) !== '2P Lite 전환';
  }).map(function(row) { return String(row[headers['Request ID'] - 1] || ''); }).filter(String);
  const requestIds2P = rows.filter(function(row) {
    return !archiveLinkArchivedV1_(row,headers) && String(row[headers['Request Type'] - 1]) === '2P Lite 전환';
  }).map(function(row) { return String(row[headers['Request ID'] - 1] || ''); }).filter(String);
  const filterViewIds3P = ensureRequestFilterViews_(requestIds3P, taskSheet);
  const filterViewIds2P = ensureRequestFilterViews_(requestIds2P, taskSheet2P);

  target.getRange('A1:M300').breakApart().clearContent().clearFormat().setFontFamily('Roboto');
  target.setHiddenGridlines(true);
  target.setRowHeights(1, 300, 22);
  target.getRange('A1:M1').merge()
    .setValue('LISTING OPERATIONS')
    .setBackground('#FFFFFF').setFontColor('#202938')
    .setFontWeight('bold').setFontSize(17)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  target.getRange('A2:M2').merge()
    .setValue('WORKFLOW  /  REQUEST TYPE SWIMLANE  ·  DONE 최근 14일')
    .setBackground('#FFFFFF').setFontColor('#8791A1')
    .setFontSize(9).setFontWeight('bold')
    .setBorder(false, false, true, false, false, false, '#E1E5EA', SpreadsheetApp.BorderStyle.SOLID);

  statuses.forEach(function(status, statusIndex) {
    const startColumn = 2 + statusIndex * 3;
    const count = visibleRows.filter(function(row) {
      return normalizeRequestStatusV2_(row[headers['Status'] - 1]) === status;
    }).length;
    const headerText = status === APP_CONFIG.STATUS.DONE && oldDoneCount > 0
      ? statusLabels[statusIndex] + '    ' + count + '  ·  이전 ' + oldDoneCount
      : statusLabels[statusIndex] + '    ' + count;
    target.getRange(4, startColumn, 1, 3).merge()
      .setValue(headerText)
      .setBackground(statusColors[statusIndex]).setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });

  let laneStartRow = 6;
  requestTypes.forEach(function(requestType, typeIndex) {
    const groups = statuses.map(function(status) {
      return visibleRows.filter(function(row) {
        return normalizeRequestStatusV2_(row[headers['Status'] - 1]) === status &&
          row[headers['Request Type'] - 1] === requestType;
      });
    });
    const totalCardsInLane = groups.reduce(function(sum, group) { return sum + group.length; }, 0);
    const maxCards = Math.max(1, Math.max.apply(null, groups.map(function(group) { return group.length; })));
    const laneHeight = totalCardsInLane === 0 ? 5 : maxCards * 8 + 1;
    const laneRange = target.getRange(laneStartRow, 1, laneHeight, 13);
    laneRange.setBackground('#F7F8FA')
      .setBorder(false, false, true, false, false, false, '#CDD3DB', SpreadsheetApp.BorderStyle.SOLID);

    target.getRange(laneStartRow, 1, laneHeight, 1).merge()
      .setValue(laneLabels[typeIndex])
      .setBackground(laneColors[typeIndex]).setFontColor('#FFFFFF')
      .setFontWeight('bold').setFontSize(10)
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setTextRotation(90);

    groups.forEach(function(cards, statusIndex) {
      const startColumn = 2 + statusIndex * 3;
      if (!cards.length) {
        target.getRange(laneStartRow + 1, startColumn, 2, 3).merge()
          .setValue('No requests')
          .setFontColor('#AFB6C1').setFontSize(9)
          .setHorizontalAlignment('center').setVerticalAlignment('middle');
        return;
      }

      cards.forEach(function(row, cardIndex) {
        const cardRow = laneStartRow + 1 + cardIndex * 8;
        const requestId = String(row[headers['Request ID'] - 1] || '');
        const brand = String(row[headers['Brand'] - 1] || '브랜드 미지정');
        const priority = String(row[headers['Priority'] - 1] || 'NORMAL');
        const assignee = String(row[headers['Active Assignees'] - 1] || '미배정');
        const node = String(row[headers['Node'] - 1] || '미지정');
        const total = Number(row[headers['Total VI'] - 1] || 0);
        const done3P = Number(row[headers['3P Completed VI'] - 1] || 0);
        const done2P = Number(row[headers['2P Completed VI'] - 1] || 0);
        const linkedTaskSheet = requestType === '2P Lite 전환' ? taskSheet2P : taskSheet;
        const linkedFilterIds = requestType === '2P Lite 전환' ? filterViewIds2P : filterViewIds3P;
        const filterViewId = linkedFilterIds[requestId];
        const url = archiveLinkArchivedV1_(row,headers) ? archivedLinks[requestId] || null : filterViewId
          ? SpreadsheetApp.getActiveSpreadsheet().getUrl() + '#gid=' + linkedTaskSheet.getSheetId() + '&fvid=' + filterViewId
          : SpreadsheetApp.getActiveSpreadsheet().getUrl() + '#gid=' + linkedTaskSheet.getSheetId();

        const cardRange = target.getRange(cardRow, startColumn, 7, 3);
        cardRange.setBackground('#FFFFFF')
          .setBorder(true, true, true, true, false, false, '#E1E5EA', SpreadsheetApp.BorderStyle.SOLID);
        target.getRange(cardRow, startColumn, 7, 1)
          .setBorder(false, true, false, false, false, false, statusColors[statusIndex], SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

        const idText = SpreadsheetApp.newRichTextValue()
          .setText(requestId).setLinkUrl(url).build();
        target.getRange(cardRow, startColumn, 1, 2).merge()
          .setRichTextValue(idText).setFontColor('#4A5668')
          .setFontWeight('bold').setFontSize(9)
          .setVerticalAlignment('middle');
        const priorityBg = priority === 'HIGH' ? '#F9E7E5' : '#F0F2F5';
        const priorityFg = priority === 'HIGH' ? '#A64D45' : '#687386';
        target.getRange(cardRow, startColumn + 2)
          .setValue(priority).setBackground(priorityBg).setFontColor(priorityFg)
          .setFontWeight('bold').setFontSize(8).setHorizontalAlignment('center');

        target.getRange(cardRow + 1, startColumn, 1, 3).merge()
          .setValue(brand).setFontColor('#263142')
          .setFontWeight('bold').setFontSize(12)
          .setVerticalAlignment('middle');
        target.getRange(cardRow + 2, startColumn, 1, 3).merge()
          .setValue('NODE  ' + node + '    ·    ASSIGNEE  ' + assignee)
          .setFontColor('#7B8492').setFontSize(8)
          .setVerticalAlignment('middle');

        if (requestType === '3P Listing') {
          drawSwimlaneProgress_(target, cardRow + 3, startColumn, '3P LISTING', done3P, total, statusColors[statusIndex]);
        } else if (requestType === '2P Lite 전환') {
          drawSwimlaneProgress_(target, cardRow + 3, startColumn, '2P LITE', done2P, total, statusColors[statusIndex]);
        } else {
          drawSwimlaneProgress_(target, cardRow + 3, startColumn, '3P LISTING', done3P, total, '#5475B8');
          drawSwimlaneProgress_(target, cardRow + 5, startColumn, '2P LITE', done2P, total, '#5A9D91');
        }

        target.setRowHeight(cardRow, 20);
        target.setRowHeight(cardRow + 1, 27);
        target.setRowHeight(cardRow + 2, 18);
        target.setRowHeights(cardRow + 3, 4, 18);
        target.setRowHeight(cardRow + 7, 8);
      });
    });
    laneStartRow += laneHeight;
  });

  target.setColumnWidth(1, 48);
  for (let column = 2; column <= 13; column += 1) target.setColumnWidth(column, 86);
  target.setRowHeight(1, 38);
  target.setRowHeight(2, 24);
  target.setRowHeight(3, 12);
  target.setRowHeight(4, 34);
  target.setRowHeight(5, 12);
  target.setFrozenRows(4);
}

// Swimlane 카드 안에 수치와 SPARKLINE 진행률 막대를 그림.
function drawSwimlaneProgress_(sheet, row, column, label, completed, total, color) {
  const percent = total > 0 ? Math.max(0, Math.min(1, completed / total)) : 0;
  const remaining = Math.max(total - completed, 0);
  sheet.getRange(row, column, 1, 3).merge()
    .setValue(label + '   ' + completed + '/' + total + '  ·  남음 ' + remaining + '  ·  ' + Math.round(percent * 100) + '%')
    .setFontColor('#606B7B').setFontSize(8)
    .setVerticalAlignment('bottom');
  sheet.getRange(row + 1, column, 1, 3).merge()
    .setFormula('=SPARKLINE(' + percent + ',{"charttype","bar";"max",1;"color1","' + color + '"})')
    .setBackground('#E7EAF0');
}

// 삭제 전 레거시 DASHBOARD 탭이 남아 있을 때만 다시 그림.
// 운영 Dashboard 팝업은 operations_dashboard.gs에서 원본 데이터를 직접 집계함.
function refreshDashboard() {
  const target = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD');
  if (!target) return {skipped:true, reason:'DASHBOARD sheet removed'};
  const source = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  const headers = getHeaderMap_(source);
  const rows = getDataRows_(source).filter(function(row){return normalizeRequestStatusV2_(row[headers['Status']-1]) !== APP_CONFIG.STATUS.DROP;});
  const todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const statusOf = function(row) { return normalizeRequestStatusV2_(row[headers['Status'] - 1]); };
  const completedToday = rows.filter(function(row) {
    const value = row[headers['Completed At'] - 1];
    return value instanceof Date &&
      Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd') === todayKey;
  }).length;

  const leadTimes = numericValues_(rows, headers['Lead Time (Hours)']);
  const workTimes = numericValues_(rows, headers['Work Time (Hours)']);
  const values = [
    rows.filter(function(row) { return statusOf(row) !== APP_CONFIG.STATUS.DONE; }).length,
    rows.filter(function(row) { return statusOf(row) === APP_CONFIG.STATUS.ONGOING; }).length,
    completedToday,
    rows.map(function(row) { return Math.max(Number(row[headers['3P Completed VI'] - 1] || 0), Number(row[headers['2P Completed VI'] - 1] || 0)); }).reduce(sum_, 0),
    average_(leadTimes),
    average_(workTimes),
  ];
  const labels = [
    'Open Requests', 'ONGOING', 'Completed Today',
    'Total Completed VI', 'Avg Lead Time', 'Avg Work Time',
  ];

  target.getRange('A1:F30').breakApart().clearContent().clearFormat();
  target.setHiddenGridlines(true);

  target.getRange('A1:F1').merge()
    .setValue('LISTING OPERATIONS — DASHBOARD')
    .setBackground('#16324F')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('left');

  target.getRange('A2:F2').merge()
    .setValue('요청 진행 현황과 처리시간 요약')
    .setBackground('#EAF2FF')
    .setFontColor('#16324F');

  target.getRange('A4:F4').setValues([labels]);
  target.getRange('A5:F5').setValues([values]);
  target.getRange('A4:F4')
    .setFontWeight('bold')
    .setFontColor('#16324F')
    .setHorizontalAlignment('center')
    .setWrap(true);
  target.getRange('A5:F5')
    .setFontWeight('bold')
    .setFontSize(18)
    .setFontColor('#2563EB')
    .setHorizontalAlignment('center');
  target.getRange('A4:F5')
    .setBackground('#F5F7FA')
    .setBorder(true, true, true, true, true, true, '#D7DEE8', SpreadsheetApp.BorderStyle.SOLID);
  target.getRange('C4:D5').setBackground('#DDF4E4');
  target.getRange('E5:F5').setNumberFormat('[h]:mm');

  for (let column = 1; column <= 6; column += 1) {
    target.setColumnWidth(column, 145);
  }
  target.setRowHeight(1, 40);
  target.setRowHeight(2, 32);
  target.setRowHeight(4, 38);
  target.setRowHeight(5, 42);
  target.setFrozenRows(2);
}

// 지정된 컬럼에서 0이 아닌 숫자 값만 추출함.
function numericValues_(rows, oneBasedColumn) {
  return rows.map(function(row) { return Number(row[oneBasedColumn - 1]); })
    .filter(function(value) { return Number.isFinite(value) && value !== 0; });
}

// 숫자 배열의 평균을 계산함.
function average_(values) {
  return values.length ? values.reduce(sum_, 0) / values.length : 0;
}

// reduce에서 사용하는 합계 함수임.
function sum_(a, b) {
  return a + b;
}


// TASKS의 Request ID 자동 필터를 해제하고 전체 Task를 다시 표시함.
function clearTaskRequestFilter() {
  const taskSheet = getRequiredSheet_(APP_CONFIG.SHEETS.TASKS);
  const filter = taskSheet.getFilter();
  if (filter) filter.remove();
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(taskSheet);
}


// Request ID별 Google Sheets 필터 보기를 생성하거나 갱신함.
// 필터 보기는 링크를 연 사용자에게만 적용되어 다른 작업자의 화면을 바꾸지 않음.
// Execution-local only: never persists metadata across user actions.
var FILTER_VIEW_REFRESH_CONTEXT_ = null;
const FILTER_VIEW_ROW_CHUNK_=1000;
function ensureTaskFilterCapacity_(taskSheet,requiredLastRow){
  const current=taskSheet.getMaxRows();if(requiredLastRow<=current-100)return false;
  const target=Math.ceil((requiredLastRow+500)/FILTER_VIEW_ROW_CHUNK_)*FILTER_VIEW_ROW_CHUNK_;
  taskSheet.insertRowsAfter(current,target-current);
  const ss=SpreadsheetApp.getActiveSpreadsheet(),meta=Sheets.Spreadsheets.get(ss.getId(),{fields:'sheets(properties(sheetId),filterViews(filterViewId,title,range))'}),entry=(meta.sheets||[]).find(function(s){return s.properties&&s.properties.sheetId===taskSheet.getSheetId();}),requests=[];
  ((entry&&entry.filterViews)||[]).forEach(function(view){if(String(view.title||'').indexOf('[AUTO] ')!==0)return;const range=Object.assign({},view.range||{});range.endRowIndex=target;requests.push({updateFilterView:{filter:{filterViewId:view.filterViewId,range:range},fields:'range'}});});
  if(requests.length)Sheets.Spreadsheets.batchUpdate({requests:requests},ss.getId());
  FILTER_VIEW_REFRESH_CONTEXT_=null;return true;
}
function withFilterViewRefreshContext_(fn) {
  if(FILTER_VIEW_REFRESH_CONTEXT_)return fn();
  FILTER_VIEW_REFRESH_CONTEXT_={metadata:null};
  try{return fn();}finally{FILTER_VIEW_REFRESH_CONTEXT_=null;}
}
function filterViewOwnedFields_(filter) {
  function ordered(value){
    if(Array.isArray(value))return value.map(ordered);
    if(value&&typeof value==='object'){const result={};Object.keys(value).sort().forEach(function(k){result[k]=ordered(value[k]);});return result;}
    return value;
  }
  const range=Object.assign({startRowIndex:0,startColumnIndex:0},filter.range||{});
  return JSON.stringify(ordered({title:filter.title,range:range,criteria:filter.criteria||{},filterSpecs:filter.filterSpecs||[]}));
}
function ensureRequestFilterViews_(requestIds, taskSheet) {
  const ss=SpreadsheetApp.getActiveSpreadsheet(),spreadsheetId=ss.getId(),sheetId=taskSheet.getSheetId();
  const context=FILTER_VIEW_REFRESH_CONTEXT_||{metadata:null};
  if(!context.metadata)context.metadata=Sheets.Spreadsheets.get(spreadsheetId,{fields:'sheets(properties(sheetId),filterViews(filterViewId,title,range,criteria,filterSpecs))'});
  const sheets=context.metadata.sheets||(context.metadata.sheets=[]);
  let target=sheets.find(function(s){return s.properties&&s.properties.sheetId===sheetId;});
  if(!target){target={properties:{sheetId:sheetId},filterViews:[]};sheets.push(target);}
  const views=target.filterViews||(target.filterViews=[]),existing={},byRequest={};
  views.forEach(function(v){if(String(v.title||'').indexOf('[AUTO] ')===0){const id=v.title.substring(7);existing[id]=v.filterViewId;byRequest[id]=v;}});
  const migrationKey='AUTO_FILTER_MAX_ROWS_V2_'+sheetId,props=PropertiesService.getDocumentProperties();
  if(props.getProperty(migrationKey)!=='DONE'){
    const maxRows=taskSheet.getMaxRows(),updates=views.filter(function(v){return String(v.title||'').indexOf('[AUTO] ')===0&&(!v.range||v.range.endRowIndex!==maxRows);}).map(function(v){const range=Object.assign({},v.range||{sheetId:sheetId,startRowIndex:0,startColumnIndex:0,endColumnIndex:taskSheet.getLastColumn()});range.endRowIndex=maxRows;v.range=range;return {updateFilterView:{filter:{filterViewId:v.filterViewId,range:range},fields:'range'}};});
    if(updates.length)Sheets.Spreadsheets.batchUpdate({requests:updates},spreadsheetId);props.setProperty(migrationKey,'DONE');
  }
  const ids=Array.from(new Set(requestIds.filter(Boolean))),pending=[];
  const endRow=Math.max(taskSheet.getMaxRows(),2),endCol=taskSheet.getLastColumn();
  ids.forEach(function(id){
    const condition={type:'TEXT_EQ',values:[{userEnteredValue:id}]};
    const filter={title:'[AUTO] '+id,range:{sheetId:sheetId,startRowIndex:0,endRowIndex:endRow,startColumnIndex:0,endColumnIndex:endCol},criteria:{0:{condition:condition}},filterSpecs:[{columnIndex:0,filterCriteria:{condition:condition}}]};
    const old=byRequest[id];
    if(old&&filterViewOwnedFields_(old)===filterViewOwnedFields_(filter))return;
    if(old)filter.filterViewId=old.filterViewId;
    pending.push({id:id,filter:filter,old:old,request:old?{updateFilterView:{filter:filter,fields:'title,range,criteria,filterSpecs'}}:{addFilterView:{filter:filter}}});
  });
  for(let start=0;start<pending.length;start+=100){
    const group=pending.slice(start,start+100),response=Sheets.Spreadsheets.batchUpdate({requests:group.map(function(p){return p.request;})},spreadsheetId);
    group.forEach(function(p,i){
      if(p.old){Object.assign(p.old,p.filter);return;}
      const reply=(response.replies||[])[i],created=reply&&reply.addFilterView&&reply.addFilterView.filter;
      if(!created||created.filterViewId==null)throw Error('필터 보기 저장 응답을 확인할 수 없습니다. 새로고침 후 확인하세요.');
      const view=Object.assign({},p.filter,created);views.push(view);byRequest[p.id]=view;existing[p.id]=view.filterViewId;
    });
  }
  return existing;
}

// REQUESTS의 Request ID를 요청별 필터 보기 링크로 갱신함.
function refreshRequestIdLinks_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requestSheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  if (requestSheet.getLastRow() < 2) return;
  const headers = getHeaderMap_(requestSheet);
  const rowCount = requestSheet.getLastRow() - 1;
  const ids = requestSheet.getRange(2, headers['Request ID'], rowCount, 1).getDisplayValues()
    .map(function(row){return String(row[0] || '').trim();});
  const types = requestSheet.getRange(2, headers['Request Type'], rowCount, 1).getDisplayValues()
    .map(function(row){return String(row[0] || '');});
  const task3P = getRequiredSheet_(APP_CONFIG.SHEETS.TASKS);
  const crawl3P = getRequiredSheet_(APP_CONFIG.SHEETS.CRAWL_TASKS);
  const task2P = getRequiredSheet_(APP_CONFIG.SHEETS.TASKS_2P);
  const crawlFlags = headers['크롤링 필요'] ? requestSheet.getRange(2,headers['크롤링 필요'],rowCount,1).getValues() : [];
  const archiveStates=headers['Archive Status']?requestSheet.getRange(2,headers['Archive Status'],rowCount,1).getDisplayValues():[];
  const archived=index=>archiveStates[index]&&String(archiveStates[index][0]).toUpperCase()==='ARCHIVED';
  const ids3P = [], idsCrawl = [], ids2P = [];
  ids.forEach(function(id,index){
    if (!id || archived(index)) return;
    if (types[index] === '2P Lite 전환') ids2P.push(id);
    else if (crawlFlags[index] && crawlFlags[index][0]===true) idsCrawl.push(id);
    else ids3P.push(id);
  });
  const views3P = ensureRequestFilterViews_(ids3P, task3P);
  const viewsCrawl = ensureRequestFilterViews_(idsCrawl, crawl3P);
  const views2P = ensureRequestFilterViews_(ids2P, task2P);
  const base = ss.getUrl();
  const rich = ids.map(function(id,index){
    const builder = SpreadsheetApp.newRichTextValue().setText(id);
    if (!id) return [builder.build()];
    const use2P = types[index] === '2P Lite 전환', useCrawl = idsCrawl.indexOf(id)>=0;
    const sheet = use2P ? task2P : (useCrawl ? crawl3P : task3P);
    const fvid = (use2P ? views2P : (useCrawl ? viewsCrawl : views3P))[id];
    const gid = sheet.getSheetId();
    const url = base + '#gid=' + gid + (fvid ? '&fvid=' + fvid : '');
    return [builder.setLinkUrl(url).build()];
  });
  // Preserve archive links; a later worker also cleans up recreated managed views.
  for(let start=0;start<rowCount;){if(archived(start)){start++;continue;}let end=start+1;while(end<rowCount&&!archived(end))end++;requestSheet.getRange(start+2,headers['Request ID'],end-start,1).setRichTextValues(rich.slice(start,end));start=end;}
}

// 성공 응답 전에 신규 REQ에 최소 Task 시트 링크를 보장함.
function setNewRequestBaseLinkV1_(requestId,requestRow,requestType,crawlRequired) {
  const ss=SpreadsheetApp.getActiveSpreadsheet(),requestSheet=getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  const archiveHeaders=getHeaderMap_(requestSheet);if(archiveHeaders['Archive Status']&&String(requestSheet.getRange(requestRow,archiveHeaders['Archive Status']).getValue()).toUpperCase()==='ARCHIVED')return true;
  const taskSheet=requestType==='2P Lite 전환'
    ? getRequiredSheet_(APP_CONFIG.SHEETS.TASKS_2P)
    : (crawlRequired?getRequiredSheet_(APP_CONFIG.SHEETS.CRAWL_TASKS):getRequiredSheet_(APP_CONFIG.SHEETS.TASKS));
  const gid=taskSheet.getSheetId(),url=ss.getUrl()+'#gid='+gid;
  const rich=SpreadsheetApp.newRichTextValue().setText(String(requestId)).setLinkUrl(url).build();
  requestSheet.getRange(requestRow,1).setRichTextValue(rich);
  return true;
}

// 신규 요청 한 건만 필터 링크로 갱신함. 전체 View는 다시 그리지 않음.
function refreshNewRequestIdLinkV1_(requestId,requestRow,requestType,crawlRequired) {
  const ss=SpreadsheetApp.getActiveSpreadsheet(),requestSheet=getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  const archiveHeaders=getHeaderMap_(requestSheet);if(archiveHeaders['Archive Status']&&String(requestSheet.getRange(requestRow,archiveHeaders['Archive Status']).getValue()).toUpperCase()==='ARCHIVED')return true;
  const taskSheet=requestType==='2P Lite 전환'
    ? getRequiredSheet_(APP_CONFIG.SHEETS.TASKS_2P)
    : (crawlRequired?getRequiredSheet_(APP_CONFIG.SHEETS.CRAWL_TASKS):getRequiredSheet_(APP_CONFIG.SHEETS.TASKS));
  const gid=taskSheet.getSheetId();
  // 필터 보기가 확정된 뒤에만 링크를 기록함. 실패 시 일반 텍스트를 유지하고 호출자가 재시도함.
  const fvid=ensureRequestFilterViews_([String(requestId)],taskSheet)[String(requestId)];
  if(!fvid)throw new Error('신규 Request 필터 보기를 확인할 수 없습니다.');
  const url=ss.getUrl()+'#gid='+gid+'&fvid='+fvid;
  const rich=SpreadsheetApp.newRichTextValue().setText(String(requestId)).setLinkUrl(url).build();
  requestSheet.getRange(requestRow,1).setRichTextValue(rich);
  return true;
}
