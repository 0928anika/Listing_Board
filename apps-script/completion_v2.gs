// TASKS에서 리더 배정 범위와 실제 단계별 작업자를 분리한 운영 구조임.
const TASK_HEADERS_STAGE_V3 = Object.freeze([
  'Request ID','Task/VI ID','품번','등록상품명','사이즈','색상','판매가','정상가','Brand',
  'Memo','Assigned Pool','3P Assignee','3P Completed','2P Assignee','2P Completed',
  '3P Completed At','2P Completed At','Completed At','Work Status','Node','Notes'
]);

// 기존 TASKS 데이터를 단계별 담당자 구조로 안전하게 이전함.
function setupTaskStageAssigneesV3(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const taskSheet=getRequiredSheet_(APP_CONFIG.SHEETS.TASKS);
  const requestSheet=getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  const oldCols=taskSheet.getLastColumn();
  const oldHeaders=taskSheet.getRange(1,1,1,oldCols).getDisplayValues()[0];
  const oldMap=oldHeaders.reduce(function(m,h,i){if(h)m[h]=i;return m;},{});
  const oldRows=taskSheet.getLastRow()<2?[]:taskSheet.getRange(2,1,taskSheet.getLastRow()-1,oldCols).getValues()
    .filter(function(row){return String(row[oldMap['Request ID']]||'').trim();});
  const rh=getHeaderMap_(requestSheet),requestInfo={};
  getDataRows_(requestSheet).forEach(function(row){
    requestInfo[String(row[rh['Request ID']-1])]={
      type:String(row[rh['Request Type']-1]||'3P Listing'),
      pool:String(row[rh['Active Assignees']-1]||'')
    };
  });
  const rows=oldRows.map(function(row){
    function val(name){return typeof oldMap[name]==='undefined'?'':row[oldMap[name]];}
    const id=String(val('Request ID')),info=requestInfo[id]||{type:'3P Listing',pool:''};
    const has3=info.type!=='2P Lite 전환',has2=info.type!=='3P Listing';
    const done3=val('3P Completed')===true||val('3P Listing 완료')===true;
    const done2=val('2P Completed')===true||val('2P Lite 완료')===true;
    const legacyAssignee=String(val('Active Assignees')||'');
    const time=val('Completed At');
    const finalDone=(has3?done3:true)&&(has2?done2:true);
    return [id,val('Task/VI ID'),val('품번'),val('등록상품명'),val('사이즈'),val('색상'),
      val('판매가'),val('정상가'),val('Brand'),typeof oldMap['Memo']==='undefined'?val('메모'):val('Memo'),info.pool,
      has3?(String(val('3P Assignee')||'')||(done3?legacyAssignee:'')):'N/A',has3?done3:'N/A',
      has2?(String(val('2P Assignee')||'')||(done2?legacyAssignee:'')):'N/A',has2?done2:'N/A',
      val('3P Completed At')||(done3?time:''),val('2P Completed At')||(done2?time:''),
      finalDone?time:'',finalDone?'작업 완료':(done3?'3P Listing 완료':(done2?'2P Lite 완료':'작업 대기')),
      val('Node'),val('Notes')||val('Notes (Placeholder)')];
  });
  taskSheet.getDataRange().clearDataValidations();taskSheet.clear();
  if(taskSheet.getMaxColumns()<TASK_HEADERS_STAGE_V3.length)taskSheet.insertColumnsAfter(taskSheet.getMaxColumns(),TASK_HEADERS_STAGE_V3.length-taskSheet.getMaxColumns());
  taskSheet.getRange(1,1,1,TASK_HEADERS_STAGE_V3.length).setValues([TASK_HEADERS_STAGE_V3]);
  if(rows.length)taskSheet.getRange(2,1,rows.length,TASK_HEADERS_STAGE_V3.length).setValues(rows);
  formatTaskStageV3_(taskSheet,requestInfo);
  Object.keys(requestInfo).forEach(syncRequestStageV3_);
  ss.toast('TASKS를 단계별 담당자 구조로 전환했습니다.','Listing Ops',5);
}

// 적용 가능한 단계만 작업자 드롭다운과 체크박스를 활성화함.
function formatTaskStageV3_(sheet,requestInfo){
  sheet.setFrozenRows(1);sheet.setFrozenColumns(2);sheet.setHiddenGridlines(true);
  const h=getHeaderMap_(sheet);
  sheet.getRange(1,1,1,TASK_HEADERS_STAGE_V3.length).setBackground('#263650').setFontColor('#FFF').setFontWeight('bold').setWrap(true);
  const cfg=getRequiredSheet_(APP_CONFIG.SHEETS.CONFIG);
  const workerRule=SpreadsheetApp.newDataValidation().requireValueInRange(cfg.getRange('C2:C'),true).setAllowInvalid(false).build();
  if(sheet.getLastRow()>=2){
    const ids=sheet.getRange(2,1,sheet.getLastRow()-1,1).getDisplayValues();
    ids.forEach(function(row,i){
      const rn=i+2,type=(requestInfo[String(row[0])]||{type:'3P Listing'}).type;
      if(type!=='2P Lite 전환'){
        sheet.getRange(rn,h['3P Assignee']).setDataValidation(workerRule).setBackground('#EAF4FF');
        sheet.getRange(rn,h['3P Completed']).insertCheckboxes().setBackground('#EAF4FF');
      }else sheet.getRange(rn,h['3P Assignee'],1,2).setValues([['N/A','N/A']]).setBackground('#F2F4F7').setFontColor('#9AA3AF');
      if(type!=='3P Listing'){
        sheet.getRange(rn,h['2P Assignee']).setDataValidation(workerRule).setBackground('#FFF3E8');
        sheet.getRange(rn,h['2P Completed']).insertCheckboxes().setBackground('#FFF3E8');
      }else sheet.getRange(rn,h['2P Assignee'],1,2).setValues([['N/A','N/A']]).setBackground('#F2F4F7').setFontColor('#9AA3AF');
    });
  }
  ['3P Completed At','2P Completed At','Completed At'].forEach(function(name){
    if(h[name])sheet.getRange(2,h[name],Math.max(sheet.getMaxRows()-1,1),1).setNumberFormat('yyyy-mm-dd hh:mm');
  });
  [145,190,110,220,90,100,100,100,105,220,180,115,105,115,105,145,145,145,125,100,200]
    .forEach(function(w,i){sheet.setColumnWidth(i+1,w);});
}

// 단계 담당자 또는 완료 체크 변경 시 자동관리 컬럼과 요청 지표를 갱신함.
// 별도 권한 없이 즉시 체크 변경을 처리하는 단순 트리거임.
function onEdit(e){handleTaskRouterV7_(e);}

// Space 일괄 체크 뒤 선택 영역을 이동하면 누락된 시각을 한 번 더 보정함.
function onSelectionChange(e){ 
  // Retired automatic handler: deliberately no reads or writes.
 }

function handleTaskStageEdit(e){return handleOperationalTaskEdit_(e);}

// 단계별 완료 시각과 최종 상태를 계산함.
function applyTaskStageV3_(sheet,row,requestId){return handleOperationalTaskEdit_({range:sheet.getRange(row,1,1,sheet.getLastColumn())});}

// 단계별 체크 수를 REQUESTS의 진행률과 완료 시각에 반영함.
function syncRequestStageV3_(requestId){return withOpsDocumentLock_(function(){return syncRequestSummaries_([requestId]);});}


// 현재 선택한 3P/2P 체크박스 범위를 한 번에 완료 처리함.
function completeSelectedStageCheckboxes(){setSelectedStageCheckboxes_(true);}

// 현재 선택한 3P/2P 체크박스 범위를 한 번에 완료 해제함.
function clearSelectedStageCheckboxes(){setSelectedStageCheckboxes_(false);}

// 선택 범위에서 실제 체크박스 셀만 변경하고 시간·진행률을 함께 동기화함.
function setSelectedStageCheckboxes_(checked){return withOpsDocumentLock_(function(){
 const sheet=SpreadsheetApp.getActiveSheet(),range=sheet.getActiveRange(),h=getHeaderMap_(sheet),a=taskAdapter_(h),start=Math.max(range.getRow(),2),end=range.getLastRow();
 const cols=[a.d3,a.d2].filter(function(c){return c>=0&&c+1>=range.getColumn()&&c+1<=range.getLastColumn();});
 if(!cols.length||end<start)throw Error('완료 체크박스 범위를 선택하세요.');
 const rows=sheet.getRange(start,1,end-start+1,sheet.getLastColumn()).getValues(),patches=[];
 rows.forEach(function(row,i){if(!row[a.id]||(checked&&!isViValidV3_(row,h)))return;cols.forEach(function(c){if(typeof row[c]==='boolean'){
   if(row[c]!==checked){patches.push({row:start+i,col:c+1,value:checked===true});row[c]=checked===true;}
 }});});
 writeOpsPatches_(sheet,patches);
 return handleOperationalTaskEdit_({range:range},{sheet:sheet,h:h,start:start,count:end-start+1,rows:rows,patches:[]});
});}


// 설치형 편집 트리거가 호출하는 직렬 처리 래퍼임.
function installedTaskStageEdit(e){ 
  // Retired automatic handler: deliberately no reads or writes.
 }

// 체크 상태와 단계별 완료 시각이 어긋난 행을 일괄 보정함.
function repairMissingStageTimestamps(e){ 
  // Retired automatic handler: deliberately no reads or writes.
 }

// 설치형 편집 트리거와 1분 주기 누락 보정 트리거를 한 번만 설치함.
function installReliableTaskEditTriggers(){return configureProductionHardening();}


// 화면에서 불필요한 식별자와 중복 완료 시각을 정리한 TASKS V4 구조임.
const TASK_HEADERS_LAYOUT_V4=Object.freeze([
  'Request ID','품번','등록상품명','사이즈','색상','판매가','정상가','Brand','Memo','Assigned Pool',
  '3P Assignee','3P Completed','2P Assignee','2P Completed',
  '3P Completed At','2P Completed At','Work Status','Node','Notes','Internal Task ID'
]);

function setupTaskLayoutV4(){
  const sheet=getRequiredSheet_(APP_CONFIG.SHEETS.TASKS),cols=sheet.getLastColumn();
  const headers=sheet.getRange(1,1,1,cols).getDisplayValues()[0];
  const map=headers.reduce(function(m,h,i){if(h)m[h]=i;return m;},{});
  const rows=sheet.getLastRow()<2?[]:sheet.getRange(2,1,sheet.getLastRow()-1,cols).getValues()
    .filter(function(r){return String(r[map['Request ID']]||'').trim();});
  function value(row,name){return typeof map[name]==='undefined'?'':row[map[name]];}
  const migrated=rows.map(function(row,index){
    return [value(row,'Request ID'),value(row,'품번'),value(row,'등록상품명'),value(row,'사이즈'),value(row,'색상'),
      value(row,'판매가'),value(row,'정상가'),value(row,'Brand'),typeof map['Memo']==='undefined'?value(row,'메모'):value(row,'Memo'),value(row,'Assigned Pool'),
      value(row,'3P Assignee'),value(row,'3P Completed'),value(row,'2P Assignee'),value(row,'2P Completed'),
      value(row,'3P Completed At'),value(row,'2P Completed At'),value(row,'Work Status'),value(row,'Node'),value(row,'Notes'),
      value(row,'Internal Task ID')||value(row,'Task/VI ID')||String(value(row,'Request ID'))+'-ROW-'+String(index+1).padStart(3,'0')];
  });
  sheet.getDataRange().clearDataValidations();sheet.clear();
  if(sheet.getMaxColumns()<TASK_HEADERS_LAYOUT_V4.length)sheet.insertColumnsAfter(sheet.getMaxColumns(),TASK_HEADERS_LAYOUT_V4.length-sheet.getMaxColumns());
  sheet.getRange(1,1,1,TASK_HEADERS_LAYOUT_V4.length).setValues([TASK_HEADERS_LAYOUT_V4]);
  if(migrated.length)sheet.getRange(2,1,migrated.length,TASK_HEADERS_LAYOUT_V4.length).setValues(migrated);
  formatTaskLayoutV4_(sheet);
  const layoutHeaders=getHeaderMap_(sheet);
  if(layoutHeaders['Internal Task ID'])sheet.hideColumns(layoutHeaders['Internal Task ID']);
  sheet.setFrozenColumns(3);
  repairTaskLayoutV6Timestamps();
  SpreadsheetApp.getActive().toast('TASKS 컬럼을 정리하고 등록상품명까지 고정했습니다.','Listing Ops',5);
}

function formatTaskLayoutV4_(sheet){
  sheet.setFrozenRows(1);sheet.setHiddenGridlines(true);
  const h=getHeaderMap_(sheet);
  sheet.getRange(1,1,1,TASK_HEADERS_LAYOUT_V4.length).setBackground('#263650').setFontColor('#FFF').setFontWeight('bold').setWrap(true);
  const requestSheet=getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS),rh=getHeaderMap_(requestSheet),types={};
  getDataRows_(requestSheet).forEach(function(r){types[String(r[rh['Request ID']-1])]=String(r[rh['Request Type']-1]||'3P Listing');});
  const cfg=getRequiredSheet_(APP_CONFIG.SHEETS.CONFIG);
  const workerRule=SpreadsheetApp.newDataValidation().requireValueInRange(cfg.getRange('C2:C'),true).setAllowInvalid(false).build();
  if(sheet.getLastRow()>=2){
    const ids=sheet.getRange(2,1,sheet.getLastRow()-1,1).getDisplayValues();
    ids.forEach(function(r,i){
      const row=i+2,type=types[String(r[0])]||'3P Listing';
      if(type!=='2P Lite 전환'){sheet.getRange(row,h['3P Assignee']).setDataValidation(workerRule).setBackground('#EAF4FF');sheet.getRange(row,h['3P Completed']).insertCheckboxes().setBackground('#EAF4FF');}
      else sheet.getRange(row,h['3P Assignee'],1,2).setValues([['N/A','N/A']]).setBackground('#F2F4F7').setFontColor('#9AA3AF');
      if(type!=='3P Listing'){sheet.getRange(row,h['2P Assignee']).setDataValidation(workerRule).setBackground('#FFF3E8');sheet.getRange(row,h['2P Completed']).insertCheckboxes().setBackground('#FFF3E8');}
      else sheet.getRange(row,h['2P Assignee'],1,2).setValues([['N/A','N/A']]).setBackground('#F2F4F7').setFontColor('#9AA3AF');
    });
  }
  ['3P Completed At','2P Completed At'].forEach(function(name){
    if(h[name])sheet.getRange(2,h[name],Math.max(sheet.getMaxRows()-1,1),1).setNumberFormat('yyyy-mm-dd hh:mm');
  });
  [145,110,220,90,100,100,100,105,220,180,115,105,115,105,145,145,125,100,200,150]
    .forEach(function(w,i){sheet.setColumnWidth(i+1,w);});
}

function handleTaskLayoutV4Edit(e){return handleOperationalTaskEdit_(e);}

function syncRequestLayoutV4_(requestId){return withOpsDocumentLock_(function(){return syncRequestSummaries_([requestId]);});}

function repairTaskLayoutV4Timestamps(){ 
  // Manual compatibility only; VI timestamps are deprecated.
 return rebuildRequestSummaries(); }


// TASKS V6: 중복 컬럼(Node, Notes, Assigned Pool)을 제거한 운영 레이아웃임.
const TASK_HEADERS_LAYOUT_V6 = [
  'Request ID', '품번', '등록상품명', '사이즈', '색상', '판매가', '정상가', 'Brand',
  'Memo', '3P Assignee', '3P Completed', '2P Assignee', '2P Completed',
  'VI Status', 'Invalid Reason', 'Work Status', 'Internal Task ID'
];

// 기존 TASKS 데이터를 헤더 기준으로 V6 순서로 옮기고 시스템 ID를 숨김.
function setupTaskLayoutV6() {
  const sheet = getRequiredSheet_(APP_CONFIG.SHEETS.TASKS);
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const oldLastColumn = Math.max(sheet.getLastColumn(), 1);
  const oldValues = sheet.getRange(1, 1, lastRow, oldLastColumn).getValues();
  const oldHeaders = oldValues[0].map(function(v) { return String(v || '').trim(); });
  const index = {};
  oldHeaders.forEach(function(header, i) { if (header) index[header] = i; });
  const migrated = oldValues.slice(1).map(function(row) {
    return TASK_HEADERS_LAYOUT_V6.map(function(header) {
      const sourceIndex=header==='Memo'&&index[header]===undefined?index['메모']:index[header];
      return sourceIndex === undefined ? '' : row[sourceIndex];
    });
  });

  if (sheet.getMaxColumns() < TASK_HEADERS_LAYOUT_V6.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), TASK_HEADERS_LAYOUT_V6.length - sheet.getMaxColumns());
  }
  sheet.showColumns(1, Math.min(sheet.getMaxColumns(), TASK_HEADERS_LAYOUT_V6.length));
  sheet.getRange(1, 1, lastRow, Math.max(oldLastColumn, TASK_HEADERS_LAYOUT_V6.length))
    .clearContent().clearDataValidations();
  sheet.getRange(1, 1, 1, TASK_HEADERS_LAYOUT_V6.length).setValues([TASK_HEADERS_LAYOUT_V6]);
  if (migrated.length) sheet.getRange(2, 1, migrated.length, TASK_HEADERS_LAYOUT_V6.length).setValues(migrated);

  sheet.getRange(1, 1, 1, TASK_HEADERS_LAYOUT_V6.length).setBackground('#263650').setFontColor('#FFFFFF')
    .setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(3);

  const rows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2, 6, rows, 2).setNumberFormat('#,##0');
  const headers = getHeaderMap_(sheet);
  if(headers['Memo'])sheet.setColumnWidth(headers['Memo'],220);
  if (sheet.getLastRow() >= 2) {
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, TASK_HEADERS_LAYOUT_V6.length).getValues();
    data.forEach(function(row, i) {
      const rn = i + 2;
      const assignee3=row[headers['3P Assignee']-1],completed3=row[headers['3P Completed']-1];
      const assignee2=row[headers['2P Assignee']-1],completed2=row[headers['2P Completed']-1];
      if (String(assignee3) !== 'N/A') sheet.getRange(rn, headers['3P Completed']).insertCheckboxes().setValue(completed3 === true);
      else sheet.getRange(rn, headers['3P Completed']).clearDataValidations().setValue('N/A');
      if (String(assignee2) !== 'N/A') sheet.getRange(rn, headers['2P Completed']).insertCheckboxes().setValue(completed2 === true);
      else sheet.getRange(rn, headers['2P Completed']).clearDataValidations().setValue('N/A');
    });
  }
  repairTaskLayoutV6Timestamps();
  SpreadsheetApp.getActive().toast('TASKS 중복 컬럼을 제거했습니다.', 'Listing Ops', 4);
}

// TASKS V6에서 담당자·체크박스 변경을 처리하고 단계별 완료 시각을 고정 기록함.
function handleTaskLayoutV6Edit(e){return handleOperationalTaskEdit_(e);}

// TASKS V6 기준으로 요청별 완료 수량과 상태를 다시 집계함.
function syncRequestLayoutV6_(requestId){return withOpsDocumentLock_(function(){return syncRequestSummaries_([requestId]);});}

// 대량 체크 시 누락된 단계별 완료 시각과 요청 집계를 보완함.
function repairTaskLayoutV6Timestamps(){ 
  // Manual compatibility only; VI timestamps are deprecated.
 return rebuildRequestSummaries(); }


// 편집된 작업 시트 유형에 맞는 완료 처리 함수를 호출함.
function handleTaskRouterV7_(e){
 if(!e||!e.range)return;
 const name=e.range.getSheet().getName();
 if([APP_CONFIG.SHEETS.TASKS,APP_CONFIG.SHEETS.CRAWL_TASKS,APP_CONFIG.SHEETS.TASKS_2P].indexOf(name)>=0){
   if(e.range.getLastRow()<2)return;
   let needsStart=false;
   const result=withOpsDocumentLock_(function(){
     const context=readOperationalEditContext_(e);
     if(!guardTaskCheckboxStartV2_(e,context)){needsStart=context.needsStart;return;}
     const result=handleOperationalTaskEdit_(e,context);
     if(context.rejected)SpreadsheetApp.getActive().toast(context.rejected+'건은 담당자/VALID/작업 시작 조건을 확인하세요. 완료 체크를 해제했습니다.','완료 처리',8);
     return result;
   });
   // A dialog releases locks. Re-read and revalidate after the start confirmation.
   if(needsStart){
     const accepted=guardTaskCheckboxStartLegacy_(e);
     return withOpsDocumentLock_(function(){
     const context=readOperationalEditContext_(e);
     context.noPrompt=true;
     if(accepted)guardTaskCheckboxStartV2_(e,context);
     const result=handleOperationalTaskEdit_(e,context);
     if(context.rejected)SpreadsheetApp.getActive().toast('작업 세션이 변경되어 완료 체크를 해제했습니다.','완료 처리',8);
     return result;
     });
   }
   return result;
 }
 if(name===APP_CONFIG.SHEETS.REQUESTS){
   const requestHeaders=getHeaderMap_(e.range.getSheet()),activeColumn=requestHeaders['Active Assignees'],statusColumn=requestHeaders['Status'];
   // A simple onEdit is limited to 30 seconds.  Assignment and cancellation
   // can touch many Task rows, so the authorized installable handler owns them.
   if(statusColumn&&requestStatusEditNeedsInstalledV4_(e,requestHeaders)){
     (e.source||SpreadsheetApp.getActiveSpreadsheet()).toast('요청 상태 변경 처리 중…','Listing Ops',30);
     return;
   }
   if(activeColumn&&e.range.getColumn()<=activeColumn&&e.range.getLastColumn()>=activeColumn){
     (e.source||SpreadsheetApp.getActiveSpreadsheet()).toast(e.range.getNumRows()===1&&e.range.getNumColumns()===1?'자동분배 처리 중…':'Active Assignees는 요청별로 한 셀씩 수정해주세요.','Listing Ops',e.range.getNumRows()===1&&e.range.getNumColumns()===1?30:8);
     return;
   }
   return withOpsDocumentLock_(function(){if(guardManualRequestDone_(e,requestHeaders))return;handleRequestAssignmentDateV11_(e);handleAllocationManagementEditV2_(e);});
 }
}

// 2P_TASKS의 담당자와 완료 체크 변경을 처리함.
function handle2PTaskEditV7_(e){return handleOperationalTaskEdit_(e);}

// 2P Lite 단독 요청의 완료 수량과 상태를 REQUESTS에 반영함.
function sync2PRequestV7_(requestId){return withOpsDocumentLock_(function(){return syncRequestSummaries_([requestId]);});}

// 대량 체크 시 누락된 2P 완료 시각을 보완함.
function repair2PTaskTimestampsV7_(){ 
  // Manual compatibility only; VI timestamps are deprecated.
 return rebuildRequestSummaries(); }


// REQUESTS에서 담당자가 처음 지정되면 Started At 날짜를 고정 기록함.
function handleRequestAssignmentDateV11_(e) {
  const sheet = e.range.getSheet();
  if (e.range.getRow() < 2) return;
  const headers = getHeaderMap_(sheet);
  const firstRow = e.range.getRow(), count = e.range.getNumRows();
  const firstCol = e.range.getColumn(), lastCol = e.range.getLastColumn();
  const now = new Date();

  // 담당자 배정은 실제 착수 시간이 아니므로 Started At을 기록하지 않음.
  if (headers['Status'] && firstCol <= headers['Status'] && lastCol >= headers['Status']) {
    const statuses = sheet.getRange(firstRow, headers['Status'], count, 1).getDisplayValues();
    statuses.forEach(function(row, index) {
      const nextStatus = normalizeRequestStatusV2_(row[0]);
      if (!nextStatus || nextStatus === APP_CONFIG.STATUS.DONE) return; // native edits cannot complete requests
      const previousStatus = count === 1 && e.oldValue ? normalizeRequestStatusV2_(e.oldValue) : '';
      const requestId = String(sheet.getRange(firstRow + index, headers['Request ID']).getValue() || '');
      if(requestId)markRequestReconcileNeeded_([requestId],{tasks:true,sessions:true},'request-status');
      applyRequestTimingV12_(sheet, firstRow + index, headers, nextStatus, previousStatus, now);
      if(requestId)reconcileRequestSummary_(requestId,{preserveStatus:true,reason:'request-status'});
      if (requestId) logActivity_(requestId, 'STATUS_CHANGED', previousStatus, nextStatus);
    });
  }
}
