// Barcode availability is request metadata, not a timed work stage.
// YES/NO are intentionally language-neutral persisted values; the UI translates them.
const BARCODE_PROVIDED_HEADER_V1_ = 'Barcode Provided';
const BARCODE_PROVIDED_VALUES_V1_ = Object.freeze(['YES', 'NO']);
const REMOVED_BARCODE_COLUMN_SLOTS_V1_ = Object.freeze([
  Object.freeze({legacy:'Barcode Assignee', reserved:'Reserved (No Shift) 1'}),
  Object.freeze({legacy:'Barcode Started At', reserved:'Reserved (No Shift) 2'}),
  Object.freeze({legacy:'Barcode Completed At', reserved:'Reserved (No Shift) 3'})
]);

function normalizeBarcodeProvidedV1_(value) {
  const normalized=String(value==null?'':value).trim().toUpperCase();
  return BARCODE_PROVIDED_VALUES_V1_.indexOf(normalized)>=0?normalized:'';
}

function isBarcodeProvidedKnownV1_(value) {
  return !!normalizeBarcodeProvidedV1_(value);
}

// Must be called while the shared document lock is held.
// The first valid answer is immutable so all sessions for a Request ID use one cohort.
function captureBarcodeProvidedV1_(requestSheet,requestRecord,requestHeaders,suppliedValue) {
  const column=requestHeaders&&requestHeaders[BARCODE_PROVIDED_HEADER_V1_];
  if(!column)throw Error('REQUESTS의 Barcode Provided 열이 없습니다. 관리자 전환 작업을 먼저 실행해 주세요.');
  const suppliedRaw=String(suppliedValue==null?'':suppliedValue).trim();
  const supplied=normalizeBarcodeProvidedV1_(suppliedRaw);
  if(suppliedRaw&&!supplied)throw Error('바코드 제공 여부는 YES 또는 NO만 입력할 수 있습니다.');
  const current=normalizeBarcodeProvidedV1_(requestRecord.values[column-1]);
  if(current){
    if(supplied&&supplied!==current)throw Error('바코드 제공 여부는 이미 기록되어 변경할 수 없습니다.');
    return current;
  }
  if(!supplied)throw Error('2P 작업 전에 바코드 제공 여부를 선택해 주세요.');
  requestSheet.getRange(requestRecord.row,column).setValue(supplied);
  requestRecord.values[column-1]=supplied;
  return supplied;
}

function barcodeWorkflowHeaderPositionsV1_(headers) {
  const positions=Object.create(null);
  headers.forEach(function(value,index){
    const name=String(value||'').trim();
    if(!name)return;
    if(!positions[name])positions[name]=[];
    positions[name].push(index+1);
  });
  return positions;
}

function barcodeWorkflowSpreadsheetV1_(spreadsheetId) {
  if(spreadsheetId)return SpreadsheetApp.openById(String(spreadsheetId));
  const spreadsheet=SpreadsheetApp.getActiveSpreadsheet();
  if(!spreadsheet)throw Error('대상 Spreadsheet를 찾을 수 없습니다. Spreadsheet ID를 지정해 주세요.');
  return spreadsheet;
}

function barcodeWorkflowPlanV1_(spreadsheet) {
  const sheet=spreadsheet.getSheetByName(APP_CONFIG.SHEETS.REQUESTS);
  if(!sheet)throw Error('필수 시트를 찾을 수 없습니다: '+APP_CONFIG.SHEETS.REQUESTS);
  const lastColumn=Math.max(sheet.getLastColumn(),1);
  const headers=sheet.getRange(1,1,1,lastColumn).getDisplayValues()[0];
  const positions=barcodeWorkflowHeaderPositionsV1_(headers);
  const providedColumns=positions[BARCODE_PROVIDED_HEADER_V1_]||[],statusColumns=positions['Barcode Status']||[];
  if(providedColumns.length>1)throw Error('REQUESTS에 중복 헤더가 있습니다: '+BARCODE_PROVIDED_HEADER_V1_);
  const providedColumn=providedColumns[0]||statusColumns[0]||0;
  const providedAction=providedColumns.length?'READY':providedColumn?'REUSE':'APPEND';
  // When a partial setup left both names behind, Barcode Provided is primary.
  // With no Provided column, the first Status slot is reused and any extras are auxiliary.
  const auxiliaryStatusColumns=providedColumns.length?statusColumns.slice():statusColumns.slice(1);
  const removed=[];
  auxiliaryStatusColumns.forEach(function(column){removed.push({header:'Barcode Status',column:column,action:'DELETE'});});
  REMOVED_BARCODE_COLUMN_SLOTS_V1_.forEach(function(item){
    [item.legacy,item.reserved].forEach(function(name){
      (positions[name]||[]).forEach(function(column){removed.push({header:name,column:column,action:'DELETE'});});
    });
  });
  const deleteColumns=removed.map(function(item){return item.column;});
  if(deleteColumns.length>=lastColumn)throw Error('REQUESTS에서 바코드 측정 열을 안전하게 삭제할 수 없습니다.');
  return {sheet:sheet,providedColumn:providedColumn,providedAction:providedAction,auxiliaryStatusColumns:auxiliaryStatusColumns,removed:removed,deleteColumns:deleteColumns,lastRow:sheet.getLastRow(),maxRows:sheet.getMaxRows()};
}

// Read-only migration preview.
function previewBarcodeProvidedWorkflowV1(spreadsheetId) {
  const plan=barcodeWorkflowPlanV1_(barcodeWorkflowSpreadsheetV1_(spreadsheetId));
  const currentProvidedColumn=plan.providedColumn||plan.sheet.getLastColumn()+1;
  const finalProvidedColumn=currentProvidedColumn-plan.deleteColumns.filter(function(column){return column<currentProvidedColumn;}).length;
  const result={
    ok:true,
    provided:{action:plan.providedAction,column:finalProvidedColumn,currentColumn:currentProvidedColumn},
    removed:plan.removed.map(function(item){return {header:item.header,action:item.action,column:item.column||null};})
  };
  console.log(JSON.stringify(result));
  return result;
}

function applyBarcodeProvidedValidationV1_(sheet,column) {
  if(sheet.getMaxRows()<2)return;
  const range=sheet.getRange(2,column,sheet.getMaxRows()-1,1);
  if(SpreadsheetApp.newDataValidation){
    range.setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(BARCODE_PROVIDED_VALUES_V1_.slice(),true).setAllowInvalid(false).build());
  }
  range.setHorizontalAlignment('center').setNumberFormat('@');
  sheet.setColumnWidth(column,118);
}

function barcodeSessionRetirementPlanV1_(spreadsheet) {
  const sheet=spreadsheet.getSheetByName('WORK_SESSIONS');
  if(!sheet)return {sheet:sheet,width:0,values:[],headers:{},active:[]};
  const width=Math.min(Math.max(sheet.getLastColumn(),10),13);
  const lastRow=Math.max(sheet.getLastRow(),1);
  const values=sheet.getRange(1,1,lastRow,width).getValues();
  const headers=values[0]||[];
  ['Stage','Status','Started At','Ended At','Duration Minutes','Updated By','Note'].forEach(function(name){
    const count=headers.filter(function(header){return String(header||'')===name;}).length;
    if(count!==1)throw Error('WORK_SESSIONS의 '+name+' 열이 없거나 중복되어 BARCODE 세션을 안전하게 종료할 수 없습니다.');
  });
  const h={};headers.forEach(function(name,index){if(name)h[String(name)]=index;});
  const active=[];
  values.slice(1).forEach(function(row,index){if(String(row[h.Stage]||'')==='BARCODE'&&String(row[h.Status]||'')==='ACTIVE')active.push({row:index+2,values:row});});
  return {sheet:sheet,width:width,values:values,headers:h,active:active};
}

// Preserve old rows as audit history, but no retired BARCODE session may stay active.
function retireActiveBarcodeSessionsV1_(plan) {
  plan=plan||barcodeSessionRetirementPlanV1_();
  if(!plan.sheet||!plan.active.length)return 0;
  const sheet=plan.sheet,h=plan.headers;
  const now=new Date(),user=typeof workSessionUser_==='function'?workSessionUser_():'',changes=[];
  plan.active.forEach(function(item){
    const row=item.values;
    const marker='바코드 작업단계 폐기로 자동 종료',oldNote=String(row[h.Note]||'').trim();
    const note=oldNote.indexOf(marker)>=0?oldNote:(oldNote?oldNote+' | '+marker:marker);
    const updates=[
      {column:h['Ended At']+1,value:now},
      {column:h['Duration Minutes']+1,value:row[h['Started At']] instanceof Date?Math.max(0,(now-row[h['Started At']])/60000):Math.max(0,Number(row[h['Duration Minutes']])||0)},
      {column:h['Updated By']+1,value:user},
      {column:h.Note+1,value:note}
    ];
    ['Start Completed VI','End Completed VI','Completed VI Delta'].forEach(function(name){if(h[name]!==undefined)updates.push({column:h[name]+1,value:''});});
    changes.push({row:item.row,updates:updates,statusColumn:h.Status+1});
  });
  // One row write includes Ended At, Duration, Status and all VI fields together.
  // This also supports the old 10-column BARCODE migration schema.
  changes.forEach(function(change){
    const values=sheet.getRange(change.row,1,1,sheet.getLastColumn()).getValues()[0];
    change.updates.forEach(function(update){values[update.column-1]=update.value;});
    values[change.statusColumn-1]='PAUSED';
    sheet.getRange(change.row,1,1,values.length).setValues([values]);
  });
  return changes.length;
}

function barcodeProvidedMigrationValuesV1_(plan) {
  const rowCount=Math.max(plan.lastRow-1,0),sheet=plan.sheet;
  if(!rowCount)return {changes:[],clearedLegacyValues:0,mergedLegacyAnswers:0};
  const primary=plan.providedColumn?sheet.getRange(2,plan.providedColumn,rowCount,1).getValues().map(function(row){return row[0];}):Array(rowCount).fill('');
  const output=primary.slice(),invalid=[];
  if(plan.providedAction==='READY')primary.forEach(function(value,index){
    if(String(value==null?'':value).trim()&&!normalizeBarcodeProvidedV1_(value))invalid.push(index+2);
  });
  if(invalid.length)throw Error('Barcode Provided에 YES/NO가 아닌 값이 있습니다: '+invalid.slice(0,10).join(', '));
  let clearedLegacyValues=0,mergedLegacyAnswers=0;
  if(plan.providedAction==='REUSE')primary.forEach(function(value,index){
    const normalized=normalizeBarcodeProvidedV1_(value);
    if(String(value==null?'':value).trim()&&!normalized)clearedLegacyValues++;
    output[index]=normalized;
  });
  plan.auxiliaryStatusColumns.forEach(function(column){
    sheet.getRange(2,column,rowCount,1).getValues().forEach(function(row,index){
      const raw=row[0],normalized=normalizeBarcodeProvidedV1_(raw);
      if(!normalized){if(String(raw==null?'':raw).trim())clearedLegacyValues++;return;}
      if(!normalizeBarcodeProvidedV1_(output[index])){output[index]=normalized;mergedLegacyAnswers++;}
    });
  });
  const changes=[];
  if(plan.providedAction!=='APPEND')output.forEach(function(value,index){
    const normalized=normalizeBarcodeProvidedV1_(value),current=String(primary[index]==null?'':primary[index]).trim();
    if(current!==normalized)changes.push({row:index+2,value:normalized});
  });
  return {changes:changes,clearedLegacyValues:clearedLegacyValues,mergedLegacyAnswers:mergedLegacyAnswers};
}

// One-time, idempotent production migration.
// Barcode Status is renamed in place and its incompatible status values are cleared.
// The three obsolete assignment/time columns are physically deleted. Barcode Provided
// is explicitly shown even when its former column had been hidden.
function installBarcodeProvidedWorkflowV1(spreadsheetId) {
  const spreadsheet=barcodeWorkflowSpreadsheetV1_(spreadsheetId);
  return withOpsDocumentLock_(function(){
    // Preflight both sheets before the first write so a malformed session schema
    // cannot leave REQUESTS halfway migrated.
    const plan=barcodeWorkflowPlanV1_(spreadsheet),sessionPlan=barcodeSessionRetirementPlanV1_(spreadsheet),valuesPlan=barcodeProvidedMigrationValuesV1_(plan),sheet=plan.sheet;
    let providedColumn=plan.providedColumn;
    if(plan.providedAction==='APPEND'){
      providedColumn=sheet.getLastColumn()+1;
      sheet.getRange(1,providedColumn).setValue(BARCODE_PROVIDED_HEADER_V1_);
    }else if(plan.providedAction==='REUSE'){
      sheet.getRange(1,providedColumn).setValue(BARCODE_PROVIDED_HEADER_V1_);
    }
    valuesPlan.changes.forEach(function(change){sheet.getRange(change.row,providedColumn).setValue(change.value);});
    const deleteColumns=plan.deleteColumns.slice().sort(function(a,b){return b-a;});
    deleteColumns.forEach(function(column){
      sheet.deleteColumn(column);
    });
    providedColumn-=deleteColumns.filter(function(column){return column<providedColumn;}).length;
    sheet.showColumns(providedColumn);
    sheet.getRange(1,providedColumn).setBackground('#24364B').setFontColor('#FFFFFF').setFontWeight('bold');
    applyBarcodeProvidedValidationV1_(sheet,providedColumn);

    const retiredSessions=retireActiveBarcodeSessionsV1_(sessionPlan);
    if(typeof bumpMyWorkSidebarRevisionV2_==='function')bumpMyWorkSidebarRevisionV2_([]);
    const result={ok:true,providedColumn:providedColumn,reusedExistingColumn:plan.providedAction==='REUSE',clearedLegacyValues:valuesPlan.clearedLegacyValues,mergedLegacyAnswers:valuesPlan.mergedLegacyAnswers,deletedColumns:deleteColumns.length,retiredActiveSessions:retiredSessions};
    console.log(JSON.stringify(result));
    return result;
  });
}

function verifyBarcodeProvidedWorkflowV1(spreadsheetId) {
  const spreadsheet=barcodeWorkflowSpreadsheetV1_(spreadsheetId),plan=barcodeWorkflowPlanV1_(spreadsheet);
  if(plan.providedAction!=='READY')throw Error('Barcode Provided 전환이 완료되지 않았습니다.');
  if(plan.removed.some(function(item){return item.action==='DELETE';}))throw Error('기존 바코드 담당자/시간측정 열이 남아 있습니다.');
  const column=plan.providedColumn,invalid=[];
  if(typeof plan.sheet.isColumnHiddenByUser==='function'&&plan.sheet.isColumnHiddenByUser(column))throw Error('Barcode Provided 열이 숨겨져 있습니다.');
  if(plan.lastRow>1)plan.sheet.getRange(2,column,plan.lastRow-1,1).getDisplayValues().forEach(function(row,index){
    if(String(row[0]||'').trim()&&!normalizeBarcodeProvidedV1_(row[0]))invalid.push(index+2);
  });
  if(invalid.length)throw Error('Barcode Provided 값 오류 행: '+invalid.slice(0,10).join(', '));
  const sessionPlan=barcodeSessionRetirementPlanV1_(spreadsheet);
  const activeBarcodeSessions=sessionPlan.active.length;
  if(activeBarcodeSessions)throw Error('종료되지 않은 BARCODE 세션이 '+activeBarcodeSessions+'건 있습니다.');
  return {ok:true,providedColumn:column,activeBarcodeSessions:0};
}
