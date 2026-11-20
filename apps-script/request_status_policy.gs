// Manual Status choices are distinct from canonical/reporting statuses (which include DONE).
function manualRequestStatuses_(){return ['TO DO','ONGOING','ON HOLD','DROP'];}
function requestStatusInputMessage_(){return 'DONE은 작업함의 정상 완료 처리로만 기록됩니다.';}
function requestStatusValidationRule_(){
  return SpreadsheetApp.newDataValidation().requireValueInList(manualRequestStatuses_(),true)
    .setAllowInvalid(false).setHelpText(requestStatusInputMessage_()).build();
}
function requestStatusApiRule_(){return {condition:{type:'ONE_OF_LIST',values:manualRequestStatuses_().map(v=>({userEnteredValue:v}))},strict:true,showCustomUi:true,inputMessage:requestStatusInputMessage_()};}
function requestStatusGrid_(sheet,row,column,count){return {sheetId:sheet.getSheetId(),startRowIndex:row-1,endRowIndex:row-1+count,startColumnIndex:column-1,endColumnIndex:column};}
function applyRequestStatusRange_(sheet,row,column,count){
  if(count<1)return;
  Sheets.Spreadsheets.batchUpdate({requests:[{setDataValidation:{range:requestStatusGrid_(sheet,row,column,count),rule:requestStatusApiRule_()}}]},SpreadsheetApp.getActiveSpreadsheet().getId());
}
// Does not read/rewrite cell values, formatting, protections or conditional rules.
function applyRequestStatusValidationV1(){
  assertDashboardAdmin_();
  const result=withInstalledRequestStatusLockV4_(function(){
    const sheet=getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS),headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0];
    const matches=[];headers.forEach((value,i)=>{if(String(value).trim()==='Status')matches.push(i+1);});
    if(matches.length!==1)throw Error('REQUESTS에 Status 헤더가 정확히 하나 있어야 합니다.');
    const rows=Math.max(0,sheet.getLastRow()-1),validationRows=Math.max(0,sheet.getMaxRows()-1);
    applyRequestStatusRange_(sheet,2,matches[0],validationRows);
    return {rows:rows,validationRows:validationRows};
  });
  console.log('REQUESTS Status validation 적용 완료: '+result.rows+'행 (빈 행 포함 '+result.validationRows+'행). 기존 값은 변경하지 않았습니다.');
  return result;
}
function assertManualRequestStatus_(status){
  const normalized=assertAllowedStatus_(status);
  if(normalized===APP_CONFIG.STATUS.DONE)throw Error(requestStatusInputMessage_());
  return normalized;
}
// Compose into the caller's existing atomic commit. Never leave a separate
// clear-validation operation committed if a later value write fails.
function requestStatusSystemWrites_(sheet,row,column,values){
  const range=requestStatusGrid_(sheet,row,column,values.length);
  return [{setDataValidation:{range:range}},
    {updateCells:{start:{sheetId:sheet.getSheetId(),rowIndex:row-1,columnIndex:column-1},rows:values.map(v=>({values:[{userEnteredValue:{stringValue:String(v)}}]})),fields:'userEnteredValue'}},
    {setDataValidation:{range:range,rule:requestStatusApiRule_()}}];
}
function writeRequestStatuses_(sheet,row,column,values){
  if(values.indexOf(APP_CONFIG.STATUS.DONE)<0){sheet.getRange(row,column,values.length,1).setValues(values.map(v=>[v]));return;}
  SpreadsheetApp.flush();
  Sheets.Spreadsheets.batchUpdate({requests:requestStatusSystemWrites_(sheet,row,column,values)},SpreadsheetApp.getActiveSpreadsheet().getId());
}
// Existing edit handlers call this under their existing lock. A multi-cell
// edit has no oldValues: never guess and erase historical DONE records.
function guardManualRequestDone_(e,headers){
  const range=e.range,column=headers.Status;
  if(!column||range.getRow()<2||range.getColumn()>column||range.getLastColumn()<column)return false;
  const sheet=range.getSheet(),values=sheet.getRange(range.getRow(),column,range.getNumRows(),1).getValues();
  if(!values.some(r=>normalizeRequestStatusV2_(r[0])===APP_CONFIG.STATUS.DONE))return false;
  const single=range.getNumRows()===1&&range.getNumColumns()===1;
  if(single&&normalizeRequestStatusV2_(e.value)!==APP_CONFIG.STATUS.DONE)return true; // stale event; don't undo a system commit
  if(single&&normalizeRequestStatusV2_(e.oldValue)!==APP_CONFIG.STATUS.DONE){
    const old=e.oldValue===undefined?'':String(e.oldValue);
    if(!old||manualRequestStatuses_().indexOf(normalizeRequestStatusV2_(old))>=0)sheet.getRange(range.getRow(),column).setValue(old?normalizeRequestStatusV2_(old):'');
  }
  // Simple onEdit cannot call the authorized Sheets service.
  sheet.getRange(range.getRow(),column,range.getNumRows(),1).setDataValidation(requestStatusValidationRule_());
  (e.source||SpreadsheetApp.getActiveSpreadsheet()).toast(requestStatusInputMessage_()+(single?'':' 여러 셀 붙여넣기의 이전 값은 확인할 수 없습니다. 관리자에게 확인을 요청하세요.'),'Status 입력 제한',10);
  return true;
}
