// Explicit administrator maintenance for the retired per-VI timestamp.
function removeDeprecatedViCompletedAtColumns(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),removed=[];
  [APP_CONFIG.SHEETS.TASKS,APP_CONFIG.SHEETS.CRAWL_TASKS,APP_CONFIG.SHEETS.TASKS_2P].forEach(function(name){
    const sheet=ss.getSheetByName(name);if(!sheet)return;
    ['3P Completed At','2P Completed At'].forEach(function(header){
      const column=getHeaderMap_(sheet)[header];
      if(column){sheet.deleteColumn(column);removed.push(name+' · '+header);}
    });
  });
  ss.toast(removed.length?removed.join(', ')+'에서 컬럼을 제거했습니다.':'제거할 컬럼이 없습니다.','관리자',6);
  return {removed:removed};
}
function removeDeprecated3PCompletedAtColumns(){return removeDeprecatedViCompletedAtColumns();}
