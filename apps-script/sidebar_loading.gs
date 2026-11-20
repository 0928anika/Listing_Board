// Request records for session commands; never read the entire REQUESTS body.
function readSessionRequestRecords_(ids,table) {
  if(!table){const sheet=getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);table={sheet:sheet,headers:getHeaderMap_(sheet),byId:Object.create(null),rows:[]};}
  const sheet=table.sheet,last=sheet.getLastRow(),width=sheet.getLastColumn();
  Array.from(new Set(ids.map(String))).forEach(function(id){
    if(table.byId[id])return;
    const found=last<2?[]:sheet.getRange(2,table.headers['Request ID'],last-1,1).createTextFinder(id).matchEntireCell(true).matchCase(true).useRegularExpression(false).findAll();
    if(found.length!==1)throw Error('요청 ID가 없거나 중복되었습니다: '+id);
    const row=found[0].getRow(),values=sheet.getRange(row,1,1,width).getValues()[0];
    if(String(values[table.headers['Request ID']-1])!==id)throw Error('요청 행이 변경되었습니다. 다시 시도하세요.');
    table.byId[id]=[{row:row,values:values}];table.rows.push(values);
  });return table;
}
// Serialize the already-loaded session/progress snapshot while still holding the mutation lock.
function sidebarSessionResult_(result,payload,requestTable,viContext) {
  if(!payload.sidebarResult)return result;
  const worker=String(payload.worker),counts=[];
  Object.keys(viContext.sources).forEach(function(source){const table=viContext.sources[source];
    table.loadedRequestIds.forEach(function(id){const m=table.index[id]||emptyTaskAggregate_(),w=m.workers[worker]||{},record=requestTable.byId[id][0];
      const v=function(name){return opsValue_(record.values,requestTable.headers,name);};
      counts.push({requestId:id,assigned3P:w.assigned3P||0,done3P:w.done3P||0,assigned2P:w.assigned2P||0,done2P:w.done2P||0,status:normalizeRequestStatusV2_(v('Status'))||APP_CONFIG.STATUS.TO_DO,totalVi:Number(v('Total VI'))||0,completed3P:Number(v('3P Completed VI'))||0,completed2P:Number(v('2P Completed VI'))||0,barcodeProvided:normalizeBarcodeProvidedV1_(v('Barcode Provided'))});
    });
  });
  result.progress=counts;if(counts.length)result.revision=getMyWorkSidebarRevision();return result;
}
// Initial sidebar needs only currently active sessions; never scan a worker's paused history.
function sidebarWorkerSessions_(worker) {
  return getWorkSessionSnapshot_(true,sidebarWorkerSessionRows_(worker));
}
function sidebarWorkerSessionRows_(worker) {
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('WORK_SESSIONS'),last=sheet?sheet.getLastRow():0;
  if(last<2)return [];
  const positions=sheet.getRange(2,5,last-1,1).createTextFinder('ACTIVE').matchEntireCell(true).matchCase(true).useRegularExpression(false).findAll().map(function(cell){return cell.getRow();}).sort(function(a,b){return a-b;}),groups=[],rows=[];
  positions.forEach(function(row){const g=groups[groups.length-1];if(g&&g.start+g.count===row)g.count++;else groups.push({start:row,count:1});});
  groups.forEach(function(g){sheet.getRange(g.start,1,g.count,10).getValues().forEach(function(row){if(String(row[4])==='ACTIVE'&&String(row[1])===worker)rows.push(row);});});
  return rows;
}
