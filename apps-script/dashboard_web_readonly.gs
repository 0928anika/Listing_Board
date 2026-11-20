// Web dashboard adapter. No active-spreadsheet fallback is added to shared helpers.
const DASHBOARD_WEB_SOURCE_KEY_V1 = 'DASHBOARD_WEB_SOURCE_SPREADSHEET_ID_V1';

// Public editor entry; reject web execution before touching configuration.
function configureReadOnlyDashboardWeb() {
  if (!SpreadsheetApp.getActiveSpreadsheet()) {
    throw Error('운영 스프레드시트에서 연 Apps Script 편집기에서만 실행할 수 있습니다.');
  }
  return configureReadOnlyDashboardWeb_();
}

// Run once from the bound project's editor as a dashboard administrator.
// Private suffix prevents invocation through google.script.run.
function configureReadOnlyDashboardWeb_() {
  assertDashboardAdmin_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw Error('운영 스프레드시트에서 연 Apps Script 편집기에서 실행해 주세요.');
  if (!ss.getSheetByName(APP_CONFIG.SHEETS.REQUESTS)) throw Error('REQUESTS 시트를 찾을 수 없습니다.');
  PropertiesService.getScriptProperties().setProperty(DASHBOARD_WEB_SOURCE_KEY_V1, ss.getId());
  console.log('조회 전용 대시보드의 시트 연결을 저장했습니다. 데이터와 트리거는 변경하지 않았습니다.');
}

// Access is restricted to the company domain by the web-app deployment.
// Web viewers do not use the sheet popup approval list.
function showReadOnlyDashboardWeb_() {
  return HtmlService.createHtmlOutputFromFile('OperationsDashboardWeb')
    .setTitle('Operations Dashboard · 조회 전용')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getReadOnlyDashboardWebData(payload) {
  const id = PropertiesService.getScriptProperties().getProperty(DASHBOARD_WEB_SOURCE_KEY_V1);
  if (!id) throw Error('웹 대시보드의 운영 시트 연결이 설정되지 않았습니다. 관리자에게 문의해 주세요.');
  // ID is server configuration, never a client-supplied spreadsheet selector.
  const ss = SpreadsheetApp.openById(id);
  return readOnlyDashboardWebSnapshot_(ss, payload);
}

function readOnlyDashboardWebRequests_(ss) {
  const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.REQUESTS);
  if (!sheet) throw Error('필수 시트를 찾을 수 없습니다: REQUESTS');
  const values = sheet.getDataRange().getValues(), headers = {}, byId = Object.create(null);
  (values[0] || []).forEach(function(v, i) { if (v) headers[String(v)] = i + 1; });
  if (!headers['Request ID']) throw Error('REQUESTS 필수 헤더 누락: Request ID');
  values.slice(1).forEach(function(row, i) {
    const id = String(row[headers['Request ID'] - 1] || '');
    if (id) { if (!byId[id]) byId[id] = []; byId[id].push({row:i + 2, values:row}); }
  });
  return {headers:headers, byId:byId};
}

// Snapshot parsing mirrors the popup; pure aggregation stays shared.
function readOnlyDashboardWebSnapshot_(ss,payload){


  const tz = 'Asia/Seoul';
  const now = new Date();
  const key=function(d){return d instanceof Date&&Number.isFinite(d.getTime())?Utilities.formatDate(d,tz,'yyyy-MM-dd'):'';};
  const today=key(now),p=payload||{},from=p.from||key(new Date(now.getTime()-29*86400000)),to=p.to||today;
  function valid(d){return /^\d{4}-\d{2}-\d{2}$/.test(d)&&!isNaN(Date.parse(d))&&new Date(d).toISOString().slice(0,10)===d;}
  if(!valid(from)||!valid(to)||from>to||Date.parse(to)-Date.parse(from)>366*86400000)throw Error('조회 기간은 시작일 이후 최대 366일로 선택하세요.');
  const rt=readOnlyDashboardWebRequests_(ss),h=rt.headers;
  ['Request ID','Node','Status','Requested At','Completed At','Request Type','Total VI','3P Completed VI','2P Completed VI'].forEach(function(n){if(!h[n])throw Error('REQUESTS 필수 헤더 누락: '+n);});
  const barcodeProvidedReady=!!h['Barcode Provided'];
  const requests=Object.keys(rt.byId).map(function(id){if(rt.byId[id].length!==1)throw Error('중복 Request ID: '+id);const r=rt.byId[id][0].values,v=function(n){return opsValue_(r,h,n);};
    const start=v('Requested At'),end=v('Completed At');
    return {id:id,node:String(v('Node')||'미지정'),brand:String(v('Brand')||''),type:String(v('Request Type')),status:normalizeRequestStatusV2_(v('Status'))||APP_CONFIG.STATUS.TO_DO,
      allocationStatus:String(v('Allocation Status')||''),total:Number(v('Total VI'))||0,done3:Number(v('3P Completed VI'))||0,done2:Number(v('2P Completed VI'))||0,
      barcodeProvided:normalizeDashboardBarcodeProvided_(v('Barcode Provided')),requested:key(start),completed:key(end),
      lead:start instanceof Date&&end instanceof Date&&end>=start?(end-start)/3600000:null};});
  const sheet=ss.getSheetByName('WORK_SESSIONS'),sessions=[];let sessionReady=false,longTermReady=true;
  if(sheet){const values=sheet.getRange(1,1,Math.max(1,sheet.getLastRow()),13).getValues();sessionReady=WORK_SESSION_COLUMNS_.every(function(n,i){return values[0][i]===n;});
    if(sessionReady)values.slice(1).forEach(function(r){if(!r[0]&&!r[1])return;sessions.push({worker:String(r[1]||''),requestId:String(r[2]||''),stage:String(r[3]||''),status:String(r[4]||''),started:r[5] instanceof Date?r[5].getTime():null,endedAt:r[6] instanceof Date?r[6].getTime():null,ended:key(r[6]),minutes:typeof r[7]==='number'&&Number.isFinite(r[7])?Math.max(0,r[7]):null,delta:typeof r[12]==='number'&&Number.isFinite(r[12])?r[12]:null,sessionCount:1,aggregate:false});});}
  const longTerm=ss.getSheetByName('LONG_TERM_SESSIONS');
  if(longTerm){const values=longTerm.getRange(1,1,Math.max(1,longTerm.getLastRow()),DASHBOARD_LONG_TERM_SESSION_COLUMNS_V1.length).getValues();longTermReady=DASHBOARD_LONG_TERM_SESSION_COLUMNS_V1.every(function(n,i){return values[0][i]===n;});
    if(longTermReady)values.slice(1).forEach(function(r){if(r[1]!=='COMMITTED')return;sessions.push({worker:String(r[3]||''),requestId:String(r[2]||''),stage:String(r[4]||''),status:'PAUSED',started:r[9] instanceof Date?r[9].getTime():null,endedAt:r[10] instanceof Date?r[10].getTime():null,ended:String(r[5]||''),minutes:typeof r[6]==='number'&&Number.isFinite(r[6])?Math.max(0,r[6]):null,delta:typeof r[7]==='number'&&Number.isFinite(r[7])?r[7]:null,sessionCount:Math.max(1,Number(r[8])||1),aggregate:true,activeMinutes:typeof r[11]==='number'?r[11]:null,requestActiveMinutes:typeof r[12]==='number'?r[12]:null});});}
  return buildOperationsDashboard_(requests,sessions,{from:from,to:to,node:String(p.node||''),worker:String(p.worker||''),today:today,timezone:tz,updatedAt:Utilities.formatDate(now,tz,'yyyy-MM-dd HH:mm:ss'),sessionReady:sessionReady,longTermReady:longTermReady,barcodeProvidedReady:barcodeProvidedReady,canManageAssignments:false});
}
