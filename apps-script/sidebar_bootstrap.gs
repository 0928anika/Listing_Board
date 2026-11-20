// Progressive Sidebar reads. No Task writes or persisted indexes.
function sidebarReadResponse_(name,fn){return fn();}
function sidebarBootstrapContext_(requestedWorker) {
  const workers=getWebConfig_().workers||[],props=PropertiesService.getUserProperties();
  if(requestedWorker!==undefined){if(requestedWorker&&workers.indexOf(requestedWorker)<0)throw Error('작업자를 확인하세요.');props.setProperty('LISTING_OPS_MY_WORKER',requestedWorker);}
  const saved=props.getProperty('LISTING_OPS_MY_WORKER')||'',worker=workers.indexOf(saved)>=0?saved:'',revision=getMyWorkSidebarRevision(),rt=readOpsTable_(APP_CONFIG.SHEETS.REQUESTS);
  const snapshot=worker?sidebarWorkerSessions_(worker):{sessions:[],active:[]};
  const sessions={sessions:snapshot.sessions.filter(function(s){return s.worker===worker;}),active:snapshot.active.filter(function(s){return s.worker===worker;})};
  return {worker:worker,workers:workers,revision:revision,rt:rt,sessions:sessions};
}
function sidebarCardShell_(record,h) {
  function v(name){return opsValue_(record.values,h,name);}
  const originalFileUrl=String(v('원본 요청 파일')||'');
  return {requestId:String(v('Request ID')),requestType:String(v('Request Type')||'3P Listing'),status:normalizeRequestStatusV2_(v('Status'))||APP_CONFIG.STATUS.TO_DO,requestor:String(v('Requestor')),brand:String(v('Brand')),node:String(v('Node')),aCode:String(v('A코드')),source:String(v('Source')),priority:String(v('Priority')),assignee:String(v('Active Assignees')),
    totalVi:Number(v('Total VI'))||0,completed3P:Number(v('3P Completed VI'))||0,completed2P:Number(v('2P Completed VI'))||0,
    assigned3P:null,done3P:null,assigned2P:null,done2P:null,progressPending:true,detailsPending:true,notes:'',originalFileUrl:/^https?:\/\//i.test(originalFileUrl)?originalFileUrl:'',taskViewUrls:{threeP:'',twoP:''},
    barcodeProvided:normalizeBarcodeProvidedV1_(v('Barcode Provided')),crawlRequired:v('크롤링 필요')===true,crawlStage:String(v('크롤링 단계')),crawlAssignee:String(v('크롤링 작업자')),preparationStatus:String(v('자료 준비 상태'))};
}
function sidebarRequestVisible_(record,h){return [APP_CONFIG.STATUS.DONE,APP_CONFIG.STATUS.DROP].indexOf(normalizeRequestStatusV2_(opsValue_(record.values,h,'Status')))<0;}
function getMyWorkSidebarInitial(worker){return sidebarReadResponse_('initial',function(){
  const b=sidebarBootstrapContext_(worker),h=b.rt.headers,active=new Set(b.sessions.active.map(function(s){return s.requestId;})),candidates=[],cards=[];
  Object.keys(b.rt.byId).forEach(function(id){const records=b.rt.byId[id];if(records.length!==1)throw Error('중복 요청 ID');const record=records[0],status=normalizeRequestStatusV2_(opsValue_(record.values,h,'Status')),crawl=b.worker&&opsValue_(record.values,h,'크롤링 필요')===true&&String(opsValue_(record.values,h,'크롤링 작업자'))===b.worker&&String(opsValue_(record.values,h,'크롤링 단계'))!=='CRAWL_COMPLETED';
    if(!b.worker||status===APP_CONFIG.STATUS.DROP||status===APP_CONFIG.STATUS.DONE&&!active.has(id))return;
    if(active.has(id)||crawl||splitWorkersV2_(opsValue_(record.values,h,'Active Assignees')).indexOf(b.worker)>=0){const shell=sidebarCardShell_(record,h);shell.taskSource=opsSourceForRequest_(record.values,h);candidates.push(shell);if(active.has(id)||crawl)cards.push(shell);}});
  return {spreadsheetId:SpreadsheetApp.getActiveSpreadsheet().getId(),worker:b.worker,workers:b.workers,revision:b.revision,requests:cards,candidateRequests:candidates,workSessions:b.sessions,updatedAt:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm'),membershipPending:!!candidates.length,progressive:true};
});}
function sidebarCheckView_(payload){const worker=PropertiesService.getUserProperties().getProperty('LISTING_OPS_MY_WORKER')||'',revision=getMyWorkSidebarRevision();return {worker:worker,revision:revision,stale:!payload||payload.worker!==worker||payload.revision!==revision};}
function sidebarFinishView_(check,data){const latest=sidebarCheckView_(check);return latest.stale?latest:Object.assign(check,data);}
// Read one column only for the requested physical rows in a single Sheets batch call.
function sidebarBatchColumnRows_(sheet,column,rows){
  const unique=Array.from(new Set(rows.map(Number))).filter(function(row){return row>=2;}).sort(function(a,b){return a-b;});
  if(!unique.length)return Object.create(null);
  const groups=[];unique.forEach(function(row){const group=groups[groups.length-1];if(group&&group.start+group.count===row)group.count++;else groups.push({start:row,count:1});});
  let letter='',value=Number(column);while(value>0){value--;letter=String.fromCharCode(65+value%26)+letter;value=Math.floor(value/26);}
  const prefix="'"+sheet.getName().replace(/'/g,"''")+"'!",ranges=groups.map(function(group){return prefix+letter+group.start+':'+letter+(group.start+group.count-1);});
  const result=Sheets.Spreadsheets.Values.batchGet(SpreadsheetApp.getActiveSpreadsheet().getId(),{ranges:ranges,majorDimension:'ROWS',valueRenderOption:'UNFORMATTED_VALUE'});
  if(!result||!Array.isArray(result.valueRanges)||result.valueRanges.length!==groups.length)throw Error('작업 후보 행 조회가 불완전합니다. 다시 시도하세요.');
  const found=Object.create(null);groups.forEach(function(group,index){const values=result.valueRanges[index].values||[];for(let offset=0;offset<group.count;offset++)found[group.start+offset]=values[offset]&&values[offset][0]!==undefined?values[offset][0]:'';});
  return found;
}
function discoverMyWorkSidebarRequests(payload){return sidebarReadResponse_('discover',function(){
  const check=sidebarCheckView_(payload);if(check.stale)return check;
  if(Array.isArray(payload.candidates)){
    const allowed=[APP_CONFIG.SHEETS.TASKS,APP_CONFIG.SHEETS.CRAWL_TASKS,APP_CONFIG.SHEETS.TASKS_2P],active=new Set((payload.activeRequestIds||[]).map(String)),byId={},groups={},cards=[];
    payload.candidates.forEach(function(c){if(c&&allowed.indexOf(c.taskSource)>=0&&c.requestId){const id=String(c.requestId),crawlMine=c.crawlRequired&&c.crawlAssignee===check.worker&&c.crawlStage!=='CRAWL_COMPLETED';byId[id]=c;if(crawlMine){cards.push(Object.assign({},c,{assigned3P:0,done3P:0,assigned2P:0,done2P:0,progressPending:false}));return;}(groups[c.taskSource]||(groups[c.taskSource]=[])).push(id);}});
    // Keep assigned VALID work after all checks: PAUSE is not Request DONE.
    Object.keys(groups).forEach(function(source){const table=readSidebarTasksForRequests_(source,groups[source],{});groups[source].forEach(function(id){const m=table.index[id]||emptyTaskAggregate_(),personal=m.workers[check.worker],candidate=byId[id];if(active.has(id)||personal&&(personal.assigned3P>0||personal.assigned2P>0))cards.push(Object.assign({},candidate,personal||{assigned3P:0,done3P:0,assigned2P:0,done2P:0},{progressPending:false}));});});
    return sidebarFinishView_(check,{requests:cards});
  }
  const rt=readOpsTable_(APP_CONFIG.SHEETS.REQUESTS),h=rt.headers,names={},found=new Set();
  Object.keys(rt.byId).forEach(function(id){const r=rt.byId[id][0];if(sidebarRequestVisible_(r,h)){names[opsSourceForRequest_(r.values,h)]=true;if(opsValue_(r.values,h,'크롤링 필요')===true&&String(opsValue_(r.values,h,'크롤링 작업자'))===check.worker&&String(opsValue_(r.values,h,'크롤링 단계'))!=='CRAWL_COMPLETED')found.add(id);}});
  // Assignee search only produces candidates. Validate candidate VI Status first, then read IDs for surviving rows only.
  Object.keys(names).forEach(function(name){const sheet=getRequiredSheet_(name),hd=getHeaderMap_(sheet),n=sheet.getLastRow()-1;if(n<1||!hd['Request ID'])return;
    const escaped=String(check.worker).replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
    const pattern='(^|\\s*,\\s*)'+escaped+'(\\s*,\\s*|$)',matchedRows=new Set();
    ['3P Assignee','2P Assignee'].forEach(function(field){if(!hd[field])return;sheet.getRange(2,hd[field],n,1).createTextFinder(pattern).matchCase(true).useRegularExpression(true).findAll().forEach(function(cell){matchedRows.add(cell.getRow());});});
    if(!matchedRows.size)return;
    const candidates=Array.from(matchedRows),viStatus=hd['VI Status']?sidebarBatchColumnRows_(sheet,hd['VI Status'],candidates):Object.create(null);
    const alive=candidates.filter(function(rowNumber){return !hd['VI Status']||String(viStatus[rowNumber]||'VALID').trim().toUpperCase()!=='INVALID';});
    if(!alive.length)return;
    const ids=sidebarBatchColumnRows_(sheet,hd['Request ID'],alive);
    alive.forEach(function(rowNumber){const id=String(ids[rowNumber]||''),records=rt.byId[id];if(records&&records.length===1&&sidebarRequestVisible_(records[0],h)&&opsSourceForRequest_(records[0].values,h)===name)found.add(id);});
  });
  return sidebarFinishView_(check,{requests:Object.keys(rt.byId).filter(function(id){return found.has(id);}).map(function(id){return sidebarCardShell_(rt.byId[id][0],h);})});
});}
function getMyWorkSidebarBatch(payload){return sidebarReadResponse_('progress',function(){
  const check=sidebarCheckView_(payload);if(check.stale)return check;
  const ids=Array.from(new Set((payload.requestIds||[]).map(String))),active=new Set((payload.activeRequestIds||[]).map(String));if(ids.length>8)throw Error('최대 8개 요청을 조회하세요.');
  const rt=readOpsTable_(APP_CONFIG.SHEETS.REQUESTS),h=rt.headers,context={sidebarTaskLookup:true},sources=scopedSourcesForRequests_(rt,ids,context),cards=[];
  ids.forEach(function(id){const record=rt.byId[id][0];if(!sidebarRequestVisible_(record,h)&&!active.has(id))return;const table=sources[opsSourceForRequest_(record.values,h)],m=table.index[id]||emptyTaskAggregate_(),personal=m.workers[check.worker],r=sidebarCardShell_(record,h);
    const crawlMine=r.crawlRequired&&r.crawlAssignee===check.worker&&r.crawlStage!=='CRAWL_COMPLETED';
    if(!active.has(id)&&!crawlMine&&(!personal||!personal.assigned3P&&!personal.assigned2P))return;
    Object.assign(r,personal||{assigned3P:0,done3P:0,assigned2P:0,done2P:0},{progressPending:false});cards.push(r);
  });return sidebarFinishView_(check,{requests:cards,requestIds:ids});
});}
function getMyWorkSidebarDetails(payload){return sidebarReadResponse_('details',function(){
  const check=sidebarCheckView_(payload);if(check.stale)return check;
  const rt=readOpsTable_(APP_CONFIG.SHEETS.REQUESTS),h=rt.headers,types={},cards=[];
  const ids=Array.from(new Set((payload.requestIds||[]).map(String)));if(ids.length>8)throw Error('최대 8개 요청');
  ids.forEach(function(id){const records=rt.byId[id];if(!records||records.length!==1)throw Error('요청 ID 확인');const r=records[0],v=function(name){return opsValue_(r.values,h,name);};types[id]=String(v('Request Type'));
    let url=String(v('원본 요청 파일')||'');if(h['원본 요청 파일']){const rich=rt.sheet.getRange(r.row,h['원본 요청 파일']).getRichTextValues()[0][0];if(rich)url=rich.getLinkUrl()||(rich.getRuns().find(function(run){return run.getLinkUrl();})||{getLinkUrl:function(){return '';}}).getLinkUrl()||url;}
    if(!/^https?:\/\//i.test(url)&&h['원본 요청 파일']){const note=rt.sheet.getRange(r.row,h['원본 요청 파일']).getNotes()[0][0];if(/^https?:\/\//i.test(note))url=note;}
    if(url&&!/^https?:\/\//i.test(url)){try{const files=getOriginalRequestFolderV2_().getFilesByName(url);if(files.hasNext())url=files.next().getUrl();}catch(ignored){}}
    cards.push({requestId:id,notes:String(v('Requestor Notes')),originalFileUrl:/^https?:\/\//i.test(url)?url:'',crawlResultFileUrl:String(v('크롤링 결과 파일')||'')});
  });const urls=getMyWorkFilterViewUrlsV2_(types);cards.forEach(function(r){r.taskViewUrls=urls[r.requestId]||{};});return sidebarFinishView_(check,{requests:cards});
});}
