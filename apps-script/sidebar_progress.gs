// Read-only UI progress. No persisted row index, summary or session metrics.
// Manual refresh ignores summary revision as a freshness gate. One bounded page,
// no Drive lookup, filter-view rebuilding, reconciliation or full Sidebar call.
function getMyWorkSidebarLight(payload){
  payload=payload||{};
  const b=sidebarBootstrapContext_(),h=b.rt.headers,ids=Array.from(new Set((payload.requestIds||[]).map(String)));
  if(ids.length>8)throw Error('최대 8개 요청을 조회하세요.');
  if(payload.worker!==b.worker)return {mode:'stale',worker:b.worker};
  const active=new Set(b.sessions.active.map(s=>s.requestId)),structure=[];
  Object.keys(b.rt.byId).forEach(function(id){const records=b.rt.byId[id];if(records.length!==1)throw Error('중복 요청 ID');const r=records[0],v=n=>opsValue_(r.values,h,n),crawl=v('크롤링 필요')===true;
    if(!b.worker||!active.has(id)&&!sidebarRequestVisible_(r,h))return;
    if(active.has(id)||splitWorkersV2_(v('Active Assignees')).indexOf(b.worker)>=0||crawl&&String(v('크롤링 작업자'))===b.worker&&String(v('크롤링 단계'))!=='CRAWL_COMPLETED')structure.push({requestId:id,requestType:String(v('Request Type')||'3P Listing'),crawlRequired:crawl,crawlStage:String(v('크롤링 단계')||'')});
  });
  function key(list){return JSON.stringify(list.map(r=>[String(r.requestId),String(r.requestType||'3P Listing'),!!r.crawlRequired,String(r.crawlStage||'')]).sort((a,b)=>a[0].localeCompare(b[0])));}
  if(key(structure)!==key(payload.structure||[]))return {mode:'reload',worker:b.worker};
  const allowed=new Set(structure.map(r=>r.requestId));
  if(ids.some(id=>!allowed.has(id)))return {mode:'reload',worker:b.worker};
  const sources=scopedSourcesForRequests_(b.rt,ids,{sidebarTaskLookup:true}),cards=[];let membershipChanged=false;
  ids.forEach(function(id){const record=b.rt.byId[id][0],r=sidebarCardShell_(record,h),m=sources[opsSourceForRequest_(record.values,h)].index[id]||emptyTaskAggregate_(),w=m.workers[b.worker]||{};
    // Numeric progress and button readiness derive from this same Task snapshot.
    const crawlMine=r.crawlRequired&&r.crawlAssignee===b.worker&&r.crawlStage!=='CRAWL_COMPLETED';
    if(!active.has(id)&&!crawlMine&&!(w.assigned3P||w.assigned2P))membershipChanged=true;
    cards.push({requestId:id,status:r.status,barcodeProvided:r.barcodeProvided,assigned3P:w.assigned3P||0,done3P:w.done3P||0,assigned2P:w.assigned2P||0,done2P:w.done2P||0,totalVi:m.total,completed3P:m.done3,completed2P:m.done2,progressPending:false});
  });
  if(membershipChanged)return {mode:'reload',worker:b.worker};
  return {mode:'light',worker:b.worker,revision:b.revision,requests:cards,workSessions:b.sessions,updatedAt:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss')};
}
function sidebarProgressMatches_(sheet,headers,requestId) {
  if(!headers['Request ID'])throw Error('Request ID 헤더가 없습니다.');
  const last=sheet.getLastRow();
  if(last<2)return [];
  // Searches one ID column on Google's side; not an indexed/O(1) lookup.
  return sheet.getRange(2,headers['Request ID'],last-1,1).createTextFinder(requestId)
    .matchEntireCell(true).matchCase(true).useRegularExpression(false).findAll()
    .map(function(cell){return cell.getRow();}).sort(function(a,b){return a-b;});
}
function getMyWorkProgressCounts(payload) {
  payload=payload||{};
  const worker=PropertiesService.getUserProperties().getProperty('LISTING_OPS_MY_WORKER')||'';
  const revision=getMyWorkSidebarRevision(); // Capture before reads; never consume a later edit's revision.
  function full(unsupported){if(payload.allowShellFallback)return {mode:'reload',progressUnsupported:unsupported?id:''};return {mode:'full',data:getMyWorkSidebarData(false),progressUnsupported:unsupported?id:''};}
  const id=String(payload.requestId||''),stage=String(payload.stage||'');
  if(!worker||worker!==payload.worker||!id||id.length>200||['3P','2P'].indexOf(stage)<0)return full();
  const requests=getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS),rh=getHeaderMap_(requests),matches=sidebarProgressMatches_(requests,rh,id);
  if(matches.length!==1)return full();
  const request=requests.getRange(matches[0],1,1,requests.getLastColumn()).getValues()[0];
  if(String(opsValue_(request,rh,'Request ID'))!==id)return {mode:'retry'};
  const type=String(opsValue_(request,rh,'Request Type')||'3P Listing'),status=normalizeRequestStatusV2_(opsValue_(request,rh,'Status')),source=opsSourceForRequest_(request,rh);
  // Structural/terminal changes use only the existing full path, before Task values are read.
  if(type!==payload.requestType||normalizeBarcodeProvidedV1_(opsValue_(request,rh,'Barcode Provided'))!==normalizeBarcodeProvidedV1_(payload.barcodeProvided)||
      Boolean(opsValue_(request,rh,'크롤링 필요')===true)!==Boolean(payload.crawlRequired)||
      String(opsValue_(request,rh,'크롤링 단계')||'')!==String(payload.crawlStage||'')||
      [APP_CONFIG.STATUS.DONE,APP_CONFIG.STATUS.DROP,APP_CONFIG.STATUS.ON_HOLD].indexOf(status)>=0||stage==='2P'&&type==='3P Listing'||stage==='3P'&&type==='2P Lite 전환')return full();
  const sheet=getRequiredSheet_(source),headers=getHeaderMap_(sheet),rows=sidebarProgressMatches_(sheet,headers,id),groups=[];
  rows.forEach(function(row){const g=groups[groups.length-1];if(g&&g.start+g.count===row)g.count++;else groups.push({start:row,count:1});});
  // Avoid many small read RPCs or a near-full-source read for an unusually large request.
  if(groups.length>8||rows.length>10000)return full(true);
  const names=['Request ID','VI Status','3P Assignee','3P Completed','2P Assignee','2P Completed'];
  const cols=names.map(function(name){return headers[name];}).filter(Boolean),first=Math.min.apply(null,cols),last=Math.max.apply(null,cols),localHeaders={};
  names.forEach(function(name){if(headers[name])localHeaders[name]=headers[name]-first+1;});
  const values=[];
  groups.forEach(function(g){Array.prototype.push.apply(values,sheet.getRange(g.start,first,g.count,last-first+1).getValues());});
  if(values.some(function(row){return String(opsValue_(row,localHeaders,'Request ID'))!==id;}))return {mode:'retry'};
  const sources={};sources[source]={rows:values,headers:localHeaders,adapter:taskAdapter_(localHeaders)};
  const m=buildTaskIndex_(sources)[source][id]||emptyTaskAggregate_(),w=m.workers[worker]||{};
  return {mode:'progress',revision:revision,worker:worker,requestId:id,
    counts:['3P','2P'].filter(function(s){return s==='3P'?type!=='2P Lite 전환':type!=='3P Listing';}).map(function(s){
      return {stage:s,assigned:Number(w['assigned'+s])||0,completed:Number(w['done'+s])||0};
    }),totalVi:m.total,completed3P:m.done3,completed2P:m.done2};
}
