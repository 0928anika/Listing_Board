// Deferred assignment: projected live reads, unchanged product grouping,
// atomic Sheets API writes in bounded groups. No writes during draft editing.
function requireAssignmentHeadersV4_(headers,names,sourceName){
  const missing=names.filter(function(name){return !headers[name];});
  if(missing.length)throw Error((sourceName?sourceName+' ':'')+'필수 열 누락: '+missing.join(', '));
  return headers;
}

function assignmentTableV3_(name) {
  const sheet = getRequiredSheet_(name);
  // Keep native Date/boolean values used by existing fingerprints, without
  // separate last-row / last-column service lookups for every source.
  const rows = sheet.getDataRange().getValues();
  const headers = {}, byRequest = Object.create(null);
  (rows[0] || []).forEach(function(value,i){if(value)headers[String(value)]=i+1;});
  requireAssignmentHeadersV4_(headers,['Request ID'],name);
  const idIndex=headers['Request ID']-1;
  rows.slice(1).forEach(function(row,i){const id=String(row[idIndex]||'');if(id){if(!byRequest[id])byRequest[id]=[];byRequest[id].push({row:i+2,values:row});}});
  return {sheet:sheet,id:sheet.getSheetId(),headers:headers,byRequest:byRequest};
}
function assignmentSourceV3_(r) {
  return r.requestType==='2P Lite 전환' ? APP_CONFIG.SHEETS.TASKS_2P
    : r.crawlRequired ? APP_CONFIG.SHEETS.CRAWL_TASKS : APP_CONFIG.SHEETS.TASKS;
}
// Only these Task fields are consumed by assignment/grouping/progress/summary.
// Internal Task ID also binds optimistic validation to the actual work item.
function assignmentFieldsV6_(name){
  return ['Request ID','Internal Task ID'].concat(name===APP_CONFIG.SHEETS.TASKS_2P
    ?['업체상품ID','Product ID','2P Assignee','2P Completed','VI Status']
    :['품번','등록상품명','3P Assignee','3P Completed','2P Assignee','2P Completed','VI Status']);
}
function assignmentRangeV6_(part){
  return "'"+part.name.replace(/'/g,"''")+"'!"+columnLetterV4_(part.left)+part.start+':'+columnLetterV4_(part.right)+part.end;
}
function assignmentResponseRangeV6_(range){
  const match=/^(?:'((?:[^']|'')+)'|([^!]+))!([A-Z]+)([0-9]+)(?::([A-Z]+)([0-9]+))?$/.exec(String(range||''));
  return match?[(match[1]||match[2]).replace(/''/g,"'"),match[3],Number(match[4]),match[5]||match[3],Number(match[6]||match[4])].join('|'):'';
}
// Split only oversized responses; ordinary reads retain a single Values call.
// Errors other than payload size, malformed responses, or exhausted fallback
// never become an empty table. A caller cannot save partially loaded records.
function assignmentReadPartsV6_(spreadsheetId,parts,accept,stats){
  if(!parts.length)return;
  if(Date.now()-stats.started>180000)throw Error('배정 데이터 조회 시간이 초과되었습니다. 저장하지 않았습니다.');
  const ranges=parts.map(assignmentRangeV6_);let result;
  try{
    stats.calls++;
    result=Sheets.Spreadsheets.Values.batchGet(spreadsheetId,{ranges:ranges,majorDimension:'ROWS',valueRenderOption:'UNFORMATTED_VALUE',dateTimeRenderOption:'SERIAL_NUMBER'});
  }catch(error){
    if(!/(?:\b413\b|response too large|response size.*exceed)/i.test(String(error)))throw error;
    stats.fallbacks++;
    if(parts.length>1){const middle=Math.ceil(parts.length/2);assignmentReadPartsV6_(spreadsheetId,parts.slice(0,middle),accept,stats);assignmentReadPartsV6_(spreadsheetId,parts.slice(middle),accept,stats);return;}
    const p=parts[0];let first,second;
    if(p.end>p.start){const middle=Math.floor((p.start+p.end)/2);first=Object.assign({},p,{end:middle});second=Object.assign({},p,{start:middle+1});}
    else if(p.right>p.left){const middle=Math.floor((p.left+p.right)/2);first=Object.assign({},p,{right:middle});second=Object.assign({},p,{left:middle+1});}
    else throw Error('배정 필수 셀의 응답이 너무 큽니다. 저장하지 않았습니다.');
    assignmentReadPartsV6_(spreadsheetId,[first],accept,stats);assignmentReadPartsV6_(spreadsheetId,[second],accept,stats);return;
  }
  if(!result||!Array.isArray(result.valueRanges)||result.valueRanges.length!==parts.length)throw Error('작업 데이터 조회가 불완전합니다. 저장하지 않았습니다.');
  result.valueRanges.forEach(function(value,index){
    const p=parts[index];
    if(!value||assignmentResponseRangeV6_(value.range)!==assignmentResponseRangeV6_(ranges[index])||value.majorDimension&&value.majorDimension!=='ROWS'||value.values!==undefined&&!Array.isArray(value.values))throw Error('작업 데이터 조회 범위가 일치하지 않습니다. 저장하지 않았습니다.');
    const rows=value.values||[];
    if(rows.length>p.end-p.start+1||rows.some(r=>!Array.isArray(r)||r.length>p.right-p.left+1||r.some(v=>v!==null&&!['string','number','boolean'].includes(typeof v))))throw Error('작업 데이터 응답 형식이 잘못되었습니다. 저장하지 않았습니다.');
    accept(p,rows);
  });
}
function assignmentTablesV5_(names) {
  const unique=names.filter(function(name,i){return names.indexOf(name)===i;}),tables={},parts=[],stats={started:Date.now(),calls:0,fallbacks:0},rowsBySheet={};
  if(!unique.length)return tables;
  unique.forEach(function(name){
    const sheet=getRequiredSheet_(name),width=sheet.getLastColumn(),last=sheet.getLastRow(),headers={};
    const labels=width?sheet.getRange(1,1,1,width).getValues()[0]:[],fields=assignmentFieldsV6_(name);
    labels.forEach(function(value,i){if(value){const key=String(value);if(headers[key]&&fields.indexOf(key)>=0)throw Error(name+' 중복 배정 헤더: '+key);headers[key]=i+1;}});
    requireAssignmentHeadersV4_(headers,['Request ID'],name);
    const columns=fields.filter(n=>headers[n]).map(n=>headers[n]).sort((a,b)=>a-b),groups=[];
    columns.forEach(function(col){const g=groups[groups.length-1];if(g&&g.right+1===col)g.right=col;else groups.push({name:name,left:col,right:col,start:2,end:last});});
    if(last>=2)Array.prototype.push.apply(parts,groups);
    tables[name]={id:sheet.getSheetId(),headers:headers,byRequest:Object.create(null)};
    rowsBySheet[name]={rows:[],last:last,columns:columns.length};
  });
  assignmentReadPartsV6_(SpreadsheetApp.getActiveSpreadsheet().getId(),parts,function(p,values){
    const rows=rowsBySheet[p.name].rows;
    for(let row=p.start;row<=p.end;row++){
      const output=rows[row-2]||(rows[row-2]=[]),input=values[row-p.start]||[];
      for(let col=p.left;col<=p.right;col++)output[col-1]=input[col-p.left]===undefined||input[col-p.left]===null?'':input[col-p.left];
    }
  },stats);
  unique.forEach(function(name){
    const table=tables[name],loaded=rowsBySheet[name];
    loaded.rows.forEach(function(row,i){const id=String(row[table.headers['Request ID']-1]||'');if(id){if(!table.byRequest[id])table.byRequest[id]=[];table.byRequest[id].push({row:i+2,values:row});}});
  });
  if(typeof opsTiming_==='function')opsTiming_('assignment_read',stats.started,{sources:unique.map(name=>({sheet:name,rows:Math.max(0,rowsBySheet[name].last-1),columns:rowsBySheet[name].columns})),ranges:parts.map(assignmentRangeV6_),calls:stats.calls,fallbacks:stats.fallbacks});
  return tables;
}
// Summary-only legacy v5 tokens remain unchanged. Task tokens use a separate
// namespace so pre-update drafts cannot silently bypass the new contract.
function assignmentRelevantVersionV4_(request,records,table){
  const fields=assignmentFieldsV6_(assignmentSourceV3_(request)),h=table&&!assignmentIsCrawlV3_(request)?table.headers:{};
  const value=(r,n)=>h[n]&&r.values[h[n]-1]!==undefined?r.values[h[n]-1]:'';
  const tasks=records.map(r=>({identity:[value(r,'Request ID'),value(r,'Internal Task ID')],row:r.row,values:fields.map(n=>value(r,n))}));
  const state=[assignmentVersionV3_(request,[]),fields.map(n=>[n,h[n]||0]),tasks];
  return 'assignment-v6:'+Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(state),Utilities.Charset.UTF_8).map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');
}
function assignmentRequestV3_(row,h) {
  function v(name){return h[name]?row[h[name]-1]:'';}
  return {requestId:String(v('Request ID')),requestType:String(v('Request Type')||'3P Listing'),
    status:normalizeRequestStatusV2_(v('Status'))||APP_CONFIG.STATUS.TO_DO,assignee:String(v('Active Assignees')||''),
    crawlRequired:v('크롤링 필요')===true,crawlStage:String(v('크롤링 단계')||''),crawlAssignee:String(v('크롤링 작업자')||''),
    totalVi:Number(v('Total VI'))||0,completed3P:Number(v('3P Completed VI'))||0,completed2P:Number(v('2P Completed VI'))||0};
}
function assignmentIsCrawlV3_(r){return r.crawlRequired&&r.crawlStage!=='CRAWL_COMPLETED';}
function assignmentVersionV3_(r,records) {
  const state=[r.requestId,r.requestType,r.status,r.assignee,r.crawlRequired,r.crawlStage,r.crawlAssignee,r.totalVi,r.completed3P,r.completed2P,records];
  // v5 dates use Sheets serial values. Older open consoles must refresh rather
  // than silently rebase a draft made with a different fingerprint encoding.
  return 'v5:'+Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(state),Utilities.Charset.UTF_8)
    .map(function(b){return ('0'+((b+256)%256).toString(16)).slice(-2);}).join('');
}
function assignmentWorkerProgressV1_(table,requestId) {
  const records=table.byRequest[requestId]||[],h=table.headers;
  const a3=(h['3P Assignee']||0)-1,d3=(h['3P Completed']||0)-1;
  const a2=(h['2P Assignee']||0)-1,d2=(h['2P Completed']||0)-1;
  const workers={};
  records.forEach(function(record){
    const row=record.values;
    if(!isViValidV3_(row,h))return;
    [[a3,d3,'3P'],[a2,d2,'2P']].forEach(function(stage){
      if(stage[0]<0)return;
      splitWorkersV2_(row[stage[0]]).forEach(function(worker){
        const progress=workers[worker]||(workers[worker]={assigned3P:0,done3P:0,assigned2P:0,done2P:0});
        progress['assigned'+stage[2]]++;
        if(stage[1]>=0&&row[stage[1]]===true)progress['done'+stage[2]]++;
      });
    });
  });
  return workers;
}
// Capture Task edits too, not just Active Assignees; reused at save time.
function attachAssignmentVersionsV3_(requests) {
  const active=requests.filter(function(r){return r.status!==APP_CONFIG.STATUS.DONE&&r.status!==APP_CONFIG.STATUS.DROP;});
  const tables=assignmentTablesV5_(active.filter(function(r){return !assignmentIsCrawlV3_(r);}).map(assignmentSourceV3_));
  requests.forEach(function(r){
    if(r.status===APP_CONFIG.STATUS.DONE||r.status===APP_CONFIG.STATUS.DROP)return;
    const name=assignmentSourceV3_(r);
    r.assignmentVersion=assignmentRelevantVersionV4_(r,assignmentIsCrawlV3_(r)?[]:(tables[name].byRequest[r.requestId]||[]),tables[name]);
    r.workerProgress=assignmentIsCrawlV3_(r)?{}:assignmentWorkerProgressV1_(tables[name],r.requestId);
  });
}
// Background preparation: keep the initial list read small, then verify its
// REQUESTS snapshot before attaching relevant Task fingerprints. Never write sheets.
function getAssignmentPreparationV4(expected) {
  if(!Array.isArray(expected))throw Error('조회할 요청 목록이 필요합니다.');
  const table=assignmentTableV3_(APP_CONFIG.SHEETS.REQUESTS),valid=[],results=[];
  expected.forEach(function(e){
    const found=table.byRequest[e.requestId]||[];
    if(found.length!==1){results.push({id:e.requestId,error:'요청이 없거나 중복되었습니다. 새로고침하세요.'});return;}
    const r=assignmentRequestV3_(found[0].values,table.headers);
    if(!e.summaryVersion||e.summaryVersion!==assignmentVersionV3_(r,[])){
      results.push({id:e.requestId,error:'목록 조회 이후 요청이 변경됐습니다. 새로고침하세요.'});return;
    }
    valid.push(r);
  });
  attachAssignmentVersionsV3_(valid);
  valid.forEach(function(r){results.push({id:r.requestId,version:r.assignmentVersion||'',workerProgress:r.workerProgress||{}});});
  return {results:results,workSessions:getWorkSessionSnapshot_(true)};
}
function assignmentCellV3_(value){return {userEnteredValue:{stringValue:String(value)}};}
// Only target, unfinished, valid owner cells are written. Never write checkbox columns.
function assignmentColumnWritesV3_(table,column,changes) {
  const writes=[];let run=null;
  changes.sort(function(a,b){return a.row-b.row;}).forEach(function(change){
    if(!run||run.start.rowIndex+run.rows.length!==change.row-1){
      run={start:{sheetId:table.id,rowIndex:change.row-1,columnIndex:column-1},rows:[],fields:'userEnteredValue'};
      writes.push({updateCells:run});
    }
    run.rows.push({values:[assignmentCellV3_(change.value)]});
  });
  return writes.reduce(function(out,write){
    const u=write.updateCells,compressed=[];let run=null;
    u.rows.forEach(function(row,i){
      const value=row.values[0].userEnteredValue.stringValue;
      if(!run||run.cell.userEnteredValue.stringValue!==value){
        run={range:{sheetId:table.id,startRowIndex:u.start.rowIndex+i,endRowIndex:u.start.rowIndex+i+1,startColumnIndex:column-1,endColumnIndex:column},cell:assignmentCellV3_(value),fields:'userEnteredValue'};
        compressed.push({repeatCell:run});
      }else run.range.endRowIndex++;
    });
    // Sparse / alternating ownership can be cheaper as the original update.
    // Never bridge gaps: completed, invalid and unrelated cells stay untouched.
    Array.prototype.push.apply(out,JSON.stringify(compressed).length<JSON.stringify(write).length?compressed:[write]);
    return out;
  },[]);
}
function assignmentPlanV3_(r,workers,table,requestTable,record) {
  const writes=[],affected=workers.slice(),h=requestTable.headers;
  function remember(value){splitWorkersV2_(value).forEach(function(w){if(affected.indexOf(w)<0)affected.push(w);});}
  function requestCell(name,value,clearValidation){
    // Assignment is ownership-only. Request workflow status changes only via
    // the explicit status/work-session paths, never as a side effect of save.
    if(name==='Status')throw Error('배정 저장은 요청 Status를 변경할 수 없습니다.');
    if(!h[name])throw Error('필수 열 누락: '+name);
    const update={start:{sheetId:requestTable.id,rowIndex:record.row-1,columnIndex:h[name]-1},rows:[{values:[assignmentCellV3_(value)]}],fields:'userEnteredValue'};
    if(clearValidation)update.fields+=',dataValidation';
    writes.push({updateCells:update});
  }
  const before=assignmentIsCrawlV3_(r)?r.crawlAssignee:r.assignee;
  remember(before);
  if(assignmentIsCrawlV3_(r)){
    if(workers.length!==1)throw Error('크롤링 담당자는 1명만 지원합니다.');
    requestCell('크롤링 작업자',workers[0]);requestCell('크롤링 단계','CRAWL_ASSIGNED');requestCell('자료 준비 상태','크롤링 배정됨');
    return {writes:writes,affected:affected,before:before,crawl:true};
  }
  const records=table.byRequest[r.requestId]||[],rows=records.map(function(x){return x.values;}),th=table.headers;
  const lite=r.requestType==='2P Lite 전환';
  const sourceName=assignmentSourceV3_(r);
  const required=lite
    ? ['Request ID','업체상품ID','Product ID','2P Assignee','2P Completed','VI Status']
    : ['Request ID','품번','등록상품명','3P Assignee','3P Completed','VI Status'].concat(r.requestType==='3P Listing'?[]:['2P Assignee','2P Completed']);
  requireAssignmentHeadersV4_(th,required,sourceName);
  const three=lite?null:[th['3P Assignee']-1,th['3P Completed']-1];
  const two=r.requestType==='3P Listing'?null:[th['2P Assignee']-1,th['2P Completed']-1];
  const skuIndex=(lite?th['업체상품ID']:th['품번'])-1;
  const nameIndex=(lite?th['Product ID']:th['등록상품명'])-1;
  const stages=lite?[two]:(r.requestType==='3P Listing'?[three]:[three,two]);
  let eligible=0;
  stages.forEach(function(stage){
    const targets=[];
    rows.forEach(function(row,i){remember(row[stage[0]]);if(isViValidV3_(row,th)&&row[stage[1]]!==true)targets.push(i);});
    eligible+=targets.length;
    const assigned=buildContiguousGroupAssignmentsV2_(rows,targets,workers,skuIndex,nameIndex);
    const changes=targets.filter(function(i){return rows[i][stage[0]]!==assigned[i];}).map(function(i){return {row:records[i].row,value:assigned[i]};});
    Array.prototype.push.apply(writes,assignmentColumnWritesV3_(table,stage[0]+1,changes));
  });
  if(!eligible)throw Error('배정할 미완료 Task가 없습니다. 새로고침 후 확인하세요.');
  requestCell('Active Assignees',workers.join(', '),true);
  if(h['Allocation Status'])requestCell('Allocation Status',ALLOCATION_STATUS_V2.AUTO);
  return {writes:writes,affected:affected,before:before,crawl:false};
}
function assignmentLogSerialV5_(time){
  // Sheets serial date, formatted in the spreadsheet's timezone like appendRow(Date).
  const tz=Session.getScriptTimeZone()||'Asia/Seoul';
  const local=Utilities.formatDate(time,tz,"yyyy-MM-dd'T'HH:mm:ss");
  return Date.parse(local+'Z')/86400000+25569;
}
// Reuse the authoritative snapshots already read for the atomic owner write.
// Summary failure must never turn a successful assignment into a failed save.
function reconcileAssignmentPlans_(plans,rt,tables){
  try{
  const sources={},tableMap={};tableMap[rt.id]=rt;
  Object.keys(tables).forEach(function(name){tableMap[tables[name].id]=tables[name];});
  Object.keys(tableMap).forEach(function(id){const t=tableMap[id];t.rowMap={};Object.keys(t.byRequest).forEach(key=>t.byRequest[key].forEach(r=>t.rowMap[r.row-1]=r.values));});
  plans.forEach(function(p){p.writes.forEach(function(w){
    const u=w.updateCells,g=w.repeatCell&&w.repeatCell.range,t=tableMap[u?u.start.sheetId:g&&g.sheetId];if(!t)return;
    const rows=t.rowMap;
    function value(cell){const v=cell.userEnteredValue||{};return v.stringValue!==undefined?v.stringValue:v.numberValue!==undefined?v.numberValue:v.boolValue!==undefined?v.boolValue:'';}
    if(u)u.rows.forEach((r,i)=>r.values.forEach((c,j)=>{if(rows[u.start.rowIndex+i])rows[u.start.rowIndex+i][u.start.columnIndex+j]=value(c);}));
    else if(g)for(let i=g.startRowIndex;i<g.endRowIndex;i++)if(rows[i])rows[i][g.startColumnIndex]=value(w.repeatCell.cell);
  });});
  Object.keys(tables).forEach(function(name){const t=tables[name];sources[name]={headers:t.headers,adapter:taskAdapter_(t.headers),rows:Object.keys(t.byRequest).reduce((a,id)=>a.concat(t.byRequest[id].map(r=>r.values)),[])};});
  const ids=plans.filter(p=>!p.crawl).map(p=>p.id);if(!ids.length)return true;
  reconcileRequestSummaries_(ids,{requestTable:{sheet:rt.sheet,headers:rt.headers,byId:rt.byRequest},taskSources:sources,
    domains:{tasks:true},preserveStatus:true,preferredAllocation:ALLOCATION_STATUS_V2.AUTO,notify:false,reason:'assignment'});return true;}
  catch(error){console.error('Assignment summary pending: '+String(error));return false;}
}
function saveAssignmentBatchV3(changes) {
  if(!Array.isArray(changes)||!changes.length||changes.length>300)throw Error('1~300건의 배정 변경을 전달하세요.');
  const ids=Object.create(null);changes.forEach(function(c){if(!c||typeof c.requestId!=='string'||!c.requestId||ids[c.requestId])throw Error('요청 ID가 없거나 중복되었습니다.');ids[c.requestId]=true;});
  const started=Date.now(),lock=LockService.getScriptLock(),timings={};
  if(!lock.tryLock(30000))throw Error('다른 사용자가 저장 중입니다. 새로고침 후 확인하세요.');
  timings.lockMs=Date.now()-started;
  try{
    const preparationStarted=Date.now();
    const rt=assignmentTableV3_(APP_CONFIG.SHEETS.REQUESTS),results=[],plans=[];
    const allowed=getWebConfig_().workers||[],ss=SpreadsheetApp.getActiveSpreadsheet(),spreadsheetId=ss.getId();
    const sourceNames=[];
    changes.forEach(function(c){
      const found=rt.byRequest[c.requestId]||[];if(found.length!==1)return;
      const r=assignmentRequestV3_(found[0].values,rt.headers);
      if(r.status!==APP_CONFIG.STATUS.DONE&&r.status!==APP_CONFIG.STATUS.DROP&&!assignmentIsCrawlV3_(r))sourceNames.push(assignmentSourceV3_(r));
    });
    const sourceStarted=Date.now(),tables=assignmentTablesV5_(sourceNames);
    timings.sourceReadMs=Date.now()-sourceStarted;
    changes.forEach(function(c){
      try{
        const found=rt.byRequest[String(c.requestId)]||[];
        if(found.length!==1)throw Error('요청이 없거나 ID가 중복되었습니다.');
        const record=found[0],r=assignmentRequestV3_(record.values,rt.headers);
        if(r.status===APP_CONFIG.STATUS.DONE||r.status===APP_CONFIG.STATUS.DROP)throw Error('이미 DONE 또는 DROP 상태인 요청입니다.');
        const workers=Array.isArray(c.workers)?c.workers:[];
        if(!workers.length||workers.some(function(w,i){return allowed.indexOf(w)<0||workers.indexOf(w)!==i;}))throw Error('등록된 작업자를 중복 없이 선택하세요.');
        const name=assignmentSourceV3_(r),crawl=assignmentIsCrawlV3_(r);
        const table=tables[name],records=crawl?[]:(table.byRequest[r.requestId]||[]);
        if(!c.version||c.version!==assignmentRelevantVersionV4_(r,records,table)){
          results.push({id:c.requestId,ok:false,code:'CONFLICT',message:'조회 이후 담당자·작업 데이터가 변경됐습니다. 새로고침 후 다시 확인하세요.'});return;
        }
        const plan=assignmentPlanV3_(r,workers,table,rt,record);
        plan.id=r.requestId;plan.workers=workers;plan.request=r;plan.records=records;
        plans.push(plan);
      }catch(e){results.push({id:c.requestId,ok:false,code:'INVALID',message:e.message||String(e)});}
    });
    timings.prepareMs=Date.now()-preparationStarted;
    timings.writeMs=0;timings.responseVersionMs=0;
    let index=0,writeCalls=0,payloadBytes=0;const affected=[];
    while(index<plans.length){
      if(Date.now()-started>240000){plans.slice(index).forEach(function(p){results.push({id:p.id,ok:false,code:'NOT_RUN',message:'실행 시간 보호를 위해 미저장. 다시 저장할 수 있습니다.'});});break;}
      const group=[],writes=[];let bytes=0;
      while(index<plans.length&&group.length<100){
        const p=plans[index],size=JSON.stringify(p.writes).length;
        if(group.length&&bytes+size>800000)break;
        group.push(p);Array.prototype.push.apply(writes,p.writes);bytes+=size;index++;
      }
      try{
        const combined=writes;
        markRequestReconcileNeeded_(group.filter(p=>!p.crawl).map(p=>p.id),{tasks:true},'assignment');
        const writeStarted=Date.now();
        try{Sheets.Spreadsheets.batchUpdate({requests:combined},spreadsheetId);writeCalls++;payloadBytes+=bytes;}
        finally{timings.writeMs+=Date.now()-writeStarted;}
      }catch(e){
        group.forEach(function(p){results.push({id:p.id,ok:false,code:'UNKNOWN',message:'저장 응답 확인 필요: '+(e.message||String(e))});});
        plans.slice(index).forEach(function(p){results.push({id:p.id,ok:false,code:'NOT_RUN',message:'앞선 저장 오류로 미실행'});});break;
      }
      const versionStarted=Date.now();
      const summaryReady=reconcileAssignmentPlans_(group,rt,tables);
      group.forEach(function(p){
        p.affected.forEach(function(w){if(affected.indexOf(w)<0)affected.push(w);});
        // Mirror only the same owner updates to produce the next optimistic version.
        if(p.crawl){p.request.crawlAssignee=p.workers[0];p.request.crawlStage='CRAWL_ASSIGNED';}
        else{
          p.request.assignee=p.workers.join(', ');
          const map={};p.records.forEach(function(r){map[r.row-1]=r;});
          p.writes.forEach(function(w){const table=tables[assignmentSourceV3_(p.request)];
            if(w.repeatCell){const u=w.repeatCell,g=u.range;if(g.sheetId!==table.id)return;
              for(let i=g.startRowIndex;i<g.endRowIndex;i++)map[i].values[g.startColumnIndex]=u.cell.userEnteredValue.stringValue;
            }else if(w.updateCells){const u=w.updateCells;if(u.start.sheetId!==table.id)return;
              u.rows.forEach(function(row,i){map[u.start.rowIndex+i].values[u.start.columnIndex]=row.values[0].userEnteredValue.stringValue;});
            }
          });
        }
        p.request=assignmentRequestV3_(rt.byRequest[p.id][0].values,rt.headers);
        results.push({id:p.id,ok:true,workers:p.workers,summaryPending:!summaryReady,version:summaryReady?assignmentRelevantVersionV4_(p.request,p.records,tables[assignmentSourceV3_(p.request)]):''});
      });
      timings.responseVersionMs+=Date.now()-versionStarted;
    }
    const notifyStarted=Date.now();
    let warning='';
    if(affected.length)try{bumpMyWorkSidebarRevisionV2_(affected);}catch(e){warning='배정은 저장됐지만 작업자 화면 갱신 알림에 실패했습니다. 해당 화면을 새로고침하세요.';}
    timings.notifyMs=Date.now()-notifyStarted;
    if(typeof opsTiming_==='function')opsTiming_('assignment_batch',started,{requests:changes.length,timings:timings,writeCalls:writeCalls,payloadBytes:payloadBytes});
    return {results:results,warning:warning,elapsedMs:Date.now()-started,writeCalls:writeCalls,timings:timings};
  }finally{lock.releaseLock();}
}

// Build the same atomic assignment plan used by the admin console, but trust
// only an installable edit event whose value is still current under the lock.
function saveDirectAssignmentEditsV4_(edits){
  if(!Array.isArray(edits)||!edits.length||edits.length>50)throw Error('한 번에 최대 50개 요청만 수기 배정할 수 있습니다.');
  const started=Date.now(),lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))throw Error('다른 배정 작업이 진행 중입니다. 잠시 후 다시 시도해주세요.');
  try{
    const rt=assignmentTableV3_(APP_CONFIG.SHEETS.REQUESTS),allowed=getWebConfig_().workers||[],prepared=[],results=[],sourceNames=[];
    edits.forEach(function(edit){
      const found=rt.byRequest[String(edit.requestId)]||[];
      if(found.length!==1){results.push({id:edit.requestId,ok:false,code:'INVALID',message:'요청이 없거나 ID가 중복되었습니다.'});return;}
      const record=found[0],request=assignmentRequestV3_(record.values,rt.headers),expected=String(edit.expectedAssignee||'');
      if(request.assignee!==expected){results.push({id:request.requestId,ok:false,code:'STALE',message:'더 최근의 담당자 편집이 있어 이전 처리를 건너뛰었습니다.'});return;}
      if(request.status===APP_CONFIG.STATUS.DONE||request.status===APP_CONFIG.STATUS.DROP){results.push({id:request.requestId,ok:false,code:'INVALID',message:'DONE 또는 DROP 상태인 요청은 수기 자동분배할 수 없습니다.'});return;}
      const workers=splitWorkersV2_(expected);
      if(workers.some(function(w,i){return allowed.indexOf(w)<0||workers.indexOf(w)!==i;})){
        results.push({id:request.requestId,ok:false,code:'INVALID',message:'CONFIG에 등록된 작업자를 중복 없이 입력하세요.'});return;
      }
      prepared.push({edit:edit,record:record,request:request,workers:workers});
      if(workers.length&&!assignmentIsCrawlV3_(request))sourceNames.push(assignmentSourceV3_(request));
    });
    const tables=assignmentTablesV5_(sourceNames);
    const plans=[],affected=[],reconcile=[];
    prepared.forEach(function(item){
      try{
        let plan;
        if(!item.workers.length){
          reconcile.push(item);
          return;
        }else{
          const name=assignmentSourceV3_(item.request),table=assignmentIsCrawlV3_(item.request)?null:tables[name];
          plan=assignmentPlanV3_(item.request,item.workers,table,rt,item.record);
          const taskChanged=!plan.crawl&&plan.writes.some(function(write){return write.repeatCell&&write.repeatCell.range.sheetId===table.id||write.updateCells&&write.updateCells.start.sheetId===table.id;});
          const allocation=rt.headers['Allocation Status']?String(item.record.values[rt.headers['Allocation Status']-1]||''):'';
          const crawlUnchanged=plan.crawl&&item.request.crawlAssignee===item.workers[0]&&item.request.crawlStage==='CRAWL_ASSIGNED';
          if((!plan.crawl&&!taskChanged&&allocation===ALLOCATION_STATUS_V2.AUTO)||(plan.crawl&&crawlUnchanged)){
            results.push({id:item.request.requestId,ok:true,code:'NOOP'});return;
          }
        }
        plan.id=item.request.requestId;plan.after=item.workers.join(', ');plan.before=String(item.edit.oldAssignee||'');
        plan.affected.forEach(function(w){if(affected.indexOf(w)<0)affected.push(w);});
        plans.push(plan);
      }catch(error){results.push({id:item.request.requestId,ok:false,code:'INVALID',message:error.message||String(error)});}
    });
    if(plans.length){
      const combined=[];
      plans.forEach(function(plan){Array.prototype.push.apply(combined,plan.writes);});
      try{
        markRequestReconcileNeeded_(plans.filter(p=>!p.crawl).map(p=>p.id),{tasks:true},'direct-assignment');
        Sheets.Spreadsheets.batchUpdate({requests:combined},SpreadsheetApp.getActiveSpreadsheet().getId());
        const summaryReady=reconcileAssignmentPlans_(plans,rt,tables);
        plans.forEach(function(plan){results.push({id:plan.id,ok:true,summaryPending:!summaryReady});});
      }catch(error){
        plans.forEach(function(plan){results.push({id:plan.id,ok:false,code:'UNKNOWN',message:'저장 응답을 확인하지 못했습니다: '+(error.message||String(error))});});
      }
    }
    if(reconcile.length){
      const requestTable=readOpsTable_(APP_CONFIG.SHEETS.REQUESTS),executionContext={};
      const synced=syncRequestSummaries_(reconcile.map(function(item){return item.request.requestId;}),{requestTable:requestTable,executionContext:executionContext});
      reconcile.forEach(function(item){
        const id=item.request.requestId,after=String(synced.summaries[id]['Active Assignees']||'');
        splitWorkersV2_(item.edit.oldAssignee).concat(splitWorkersV2_(after)).forEach(function(worker){if(affected.indexOf(worker)<0)affected.push(worker);});
        logActivity_(id,'ACTIVE_ASSIGNEES_RECONCILED_FROM_TASKS',String(item.edit.oldAssignee||''),after);
        results.push({id:id,ok:true,code:'RECONCILED',assignees:splitWorkersV2_(after)});
      });
    }
    if(affected.length)try{bumpMyWorkSidebarRevisionV2_(affected);}catch(ignore){}
    if(typeof opsTiming_==='function')opsTiming_('direct_assignment_batch',started,{requests:edits.length,plans:plans.length});
    return {results:results,elapsedMs:Date.now()-started};
  }finally{lock.releaseLock();}
}

function requestStatusEditNeedsInstalledV4_(e,h){
  const statusColumn=h['Status'];
  if(!statusColumn||e.range.getColumn()>statusColumn||e.range.getLastColumn()<statusColumn)return false;
  if(normalizeRequestStatusV2_(e.value)===APP_CONFIG.STATUS.DROP||normalizeRequestStatusV2_(e.oldValue)===APP_CONFIG.STATUS.DROP)return true;
  const statuses=e.range.getSheet().getRange(e.range.getRow(),statusColumn,e.range.getNumRows(),1).getDisplayValues();
  if(statuses.some(function(row){return normalizeRequestStatusV2_(row[0])===APP_CONFIG.STATUS.DROP;}))return true;
  if(!h['Cancelled At'])return false;
  return e.range.getSheet().getRange(e.range.getRow(),h['Cancelled At'],e.range.getNumRows(),1).getValues()
    .some(function(row){return !!row[0];});
}

// 설치형 상태 변경은 대량 작업과 겹쳐도 단순 트리거의 5초 대기값으로 실패하지 않게 함.
function withInstalledRequestStatusLockV4_(fn){
  const lock=LockService.getScriptLock();
  if(typeof lock.hasLock==='function'&&lock.hasLock())return fn();
  if(!lock.tryLock(30000))throw Error('다른 변경이 30초 넘게 실행 중이라 요청 상태 처리를 시작하지 못했습니다.');
  try{return fn();}finally{lock.releaseLock();}
}

// Authorized installable onEdit entry point for expensive REQUESTS edits.
function installedRequestAssignmentEditV4(e){
  if(!e||!e.range||e.range.getRow()<2)return;
  const sheet=e.range.getSheet();
  if(sheet.getName()!==APP_CONFIG.SHEETS.REQUESTS)return;
  const h=getHeaderMap_(sheet),statusColumn=h['Status'];
  if(requestStatusEditNeedsInstalledV4_(e,h)){
    return withInstalledRequestStatusLockV4_(function(){
      if(guardManualRequestDone_(e,h))return {rejected:true};
      const single=e.range.getNumRows()===1&&e.range.getNumColumns()===1;
      const expected=single&&e.value!==undefined?normalizeRequestStatusV2_(e.value):'';
      const current=single?normalizeRequestStatusV2_(sheet.getRange(e.range.getRow(),statusColumn).getDisplayValue()):'';
      if(expected&&current!==expected)return {stale:true};
      const cancelledAtBefore=single&&h['Cancelled At']?sheet.getRange(e.range.getRow(),h['Cancelled At']).getValue():'';
      handleAllocationManagementEditV2_(e);
      const finalStatus=single?normalizeRequestStatusV2_(sheet.getRange(e.range.getRow(),statusColumn).getDisplayValue()):'';
      if(expected&&finalStatus!==expected)return {rejected:true};
      if(!(single&&expected===APP_CONFIG.STATUS.DROP&&cancelledAtBefore))handleRequestAssignmentDateV11_(e);
      return {ok:true};
    });
  }
  const column=h['Active Assignees'];
  if(!column||e.range.getColumn()>column||e.range.getLastColumn()<column)return;
  if(e.range.getNumRows()!==1||e.range.getNumColumns()!==1)return;
  const start=e.range.getRow(),count=e.range.getNumRows(),rows=sheet.getRange(start,1,count,sheet.getLastColumn()).getValues();
  const edits=rows.map(function(row,index){return {requestId:String(row[(h['Request ID']||1)-1]||''),expectedAssignee:String(row[column-1]||''),oldAssignee:count===1?String(e.oldValue||''):'' ,row:start+index};}).filter(function(edit){return edit.requestId;});
  if(!edits.length)return;
  const spreadsheet=e.source||SpreadsheetApp.getActiveSpreadsheet();
  try{
    const result=saveDirectAssignmentEditsV4_(edits),failed=result.results.filter(function(r){return !r.ok&&r.code!=='STALE';}),saved=result.results.filter(function(r){return r.ok;});
    if(failed.length){
      const uncertain=failed.some(function(r){return r.code==='UNKNOWN';});
      if(!uncertain&&count===1){
        const current=String(sheet.getRange(start,column).getValue()||''),expected=edits[0].expectedAssignee;
        if(current===expected){
          sheet.getRange(start,column).setValue(String(e.oldValue||''));
        }
      }
      spreadsheet.toast((uncertain?'자동분배 저장 결과 확인 필요: ':'자동분배 실패: ')+failed[0].message,'Listing Ops',10);
    }else if(saved.length)spreadsheet.toast('남은 VALID·미완료 Task 자동분배가 완료되었습니다.','Listing Ops',5);
  }catch(error){
    // A single-cell edit can be safely restored before any batch write.  Do
    // not overwrite a newer value that arrived while this trigger was queued.
    if(count===1){
      const current=String(sheet.getRange(start,column).getValue()||''),expected=edits[0].expectedAssignee;
      if(current===expected){
        sheet.getRange(start,column).setValue(String(e.oldValue||''));
      }
    }
    spreadsheet.toast('자동분배 실패: '+(error.message||String(error)),'Listing Ops',10);
  }
}

// One-time idempotent installer. Keeps exactly one authorized edit trigger.
function installRequestAssignmentEditTriggerV4(){
  const handler='installedRequestAssignmentEditV4',triggers=ScriptApp.getProjectTriggers(),matching=triggers.filter(function(trigger){return trigger.getHandlerFunction()===handler;});
  matching.slice(1).forEach(function(trigger){ScriptApp.deleteTrigger(trigger);});
  if(!matching.length)ScriptApp.newTrigger(handler).forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
  const sessionRepair=repairCancelledActiveWorkSessionsV1();
  const repair=repairCancelledRequestValidityV4();
  const assignmentRepair=repairTaskDerivedAssignmentSummariesV6();
  return {handler:handler,installed:matching.length?false:true,duplicatesRemoved:Math.max(matching.length-1,0),cancelledSessionRepair:sessionRepair,cancelledRepair:repair,assignmentRepair:assignmentRepair};
}
