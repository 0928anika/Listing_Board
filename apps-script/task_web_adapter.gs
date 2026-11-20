// Existing web RPC signatures, with current header-based schemas and stale-row guards.
function webTaskFingerprint_(row){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(row),Utilities.Charset.UTF_8).map(function(b){return ('0'+((b+256)%256).toString(16)).slice(-2);}).join('');}
function webTaskToken_(source,record,id){return 'OPS1:'+encodeURIComponent(JSON.stringify({source:source,row:record.row,id:id,version:webTaskFingerprint_(record.values)}));}
function readRequestWebTasks_(requestId){
  const rt=readOpsTable_(APP_CONFIG.SHEETS.REQUESTS),found=rt.byId[String(requestId)]||[];
  if(found.length!==1)throw Error('요청을 찾을 수 없거나 중복되었습니다.');
  const r=found[0].values,source=opsSourceForRequest_(r,rt.headers),table=readOpsTable_(source),h=table.headers,a=taskAdapter_(h),type=String(opsValue_(r,rt.headers,'Request Type')||'3P Listing');
  return (table.byId[String(requestId)]||[]).filter(function(record){return isViValidV3_(record.values,h);}).map(function(record){
    const row=record.values,d3=type==='2P Lite 전환'||a.d3>=0&&row[a.d3]===true,d2=type==='3P Listing'||a.d2>=0&&row[a.d2]===true;
    return {requestId:String(requestId),taskId:webTaskToken_(source,record,String(requestId)),
      productCode:String(opsValue_(row,h,'품번')||opsValue_(row,h,'업체상품ID')||''),productName:String(opsValue_(row,h,'등록상품명')||opsValue_(row,h,'업체 등록 상품명')||''),
      size:String(opsValue_(row,h,'사이즈')||''),color:String(opsValue_(row,h,'색상')||''),salePrice:Number(opsValue_(row,h,'판매가')||opsValue_(row,h,'판매가격'))||0,discountPrice:Number(opsValue_(row,h,'정상가')||opsValue_(row,h,'할인율기준가격'))||0,
      brand:String(opsValue_(r,rt.headers,'Brand')||''),node:String(opsValue_(r,rt.headers,'Node')||''),assignee:String(row[a.a3]||row[a.a2]||''),
      completed:d3&&d2,startedAt:'',completedAt:'',workStatus:d3&&d2?(type==='3P Listing'?'3P Listing 완료':'2P Lite 전환 완료'):type!=='2P Lite 전환'&&d3?'3P Listing 완료':'아무것도 안한 상태'};
  });
}
function resolveWebTask_(taskId){
  // Old clients used timestamp/obsolete columns as IDs. Refuse those instead of guessing.
  if(String(taskId).indexOf('OPS1:')!==0)throw Error('작업 목록을 새로고침한 뒤 다시 시도하세요.');
  const ref=JSON.parse(decodeURIComponent(String(taskId).slice(5)));
  if([APP_CONFIG.SHEETS.TASKS,APP_CONFIG.SHEETS.CRAWL_TASKS,APP_CONFIG.SHEETS.TASKS_2P].indexOf(ref.source)<0||!Number.isInteger(ref.row)||ref.row<2)throw Error('잘못된 Task 참조입니다.');
  const sheet=getRequiredSheet_(ref.source),h=getHeaderMap_(sheet);
  if(ref.row>sheet.getLastRow())throw Error('작업 행이 변경됐습니다. 새로고침하세요.');
  const row=sheet.getRange(ref.row,1,1,sheet.getLastColumn()).getValues()[0];
  if(String(opsValue_(row,h,'Request ID'))!==ref.id||webTaskFingerprint_(row)!==ref.version)throw Error('조회 이후 작업이 변경됐습니다. 새로고침하세요.');
  if(!isViValidV3_(row,h))throw Error('INVALID VI는 변경할 수 없습니다.');
  return {sheet:sheet,headers:h,values:row,row:ref.row,id:ref.id,adapter:taskAdapter_(h)};
}
function updateWebTaskCompletion_(taskId,value,isStatus,barcodeProvided){
  return withOpsDocumentLock_(function(){
    const t=resolveWebTask_(taskId),rt=readOpsTable_(APP_CONFIG.SHEETS.REQUESTS),found=rt.byId[t.id]||[];
    if(found.length!==1)throw Error('요청 ID가 없거나 중복되었습니다.');
    const r=found[0].values,type=String(opsValue_(r,rt.headers,'Request Type')||'3P Listing'),a=t.adapter;
    if(opsSourceForRequest_(r,rt.headers)!==t.sheet.getName())throw Error('요청의 작업 원본이 변경됐습니다. 새로고침하세요.');
    if(normalizeRequestStatusV2_(opsValue_(r,rt.headers,'Status'))===APP_CONFIG.STATUS.DROP)throw Error('DROP 상태인 요청입니다.');
    let done3,done2;
    if(isStatus){
      if(['아무것도 안한 상태','3P Listing 완료','2P Lite 전환 완료'].indexOf(value)<0)throw Error('허용되지 않은 작업 상태입니다.');
      if(type==='3P Listing'&&value==='2P Lite 전환 완료'||type==='2P Lite 전환'&&value==='3P Listing 완료')throw Error('요청 유형에 맞는 단계를 선택하세요.');
      done3=value!=='아무것도 안한 상태';done2=value==='2P Lite 전환 완료';
    }else{done3=value===true;done2=value===true;}
    const stages=[[a.a3,a.d3,type!=='2P Lite 전환',done3],[a.a2,a.d2,type!=='3P Listing',done2]];
    stages.forEach(function(stage){
      if(!stage[2]||stage[1]<0)return;
      if(stage[3]&&!splitWorkersV2_(t.values[stage[0]]).length)throw Error('담당자를 먼저 배정하세요.');
      if(stage[3]&&stage[1]===a.d2&&type==='3P Listing 후 2P Lite 전환'&&(a.d3<0||t.values[a.d3]!==true))throw Error('3P 작업을 완료한 뒤 2P 작업을 완료하세요.');
    });
    if(type!=='3P Listing'&&done2)captureBarcodeProvidedV1_(rt.sheet,found[0],rt.headers,barcodeProvided);
    const patches=[];
    stages.forEach(function(stage){
      if(!stage[2]||stage[1]<0)return;
      if(t.values[stage[1]]!==stage[3]){patches.push({row:t.row,col:stage[1]+1,value:stage[3]});t.values[stage[1]]=stage[3];}
    });
    writeOpsPatches_(t.sheet,patches);
    const cols=[a.d3,a.d2].filter(function(c){return c>=0;}).map(function(c){return c+1;});
    handleOperationalTaskEdit_({range:t.sheet.getRange(t.row,Math.min.apply(null,cols),1,Math.max.apply(null,cols)-Math.min.apply(null,cols)+1)},
      {sheet:t.sheet,h:t.headers,start:t.row,count:1,rows:[t.values],patches:[],requestTable:rt});
    return {ok:true,taskId:String(taskId),requestId:t.id};
  });
}
