// COMPLETE shares the session command transport but commits Request DONE and
// session closure together. Task validation is always fresh under the write lock.
function completeProof_(row){const m=/(?:^|\s\|\s)COMPLETE:([a-zA-Z0-9-]{16,80})(?:$|\s\|\s)/.exec(String(row[9]||''));return m?m[1]:'';}
function completeValidEnd_(row){return sessionValidEnd_(row)&&row[5] instanceof Date&&Number.isFinite(row[5].getTime())&&row[6]>=row[5]&&Math.abs((row[6]-row[5])-row[7]*60000)<=1.1;}
function completeReadCommitted_(p,receipt,providedRows,providedTable){
  if(!receipt||!receipt.complete)return null;
  const rows=providedRows||workSessionRows_(),matches=rows.filter(r=>String(r[0])===String(receipt.sessionId));
  if(matches.length!==1)return null;const row=matches[0],proof=receipt.complete;
  if(String(row[1])!==p.worker||String(row[2])!==p.requestId||String(row[3])!==p.stage||!completeValidEnd_(row))return null;
  if(Math.abs(row[6].getTime()-proof.sessionEndedAt)>1)return null;
  if(!proof.adopted&&completeProof_(row)!==proof.operationId)return null;
  const rt=providedTable||readSessionRequestRecords_([p.requestId]),record=rt.byId[p.requestId][0],h=rt.headers;
  const status=normalizeRequestStatusV2_(opsValue_(record.values,h,'Status')),ended=opsValue_(record.values,h,'Completed At');
  if(proof.requestDone&&(status!==APP_CONFIG.STATUS.DONE||!(ended instanceof Date)||Math.abs(ended.getTime()-proof.completedAt)>1||rows.some(r=>String(r[2])===p.requestId&&r[4]==='ACTIVE')))return null;
  const workerRows=rows.filter(r=>String(r[1])===p.worker);
  if(workerRows.filter(r=>r[4]==='ACTIVE').length>1)return null;
  return {ok:true,state:'COMMITTED',action:'COMPLETE',operationId:p.operationId,sessionId:receipt.sessionId,sessionStatus:'PAUSED',
    requestDone:proof.requestDone,requestPatch:{requestId:p.requestId,status:status},needsRefresh:true,
    postProcessing:receipt.postDone?'DONE':'PENDING',snapshot:getWorkSessionSnapshot_(true,workerRows)};
}
function completeReceipt_(p,row,requestDone,completedAt,adopted,operationId){
  const receipt={operationId:p.operationId,worker:p.worker,requestId:p.requestId,stage:p.stage,action:'COMPLETE',sessionId:p.sessionId,
    affected:[p.requestId],createdAt:Number(p.issuedAt),postDone:false,
    complete:{requestDone:requestDone,completedAt:requestDone?completedAt.getTime():null,sessionEndedAt:row[6].getTime(),adopted:!!adopted,operationId:operationId||p.operationId}};
  saveSessionReceipt_(receipt);
  markRequestReconcileNeeded_([p.requestId],{tasks:true,sessions:true},'complete-commit');
  PropertiesService.getScriptProperties().setProperty('SESSION_POST_PENDING','1');return receipt;
}
function completeWorkSessionCommand_(payload){
  const legacy=!payload.operationId,p=Object.assign({},payload),context={},began=Date.now();
  if(legacy){p.operationId=Utilities.getUuid();p.issuedAt=Date.now();}
  const result=withSessionCommitLock_(function(){
    context.lockWaitMs=Date.now()-began;
    try{
      const sheet=ensureWorkSessionsSheet_(),rows=workSessionRows_();
      if(legacy&&!p.sessionId){const active=rows.filter(r=>String(r[1])===String(p.worker)&&String(r[2])===String(p.requestId)&&String(r[3])===String(p.stage)&&r[4]==='ACTIVE');if(active.length===1)p.sessionId=String(active[0][0]);}
      Object.assign(p,sessionCommand_(p));
      let receipt=sessionReceipt_(p.operationId);
      if(receipt&&!sessionSameCommand_(receipt,p))throw Error('다른 세션 명령에 사용된 처리 ID입니다.');
      if(receipt&&receipt.rejected)throw Error(receipt.rejected);
      const rt=readSessionRequestRecords_([p.requestId]),record=rt.byId[p.requestId][0],h=rt.headers,v=n=>opsValue_(record.values,h,n);
      if(receipt){const committed=completeReadCommitted_(p,receipt,rows,rt);if(committed)return committed;}
      if(!receipt&&(!Number.isFinite(Number(p.issuedAt))||Math.abs(Date.now()-Number(p.issuedAt))>86400000))throw Error('오래된 미처리 완료 요청입니다. 최신 작업 상태를 확인하세요.');
      const matching=rows.map((r,i)=>({row:r,index:i})).filter(x=>String(x.row[0])===String(p.sessionId));
      if(matching.length!==1)throw Error('완료할 Session ID를 정확히 확인할 수 없습니다.');
      const target=matching[0],r=target.row;
      if(String(r[1])!==p.worker||String(r[2])!==p.requestId||String(r[3])!==p.stage)throw Error('Session ID의 작업자/요청/단계가 일치하지 않습니다.');
      const type=String(v('Request Type')||'3P Listing'),status=normalizeRequestStatusV2_(v('Status'));
      if(r[4]!=='ACTIVE'&&status!==APP_CONFIG.STATUS.DONE&&!completeProof_(r))throw Error('진행 중인 작업 세션이 아닙니다. 최신 상태를 확인하세요.');
      if(['3P','2P'].indexOf(p.stage)<0)throw Error('이 단계의 완료는 전용 완료 버튼을 사용해 주세요.');
      validateWorkStage_(type,p.stage);
      if(status===APP_CONFIG.STATUS.DROP||String(v('Archive Status')).toUpperCase()==='ARCHIVED')throw Error('DROP 또는 백업 완료 요청은 여기서 완료 처리할 수 없습니다.');
      assertWorkSessionViSchema_(sheet);
      // Fresh raw Task read; no browser counts, persisted summary or prior context.
      const vi=createWorkSessionViContext_(rt);vi.sidebarTaskLookup=true;vi.sources=scopedSourcesForRequests_(rt,[p.requestId],vi);
      Object.keys(vi.sources).forEach(name=>vi.index[name]=vi.sources[name].index);
      const source=opsSourceForRequest_(record.values,h),m=vi.index[source][p.requestId]||emptyTaskAggregate_(),mine=m.workers[p.worker]||{};
      const adoptingDone=status===APP_CONFIG.STATUS.DONE&&r[4]==='PAUSED';
      if(!adoptingDone&&!(Number(mine['assigned'+p.stage])>0&&Number(mine['done'+p.stage])>=Number(mine['assigned'+p.stage])))throw Error('내게 배정된 VALID VI를 모두 체크한 뒤 완료할 수 있습니다.');
      if(!adoptingDone&&p.stage==='2P'&&type==='3P Listing 후 2P Lite 전환'&&!(Number(mine.assigned3P)>0&&Number(mine.done3P)>=Number(mine.assigned3P)))throw Error('내게 배정된 VALID 3P VI를 모두 완료한 뒤 2P Lite를 진행할 수 있습니다.');
      const allDone=m.total>0&&(type==='3P Listing'?m.done3>=m.total:type==='2P Lite 전환'?m.done2>=m.total:m.both>=m.total);
      const now=new Date(),extra=[],tz=SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
      function requestCell(name,value){if(!h[name])throw Error('REQUESTS 필수 헤더 누락: '+name);extra.push({updateCells:{start:{sheetId:rt.sheet.getSheetId(),rowIndex:record.row-1,columnIndex:h[name]-1},rows:[{values:[sessionCell_(value,tz)]}],fields:value instanceof Date?'userEnteredValue,userEnteredFormat.numberFormat':'userEnteredValue'}});}
      if(p.stage==='2P'){
        if(!h['Barcode Provided'])throw Error('REQUESTS의 Barcode Provided 열이 없습니다.');
        const raw=String(p.barcodeProvided||'').trim(),supplied=normalizeBarcodeProvidedV1_(raw),current=normalizeBarcodeProvidedV1_(v('Barcode Provided'));
        if(raw&&!supplied)throw Error('바코드 제공 여부는 YES 또는 NO만 입력할 수 있습니다.');
        if(current&&supplied&&current!==supplied)throw Error('바코드 제공 여부는 이미 기록되어 변경할 수 없습니다.');
        if(!current&&!supplied)throw Error('2P 작업 전에 바코드 제공 여부를 선택해 주세요.');
        if(!current)requestCell('Barcode Provided',supplied);
      }
      if(r[4]!=='ACTIVE'){
        // An old PAUSE cannot close a newer ACTIVE. Only proven completion or a
        // verified already-DONE request is adopted; no timestamp is overwritten.
        const marker=completeProof_(r),requestActive=rows.some(x=>String(x[2])===p.requestId&&x[4]==='ACTIVE');
        if(!completeValidEnd_(r))throw Error('기존 종료 기록 확인이 필요합니다. 재계산 후 확인하세요.');
        if(status===APP_CONFIG.STATUS.DONE){
          if(!allDone||requestActive||!(v('Completed At') instanceof Date)||!Number.isFinite(v('Completed At').getTime()))throw Error('DONE의 실제 완료 상태가 일치하지 않습니다.');
        }else if(!marker||receipt&&receipt.complete&&receipt.complete.requestDone)throw Error('진행 중인 작업 세션이 아닙니다. 최신 상태를 확인하세요.');
        receipt=completeReceipt_(p,r,status===APP_CONFIG.STATUS.DONE,v('Completed At'),!marker,marker||p.operationId);
        return completeReadCommitted_(p,receipt,rows,rt)||sessionRetry_('VERIFYING');
      }
      if(rows.filter(x=>String(x[1])===p.worker&&x[4]==='ACTIVE').length!==1)throw Error('작업자에게 ACTIVE 세션이 중복되어 있습니다. 관리자 확인이 필요합니다.');
      if(status===APP_CONFIG.STATUS.DONE&&!allDone)throw Error('DONE 요청의 실제 Task 완료 상태가 일치하지 않습니다.');
      const completedAt=allDone&&status===APP_CONFIG.STATUS.DONE&&v('Completed At') instanceof Date&&Number.isFinite(v('Completed At').getTime())?v('Completed At'):now;
      const targets=allDone?rows.map((row,index)=>({row,index})).filter(x=>String(x.row[2])===p.requestId&&x.row[4]==='ACTIVE'):[target];
      const changes=targets.map(function(x){if(!(x.row[5] instanceof Date)||!Number.isFinite(x.row[5].getTime())||x.row[5]>now)throw Error('ACTIVE 세션의 시작 시각을 확인하세요.');const next=finalizeWorkSessionViMetrics_(x.row,now,(x.index===target.index?'사용자 완료':'요청 완료로 자동 종료')+' | COMPLETE:'+p.operationId,vi);return {row:x.index+2,record:x.row,next:next};});
      const own=changes.find(c=>c.row===target.index+2).next;
      if(allDone){if(!h.Status)throw Error('REQUESTS Status 헤더가 없습니다.');extra.push.apply(extra,requestStatusSystemWrites_(rt.sheet,record.row,h.Status,[APP_CONFIG.STATUS.DONE]));requestCell('Completed At',completedAt);}
      receipt=completeReceipt_(p,own,allDone,completedAt,false);context.receipt=receipt;context.extraRequests=extra;
      commitSessionRows_(sheet,changes,null,context);
      return completeReadCommitted_(p,receipt)||sessionRetry_('VERIFYING');
    }catch(error){
      if(context.writeAttempted)return sessionRetry_('VERIFYING');
      return {ok:false,state:'REJECTED',message:String(error.message||error)};
    }
  },legacy?30000:1200);
  result.operationId=p.operationId;result.timing={serverMs:Date.now()-began,lockWaitMs:context.lockWaitMs===undefined?Date.now()-began:context.lockWaitMs};
  if(legacy&&result.state==='REJECTED')throw Error(result.message);
  return result;
}
