// 고빈도 편집은 버리고, 종결 상태와 비정상 복구처럼 운영상 필요한 사건만 남김.
function shouldLogActivityV2_(action,afterValue){
 const name=String(action||'').trim(),after=String(afterValue==null?'':afterValue).trim();
 if(name==='REQUEST_COMPLETED'||name==='SUBMISSION_PARTIAL_ROLLBACK'||name==='DROP_ACTIVE_SESSIONS_REPAIRED')return true;
 if(name==='STATUS_CHANGED'||name==='REQUEST_STATUS_CHANGED')return normalizeRequestStatusV2_(after)===APP_CONFIG.STATUS.DONE||normalizeRequestStatusV2_(after)===APP_CONFIG.STATUS.DROP;
 return false;
}

function logActivity_(requestId,action,beforeValue,afterValue){
 if(!shouldLogActivityV2_(action,afterValue))return false;
 try{getRequiredSheet_(APP_CONFIG.SHEETS.ACTIVITY_LOG).appendRow([new Date(),getCurrentUser_(),requestId||'',action||'',serializeLogValue_(beforeValue),serializeLogValue_(afterValue)]);return true;}
 catch(error){console.warn('Activity log write failed; primary operation may already be saved.');return false;}
}

// 더 이상 사용하지 않는 과거 목표 VI 분배 시트를 제거함.
function removeDeprecatedAllocationsSheetV1(){
 assertDashboardAdmin_();
 const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName('ALLOCATIONS');
 if(!sheet)return {ok:true,removed:false};
 ss.deleteSheet(sheet);
 return {ok:true,removed:true};
}

// 로그 셀에 객체가 그대로 들어가지 않도록 안전한 문자열로 변환함.
function serializeLogValue_(value) {
  if (value === null || typeof value === 'undefined') return '';
  if (typeof value === 'string') return value;
  try {
    return JSON.stringify(value);
  } catch (error) {
    return String(value);
  }
}
