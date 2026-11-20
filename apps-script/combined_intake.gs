// 요청과 상품 Task를 한 번에 등록하는 통합 팝업을 열기.
function showCombinedIntakeSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('TaskSidebar')
    .setTitle('신규 Listing 요청');
  SpreadsheetApp.getUi().showModalDialog(
    html.setWidth(1200).setHeight(760),
    '신규 Listing 요청 + Task 등록'
  );
}

function getCombinedIntakeBootstrapData() {
  const configSheet = getRequiredSheet_(APP_CONFIG.SHEETS.CONFIG);
  const headers = ensureConfigNodeACodeColumns_();
  const lastRow = Math.max(configSheet.getLastRow(), 2);
  const values = configSheet.getRange(2, 1, lastRow - 1, configSheet.getLastColumn()).getDisplayValues();
  const cleanUnique = function(list) {
    return list.filter(function(value) {
      const text = String(value || '').trim();
      return text && text.indexOf('TODO:') !== 0;
    }).filter(function(value, index, array) { return array.indexOf(value) === index; });
  };
  const aCodesByNode = {};
  values.forEach(function(row) {
    const node = String(row[headers['Node'] - 1] || '').trim();
    const aCode = String(row[headers['A코드'] - 1] || '').trim();
    if (!node || !aCode || aCode.indexOf('TODO:') === 0) return;
    if (!aCodesByNode[node]) aCodesByNode[node] = [];
    if (aCodesByNode[node].indexOf(aCode) < 0) aCodesByNode[node].push(aCode);
  });
  return {
    priorities: cleanUnique(values.map(function(row) { return row[(headers['Priority'] || 2) - 1]; })),
    brands: cleanUnique(values.map(function(row) { return row[(headers['Brand'] || 4) - 1]; })),
    nodes: cleanUnique(values.map(function(row) { return row[headers['Node'] - 1]; })),
    aCodesByNode: aCodesByNode,
  };
}

// 확정된 3P 표준양식 원본을 서식 변경 없이 내려줌.
function getThreePTemplateDownloadV2() {
  return {
    name: '3P_Listing_표준양식.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    base64: getThreePTemplateBase64V2_()
  };
}

function getThreePTemplateBase64V2_() {
  return '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';
}

function ensureConfigNodeACodeColumns_() {
  const sheet = getRequiredSheet_(APP_CONFIG.SHEETS.CONFIG);
  let headers = getHeaderMap_(sheet);
  ['Node','A코드'].forEach(function(name) {
    if (headers[name]) return;
    const column = sheet.getLastColumn() + 1;
    sheet.getRange(1, column).setValue(name)
      .setBackground('#263650').setFontColor('#FFFFFF').setFontWeight('bold');
    headers[name] = column;
  });
  sheet.setColumnWidth(headers['Node'], 150);
  sheet.setColumnWidth(headers['A코드'], 150);
  return headers;
}

function applyNodeACodeConfigUpdate() {
  ensureConfigNodeACodeColumns_();
  SpreadsheetApp.getActive().toast('CONFIG에 Node별 A코드 입력 열을 준비했습니다.', 'Listing Ops', 4);
}

function ensureRequestIntakeMetadataColumnsV2_() {
  const sheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
  const headers = getHeaderMap_(sheet);
  ['Source','A코드','원본 요청 파일','자료 준비 상태','크롤링 필요','크롤링 단계','크롤링 결과 파일','크롤링 작업자','크롤링 완료일','Submission ID','Submission File Name'].forEach(function(name) {
    if (headers[name]) return;
    const column = sheet.getLastColumn() + 1;
    sheet.getRange(1,column).setValue(name).setBackground('#263650').setFontColor('#FFFFFF').setFontWeight('bold');
    headers[name] = column;
  });
  ['Submission ID','Submission File Name'].forEach(function(name){if(headers[name])sheet.hideColumns(headers[name]);});
  return headers;
}

function extractDriveFolderIdV2_(value) {
  const text = String(value || '').trim();
  const match = text.match(/\/folders\/([a-zA-Z0-9_-]+)/) || text.match(/^([a-zA-Z0-9_-]{10,})$/);
  return match ? match[1] : '';
}

function getOriginalRequestFolderV2_() {
  const props = PropertiesService.getDocumentProperties();
  const folderId = props.getProperty('LISTING_ORIGINAL_FILE_FOLDER_ID') || '';

  // 관리자가 이미 저장 폴더를 지정한 경우
  if (folderId) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      folder.getName(); // 실제 접근 권한 확인
      return folder;
    } catch (err) {
      throw new Error(
        '원본 파일 저장 폴더에 접근할 수 없습니다. ' +
        '관리자에게 저장 폴더 권한을 요청해 주세요.'
      );
    }
  }

  // 아직 아무 폴더도 설정하지 않은 최초 상태에서만 생성
  const folder = DriveApp.createFolder('Listing Ops 원본 요청 파일');
  props.setProperty('LISTING_ORIGINAL_FILE_FOLDER_ID', folder.getId());
  return folder;
}

function configureOriginalRequestFolderV2() {
  const ui = SpreadsheetApp.getUi();
  const current = getOriginalRequestFolderV2_();
  const response = ui.prompt('원본 파일 저장 폴더 설정',
    '현재 폴더: '+current.getName()+'\n'+current.getUrl()+'\n\n변경할 개인/공유 드라이브 폴더 URL을 붙여넣으세요.',
    ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const folderId = extractDriveFolderIdV2_(response.getResponseText());
  if (!folderId) throw new Error('Google Drive 폴더 URL 또는 폴더 ID를 확인해 주세요.');
  const folder = DriveApp.getFolderById(folderId);
  folder.getName();
  PropertiesService.getDocumentProperties().setProperty('LISTING_ORIGINAL_FILE_FOLDER_ID',folderId);
  ui.alert('저장 폴더가 변경되었습니다.\n'+folder.getName()+'\n'+folder.getUrl());
}

function requestOriginalFileNameV2_(fileData,requestId) {
  const sourceName=fileData&&fileData.name?fileData.name:(fileData&&fileData.blob&&fileData.blob.getName?fileData.blob.getName():'upload.bin');
  const safeName=String(sourceName).replace(/[\\/:*?"<>|]/g,'_');
  return String(requestId)+'_'+safeName;
}

function saveOriginalRequestFileV2_(fileData, requestId) {
  if (!fileData || (!fileData.base64 && !fileData.blob)) throw new Error('업로드 원본 파일을 찾지 못했습니다.');
  const safeName = requestOriginalFileNameV2_(fileData,requestId);
  const blob = fileData.blob
    ? fileData.blob.copyBlob().setName(safeName)
    : Utilities.newBlob(Utilities.base64Decode(fileData.base64),String(fileData.mimeType || 'application/octet-stream'),safeName);
  return getOriginalRequestFolderV2_().createFile(blob);
}

function applyTwoPTemplateUpdate() {
  configure2PTasksSheet_();
  return '2P_TASKS 외부 사이트 A~I 업로드 형식 반영 완료';
}

// 현재 Task 헤더 위치에 맞춰 한 행을 만들고, 새 Memo 열을 포함한 미지정 열은 빈칸으로 둠.
function buildIntakeTaskRowV1_(headers,columnCount,values,sheetName) {
  const row=Array(columnCount).fill('');
  Object.keys(values).forEach(function(name){
    const column=headers[name];
    if(!column)throw new Error((sheetName||'Task')+' 필수 열을 찾을 수 없습니다: '+name);
    row[column-1]=values[name];
  });
  if(headers['Memo'])row[headers['Memo']-1]='';
  return row;
}

// 숫자로 입력한 가격은 기존 숫자 저장을 유지하고, 안내 문구는 내부 공백/기호를 보존함.
function intakePriceCellValueV1_(value) {
  const text=String(value==null?'':value).trim();
  const numeric=text.replace(/[\s,₩원%]/g,'');
  if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(numeric) && Number.isFinite(Number(numeric))) return Number(numeric);
  // 자유 입력 문구가 '='로 시작해도 시트 수식으로 실행하지 않음.
  return text.charAt(0)==='=' ? "'"+text : text;
}

function submissionIdV1_(value) {
  const id=String(value||'').trim();
  if(!/^SUB-[A-Za-z0-9_-]{12,100}$/.test(id))throw new Error('제출 식별자를 확인할 수 없습니다. 다시 열어 등록해 주세요.');
  return id;
}

function submissionTaskRowsV1_(requestId) {
  const names=[APP_CONFIG.SHEETS.TASKS,APP_CONFIG.SHEETS.CRAWL_TASKS,APP_CONFIG.SHEETS.TASKS_2P];
  const ranges=names.map(function(name){return "'"+name.replace(/'/g,"''")+"'!A2:A";});
  const result=Sheets.Spreadsheets.Values.batchGet(SpreadsheetApp.getActiveSpreadsheet().getId(),{
    ranges:ranges,majorDimension:'ROWS',valueRenderOption:'UNFORMATTED_VALUE'
  });
  if(!result||!Array.isArray(result.valueRanges)||result.valueRanges.length!==names.length)throw new Error('Task 저장 상태를 확인하지 못했습니다. 데이터는 삭제하지 않았습니다.');
  const rowsBySheet=Object.create(null),counts=Object.create(null);
  names.forEach(function(name,index){
    const rows=[];((result.valueRanges[index]||{}).values||[]).forEach(function(row,i){if(String(row[0]||'')===requestId)rows.push(i+1);});
    rowsBySheet[name]=rows;counts[name]=rows.length;
  });
  return {names:names,rowsBySheet:rowsBySheet,counts:counts};
}

function deleteSubmissionTaskRowsV1_(taskState) {
  const ss=SpreadsheetApp.getActiveSpreadsheet(),requests=[];
  taskState.names.forEach(function(name){
    const sheet=getRequiredSheet_(name),indices=taskState.rowsBySheet[name].slice().sort(function(a,b){return a-b;}),runs=[];
    indices.forEach(function(index){const last=runs[runs.length-1];if(last&&last.end===index)last.end=index+1;else runs.push({start:index,end:index+1});});
    runs.reverse().forEach(function(run){requests.push({deleteDimension:{range:{sheetId:sheet.getSheetId(),dimension:'ROWS',startIndex:run.start,endIndex:run.end}}});});
  });
  if(requests.length)Sheets.Spreadsheets.batchUpdate({requests:requests},ss.getId());
}

function trashSubmissionOriginalV1_(fileUrl,fileName) {
  const id=String(fileUrl||'').match(/\/d\/([A-Za-z0-9_-]+)/);
  if(id){DriveApp.getFileById(id[1]).setTrashed(true);return;}
  if(!fileName)return;
  const files=getOriginalRequestFolderV2_().getFilesByName(String(fileName));
  while(files.hasNext())files.next().setTrashed(true);
}

// Script Lock을 이미 보유한 상태에서 Submission 저장 상태를 판별/롤백함.
function resolveCombinedListingSubmissionLockedV1_(id) {
    const requestSheet=getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS),headers=getHeaderMap_(requestSheet);
    if(!headers['Submission ID'])return {state:'FAILED'};
    const rows=requestSheet.getLastRow()>1?requestSheet.getRange(2,1,requestSheet.getLastRow()-1,requestSheet.getLastColumn()).getValues():[];
    const found=[];rows.forEach(function(row,index){if(String(row[headers['Submission ID']-1]||'')===id)found.push({row:index+2,values:row});});
    if(!found.length)return {state:'FAILED'};
    if(found.length!==1)throw new Error('동일 Submission ID가 여러 요청에 존재하여 안전상 삭제하지 않았습니다. 관리자에게 문의해 주세요.');
    const record=found[0],value=function(name){return headers[name]?record.values[headers[name]-1]:'';},requestId=String(value('Request ID')||'');
    const sameRequestId=rows.filter(function(row){return String(row[headers['Request ID']-1]||'')===requestId;}).length;
    if(!requestId||sameRequestId!==1)throw new Error('Request ID 매핑이 유일하지 않아 안전상 삭제하지 않았습니다. 관리자에게 문의해 주세요.');
    const requestType=String(value('Request Type')||''),crawlRequired=value('크롤링 필요')===true,totalVi=Number(value('Total VI')||0);
    const taskState=submissionTaskRowsV1_(requestId),expected=crawlRequired?0:totalVi;
    const expectedSheet=requestType==='2P Lite 전환'?APP_CONFIG.SHEETS.TASKS_2P:(crawlRequired?APP_CONFIG.SHEETS.CRAWL_TASKS:APP_CONFIG.SHEETS.TASKS);
    const totalTasks=taskState.names.reduce(function(total,name){return total+taskState.counts[name];},0);
    const fileSaved=!!String(value('원본 요청 파일')||'');
    if(fileSaved&&totalTasks===expected&&taskState.counts[expectedSheet]===expected) {
      return {state:'COMPLETE',ok:true,requestId:requestId,createdCount:expected,crawlRequired:crawlRequired,requestLinkCreated:true};
    }
    // 삭제 범위는 유일한 Submission ID에서 역참조한 Request ID로만 제한함.
    deleteSubmissionTaskRowsV1_(taskState);
    trashSubmissionOriginalV1_(value('원본 요청 파일'),value('Submission File Name'));
    requestSheet.deleteRow(record.row);
    logActivity_(requestId,'SUBMISSION_PARTIAL_ROLLBACK','',JSON.stringify({submissionId:id,taskCounts:taskState.counts,fileSaved:fileSaved}));
    return {state:'ROLLED_BACK_PARTIAL',requestId:requestId};
}

// 502 등 응답 유실 시 완전 성공/실패를 확인하고 부분 저장만 제출 단위로 롤백함.
function resolveCombinedListingSubmissionV1_(submissionId) {
  const id=submissionIdV1_(submissionId),lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))throw new Error('등록 처리가 아직 진행 중입니다. 잠시 후 자동으로 다시 확인해 주세요.');
  try {
    return resolveCombinedListingSubmissionLockedV1_(id);
  } finally { lock.releaseLock(); }
}

// 공개 진입점: HTML 클라이언트에서 Submission 저장 상태를 확인함.
function resolveCombinedListingSubmissionV1(submissionId) {
  return resolveCombinedListingSubmissionV1_(submissionId);
}

// 필수 데이터 저장 이후 Task 컨트롤과 REQ 링크만 후처리함.
function finalizeCombinedListingPresentationV1_(submissionId) {
  const id=submissionIdV1_(submissionId),lock=LockService.getScriptLock();
  if(!lock.tryLock(30000))throw new Error('다른 등록 후처리를 진행 중입니다.');
  try {
    const requestSheet=getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS),headers=getHeaderMap_(requestSheet);
    if(!headers['Submission ID'])throw new Error('Submission ID 열이 없습니다.');
    const rows=requestSheet.getLastRow()>1?requestSheet.getRange(2,1,requestSheet.getLastRow()-1,requestSheet.getLastColumn()).getValues():[];
    const found=[];rows.forEach(function(row,index){if(String(row[headers['Submission ID']-1]||'')===id)found.push({row:index+2,values:row});});
    if(found.length!==1)throw new Error('후처리할 요청을 유일하게 확인할 수 없습니다.');
    const record=found[0],value=function(name){return headers[name]?record.values[headers[name]-1]:'';};
    const requestId=String(value('Request ID')||''),requestType=String(value('Request Type')||''),crawlRequired=value('크롤링 필요')===true;
    const taskState=submissionTaskRowsV1_(requestId);
    // 클릭 가능한 필터 링크를 먼저 확정하고, 상대적으로 느린 Task 컨트롤은 그 다음 적용함.
    const linked=refreshNewRequestIdLinkV1_(requestId,record.row,requestType,crawlRequired);
    if(!crawlRequired){
      const taskName=requestType==='2P Lite 전환'?APP_CONFIG.SHEETS.TASKS_2P:APP_CONFIG.SHEETS.TASKS;
      const sheet=getRequiredSheet_(taskName),indices=taskState.rowsBySheet[taskName].slice().sort(function(a,b){return a-b;}),runs=[];
      indices.forEach(function(index){const last=runs[runs.length-1];if(last&&last.end===index)last.end=index+1;else runs.push({start:index,end:index+1});});
      runs.forEach(function(run){applyNewTaskControlsV2_(sheet,run.start+1,run.end-run.start,requestType);});
    }
    return {ok:true,requestId:requestId,requestLinkCreated:linked};
  } finally { lock.releaseLock(); }
}

// 공개 진입점: HTML 클라이언트에서 신규 요청의 필터 링크와 Task 컨트롤을 후처리함.
function finalizeCombinedListingPresentationV1(submissionId) {
  return finalizeCombinedListingPresentationV1_(submissionId);
}

function submitCombinedListingRequest(payload) {
  if (!payload) throw new Error('입력 데이터가 없습니다.');
  const requester = String(payload.requester || '').trim();
  const brand = String(payload.brand || '').trim();
  const priority = String(payload.priority || '').trim() || 'NORMAL';
  const notes = String(payload.notes || '').trim();
  const requestType = String(payload.requestType || '3P Listing');
  const node = String(payload.node || '').trim();
  const aCode = String(payload.aCode || '').trim();
  const isTwoPOnly = requestType === '2P Lite 전환';
  const source = isTwoPOnly ? '' : String(payload.source || '').trim();
  const originalFile = payload.originalFile || null;
  const submissionId = submissionIdV1_(payload.submissionId);
  const rawItems = Array.isArray(payload.items) ? payload.items : [];
  const items = rawItems.map(function(item) {
    if (!Array.isArray(item)) return item;
    if (isTwoPOnly) return {
      vendorProductId:item[0], productId:item[1], option3pId:item[2], barcode:item[3],
      vendorProductCode:item[4], coupangDisplayName:item[5], vendorRegisteredName:item[6],
      registeredOptionName:item[7]
    };
    return {productCode:item[0], productName:item[1], size:item[2], color:item[3], salePrice:item[4], discountPrice:item[5]};
  });
  const crawlRequired = !isTwoPOnly && (payload.crawlRequired === true || items.length === 0);
  const intakeConfig = getCombinedIntakeBootstrapData();

  if (!requester) throw new Error('요청자를 입력해 주세요.');
  if (!brand) throw new Error('브랜드를 선택해 주세요.');
  if (intakeConfig.brands.indexOf(brand) < 0) throw new Error('CONFIG에 등록된 브랜드를 선택해 주세요.');
  if (!node) throw new Error('Node를 선택해 주세요.');
  if (intakeConfig.nodes.indexOf(node) < 0) throw new Error('CONFIG에 등록된 Node를 선택해 주세요.');
  const allowedACodes = intakeConfig.aCodesByNode[node] || [];
  if (allowedACodes.length && allowedACodes.indexOf(aCode) < 0) throw new Error('선택한 Node에 등록된 A코드를 선택해 주세요.');
  if (!isTwoPOnly && !source) throw new Error('Source를 입력해 주세요.');
  if (!originalFile) throw new Error('요청 원본 파일을 업로드해 주세요.');
  if (!crawlRequired && !items.length) throw new Error('표준 양식 파일에서 상품 데이터를 찾지 못했습니다.');

  if (!crawlRequired) items.forEach(function(item, index) {
    if (!isTwoPOnly) {
      const required3P = [item.productCode, item.productName, item.size, item.color, item.salePrice, item.discountPrice];
      if (required3P.some(function(value) { return String(value == null ? '' : value).trim() === ''; })) {
        throw new Error((index + 1) + '번째 상품의 필수값 6개를 모두 입력해 주세요.');
      }
    }
  });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('다른 사용자의 등록을 처리 중입니다. 잠시 후 다시 시도해 주세요.');
  try {
    // 동일 Submission ID를 먼저 판별해 응답 유실 후 재전송의 중복 생성을 막음.
    const priorSubmission=resolveCombinedListingSubmissionLockedV1_(submissionId);
    if(priorSubmission.state!=='FAILED')return priorSubmission;
    const metadataHeaders = ensureRequestIntakeMetadataColumnsV2_();
    const createdRequest = createRequest({
      requester: requester, brand: brand, priority: priority, assignee: '',
      node: node, requestType: requestType, totalVi: crawlRequired ? 0 : items.length, notes: notes,
      submissionId:submissionId,submissionFileName:'',source:source,aCode:aCode,
      preparationStatus:crawlRequired?'크롤링 담당자 배정 필요':'바로 리스팅 가능',
      crawlRequired:crawlRequired,crawlStage:crawlRequired?'CRAWL_REQUESTED':'N/A',
      returnDetails:true,skipRefresh: true
    });
    const requestId=createdRequest.requestId,requestRow=createdRequest.row;
    const requestSheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS);
    const submissionFileName=requestOriginalFileNameV2_(originalFile,requestId);
    // Drive 저장 직전에 최종 REQ 파일명을 기록해 URL 기록 실패 시에도 정확히 롤백할 수 있게 함.
    requestSheet.getRange(requestRow,metadataHeaders['Submission File Name']).setValue(submissionFileName);
    const savedOriginalFile = saveOriginalRequestFileV2_(originalFile,requestId);
    requestSheet.getRange(requestRow,metadataHeaders['원본 요청 파일'])
      .setValue(savedOriginalFile.getUrl()).setNote(savedOriginalFile.getUrl());

    if (crawlRequired) {
      logActivity_(requestId, 'CRAWL_REQUEST_CREATED', '', 'VI 미확정 | ' + requestType);
    } else if (isTwoPOnly) {
      const twoPLayout = getTwoPTaskLayoutV8_();
      const sheet2P = twoPLayout.sheet;
      const h2P = twoPLayout.headers;
      // 시트 메타데이터는 등록당 한 번만 읽고 모든 신규 행에 재사용함.
      const width2P = sheet2P.getLastColumn();
      const rows2P = items.map(function(item, index) {
        const internalId = requestId + '-2P-' + String(index + 1).padStart(3, '0');
        const row = Array(width2P).fill('');
        function put(name, value) { if (h2P[name]) row[h2P[name] - 1] = value; }
        put('Request ID', requestId); put('업체상품ID', String(item.vendorProductId || '').trim());
        put('Product ID', String(item.productId || '').trim()); put('3P 옵션 ID', String(item.option3pId || '').trim());
        put('바코드', String(item.barcode || '').trim());
        put('업체상품코드', String(item.vendorProductCode || '').trim()); put('쿠팡노출상품명', String(item.coupangDisplayName || '').trim());
        put('업체 등록 상품명', String(item.vendorRegisteredName || '').trim()); put('등록 옵션명', String(item.registeredOptionName || '').trim());
        put('2P Completed', false); put('VI Status', 'VALID'); put('Work Status', '작업 대기'); put('Internal Task ID', internalId);
        return row;
      });
      const target2P = Math.max(sheet2P.getLastRow() + 1, 2);
      ensureTaskFilterCapacity_(sheet2P,target2P+rows2P.length-1);
      sheet2P.getRange(target2P, 1, rows2P.length, width2P).setValues(rows2P);
    } else {
      const sheet3P = getRequiredSheet_(APP_CONFIG.SHEETS.TASKS);
      const h3P = getHeaderMap_(sheet3P), width3P = sheet3P.getLastColumn();
      const rows3P = items.map(function(item, index) {
        const internalId = requestId + '-VI-' + String(index + 1).padStart(3, '0');
        return buildIntakeTaskRowV1_(h3P,width3P,{
          'Request ID':requestId,'품번':String(item.productCode).trim(),'등록상품명':String(item.productName).trim(),
          '사이즈':String(item.size).trim(),'색상':String(item.color).trim(),'판매가':intakePriceCellValueV1_(item.salePrice),
          '정상가':intakePriceCellValueV1_(item.discountPrice),'Brand':brand,'3P Assignee':'','3P Completed':false,
          '2P Assignee':requestType === '3P Listing' ? 'N/A' : '',
          '2P Completed':requestType === '3P Listing' ? 'N/A' : false,
          'VI Status':'VALID','Invalid Reason':'','Work Status':'작업 대기','Internal Task ID':internalId
        },sheet3P.getName());
      });
      const target3P = Math.max(sheet3P.getLastRow() + 1, 2);
      ensureTaskFilterCapacity_(sheet3P,target3P+rows3P.length-1);
      sheet3P.getRange(target3P, 1, rows3P.length, width3P).setValues(rows3P);
    }

    logActivity_(requestId, 'REQUEST_AND_TASKS_CREATED', '', items.length + ' tasks | ' + requestType);
    return {ok:true, requestId:requestId, createdCount:items.length, crawlRequired:crawlRequired,requestLinkCreated:false};
  } finally {
    lock.releaseLock();
  }
}

// 공개 진입점: HTML form의 파일을 Blob으로 받아 Base64 변환 없이 등록함.
function submitCombinedListingRequestFormV1(form) {
  if(!form||!form.payloadJson)throw new Error('요청 데이터를 확인할 수 없습니다.');
  const payload=JSON.parse(String(form.payloadJson));
  const blob=form.originalUpload;
  if(!blob||typeof blob.getName!=='function'||!blob.getName())throw new Error('요청 원본 파일을 업로드해 주세요.');
  payload.originalFile={blob:blob,name:blob.getName(),mimeType:blob.getContentType()};
  return submitCombinedListingRequest(payload);
}

function showCrawlResultDialog() {
  PropertiesService.getUserProperties().deleteProperty('CRAWL_RESULT_DEFAULT_REQUEST');
  PropertiesService.getUserProperties().deleteProperty('CRAWL_RESULT_DEFAULT_WORKER');
  const html = HtmlService.createHtmlOutputFromFile('CrawlResultSidebar').setWidth(1100).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, '크롤링 결과 등록');
}

function showCrawlResultDialogForRequest(requestId,worker) {
  const props=PropertiesService.getUserProperties();
  props.setProperty('CRAWL_RESULT_DEFAULT_REQUEST',String(requestId||''));
  props.setProperty('CRAWL_RESULT_DEFAULT_WORKER',String(worker||''));
  const html=HtmlService.createHtmlOutputFromFile('CrawlResultSidebar').setWidth(1100).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html,'크롤링 결과 제출');
}

function getPendingCrawlRequests() {
  const sheet = getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS), h = getHeaderMap_(sheet);
  return getDataRows_(sheet).filter(function(row){
    return h['크롤링 필요'] && row[h['크롤링 필요']-1] === true && String(row[h['크롤링 단계']-1]||'') !== 'CRAWL_COMPLETED';
  }).map(function(row){return {
    requestId:String(row[h['Request ID']-1]||''), brand:String(row[h['Brand']-1]||''),
    requestType:String(row[h['Request Type']-1]||''), requestor:String(row[h['Requestor']-1]||'')
  };});
}

function getCrawlResultBootstrapData() {
  const props=PropertiesService.getUserProperties();
  return {
    requests:getPendingCrawlRequests(),
    workers:(getWebConfig_().workers || []).slice(),
    defaultRequest:props.getProperty('CRAWL_RESULT_DEFAULT_REQUEST')||'',
    defaultWorker:props.getProperty('CRAWL_RESULT_DEFAULT_WORKER')||''
  };
}

function submitCrawlResult(payload) {
  if (!payload || !payload.requestId || !payload.originalFile) throw new Error('요청과 크롤링 결과 파일을 확인해 주세요.');
  const crawlWorker=String(payload.worker||'').trim();
  if (!crawlWorker || (getWebConfig_().workers||[]).indexOf(crawlWorker)<0) throw new Error('CONFIG에 등록된 크롤링 작업자를 선택해 주세요.');
  const items = (Array.isArray(payload.items) ? payload.items : []).map(function(item) {
    return Array.isArray(item)
      ? {productCode:item[0],productName:item[1],size:item[2],color:item[3],salePrice:item[4],discountPrice:item[5]}
      : item;
  });
  if (!items.length) throw new Error('크롤링 결과에서 VI를 찾지 못했습니다.');
  items.forEach(function(item,index){
    if ([item.productCode,item.productName,item.size,item.color,item.salePrice,item.discountPrice].some(function(v){return String(v==null?'':v).trim()==='';})) throw new Error((index+1)+'번째 VI의 필수값이 누락되었습니다.');
  });
  const lock=LockService.getScriptLock(); if(!lock.tryLock(30000)) throw new Error('다른 등록 작업을 처리 중입니다.');
  try {
    const requestSheet=getRequiredSheet_(APP_CONFIG.SHEETS.REQUESTS), requestHeaders=getHeaderMap_(requestSheet);
    ['Node','Brand','Request Type','크롤링 필요','크롤링 단계','크롤링 작업자'].forEach(function(name){
      if(!requestHeaders[name])throw new Error('REQUESTS 필수 헤더가 없습니다: '+name);
    });
    const requestRow=findRowByRequestId_(requestSheet,payload.requestId);
    if (requestSheet.getRange(requestRow,requestHeaders['크롤링 필요']).getValue() !== true) throw new Error('크롤링 필요 요청이 아닙니다.');
    if (String(requestSheet.getRange(requestRow,requestHeaders['크롤링 단계']).getValue()) === 'CRAWL_COMPLETED') throw new Error('이미 결과 등록이 완료된 요청입니다.');
    const assignedCrawlWorker=String(requestSheet.getRange(requestRow,requestHeaders['크롤링 작업자']).getValue()||'');
    if (assignedCrawlWorker && assignedCrawlWorker!==crawlWorker) throw new Error('이 요청은 '+assignedCrawlWorker+' 작업자에게 배정되어 있습니다.');
    const node=String(requestSheet.getRange(requestRow,requestHeaders['Node']).getValue()||'').trim();
    const brand=String(requestSheet.getRange(requestRow,requestHeaders['Brand']).getValue()||'');
    const requestType=String(requestSheet.getRange(requestRow,requestHeaders['Request Type']).getValue()||'3P Listing');
    const h=ensureRequestIntakeMetadataColumnsV2_();
    const taskSheet=getRequiredSheet_(APP_CONFIG.SHEETS.CRAWL_TASKS),taskHeaders=getHeaderMap_(taskSheet),taskWidth=taskSheet.getLastColumn();
    const rows=items.map(function(item,index){return buildIntakeTaskRowV1_(taskHeaders,taskWidth,{
      'Request ID':payload.requestId,'품번':String(item.productCode).trim(),'등록상품명':String(item.productName).trim(),
      '사이즈':String(item.size).trim(),'색상':String(item.color).trim(),'판매가':intakePriceCellValueV1_(item.salePrice),
      '정상가':intakePriceCellValueV1_(item.discountPrice),'Brand':brand,'3P Assignee':'','3P Completed':false,
      '2P Assignee':requestType==='3P Listing'?'N/A':'','2P Completed':requestType==='3P Listing'?'N/A':false,
      'VI Status':'VALID','Invalid Reason':'','Work Status':'작업 대기',
      'Internal Task ID':payload.requestId+'-CR-VI-'+String(index+1).padStart(3,'0')
    },taskSheet.getName());});
    const resultFile=saveOriginalRequestFileV2_(payload.originalFile,payload.requestId+'_CRAWL');
    const target=Math.max(taskSheet.getLastRow()+1,2);
    ensureTaskFilterCapacity_(taskSheet,target+rows.length-1);
    taskSheet.getRange(target,1,rows.length,taskWidth).setValues(rows);
    applyNewTaskControlsV2_(taskSheet,target,rows.length,requestType);
    requestSheet.getRange(requestRow,h['Total VI']).setValue(rows.length);
    requestSheet.getRange(requestRow,h['크롤링 단계']).setValue('CRAWL_COMPLETED');
    requestSheet.getRange(requestRow,h['자료 준비 상태']).setValue('크롤링 완료 · 3P 리스팅 가능');
    requestSheet.getRange(requestRow,h['크롤링 결과 파일']).setValue(resultFile.getUrl()).setNote(resultFile.getUrl());
    requestSheet.getRange(requestRow,h['크롤링 작업자']).setValue(crawlWorker);
    requestSheet.getRange(requestRow,h['크롤링 완료일']).setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm');
    if (typeof pauseActiveStageSessions_==='function') pauseActiveStageSessions_(payload.requestId,'CRAWL',new Date(),'크롤링 결과 제출 완료');
    if (typeof bumpMyWorkSidebarRevisionV2_==='function') bumpMyWorkSidebarRevisionV2_([crawlWorker]);
    logActivity_(payload.requestId,'CRAWL_RESULT_CREATED','',rows.length+' VI');
    return {ok:true,requestId:payload.requestId,createdCount:rows.length};
  } finally { lock.releaseLock(); }
}
