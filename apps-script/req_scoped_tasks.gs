// Execution-local only. A scoped source never implies coverage of other requests.
function readTasksForRequests_(sourceName, requestIds, executionContext) {
  const context=executionContext||{}, ids=Array.from(new Set(requestIds.map(String)));
  if(!context.scopedTasks)context.scopedTasks=Object.create(null);
  let table=context.scopedTasks[sourceName];
  if(!table){
    const sheet=getRequiredSheet_(sourceName),headers=getHeaderMap_(sheet);
    table={sheet:sheet,headers:headers,adapter:taskAdapter_(headers),rows:[],rowNumbers:[],byId:Object.create(null),loadedRequestIds:new Set(),positions:Object.create(null),index:Object.create(null),scoped:true};
    context.scopedTasks[sourceName]=table;
  }
  const missing=ids.filter(function(id){return !table.loadedRequestIds.has(id);});
  if(!missing.length)return table;
  const sheet=table.sheet,last=sheet.getLastRow(),idCol=table.headers['Request ID'],width=sheet.getLastColumn();
  const positions=Object.create(null);missing.forEach(function(id){positions[id]=[];});
  const lookupStarted=context.readMetrics?Date.now():0,lookupStrategy='idColumn';
  if(last>=2){
    if(!table.idRows)table.idRows=sheet.getRange(2,idCol,last-1,1).getValues().map(function(row){return String(row[0]||'');});
    table.idRows.forEach(function(id,i){if(positions[id])positions[id].push(i+2);});
  }
  if(context.readMetrics)context.readMetrics.push({phase:'lookup',strategy:lookupStrategy,ms:Date.now()-lookupStarted});
  const wanted=Object.create(null),groups=[],records=Object.create(null);
  missing.forEach(function(id){records[id]=[];positions[id].forEach(function(row){wanted[row]=id;});});
  Object.keys(wanted).map(Number).sort(function(a,b){return a-b;}).forEach(function(row){
    const g=groups[groups.length-1];if(g&&g.start+g.count===row)g.count++;else groups.push({start:row,count:1});
  });
  // Do not publish partial coverage if a read or row identity validation fails.
  groups.forEach(function(g){sheet.getRange(g.start,1,g.count,width).getValues().forEach(function(values,i){
    const row=g.start+i,id=String(values[idCol-1]||'');
    if(id!==wanted[row])throw Error('Task 행 위치가 변경되었습니다. 다시 시도하세요.');
    records[id].push({row:row,values:values});
  });});
  missing.forEach(function(id){table.byId[id]=records[id];table.positions[id]=positions[id];table.loadedRequestIds.add(id);});
  rebuildScopedTaskRows_(table);
  return table;
}

function rebuildScopedTaskRows_(table) {
  const records=Object.keys(table.byId).flatMap(function(id){return table.byId[id];}).sort(function(a,b){return a.row-b.row;});
  table.rows=records.map(function(r){return r.values;});table.rowNumbers=records.map(function(r){return r.row;});
  const source={};source[table.sheet.getName()]=table;
  table.index=buildTaskIndex_(source)[table.sheet.getName()];
}

// Sidebar reads normally contain only a handful of Request IDs. Locate those
// rows inside Sheets instead of transferring the entire Request ID column.
// Larger lookups keep the single-column strategy because one very large
// regular expression can be slower and is unnecessary for the Sidebar.
function readSidebarTasksForRequests_(sourceName, requestIds, executionContext) {
  const context=executionContext||{},ids=Array.from(new Set(requestIds.map(String)));
  if(!ids.length||ids.length>64||ids.some(function(id){return !id;}))return readTasksForRequests_(sourceName,ids,context);
  if(!context.scopedTasks)context.scopedTasks=Object.create(null);
  let table=context.scopedTasks[sourceName];
  if(!table){
    const sheet=getRequiredSheet_(sourceName),headers=getHeaderMap_(sheet);
    table={sheet:sheet,headers:headers,adapter:taskAdapter_(headers),rows:[],rowNumbers:[],byId:Object.create(null),loadedRequestIds:new Set(),positions:Object.create(null),index:Object.create(null),scoped:true};
    context.scopedTasks[sourceName]=table;
  }
  const missing=ids.filter(function(id){return !table.loadedRequestIds.has(id);});
  if(!missing.length)return table;
  // A previous fallback already paid for the full ID column; reuse it.
  if(table.idRows)return readTasksForRequests_(sourceName,ids,context);
  const escaped=missing.map(function(id){return id.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');});
  // Use a capturing group: Google Sheets uses RE2, where lookarounds and some
  // non-capturing-group variants are unavailable.
  const pattern='^('+escaped.join('|')+')$';
  if(pattern.length>2048)return readTasksForRequests_(sourceName,ids,context);
  const sheet=table.sheet,last=sheet.getLastRow(),idCol=table.headers['Request ID'],width=sheet.getLastColumn();
  const wanted=Object.create(null),records=Object.create(null),positions=Object.create(null);
  missing.forEach(function(id){wanted[id]=true;records[id]=[];positions[id]=[];});
  const rows=last<2?[]:sheet.getRange(2,idCol,last-1,1).createTextFinder(pattern)
    .matchEntireCell(true).matchCase(true).useRegularExpression(true).findAll()
    .map(function(cell){return cell.getRow();}).sort(function(a,b){return a-b;});
  const groups=[];rows.forEach(function(row){const group=groups[groups.length-1];if(group&&group.start+group.count===row)group.count++;else groups.push({start:row,count:1});});
  // Validate the row identity after the search so a concurrent sort cannot
  // publish partial or mismatched coverage.
  groups.forEach(function(group){sheet.getRange(group.start,1,group.count,width).getValues().forEach(function(values,offset){
    const row=group.start+offset,id=String(values[idCol-1]||'');
    if(!wanted[id])throw Error('Task 행 위치가 변경되었습니다. 다시 시도하세요.');
    records[id].push({row:row,values:values});positions[id].push(row);
  });});
  missing.forEach(function(id){table.byId[id]=records[id];table.positions[id]=positions[id];table.loadedRequestIds.add(id);});
  rebuildScopedTaskRows_(table);
  return table;
}
// Overlay guard-validated edits using real Sheet coordinates, never compact array offsets.
function overlayScopedTaskEdit_(context,sourceName,start,rows) {
  const table=context.scopedTasks&&context.scopedTasks[sourceName];if(!table)return;
  const records=Object.create(null);Object.keys(table.byId).forEach(function(id){table.byId[id].forEach(function(r){records[r.row]=r;});});
  rows.forEach(function(values,i){const record=records[start+i];if(!record)return;
    if(String(record.values[table.adapter.id])!==String(values[table.adapter.id]))throw Error('Task 행 위치가 변경되었습니다. 다시 시도하세요.');
    record.values=values.slice();
  });rebuildScopedTaskRows_(table);
}
function scopedSourcesForRequests_(requestTable,ids,context) {
  const grouped=Object.create(null),sources={};
  ids.forEach(function(id){const records=requestTable.byId[String(id)]||[];if(records.length!==1)throw Error('요청 ID가 없거나 중복되었습니다: '+id);
    const name=opsSourceForRequest_(records[0].values,requestTable.headers);(grouped[name]||(grouped[name]=[])).push(String(id));
  });
  const reader=context&&context.sidebarTaskLookup?readSidebarTasksForRequests_:readTasksForRequests_;
  Object.keys(grouped).forEach(function(name){sources[name]=reader(name,grouped[name],context);});return sources;
}
