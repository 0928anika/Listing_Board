/** A009: separate 2P/3P work, end-to-end listing efficiency and backlog effort. */
function laLayouts_(tabs){
  const layouts={},model=tabs._regular,metrics=model.metrics,workerDays=tabs._workerDays||[],leadSamples=tabs._monthlyLeadTime||[],today=tabs._today||'';
  const sum=(rs,i)=>rs.reduce((n,r)=>n+(typeof r[i]==='number'?r[i]:0),0),show=v=>v===null||v===undefined?'':v;
  function section(list,title,headers,rows,formats,col){const s={title,headers,rows:rows.length?rows.map(r=>r.map(show)):[headers.map((_,i)=>i===0?'기록 없음':'')],formats:formats||{},col:col||2,noTotal:true};list.push(s);return s;}
  const dayHeaders=['기간','작업자/범위','2P VI','2P d','2P VI/d','3P VI','리스팅 d','3P VI/d'];
  const dayFormats={0:'@',1:'@',2:'#,##0',3:'0.00',4:'0.0',5:'#,##0',6:'0.00',7:'0.0'};
  function dayTable(list,title,rows,headers){return section(list,title,headers||dayHeaders,rows.map(r=>[r[0],r[1],r[2],r[12],r[7],r[3],r[13],r[8]]),dayFormats);}
  function median(a){const b=a.slice().sort((a,b)=>a-b),n=b.length;return n?(b[Math.floor((n-1)/2)]+b[Math.floor(n/2)])/2:'';}
  ['월별 비교','주별 비교'].forEach(name=>{const rows=metrics[name]||[],list=layouts[name]=[],weekly=name==='주별 비교',periods=Array.from(new Set(rows.map(r=>r[0]))).sort().reverse(),blocks=[];
    periods.forEach(p=>{let b=blocks[blocks.length-1];if(!b||b.length>=6||weekly&&b[0].slice(0,7)!==p.slice(0,7)){b=[];blocks.push(b);}b.push(p);});
    blocks.forEach(ps=>{const source=ps.map(p=>rows.find(r=>r[0]===p)),items=[['3P VI',3,'#,##0'],['리스팅 d',13,'0.00'],['3P VI/d',8,'0.0'],['2P VI',2,'#,##0'],['2P d',12,'0.00'],['2P VI/d',7,'0.0']];
      const body=items.map(([label,index])=>[label,...source.map(r=>r[index]),index===8?sum(source,13)>0?sum(source,3)/sum(source,13):'':index===7?sum(source,12)>0?sum(source,2)/sum(source,12):'':sum(source,index)]),s=section(list,'내부 3P·2P VI/d | '+ps[0]+' ~ '+ps[ps.length-1],['지표',...ps,'전체'],body);s.rowFormats=items.map(i=>i[2]);});
    if(!blocks.length)dayTable(list,'내부 3P·2P VI/d',[]);
    const lookup=new Map(rows.map(r=>[r[0],r])),changes=periods.map(p=>{const r=lookup.get(p),d=new Date((weekly?p:p+'-01')+'T00:00:00Z');if(weekly)d.setUTCDate(d.getUTCDate()-7);else d.setUTCMonth(d.getUTCMonth()-1);const prior=lookup.get(d.toISOString().slice(0,weekly?10:7)),growth=i=>prior&&prior[i]>0&&typeof r[i]==='number'?r[i]/prior[i]-1:null;return [p,r[8],prior?prior[8]:null,growth(8),r[7],prior?prior[7]:null,growth(7),p===(weekly?laWeek_(today):today.slice(0,7))?'진행 중':'기록 기간'];});
    section(list,'직전 달력기간 대비 3P·2P 생산성',['기간','3P VI/d','직전 3P VI/d','3P 생산성 증감률','2P VI/d','직전 2P VI/d','2P 생산성 증감률','기간 상태'],changes,{0:'@',3:'+0.0%;-0.0%;0.0%',6:'+0.0%;-0.0%;0.0%',7:'@'});
  });
  section(layouts['월별 비교'],'월별 크롤링+리스팅 전체 효율',['월','리스팅 VI','크롤링 d','리스팅 d','크롤링+리스팅 d','리스팅 VI / (크롤링 d + 리스팅 d)'],model.monthly.map(m=>[m.month,m.vi,m.crawlD,m.listingD,m.totalD,m.totalRate]),{0:'@',1:'#,##0',2:'0.00',3:'0.00',4:'0.00',5:'0.0'});
  const worker=layouts['작업자 순위']=[],workerRows=metrics['작업자 월별']||[];
  const periods=Array.from(new Set(workerRows.map(r=>r[0]))).sort().reverse();
  periods.forEach(month=>{const rs=workerRows.filter(r=>r[0]===month);['3P','2P'].forEach((stage,j)=>{const vi=j===0?3:2,rate=j===0?8:7,stageDays=j===0?11:10;const eligible=rs.filter(r=>r[stageDays]>=(tabs._minRankDays||1)&&r[j===0?13:12]>0&&(r[stageDays]===undefined?r[vi]!==0:r[stageDays]>0)).sort((a,b)=>b[rate]-a[rate]||a[1].localeCompare(b[1]));let previous=null,rank=0;
      const ranked=eligible.map((r,i)=>{if(r[rate]!==previous)rank=i+1;previous=r[rate];return [rank,r[1],r[vi],r[j===0?13:12],r[rate]];});if(ranked.length)ranked.push(['TOTAL','',sum(eligible,vi),sum(eligible,j===0?13:12),sum(eligible,j===0?13:12)>0?sum(eligible,vi)/sum(eligible,j===0?13:12):'']);
      const s=section(worker,month+' | '+stage+' VI/d 순위',['순위','작업자','완료 VI','d','VI/d'],ranked,{0:'0',1:'@',2:'#,##0',3:'0.00',4:'0.0'},j===0?2:8);if(j===1)s.alongside=true;
    });dayTable(worker,month+' | 작업자 VI/d',rs);});if(!periods.length)dayTable(worker,'대표 지표 | 작업자 VI/d',[]);
  const daily=layouts['일별 기록']=[],dailyRows=metrics['일별 기록']||[],dailyMonths=Array.from(new Set(dailyRows.map(r=>r[0].slice(0,7)))).sort().reverse();
  dailyMonths.forEach(month=>{const count=new Date(Date.UTC(Number(month.slice(0,4)),Number(month.slice(5,7)),0)).getUTCDate(),dates=Array.from({length:count},(_,i)=>month+'-'+String(i+1).padStart(2,'0')),source=dailyRows.filter(r=>r[0].slice(0,7)===month),people=workerRows.filter(r=>r[0]===month),names=people.map(r=>r[1]);
    const headers=['작업자',...dates.flatMap(d=>[d,d]),'월 합계','월 합계','근무비율','2P d','리스팅 d','2P VI/d','3P VI/d'];
    const body=people.map(p=>{const perDay=dates.flatMap(d=>{const r=source.find(r=>r[0]===d&&laName_(r[1])===laName_(p[1]));return r?[r[2],r[3]]:['',''];});const record=workerDays.find(g=>g.month===month&&!g.external&&laName_(g.worker)===laName_(p[1]));return [p[1],...perDay,p[2],p[3],record?record.ratio:1,p[12],p[13],p[7],p[8]];});
    if(body.length){const d=sum(people,6),totalDates=dates.flatMap(day=>{const rs=source.filter(r=>r[0]===day);return rs.length?[sum(rs,2),sum(rs,3)]:['',''];});body.push(['TOTAL',...totalDates,sum(people,2),sum(people,3),'',sum(people,12),sum(people,13),sum(people,12)?sum(people,2)/sum(people,12):'',sum(people,13)?sum(people,3)/sum(people,13):'']);}
    const formats={0:'@'};for(let i=1;i<headers.length;i++)formats[i]=i<=dates.length*2+2?'#,##0':'0.00';const s=section(daily,month+' | 작업자별 완료 VI 및 VI/d',headers,body,formats);s.subheaders=['',...dates.concat(['월']).flatMap(()=>['2P','3P']),'비율','d','d','2P','3P'];s.headerGroups=dates.concat(['월']).map((_,i)=>({offset:1+i*2,width:2}));
  });if(!daily.length)dayTable(daily,'대표 지표 | 일별 VI/d',[]);
  const ext=layouts['외부 지원']=[];dayTable(ext,'대표 지표 | 외부 작업자 월별 VI/d',metrics['외부 지원']||[]);dayTable(ext,'외부 작업자 주별 VI/d (월요일 시작)',metrics['외부 주별']||[]);dayTable(ext,'외부 작업자 일별 VI/d',metrics['외부 일별']||[],['날짜',...dayHeaders.slice(1)]);ext[ext.length-1].formats={...dayFormats,0:'yyyy-mm-dd'};
  const node=layouts['Node별']=[];dayTable(node,'대표 지표 | Node별 내부 VI/d',metrics['Node별']||[]);const nodes=tabs['Node별'][1],nodeMonths=Array.from(new Set(nodes.map(r=>r[0]))).sort().reverse();
  nodeMonths.forEach(month=>{const rows=nodes.filter(r=>r[0]===month).map(r=>{const g=(tabs._nodeCompletion||{})[JSON.stringify([month,r[1]])]||{current:0,prior:0,unknown:0},n=g.current+g.prior+g.unknown;return [r[1],r[2],r[3],g.current,g.prior,g.unknown,n?g.prior/n:''];});if(rows.length){const samples=leadSamples.filter(s=>s.month===month);rows.push(['TOTAL',sum(rows,1),sum(rows,2),sum(rows,3),sum(rows,4),sum(rows,5),sum(rows,2)>0?sum(rows,4)/sum(rows,2):'']);}section(node,month+' | 요청 유입·완료 (전체 요청)',['Node','접수','완료','당월 접수 완료','이전 월 접수 완료','접수일 미확인/오류','이월 완료율'],rows,{0:'@',1:'#,##0',2:'#,##0',3:'#,##0',4:'#,##0',5:'#,##0',6:'0.0%'});});
  const compare=layouts['내부·외부 비교']=[];dayTable(compare,'대표 지표 | 내부·외부 VI/d',metrics['내부·외부 비교']||[]);
  const time=layouts['작업시간']=[],recorded=tabs._recordedTime||{monthly:[],daily:[]};section(time,'근무자별 월 기록시간 (확인용)',['월','작업자','구분','크롤링(h)','3P(h)','2P(h)','합계(h)'],recorded.monthly.map(r=>r.slice(0,7)),{0:'@',1:'@',2:'@',3:'0.00',4:'0.00',5:'0.00',6:'0.00'});
  Array.from(new Set(recorded.monthly.map(r=>r[0]))).sort().reverse().forEach(month=>{const days=new Date(Date.UTC(Number(month.slice(0,4)),Number(month.slice(5,7)),0)).getUTCDate(),dates=Array.from({length:days},(_,i)=>month+'-'+String(i+1).padStart(2,'0')),people=recorded.monthly.filter(r=>r[0]===month),source=recorded.daily.filter(r=>r[0].slice(0,7)===month);const body=people.map(p=>[p[1],...dates.map(day=>{const r=source.find(r=>r[0]===day&&laName_(r[1])===laName_(p[1]));return r?r[6]:'';}),p[6]]),formats={0:'@'};for(let i=1;i<=dates.length+1;i++)formats[i]='0.00';section(time,month+' | 근무자별 일별 기록시간(h)',['작업자',...dates,'월 합계(h)'],body,formats);});



  const crawl=layouts['크롤링 효과']=[];
  section(crawl,'월별 C/N 리스팅 효율',['월','C VI','C d','C VI/d','N VI','N d','N VI/d','C/N 배수','C VI 비중'],model.monthly.map(m=>[m.month,m.cVI,m.cD,m.cRate,m.nVI,m.nD,m.nRate,m.cRate>0&&m.nRate>0?m.cRate/m.nRate:null,m.share]),{0:'@',1:'#,##0',2:'0.00',3:'0.0',4:'#,##0',5:'0.00',6:'0.0',7:'0.00"×"',8:'0.0%'});
  section(crawl,'월별 크롤링 투입 및 처리량',['월','크롤링 VI','크롤링 d','크롤링 VI/d','리스팅 VI','리스팅 d','리스팅 VI / (크롤링 d + 리스팅 d)'],model.monthly.map(m=>[m.month,m.crawlVI,m.crawlD,m.crawlRate,m.vi,m.listingD,m.totalRate]),{0:'@',1:'#,##0',2:'0.00',3:'0.0',4:'#,##0',5:'0.00',6:'0.0'});
  section(crawl,'월별 크롤링 작업자 물량',['월','크롤링 작업자','크롤링 VI','크롤링 d','크롤링 VI/d','완료 크롤링 요청 수','동일 요청 현재 리스팅 VI'],model.crawlRows.map(g=>[g.month,g.worker,g.vi,g.d,g.d>0?g.vi/g.d:null,g.count,g.listing]),{0:'@',1:'@',2:'#,##0',3:'0.00',4:'0.0',5:'0',6:'#,##0'});
  const effectHeaders=['월','작업자','C VI','C d','C VI/d','N VI','N d','N VI/d','C/N 배수'],effectFormats={0:'@',1:'@',2:'#,##0',3:'0.00',4:'0.0',5:'#,##0',6:'0.00',7:'0.0',8:'0.00"×"'};
  section(crawl,'같은 작업자 C/N 월별 비교',effectHeaders,model.workers,effectFormats);section(crawl,'같은 Node C/N 월별 비교',['월','Node',...effectHeaders.slice(2)],model.nodes,effectFormats);
  function leadRows(byNode){const groups=new Map();leadSamples.forEach(r=>{const key=JSON.stringify(byNode?[r.month,r.node]:[r.month]);if(!groups.has(key))groups.set(key,[]);groups.get(key).push(r);});return Array.from(groups.keys()).sort().reverse().map(key=>{const all=groups.get(key),valid=all.filter(r=>r.adjusted!==null);return [...JSON.parse(key),all.length,median(all.map(r=>r.hours/24)),median(valid.map(r=>r.adjusted/24)),valid.length];});}
  const leadHeaders=['월','완료 요청 수','리드타임 중앙값 (전체·달력일)','홀드 제외 리드타임 중앙값 (유효 표본·달력일)','홀드 제외 표본 수'];
  section(layouts['월별 비교'],'요청 리드타임 및 홀드 제외 리드타임',leadHeaders,leadRows(false),{0:'@',1:'0',2:'0.00',3:'0.00',4:'0'});
  section(layouts['Node별'],'Node별 리드타임 및 홀드 제외 리드타임',['월','Node',...leadHeaders.slice(1)],leadRows(true),{0:'@',1:'@',2:'0',3:'0.00',4:'0.00',5:'0'});
  const plan=layouts['투입 계획']=[],b=model.backlog,gs=Object.values(b),active=[b.ready,b.pending,b.manual],sumField=(rs,key)=>rs.some(r=>r[key]===null)?null:rs.reduce((n,r)=>n+r[key],0);
  section(plan,'현재 남은 요청·VI와 필요 투입 d',['범위','남은 요청 수','남은 VI (평균 추정 포함)','크롤링 d','리스팅 d','필요 총 d'],[
    ['전체 (홀드 재개 가정)',sumField(gs,'count'),sumField(gs,'vi'),model.current.allCrawlD,model.current.allListingD,model.current.allD],
    ['진행 가능',sumField(active,'count'),sumField(active,'vi'),model.current.crawlD,model.current.listingD,model.current.totalD],
    ['ON HOLD (재개 후)',b.held.count,b.held.vi,model.current.heldCrawlD,model.current.heldListingD,model.current.heldD]
  ],{0:'@',1:'0',2:'#,##0',3:'0.00',4:'0.00',5:'0.00'});
  section(plan,'현재 잔량 계산에 사용한 최근 생산성',['공정','기준 월','VI/d'],[['크롤링',model.basis.crawlRate,model.monthly.find(m=>m.month===model.basis.crawlRate)?.crawlRate],['C 리스팅',model.basis.cRate,model.monthly.find(m=>m.month===model.basis.cRate)?.cRate],['N 리스팅',model.basis.nRate,model.monthly.find(m=>m.month===model.basis.nRate)?.nRate]],{0:'@',1:'@',2:'0.0'});
  section(plan,'월별 생산성으로 현재 잔량 처리 시 필요 d',['기준 월','크롤링 d (진행 가능)','리스팅 d (진행 가능)','필요 총 d (진행 가능)','홀드 재개 시 추가 d','전체 필요 d'],model.forecast.map(m=>[m.month,m.crawlD,m.listingD,m.totalD,m.heldD,m.allD]),{0:'@',1:'0.00',2:'0.00',3:'0.00',4:'0.00',5:'0.00'});
  section(plan,'전체 요청당 평균 VI 기준 월별 필요 d',['기준 월','공통 요청당 평균 VI','크롤링 없이 1개 d','크롤링 없이 5개 d','크롤링 해서 1개 d','크롤링 해서 5개 d','크롤링 해서 10개 d'],model.costs.map(m=>[m.month,m.average,m.n1,m.n5,m.c1,m.c5,m.c10]),{0:'@',1:'0.0',2:'0.00',3:'0.00',4:'0.00',5:'0.00',6:'0.00'});
  Object.values(layouts).forEach(sections=>{let next=6,last=null;sections.forEach(s=>{s.row=s.alongside&&last?last.row:next;next=Math.max(next,s.row+s.rows.length+(s.subheaders?3:2)+3);last=s;});});return layouts;
}

function laClearReportMerges_(sheet){
  // A merged title can extend beyond getDataRange()'s last populated column.
  // Unmerge each complete merged range, including otherwise empty title cells.
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).getMergedRanges().forEach(range=>range.breakApart());
}
function laReportValue_(value,header){
  if(header==='날짜'){
    const day=typeof value==='number'?laDay_(value):value;
    if(typeof day==='string'&&/^\d{4}-\d{2}-\d{2}$/.test(day))return Date.parse(day+'T00:00:00Z')/86400000+25569;
  }
  // Period labels are text. Sheets may otherwise coerce them into date serials.
  if(typeof value==='string'&&/^\d{4}-\d{2}(?:-\d{2})?$/.test(value))return "'"+value;
  return typeof value==='string'&&/^[=+@-]/.test(value)?"'"+value:value;
}
function laRenderReports_(ss,layouts,updated){
  Object.keys(layouts).forEach(name=>{
    const sections=layouts[name],height=Math.max(10,...sections.map(s=>s.row+s.rows.length+5)),width=Math.max(13,...sections.map(s=>s.col+s.headers.length));
    let sheet=ss.getSheetByName(name);if(!sheet)sheet=ss.insertSheet(name);
    if(sheet.getMaxRows()<height)sheet.insertRowsAfter(sheet.getMaxRows(),height-sheet.getMaxRows());if(sheet.getMaxColumns()<width)sheet.insertColumnsAfter(sheet.getMaxColumns(),width-sheet.getMaxColumns());
    laClearReportMerges_(sheet);sheet.clear();sheet.setHiddenGridlines(true);sheet.setFrozenRows(3);sheet.setFrozenColumns(1);sheet.setTabColor(name==='외부 지원'?'#b7791f':'#304e76');
    sheet.getRange(1,1,height,width).setBackground('#f4f7fb').setFontFamily('Arial').setFontSize(10).setFontColor('#24354b').setVerticalAlignment('middle');sheet.setColumnWidth(1,24);sheet.setColumnWidths(2,width-1,112);sheet.setColumnWidth(2,170);sheet.setRowHeights(1,height,24);
    sheet.getRange(1,2,1,width-2).merge().setValue('LISTING  /  '+name).setFontSize(18).setFontWeight('bold').setFontColor('#183555');sheet.setRowHeight(1,38);
    sheet.getRange(2,2,1,width-2).merge().setValue('최근 갱신 '+updated+' KST  |  A009  |  수치는 원본 가용 이력 기준').setFontColor('#62758b');
    const note=name==='작업시간'?'개인별 기록 h 확인용 / 생산성 비교·순위에 사용하지 않음':name==='투입 계획'?'VI/d 기반 필요 근무일 d / 크롤링 방식은 크롤링+리스팅 포함 / 잔량 VI 미확정은 전체 요청 평균으로 추정 / 빈칸은 해당 생산성 표본 없음':name==='크롤링 효과'?'C=CRAWL_COMPLETED, 그 외 N / C d·N d는 해당 타겟 리스팅 투입 / 크롤링 d는 크롤링 작업 투입': 'CONFIG 근무비율 적용 / 2P·3P 각각의 기록일로 VI/d 계산 / 진행 중 기간은 부분 실적';
    sheet.getRange(3,2,1,width-2).merge().setValue(note).setFontSize(9).setFontColor('#62758b');
    sections.forEach(s=>{
      const w=s.headers.length,r=s.row,c=s.col,bodyRow=r+(s.subheaders?3:2);
      sheet.getRange(r,c,1,w).merge().setValue((s.reference?'참고 | ':'')+s.title).setFontWeight('bold').setFontSize(11).setFontColor('#183555');sheet.setRowHeight(r,30);
      sheet.getRange(r+1,c,1,w).setNumberFormat('@').setValues([s.headers.map(v=>laReportValue_(v,''))]).setBackground('#304e76').setFontColor('#ffffff').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');sheet.setRowHeight(r+1,36);
      if(s.subheaders){sheet.getRange(r+2,c,1,w).setNumberFormat('@').setValues([s.subheaders.map(v=>laReportValue_(v,''))]).setBackground('#d8e5f4').setFontWeight('bold').setHorizontalAlignment('center');sheet.getRange(r+1,c,2,1).merge();const groups=s.headerGroups||Array.from({length:(w-1)/2},(_,i)=>({offset:1+i*2,width:2}));groups.forEach(g=>sheet.getRange(r+1,c+g.offset,1,g.width).merge());sheet.setColumnWidths(c+1,w-1,76);}
      const body=sheet.getRange(bodyRow,c,s.rows.length,w);body.setNumberFormat('#,##0.00');
      // Set temporal column formats BEFORE writing numeric date values.
      s.headers.forEach((header,i)=>{if(header==='날짜')sheet.getRange(bodyRow,c+i,s.rows.length,1).setNumberFormat('yyyy-mm-dd');else if(header==='월')sheet.getRange(bodyRow,c+i,s.rows.length,1).setNumberFormat('yyyy-mm');});
      body.setValues(s.rows.map(row=>row.map((v,i)=>laReportValue_(v,s.headers[i])))).setHorizontalAlignment('right');
      body.setBackgrounds(s.rows.map((_,i)=>Array(w).fill(i%2===0?'#ffffff':'#eaf0f8')));sheet.getRange(bodyRow,c,s.rows.length,1).setHorizontalAlignment('left');
      s.rows.forEach((row,i)=>{if(row[0]==='TOTAL')sheet.getRange(bodyRow+i,c,1,w).setBackground('#d8e5f4').setFontWeight('bold');});
      Object.keys(s.formats).forEach(i=>sheet.getRange(bodyRow,c+Number(i),s.rows.length,1).setNumberFormat(s.formats[i]));
      if(s.rowFormats)s.rowFormats.forEach((fmt,i)=>{sheet.getRange(bodyRow+i,c+1,1,w-1).setNumberFormat(fmt);if(String(s.rows[i][0]).indexOf('합계')>=0)sheet.getRange(bodyRow+i,c,1,w).setBackground('#d8e5f4').setFontWeight('bold');});
      sheet.getRange(r+1,c,s.rows.length+(s.subheaders?2:1),w).setBorder(true,true,true,true,true,true,'#d6dfeb',SpreadsheetApp.BorderStyle.SOLID);body.setWrap(true);
      if(s.noteBlock)s.rows.forEach((_,i)=>{sheet.getRange(bodyRow+i,c+1,1,width-c).merge().setHorizontalAlignment('left').setWrap(true);sheet.setRowHeight(bodyRow+i,48);});
    });
    if(name==='데이터 점검'){sheet.setColumnWidth(3,180);sheet.setColumnWidth(4,650);sheet.getRange(6,2,height-5,3).setWrap(true);}
  });
}


