/** A005 report layouts. Pure layout planning, followed by destination-only rendering. */
function laLayouts_(tabs){
  const layouts={};
  const workerDays=tabs._workerDays||[];
  function dailySummary(gs){const days=gs.reduce((n,g)=>n+g.days,0),equiv=gs.reduce((n,g)=>n+g.equivalentDays,0),v2=gs.reduce((n,g)=>n+g.v2,0),v3=gs.reduce((n,g)=>n+g.v3,0);return [days,gs.length===1?gs[0].ratio:'',days?v2/days:'',days?v3/days:'',days?(v2+v3)/days:'',equiv?v2/equiv:'',equiv?v3/equiv:'',equiv?(v2+v3)/equiv:''];}
  function section(list,title,headers,rows,formats,col){list.push({title:title,headers:headers,rows:rows.length?rows:[headers.map((_,i)=>i===0?'기록 없음':'')],formats:formats||{},col:col||2});}
  function matrix(list,title,rows,metrics,byStartMonth){
    const periods=Array.from(new Set(rows.map(r=>r[0]))).sort();
    if(!periods.length){section(list,title,['항목','값'],[]);return;}
    const blocks=[];
    if(byStartMonth){periods.forEach(p=>{let block=blocks[blocks.length-1];if(!block||block[0].slice(0,7)!==p.slice(0,7)){block=[];blocks.push(block);}block.push(p);});}
    else for(let i=0;i<periods.length;i+=6)blocks.push(periods.slice(i,i+6));
    blocks.forEach(ps=>{const body=metrics.map(m=>[m[0]].concat(ps.map(p=>{const r=rows.find(r=>r[0]===p);return r?r[m[1]]:'';})));
      const source=rows.filter(r=>ps.includes(r[0]));
      body.forEach((r,j)=>{const ix=metrics[j][1],sum=k=>source.reduce((n,x)=>n+(typeof x[k]==='number'?x[k]:0),0);r.push([2,3,4,5,6,7].includes(ix)?sum(ix):ix===8?(sum(5)>0?sum(2)/sum(5):''):ix===9?(sum(6)>0?sum(3)/sum(6):''):'');});
      section(list,title+' | '+(byStartMonth?ps[0].slice(0,7)+' 시작 주 / ':'')+ps[0]+' ~ '+ps[ps.length-1],['지표'].concat(ps).concat(['전체']),body);list[list.length-1].rowFormats=metrics.map(m=>m[2]||'#,##0.0');
    });
  }
  ['월별 비교','주별 비교'].forEach(n=>{const list=layouts[n]=[],rs=tabs[n][1];
    matrix(list,'01  내부 실적 비교',rs,[['3P 완료 VI',2,'#,##0'],['2P 완료 VI',3,'#,##0'],['완료 VI 합계',4,'#,##0'],['3P 작업시간(h)',5],['2P 작업시간(h)',6],['작업시간 합계(h)',7],['3P 생산성(VI/h)',8],['2P 생산성(VI/h)',9]],n==='주별 비교');
    matrix(list,'02  직전 달력기간 대비',rs,[['완료 VI 증감률',13,'+0.0%;-0.0%;0.0%'],['작업시간 증감률',15,'+0.0%;-0.0%;0.0%'],['작업 REQ 수',11,'#,##0'],['기록일 수',10,'#,##0'],['기간 상태',16,'@']],n==='주별 비교');
  });
  const leadSamples=tabs._monthlyLeadTime||[],leadMonths=Array.from(new Set(tabs['월별 비교'][1].map(r=>r[0]).concat(leadSamples.map(s=>s.month)))).sort();
  function leadStats(samples){const values=samples.map(s=>s.hours).sort((a,b)=>a-b),n=values.length;return n?[(values[Math.floor((n-1)/2)]+values[Math.floor(n/2)])/2,values.reduce((a,b)=>a+b,0)/n,n]:['','',0];}
  if(!leadMonths.length)section(layouts['월별 비교'],'03  요청 리드타임 (전체 완료 요청)',['항목','값'],[]);
  for(let i=0;i<leadMonths.length;i+=6){const months=leadMonths.slice(i,i+6),stats=months.map(m=>leadStats(leadSamples.filter(s=>s.month===m))),all=leadStats(leadSamples.filter(s=>months.includes(s.month))),labels=['리드타임 중앙값(h)','리드타임 평균(h)','유효 완료 요청 수'];
    section(layouts['월별 비교'],'03  요청 리드타임 (전체 완료 요청) | '+months[0]+' ~ '+months[months.length-1],['지표'].concat(months,['전체']),labels.map((label,j)=>[label].concat(stats.map(s=>s[j]),[all[j]])));
    layouts['월별 비교'].at(-1).rowFormats=['#,##0.0','#,##0.0','#,##0'];
  }
  function workerSections(list,rows,label,ranked){
    const periods=Array.from(new Set(rows.map(r=>r[0]))).sort().reverse();
    periods.forEach(p=>{const rs=rows.filter(r=>r[0]===p);if(ranked!==false)['3P','2P'].forEach((stage,i)=>{
      const rank=i===0?12:13,vi=i===0?2:3,hour=i===0?5:6,rate=i===0?8:9;
      const sorted=rs.slice().sort((a,b)=>(a[rank]||Infinity)-(b[rank]||Infinity)||a[1].localeCompare(b[1]));
      section(list,label+' '+p+' | '+stage+' 생산성 순위',['순위','작업자','완료 VI','시간(h)','VI/h'],sorted.map(r=>[r[rank],r[1],r[vi],r[hour],r[rate]]),{0:'#,##0',1:'@',2:'#,##0',3:'0.00',4:'0.0'},i===0?2:8);
      if(i===1)list[list.length-1].alongside=true;
    });
    section(list,label+' '+p+' | 작업시간 및 작업량',['작업자','3P 시간(h)','2P 시간(h)','합계 시간(h)','3P VI','2P VI','기록일','작업 REQ'],rs.slice().sort((a,b)=>b[7]-a[7]).map(r=>[r[1],r[5],r[6],r[7],r[2],r[3],r[10],r[11]]),{0:'@',4:'#,##0',5:'#,##0',6:'#,##0',7:'#,##0'});
    });if(!periods.length)section(list,label,['작업자','값'],[]);
  }
  workerSections(layouts['작업자 월별']=[],tabs['작업자 월별'][1],'월간');
  const daily=layouts['일별 기록']=[],rs=tabs['일별 기록'][1],dailyMonths=Array.from(new Set(rs.map(r=>r[0].slice(0,7)))).sort().reverse();
  dailyMonths.forEach(month=>{
    const count=new Date(Date.UTC(Number(month.slice(0,4)),Number(month.slice(5,7)),0)).getUTCDate(),ds=Array.from({length:count},(_,i)=>month+'-'+String(i+1).padStart(2,'0'));
    const monthlyRows=rs.filter(r=>r[0].slice(0,7)===month),workers=Array.from(new Set(monthlyRows.map(r=>r[1]))).sort();
    const viHeaders=['작업자'].concat(ds.flatMap(d=>[d,d]),['월 합계','월 합계']);
    section(daily,month+' | 작업자별 완료 VI',viHeaders,workers.map(w=>[w].concat(ds.flatMap(d=>{const r=monthlyRows.find(r=>r[0]===d&&r[1]===w);return r?[r[3],r[2]]:['',''];}),[monthlyRows.filter(r=>r[1]===w).reduce((n,r)=>n+r[3],0),monthlyRows.filter(r=>r[1]===w).reduce((n,r)=>n+r[2],0)])),Object.fromEntries(viHeaders.slice(1).map((_,i)=>[i+1,'#,##0'])));
    daily[daily.length-1].subheaders=[''].concat(ds.concat(['月']).flatMap(()=>['2P','3P']));
    const viSection=daily[daily.length-1],baseWidth=viHeaders.length,people=workerDays.filter(g=>g.month===month&&!g.external);
    viSection.headerGroups=ds.concat(['월 합계']).map((_,i)=>({offset:1+i*2,width:2}));
    viSection.headers.push('근무일수','근무비율','1일 평균','1일 평균','1일 평균','보정 1일 평균','보정 1일 평균','보정 1일 평균');
    viSection.subheaders.push('일','비율','2P','3P','TOTAL','2P','3P','TOTAL');
    viSection.headerGroups.push({offset:baseWidth+2,width:3},{offset:baseWidth+5,width:3});
    viSection.rows.forEach(r=>r.push(...dailySummary(people.filter(g=>g.worker===r[0]))));
    for(let j=0;j<8;j++)viSection.formats[baseWidth+j]=j===0?'#,##0':'0.00';
    viSection.summaryTail={offset:baseWidth,values:dailySummary(people)};
    section(daily,month+' | 작업자별 작업시간(h)',['작업자'].concat(ds,['월 합계']),workers.map(w=>[w].concat(ds.map(d=>{const r=monthlyRows.find(r=>r[0]===d&&r[1]===w);return r?r[7]:'';}),[monthlyRows.filter(r=>r[1]===w).reduce((n,r)=>n+r[7],0)])),Object.fromEntries(ds.concat(['월 합계']).map((_,j)=>[j+1,'0.00'])));
  });
  if(!daily.length)section(daily,'일별 기록',['작업자','값'],[]);
  const ext=layouts['외부 지원']=[],external=tabs['외부 지원'][1];
  const monthly=external.filter(r=>r[0]==='월').map(r=>r.slice(1)),weeks=external.filter(r=>r[0]==='주').map(r=>r.slice(1));
  workerSections(ext,monthly,'외부 월간',false);workerSections(ext,weeks,'외부 주간(월요일 시작)',false);
  section(ext,'외부 일별 작업량·시간',['날짜','작업자','3P VI','2P VI','3P 시간(h)','2P 시간(h)'],external.filter(r=>r[0]==='일').map(r=>[r[1],r[2],r[3],r[4],r[6],r[7]]),{0:'yyyy-mm-dd',1:'@',2:'#,##0',3:'#,##0'});
  const ns=layouts['Node별']=[],nodes=tabs['Node별'][1],months=Array.from(new Set(nodes.map(r=>r[0]))).sort().reverse();
  months.forEach(p=>{const rs=nodes.filter(r=>r[0]===p);section(ns,p+' | 요청 유입·완료 (전체 요청)',['Node','접수','완료','리드타임 중앙값(h)','당월 접수 완료','이전 월 접수 완료','접수일 미확인/오류','이월 완료율(딜레이율)'],rs.map(r=>{const g=(tabs._nodeCompletion||{})[JSON.stringify([p,r[1]])]||{current:0,prior:0,unknown:0};const total=g.current+g.prior+g.unknown;return [r[1],r[2],r[3],r[4],g.current,g.prior,g.unknown,total?g.prior/total:''];}),{0:'@',1:'#,##0',2:'#,##0',4:'#,##0',5:'#,##0',6:'#,##0',7:'0.0%'});ns[ns.length-1].delayTotal=true;section(ns,p+' | 내부 작업 실적',['Node','3P VI','2P VI','3P 시간(h)','2P 시간(h)'],rs.map(r=>[r[1],r[5],r[6],r[7],r[8]]),{0:'@',1:'#,##0',2:'#,##0'});});if(!ns.length)section(ns,'Node별',['Node','값'],[]);
  const compare=layouts['내부·외부 비교']=[],compareMonths=Array.from(new Set(workerDays.map(g=>g.month))).sort().reverse();
  compareMonths.forEach(month=>{const people=workerDays.filter(g=>g.month===month),headers=['그룹','작업자 수','근무일수 합계(인일)','환산 근무일수','2P 완료 VI','3P 완료 VI','2P 1인·1일 평균','3P 1인·1일 평균','TOTAL 1인·1일 평균','2P 보정 평균','3P 보정 평균','TOTAL 보정 평균'];
    function row(label,gs){const summary=dailySummary(gs);return [label,gs.length,summary[0],gs.reduce((n,g)=>n+g.equivalentDays,0),gs.reduce((n,g)=>n+g.v2,0),gs.reduce((n,g)=>n+g.v3,0)].concat(summary.slice(2));}
    section(compare,month+' | 내부·외부 1인·1일 작업량',headers,[row('내부',people.filter(g=>!g.external)),row('외부 지원',people.filter(g=>g.external))],{0:'@',1:'#,##0',2:'#,##0',3:'0.00',4:'#,##0',5:'#,##0'});compare[compare.length-1].customTotal=row('TOTAL',people);
  });if(!compare.length)section(compare,'내부·외부 비교',['그룹','값'],[]);
  const crawl=layouts['크롤링 효과']=[],ca=tabs._crawl;
  function crawlSection(title,headers,rows,formats){section(crawl,title,headers,rows,formats);crawl[crawl.length-1].noTotal=true;}
  crawlSection('01  해석 및 집계 기준',['항목','기준'],[
    ['분류','REQUESTS 현재 크롤링 단계만 사용: CRAWL_COMPLETED=크롤링 / N/A=비크롤링 / 나머지=확인 불가. 대소문자·앞뒤 공백 정규화.'],
    ['소급 적용','완료일·크롤링 필요·최초 요청 유형은 사용하지 않음. 현재 단계를 과거 실적에 소급하므로 당시 작업 방식의 확정 이력은 아님.'],
    ['분석 대상','CONFIG 외부 목록을 제외한 내부 작업자의 3P 세션과 COMMITTED 장기 집계. 2P와 CRAWL 준비시간 제외.'],
    ['생산성','시간이 0보다 큰 기록의 순증 VI 합계 ÷ 같은 기록의 작업시간 합계. 0시간 기록의 VI는 물량에만 포함하고 별도 표시. 음수 보정 유지.'],
    ['배수','크롤링 등록 VI/h ÷ 비크롤링 등록 VI/h. 양쪽 모두 양수일 때만 표시. 관찰된 차이이며 크롤링의 인과 효과나 팀 전체 생산성 아님.'],
    ['비중','전체 비중은 확인 불가 물량까지 분모에 포함. 분류 확인 내 비중은 크롤링+비크롤링만 분모. 각 그룹 순증 VI가 음수이거나 분모 0이면 비중 공란.'],
    ['추이','세션 종료일(KST) 귀속. 일별은 기록 있는 날짜만 표시. 최근 7일은 해당 날짜 포함 7개 달력일 합산 VI/합산 시간.'],
    ['비교 한계','같은 작업자·Node도 브랜드, 난이도, 시기 차이는 남음. 그룹 간 동일 요청 중복 집계 가능. 원본 요청이 없으면 확인 불가.'],
    ['크롤링 개인 평균','전체 보유 이력 기준: CRAWL_COMPLETED 요청의 현재 Total VI를 현재 크롤링 작업자에게 1회 귀속. CRAWL 세션 종료일(없으면 시작일)의 작업자별 고유 일수로 나눔. 0시간·ACTIVE 기록도 날짜가 있으면 포함.'],
    ['개인 평균 한계','진행 중 요청의 기록일도 분모에 포함. 여러 요청 동시 작업은 하루로 중복 제거. 완료일 미사용으로 일별 제출 VI나 월별 개인 평균은 추정하지 않음. 현재 Total VI는 제출 당시 물량과 다를 수 있음.'],
    ['백업 누락 주의','현재 장기 집계는 CRAWL을 보존하지 않음. 요청-작업자 연결 기록이 없거나 물량 오류가 있으면 전체 평균 공란, 기록 연결 물량의 참고 평균 표시. 연결률 100%도 일부 작업일 누락까지 보증하지 않음.'],
    ['확인 방법','최종 CONFIG 성공 시각과 A005 버전 확인. 기존 분석 지표의 계산 기준은 유지.']
  ],{0:'@',1:'@'});crawl[0].noteBlock=true;
  if(ca){
    crawlSection('02  월별 전환 비중 및 등록 생산성',['월','전체 순증 VI','크롤링 VI','비크롤링 VI','확인 불가 VI','크롤링 비중(전체)','크롤링 비중(분류 확인 내)','분류 확인 비중','크롤링 등록 VI/h','비크롤링 등록 VI/h','등록 생산성 배수'],ca.trends,{0:'@',1:'#,##0',2:'#,##0',3:'#,##0',4:'#,##0',5:'0.0%',6:'0.0%',7:'0.0%',10:'0.00"배"'});
    crawlSection('03  월별 비교 표본 및 0시간 실적',['월','그룹','순증 VI','작업시간(h)','시간 유효 순증 VI','등록 VI/h','작업 요청 수','기록 행 수','0시간 순증 VI'],ca.monthly,{0:'@',1:'@',2:'#,##0',4:'#,##0',6:'#,##0',7:'#,##0',8:'#,##0'});
    crawlSection('04  일별 및 최근 7일 추이',['날짜','크롤링 VI','비크롤링 VI','확인 불가 VI','크롤링 비중(전체)','크롤링 비중(분류 확인 내)','분류 확인 비중','크롤링 등록 VI/h','비크롤링 등록 VI/h','등록 생산성 배수','전체 순증 VI','최근 7일 크롤링 비중(전체)','최근 7일 크롤링 VI/h','최근 7일 비크롤링 VI/h','최근 7일 생산성 배수'],ca.daily,{0:'yyyy-mm-dd',1:'#,##0',2:'#,##0',3:'#,##0',4:'0.0%',5:'0.0%',6:'0.0%',9:'0.00"배"',10:'#,##0',11:'0.0%',14:'0.00"배"'});
    const comparisonHeaders=['월','작업자','크롤링 시간 유효 VI','크롤링 시간(h)','크롤링 등록 VI/h','비크롤링 시간 유효 VI','비크롤링 시간(h)','비크롤링 등록 VI/h','등록 생산성 배수','크롤링 작업 요청 수','비크롤링 작업 요청 수','확인 불가 VI'];
    const comparisonFormats={0:'@',1:'@',2:'#,##0',5:'#,##0',8:'0.00"배"',9:'#,##0',10:'#,##0',11:'#,##0'};
    crawlSection('05  같은 작업자 월별 비교 (양쪽 표본 확인)',comparisonHeaders,ca.workers,comparisonFormats);
    const nodeHeaders=comparisonHeaders.slice();nodeHeaders[1]='Node';
    crawlSection('06  같은 Node 월별 비교',nodeHeaders,ca.nodes,comparisonFormats);
    crawlSection('07  확인 불가 기록 점검',['월','사유','순증 VI','작업시간(h)','작업 요청 수','기록 행 수'],ca.reasons,{0:'@',1:'@',2:'#,##0',4:'#,##0',5:'#,##0'});
    if(ca.preparation){
      crawlSection('08  크롤링 작업자별 1일 평균 VI (전체 보유 기록)',['크롤링 작업자','완료 요청 총 VI','완료 요청 수','기록일 수','총 VI/기록일','기록 연결 VI','기록 연결 요청 수','요청 기록 연결률','연결 VI/기록일(참고)','첫 기록일','마지막 기록일'],ca.preparation.rows,{0:'@',1:'#,##0',2:'#,##0',3:'#,##0',4:'0.0',5:'#,##0',6:'#,##0',7:'0.0%',8:'0.0',9:'@',10:'@'});
      crawlSection('09  크롤링 개인 평균 누락 점검',['요청 ID','크롤링 작업자','사유','현재 Total VI'],ca.preparation.audit,{0:'@',1:'@',2:'@',3:'#,##0'});
    }
  }
  Object.values(layouts).forEach(sections=>sections.forEach(s=>{
    if(s.noTotal)return;
    if(s.rowFormats||s.rows[0][0]==='기록 없음')return;
    const total=s.headers.map(()=>''),sum=i=>s.rows.reduce((n,r)=>n+(typeof r[i]==='number'?r[i]:0),0);total[0]='TOTAL';
    s.headers.forEach((h,i)=>{if(i===0)return;if(/VI|시간|완료|접수|^\d{4}-\d{2}-\d{2}$|월 합계/.test(h)&&!h.includes('VI/h')&&!h.includes('중앙값'))total[i]=sum(i);});
    const rate=s.headers.indexOf('VI/h');if(rate>=0){const v=s.headers.indexOf('완료 VI'),h=s.headers.indexOf('시간(h)');total[rate]=sum(h)>0?sum(v)/sum(h):'';}
    if(s.summaryTail)s.summaryTail.values.forEach((v,i)=>total[s.summaryTail.offset+i]=v);
    if(s.delayTotal){total[6]=sum(6);total[7]=sum(2)>0?sum(5)/sum(2):'';}
    if(s.customTotal){s.rows.push(s.customTotal);return;}
    if(total.some((v,i)=>i>0&&typeof v==='number'))s.rows.push(total);
  }));
  Object.keys(layouts).forEach(n=>{let next=6,last=null;layouts[n].forEach(s=>{s.row=s.alongside&&last?last.row:next;const end=s.row+s.rows.length+(s.subheaders?3:2);next=Math.max(next,end+3);last=s;});});return layouts;
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
    sheet.getRange(2,2,1,width-2).merge().setValue('최근 갱신 '+updated+' KST  |  A005  |  수치는 원본 가용 이력 기준').setFontColor('#62758b');
    const note=name==='크롤링 효과'?'내부 3P 등록 생산성 / 現 크롤링 단계 소급 분류 / 준비시간 제외 / 인과 효과 아님':name==='외부 지원'?'CONFIG 지정 외부 지원만 표시 · 내부 전체 실적 및 내부 순위에서 제외':name==='내부·외부 비교'?'2P·3P 별도 비교 / 1인·1일 평균 = 총 VI ÷ 작업자별 근무일수 합계 / 보정 = 총 VI ÷ 환산 근무일수 합계':name==='Node별'?'이월 완료율 = 이전 월 접수 완료 ÷ 전체 당월 완료 (납기 위반율 아님) / 접수일 미확인·오류는 별도 표시':'내부 실적 · 외부 지원 제외 / 진행 중인 기간은 부분 실적 / 순위 공란은 최소시간 미달 또는 시간 없음';
    sheet.getRange(3,2,1,width-2).merge().setValue(note).setFontSize(9).setFontColor('#62758b');
    sections.forEach(s=>{
      const w=s.headers.length,r=s.row,c=s.col,bodyRow=r+(s.subheaders?3:2);
      sheet.getRange(r,c,1,w).merge().setValue(s.title).setFontWeight('bold').setFontSize(11).setFontColor('#183555');sheet.setRowHeight(r,30);
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


