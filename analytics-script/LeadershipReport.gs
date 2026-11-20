/** One-off November numeric dashboard. Never called by refreshListingAnalytics or a trigger. */
const CRAWLER_HUB_REPORT_OPTIONS={
  asOf:'2026-11-20',
  minRecordedDays:1,
  exampleMinDays:2,
  septemberStart:'AUTO'
};
function buildCrawlerHubLeadershipReport(){return buildNovemberReport();}
function buildNovemberReport(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),c=laConfig_(ss),o=Object.assign({},CRAWLER_HUB_REPORT_OPTIONS);
  const lock=LockService.getDocumentLock();if(!lock.tryLock(1000))throw Error('다른 갱신이 실행 중입니다. 끝난 뒤 보고서를 만드세요.');
  try{
    const data=laRead_(c.id),report=crhBuild_(data,c,o);
    crhRender_(ss,report);
    console.log(JSON.stringify({report:'11월 보고',asOf:o.asOf,matchedWorkers:report.paired.count,quality:report.quality}));
  }finally{lock.releaseLock();}
}
function crhMedian_(values){const a=values.filter(v=>typeof v==='number'&&Number.isFinite(v)).sort((a,b)=>a-b),n=a.length;return n?(a[Math.floor((n-1)/2)]+a[Math.floor(n/2)])/2:null;}
function crhRatio_(a,b){return typeof a==='number'&&typeof b==='number'&&a>0&&b>0?a/b:null;}
function crhRate_(vi,days){return days>0?vi/days:null;}
function crhName_(v){return String(v||'').trim().toLowerCase();}
function crhWeight_(c,name){return Object.prototype.hasOwnProperty.call(c.weights||{},crhName_(name))?c.weights[crhName_(name)]:1;}
function crhDays_(keys,c){let n=0;keys.forEach(key=>{n+=crhWeight_(c,JSON.parse(key)[0]);});return n;}
function crhNewMonth_(){return {vi:0,crawlVi:0,unknownVi:0,unknownWorked:new Set(),listingDays:new Set(),crawlDays:new Set(),union:new Set(),worked:new Set(),crawlWorked:new Set(),linkedWorked:new Set(),prepared:0,preparedRequests:0,fallbackVi:0,fallbackRequests:0,crawlerNames:new Set(),invalidVi:0,missingWeight:new Set()};}
function crhRead_(data,c,o){
  if(!/^2026-11-(0[1-9]|[12]\d|30)$/.test(o.asOf))throw Error('이 일회용 보고서는 2026년 9·10·11월용입니다. asOf를 2026-11-DD로 입력하세요.');
  if(!Number.isInteger(o.minRecordedDays)||o.minRecordedDays<1)throw Error('minRecordedDays는 1 이상의 정수여야 합니다.');
  const requests=Object.create(null),rt=laTable_(data.REQUESTS,['Request ID','Node','크롤링 단계','크롤링 작업자','Total VI'],'REQUESTS');
  rt.rows.forEach(r=>{const v=n=>r[rt.h[n]],id=String(v('Request ID')||'').trim();if(!id)return;if(requests[id])throw Error('중복 Request ID: '+id);requests[id]={id:id,node:String(v('Node')||'미지정'),brand:String(v('Brand')||'미지정'),stage:String(v('크롤링 단계')||'').trim().toUpperCase(),creator:String(v('크롤링 작업자')||'').trim(),total:v('Total VI'),workStart:v('Started At'),status:v('Status'),start:v('Requested At'),end:v('Completed At'),blocked:v('Blocked Time (Hours)')};});
  const records=[],crawlRecords=[],seen=new Set(),rawGroups=new Set(),aggregates=new Set(),quality={invalidListing:0,missingCrawlDate:0,unknownPreparedVi:0,missingRatio:new Set(),missingCreator:new Set()};
  const inside=day=>day>='2026-09-01'&&day<=o.asOf;
  const external=name=>(c.external||new Set()).has(crhName_(name));
  const st=laTable_(data.WORK_SESSIONS,['Worker','Request ID','Stage','Status','Ended At','Completed VI Delta'],'WORK_SESSIONS');
  st.rows.forEach(r=>{
    const v=n=>r[st.h[n]],stage=String(v('Stage')||'').trim().toUpperCase(),name=String(v('Worker')||'').trim(),id=String(v('Request ID')||'').trim();
    if(!name||external(name)||!id)return;
    if(stage==='CRAWL'){
      const day=laDay_(v('Ended At'))||laDay_(v('Started At'));if(!day){quality.missingCrawlDate++;return;}
      if(day<=o.asOf)crawlRecords.push({worker:name,name:crhName_(name),id:id,day:day});return;
    }
    if(stage!=='3P'||String(v('Status')||'').toUpperCase()==='ACTIVE')return;
    const day=laDay_(v('Ended At'));if(!day||!inside(day))return;
    const key=String(v('Session ID')||'');if(key){if(seen.has(key))throw Error('중복 Session ID: '+key);seen.add(key);}
    rawGroups.add(JSON.stringify([id,crhName_(name),day]));
    const valid=laNumber_(v('Completed VI Delta'));if(!valid)quality.invalidListing++;
    records.push({worker:name,name:crhName_(name),id:id,day:day,vi:valid?v('Completed VI Delta'):0,valid:valid});
  });
  if(data.LONG_TERM_SESSIONS){const lt=laTable_(data.LONG_TERM_SESSIONS,['Aggregate ID','Status','Request ID','Worker','Stage','Work Date','Completed VI Delta'],'LONG_TERM_SESSIONS');lt.rows.forEach(r=>{
    const v=n=>r[lt.h[n]],name=String(v('Worker')||'').trim(),id=String(v('Request ID')||'').trim(),day=String(v('Work Date')||'');
    if(v('Status')!=='COMMITTED'||v('Stage')!=='3P'||!name||external(name)||!id||!/^\d{4}-\d{2}-\d{2}$/.test(day)||!inside(day))return;
    const key=String(v('Aggregate ID')||'');if(!key||aggregates.has(key))throw Error('장기 집계 ID 누락/중복: '+key);aggregates.add(key);
    if(rawGroups.has(JSON.stringify([id,crhName_(name),day])))throw Error('원본/장기 집계 중첩: '+id+' '+day);
    const valid=laNumber_(v('Completed VI Delta'));if(!valid)quality.invalidListing++;
    records.push({worker:name,name:crhName_(name),id:id,day:day,vi:valid?v('Completed VI Delta'):0,valid:valid});
  });}
  records.concat(crawlRecords).forEach(s=>{if(!Object.prototype.hasOwnProperty.call(c.weights||{},s.name))quality.missingRatio.add(s.name);});
  return {requests:requests,records:records,crawlRecords:crawlRecords,quality:quality,inside:inside};
}
function crhBucket_(){return {vi:0,days:new Set(),requests:new Set()};}
function crhAdd_(bucket,s){bucket.vi+=s.vi;bucket.days.add(JSON.stringify([s.name,s.day]));bucket.requests.add(s.id);}
function crhPure_(records,requests,classifier){
  const days=new Map(),groups=new Map();
  records.forEach(s=>{const dayKey=JSON.stringify([s.name,s.day]);if(!days.has(dayKey))days.set(dayKey,[]);days.get(dayKey).push(s);});
  let mixed=0;
  days.forEach(list=>{
    const classes=new Set(list.map(s=>classifier(s,requests[s.id])));if(classes.size!==1||classes.has(null)||list.some(s=>!s.valid)){mixed+=list.length;return;}
    const label=Array.from(classes)[0];list.forEach(s=>{const key=JSON.stringify([s.day.slice(0,7),s.name,label]);if(!groups.has(key))groups.set(key,crhBucket_());crhAdd_(groups.get(key),s);});
  });return {groups:groups,mixed:mixed};
}
function crhWorkerEffects_(records,requests,c,minDays){
  const pure=crhPure_(records,requests,(_,r)=>r&&r.stage==='CRAWL_COMPLETED'?'crawl':r&&r.stage==='N/A'?'manual':null),workers=new Map();
  pure.groups.forEach((bucket,key)=>{const [month,name,mode]=JSON.parse(key),k=JSON.stringify([month,name]);if(!workers.has(k))workers.set(k,{});workers.get(k)[mode]=bucket;});
  const monthly=new Map(),byWorker=new Map();
  workers.forEach((g,key)=>{
    if(!g.crawl||!g.manual||g.crawl.days.size<minDays||g.manual.days.size<minDays)return;
    const a=crhRate_(g.crawl.vi,crhDays_(g.crawl.days,c)),b=crhRate_(g.manual.vi,crhDays_(g.manual.days,c)),ratio=crhRatio_(a,b);if(ratio===null)return;
    const [month,name]=JSON.parse(key);byWorker.set(key,ratio);if(!monthly.has(month))monthly.set(month,[]);monthly.get(month).push(ratio);
  });
  const pairs=[];byWorker.forEach((nov,key)=>{const [month,name]=JSON.parse(key);if(month!=='2026-11')return;const oct=byWorker.get(JSON.stringify(['2026-10',name]));if(oct)pairs.push({name:name,oct:oct,nov:nov,change:nov/oct});});
  return {monthly:monthly,paired:{count:pairs.length,oct:crhMedian_(pairs.map(p=>p.oct)),nov:crhMedian_(pairs.map(p=>p.nov)),change:crhMedian_(pairs.map(p=>p.change))},mixed:pure.mixed};
}
function crhBuild_(data,c,o){
  o=Object.assign({minRecordedDays:1,exampleMinDays:2,septemberStart:'AUTO'},o);const parsed=crhRead_(data,c,o);const september=crhSeptemberWindow_(parsed.records,o.septemberStart);const originalInside=parsed.inside;parsed.inside=day=>originalInside(day)&&(day.slice(0,7)!=='2026-09'||(september.start!==null&&day>=september.start));parsed.records=parsed.records.filter(s=>parsed.inside(s.day));const months=['2026-09','2026-10','2026-11'],stats={},comparable={},latest=new Map();
  months.forEach(m=>{stats[m]=crhNewMonth_();comparable[m]=crhNewMonth_();});
  const dayLimit=Number(o.asOf.slice(8,10));
  parsed.crawlRecords.forEach(s=>{
    const key=JSON.stringify([s.id,s.name]);if(!latest.has(key)||latest.get(key)<s.day)latest.set(key,s.day);
    if(!parsed.inside(s.day))return;const m=s.day.slice(0,7),d=JSON.stringify([s.name,s.day]);stats[m].crawlDays.add(d);stats[m].union.add(d);stats[m].crawlerNames.add(s.name);
    if(Number(s.day.slice(8,10))<=dayLimit){comparable[m].crawlDays.add(d);comparable[m].union.add(d);}
  });
  Object.values(parsed.requests).forEach(r=>{
    if(r.stage!=='CRAWL_COMPLETED')return;
    if((c.external||new Set()).has(crhName_(r.creator)))return;
    if(!r.creator)parsed.quality.missingCreator.add(r.id);
    const sessionDay=latest.get(JSON.stringify([r.id,crhName_(r.creator)])),day=sessionDay||laDay_(r.workStart);if(!day){parsed.quality.unknownPreparedVi+=laNumber_(r.total)&&r.total>0?r.total:0;return;}
    if(!parsed.inside(day))return;const m=day.slice(0,7);
    if(laNumber_(r.total)&&r.total>=0){stats[m].prepared+=r.total;stats[m].preparedRequests++;if(r.creator)stats[m].crawlerNames.add(crhName_(r.creator));if(!sessionDay){stats[m].fallbackVi+=r.total;stats[m].fallbackRequests++;}}
  });
  function add(g,s){const d=JSON.stringify([s.name,s.day]),r=parsed.requests[s.id];g.vi+=s.vi;g.listingDays.add(d);g.union.add(d);g.worked.add(s.id);if(!s.valid)g.invalidVi++;
    if(r&&r.stage==='CRAWL_COMPLETED'){g.crawlVi+=s.vi;g.crawlWorked.add(s.id);if(latest.has(JSON.stringify([r.id,crhName_(r.creator)])))g.linkedWorked.add(s.id);}else if(!r||r.stage!=='N/A'){g.unknownVi+=s.vi;g.unknownWorked.add(s.id);}}
  parsed.records.forEach(s=>{const m=s.day.slice(0,7);add(stats[m],s);if(Number(s.day.slice(8,10))<=dayLimit)add(comparable[m],s);});
  const effects=crhWorkerEffects_(parsed.records,parsed.requests,c,o.minRecordedDays),cohort=crhFixedCohort_(parsed.records,c),examples=crhExamples_(parsed.records,parsed.requests,c,o.exampleMinDays);
  function summarize(g){const d=crhDays_(g.listingDays,c),cd=crhDays_(g.crawlDays,c),union=crhDays_(g.union,c);return {complete:g.fallbackRequests===0&&g.unknownWorked.size===0&&g.invalidVi===0&&(g.crawlWorked.size===0||g.crawlWorked.size===g.linkedWorked.size),vi:g.vi,crawlers:g.crawlerNames.size,listers:new Set(Array.from(g.listingDays,k=>JSON.parse(k)[0])).size,listingD:d,crawlD:cd,totalD:union,listingRate:crhRate_(g.vi,d),totalRate:crhRate_(g.vi,union),share:g.vi>0&&g.crawlVi>=0&&g.crawlVi<=g.vi?g.crawlVi/g.vi:null,prepared:g.prepared,preparedRequests:g.preparedRequests,fallbackVi:g.fallbackVi,fallbackRequests:g.fallbackRequests,crawlRate:g.fallbackRequests===0?crhRate_(g.prepared,cd):null,coverage:g.crawlWorked.size?g.linkedWorked.size/g.crawlWorked.size:null,unknownVi:g.unknownVi,invalidVi:g.invalidVi,requests:g.worked.size};}
  const hold={};months.forEach(m=>hold[m]=[]);Object.values(parsed.requests).forEach(r=>{if(r.status!=='DONE'||!laNumber_(r.start)||!laNumber_(r.end)||r.end<r.start||!laNumber_(r.blocked)||r.blocked<0||r.blocked>(r.end-r.start)*24)return;const day=laDay_(r.end);if(!parsed.inside(day)||!hold[day.slice(0,7)])return;hold[day.slice(0,7)].push(r.end-r.start-r.blocked/24);});
  const summary={},matchedPeriod={};months.forEach(m=>{summary[m]=summarize(stats[m]);matchedPeriod[m]=summarize(comparable[m]);});
  const monthlyRows=[
    ['3P 리스팅 VI',...months.map(m=>summary[m].vi)],['리스팅 VI/d',...months.map(m=>summary[m].listingRate)],
    ['크롤링 물량 비중',...months.map(m=>summary[m].share)],['크롤링 VI (귀속월 추정)',...months.map(m=>summary[m].prepared)],
    ['크롤링 참여 인원',...months.map(m=>summary[m].crawlers)],['크롤링 d',...months.map(m=>summary[m].crawlD)],['크롤링 VI/d',...months.map(m=>summary[m].crawlRate)],
    ['크롤링+리스팅 총 d (중복 제거)',...months.map(m=>summary[m].totalD)],['공정 합산 관측 VI/d',...months.map(m=>summary[m].totalRate)],
    ['같은 집단 내 크롤링 효과 배수',...months.map(m=>crhMedian_(effects.monthly.get(m)||[]))],
    ['각 월 1~'+dayLimit+'일 리스팅 VI/d',...months.map(m=>matchedPeriod[m].listingRate)],
    ['CRAWL 기록 연결률 (관련 요청)',...months.map(m=>summary[m].coverage)],
    ['크롤링 분류 확인 VI 비중',...months.map(m=>summary[m].vi>0&&summary[m].unknownVi>=0&&summary[m].unknownVi<=summary[m].vi?(summary[m].vi-summary[m].unknownVi)/summary[m].vi:null)],
    ['전체 완료요청 홀드제외일 (달력일)',...months.map(m=>crhMedian_(hold[m]))]
  ];
  const growth=crhRatio_(summary['2026-11'].listingRate,summary['2026-10'].listingRate),totalGrowth=summary['2026-11'].complete&&summary['2026-10'].complete?crhRatio_(summary['2026-11'].totalRate,summary['2026-10'].totalRate):null;
  const quality={invalidListing:parsed.quality.invalidListing,missingCrawlDate:parsed.quality.missingCrawlDate,unknownPreparedVi:parsed.quality.unknownPreparedVi,missingRatio:Array.from(parsed.quality.missingRatio),missingCreator:Array.from(parsed.quality.missingCreator),mixedComparisonRecords:effects.mixed};
  return {asOf:o.asOf,september:september,exampleMinDays:o.exampleMinDays,examples:examples,effectCounts:Object.fromEntries(months.map(m=>[m,(effects.monthly.get(m)||[]).length])),summary:summary,matchedPeriod:matchedPeriod,monthlyRows:monthlyRows,paired:effects.paired,cohort:cohort,holdCounts:months.map(m=>hold[m].length),growth:growth,totalGrowth:totalGrowth,quality:quality};
}
function crhText_(value,style){if(value===null||value===undefined||!Number.isFinite(value))return '표본 없음';if(style==='percent')return (value*100).toFixed(1)+'%';if(style==='ratio')return value.toFixed(2)+'×';return value.toLocaleString('en-US',{maximumFractionDigits:style==='integer'?0:1});}
function crhFixedCohort_(records,c){
  const groups=new Map();records.forEach(s=>{const month=s.day.slice(0,7);if(!['2026-10','2026-11'].includes(month))return;const key=JSON.stringify([month,s.name]);if(!groups.has(key))groups.set(key,crhBucket_());crhAdd_(groups.get(key),s);});
  const oct=[],nov=[],ratios=[];groups.forEach((g,key)=>{const [month,name]=JSON.parse(key);if(month!=='2026-11')return;const previous=groups.get(JSON.stringify(['2026-10',name]));if(!previous)return;const a=crhRate_(previous.vi,crhDays_(previous.days,c)),b=crhRate_(g.vi,crhDays_(g.days,c)),ratio=crhRatio_(b,a);if(ratio===null)return;oct.push(a);nov.push(b);ratios.push(ratio);});
  return {count:ratios.length,oct:crhMedian_(oct),nov:crhMedian_(nov),change:crhMedian_(ratios)};
}
function crhSeptemberWindow_(records,setting){
  const daily=new Map();records.filter(s=>s.valid).forEach(s=>{if(!daily.has(s.day))daily.set(s.day,new Set());daily.get(s.day).add(s.name);});
  const september=Array.from(daily.keys()).filter(d=>d.slice(0,7)==='2026-09').sort();
  if(setting&&setting!=='AUTO'){
    if(!/^2026-09-(0[1-9]|[12]\d|30)$/.test(setting))throw Error('septemberStart는 AUTO 또는 2026-09-DD이어야 합니다.');
    return {start:setting,method:'사용자 지정',activeDays:september.filter(d=>d>=setting).length};
  }
  const october=Array.from(daily.keys()).filter(d=>d.slice(0,7)==='2026-10').map(d=>daily.get(d).size),threshold=Math.max(1,(crhMedian_(october)||1)*0.5);
  for(const start of september){
    const end=new Date(Date.parse(start+'T00:00:00Z')+6*86400000).toISOString().slice(0,10),window=september.filter(d=>d>=start&&d<=end);
    if(window.length>=3&&window.reduce((n,d)=>n+daily.get(d).size,0)/window.length>=threshold)return {start:start,method:'7일 중 3일 이상 기록·10월 중앙 참여인원 50% 이상',activeDays:september.filter(d=>d>=start).length};
  }
  return {start:september[0]||null,method:'연속 기록 기준 미충족·첫 유효 기록부터 참고',activeDays:september.length};
}
function crhExamples_(records,requests,c,minDays){
  if(!Number.isInteger(minDays)||minDays<1)throw Error('exampleMinDays는 1 이상의 정수여야 합니다.');
  const pure=crhPure_(records,requests,(_,r)=>r&&r.stage==='CRAWL_COMPLETED'?'crawl':r&&r.stage==='N/A'?'manual':null),groups=new Map(),display=new Map(),result={};
  records.forEach(s=>{if(!display.has(s.name))display.set(s.name,s.worker);});
  pure.groups.forEach((bucket,key)=>{const [month,name,mode]=JSON.parse(key),k=JSON.stringify([month,name]);if(!groups.has(k))groups.set(k,{});groups.get(k)[mode]=bucket;});
  ['2026-09','2026-10','2026-11'].forEach(month=>result[month]={eligible:0,improved:0,rows:[]});
  groups.forEach((g,key)=>{
    if(!g.crawl||!g.manual||g.crawl.days.size<minDays||g.manual.days.size<minDays)return;
    const [month,name]=JSON.parse(key),manualD=crhDays_(g.manual.days,c),crawlD=crhDays_(g.crawl.days,c),manualRate=crhRate_(g.manual.vi,manualD),crawlRate=crhRate_(g.crawl.vi,crawlD),ratio=crhRatio_(crawlRate,manualRate);if(ratio===null)return;
    const out=result[month];out.eligible++;if(ratio<=1)return;out.improved++;out.rows.push({worker:display.get(name),manualRate:manualRate,crawlRate:crawlRate,ratio:ratio,manualD:manualD,crawlD:crawlD,manualVi:g.manual.vi,crawlVi:g.crawl.vi});
  });
  Object.values(result).forEach(g=>{g.rows.sort((a,b)=>b.ratio-a.ratio||b.crawlVi-a.crawlVi||a.worker.localeCompare(b.worker));g.rows=g.rows.slice(0,5);});return result;
}
function crhRender_(ss,r){
  const name='11월 보고',oldName='Crawler Hub 리더 보고';let sheet=ss.getSheetByName(name);
  if(!sheet){const old=ss.getSheetByName(oldName);if(old&&old.getRange('A1').getValue()==='CRAWLER HUB | IMPACT REVIEW'){old.setName(name);sheet=old;}}
  const marker=sheet?sheet.getRange('A1').getValue():'';
  if(sheet&&!['CRAWLER HUB | IMPACT REVIEW','11월 보고 | CRAWLER HUB'].includes(marker))throw Error('동일 이름의 기존 탭이 있습니다. 이름을 바꿔 보존한 뒤 다시 실행하세요.');
  if(!sheet)sheet=ss.insertSheet(name);
  if(sheet.getMaxRows()<48)sheet.insertRowsAfter(sheet.getMaxRows(),48-sheet.getMaxRows());if(sheet.getMaxColumns()<12)sheet.insertColumnsAfter(sheet.getMaxColumns(),12-sheet.getMaxColumns());
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).getMergedRanges().forEach(range=>range.breakApart());sheet.clear();
  sheet.getRange(1,1,48,12).setBackground('#f4f7fb').setFontFamily('Arial').setFontSize(9).setFontColor('#23324b').setVerticalAlignment('middle');sheet.setColumnWidths(1,12,90);sheet.setRowHeights(1,48,18);sheet.setHiddenGridlines(true);sheet.setFrozenRows(0);sheet.setTabColor('#276b73');
  function text(row,col,width,value,bg,color,size){const cell=sheet.getRange(row,col,1,width);if(width>1)cell.merge();cell.setNumberFormat('@').setValue(laReportValue_(String(value),'')).setWrap(true);if(bg)cell.setBackground(bg);if(color)cell.setFontColor(color);if(size)cell.setFontSize(size);return cell;}
  const f=crhText_,months=['2026-09','2026-10','2026-11'],sep=r.september.start?'9/'+Number(r.september.start.slice(8,10))+'~30':'9월 기록 없음',nov=r.summary['2026-11'],oct=r.summary['2026-10'];
  const effect=r.monthlyRows.find(row=>row[0]==='같은 집단 내 크롤링 효과 배수'),novEffect=effect?effect[3]:null;
  const change=value=>value===null?'비교 표본 없음':(value>=1?'+':'')+((value-1)*100).toFixed(1)+'%';
  text(1,1,12,'11월 보고 | CRAWLER HUB','#15364a','#ffffff',22).setFontWeight('bold');sheet.setRowHeight(1,38);
  text(2,1,12,'9월 '+sep+' · 10월 전체 · 11월 1~'+r.asOf.slice(8,10)+'일  |  내부 3P 리스팅 · 기준일 '+r.asOf,null,'#607087',9);
  const cards=[['11월 리스팅 물량',f(nov.vi,'integer'),'완료 VI'],['10월 대비 리스팅 VI/d',f(r.growth,'ratio'),change(r.growth)],['11월 크롤링 효과',f(novEffect,'ratio'),'C/N 배수 중앙값 · '+r.effectCounts['2026-11']+'명'],['같은 참여 집단의 VI/d 변화',f(r.cohort.change,'ratio'),change(r.cohort.change)+' · '+r.cohort.count+'명']];
  cards.forEach((card,i)=>{const col=1+i*3;text(4,col,3,card[0],'#dce9ed','#33586a',9);text(5,col,3,card[1],'#ffffff','#15364a',25).setFontWeight('bold');text(6,col,3,card[2],'#ffffff','#38726c',9);});sheet.setRowHeight(5,39);
  text(8,1,12,'01  전환 규모와 생산성 — 월별 숫자 비교','#15364a','#ffffff',11).setFontWeight('bold');
  text(9,1,6,'지표');[sep,'10월 전체','11월 1~'+r.asOf.slice(8,10)+'일'].forEach((v,i)=>text(9,7+i*2,2,v,'#dce9ed').setFontWeight('bold'));
  const matrix=[
    ['리스팅 완료 VI',m=>r.summary[m].vi,'integer'],['리스팅 VI/d',m=>r.summary[m].listingRate,''],
    ['11월 VI/d ÷ 해당 월 VI/d',m=>crhRatio_(nov.listingRate,r.summary[m].listingRate),'ratio'],
    ['크롤링 VI (세션 없으면 Started At 월)',m=>r.summary[m].prepared,'integer'],
    ['크롤링 요청 수',m=>r.summary[m].preparedRequests,'integer'],['크롤링 물량 비중 (리스팅 VI 기준)',m=>r.summary[m].share,'percent'],
    ['동일 작업자 내 크롤링 효과 배수',m=>{const row=r.monthlyRows.find(x=>x[0]==='같은 집단 내 크롤링 효과 배수');return row?row[months.indexOf(m)+1]:null;},'ratio'],
    ['크롤링 참여 인원 (요청·세션 기준)',m=>r.summary[m].crawlers,'integer'],
    ['각 월 1~'+r.asOf.slice(8,10)+'일 리스팅 VI/d',m=>r.matchedPeriod[m].listingRate,''],
    ['Started At로 귀속한 크롤링 VI',m=>r.summary[m].fallbackVi,'integer']
  ];
  matrix.forEach((row,i)=>{const rr=10+i;sheet.getRange(rr,1,1,12).setBackground(i%2?'#eaf0f5':'#ffffff');text(rr,1,6,row[0]);months.forEach((m,j)=>text(rr,7+j*2,2,f(row[1](m),row[2]),j===2?'#e2f0e9':null).setHorizontalAlignment('right').setFontWeight([1,2,6].includes(i)?'bold':'normal'));});
  text(21,1,12,'02  리스팅 효율 개선 사례 — 월별 C/N 배수 상위 5명','#15364a','#ffffff',11).setFontWeight('bold');
  months.forEach((m,i)=>{
    const col=1+i*4,g=r.examples[m];text(22,col,4,(i===0?sep:i===1?'10월':'11월')+' · 개선 '+g.improved+'/'+g.eligible+'명','#dce9ed').setFontWeight('bold');
    ['작업자 · 표본 d','N VI/d','C VI/d','배수'].forEach((label,j)=>text(23,col+j,1,label,'#e7edf3',null,8));
    for(let k=0;k<5;k++){const row=24+k,item=g.rows[k];sheet.setRowHeight(row,29);if(!item){text(row,col,4,k===0?'유효 개선 사례 없음':'—','#ffffff','#9aa7b6',8);continue;}
      text(row,col,1,item.worker+'\nN'+f(item.manualD)+' / C'+f(item.crawlD),'#ffffff',null,8);text(row,col+1,1,f(item.manualRate),'#ffffff',null,8);text(row,col+2,1,f(item.crawlRate),'#ffffff',null,8);text(row,col+3,1,f(item.ratio,'ratio'),'#d7eddf','#1d6448',12).setFontWeight('bold');
    }
  });
  text(29,1,12,'C=크롤링, N=비크롤링. 각 방식 '+r.exampleMinDays+'기록일 이상·양쪽 VI/d>0. 상위 5명은 개선 사례이며 전체 효과는 위 중앙값·유효 인원으로 확인.',null,'#607087',8);sheet.setRowHeight(29,26);
  text(31,1,12,'03  같은 물량에 필요한 d와 크롤링+리스팅 투입','#15364a','#ffffff',11).setFontWeight('bold');
  const costs=[['리스팅 1,000 VI당 d',m=>r.summary[m].listingRate>0?f(1000/r.summary[m].listingRate):'표본 없음'],['크롤링 관측 d',m=>r.summary[m].fallbackRequests>0?f(r.summary[m].crawlD)+' + 미복원':f(r.summary[m].crawlD)],['총 관측 d (같은 사람·날짜 중복 제거)',m=>f(r.summary[m].totalD)],['공정 합산 VI/d (크롤링 기록 완전 시)',m=>r.summary[m].complete?f(r.summary[m].totalRate):'미복원'],['전체 완료요청 홀드제외 중앙값 (달력일)',m=>{const row=r.monthlyRows.find(x=>x[0].includes('홀드제외'));return row?f(row[months.indexOf(m)+1]):'표본 없음';}]];
  text(32,1,6,'지표');[sep,'10월','11월'].forEach((v,i)=>text(32,7+i*2,2,v,'#dce9ed'));
  costs.forEach((row,i)=>{text(33+i,1,6,row[0]);months.forEach((m,j)=>text(33+i,7+j*2,2,row[1](m),i%2?'#eaf0f5':'#ffffff').setHorizontalAlignment('right'));});
  const before=oct.listingRate>0?1000/oct.listingRate:null,after=nov.listingRate>0?1000/nov.listingRate:null,saved=before!==null&&after!==null?before-after:null;
  text(39,1,12,'10월→11월: 리스팅 VI/d '+change(r.growth)+'  |  1,000 VI당 기록 d 차이 '+f(saved)+'d  |  같은 집단 효과 배수 변화 '+f(r.paired.change,'ratio'),'#dce9ed','#15364a',11).setFontWeight('bold');sheet.setRowHeight(39,30);
  const fallback=months.reduce((n,m)=>n+r.summary[m].fallbackRequests,0),notes=[
    'd=고유 기록일×CONFIG 비율. 9월 시작 '+(r.september.start||'없음')+' ('+r.september.method+'). 기록 밀도만으로 선택하며 효율이 높은 기간을 고르지 않음.',
    '기간 길이가 달라 월 총량은 단순 증감률로 비교하지 않음. 9월 1~'+r.asOf.slice(8,10)+'일 값도 선택된 유효 시작일 이후만 포함. 같은 집단 비교 '+r.cohort.count+'명·C/N 배수 변화 '+r.paired.count+'명.',
    '귀속: CRAWL 세션 우선, 없으면 요청 Started At('+fallback+'건). Started At은 귀속월만 보완하며 크롤링 작업일을 생성하지 않음. 과거 총 투입 d는 미복원일 수 있음.',
    '상위 사례는 양쪽 방식 기록이 있는 작업자만 선정. 같은 날 혼합 작업은 C/N 비교에서 제외. 현재 요청 속성 소급·난이도·배정 차이가 남는 관측 비교. 허브 사용 자체의 인과 증명은 아님.',
    '점검: 월 미귀속 준비 VI '+f(r.quality.unknownPreparedVi,'integer')+', VI 누락 '+r.quality.invalidListing+'행, 비율 미설정 '+r.quality.missingRatio.length+'명(기본 1). 외부 지원 제외. 이 탭은 별도 실행 시에만 갱신.'
  ];notes.forEach((note,i)=>{text(41+i,1,12,note,null,'#607087',8);sheet.setRowHeight(41+i,25);});ss.setActiveSheet(sheet);
}
