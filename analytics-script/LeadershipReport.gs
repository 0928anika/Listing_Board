/** One-off leadership snapshot. Never called by refreshListingAnalytics or a trigger. */
const CRAWLER_HUB_REPORT_OPTIONS={
  asOf:'2026-11-20',
  minRecordedDays:1
};
function buildCrawlerHubLeadershipReport(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),c=laConfig_(ss),o=Object.assign({},CRAWLER_HUB_REPORT_OPTIONS);
  const lock=LockService.getDocumentLock();if(!lock.tryLock(1000))throw Error('다른 갱신이 실행 중입니다. 끝난 뒤 보고서를 만드세요.');
  try{
    const data=laRead_(c.id),report=crhBuild_(data,c,o);
    crhRender_(ss,report);
    console.log(JSON.stringify({report:'Crawler Hub 리더 보고',asOf:o.asOf,matchedWorkers:report.paired.count,quality:report.quality}));
  }finally{lock.releaseLock();}
}
function crhMedian_(values){const a=values.filter(v=>typeof v==='number'&&Number.isFinite(v)).sort((a,b)=>a-b),n=a.length;return n?(a[Math.floor((n-1)/2)]+a[Math.floor(n/2)])/2:null;}
function crhRatio_(a,b){return typeof a==='number'&&typeof b==='number'&&a>0&&b>0?a/b:null;}
function crhRate_(vi,days){return days>0?vi/days:null;}
function crhName_(v){return String(v||'').trim().toLowerCase();}
function crhWeight_(c,name){return Object.prototype.hasOwnProperty.call(c.weights||{},crhName_(name))?c.weights[crhName_(name)]:1;}
function crhDays_(keys,c){let n=0;keys.forEach(key=>{n+=crhWeight_(c,JSON.parse(key)[0]);});return n;}
function crhNewMonth_(){return {vi:0,crawlVi:0,unknownVi:0,unknownWorked:new Set(),listingDays:new Set(),crawlDays:new Set(),union:new Set(),worked:new Set(),crawlWorked:new Set(),linkedWorked:new Set(),prepared:0,preparedRequests:0,invalidVi:0,missingWeight:new Set()};}
function crhRead_(data,c,o){
  if(!/^2026-11-(0[1-9]|[12]\d|30)$/.test(o.asOf))throw Error('이 일회용 보고서는 2026년 9·10·11월용입니다. asOf를 2026-11-DD로 입력하세요.');
  if(!Number.isInteger(o.minRecordedDays)||o.minRecordedDays<1)throw Error('minRecordedDays는 1 이상의 정수여야 합니다.');
  const requests=Object.create(null),rt=laTable_(data.REQUESTS,['Request ID','Node','크롤링 단계','크롤링 작업자','Total VI'],'REQUESTS');
  rt.rows.forEach(r=>{const v=n=>r[rt.h[n]],id=String(v('Request ID')||'').trim();if(!id)return;if(requests[id])throw Error('중복 Request ID: '+id);requests[id]={id:id,node:String(v('Node')||'미지정'),brand:String(v('Brand')||'미지정'),stage:String(v('크롤링 단계')||'').trim().toUpperCase(),creator:String(v('크롤링 작업자')||'').trim(),total:v('Total VI'),status:v('Status'),start:v('Requested At'),end:v('Completed At'),blocked:v('Blocked Time (Hours)')};});
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
  o=Object.assign({minRecordedDays:1},o);const parsed=crhRead_(data,c,o),months=['2026-09','2026-10','2026-11'],stats={},comparable={},latest=new Map();
  months.forEach(m=>{stats[m]=crhNewMonth_();comparable[m]=crhNewMonth_();});
  const dayLimit=Number(o.asOf.slice(8,10));
  parsed.crawlRecords.forEach(s=>{
    const key=JSON.stringify([s.id,s.name]);if(!latest.has(key)||latest.get(key)<s.day)latest.set(key,s.day);
    if(!parsed.inside(s.day))return;const m=s.day.slice(0,7),d=JSON.stringify([s.name,s.day]);stats[m].crawlDays.add(d);stats[m].union.add(d);
    if(Number(s.day.slice(8,10))<=dayLimit){comparable[m].crawlDays.add(d);comparable[m].union.add(d);}
  });
  Object.values(parsed.requests).forEach(r=>{
    if(r.stage!=='CRAWL_COMPLETED')return;
    if((c.external||new Set()).has(crhName_(r.creator)))return;
    if(!r.creator){parsed.quality.missingCreator.add(r.id);return;}
    const day=latest.get(JSON.stringify([r.id,crhName_(r.creator)]));if(!day){parsed.quality.unknownPreparedVi+=laNumber_(r.total)?r.total:0;return;}
    if(!parsed.inside(day))return;const m=day.slice(0,7);
    if(laNumber_(r.total)&&r.total>=0){stats[m].prepared+=r.total;stats[m].preparedRequests++;}
  });
  function add(g,s){const d=JSON.stringify([s.name,s.day]),r=parsed.requests[s.id];g.vi+=s.vi;g.listingDays.add(d);g.union.add(d);g.worked.add(s.id);if(!s.valid)g.invalidVi++;
    if(r&&r.stage==='CRAWL_COMPLETED'){g.crawlVi+=s.vi;g.crawlWorked.add(s.id);if(latest.has(JSON.stringify([r.id,crhName_(r.creator)])))g.linkedWorked.add(s.id);}else if(!r||r.stage!=='N/A'){g.unknownVi+=s.vi;g.unknownWorked.add(s.id);}}
  parsed.records.forEach(s=>{const m=s.day.slice(0,7);add(stats[m],s);if(Number(s.day.slice(8,10))<=dayLimit)add(comparable[m],s);});
  const effects=crhWorkerEffects_(parsed.records,parsed.requests,c,o.minRecordedDays),cohort=crhFixedCohort_(parsed.records,c);
  function summarize(g){const d=crhDays_(g.listingDays,c),cd=crhDays_(g.crawlDays,c),union=crhDays_(g.union,c);return {complete:g.unknownWorked.size===0&&g.invalidVi===0&&(g.crawlWorked.size===0||g.crawlWorked.size===g.linkedWorked.size),vi:g.vi,crawlers:new Set(Array.from(g.crawlDays,k=>JSON.parse(k)[0])).size,listers:new Set(Array.from(g.listingDays,k=>JSON.parse(k)[0])).size,listingD:d,crawlD:cd,totalD:union,listingRate:crhRate_(g.vi,d),totalRate:crhRate_(g.vi,union),share:g.vi>0&&g.crawlVi>=0&&g.crawlVi<=g.vi?g.crawlVi/g.vi:null,prepared:g.prepared,preparedRequests:g.preparedRequests,crawlRate:crhRate_(g.prepared,cd),coverage:g.crawlWorked.size?g.linkedWorked.size/g.crawlWorked.size:null,unknownVi:g.unknownVi,invalidVi:g.invalidVi,requests:g.worked.size};}
  const hold={};months.forEach(m=>hold[m]=[]);Object.values(parsed.requests).forEach(r=>{if(r.status!=='DONE'||!laNumber_(r.start)||!laNumber_(r.end)||r.end<r.start||!laNumber_(r.blocked)||r.blocked<0||r.blocked>(r.end-r.start)*24)return;const day=laDay_(r.end);if(day>o.asOf||!hold[day.slice(0,7)])return;hold[day.slice(0,7)].push(r.end-r.start-r.blocked/24);});
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
  return {asOf:o.asOf,summary:summary,matchedPeriod:matchedPeriod,monthlyRows:monthlyRows,paired:effects.paired,cohort:cohort,holdCounts:months.map(m=>hold[m].length),growth:growth,totalGrowth:totalGrowth,quality:quality};
}
function crhText_(value,style){if(value===null||value===undefined||!Number.isFinite(value))return '표본 없음';if(style==='percent')return (value*100).toFixed(1)+'%';if(style==='ratio')return value.toFixed(2)+'×';return value.toLocaleString('en-US',{maximumFractionDigits:style==='integer'?0:1});}
function crhFixedCohort_(records,c){
  const groups=new Map();records.forEach(s=>{const month=s.day.slice(0,7);if(!['2026-10','2026-11'].includes(month))return;const key=JSON.stringify([month,s.name]);if(!groups.has(key))groups.set(key,crhBucket_());crhAdd_(groups.get(key),s);});
  const oct=[],nov=[],ratios=[];groups.forEach((g,key)=>{const [month,name]=JSON.parse(key);if(month!=='2026-11')return;const previous=groups.get(JSON.stringify(['2026-10',name]));if(!previous)return;const a=crhRate_(previous.vi,crhDays_(previous.days,c)),b=crhRate_(g.vi,crhDays_(g.days,c)),ratio=crhRatio_(b,a);if(ratio===null)return;oct.push(a);nov.push(b);ratios.push(ratio);});
  return {count:ratios.length,oct:crhMedian_(oct),nov:crhMedian_(nov),change:crhMedian_(ratios)};
}
function crhRender_(ss,r){
  const name='Crawler Hub 리더 보고';let sheet=ss.getSheetByName(name);
  if(sheet&&sheet.getRange('A1').getValue()!=='CRAWLER HUB | IMPACT REVIEW')throw Error('동일 이름의 기존 탭이 있습니다. 이름을 바꿔 보존한 뒤 다시 실행하세요.');
  if(!sheet)sheet=ss.insertSheet(name);
  if(sheet.getMaxRows()<50)sheet.insertRowsAfter(sheet.getMaxRows(),50-sheet.getMaxRows());if(sheet.getMaxColumns()<12)sheet.insertColumnsAfter(sheet.getMaxColumns(),12-sheet.getMaxColumns());
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).getMergedRanges().forEach(range=>range.breakApart());sheet.clear();
  sheet.getRange(1,1,50,12).setBackground('#f5f7fb').setFontFamily('Arial').setFontSize(9).setFontColor('#23324b').setVerticalAlignment('middle');sheet.setColumnWidths(1,12,86);sheet.setRowHeights(1,50,19);sheet.setHiddenGridlines(true);sheet.setFrozenRows(0);sheet.setTabColor('#276b73');
  function text(row,col,width,value,bg,color,size){const cell=sheet.getRange(row,col,1,width).merge();cell.setNumberFormat('@').setValue(laReportValue_(String(value),'')).setWrap(true);if(bg)cell.setBackground(bg);if(color)cell.setFontColor(color);if(size)cell.setFontSize(size);return cell;}
  text(1,1,12,'CRAWLER HUB | IMPACT REVIEW','#17334c','#ffffff',20).setFontWeight('bold');sheet.setRowHeight(1,38);
  text(2,1,12,'2026.09 · 10 · 11 월별 팀 성과  |  11월은 '+r.asOf+'까지  |  내부 3P 리스팅 · 일회성 스냅샷',null,'#607087',10);
  text(3,1,12,'크롤링 확대 → 인력 증가와 효율 분리 → 같은 참여 집단의 변화 → 총 투입 d 대비 성과',null,'#17334c',11).setFontWeight('bold');
  const nov=r.summary['2026-11'],oct=r.summary['2026-10'];
  const cards=[['11월 크롤링 물량 비중',crhText_(nov.share,'percent')],['10월 대비 리스팅 VI/d',crhText_(r.growth,'ratio')],['같은 참여 집단의 VI/d 변화',crhText_(r.cohort.change,'ratio')],['총 투입 d 대비 VI/d 변화',!nov.complete||!oct.complete?'기록 확인 필요':crhText_(r.totalGrowth,'ratio')]];
  cards.forEach((card,i)=>{text(5,1+i*3,3,card[0],'#dce9ed','#33586a',9);text(6,1+i*3,3,card[1],'#ffffff','#17334c',23).setFontWeight('bold');});sheet.setRowHeight(6,38);
  text(8,1,12,'01  월별 전환 · 물량 · 보정 근무일당 성과','#17334c','#ffffff',11).setFontWeight('bold');
  text(9,1,6,'지표');['9월 전체','10월 전체','11월 1~'+r.asOf.slice(8,10)+'일'].forEach((m,i)=>text(9,7+i*2,2,m,'#dce9ed').setFontWeight('bold'));
  r.monthlyRows.forEach((row,i)=>{const rr=10+i;sheet.getRange(rr,1,1,12).setBackground(i%2?'#edf2f7':'#ffffff');text(rr,1,6,row[0]);row.slice(1).forEach((v,j)=>{const style=[2,11,12].includes(i)?'percent':i===9?'ratio':[0,3,4].includes(i)?'integer':'';text(rr,7+j*2,2,crhText_(v,style)).setHorizontalAlignment('right');});});
  text(24,1,12,'02  크롤링 인력이 늘어난 효과와 생산성 변화를 구분','#17334c','#ffffff',11).setFontWeight('bold');
  text(25,1,8,'크롤링 참여 인원 (기록이 확인되는 인원)');text(25,9,4,crhText_(oct.crawlers,'integer')+'명 → '+crhText_(nov.crawlers,'integer')+'명');
  text(26,1,8,'인원 증가를 d로 보정한 크롤링 VI/d: 10월 → 11월');text(26,9,4,crhText_(oct.crawlRate)+' → '+crhText_(nov.crawlRate));
  text(27,1,8,'동일 참여 집단의 리스팅 VI/d: 10월 → 11월 ('+r.cohort.count+'명 집계)');text(27,9,4,crhText_(r.cohort.oct)+' → '+crhText_(r.cohort.nov));
  text(28,1,8,'동일 집단 내 크롤링/비크롤링 효과 배수: 10월 → 11월 ('+r.paired.count+'명 집계)');text(28,9,4,crhText_(r.paired.oct,'ratio')+' → '+crhText_(r.paired.nov,'ratio'));
  text(29,1,12,'개인별 비교·순위는 표시하지 않음. 동일 참여 집단은 두 달 모두 자료가 있는 인원만 집계. 효과 배수는 같은 날 두 방식이 섞인 기록일을 제외.',null,'#607087',8);
  text(31,1,12,'03  크롤링과 리스팅을 합친 총 투입 대비 성과','#17334c','#ffffff',11).setFontWeight('bold');
  text(32,1,8,'총 관측 d: 10월 → 11월 (같은 사람·같은 날짜 중복 제거)');text(32,9,4,crhText_(oct.totalD)+' → '+crhText_(nov.totalD));
  text(33,1,8,'공정 합산 관측 VI/d: 10월 → 11월');text(33,9,4,crhText_(oct.totalRate)+' → '+crhText_(nov.totalRate)).setFontWeight('bold');
  const oldCost=oct.complete&&oct.totalRate>0?1000/oct.totalRate:null,newCost=nov.complete&&nov.totalRate>0?1000/nov.totalRate:null;
  text(34,1,8,'같은 1,000 VI 기준 관측 d: 10월 → 11월');text(34,9,4,crhText_(oldCost)+'d → '+crhText_(newCost)+'d');
  text(35,1,12,'총 d는 관측된 크롤링·3P 리스팅 기록일의 합집합. 크롤링 과거 이력이 누락되면 총 d와 VI/d도 불완전하므로 아래 연결률·누락 범위를 함께 확인.',null,'#607087',8);
  text(37,1,12,'04  리더에게 전달할 핵심 해석','#17334c','#ffffff',11).setFontWeight('bold');
  let conclusion;
  if(r.growth>1&&r.cohort.change>1&&r.totalGrowth>1)conclusion='11월은 물량 확대와 함께, 같은 참여 집단의 리스팅 VI/d 및 총 관측 d 대비 성과가 모두 개선 방향. 허브 활용 확대의 기여와 일치하는 관측 근거.';
  else if(r.growth>1)conclusion='11월 팀 리스팅 VI/d 개선은 관찰됨. 인력 확대·업무 구성·기록 누락의 영향을 구분하여 총 투입 및 동일 집단 지표와 함께 판단.';
  else conclusion='월별 유효 표본과 VI/d를 먼저 확인. 수치로 확인되지 않은 향상은 주장하지 않으며 표본 없음은 0이나 효과 없음과 구분.';
  text(38,1,12,conclusion,'#dce9ed','#17334c',10).setFontWeight('bold');sheet.setRowHeight(38,34);
  const q=r.quality;
  const notes=[
    '단위: d=작업자 고유 기록일×CONFIG 근무비율(0.5 등), VI/d=완료 VI÷d. 시간 기록 미사용. 외부 지원 제외. 개인별 우열 비교가 아닌 팀 월별 비교.',
    '기간: 9·10월은 전체, 11월은 기준일까지. 각 월 1~'+r.asOf.slice(8,10)+'일 지표도 표시. 월 총량만으로 생산성 향상을 판단하지 않음.',
    '귀속: 현재 CRAWL_COMPLETED를 과거 작업에 연결. 크롤링 VI는 마지막 CRAWL 기록월로 추정. 허브 사용·도입일·당시 난이도는 별도 기록 없음.',
    '누락: 과거 CRAWL 백업 누락 가능. 기록에 연결되지 않은 준비 VI '+crhText_(q.unknownPreparedVi,'integer')+'; 비율 미설정 '+q.missingRatio.length+'명(기본 1). 총 투입 d는 실제 출근일 전체가 아님.',
    '비교: 동일 집단은 개인별 월 변화의 중앙값만 집계. 난이도·배정 차이가 남는 관측 비교임. 홀드제외일은 전체 완료요청 중 유효 홀드기록 표본만 사용(n='+r.holdCounts.join('/')+').',
    '점검: VI 누락 '+q.invalidListing+'행, CRAWL 날짜 누락 '+q.missingCrawlDate+'행, 크롤링 작업자 누락 '+q.missingCreator.length+'요청. 이 탭은 전체 새로고침으로 변경되지 않음.'
  ];notes.forEach((note,i)=>{text(40+i,1,12,note,null,'#607087',8);sheet.setRowHeight(40+i,25);});
  ss.setActiveSheet(sheet);
}
