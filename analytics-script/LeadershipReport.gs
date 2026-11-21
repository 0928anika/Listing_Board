/** One-off November numeric dashboard. Never called by refreshListingAnalytics or a trigger. */
const CRAWLER_HUB_REPORT_OPTIONS={
  asOf:'2026-11-21',
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
function crhNewMonth_(){return {vi:0,crawlVi:0,unknownVi:0,unknownWorked:new Set(),listingDays:new Set(),crawlDays:new Set(),union:new Set(),worked:new Set(),crawlWorked:new Set(),linkedWorked:new Set(),prepared:0,preparedRequests:0,fallbackVi:0,fallbackRequests:0,estimatedCrawlD:0,estimatedRequests:0,unestimatedRequests:0,crawlerNames:new Set(),invalidVi:0,missingWeight:new Set()};}
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
  o=Object.assign({minRecordedDays:1,exampleMinDays:2,septemberStart:'AUTO'},o);const parsed=crhRead_(data,c,o);const september=crhSeptemberWindow_(parsed.records,o.septemberStart);const originalInside=parsed.inside;parsed.inside=day=>originalInside(day)&&(day.slice(0,7)!=='2026-09'||(september.start!==null&&day>=september.start));parsed.records=parsed.records.filter(s=>parsed.inside(s.day));const months=['2026-09','2026-10','2026-11'],stats={},comparable={},latest=new Map(),anyCrawl=new Map(),coveredCrawl=new Set();
  months.forEach(m=>{stats[m]=crhNewMonth_();comparable[m]=crhNewMonth_();});
  const dayLimit=Number(o.asOf.slice(8,10));
  parsed.crawlRecords.forEach(s=>{
    if(!anyCrawl.has(s.id)||anyCrawl.get(s.id)<s.day)anyCrawl.set(s.id,s.day);coveredCrawl.add(s.id);const key=JSON.stringify([s.id,s.name]);if(!latest.has(key)||latest.get(key)<s.day)latest.set(key,s.day);
    if(!parsed.inside(s.day))return;const m=s.day.slice(0,7),d=JSON.stringify([s.name,s.day]);stats[m].crawlDays.add(d);stats[m].union.add(d);stats[m].crawlerNames.add(s.name);
    if(Number(s.day.slice(8,10))<=dayLimit){comparable[m].crawlDays.add(d);comparable[m].union.add(d);}
  });
  Object.values(parsed.requests).forEach(r=>{
    if(r.stage!=='CRAWL_COMPLETED')return;
    if((c.external||new Set()).has(crhName_(r.creator)))return;
    if(!r.creator)parsed.quality.missingCreator.add(r.id);
    const sessionDay=latest.get(JSON.stringify([r.id,crhName_(r.creator)]))||anyCrawl.get(r.id),day=sessionDay||laDay_(r.workStart);if(!day){parsed.quality.unknownPreparedVi+=laNumber_(r.total)&&r.total>0?r.total:0;return;}
    if(!parsed.inside(day))return;const m=day.slice(0,7);
    if(laNumber_(r.total)&&r.total>=0){stats[m].prepared+=r.total;stats[m].preparedRequests++;if(r.creator)stats[m].crawlerNames.add(crhName_(r.creator));if(!sessionDay){stats[m].fallbackVi+=r.total;stats[m].fallbackRequests++;if(r.creator){if(!Object.prototype.hasOwnProperty.call(c.weights||{},crhName_(r.creator)))parsed.quality.missingRatio.add(crhName_(r.creator));const estimate=crhWeight_(c,r.creator);stats[m].estimatedCrawlD+=estimate;stats[m].estimatedRequests++;coveredCrawl.add(r.id);if(Number(day.slice(8,10))<=dayLimit)comparable[m].estimatedCrawlD+=estimate;}else stats[m].unestimatedRequests++;}}
  });
  function add(g,s){const d=JSON.stringify([s.name,s.day]),r=parsed.requests[s.id];g.vi+=s.vi;g.listingDays.add(d);g.union.add(d);g.worked.add(s.id);if(!s.valid)g.invalidVi++;
    if(r&&r.stage==='CRAWL_COMPLETED'){g.crawlVi+=s.vi;g.crawlWorked.add(s.id);if(coveredCrawl.has(r.id))g.linkedWorked.add(s.id);}else if(!r||r.stage!=='N/A'){g.unknownVi+=s.vi;g.unknownWorked.add(s.id);}}
  parsed.records.forEach(s=>{const m=s.day.slice(0,7);add(stats[m],s);if(Number(s.day.slice(8,10))<=dayLimit)add(comparable[m],s);});
  const effects=crhWorkerEffects_(parsed.records,parsed.requests,c,o.minRecordedDays),cohort=crhFixedCohort_(parsed.records,c),examples=crhExamples_(parsed.records,parsed.requests,c,o.exampleMinDays);
  function summarize(g){const d=crhDays_(g.listingDays,c),cd=crhDays_(g.crawlDays,c),observedUnion=crhDays_(g.union,c),union=observedUnion+g.estimatedCrawlD;return {complete:g.unestimatedRequests===0&&g.unknownWorked.size===0&&g.invalidVi===0&&(g.crawlWorked.size===0||g.crawlWorked.size===g.linkedWorked.size),vi:g.vi,crawlers:g.crawlerNames.size,listers:new Set(Array.from(g.listingDays,k=>JSON.parse(k)[0])).size,listingD:d,observedCrawlD:cd,estimatedCrawlD:g.estimatedCrawlD,estimatedRequests:g.estimatedRequests,crawlD:cd+g.estimatedCrawlD,observedTotalD:observedUnion,totalD:union,listingRate:crhRate_(g.vi,d),totalRate:crhRate_(g.vi,union),share:g.vi>0&&g.crawlVi>=0&&g.crawlVi<=g.vi?g.crawlVi/g.vi:null,prepared:g.prepared,preparedRequests:g.preparedRequests,fallbackVi:g.fallbackVi,fallbackRequests:g.fallbackRequests,crawlRate:g.unestimatedRequests===0?crhRate_(g.prepared,cd+g.estimatedCrawlD):null,coverage:g.crawlWorked.size?g.linkedWorked.size/g.crawlWorked.size:null,unknownVi:g.unknownVi,invalidVi:g.invalidVi,requests:g.worked.size};}
  const hold={};months.forEach(m=>hold[m]=[]);Object.values(parsed.requests).forEach(r=>{if(r.status!=='DONE'||!laNumber_(r.start)||!laNumber_(r.end)||r.end<r.start||!laNumber_(r.blocked)||r.blocked<0||r.blocked>(r.end-r.start)*24)return;const day=laDay_(r.end);if(!parsed.inside(day)||!hold[day.slice(0,7)])return;hold[day.slice(0,7)].push(r.end-r.start-r.blocked/24);});
  const summary={},matchedPeriod={};months.forEach(m=>{summary[m]=summarize(stats[m]);matchedPeriod[m]=summarize(comparable[m]);});
  const monthlyRows=[
    ['3P 리스팅 VI',...months.map(m=>summary[m].vi)],['리스팅 VI/d',...months.map(m=>summary[m].listingRate)],
    ['크롤링 물량 비중',...months.map(m=>summary[m].share)],['크롤링 VI (귀속월 추정)',...months.map(m=>summary[m].prepared)],
    ['크롤링 참여 인원',...months.map(m=>summary[m].crawlers)],['크롤링 d',...months.map(m=>summary[m].crawlD)],['크롤링 VI/d',...months.map(m=>summary[m].crawlRate)],
    ['크롤링+리스팅 d (관측+추정)',...months.map(m=>summary[m].totalD)],['공정 합산 모형 VI/d',...months.map(m=>summary[m].totalRate)],
    ['같은 집단 내 크롤링 효과 배수',...months.map(m=>crhMedian_(effects.monthly.get(m)||[]))],
    ['각 월 1~'+dayLimit+'일 리스팅 VI/d',...months.map(m=>matchedPeriod[m].listingRate)],
    ['CRAWL 기록 연결률 (관련 요청)',...months.map(m=>summary[m].coverage)],
    ['크롤링 분류 확인 VI 비중',...months.map(m=>summary[m].vi>0&&summary[m].unknownVi>=0&&summary[m].unknownVi<=summary[m].vi?(summary[m].vi-summary[m].unknownVi)/summary[m].vi:null)],
    ['전체 완료요청 홀드제외일 (달력일)',...months.map(m=>crhMedian_(hold[m]))]
  ];
  const growth=crhRatio_(summary['2026-11'].listingRate,summary['2026-10'].listingRate),totalGrowth=summary['2026-11'].complete&&summary['2026-10'].complete?crhRatio_(summary['2026-11'].totalRate,summary['2026-10'].totalRate):null;
  const quality={invalidListing:parsed.quality.invalidListing,missingCrawlDate:parsed.quality.missingCrawlDate,unknownPreparedVi:parsed.quality.unknownPreparedVi,missingRatio:Array.from(parsed.quality.missingRatio),missingCreator:Array.from(parsed.quality.missingCreator),mixedComparisonRecords:effects.mixed};
  return {asOf:o.asOf,september:september,exampleMinDays:o.exampleMinDays,examples:examples,effectCounts:Object.fromEntries(months.map(m=>[m,(effects.monthly.get(m)||[]).length])),summary:summary,matchedPeriod:matchedPeriod,monthlyRows:monthlyRows,paired:effects.paired,cohort:cohort,requestCosts:crhRequestCosts_(parsed,stats,summary),november:crhNovember_(parsed,c,o),holdCounts:months.map(m=>hold[m].length),growth:growth,totalGrowth:totalGrowth,quality:quality};
}
function crhText_(value,style){if(value===null||value===undefined||!Number.isFinite(value))return '표본 없음';if(style==='percent')return (value*100).toFixed(1)+'%';if(style==='ratio')return value.toFixed(2)+'×';return value.toLocaleString('en-US',{maximumFractionDigits:style==='integer'?0:style==='days'?2:1});}
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
function crhRequestCosts_(parsed,stats,summary){
  const result={};Object.keys(stats).forEach(month=>{
    const sizes=Array.from(stats[month].worked,id=>parsed.requests[id]).filter(r=>r&&laNumber_(r.total)&&r.total>0).map(r=>r.total),average=sizes.length?sizes.reduce((a,b)=>a+b,0)/sizes.length:null,s=summary[month];
    const listing=average!==null&&s.listingRate>0?average/s.listingRate:null,total=average!==null&&s.complete&&s.totalRate>0?average/s.totalRate:null;
    result[month]={averageVi:average,sizeSample:sizes.length,workedRequests:stats[month].worked.size,listingPerRequest:listing,totalPerRequest:total,five:total!==null?total*5:null,ten:total!==null?total*10:null};
  });
  const common=result['2026-11'].averageVi;Object.keys(result).forEach(month=>{const s=summary[month];result[month].sameSizePerRequest=common!==null&&s.complete&&s.totalRate>0?common/s.totalRate:null;});return result;
}
function crhExamples_(records,requests,c,minDays){
  if(!Number.isInteger(minDays)||minDays<1)throw Error('exampleMinDays는 1 이상의 정수여야 합니다.');
  const mode=(_,r)=>r&&r.stage==='CRAWL_COMPLETED'?'crawl':r&&r.stage==='N/A'?'manual':null,pure=crhPure_(records,requests,mode),groups=new Map(),all=new Map(),display=new Map(),rawModes=new Map(),result={};
  records.forEach(s=>{display.set(s.name,s.worker);const key=JSON.stringify([s.day.slice(0,7),s.name]);if(!all.has(key))all.set(key,crhBucket_());crhAdd_(all.get(key),s);if(!rawModes.has(key))rawModes.set(key,new Set());const k=mode(s,requests[s.id]);if(k)rawModes.get(key).add(k);});
  pure.groups.forEach((bucket,key)=>{const [month,name,k]=JSON.parse(key),group=JSON.stringify([month,name]);if(!groups.has(group))groups.set(group,{});groups.get(group)[k]=bucket;});
  const mean=a=>a.length?a.reduce((n,v)=>n+v,0)/a.length:null;
  ['2026-09','2026-10','2026-11'].forEach(month=>result[month]={eligible:0,improved:0,strictEligible:0,smallSample:0,workerCount:0,oneMode:0,mixedExcluded:0,rows:[],rates:[],ratios:[],candidates:[],allWorkers:[]});
  all.forEach((g,key)=>{const [month,name]=JSON.parse(key),rate=crhRate_(g.vi,crhDays_(g.days,c)),out=result[month];out.workerCount++;if(rate!==null){out.rates.push(rate);out.allWorkers.push({worker:display.get(name),name,rate,d:crhDays_(g.days,c)});}const kinds=rawModes.get(key),pair=groups.get(key);if(kinds.size<2)out.oneMode++;else if(!pair||!pair.crawl||!pair.manual)out.mixedExcluded++;});
  groups.forEach((g,key)=>{
    if(!g.crawl||!g.manual)return;const [month,name]=JSON.parse(key),manualD=crhDays_(g.manual.days,c),crawlD=crhDays_(g.crawl.days,c),manualRate=crhRate_(g.manual.vi,manualD),crawlRate=crhRate_(g.crawl.vi,crawlD),ratio=crhRatio_(crawlRate,manualRate);if(ratio===null)return;
    const out=result[month],small=g.crawl.days.size<minDays||g.manual.days.size<minDays;out.eligible++;out.ratios.push(ratio);if(small)out.smallSample++;else out.strictEligible++;
    if(ratio>1){out.improved++;out.candidates.push({worker:display.get(name),name,kind:small?'C/N·소표본':'C/N',manualRate,crawlRate,ratio,manualD,crawlD,manualVi:g.manual.vi,crawlVi:g.crawl.vi});}
  });
  Object.keys(result).forEach(month=>{
    const out=result[month];out.candidates.sort((a,b)=>b.ratio-a.ratio||b.crawlVi-a.crawlVi||a.worker.localeCompare(b.worker));out.rows=out.candidates.slice(0,5);const selected=new Set(out.rows.map(r=>r.name));
    const previous=month==='2026-11'?'2026-10':month==='2026-10'?'2026-09':null,changes=[];
    out.allWorkers.forEach(g=>{if(selected.has(g.name)||!previous)return;const prior=all.get(JSON.stringify([previous,g.name]));if(!prior)return;const before=crhRate_(prior.vi,crhDays_(prior.days,c)),ratio=crhRatio_(g.rate,before);if(ratio!==null&&ratio>1)changes.push({worker:g.worker,name:g.name,kind:'전월비',manualRate:before,crawlRate:g.rate,ratio,manualD:crhDays_(prior.days,c),crawlD:g.d});});
    changes.sort((a,b)=>b.ratio-a.ratio||a.worker.localeCompare(b.worker));for(const g of changes){if(out.rows.length>=5)break;out.rows.push(g);selected.add(g.name);}
    for(const g of out.allWorkers.slice().sort((a,b)=>b.rate-a.rate||a.worker.localeCompare(b.worker))){if(out.rows.length>=5)break;if(selected.has(g.name))continue;out.rows.push({worker:g.worker,name:g.name,kind:'월 생산성 참고',manualRate:null,crawlRate:g.rate,ratio:null,manualD:null,crawlD:g.d});selected.add(g.name);}
    out.medianRate=crhMedian_(out.rates);out.meanRate=mean(out.rates);out.medianRatio=crhMedian_(out.ratios);out.meanRatio=mean(out.ratios);delete out.candidates;delete out.allWorkers;
  });return result;
}
function crhNovember_(parsed,c,o){
  const count=Number(o.asOf.slice(8,10)),daily=new Map();
  for(let d=1;d<=count;d++)daily.set('2026-11-'+String(d).padStart(2,'0'),{vi:0,crawlVi:0,unknownVi:0,days:new Set(),workers:new Set()});
  parsed.records.filter(s=>daily.has(s.day)).forEach(s=>{const g=daily.get(s.day),r=parsed.requests[s.id];g.vi+=s.vi;g.days.add(JSON.stringify([s.name,s.day]));g.workers.add(s.name);if(r&&r.stage==='CRAWL_COMPLETED')g.crawlVi+=s.vi;else if(!r||r.stage!=='N/A')g.unknownVi+=s.vi;});
  function pool(groups){const vi=groups.reduce((n,g)=>n+g.vi,0),crawl=groups.reduce((n,g)=>n+g.crawlVi,0),d=groups.reduce((n,g)=>n+crhDays_(g.days,c),0);return {vi,d,share:vi>0&&crawl>=0&&crawl<=vi?crawl/vi:null,rate:crhRate_(vi,d),activeDays:groups.filter(g=>g.days.size).length};}
  const keys=Array.from(daily.keys()),rows=keys.map((date,i)=>{const g=daily.get(date),rolling=pool(keys.slice(Math.max(0,i-6),i+1).map(d=>daily.get(d)));return [date,g.vi>0&&g.crawlVi>=0&&g.crawlVi<=g.vi?g.crawlVi/g.vi:null,g.vi,crhRate_(g.vi,crhDays_(g.days,c)),rolling.rate,g.workers.size];});
  const split=Math.floor(count/2),first=pool(keys.slice(0,split).map(d=>daily.get(d))),second=pool(keys.slice(split).map(d=>daily.get(d))),low=pool(Array.from(daily.values()).filter(g=>g.vi>0&&g.crawlVi/g.vi<0.5)),high=pool(Array.from(daily.values()).filter(g=>g.vi>0&&g.crawlVi/g.vi>=0.5));
  return {rows,split,first,second,low,high,halfRatio:crhRatio_(second.rate,first.rate),shareRatio:crhRatio_(high.rate,low.rate)};
}
function crhRender_(ss,r){
  const name='11월 보고';let sheet=ss.getSheetByName(name);
  if(!sheet){const old=ss.getSheetByName('Crawler Hub 리더 보고');if(old&&old.getRange('A1').getValue()==='CRAWLER HUB | IMPACT REVIEW'){old.setName(name);sheet=old;}}
  if(sheet&&!['CRAWLER HUB | IMPACT REVIEW','11월 보고 | CRAWLER HUB'].includes(sheet.getRange('A1').getValue()))throw Error('동일 이름의 기존 탭이 있습니다. 이름을 바꿔 보존한 뒤 다시 실행하세요.');
  if(!sheet)sheet=ss.insertSheet(name);
  const height=72;if(sheet.getMaxRows()<height)sheet.insertRowsAfter(sheet.getMaxRows(),height-sheet.getMaxRows());if(sheet.getMaxColumns()<12)sheet.insertColumnsAfter(sheet.getMaxColumns(),12-sheet.getMaxColumns());
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).getMergedRanges().forEach(x=>x.breakApart());sheet.clear();sheet.getRange(1,1,height,12).setBackground('#f4f7fb').setFontFamily('Arial').setFontSize(9).setFontColor('#23324b').setVerticalAlignment('middle');sheet.setColumnWidths(1,12,94);sheet.setRowHeights(1,height,19);sheet.setHiddenGridlines(true);sheet.setFrozenRows(0);sheet.setTabColor('#276b73');
  const grid=Array.from({length:height},()=>Array(12).fill('')),merges=[];sheet.getRange(1,1,height,12).setNumberFormat('@');
  function text(row,col,width,value,bg,color,size){const cell=sheet.getRange(row,col,1,width);if(width>1)merges.push([row,col,1,width]);grid[row-1][col-1]=laReportValue_(String(value),'');cell.setWrap(true);if(bg)cell.setBackground(bg);if(color)cell.setFontColor(color);if(size)cell.setFontSize(size);return cell;}
  function section(row,title){text(row,1,12,title,'#15364a','#ffffff',11).setFontWeight('bold');sheet.setRowHeight(row,24);}
  const f=crhText_,months=['2026-09','2026-10','2026-11'],sep=r.september.start?'9/'+Number(r.september.start.slice(8,10))+'~30':'9월 기록 없음',nov=r.summary['2026-11'],oct=r.summary['2026-10'],labels=[sep,'10월','11월 1~'+r.asOf.slice(8,10)+'일'];
  const change=x=>x===null?'비교 없음':(x>=1?'+':'')+((x-1)*100).toFixed(1)+'%',value=(v,style)=>typeof v==='string'?v:f(v,style);
  function monthHeader(row,label){text(row,1,6,label);labels.forEach((m,j)=>text(row,7+j*2,2,m,'#dce9ed').setFontWeight('bold'));}
  function monthRow(row,label,fn,style,bold){text(row,1,6,label);months.forEach((m,j)=>text(row,7+j*2,2,value(fn(m),style),j===2?'#e2f0e9':row%2?'#ffffff':'#eaf0f5').setHorizontalAlignment('right').setFontWeight(bold?'bold':'normal'));}
  text(1,1,12,'11월 보고 | CRAWLER HUB','#15364a','#ffffff',22).setFontWeight('bold');sheet.setRowHeight(1,38);
  text(2,1,12,sep+' · 10월 전체 · 11월 기준일 '+r.asOf+' | 내부 3P | d=CONFIG 근무비율 보정 근무일',null,'#607087',9);
  const cards=[['11월 리스팅 VI',f(nov.vi,'integer'),'완료 물량'],['10월 대비 VI/d',f(r.growth,'ratio'),change(r.growth)],['11월 C/N 효과 중앙값',f(r.examples['2026-11'].medianRatio,'ratio'),r.examples['2026-11'].eligible+'명 비교'],['같은 참여 집단의 변화',f(r.cohort.change,'ratio'),change(r.cohort.change)+' · '+r.cohort.count+'명']];
  cards.forEach((card,i)=>{const col=1+i*3;text(4,col,3,card[0],'#dce9ed','#33586a',9);text(5,col,3,card[1],'#ffffff','#15364a',24).setFontWeight('bold');text(6,col,3,card[2],'#ffffff','#38726c',9);});sheet.setRowHeight(5,37);
  section(8,'01  월별 물량·생산성·투입 d');monthHeader(9,'지표');
  const matrix=[['리스팅 VI',m=>r.summary[m].vi,'integer'],['리스팅 VI/d',m=>r.summary[m].listingRate,''],['크롤링 VI',m=>r.summary[m].prepared,'integer'],['크롤링 비중 (리스팅 VI 기준)',m=>r.summary[m].share,'percent'],['전체 C/N 배수 중앙값',m=>r.examples[m].medianRatio,'ratio'],['작업한 요청 수',m=>r.summary[m].requests,'integer'],['준비+리스팅 d (관측+추정)',m=>r.summary[m].complete?r.summary[m].totalD:'기록 확인 필요','days'],['그중 요청당 1일 가정 크롤링 d',m=>r.summary[m].estimatedCrawlD,'days']];
  matrix.forEach((row,i)=>monthRow(10+i,row[0],row[1],row[2],[1,4,6].includes(i)));
  section(19,'02  작업자 사례 TOP 5 — C/N 우선, 부족 시 전월비·월 생산성 참고');
  months.forEach((m,i)=>{
    const col=1+i*4,g=r.examples[m];text(20,col,4,labels[i]+' · 월 전체 '+g.workerCount+'명 / C/N '+g.eligible+'명','#dce9ed',null,9).setFontWeight('bold');
    ['작업자·기준','기준 VI/d','현재 VI/d','배수'].forEach((v,j)=>text(21,col+j,1,v,'#e7edf3',null,8));
    for(let k=0;k<5;k++){const row=22+k,item=g.rows[k];sheet.setRowHeight(row,34);if(!item){text(row,col,4,g.workerCount<5?'월 작업자 '+g.workerCount+'명':'비교 가능한 사례 없음','#ffffff','#9aa7b6',8);continue;}
      text(row,col,1,item.worker+'\n'+item.kind,'#ffffff',null,8);text(row,col+1,1,item.manualRate===null?'—':f(item.manualRate),'#ffffff',null,8);text(row,col+2,1,f(item.crawlRate),'#ffffff',null,8);text(row,col+3,1,item.ratio===null?'참고':f(item.ratio,'ratio'),item.kind==='월 생산성 참고'?'#e7edf3':'#d7eddf','#1d6448',11).setFontWeight('bold');
    }
    text(27,col,4,'전체 VI/d 중앙 '+f(g.medianRate)+' / 평균 '+f(g.meanRate)+' · '+g.workerCount+'명','#dce9ed',null,9);
    text(28,col,4,'전체 C/N 중앙 '+f(g.medianRatio,'ratio')+' / 평균 '+f(g.meanRatio,'ratio')+' · '+g.eligible+'명','#dce9ed',null,9);
    text(29,col,4,'한 방식만 '+g.oneMode+'명 · 혼합일 제외 '+g.mixedExcluded+'명 · C/N 소표본 '+g.smallSample+'명',null,'#607087',8);
  });sheet.setRowHeight(27,30);sheet.setRowHeight(28,30);sheet.setRowHeight(29,29);
  text(30,1,12,'C/N=비크롤링→크롤링, 전월비=전월→이번월, 참고=월 VI/d만 표시. 소표본은 한쪽이 '+r.exampleMinDays+'기록일 미만. 평균·중앙값은 상위 5명이 아닌 전체 해당 집단.',null,'#607087',8);sheet.setRowHeight(30,27);
  section(32,'03  요청 몇 개에 몇 d가 필요한가? — 평균 요청 크기 기준 계획값');monthHeader(33,'작업한 요청의 현재 크기로 환산');
  monthRow(34,'요청 1개 평균 VI (표본 요청 수)',m=>f(r.requestCosts[m].averageVi)+' VI · n='+r.requestCosts[m].sizeSample+'/'+r.requestCosts[m].workedRequests,'');
  monthRow(35,'리스팅만 · 요청 1개 필요 d',m=>r.requestCosts[m].listingPerRequest,'days');
  monthRow(36,'준비+리스팅 · 요청 1개 필요 d',m=>r.requestCosts[m].totalPerRequest,'days',true);
  monthRow(37,'준비+리스팅 · 요청 5개 필요 d',m=>r.requestCosts[m].five,'days',true);
  monthRow(38,'준비+리스팅 · 요청 10개 필요 d',m=>r.requestCosts[m].ten,'days');
  monthRow(39,'동일 크기 비교 · 11월 평균 VI의 요청 1개',m=>r.requestCosts[m].sameSizePerRequest,'days');
  text(40,1,12,'필요 d=평균 요청 VI÷해당 월 VI/d. 크롤링 세션 없는 완료 요청은 Started At에 요청당 1일×CONFIG 비율을 추정 추가(동일 날짜 여러 요청도 각각 추정). 인원 배분·대기시간은 별도.',null,'#607087',8);sheet.setRowHeight(40,30);
  section(42,'04  11월 일별 심층 — 크롤링 비중과 리스팅 효과');
  const daily=r.november.rows,left=Math.ceil(daily.length/2),headers=['일자','C 비중','완료 VI','VI/d','7일 VI/d','작업자'];
  [1,7].forEach(col=>headers.forEach((label,i)=>text(43,col+i,1,label,'#dce9ed',null,8)));
  for(let i=0;i<left;i++)for(let side=0;side<2;side++){
    const item=daily[i+side*left];if(!item)continue;const values=[item[0].slice(5),f(item[1],'percent'),f(item[2],'integer'),f(item[3]),f(item[4]),f(item[5],'integer')];
    values.forEach((v,j)=>text(44+i,1+side*6+j,1,v,i%2?'#ffffff':'#eaf0f5',null,8));
  }
  const end=45+left,n=r.november;
  text(end,1,6,'1~'+n.split+'일: C '+f(n.first.share,'percent')+' / '+f(n.first.rate)+' VI/d','#dce9ed');text(end,7,6,(n.split+1)+'~'+r.asOf.slice(8,10)+'일: C '+f(n.second.share,'percent')+' / '+f(n.second.rate)+' VI/d','#dce9ed');
  text(end+1,1,12,'후반/전반 VI/d '+f(n.halfRatio,'ratio')+' | C비중 50% 미만 '+f(n.low.rate)+' VI/d ('+n.low.activeDays+'일) → 50% 이상 '+f(n.high.rate)+' VI/d ('+n.high.activeDays+'일), '+f(n.shareRatio,'ratio'),null,'#15364a',10).setFontWeight('bold');sheet.setRowHeight(end+1,28);
  const notes=[
    '9월 '+(r.september.start||'없음')+'부터 ('+r.september.method+'). C/N은 같은 날 혼합일 제외. 전체 C/N 평균은 개인별 배수의 산술평균, 중앙값은 개인별 배수의 중앙값.',
    '요청 크기는 해당 월 작업한 고유 요청의 현재 Total VI 평균. 필요 d는 관측+추정 생산성에 이 크기를 적용한 모형값으로 실제 요청별 소요기간의 평균은 아님. 요청 크기 차이는 동일 크기 행으로 비교.',
    '추정 크롤링 d는 누락 요청당 1일×현재 작업자 근무비율. 실제 세션의 사람·날짜는 중복 제거하고 추정분은 요청별로 추가. 관측 d와 추정 d를 합친 계획 기준이며 실측 노동시간은 아님.',
    '11월 일별 7일 VI/d는 최근 7개 달력일 합산 VI÷합산 d. 전반/후반 및 C비중 구간 비교는 관측 동행성으로, 작업자·난이도·배정 차이를 제거한 인과 효과가 아님.',
    '점검: VI 누락 '+r.quality.invalidListing+'행, 월 미귀속 준비 VI '+f(r.quality.unknownPreparedVi,'integer')+', 비율 미설정 '+r.quality.missingRatio.length+'명(기본 1). 개인 크롤링 제공자 비교 없음. 별도 실행 시만 갱신.'
  ];notes.forEach((note,i)=>{text(end+3+i,1,12,note,null,'#607087',8);sheet.setRowHeight(end+3+i,26);});sheet.getRange(1,1,height,12).setValues(grid);merges.forEach(args=>sheet.getRange(...args).merge());ss.setActiveSheet(sheet);
}
