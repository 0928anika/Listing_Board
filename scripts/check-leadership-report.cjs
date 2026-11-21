const fs=require('fs'),vm=require('vm'),assert=require('node:assert/strict');
const ctx=vm.createContext({console});for(const f of ['Code.gs','Layout.gs','LeadershipReport.gs'])vm.runInContext(fs.readFileSync('analytics-script/'+f,'utf8'),ctx);
let count=0;const check=f=>{f();count++;},serial=d=>Date.parse(d+'T00:00:00Z')/86400000+25569;
const rh=['Request ID','Node','Brand','크롤링 단계','크롤링 작업자','Total VI'];
const sh=['Session ID','Worker','Request ID','Stage','Status','Started At','Ended At','Completed VI Delta'];
const data={REQUESTS:[rh],WORK_SESSIONS:[sh]};let index=0;
function request(id,stage,creator,vi){data.REQUESTS.push([id,'N','B',stage,creator,vi]);}
function session(worker,id,stage,date,vi,status='PAUSED'){data.WORK_SESSIONS.push(['S'+(++index),worker,id,stage,status,serial(date),serial(date),vi]);}
for(const month of ['09','10','11']){
 const prefix='2026-'+month;
 if(month!=='11'){request(month+'C','CRAWL_COMPLETED','CrawlerA',month==='09'?100:200);session('CrawlerA',month+'C','CRAWL',prefix+'-01','');session('L1',month+'C','3P',prefix+'-01',month==='09'?100:200);}
 else {request('11C','CRAWL_COMPLETED','CrawlerB',700);request('11D','CRAWL_COMPLETED','CrawlerA',300);session('CrawlerB','11C','CRAWL',prefix+'-01','');session('CrawlerA','11D','CRAWL',prefix+'-02','');session('L1','11C','3P',prefix+'-01',600);session('L1','11D','3P',prefix+'-02',300);session('CrawlerB','11C','3P',prefix+'-01',100);}
 request(month+'M','N/A','',100);session('L1',month+'M','3P',prefix+'-03',100);
}
request('10LATE','N/A','',1000);session('L2','10LATE','3P','2026-10-31',1000);
const config={external:new Set(),weights:{l1:0.5,crawlera:1,crawlerb:1,l2:1}},options={asOf:'2026-11-20'};
const snapshot=JSON.stringify(data),report=ctx.crhBuild_(data,config,options),nov=report.summary['2026-11'];
check(()=>assert.equal(JSON.stringify(data),snapshot));
check(()=>assert.equal(nov.vi,1100));
check(()=>assert.equal(nov.listingD,2.5));
check(()=>assert.equal(nov.crawlD,2));
check(()=>assert.equal(nov.totalD,3.5));
check(()=>assert.equal(nov.totalRate,1100/3.5));
check(()=>assert.equal(nov.crawlers,2));
check(()=>assert.equal(report.summary['2026-10'].crawlers,1));
check(()=>assert.equal(nov.prepared,1000));
check(()=>assert.equal(nov.crawlRate,500));
check(()=>assert.equal(nov.share,1000/1100));
check(()=>assert.equal(nov.coverage,1));
check(()=>assert.equal(report.summary['2026-10'].vi,1300));
check(()=>assert.equal(report.matchedPeriod['2026-10'].vi,300));
check(()=>assert.equal(report.paired.count,0));
check(()=>assert.equal(report.paired.oct,null));
check(()=>assert.equal(report.paired.nov,null));
check(()=>assert.equal(report.paired.change,null));
check(()=>assert.equal(report.cohort.count,1));
check(()=>assert.equal(report.cohort.oct,300));
check(()=>assert.equal(report.cohort.nov,1000/1.5));
check(()=>assert.equal(report.monthlyRows.length,14));
check(()=>assert.ok(report.monthlyRows.every(r=>r.length===4)));
check(()=>assert.equal(report.self,undefined));
check(()=>assert.equal(report.creators,undefined));
check(()=>assert.equal(report.sourceNode,undefined));
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS=d.WORK_SESSIONS.filter(r=>r[3]!=='CRAWL');const r=ctx.crhBuild_(d,config,options);assert.equal(r.summary['2026-11'].prepared,0);assert.equal(r.summary['2026-11'].coverage,0);assert.equal(r.quality.unknownPreparedVi,1300);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.find(r=>r[1]==='L1'&&r[2]==='11M')[6]=serial('2026-11-01');const r=ctx.crhBuild_(d,config,options);assert.equal(r.paired.count,0);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.push(d.WORK_SESSIONS[2]);assert.throws(()=>ctx.crhBuild_(d,config,options),/중복 Session/);});
check(()=>{const d=structuredClone(data);d.REQUESTS.push(d.REQUESTS[1]);assert.throws(()=>ctx.crhBuild_(d,config,options),/중복 Request/);});
check(()=>{const d=structuredClone(data);d.LONG_TERM_SESSIONS=[['Aggregate ID','Status','Request ID','Worker','Stage','Work Date','Completed VI Delta'],['A1','COMMITTED','09C','L1','3P','2026-09-01',100]];assert.throws(()=>ctx.crhBuild_(d,config,options),/중첩/);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.find(r=>r[2]==='11C'&&r[1]==='L1')[7]='';const r=ctx.crhBuild_(d,config,options);assert.equal(r.quality.invalidListing,1);assert.equal(r.summary['2026-11'].listingD,2.5);assert.equal(r.summary['2026-11'].vi,500);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.find(r=>r[2]==='11C'&&r[1]==='L1')[4]='ACTIVE';const r=ctx.crhBuild_(d,config,options);assert.equal(r.summary['2026-11'].listingD,2);assert.equal(r.summary['2026-11'].vi,500);});
check(()=>{const r=ctx.crhBuild_(data,{...config,external:new Set(['crawlerb'])},options);assert.equal(r.summary['2026-11'].crawlers,1);assert.equal(r.summary['2026-11'].vi,1000);assert.equal(r.summary['2026-11'].prepared,300);});
check(()=>assert.throws(()=>ctx.crhBuild_(data,config,{asOf:'2026-10-20'}),/2026/));
check(()=>assert.throws(()=>ctx.crhBuild_(data,config,{...options,minRecordedDays:0}),/minRecordedDays/));
check(()=>assert.equal(ctx.crhText_(null),'표본 없음'));
check(()=>assert.equal(ctx.crhText_(0),'0'));
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS=d.WORK_SESSIONS.filter(r=>r[3]!=='CRAWL');const out=ctx.crhBuild_(d,config,options);assert.ok(Number.isFinite(out.totalGrowth));assert.equal(out.summary['2026-11'].complete,false);assert.ok(out.summary['2026-11'].uncoveredCrawls>0);});
check(()=>{let touched=false;assert.throws(()=>ctx.crhRender_({getSheetByName:()=>({getRange:()=>({getValue:()=> 'user content'}),clear:()=>touched=true})},report),/기존 탭/);assert.equal(touched,false);});
check(()=>{
 const d=structuredClone(data);d.REQUESTS[0].push('Started At');d.REQUESTS.slice(1).forEach(r=>r.push(serial(r[0].startsWith('09')?'2026-09-01':r[0].startsWith('10')?'2026-10-01':'2026-11-01')));
 d.WORK_SESSIONS=d.WORK_SESSIONS.filter(r=>r[3]!=='CRAWL'||serial('2026-11-01')<=r[6]);const out=ctx.crhBuild_(d,config,options);
 assert.equal(out.summary['2026-09'].prepared,100);assert.equal(out.summary['2026-10'].prepared,200);assert.equal(out.summary['2026-10'].fallbackVi,200);assert.equal(out.summary['2026-10'].fallbackRequests,1);assert.equal(out.summary['2026-10'].crawlers,1);assert.equal(out.summary['2026-10'].crawlD,1);assert.equal(out.summary['2026-10'].estimatedCrawlD,1);assert.equal(out.summary['2026-10'].crawlRate,200);assert.equal(typeof out.totalGrowth,'number');assert.equal(out.summary['2026-11'].fallbackRequests,0);
});
check(()=>{
 const d=structuredClone(data);d.REQUESTS[0].push('Started At','Requested At');d.REQUESTS[1].push(serial('2026-10-05'),serial('2026-09-01'));d.WORK_SESSIONS=d.WORK_SESSIONS.filter(r=>!(r[2]==='09C'&&r[3]==='CRAWL'));const out=ctx.crhBuild_(d,config,options);assert.equal(out.summary['2026-09'].prepared,0);assert.equal(out.summary['2026-10'].prepared,300);
});
check(()=>{
 const d=structuredClone(data);d.REQUESTS[0].push('Started At');d.REQUESTS[2].push(serial('2026-09-05'));const out=ctx.crhBuild_(d,config,options);assert.equal(out.summary['2026-10'].prepared,200);
});
const density=[];
for(const day of ['2026-09-03','2026-09-05','2026-09-07'])density.push({day,name:'Sparse',valid:true});
for(const day of ['2026-09-16','2026-09-17','2026-09-18','2026-10-01','2026-10-02','2026-10-04'])for(const name of ['A','B','C','D'])density.push({day,name,valid:true});
check(()=>assert.equal(ctx.crhSeptemberWindow_(density,'AUTO').start,'2026-09-16'));
check(()=>assert.equal(ctx.crhSeptemberWindow_(density,'2026-09-17').start,'2026-09-17'));
check(()=>assert.equal(ctx.crhSeptemberWindow_([],'AUTO').start,null));
check(()=>assert.throws(()=>ctx.crhSeptemberWindow_(density,'2026-09-31'),/septemberStart/));
check(()=>{
 const d=structuredClone(data);for(const s of density)d.WORK_SESSIONS.push(['D'+d.WORK_SESSIONS.length,s.name,'09M','3P','PAUSED',serial(s.day),serial(s.day),10]);
 const out=ctx.crhBuild_(d,config,options);assert.equal(out.september.start,'2026-09-16');assert.equal(out.summary['2026-09'].vi,120);assert.equal(out.summary['2026-09'].prepared,0);
});
const exampleRecords=[],exampleRequests={C:{stage:'CRAWL_COMPLETED'},N:{stage:'N/A'}};
for(let i=0;i<7;i++)for(const [id,dates,vi] of [['C',['01','02'],(i+1.5)*100],['N',['03','04'],100]])for(const dd of dates)exampleRecords.push({worker:'W'+i,name:'w'+i,id,day:'2026-11-'+dd,vi,valid:true});
const examples=ctx.crhExamples_(exampleRecords,exampleRequests,{weights:{w6:0.5}},2);
check(()=>assert.equal(examples['2026-11'].eligible,7));
check(()=>assert.equal(examples['2026-11'].rows.length,5));
check(()=>assert.equal(examples['2026-11'].rows[0].worker,'W6'));
check(()=>assert.equal(examples['2026-11'].rows[0].ratio,7.5));
check(()=>assert.equal(examples['2026-11'].rows[0].manualD,1));
check(()=>assert.equal(examples['2026-11'].rows[0].manualRate,200));
check(()=>assert.equal(examples['2026-11'].rows.at(-1).worker,'W2'));
check(()=>assert.equal(examples['2026-09'].rows.length,0));
check(()=>assert.equal(ctx.crhExamples_(exampleRecords,exampleRequests,config)['2026-11'].rows.length,5));
check(()=>{assert.equal(report.requestCosts['2026-11'].averageVi,1100/3);assert.equal(report.requestCosts['2026-11'].sizeSample,3);assert.ok(Math.abs(report.requestCosts['2026-11'].totalPerRequest-3.5/3)<1e-10);assert.ok(Math.abs(report.requestCosts['2026-11'].five-3.5/3*5)<1e-10);assert.ok(Math.abs(report.requestCosts['2026-11'].ten-3.5/3*10)<1e-10);});
check(()=>{const row=report.requestCosts['2026-10'];assert.equal(row.averageVi,1300/3);assert.equal(row.totalPerRequest,1);assert.ok(Math.abs(row.sameSizePerRequest-(1100/3)/(1300/3))<1e-10);});
check(()=>{
 const d=structuredClone(data);d.REQUESTS[0].push('Started At');d.REQUESTS.slice(1).forEach(r=>r.push(serial(r[0].startsWith('09')?'2026-09-01':r[0].startsWith('10')?'2026-10-01':'2026-11-01')));d.WORK_SESSIONS=d.WORK_SESSIONS.filter(r=>r[3]!=='CRAWL');
 const out=ctx.crhBuild_(d,{...config,weights:{...config.weights,crawlera:0.5}},options);assert.equal(out.summary['2026-10'].estimatedCrawlD,0.5);assert.equal(out.summary['2026-10'].crawlRate,400);assert.equal(out.summary['2026-11'].estimatedCrawlD,1.5);assert.equal(out.summary['2026-11'].totalD,4);assert.equal(out.summary['2026-11'].complete,true);
});
check(()=>{
 const d=structuredClone(data);d.REQUESTS[0].push('Started At');d.REQUESTS.push(['EX1','N','B','CRAWL_COMPLETED','CrawlerA',100,serial('2026-11-04')],['EX2','N','B','CRAWL_COMPLETED','CrawlerA',100,serial('2026-11-04')]);const out=ctx.crhBuild_(d,{...config,weights:{...config.weights,crawlera:0.5}},options);assert.equal(out.summary['2026-11'].estimatedCrawlD,1);assert.equal(out.summary['2026-11'].estimatedRequests,2);assert.equal(out.summary['2026-11'].totalD,4);
});
check(()=>{assert.equal(report.examples['2026-11'].workerCount,2);assert.equal(report.examples['2026-11'].rows.length,0);assert.equal(report.examples['2026-11'].eligible,0);assert.equal(report.examples['2026-11'].medianRatio,null);});
check(()=>{const rows=exampleRecords.filter(s=>s.day!=='2026-11-02'&&s.day!=='2026-11-04');const out=ctx.crhExamples_(rows,exampleRequests,config,2)['2026-11'];assert.equal(out.rows.length,0);assert.equal(out.eligible,0);});
check(()=>{const d=structuredClone(data);d.REQUESTS.find(r=>r[0]==='10M')[3]='';const out=ctx.crhBuild_(d,config,options);assert.equal(out.summary['2026-10'].complete,false);assert.equal(out.summary['2026-10'].unknownRequests,1);assert.ok(out.requestCosts['2026-10'].totalPerRequest>0);assert.ok(out.requestCosts['2026-10'].five>0);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.find(r=>r[2]==='10M'&&r[3]==='3P')[7]='';const out=ctx.crhBuild_(d,config,options);assert.equal(out.summary['2026-10'].invalidVi,1);assert.ok(out.requestCosts['2026-10'].totalPerRequest>0);});
check(()=>{
 const records=[];for(let i=0;i<6;i++){records.push({worker:'P'+i,name:'p'+i,id:'C',day:'2026-10-01',vi:100,valid:true},{worker:'P'+i,name:'p'+i,id:'C',day:'2026-11-01',vi:200+i*10,valid:true});}
 const out=ctx.crhExamples_(records,exampleRequests,config,2)['2026-11'];assert.equal(out.rows.length,0);assert.equal(out.eligible,0);assert.equal(out.meanRate,225);assert.equal(out.medianRate,225);
});
check(()=>{const source=fs.readFileSync('analytics-script/Code.gs','utf8');assert.ok(!source.includes('buildCrawlerHubLeadershipReport'));assert.ok(!source.includes('Crawler Hub 리더 보고'));});
check(()=>{const w=report.november.weeks.find(w=>w.number===1);assert.equal(w.start,'2026-11-02');assert.equal(w.end,'2026-11-08');assert.equal(w.requests,2);assert.equal(w.crawlRequests,1);assert.equal(w.requestShare,0.5);assert.equal(w.vi,400);assert.equal(w.d,1);assert.equal(w.rate,400);});
check(()=>{const w=report.november.weeks.find(w=>w.number===3);assert.equal(w.start,'2026-11-16');assert.equal(w.end,'2026-11-22');assert.equal(w.through,'2026-11-20');assert.equal(w.partial,true);assert.equal(w.rate,null);});
check(()=>{const w=report.november.weeks.find(w=>w.number===0);assert.equal(w.vi,700);assert.equal(w.requests,1);assert.equal(report.november.weeks.reduce((n,w)=>n+w.vi,0),nov.vi);assert.equal(report.november.rows,undefined);});
check(()=>{
 const d=structuredClone(data);for(const [id,date,delta] of [['11D','2026-11-08',10],['11D','2026-11-09',20],['11M','2026-11-15',30],['11D','2026-11-16',40]])d.WORK_SESSIONS.push(['W'+date,'L1',id,'3P','PAUSED',serial(date),serial(date),delta]);const out=ctx.crhBuild_(d,config,options),weeks=out.november.weeks;
 assert.equal(weeks.find(w=>w.number===1).vi,410);assert.equal(weeks.find(w=>w.number===1).requests,2);assert.equal(weeks.find(w=>w.number===2).vi,50);assert.equal(weeks.find(w=>w.number===2).requests,2);assert.equal(weeks.find(w=>w.number===3).vi,40);
});
check(()=>{const cost=report.requestCosts['2026-11'];assert.equal(cost.withoutCrawlRate,null);assert.equal(cost.withoutCrawlPerRequest,null);assert.equal(cost.withoutCrawlRequests,0);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.push(['BASE','L1','11M','3P','PAUSED',serial('2026-11-04'),serial('2026-11-04'),100]);const out=ctx.crhBuild_(d,config,options),cost=out.requestCosts['2026-11'];assert.equal(cost.withoutCrawlRate,out.examples['2026-11'].totalNRate);assert.equal(cost.withoutCrawlRate,200);assert.equal(cost.withoutCrawlD,1);assert.equal(cost.withoutCrawlRequests,1);assert.ok(Math.abs(cost.withoutCrawlPerRequest-cost.averageVi/200)<1e-10);});
check(()=>{
 const records=[];for(const [name,n,c] of [['A',100,400],['B',1000,2000]])for(const [id,dates,vi] of [['N',['01','02'],n],['C',['03','04'],c]])for(const day of dates)records.push({name:name.toLowerCase(),worker:name,id,day:'2026-11-'+day,vi,valid:true});
 const out=ctx.crhExamples_(records,exampleRequests,config)['2026-11'];assert.equal(out.totalCRate,1200);assert.equal(out.totalNRate,550);assert.equal(out.totalRatio,1200/550);assert.notEqual(out.totalRatio,out.medianRatio);
});
check(()=>{
 const records=[['C','01',9530],['N','02',1000],['N','03',1000]].map(([id,day,vi])=>({name:'sparse',worker:'Sparse',id,day:'2026-10-'+day,vi,valid:true}));assert.equal(ctx.crhExamples_(records,exampleRequests,config)['2026-10'].eligible,0);
 records.push({name:'sparse',worker:'Sparse',id:'C',day:'2026-10-04',vi:9530,valid:true});const out=ctx.crhExamples_(records,exampleRequests,config)['2026-10'];assert.equal(out.eligible,1);assert.equal(out.rows[0].ratio,9.53);
});
check(()=>{
 const records=[];for(const [name,amount] of [['slow',50],['fast',200]])for(const [id,days,vi] of [['N',['01','02'],100],['C',['03','04'],amount]])for(const day of days)records.push({name,worker:name,id,day:'2026-11-'+day,vi,valid:true});
 const out=ctx.crhExamples_(records,exampleRequests,config)['2026-11'];assert.equal(out.eligible,2);assert.equal(out.rows.length,1);assert.equal(out.totalCRate,125);assert.equal(out.totalNRate,100);assert.equal(out.totalRatio,1.25);
});
check(()=>{const out=ctx.crhBuild_(data,config,{asOf:'2026-11-30'}),last=out.november.weeks.at(-1);assert.equal(last.number,5);assert.equal(last.start,'2026-11-30');assert.equal(last.end,'2026-12-06');assert.equal(last.through,'2026-11-30');assert.equal(last.partial,true);});
// Layout QA: bounded single-tab grid, no overlapping merges or personal comparisons.
check(()=>{
 const texts=[],occupied=new Set();let inserts=0,renames=0,writes=0;
 const sheet={setName:name=>{assert.equal(name,'11월 보고');renames++;},getMaxRows:()=>80,getMaxColumns:()=>12,getRange(...args){if(typeof args[0]==='string')return {getValue:()=> 'CRAWLER HUB | IMPACT REVIEW'};const [row,col,n=1,m=1]=args;assert.ok(row+n-1<=80&&col+m-1<=12);let range;range=new Proxy({}, {get(_,method){if(method==='getMergedRanges')return ()=>[];if(method==='merge')return ()=>{for(let y=row;y<row+n;y++)for(let x=col;x<col+m;x++){const key=y+':'+x;assert.equal(occupied.has(key),false,'overlapping cell '+key);occupied.add(key);}return range;};if(method==='setValues')return values=>{assert.equal(values.length,n);assert.ok(values.every(r=>r.length===m));writes++;texts.push(...values.flat().map(String));return range;};if(method==='setValue')return ()=>{throw Error('Unbatched cell write');};return ()=>range;}});return range;},clear(){},setColumnWidths(){},setRowHeights(){},setHiddenGridlines(){},setFrozenRows(){},setTabColor(){},setRowHeight(){}};
 ctx.crhRender_({getSheetByName:name=>name==='11월 보고'?null:sheet,insertSheet:()=>{inserts++;return sheet;},setActiveSheet:s=>assert.equal(s,sheet)},{...report,examples});
 assert.equal(writes,1);assert.equal(renames,1);assert.equal(inserts,0);assert.ok(texts.includes('11월 보고 | CRAWLER HUB'));assert.ok(texts.some(t=>t.includes('W6')));assert.ok(texts.some(t=>t.includes('전체 C VI/d')));assert.ok(texts.some(t=>t.includes('전체 N VI/d')));assert.ok(texts.includes('요청 비중'));assert.ok(texts.includes('N VI/d'));assert.ok(texts.includes('C VI/d'));assert.ok(!texts.some(t=>t.includes('7일 VI/d')));assert.ok(texts.includes('크롤링+리스팅 통합 VI/d'));assert.ok(texts.length>100);assert.ok(texts.every(t=>!t.includes('CrawlerA')&&!t.includes('CrawlerB')&&!t.includes('리더')&&!t.includes('VI/h')&&!t.includes('환산인일')&&!t.includes('등록')&&!t.includes('소표본')));
});
console.log(JSON.stringify({passed:count,failed:0,scope:'One-off team-month report, matched cohort, same-day union, no creator comparisons, top-5 listing examples, September coverage, partial month and one-page render bounds'}));
