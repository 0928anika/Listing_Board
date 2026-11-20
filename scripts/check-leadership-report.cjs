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
check(()=>assert.equal(report.paired.count,1));
check(()=>assert.equal(report.paired.oct,2));
check(()=>assert.equal(report.paired.nov,4.5));
check(()=>assert.equal(report.paired.change,2.25));
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
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS=d.WORK_SESSIONS.filter(r=>r[3]!=='CRAWL');assert.equal(ctx.crhBuild_(d,config,options).totalGrowth,null);});
check(()=>{let touched=false;assert.throws(()=>ctx.crhRender_({getSheetByName:()=>({getRange:()=>({getValue:()=> 'user content'}),clear:()=>touched=true})},report),/기존 탭/);assert.equal(touched,false);});
check(()=>{
 const d=structuredClone(data);d.REQUESTS[0].push('Started At');d.REQUESTS.slice(1).forEach(r=>r.push(serial(r[0].startsWith('09')?'2026-09-01':r[0].startsWith('10')?'2026-10-01':'2026-11-01')));
 d.WORK_SESSIONS=d.WORK_SESSIONS.filter(r=>r[3]!=='CRAWL'||serial('2026-11-01')<=r[6]);const out=ctx.crhBuild_(d,config,options);
 assert.equal(out.summary['2026-09'].prepared,100);assert.equal(out.summary['2026-10'].prepared,200);assert.equal(out.summary['2026-10'].fallbackVi,200);assert.equal(out.summary['2026-10'].fallbackRequests,1);assert.equal(out.summary['2026-10'].crawlers,1);assert.equal(out.summary['2026-10'].crawlD,0);assert.equal(out.summary['2026-10'].crawlRate,null);assert.equal(out.totalGrowth,null);assert.equal(out.summary['2026-11'].fallbackRequests,0);
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
check(()=>assert.throws(()=>ctx.crhExamples_(exampleRecords,exampleRequests,config,0),/exampleMinDays/));
check(()=>{const source=fs.readFileSync('analytics-script/Code.gs','utf8');assert.ok(!source.includes('buildCrawlerHubLeadershipReport'));assert.ok(!source.includes('Crawler Hub 리더 보고'));});
// Layout QA: bounded single-tab grid, no overlapping merges or personal comparisons.
check(()=>{
 const texts=[],occupied=new Set();let inserts=0,renames=0;
 const sheet={setName:name=>{assert.equal(name,'11월 보고');renames++;},getMaxRows:()=>50,getMaxColumns:()=>12,getRange(...args){if(typeof args[0]==='string')return {getValue:()=> 'CRAWLER HUB | IMPACT REVIEW'};const [row,col,n=1,m=1]=args;assert.ok(row+n-1<=50&&col+m-1<=12);let range;range=new Proxy({}, {get(_,method){if(method==='getMergedRanges')return ()=>[];if(method==='merge')return ()=>{for(let y=row;y<row+n;y++)for(let x=col;x<col+m;x++){const key=y+':'+x;assert.equal(occupied.has(key),false,'overlapping cell '+key);occupied.add(key);}return range;};if(method==='setValue')return value=>{texts.push(String(value));return range;};return ()=>range;}});return range;},clear(){},setColumnWidths(){},setRowHeights(){},setHiddenGridlines(){},setFrozenRows(){},setTabColor(){},setRowHeight(){}};
 ctx.crhRender_({getSheetByName:name=>name==='11월 보고'?null:sheet,insertSheet:()=>{inserts++;return sheet;},setActiveSheet:s=>assert.equal(s,sheet)},{...report,examples});
 assert.equal(renames,1);assert.equal(inserts,0);assert.ok(texts.includes('11월 보고 | CRAWLER HUB'));assert.ok(texts.some(t=>t.includes('W6')));assert.ok(texts.length>100);assert.ok(texts.every(t=>!t.includes('CrawlerA')&&!t.includes('CrawlerB')&&!t.includes('리더')&&!t.includes('VI/h')&&!t.includes('환산인일')&&!t.includes('등록')));
});
console.log(JSON.stringify({passed:count,failed:0,scope:'One-off team-month report, matched cohort, same-day union, no creator comparisons, top-5 listing examples, September coverage, partial month and one-page render bounds'}));
