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
check(()=>{const source=fs.readFileSync('analytics-script/Code.gs','utf8');assert.ok(!source.includes('buildCrawlerHubLeadershipReport'));assert.ok(!source.includes('Crawler Hub 리더 보고'));});
// Layout QA: bounded single-tab grid, no overlapping merges or personal comparisons.
check(()=>{
 const texts=[],occupied=new Set();let inserts=0;
 const sheet={getMaxRows:()=>50,getMaxColumns:()=>12,getRange(...args){if(typeof args[0]==='string')return {getValue:()=> 'CRAWLER HUB | IMPACT REVIEW'};const [row,col,n=1,m=1]=args;assert.ok(row+n-1<=50&&col+m-1<=12);let range;range=new Proxy({}, {get(_,method){if(method==='getMergedRanges')return ()=>[];if(method==='merge')return ()=>{for(let y=row;y<row+n;y++)for(let x=col;x<col+m;x++){const key=y+':'+x;assert.equal(occupied.has(key),false,'overlapping cell '+key);occupied.add(key);}return range;};if(method==='setValue')return value=>{texts.push(String(value));return range;};return ()=>range;}});return range;},clear(){},setColumnWidths(){},setRowHeights(){},setHiddenGridlines(){},setFrozenRows(){},setTabColor(){},setRowHeight(){}};
 ctx.crhRender_({getSheetByName:()=>sheet,insertSheet:()=>{inserts++;return sheet;},setActiveSheet:s=>assert.equal(s,sheet)},report);
 assert.equal(inserts,0);assert.ok(texts.length>70);assert.ok(texts.every(t=>!t.includes('CrawlerA')&&!t.includes('CrawlerB')&&!t.includes('L1')&&!t.includes('VI/h')&&!t.includes('환산인일')&&!t.includes('등록')));
});
console.log(JSON.stringify({passed:count,failed:0,scope:'One-off team-month report, matched cohort, same-day union, no personal comparisons, partial month and one-page render bounds'}));
