const fs=require('fs'),vm=require('vm'),assert=require('node:assert/strict');
const ctx=vm.createContext({console,SpreadsheetApp:{BorderStyle:{SOLID:'SOLID'}}});for(const f of ['Code.gs','Layout.gs'])vm.runInContext(fs.readFileSync('analytics-script/'+f,'utf8'),ctx);
let count=0;const check=f=>{f();count++;},serial=d=>Date.parse(d+'T00:00:00Z')/86400000+25569;
const rh=['Request ID','Node','Status','Requested At','Completed At','크롤링 단계','크롤링 작업자','Total VI','3P Completed VI','Request Type','Blocked Time (Hours)'];
const sh=['Session ID','Worker','Request ID','Stage','Status','Started At','Ended At','Duration Minutes','Completed VI Delta'];
const data={REQUESTS:[rh,
 ['C','N','ONGOING',serial('2026-11-01'),'','CRAWL_COMPLETED','Crawler',1000,500,'3P Listing',0],
 ['M','N','DONE',serial('2026-11-01'),serial('2026-11-05'),'N/A','',250,250,'3P Listing',48],
 ['P','N','TO DO',serial('2026-11-01'),'','CRAWL_REQUESTED','Crawler',0,0,'3P Listing',0]],
 WORK_SESSIONS:[sh,
 ['S1','Crawler','C','CRAWL','PAUSED',serial('2026-11-02'),serial('2026-11-02'),0,''],
 ['S2','Lister','C','3P','PAUSED',serial('2026-11-03'),serial('2026-11-03'),'',500],
 ['S3','Manual','M','3P','COMPLETE',serial('2026-11-03'),serial('2026-11-03'),0,250]]};
const settings={month:'2026-11',crawlers:1,listers:2,targetDays:2};
const build=(d=data,weights={})=>ctx.laBuild_(d,{min:1,external:new Set(),weights,plan:settings},'2026-11-20');
const snapshot=JSON.stringify(data),out=build(),p=out._plan;
check(()=>assert.equal(JSON.stringify(data),snapshot));
check(()=>assert.equal(out._crawl.dayMonthly.find(r=>r[1]==='크롤링')[5],500));
check(()=>assert.equal(p.raw.pendingVi,1000));
check(()=>assert.equal(p.raw.crawlDays,1));
check(()=>assert.equal(p.raw.listingDays,3));
check(()=>assert.equal(p.raw.allDays,4));
check(()=>assert.equal(p.raw.parallel,2));
check(()=>assert.equal(p.raw.sequential,3));
check(()=>assert.equal(p.raw.crawlerFte,0.5));
check(()=>assert.equal(p.raw.listerFte,1.5));
check(()=>assert.equal(p.raw.old1000,4));
check(()=>assert.equal(p.raw.new1000,3));
check(()=>assert.equal(p.impact.find(r=>r[0]==='준비+리스팅 합산 VI/d')[1],1000/3));
check(()=>assert.equal(p.impact.find(r=>r[0]==='직렬 공정의 1작업일 처리용량(병목)')[1],1000));
check(()=>{const o=build(data,{crawler:0.5,lister:0.5});assert.equal(o._crawl.monthlyPreparation.rows[0][10],2000);assert.equal(o._crawl.dayMonthly.find(r=>r[1]==='크롤링')[5],1000);assert.equal(o._plan.raw.new1000,1.5);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.push(['S4','Lister','C','3P','PAUSED',serial('2026-11-03'),serial('2026-11-03'),99999,100]);const o=build(d);assert.equal(o._dayProductivity['작업자 월별'].find(r=>r[1]==='Lister')[6],1);assert.equal(o._crawl.dayMonthly.find(r=>r[1]==='크롤링')[5],600);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS[2][7]=100000;assert.equal(build(d)._plan.raw.new1000,3);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.push(['S4','Lister','C','2P','PAUSED',serial('2026-11-03'),serial('2026-11-03'),0,50]);const r=build(d)._dayProductivity['작업자 월별'].find(r=>r[1]==='Lister');assert.equal(r[6],1);assert.equal(r[9],550);});
check(()=>{const d=structuredClone(data);d.REQUESTS.push(['C2','N','DONE',serial('2026-11-01'),serial('2026-11-05'),'CRAWL_COMPLETED','Crawler',100,100,'3P Listing',0]);d.WORK_SESSIONS.push(['S4','Crawler','C2','CRAWL','PAUSED',serial('2026-11-02'),serial('2026-11-02'),0,'']);const r=build(d)._crawl.monthlyPreparation.rows[0];assert.equal(r[4],1);assert.equal(r[2],1100);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS[1][6]=serial('2026-10-31');d.WORK_SESSIONS[1][5]=serial('2026-10-31');assert.equal(build(d)._crawl.monthlyPreparation.rows[0][0],'2026-10');});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.splice(1,1);assert.equal(build(d)._crawl.monthlyPreparation.rows[0][0],'시점 미확인');assert.equal(build(d)._plan.raw.crawlDays,null);});
check(()=>{const d=structuredClone(data);d.REQUESTS[3][5]='';assert.equal(build(d)._plan.backlog[3][1],1);assert.equal(build(d)._plan.raw.pendingVi,0);});
check(()=>{const o=ctx.laBuild_(data,{min:1,external:new Set(),weights:{},plan:{...settings,crawlRate:2000,listingRate:1000,requestVi:500}},'2026-11-20');assert.equal(o._plan.raw.pendingVi,500);assert.equal(o._plan.raw.crawlDays,0.25);assert.equal(o._plan.raw.listingDays,1);});
check(()=>{const o=ctx.laBuild_(data,{min:1,external:new Set(),plan:{}},'2026-11-20');assert.equal(o._plan.raw.parallel,null);});
const lead=out._monthlyLeadTime[0];
check(()=>assert.equal(lead.hours,96));
check(()=>assert.equal(lead.blocked,48));
check(()=>assert.equal(lead.adjusted,48));
check(()=>{const d=structuredClone(data);d.REQUESTS[2][10]='';assert.equal(build(d)._monthlyLeadTime[0].adjusted,null);});
check(()=>{const d=structuredClone(data);d.REQUESTS[2][10]=200;assert.equal(build(d)._monthlyLeadTime[0].adjusted,null);});
check(()=>{const d=structuredClone(data);d.REQUESTS[2][10]=0;assert.equal(build(d)._monthlyLeadTime[0].adjusted,96);});
check(()=>{const d=structuredClone(data);d.REQUESTS[2][10]=-1;assert.equal(build(d)._monthlyLeadTime[0].adjusted,null);});
const layouts=ctx.laLayouts_(out);
check(()=>assert.equal(layouts['월별 비교'].find(s=>s.title==='요청 리드타임 및 홀드 제외 리드타임').rows[0][3],2));
check(()=>assert.equal(layouts['월별 비교'].find(s=>s.title==='요청 리드타임 및 홀드 제외 리드타임').rows[0][2],4));
check(()=>assert.equal(layouts['투입 계획'].length,4));
check(()=>assert.ok(Object.values(layouts).every(sections=>sections.every(s=>s.rows.every(r=>r.length===s.headers.length)))));
check(()=>assert.ok(Object.values(layouts).every(sections=>sections.every(s=>s.row>=6))));
check(()=>assert.ok(layouts['외부 지원'][0].title.startsWith('대표 지표')));
check(()=>{const d=structuredClone(data);d.REQUESTS[3][2]='ON HOLD';const o=build(d);assert.equal(o._plan.backlog[4][1],1);assert.equal(o._plan.raw.pendingVi,0);assert.equal(o._plan.raw.crawlDays,0);});
check(()=>{const d=structuredClone(data);d.REQUESTS[2][2]='DROP';assert.equal(build(d)._monthlyLeadTime.length,0);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS[2][4]='ACTIVE';assert.equal(build(d)._dayProductivity['작업자 월별'].some(r=>r[1]==='Lister'),false);});
check(()=>{const o=ctx.laBuild_(data,{min:1,external:new Set(['lister']),weights:{lister:0.5},plan:settings},'2026-11-20');assert.equal(o._dayProductivity['외부 지원'][0][6],0.5);assert.equal(o._dayProductivity['외부 지원'][0][9],1000);assert.equal(o._crawl.dayMonthly.find(r=>r[1]==='크롤링')[5],'');});
// Exercise every rendered range; verifies new sections fit their declared sizes.
check(()=>{
 let writes=0;const sheets=new Map();
 function sheet(){return {getMaxRows:()=>10000,getMaxColumns:()=>200,getRange(...args){const n=args[2],m=args[3];let range;range=new Proxy({}, {get(_,method){if(method==='getMergedRanges')return ()=>[];if(method==='setValues')return values=>{assert.equal(values.length,n);assert.ok(values.every(r=>r.length===m));writes++;return range;};return ()=>range;}});return range;},setHiddenGridlines(){},setFrozenRows(){},setFrozenColumns(){},setTabColor(){},setColumnWidth(){},setColumnWidths(){},setRowHeights(){},setRowHeight(){},clear(){}};}
 ctx.laRenderReports_({getSheetByName:n=>sheets.get(n),insertSheet:n=>{const s=sheet();sheets.set(n,s);return s;}},layouts,'2026-11-20 12:00:00');
 assert.equal(sheets.size,10);assert.ok(writes>50);
});
console.log(JSON.stringify({passed:count,failed:0,scope:'Adjusted person-days, mixed-stage deduplication, monthly attribution, forecast math, missing data and hold-adjusted lead times'}));
