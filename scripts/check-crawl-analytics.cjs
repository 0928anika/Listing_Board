const fs=require('fs'),vm=require('vm'),assert=require('node:assert/strict');
const ctx=vm.createContext({console});
for(const f of ['Code.gs','Layout.gs'])vm.runInContext(fs.readFileSync('analytics-script/'+f,'utf8'),ctx);
let count=0;const check=fn=>{fn();count++;};
const serial=d=>Date.parse(d+'T00:00:00Z')/86400000+25569;
const rh=['Request ID','Node','Status','Requested At','Completed At','크롤링 단계','크롤링 필요','크롤링 완료일'];
const sh=['Session ID','Worker','Request ID','Stage','Status','Ended At','Duration Minutes','Completed VI Delta'];
const lh=['Aggregate ID','Status','Request ID','Worker','Stage','Work Date','Duration Minutes','Completed VI Delta'];
const data={REQUESTS:[rh,
 ['C','N1','DONE',serial('2026-11-01'),serial('2026-11-20'),' crawl_completed ',false,''],
 ['M','N1','DONE',serial('2026-11-01'),serial('2026-11-20'),'N/A',true,serial('2026-11-20')],
 ['U','N1','DONE',serial('2026-11-01'),serial('2026-11-20'),'CRAWL_REQUESTED',true,'']],
 WORK_SESSIONS:[sh,
 ['S1','Amy','C','3P','PAUSED',serial('2026-11-01'),60,300],
 ['S2','Amy','M','3P','PAUSED',serial('2026-11-01'),60,100],
 ['S3','Amy','U','3P','PAUSED',serial('2026-11-01'),60,100],
 ['S4','Amy','C','3P','PAUSED',serial('2026-11-01'),0,50],
 ['S5','Support','C','3P','PAUSED',serial('2026-11-01'),60,9000],
 ['S6','Amy','C','2P','PAUSED',serial('2026-11-01'),60,9000],
 ['S7','Amy','C','CRAWL','PAUSED',serial('2026-11-01'),60,9000],
 ['S8','Amy','C','3P','ACTIVE',serial('2026-11-01'),60,9000],
 ['S9','Amy','C','3P','PAUSED',serial('2026-11-07'),120,300],
 ['S10','Amy','M','3P','PAUSED',serial('2026-11-07'),60,100],
 ['S11','Amy','C','3P','PAUSED',serial('2026-11-08'),60,300],
 ['S12','Amy','M','3P','PAUSED',serial('2026-11-08'),60,100]],
 LONG_TERM_SESSIONS:[lh,['A1','COMMITTED','C','Amy','3P','2026-11-15',60,300],['A2','PENDING','C','Amy','3P','2026-11-16',60,9000]]};
const config={min:1,external:new Set(['support'])};
const build=d=>ctx.laBuild_(d,config,'2026-11-20');
const snapshot=JSON.stringify(data),out=build(data),ca=out._crawl;
check(()=>assert.equal(JSON.stringify(data),snapshot));
check(()=>assert.equal(ctx.laCrawlClass_({stage:' crawl_completed '}),'크롤링'));
check(()=>assert.equal(ctx.laCrawlClass_({stage:'N/A'}),'비크롤링'));
check(()=>assert.notEqual(ctx.laCrawlClass_({stage:''}),'비크롤링'));
check(()=>assert.notEqual(ctx.laCrawlClass_({stage:'CRAWL_ONGOING'}),'비크롤링'));
check(()=>assert.equal(ca.monthly.find(r=>r[1]==='크롤링')[2],1250));
check(()=>assert.equal(ca.monthly.find(r=>r[1]==='크롤링')[3],5));
check(()=>assert.equal(ca.monthly.find(r=>r[1]==='크롤링')[4],1200));
check(()=>assert.equal(ca.monthly.find(r=>r[1]==='크롤링')[5],240));
check(()=>assert.equal(ca.monthly.find(r=>r[1]==='크롤링')[8],50));
check(()=>assert.equal(ca.monthly.find(r=>r[1]==='크롤링')[6],1));
check(()=>assert.equal(ca.monthly.find(r=>r[1]==='비크롤링')[5],100));
check(()=>assert.equal(ca.trends[0][10],2.4));
check(()=>assert.equal(ca.trends[0][5],1250/1650));
check(()=>assert.equal(ca.trends[0][6],1250/1550));
check(()=>assert.equal(ca.trends[0][7],1550/1650));
check(()=>assert.equal(ca.workers[0][8],2.4));
check(()=>assert.equal(ca.nodes[0][8],2.4));
check(()=>assert.equal(ca.daily.find(r=>r[0]==='2026-11-01')[9],3));
check(()=>assert.equal(ca.daily.find(r=>r[0]==='2026-11-07')[14],2));
check(()=>assert.equal(ca.daily.find(r=>r[0]==='2026-11-08')[14],2));
check(()=>assert.equal(ca.daily.find(r=>r[0]==='2026-11-15')[12],300));
check(()=>assert.equal(ca.daily.find(r=>r[0]==='2026-11-15')[14],''));
check(()=>{const d=structuredClone(data);d.REQUESTS[1][5]='';assert.equal(build(d)._crawl.monthly.find(r=>r[1]==='크롤링')[2],0);});
check(()=>{const d=structuredClone(data);d.REQUESTS[0][5]='Renamed';assert.equal(build(d)._crawl.monthly.find(r=>r[1]==='크롤링')[2],0);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.push(['S13','Amy','M','3P','PAUSED',serial('2026-11-15'),60,-500]);const c=build(d)._crawl;assert.equal(c.trends[0][5],'');assert.equal(c.trends[0][10],'');});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.push(['S14','Amy','Missing','3P','PAUSED',serial('2026-11-15'),60,5]);assert.ok(build(d)._crawl.reasons.some(r=>r[1]==='요청 메타데이터 없음'));});
const layouts=ctx.laLayouts_(out),sections=layouts['크롤링 효과'];
check(()=>assert.ok(sections.every(s=>[s.title,...s.headers,...s.rows.flat()].every(v=>typeof v!=='string'||(!v.includes('VI/h')&&!v.includes('환산인일')&&!v.includes('등록'))))));
check(()=>assert.equal(sections.length,9));
check(()=>assert.ok(sections.every(s=>s.noTotal&&!s.rows.some(r=>r[0]==='TOTAL'))));
check(()=>assert.ok(sections.every(s=>s.rows.every(r=>r.length===s.headers.length))));
check(()=>assert.ok(sections.every((s,i)=>!i||s.row>sections[i-1].row+sections[i-1].rows.length)));
check(()=>{const empty=build({REQUESTS:[rh],WORK_SESSIONS:[sh]});assert.equal(empty._crawl.daily.length,0);assert.ok(ctx.laLayouts_(empty)['크롤링 효과'].length===9);});
const prepData={REQUESTS:[['Request ID','크롤링 단계','크롤링 작업자','Total VI'],['R1','CRAWL_COMPLETED','Amy',600],['R2','CRAWL_COMPLETED','Amy',400],['R3','CRAWL_COMPLETED','Bob',900],['R4','CRAWL_REQUESTED','Amy',5000]],WORK_SESSIONS:[['Request ID','Stage','Worker','Status','Started At','Ended At','Duration Minutes'],
 ['R1','CRAWL','Amy','PAUSED',serial('2026-11-01'),serial('2026-11-01'),0],
 ['R2','CRAWL','Amy','PAUSED',serial('2026-11-01'),serial('2026-11-01'),0],
 ['R1','CRAWL','AMY','ACTIVE',serial('2026-11-02'),'',0],
 ['R1','3P','Amy','PAUSED',serial('2026-11-03'),serial('2026-11-03'),60],
 ['R1','CRAWL','Amy','PAUSED',serial('2026-12-01'),serial('2026-12-01'),0]]};
const prep=ctx.laCrawlPreparation_(prepData,'2026-11-20'),amy=prep.rows.find(r=>r[0]==='Amy'),bob=prep.rows.find(r=>r[0]==='Bob');
check(()=>assert.equal(amy[1],1000));
check(()=>assert.equal(amy[3],2));
check(()=>assert.equal(amy[4],500));
check(()=>assert.equal(amy[7],1));
check(()=>assert.equal(amy[8],500));
check(()=>assert.equal(bob[4],''));
check(()=>assert.equal(bob[7],0));
check(()=>assert.ok(prep.audit.some(r=>r[0]==='R3'&&r[2].includes('記録')===false&&r[2].includes('기록일 없음'))));
check(()=>{const d=structuredClone(prepData);d.REQUESTS.push(['R5','CRAWL_COMPLETED','Amy',500]);const r=ctx.laCrawlPreparation_(d,'2026-11-20').rows.find(r=>r[0]==='Amy');assert.equal(r[4],'');assert.equal(r[8],500);assert.equal(r[7],2/3);});
check(()=>{const d=structuredClone(prepData);d.REQUESTS[1][3]='600';const r=ctx.laCrawlPreparation_(d,'2026-11-20').rows.find(r=>r[0]==='Amy');assert.equal(r[4],'');});
check(()=>{const d=structuredClone(prepData);d.WORK_SESSIONS.push(['R4','CRAWL','Amy','ACTIVE',serial('2026-11-04'),'',0]);const r=ctx.laCrawlPreparation_(d,'2026-11-20').rows.find(r=>r[0]==='Amy');assert.equal(r[3],3);assert.equal(r[4],1000/3);});
console.log(JSON.stringify({passed:count,failed:0,scope:'Synthetic crawl classification, time-valid productivity, calendar rolling windows, corrections and layouts'}));
