const fs=require('fs'),vm=require('vm'),assert=require('node:assert/strict');
const ctx=vm.createContext({console});vm.runInContext(fs.readFileSync('analytics-script/Code.gs','utf8'),ctx);
vm.runInContext(fs.readFileSync('analytics-script/Layout.gs','utf8'),ctx);
const serial=d=>Date.parse(d+'T00:00:00Z')/86400000+25569;
const rh=['Request ID','Node','Status','Requested At','Completed At'];
const sh=['Session ID','Worker','Request ID','Stage','Status','Ended At','Duration Minutes','Completed VI Delta'];
const lh=['Aggregate ID','Status','Request ID','Worker','Stage','Work Date','Duration Minutes','Completed VI Delta'];
const data={REQUESTS:[rh,['R1','N1','DONE',serial('2026-09-01'),serial('2026-10-01')]],WORK_SESSIONS:[sh,
 ['S1','Amy','R1','3P','PAUSED',serial('2026-09-30'),60,100],
 ['S2','Amy','R1','3P','COMPLETE',serial('2026-10-01'),120,100],
 ['S3','Support','R1','3P','PAUSED',serial('2026-10-01'),60,900],
 ['','Amy','R1','2P','PAUSED',serial('2026-10-02'),60,-10],
 ['S5','Amy','R1','3P','ACTIVE','',10,20]],LONG_TERM_SESSIONS:[lh,
 ['A1','COMMITTED','R1','Amy','3P','2026-09-20',60,50],
 ['A2','PENDING','R1','Amy','3P','2026-09-21',60,1000]]};
const config={min:1,external:new Set(['support'])};let count=0;
function check(fn){fn();count++;}
const build=d=>ctx.laBuild_(d,config,'2026-10-23'),out=build(data);
check(()=>assert.equal(out['월별 비교'][1][0][4],150));
check(()=>assert.equal(out['월별 비교'][1][1][4],90));
check(()=>assert.equal(out['월별 비교'][1][1][7],3));
check(()=>assert.equal(out['월별 비교'][1][1][8],50));
check(()=>assert.equal(out['월별 비교'][1][1][13],-0.4));
check(()=>assert.equal(out['외부 지원'][1].find(r=>r[0]==='월')[5],900));
check(()=>assert.equal(out['작업자 월별'][1][1][12],1));
check(()=>assert.equal(ctx.laWeek_('2026-10-01'),'2026-09-28'));
check(()=>assert.equal(ctx.laDay_(serial('2026-10-01')+0.9),'2026-10-01'));
check(()=>assert.equal(ctx.laDay_('2026-10-01'),''));
check(()=>assert.ok(out['데이터 점검'][1].some(r=>r[0]==='수기 세션 포함')));
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.push(d.WORK_SESSIONS[1]);assert.throws(()=>build(d),/중복 Session ID/);});
check(()=>{const d=structuredClone(data);d.LONG_TERM_SESSIONS.push(['A3','COMMITTED','R1','Amy','3P','2026-10-01',120,100]);assert.throws(()=>build(d),/겹칩니다/);});
check(()=>{const d=structuredClone(data);d.LONG_TERM_SESSIONS.push(d.LONG_TERM_SESSIONS[1]);assert.throws(()=>build(d),/집계 ID/);});
check(()=>{const d=structuredClone(data);d.REQUESTS.push(d.REQUESTS[1]);assert.throws(()=>build(d),/중복 Request/);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS[1][6]='60';assert.ok(build(d)['데이터 점검'][1].some(r=>r[0]==='VI/시간 누락 또는 오류'));});
check(()=>{const o=ctx.laBuild_(data,{min:1,external:new Set()},'2026-10-23');assert.equal(o['월별 비교'][1][1][4],990);});
check(()=>{const o=ctx.laBuild_(data,{min:3,external:new Set(['support'])},'2026-10-23');assert.equal(o['작업자 월별'][1][1][12],'');});
check(()=>{const d=structuredClone(data);delete d.LONG_TERM_SESSIONS;assert.ok(build(d)['데이터 점검'][1].some(r=>r[0]==='장기 집계 없음'));});
check(()=>{const d=structuredClone(data);d.REQUESTS[0].reverse();d.REQUESTS[1].reverse();assert.equal(build(d)['월별 비교'][1][0][4],150);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.push(['S8','Bob','R1','3P','PAUSED',serial('2026-10-03'),60,50]);const r=build(d)['작업자 월별'][1].filter(r=>r[0]==='2026-10');assert.deepEqual(Array.from(r,r=>r[12]),[1,1]);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.push(['S9','Amy','R1','3P','PAUSED',serial('2026-10-24'),60,10]);assert.equal(build(d)['월별 비교'][1][1][4],90);});
const layouts=ctx.laLayouts_(out);
check(()=>assert.equal(Object.keys(layouts).length,10));
check(()=>{for(const [name,sections] of Object.entries(layouts))for(const s of sections){assert.ok(!s.headers.some(h=>String(h).includes('VI/h')));if(name!=='작업시간')assert.ok(!s.headers.some(h=>/시간\(h\)|\(h\)/.test(h)));for(const row of s.rows)assert.equal(row.length,s.headers.length);if(s.subheaders)assert.equal(s.subheaders.length,s.headers.length);}});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.slice(1).forEach((r,i)=>r[6]=i%2?0:9000);d.LONG_TERM_SESSIONS.slice(1).forEach(r=>r[6]=0);const changed=ctx.laLayouts_(build(d));for(const name of Object.keys(layouts))if(name!=='작업시간')assert.deepEqual(JSON.parse(JSON.stringify(changed[name])),JSON.parse(JSON.stringify(layouts[name])));});
check(()=>{const rows=out._dayProductivity['작업자 월별'],r=rows.find(r=>r[0]==='2026-10');assert.equal(r[5],2);assert.equal(r[6],2);assert.equal(r[8],50);assert.equal(r[7],-5);});
check(()=>{const o=ctx.laBuild_(data,{...config,weights:{amy:0.5}},'2026-10-23'),r=o._dayProductivity['작업자 월별'].find(r=>r[0]==='2026-10');assert.equal(r[5],2);assert.equal(r[6],1);assert.equal(r[8],100);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.push(['SAME','Amy','R1','2P','PAUSED',serial('2026-10-01'),0,10]);const r=build(d)._dayProductivity['작업자 월별'].find(r=>r[0]==='2026-10');assert.equal(r[5],2);assert.equal(r[2],0);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.push(['MISSING','Amy','R1','3P','PAUSED',serial('2026-10-03'),'', '']);const r=build(d)._dayProductivity['작업자 월별'].find(r=>r[0]==='2026-10');assert.equal(r[5],3);assert.equal(r[3],100);assert.equal(r[8],100/3);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.push(['B','Bob','R1','3P','PAUSED',serial('2026-10-03'),0.1,75]);const o=ctx.laBuild_(d,{...config,weights:{amy:0.5}},'2026-10-23'),s=ctx.laLayouts_(o)['작업자 순위'].find(s=>s.title==='2026-10 | 3P VI/d 순위');assert.equal(s.rows[0][1],'Amy');assert.equal(s.rows[0][4],200);assert.equal(s.rows[1][1],'Bob');assert.equal(s.rows.at(-1)[4],175/1.5);});
check(()=>{const o=ctx.laBuild_(data,{...config,minDays:3},'2026-10-23'),s=ctx.laLayouts_(o)['작업자 순위'].find(s=>s.title==='2026-10 | 3P VI/d 순위');assert.equal(s.rows[0][0],'기록 없음');});
check(()=>{const r=out._recordedTime.monthly.find(r=>r[0]==='2026-10'&&r[1]==='Amy');assert.equal(r[6],3);assert.equal(r[3],0);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.push(['CR','Amy','R1','CRAWL','COMPLETE',serial('2026-10-03'),90,'']);const r=build(d)._recordedTime.monthly.find(r=>r[0]==='2026-10'&&r[1]==='Amy');assert.equal(r[3],1.5);assert.equal(r[6],4.5);});
check(()=>{const s=layouts['월별 비교'].find(s=>s.title==='요청 리드타임 및 홀드 제외 리드타임');assert.equal(s.rows[0][2],30);});
check(()=>{const s=layouts['외부 지원'].at(-1);assert.equal(s.formats[0],'yyyy-mm-dd');assert.ok(s.headers.includes('3P VI/d'));});
check(()=>{const rows=[[],['SOURCE_SPREADSHEET_ID','abcdefghijklmnopqrstuv'],['MIN_RANK_HOURS',999],['MIN_RANK_DAYS',2],['','','','','','','Amy',0.5]];const c=ctx.laConfig_({getId:()=> 'destination',getSheetByName:()=>({getDataRange:()=>({getValues:()=>rows})})});assert.equal(c.minDays,2);assert.equal(c.weights.amy,0.5);});
check(()=>assert.equal(ctx.laReportValue_('2026-11','월'),"'2026-11"));
check(()=>assert.equal(ctx.laReportValue_('2026-10-01','날짜'),serial('2026-10-01')));
check(()=>{for(const sections of Object.values(layouts))for(let i=0;i<sections.length;i++)for(let j=i+1;j<sections.length;j++){const a=sections[i],b=sections[j];const overlapRows=a.row<b.row+b.rows.length+2&&b.row<a.row+a.rows.length+2;const overlapCols=a.col<b.col+b.headers.length&&b.col<a.col+a.headers.length;assert.ok(!(overlapRows&&overlapCols));}});
check(()=>{const empty=ctx.laBuild_({REQUESTS:[rh],WORK_SESSIONS:[sh]},config,'2026-10-23');for(const sections of Object.values(ctx.laLayouts_(empty)))assert.ok(sections.length>0);});
check(()=>{let n=0;ctx.laClearReportMerges_({getMaxRows:()=>1,getMaxColumns:()=>1,getRange:()=>({getMergedRanges:()=>[{breakApart(){n++;}}]})});assert.equal(n,1);});
if(process.argv.includes('--preview')){
 const escape=v=>String(v).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;');
 const html='<!doctype html><meta charset="utf-8"><title>Listing A003 Layout Preview</title><style>body{background:#f4f7fb;color:#24354b;font:14px Arial;margin:40px}h1{color:#183555}h2{margin:32px 0 12px;font-size:17px}table{border-collapse:collapse;margin-bottom:32px;min-width:750px}th{background:#304e76;color:white;padding:12px}td{border:1px solid #d6dfeb;padding:10px;text-align:right}td:first-child{text-align:left}tr:nth-child(odd){background:#eaf0f8}tr:nth-child(even){background:white}.page{margin-bottom:70px}.note{color:#62758b}</style><h1>LISTING / A003 배치 미리보기</h1><p class="note">가상 데이터 · 실제 운영 실적 아님 · Google Sheets 적용 시 같은 제목·색상·표 구성을 사용합니다.</p>'+['월별 비교','작업자 월별'].map(n=>'<div class="page"><h1>'+n+'</h1>'+layouts[n].map(s=>'<h2>'+escape(s.title)+'</h2><table><tr>'+s.headers.map(h=>'<th>'+escape(h)+'</th>').join('')+'</tr>'+s.rows.map(row=>'<tr>'+row.map(v=>'<td>'+escape(v)+'</td>').join('')+'</tr>').join('')+'</table>').join('')+'</div>').join('');
 fs.writeFileSync('analytics-script/PREVIEW.html',html);
}
console.log(JSON.stringify({passed:count,failed:0,scope:'Local synthetic aggregation and layout tests; not live execution'}));




