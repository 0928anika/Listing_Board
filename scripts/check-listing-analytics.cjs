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
check(()=>{const g=out._workerDays.find(g=>g.month==='2026-10'&&!g.external);assert.equal(g.days,2);assert.equal(g.ratio,1);});
check(()=>{const o=ctx.laBuild_(data,{...config,weights:{amy:0.5}},'2026-10-23');const s=ctx.laLayouts_(o)['일별 기록'].find(s=>s.title==='2026-10 | 작업자별 완료 VI'),r=s.rows[0];assert.equal(r[65],2);assert.equal(r[66],0.5);assert.equal(r[67],-5);assert.equal(r[68],50);assert.equal(r[69],45);assert.equal(r[70],-10);assert.equal(r[71],100);assert.equal(r[72],90);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.push(['ZERO','Amy','R1','3P','PAUSED',serial('2026-10-03'),0,10]);assert.equal(build(d)._workerDays.find(g=>g.month==='2026-10'&&!g.external).days,2);});
check(()=>{const s=layouts['내부·외부 비교'].find(s=>s.title.startsWith('2026-10'));assert.equal(s.rows[0][6],-5);assert.equal(s.rows[0][7],50);assert.equal(s.rows[1][6],0);assert.equal(s.rows[1][7],900);assert.equal(s.rows.at(-1)[6],-10/3);assert.equal(s.rows.at(-1)[7],1000/3);});
check(()=>{const s=layouts['Node별'].find(s=>s.title.startsWith('2026-10')&&s.delayTotal);assert.equal(s.rows[0][4],0);assert.equal(s.rows[0][5],1);assert.equal(s.rows[0][7],1);assert.equal(s.rows.at(-1)[7],1);});
check(()=>{const d=structuredClone(data);d.REQUESTS.push(['SAME','N1','DONE',serial('2026-10-01'),serial('2026-10-02')],['UNKNOWN','N1','DONE','',serial('2026-10-02')]);const s=ctx.laLayouts_(build(d))['Node별'].find(s=>s.title.startsWith('2026-10')&&s.delayTotal);assert.equal(s.rows[0][4],1);assert.equal(s.rows[0][5],1);assert.equal(s.rows[0][6],1);assert.equal(s.rows[0][7],1/3);});
check(()=>{const rows=[['設定','','','','','','작업자','근무비율'],['SOURCE_SPREADSHEET_ID','abcdefghijklmnopqrstuv','','','','','Amy',0.5]];const c=ctx.laConfig_({getId:()=> 'destination',getSheetByName:()=>({getDataRange:()=>({getValues:()=>rows})})});assert.equal(c.weights.amy,0.5);});
check(()=>{const rows=[[],['SOURCE_SPREADSHEET_ID','abcdefghijklmnopqrstuv','','','','','Amy',0]];assert.throws(()=>ctx.laConfig_({getId:()=> 'destination',getSheetByName:()=>({getDataRange:()=>({getValues:()=>rows})})}),/근무비율/);});
check(()=>{const o=ctx.laBuild_(data,{...config,weights:{amy:0.5,support:1}},'2026-10-23'),s=ctx.laLayouts_(o)['내부·외부 비교'].find(s=>s.title.startsWith('2026-10'));assert.equal(s.rows[0][3],1);assert.equal(s.rows[0][10],100);assert.equal(s.rows[1][10],900);assert.equal(s.rows.at(-1)[10],500);});
check(()=>{const d=structuredClone(data);d.REQUESTS.push(['FUTURE','N1','DONE',serial('2026-10-01'),serial('2026-11-01')]);assert.ok(!ctx.laLayouts_(build(d))['Node별'].some(s=>s.title.startsWith('2026-11')));});
check(()=>{const o=ctx.laBuild_({REQUESTS:[rh],WORK_SESSIONS:[sh]},config,'2026-10-23');assert.equal(ctx.laLayouts_(o)['내부·외부 비교'][0].rows[0][0],'기록 없음');});
check(()=>{const s=layouts['일별 기록'].find(s=>s.title==='2026-10 | 작업자별 완료 VI');assert.equal(s.subheaders.length,s.headers.length);assert.equal(s.subheaders[1],'2P');assert.equal(s.subheaders[2],'3P');assert.equal(s.rows[0][1],0);assert.equal(s.rows[0][2],100);assert.equal(s.rows[0][3],-10);assert.equal(s.rows[0][4],0);assert.equal(s.rows[0][5],'');});
check(()=>{const s=layouts['일별 기록'].find(s=>s.title==='2026-10 | 작업자별 작업시간(h)');assert.equal(s.subheaders,undefined);assert.equal(s.headers.length,33);assert.equal(s.rows[0].at(-1),3);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.push(['BOB','Bob','R1','2P','PAUSED',serial('2026-10-01'),60,25]);const s=ctx.laLayouts_(build(d))['일별 기록'].find(s=>s.title==='2026-10 | 작업자별 완료 VI');assert.equal(s.rows.at(-1)[1],25);assert.equal(s.rows.at(-1)[2],100);assert.equal(s.rows.at(-1)[63],15);assert.equal(s.rows.at(-1)[64],100);});
check(()=>{const s=layouts['월별 비교'].find(s=>s.title.startsWith('03  요청 리드타임'));assert.equal(s.rows[0][2],720);assert.equal(s.rows[1][2],720);assert.equal(s.rows[2][2],1);assert.equal(s.rows[0].at(-1),720);});
check(()=>{const d=structuredClone(data);d.REQUESTS.push(['R2','N2','DONE',serial('2026-09-01'),serial('2026-09-02')],['R3','N1','DONE',serial('2026-10-01'),serial('2026-10-03')],['R4','N1','DONE',serial('2026-10-03'),serial('2026-10-01')],['R5','N1','ONGOING',serial('2026-10-01'),serial('2026-10-02')],['R6','N1','DONE','2026-10-01',serial('2026-10-02')]);const s=ctx.laLayouts_(build(d))['월별 비교'].find(s=>s.title.startsWith('03  요청 리드타임'));assert.equal(s.rows[0].at(-1),48);assert.equal(s.rows[1].at(-1),264);assert.equal(s.rows[2].at(-1),3);});
check(()=>{const d=structuredClone(data);d.REQUESTS[1][4]=serial('2026-10-24');assert.equal(build(d)._monthlyLeadTime.length,0);});
check(()=>{const d=structuredClone(data);d.REQUESTS[1][3]=d.REQUESTS[1][4];assert.equal(build(d)._monthlyLeadTime[0].hours,0);});
check(()=>assert.equal(ctx.laReportValue_('2026-10-01','날짜'),serial('2026-10-01')));
check(()=>assert.equal(ctx.laReportValue_(serial('2026-10-01'),'날짜'),serial('2026-10-01')));
check(()=>assert.equal(ctx.laReportValue_('TOTAL','날짜'),'TOTAL'));
check(()=>assert.equal(ctx.laReportValue_('2026-11','월'),"'2026-11"));
check(()=>assert.equal(ctx.laReportValue_('2026-11-01','기간'),"'2026-11-01"));
check(()=>assert.equal(ctx.laReportValue_('2026-11',''),"'2026-11"));
check(()=>assert.equal(layouts['외부 지원'].find(s=>s.title==='외부 일별 작업량·시간').formats[0],'yyyy-mm-dd'));
check(()=>{const t=structuredClone(out);const original=t['주별 비교'][1][0];t['주별 비교'][1]=['2026-09-07','2026-09-14','2026-09-21','2026-09-28','2026-10-05','2026-10-12','2026-10-19'].map(p=>{const r=Array.from(original);r[0]=p;return r;});const sections=ctx.laLayouts_(t)['주별 비교'];assert.equal(sections.length,4);assert.equal(sections[0].headers.slice(1,-1).join(','),'2026-09-07,2026-09-14,2026-09-21,2026-09-28');assert.equal(sections[1].headers.slice(1,-1).join(','),'2026-10-05,2026-10-12,2026-10-19');assert.ok(sections.every(s=>new Set(s.headers.slice(1,-1).map(p=>p.slice(0,7))).size===1));});
check(()=>{let merges=[{breakApart(){assert.equal(this,merges[0]);}},{breakApart(){assert.equal(this,merges[1]);}}],calls=0;const sheet={getMaxRows:()=>100,getMaxColumns:()=>40,getRange(...args){assert.deepEqual(args,[1,1,100,40]);calls++;return {getMergedRanges:()=>merges};}};ctx.laClearReportMerges_(sheet);assert.equal(calls,1);});
check(()=>{let merged=true,cleared=0;const sheet={getMaxRows:()=>100,getMaxColumns:()=>40,getRange(){return {getMergedRanges:()=>merged?[{breakApart(){merged=false;cleared++;}}]:[]};}};ctx.laClearReportMerges_(sheet);ctx.laClearReportMerges_(sheet);assert.equal(cleared,1);assert.equal(merged,false);});
check(()=>ctx.laClearReportMerges_({getMaxRows:()=>1,getMaxColumns:()=>1,getRange:()=>({getMergedRanges:()=>[]})}));
check(()=>assert.equal(ctx.laWeek_('2026-10-04'),'2026-09-28'));
check(()=>assert.equal(ctx.laWeek_('2026-10-05'),'2026-10-05'));
check(()=>assert.equal(layouts['작업자 주별'],undefined));
check(()=>assert.equal(layouts['데이터 점검'],undefined));
check(()=>assert.ok(layouts['외부 지원'].every(s=>!s.title.includes('순위')&&!s.headers.includes('순위'))));
check(()=>{const s=layouts['일별 기록'].find(s=>s.title==='2026-09 | 작업자별 완료 VI');assert.equal(s.headers.length,71);assert.equal(s.headers[1],'2026-09-01');assert.equal(s.headers[60],'2026-09-30');assert.equal(s.headers[61],'월 합계');assert.equal(s.rows[0][62],150);assert.equal(s.rows[0][61],0);});
check(()=>{const s=layouts['일별 기록'].find(s=>s.title==='2026-10 | 작업자별 완료 VI');assert.equal(s.headers[61],'2026-10-31');assert.equal(s.rows[0][63],-10);assert.equal(s.rows[0][64],100);});
check(()=>{const d=structuredClone(data);d.WORK_SESSIONS.push(['LEAP','Amy','R1','3P','PAUSED',serial('2024-02-29'),60,1]);const s=ctx.laLayouts_(build(d))['일별 기록'].find(s=>s.title==='2024-02 | 작업자별 완료 VI');assert.equal(s.headers[57],'2024-02-29');assert.equal(s.headers[59],'월 합계');});
check(()=>assert.equal(Object.keys(layouts).length,8));
check(()=>{const s=layouts['작업자 월별'].find(s=>s.title==='월간 2026-10 | 3P 생산성 순위');const r=s.rows.at(-1);assert.equal(r[0],'TOTAL');assert.equal(r[2],100);assert.equal(r[3],2);assert.equal(r[4],50);});
check(()=>{const s=layouts['일별 기록'].find(s=>s.title==='2026-10 | 작업자별 완료 VI');assert.equal(s.rows.at(-1)[0],'TOTAL');assert.equal(s.rows.at(-1)[64],100);assert.equal(s.rows.at(-1)[63],-10);});
check(()=>{const s=layouts['월별 비교'][0];assert.equal(s.headers.at(-1),'전체');assert.equal(s.rows.find(r=>r[0]==='3P 생산성(VI/h)').at(-1),62.5);});
check(()=>{const s=layouts['작업자 월별'].find(s=>s.title.includes('작업시간 및 작업량'));assert.equal(s.rows.at(-1)[6],'');assert.equal(s.rows.at(-1)[7],'');});
check(()=>assert.ok(layouts['월별 비교'].length>=2));
check(()=>assert.equal(layouts['월별 비교'][0].headers[1],'2026-09'));
check(()=>assert.equal(layouts['월별 비교'][0].rows.find(r=>r[0]==='완료 VI 합계')[1],150));
check(()=>assert.ok(layouts['작업자 월별'].some(s=>s.alongside)));
check(()=>{for(const sections of Object.values(layouts))for(const s of sections)for(const row of s.rows)assert.equal(row.length,s.headers.length);});
check(()=>{for(const sections of Object.values(layouts))for(let i=0;i<sections.length;i++)for(let j=i+1;j<sections.length;j++){const a=sections[i],b=sections[j];const overlapRows=a.row<b.row+b.rows.length+2&&b.row<a.row+a.rows.length+2;const overlapCols=a.col<b.col+b.headers.length&&b.col<a.col+a.headers.length;assert.ok(!(overlapRows&&overlapCols));}});
check(()=>{const empty=ctx.laBuild_({REQUESTS:[rh],WORK_SESSIONS:[sh]},config,'2026-10-23');for(const sections of Object.values(ctx.laLayouts_(empty)))assert.ok(sections.length>0);});
if(process.argv.includes('--preview')){
 const escape=v=>String(v).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;');
 const html='<!doctype html><meta charset="utf-8"><title>Listing A003 Layout Preview</title><style>body{background:#f4f7fb;color:#24354b;font:14px Arial;margin:40px}h1{color:#183555}h2{margin:32px 0 12px;font-size:17px}table{border-collapse:collapse;margin-bottom:32px;min-width:750px}th{background:#304e76;color:white;padding:12px}td{border:1px solid #d6dfeb;padding:10px;text-align:right}td:first-child{text-align:left}tr:nth-child(odd){background:#eaf0f8}tr:nth-child(even){background:white}.page{margin-bottom:70px}.note{color:#62758b}</style><h1>LISTING / A003 배치 미리보기</h1><p class="note">가상 데이터 · 실제 운영 실적 아님 · Google Sheets 적용 시 같은 제목·색상·표 구성을 사용합니다.</p>'+['월별 비교','작업자 월별'].map(n=>'<div class="page"><h1>'+n+'</h1>'+layouts[n].map(s=>'<h2>'+escape(s.title)+'</h2><table><tr>'+s.headers.map(h=>'<th>'+escape(h)+'</th>').join('')+'</tr>'+s.rows.map(row=>'<tr>'+row.map(v=>'<td>'+escape(v)+'</td>').join('')+'</tr>').join('')+'</table>').join('')+'</div>').join('');
 fs.writeFileSync('analytics-script/PREVIEW.html',html);
}
console.log(JSON.stringify({passed:count,failed:0,scope:'Local synthetic aggregation and layout tests; not live execution'}));




