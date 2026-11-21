/** A008: all productivity comparisons use VI/d. Hours live only on 작업시간. */
function laLayouts_(tabs){
  const layouts={},metrics=tabs._dayProductivity||{},workerDays=tabs._workerDays||[],leadSamples=tabs._monthlyLeadTime||[],today=tabs._today||'';
  function section(list,title,headers,rows,formats,col){const s={title,headers,rows:rows.length?rows:[headers.map((_,i)=>i===0?'기록 없음':'')],formats:formats||{},col:col||2,noTotal:true};list.push(s);return s;}
  const dayHeaders=['기간','작업자/범위','2P 완료 VI','3P 완료 VI','단계 완료 VI 합계','기록일 합계','d','2P VI/d','3P VI/d','단계 VI/d'];
  const dayFormats={0:'@',1:'@',2:'#,##0',3:'#,##0',4:'#,##0',5:'#,##0',6:'0.00',7:'0.0',8:'0.0',9:'0.0'};
  const sum=(rows,i)=>rows.reduce((n,r)=>n+(typeof r[i]==='number'?r[i]:0),0);
  function dayTable(list,title,rows,headers){return section(list,title,headers||dayHeaders,rows.map(r=>r.slice(0,10)),dayFormats);}
  function matrix(list,rows,weekly){
    const periods=Array.from(new Set(rows.map(r=>r[0]))).sort(),blocks=[];
    periods.forEach(p=>{let b=blocks[blocks.length-1];if(!b||b.length>=6||weekly&&b[0].slice(0,7)!==p.slice(0,7)){b=[];blocks.push(b);}b.push(p);});
    if(!blocks.length){section(list,'01  내부 VI/d 비교',['지표','값'],[]);return;}
    blocks.forEach(ps=>{const source=rows.filter(r=>ps.includes(r[0])),totalD=sum(source,6),items=[['3P 완료 VI',3,'#,##0'],['2P 완료 VI',2,'#,##0'],['단계 완료 VI 합계',4,'#,##0'],['d',6,'0.00'],['3P VI/d',8,'0.0'],['2P VI/d',7,'0.0'],['단계 VI/d',9,'0.0']];
      const body=items.map(([label,index])=>[label,...ps.map(p=>{const r=source.find(r=>r[0]===p);return r?r[index]:'';}),[7,8,9].includes(index)?totalD>0?sum(source,index===7?2:index===8?3:4)/totalD:'':sum(source,index)]);
      const s=section(list,'01  내부 VI/d 비교 | '+ps[0]+' ~ '+ps[ps.length-1],['지표',...ps,'전체'],body);s.rowFormats=items.map(m=>m[2]);
    });
  }
  ['월별 비교','주별 비교'].forEach(name=>{const rows=metrics[name]||[],list=layouts[name]=[],weekly=name==='주별 비교';matrix(list,rows,weekly);
    const lookup=new Map(rows.map(r=>[r[0],r]));const changes=rows.map(r=>{const d=new Date((weekly?r[0]:r[0]+'-01')+'T00:00:00Z');if(weekly)d.setUTCDate(d.getUTCDate()-7);else d.setUTCMonth(d.getUTCMonth()-1);const key=d.toISOString().slice(0,weekly?10:7),prior=lookup.get(key),rate=prior?prior[9]:'';return [r[0],r[3],r[2],r[6],r[8],r[7],r[9],rate,typeof rate==='number'&&rate>0?(r[9]/rate-1):'',r[0]===(weekly?laWeek_(today):today.slice(0,7))?'진행 중':'기록 기간'];});
    section(list,'02  직전 달력기간 대비 VI/d',['기간','3P VI','2P VI','d','3P VI/d','2P VI/d','단계 VI/d','직전 기간 단계 VI/d','생산성 증감률','기간 상태'],changes,{0:'@',1:'#,##0',2:'#,##0',3:'0.00',8:'+0.0%;-0.0%;0.0%',9:'@'});
  });
  const leadMonths=Array.from(new Set(leadSamples.map(s=>s.month))).sort().reverse();
  function median(a){const b=a.slice().sort((a,b)=>a-b),n=b.length;return n?(b[Math.floor((n-1)/2)]+b[Math.floor(n/2)])/2:'';}
  section(layouts['월별 비교'],'요청 전체 경과기간 (달력일)',['월','유효 완료 요청 수','전체 경과일 중앙값','전체 경과일 평균'],leadMonths.map(m=>{const rows=leadSamples.filter(s=>s.month===m),days=rows.map(s=>s.hours/24);return [m,rows.length,median(days),days.reduce((a,b)=>a+b,0)/days.length];}),{0:'@',1:'#,##0'});
  const worker=layouts['작업자 월별']=[],workerRows=metrics['작업자 월별']||[];
  const periods=Array.from(new Set(workerRows.map(r=>r[0]))).sort().reverse();
  periods.forEach(month=>{const rs=workerRows.filter(r=>r[0]===month);['3P','2P'].forEach((stage,j)=>{const vi=j===0?3:2,rate=j===0?8:7,stageDays=j===0?11:10;const eligible=rs.filter(r=>r[5]>=(tabs._minRankDays||1)&&r[6]>0&&(r[stageDays]===undefined?r[vi]!==0:r[stageDays]>0)).sort((a,b)=>b[rate]-a[rate]||a[1].localeCompare(b[1]));let previous=null,rank=0;
      const ranked=eligible.map((r,i)=>{if(r[rate]!==previous)rank=i+1;previous=r[rate];return [rank,r[1],r[vi],r[6],r[rate]];});if(ranked.length)ranked.push(['TOTAL','',sum(eligible,vi),sum(eligible,6),sum(eligible,6)>0?sum(eligible,vi)/sum(eligible,6):'']);
      const s=section(worker,month+' | '+stage+' VI/d 순위',['순위','작업자','완료 VI','d','VI/d'],ranked,{0:'@',1:'@',2:'#,##0',3:'0.00',4:'0.0'},j===0?2:8);if(j===1)s.alongside=true;
    });dayTable(worker,month+' | 작업자 VI/d',rs);});if(!periods.length)dayTable(worker,'대표 지표 | 작업자 VI/d',[]);
  const daily=layouts['일별 기록']=[],dailyRows=metrics['일별 기록']||[],dailyMonths=Array.from(new Set(dailyRows.map(r=>r[0].slice(0,7)))).sort().reverse();
  dailyMonths.forEach(month=>{const count=new Date(Date.UTC(Number(month.slice(0,4)),Number(month.slice(5,7)),0)).getUTCDate(),dates=Array.from({length:count},(_,i)=>month+'-'+String(i+1).padStart(2,'0')),source=dailyRows.filter(r=>r[0].slice(0,7)===month),people=workerRows.filter(r=>r[0]===month),names=people.map(r=>r[1]);
    const headers=['작업자',...dates.flatMap(d=>[d,d]),'월 합계','월 합계','기록일 합계','근무비율','d','2P VI/d','3P VI/d','단계 VI/d'];
    const body=people.map(p=>{const perDay=dates.flatMap(d=>{const r=source.find(r=>r[0]===d&&laName_(r[1])===laName_(p[1]));return r?[r[2],r[3]]:['',''];});const record=workerDays.find(g=>g.month===month&&!g.external&&laName_(g.worker)===laName_(p[1]));return [p[1],...perDay,p[2],p[3],p[5],record?record.ratio:1,p[6],p[7],p[8],p[9]];});
    if(body.length){const d=sum(people,6),totalDates=dates.flatMap(day=>{const rs=source.filter(r=>r[0]===day);return rs.length?[sum(rs,2),sum(rs,3)]:['',''];});body.push(['TOTAL',...totalDates,sum(people,2),sum(people,3),sum(people,5),'',d,d?sum(people,2)/d:'',d?sum(people,3)/d:'',d?sum(people,4)/d:'']);}
    const formats={0:'@'};for(let i=1;i<headers.length;i++)formats[i]=i<=dates.length*2+3?'#,##0':'0.00';const s=section(daily,month+' | 작업자별 완료 VI 및 VI/d',headers,body,formats);s.subheaders=['',...dates.concat(['월']).flatMap(()=>['2P','3P']),'일','비율','d','2P','3P','단계 합계'];s.headerGroups=dates.concat(['월']).map((_,i)=>({offset:1+i*2,width:2}));
  });if(!daily.length)dayTable(daily,'대표 지표 | 일별 VI/d',[]);
  const ext=layouts['외부 지원']=[];dayTable(ext,'대표 지표 | 외부 작업자 월별 VI/d',metrics['외부 지원']||[]);dayTable(ext,'외부 작업자 주별 VI/d (월요일 시작)',metrics['외부 주별']||[]);dayTable(ext,'외부 작업자 일별 VI/d',metrics['외부 일별']||[],['날짜',...dayHeaders.slice(1)]);ext[ext.length-1].formats={...dayFormats,0:'yyyy-mm-dd'};
  const node=layouts['Node별']=[];dayTable(node,'대표 지표 | Node별 내부 VI/d',metrics['Node별']||[]);const nodes=tabs['Node별'][1],nodeMonths=Array.from(new Set(nodes.map(r=>r[0]))).sort().reverse();
  nodeMonths.forEach(month=>{const rows=nodes.filter(r=>r[0]===month).map(r=>{const g=(tabs._nodeCompletion||{})[JSON.stringify([month,r[1]])]||{current:0,prior:0,unknown:0},n=g.current+g.prior+g.unknown;return [r[1],r[2],r[3],typeof r[4]==='number'?r[4]/24:'',g.current,g.prior,g.unknown,n?g.prior/n:''];});if(rows.length){const samples=leadSamples.filter(s=>s.month===month);rows.push(['TOTAL',sum(rows,1),sum(rows,2),median(samples.map(s=>s.hours/24)),sum(rows,4),sum(rows,5),sum(rows,6),sum(rows,2)>0?sum(rows,5)/sum(rows,2):'']);}section(node,month+' | 요청 유입·완료 (전체 요청)',['Node','접수','완료','리드타임 중앙값(달력일)','당월 접수 완료','이전 월 접수 완료','접수일 미확인/오류','이월 완료율'],rows,{0:'@',1:'#,##0',2:'#,##0',4:'#,##0',5:'#,##0',6:'#,##0',7:'0.0%'});});
  const compare=layouts['내부·외부 비교']=[];dayTable(compare,'대표 지표 | 내부·외부 VI/d',metrics['내부·외부 비교']||[]);
  const time=layouts['작업시간']=[],recorded=tabs._recordedTime||{monthly:[],daily:[]};section(time,'근무자별 월 기록시간 (확인용)',['월','작업자','구분','크롤링(h)','3P(h)','2P(h)','합계(h)','시간 미기록 행'],recorded.monthly,{0:'@',1:'@',2:'@',3:'0.00',4:'0.00',5:'0.00',6:'0.00',7:'#,##0'});
  Array.from(new Set(recorded.monthly.map(r=>r[0]))).sort().reverse().forEach(month=>{const days=new Date(Date.UTC(Number(month.slice(0,4)),Number(month.slice(5,7)),0)).getUTCDate(),dates=Array.from({length:days},(_,i)=>month+'-'+String(i+1).padStart(2,'0')),people=recorded.monthly.filter(r=>r[0]===month),source=recorded.daily.filter(r=>r[0].slice(0,7)===month);const body=people.map(p=>[p[1],...dates.map(day=>{const r=source.find(r=>r[0]===day&&laName_(r[1])===laName_(p[1]));return r?r[6]:'';}),p[6]]),formats={0:'@'};for(let i=1;i<=dates.length+1;i++)formats[i]='0.00';section(time,month+' | 근무자별 일별 기록시간(h)',['작업자',...dates,'월 합계(h)'],body,formats);});

  const crawl=layouts['크롤링 효과']=[],ca=tabs._crawl;
  function crawlSection(title,headers,rows,formats){section(crawl,title,headers,rows,formats);crawl[crawl.length-1].noTotal=true;}
  crawlSection('01  해석 및 집계 기준',['항목','기준'],[
    ['분류','REQUESTS 현재 크롤링 단계만 사용: CRAWL_COMPLETED=크롤링 / N/A=비크롤링 / 나머지=확인 불가. 대소문자·앞뒤 공백 정규화.'],
    ['소급 적용','완료일·크롤링 필요·최초 요청 유형은 사용하지 않음. 현재 단계를 과거 실적에 소급하므로 당시 작업 방식의 확정 이력은 아님.'],
    ['분석 대상','CONFIG 외부 목록을 제외한 내부 작업자의 3P 세션과 COMMITTED 장기 집계. 2P와 CRAWL 준비시간 제외.'],
    ['대표 생산성','순증 VI 합계 ÷ sum(작업자별 고유 기록일×CONFIG 근무비율). 예: 0.5 근무자 10기록일=5d. 0시간 실적도 포함. 음수 정정 유지.'],
    ['배수','대표 배수는 크롤링 VI/환산인일 ÷ 비크롤링 VI/환산인일. 관찰된 차이이며 인과 효과 아님. 같은 사람이 같은 날 두 유형을 작업하면 양쪽 분모에 포함되므로 합산 금지.'],
    ['비중','전체 비중은 확인 불가 물량까지 분모에 포함. 분류 확인 내 비중은 크롤링+비크롤링만 분모. 각 그룹 순증 VI가 음수이거나 분모 0이면 비중 공란.'],
    ['추이','완료/일시정지 세션의 종료일(KST) 귀속. ACTIVE·VI 누락 기록은 분자 제외·기록일 포함. 일별은 기록 있는 날짜만 표시. 최근 7일은 7개 달력일 합산 VI/환산인일.'],
    ['비교 한계','같은 작업자·Node도 브랜드, 난이도, 시기 차이는 남음. 그룹 간 동일 요청 중복 집계 가능. 원본 요청이 없으면 확인 불가.'],
    ['크롤링 개인 평균','전체 보유 이력 기준: CRAWL_COMPLETED 요청의 현재 Total VI를 현재 크롤링 작업자에게 1회 귀속. CRAWL 세션 종료일(없으면 시작일)의 작업자별 고유 일수로 나눔. 0시간·ACTIVE 기록도 날짜가 있으면 포함.'],
    ['개인 평균 한계','진행 중 요청의 기록일도 분모에 포함. 여러 요청 동시 작업은 하루로 중복 제거. 월별 물량은 요청-작업자 마지막 CRAWL 기록월로 추정 귀속. 완료일 미사용. 현재 Total VI는 제출 당시와 다를 수 있음.'],
    ['백업 누락 주의','현재 장기 집계는 CRAWL을 보존하지 않음. 요청-작업자 연결 기록이 없거나 물량 오류가 있으면 전체 평균 공란, 기록 연결 물량의 참고 평균 표시. 연결률 100%도 일부 작업일 누락까지 보증하지 않음.'],
    ['확인 방법','최종 CONFIG 성공 시각과 A008 버전 확인. 비교·순위는 VI/d, h는 작업시간 탭에서만 확인.']
  ],{0:'@',1:'@'});crawl[0].noteBlock=true;
  if(ca){
    crawlSection('02  월별 전환 비중 및 등록 생산성',['월','전체 순증 VI','크롤링 VI','비크롤링 VI','확인 불가 VI','크롤링 비중(전체)','크롤링 비중(분류 확인 내)','분류 확인 비중','크롤링 등록 VI/h','비크롤링 등록 VI/h','등록 생산성 배수'],ca.trends,{0:'@',1:'#,##0',2:'#,##0',3:'#,##0',4:'#,##0',5:'0.0%',6:'0.0%',7:'0.0%',10:'0.00"배"'});
    crawlSection('03  월별 비교 표본 및 0시간 실적',['월','그룹','순증 VI','작업시간(h)','시간 유효 순증 VI','등록 VI/h','작업 요청 수','기록 행 수','0시간 순증 VI'],ca.monthly,{0:'@',1:'@',2:'#,##0',4:'#,##0',6:'#,##0',7:'#,##0',8:'#,##0'});
    crawlSection('04  일별 및 최근 7일 추이',['날짜','크롤링 VI','비크롤링 VI','확인 불가 VI','크롤링 비중(전체)','크롤링 비중(분류 확인 내)','분류 확인 비중','크롤링 등록 VI/h','비크롤링 등록 VI/h','등록 생산성 배수','전체 순증 VI','최근 7일 크롤링 비중(전체)','최근 7일 크롤링 VI/h','최근 7일 비크롤링 VI/h','최근 7일 생산성 배수'],ca.daily,{0:'yyyy-mm-dd',1:'#,##0',2:'#,##0',3:'#,##0',4:'0.0%',5:'0.0%',6:'0.0%',9:'0.00"배"',10:'#,##0',11:'0.0%',14:'0.00"배"'});
    const comparisonHeaders=['월','작업자','크롤링 시간 유효 VI','크롤링 시간(h)','크롤링 등록 VI/h','비크롤링 시간 유효 VI','비크롤링 시간(h)','비크롤링 등록 VI/h','등록 생산성 배수','크롤링 작업 요청 수','비크롤링 작업 요청 수','확인 불가 VI'];
    const comparisonFormats={0:'@',1:'@',2:'#,##0',5:'#,##0',8:'0.00"배"',9:'#,##0',10:'#,##0',11:'#,##0'};
    crawlSection('05  같은 작업자 월별 비교 (양쪽 표본 확인)',comparisonHeaders,ca.workers,comparisonFormats);
    const nodeHeaders=comparisonHeaders.slice();nodeHeaders[1]='Node';
    crawlSection('06  같은 Node 월별 비교',nodeHeaders,ca.nodes,comparisonFormats);
    crawlSection('07  확인 불가 기록 점검',['월','사유','순증 VI','작업시간(h)','작업 요청 수','기록 행 수'],ca.reasons,{0:'@',1:'@',2:'#,##0',4:'#,##0',5:'#,##0'});
    if(ca.preparation){
      crawlSection('08  크롤링 작업자별 1일 평균 VI (전체 보유 기록)',['크롤링 작업자','완료 요청 총 VI','완료 요청 수','기록일 수','총 VI/기록일','기록 연결 VI','기록 연결 요청 수','요청 기록 연결률','연결 VI/기록일(참고)','첫 기록일','마지막 기록일'],ca.preparation.rows,{0:'@',1:'#,##0',2:'#,##0',3:'#,##0',4:'0.0',5:'#,##0',6:'#,##0',7:'0.0%',8:'0.0',9:'@',10:'@'});
      crawlSection('09  크롤링 개인 평균 누락 점검',['요청 ID','크롤링 작업자','사유','현재 Total VI'],ca.preparation.audit,{0:'@',1:'@',2:'@',3:'#,##0'});
    }
    if(ca.monthlyPreparation)crawlSection('10  월별 크롤링 작업자 물량 및 연결된 3P 등록 현황 (귀속월 추정)',['월','크롤링 작업자','크롤링 VI','완료 요청 수','CRAWL 기록일','1기록일 평균 VI','동일 요청 현재 3P 완료 VI','동일 요청 현재 등록률','물량/등록수 오류 요청 수','환산 기록 인일','크롤링 VI/환산인일'],ca.monthlyPreparation.rows,{0:'@',1:'@',2:'#,##0',3:'#,##0',4:'#,##0',5:'0.0',6:'#,##0',7:'0.0%',8:'#,##0',9:'0.00',10:'0.0'});
  }
  const plan=tabs._plan;if(plan){
    const list=layouts['투입 계획']=[];
    function planSection(title,headers,rows,formats){section(list,title,headers,rows,formats);list[list.length-1].noTotal=true;}
    planSection('01  계산 전제 및 사용법',['항목','설명'],[
      ['입력','투입 계획 설정 탭에서 기준 월·크롤링/등록 환산 인원·목표 작업일을 입력한 뒤 새로고침. 예: 0.5 근무자 2명은 환산 1명. 표본 부족·빈 입력은 공란.'],
      ['실측과 추정','시간은 사용하지 않음. CRAWL 및 3P의 VI/환산인일로 표준 물량 1,000 VI의 필요 인일을 각각 계산해 합산. 실측 소요시간이 아닌 분리 공정 시나리오.'],
      ['월별 귀속','크롤링 완료일 미사용. 완료 요청의 현재 Total VI 전체를 해당 작업자의 마지막 CRAWL 기록월에 1회 귀속. 여러 달 작업한 요청의 물량 배분은 복원하지 않음.'],
      ['3P 연결','월별 크롤링 표의 3P 완료 VI는 같은 요청들의 현재 등록 누계로 전기간·전체 작업자를 포함. 크롤링월의 등록 실적이 아님. 3P 기준 속도는 별도로 기준 월 내부 작업자의 3P 유효 기록일과 CONFIG 근무비율을 사용.'],
      ['기준 속도','크롤링 평균은 기준 월 전체 물량/작업자별 환산 기록일 합계. 개인 평균의 단순 평균 아님. 등록도 기준 월 3P 기록일과 CONFIG 근무비율 사용. 외부 지원은 별도 표로 분리.'],
      ['잔량 범위','DONE/DROP/취소 및 2P 전용 요청 제외. 현재 Total VI-3P Completed VI. ON HOLD는 별도 잔량으로 표시하고 재개 전 예측에서 제외. 크롤링 대기 미확정 VI는 입력값 또는 기준 월 평균 요청 크기로 추정. 분류/물량 확인 필요 요청은 예측에서 제외.'],
      ['기간과 인원','현재 잔량만, 신규 유입·휴일·승인대기·재작업·요청 분할 제약 제외. 크롤링/등록 인력은 별도 전담 가정. 같은 사람·날짜의 기록일은 두 공정에 겹칠 수 있으므로 합산 인일은 실측 노동량이 아님. 병행일은 용량상 이상적 하한이며 순차일도 납기 보장이 아님. 목표 인원 역시 단계별 용량 하한.'],
      ['효율 표현','1,000 VI당 비크롤링 필요 인일과 크롤링 준비+등록 필요 인일 비교. 같은 물량의 독립 공정 계획값이며 실제 절감 노동량이나 인과 효과를 확정하지 않음.'],
      ['불완전 표본','이전 백업으로 CRAWL이 빠지면 월별 물량이 시점 미확인으로 남음. 설정의 기준 속도를 보정 가능. 세션이 찍힌 날만 분모이므로 일을 하지 않은 근무일은 포함되지 않음.'],
      ['홀드 보정','월별·Node별 대표 처리기간은 전체 경과시간-Blocked Time (Hours). 완료 요청만 사용. 홀드 누락·음수·전체시간 초과는 보정 제외. 달력일이며 작업시간·SLA 위반일수가 아님.'],
      ['동일 비교 표본','홀드 보정 표의 전체/홀드/보정 중앙값은 모두 홀드값이 유효한 동일 요청만 사용. 누락 요청 수를 별도 표시. 모든 홀드가 외부 사유인지는 현재 기록으로 구분하지 못함.']
    ],{0:'@',1:'@'});list[0].noteBlock=true;
    planSection('02  적용 기준 및 입력 가정',['항목','값'],plan.assumptions,{0:'@'});
    planSection('03  현재 잔량과 미확정 물량',['상태','요청 수','확정 잔량 VI','미확정 요청 수'],plan.backlog,{0:'@',1:'#,##0',2:'#,##0',3:'#,##0'});
    planSection('04  1,000 VI당 필요 인일·효율 및 병목',['지표','값'],plan.impact,{0:'@'});
    planSection('05  잔량 소화에 필요한 환산 인원·인일·작업일 (추정)',['지표','값'],plan.forecast,{0:'@'});
  }
  if(ca){
    const core=[];
    function daySection(title,headers,rows,formats){section(core,title,headers,rows,formats);core[core.length-1].noTotal=true;}
    daySection('월별 리스팅 물량의 크롤링 비중',['월','크롤링 VI','비크롤링 VI','확인 불가 VI','전체 VI','크롤링 비중'],ca.dayTrends.map(r=>{const total=r[1]+r[4]+r[8];return [r[0],r[1],r[4],r[8],total,total>0&&r[1]>=0&&r[4]>=0&&r[8]>=0?r[1]/total:''];}),{0:'@',1:'#,##0',2:'#,##0',3:'#,##0',4:'#,##0',5:'0.0%'});
    daySection('대표 지표 | 월별 크롤링·비크롤링 등록 효율',['월','크롤링 VI','크롤링 환산인일','크롤링 VI/환산인일','비크롤링 VI','비크롤링 환산인일','비크롤링 VI/환산인일','등록 효율 배수','확인 불가 VI'],ca.dayTrends,{0:'@',1:'#,##0',4:'#,##0',7:'0.00"배"',8:'#,##0'});
    daySection('대표 지표 | 월별 표본과 환산 기록일',['월','분류','VI','기록 인일','환산인일','VI/환산인일','작업 요청 수'],ca.dayMonthly,{0:'@',1:'@',2:'#,##0',3:'#,##0',6:'#,##0'});
    daySection('대표 지표 | 일별·최근 7일 등록 효율',['날짜','크롤링 VI','비크롤링 VI','확인 불가 VI','크롤링 VI/환산인일','비크롤링 VI/환산인일','일별 효율 배수','7일 크롤링 VI/환산인일','7일 비크롤링 VI/환산인일','7일 효율 배수'],ca.dayDaily,{0:'yyyy-mm-dd',1:'#,##0',2:'#,##0',3:'#,##0',6:'0.00"배"',9:'0.00"배"'});
    const headers=['월','작업자','크롤링 VI','크롤링 환산인일','크롤링 VI/환산인일','비크롤링 VI','비크롤링 환산인일','비크롤링 VI/환산인일','효율 배수','확인 불가 VI'];
    daySection('대표 지표 | 같은 작업자 비교',headers,ca.dayWorkers,{0:'@',1:'@',2:'#,##0',5:'#,##0',8:'0.00"배"',9:'#,##0'});
    const nh=headers.slice();nh[1]='Node';daySection('대표 지표 | 같은 Node 비교',nh,ca.dayNodes,{0:'@',1:'@',2:'#,##0',5:'#,##0',8:'0.00"배"',9:'#,##0'});
    crawl.forEach(s=>{if(s.headers.some(h=>String(h).includes('VI/h'))||s.title.startsWith('08  '))s.reference=true;});crawl.splice(1,0,...core);
    const monthlyIndex=crawl.findIndex(s=>s.title.startsWith('10  '));if(monthlyIndex>=0)crawl.splice(1,0,crawl.splice(monthlyIndex,1)[0]);
  }
  // Compare elapsed lead time on the SAME requests with valid recorded hold hours.
  function holdRows(byNode){
    const groups=new Map();leadSamples.forEach(s=>{const key=JSON.stringify(byNode?[s.month,s.node]:[s.month]);if(!groups.has(key))groups.set(key,[]);groups.get(key).push(s);});
    const median=values=>{const a=values.slice().sort((a,b)=>a-b),n=a.length;return n?(a[Math.floor((n-1)/2)]+a[Math.floor(n/2)])/2:'';};
    return Array.from(groups.keys()).sort().reverse().map(key=>{const all=groups.get(key),valid=all.filter(s=>s.adjusted!==null),full=valid.reduce((n,s)=>n+s.hours,0),blocked=valid.reduce((n,s)=>n+s.blocked,0);return JSON.parse(key).concat([all.length,valid.length,all.length-valid.length,median(valid.map(s=>s.hours/24)),median(valid.map(s=>s.blocked/24)),median(valid.map(s=>s.adjusted/24)),full>0?blocked/full:'']);});
  }
  const holdHeaders=['월','완료 요청 수','홀드 보정 유효 요청 수','홀드 기록 확인 필요','전체 경과일 중앙값','홀드일 중앙값','홀드 제외 경과일 중앙값','홀드 시간 비중'];
  const hold=[];section(hold,'대표 지표 | 기록된 홀드를 제외한 처리기간 (달력일)',holdHeaders,holdRows(false),{0:'@',1:'#,##0',2:'#,##0',3:'#,##0',7:'0.0%'});hold[0].noTotal=true;layouts['월별 비교'].splice(1,0,...hold);
  const nodeHold=[];section(nodeHold,'대표 지표 | Node별 홀드 제외 처리기간 (달력일)',['월','Node'].concat(holdHeaders.slice(1)),holdRows(true),{0:'@',1:'@',2:'#,##0',3:'#,##0',4:'#,##0',8:'0.0%'});nodeHold[0].noTotal=true;layouts['Node별'].splice(1,0,...nodeHold);
  // Crawl effectiveness is exclusively day-based; remove legacy hourly tables.
  layouts['크롤링 효과']=layouts['크롤링 효과'].filter(s=>!s.headers.some(h=>String(h).includes('VI/h'))&&!s.title.startsWith('08  '));
  const audit=layouts['크롤링 효과'].find(s=>s.title.startsWith('07  '));
  if(audit){audit.headers.splice(3,1);audit.rows.forEach(r=>r.splice(3,1));audit.formats={0:'@',1:'@',2:'#,##0',3:'#,##0',4:'#,##0'};}
  function terminology(v){return typeof v==='string'?v.replace(/VI\/환산인일/g,'VI/d').replace(/환산 근무일당 완료 VI/g,'VI/d').replace(/환산 근무 인일|환산 기록 인일|환산인일/g,'d').replace(/작업기록 인일|기록 인일/g,'기록일 합계').replace(/환산 인원/g,'보정 인원').replace(/등록/g,'리스팅'):v;}
  Object.keys(layouts).forEach(n=>{let next=6,last=null;layouts[n].forEach(s=>{s.title=terminology(s.title);s.headers=s.headers.map(terminology);if(s.subheaders)s.subheaders=s.subheaders.map(terminology);s.rows=s.rows.map(r=>r.map(terminology));s.row=s.alongside&&last?last.row:next;const end=s.row+s.rows.length+(s.subheaders?3:2);next=Math.max(next,end+3);last=s;});});return layouts;
}
function laClearReportMerges_(sheet){
  // A merged title can extend beyond getDataRange()'s last populated column.
  // Unmerge each complete merged range, including otherwise empty title cells.
  sheet.getRange(1,1,sheet.getMaxRows(),sheet.getMaxColumns()).getMergedRanges().forEach(range=>range.breakApart());
}
function laReportValue_(value,header){
  if(header==='날짜'){
    const day=typeof value==='number'?laDay_(value):value;
    if(typeof day==='string'&&/^\d{4}-\d{2}-\d{2}$/.test(day))return Date.parse(day+'T00:00:00Z')/86400000+25569;
  }
  // Period labels are text. Sheets may otherwise coerce them into date serials.
  if(typeof value==='string'&&/^\d{4}-\d{2}(?:-\d{2})?$/.test(value))return "'"+value;
  return typeof value==='string'&&/^[=+@-]/.test(value)?"'"+value:value;
}
function laRenderReports_(ss,layouts,updated){
  Object.keys(layouts).forEach(name=>{
    const sections=layouts[name],height=Math.max(10,...sections.map(s=>s.row+s.rows.length+5)),width=Math.max(13,...sections.map(s=>s.col+s.headers.length));
    let sheet=ss.getSheetByName(name);if(!sheet)sheet=ss.insertSheet(name);
    if(sheet.getMaxRows()<height)sheet.insertRowsAfter(sheet.getMaxRows(),height-sheet.getMaxRows());if(sheet.getMaxColumns()<width)sheet.insertColumnsAfter(sheet.getMaxColumns(),width-sheet.getMaxColumns());
    laClearReportMerges_(sheet);sheet.clear();sheet.setHiddenGridlines(true);sheet.setFrozenRows(3);sheet.setFrozenColumns(1);sheet.setTabColor(name==='외부 지원'?'#b7791f':'#304e76');
    sheet.getRange(1,1,height,width).setBackground('#f4f7fb').setFontFamily('Arial').setFontSize(10).setFontColor('#24354b').setVerticalAlignment('middle');sheet.setColumnWidth(1,24);sheet.setColumnWidths(2,width-1,112);sheet.setColumnWidth(2,170);sheet.setRowHeights(1,height,24);
    sheet.getRange(1,2,1,width-2).merge().setValue('LISTING  /  '+name).setFontSize(18).setFontWeight('bold').setFontColor('#183555');sheet.setRowHeight(1,38);
    sheet.getRange(2,2,1,width-2).merge().setValue('최근 갱신 '+updated+' KST  |  A008  |  수치는 원본 가용 이력 기준').setFontColor('#62758b');
    const note=name==='작업시간'?'기록된 h 확인용 / 비교·순위에는 사용하지 않음 / ACTIVE 제외':name==='투입 계획'?'설정 탭의 환산 인원·기준 속도를 확인하세요 / 현재 잔량의 3P 등록까지 예측 / 2P 완료 납기 아님':name==='크롤링 효과'?'내부 3P 등록 생산성 / 現 크롤링 단계 소급 분류 / 준비시간 제외 / 인과 효과 아님':name==='외부 지원'?'CONFIG 지정 외부 지원만 표시 · 내부 전체 실적 및 내부 순위에서 제외':name==='내부·외부 비교'?'2P·3P 별도 비교 / 1인·1일 평균 = 총 VI ÷ 작업자별 근무일수 합계 / 보정 = 총 VI ÷ 환산 근무일수 합계':name==='Node별'?'이월 완료율 = 이전 월 접수 완료 ÷ 전체 당월 완료 (납기 위반율 아님) / 접수일 미확인·오류는 별도 표시':'내부 실적 · 외부 지원 제외 / 진행 중인 기간은 부분 실적 / VI/d = 완료 VI ÷ 보정 기록일 d / h는 작업시간 탭에서 확인';
    sheet.getRange(3,2,1,width-2).merge().setValue(note).setFontSize(9).setFontColor('#62758b');
    sections.forEach(s=>{
      const w=s.headers.length,r=s.row,c=s.col,bodyRow=r+(s.subheaders?3:2);
      sheet.getRange(r,c,1,w).merge().setValue((s.reference?'참고 | ':'')+s.title).setFontWeight('bold').setFontSize(11).setFontColor('#183555');sheet.setRowHeight(r,30);
      sheet.getRange(r+1,c,1,w).setNumberFormat('@').setValues([s.headers.map(v=>laReportValue_(v,''))]).setBackground('#304e76').setFontColor('#ffffff').setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');sheet.setRowHeight(r+1,36);
      if(s.subheaders){sheet.getRange(r+2,c,1,w).setNumberFormat('@').setValues([s.subheaders.map(v=>laReportValue_(v,''))]).setBackground('#d8e5f4').setFontWeight('bold').setHorizontalAlignment('center');sheet.getRange(r+1,c,2,1).merge();const groups=s.headerGroups||Array.from({length:(w-1)/2},(_,i)=>({offset:1+i*2,width:2}));groups.forEach(g=>sheet.getRange(r+1,c+g.offset,1,g.width).merge());sheet.setColumnWidths(c+1,w-1,76);}
      const body=sheet.getRange(bodyRow,c,s.rows.length,w);body.setNumberFormat('#,##0.00');
      // Set temporal column formats BEFORE writing numeric date values.
      s.headers.forEach((header,i)=>{if(header==='날짜')sheet.getRange(bodyRow,c+i,s.rows.length,1).setNumberFormat('yyyy-mm-dd');else if(header==='월')sheet.getRange(bodyRow,c+i,s.rows.length,1).setNumberFormat('yyyy-mm');});
      body.setValues(s.rows.map(row=>row.map((v,i)=>laReportValue_(v,s.headers[i])))).setHorizontalAlignment('right');
      body.setBackgrounds(s.rows.map((_,i)=>Array(w).fill(i%2===0?'#ffffff':'#eaf0f8')));sheet.getRange(bodyRow,c,s.rows.length,1).setHorizontalAlignment('left');
      s.rows.forEach((row,i)=>{if(row[0]==='TOTAL')sheet.getRange(bodyRow+i,c,1,w).setBackground('#d8e5f4').setFontWeight('bold');});
      Object.keys(s.formats).forEach(i=>sheet.getRange(bodyRow,c+Number(i),s.rows.length,1).setNumberFormat(s.formats[i]));
      if(s.rowFormats)s.rowFormats.forEach((fmt,i)=>{sheet.getRange(bodyRow+i,c+1,1,w-1).setNumberFormat(fmt);if(String(s.rows[i][0]).indexOf('합계')>=0)sheet.getRange(bodyRow+i,c,1,w).setBackground('#d8e5f4').setFontWeight('bold');});
      sheet.getRange(r+1,c,s.rows.length+(s.subheaders?2:1),w).setBorder(true,true,true,true,true,true,'#d6dfeb',SpreadsheetApp.BorderStyle.SOLID);body.setWrap(true);
      if(s.noteBlock)s.rows.forEach((_,i)=>{sheet.getRange(bodyRow+i,c+1,1,width-c).merge().setHorizontalAlignment('left').setWrap(true);sheet.setRowHeight(bodyRow+i,48);});
    });
    if(name==='데이터 점검'){sheet.setColumnWidth(3,180);sheet.setColumnWidth(4,650);sheet.getRange(6,2,height-5,3).setWrap(true);}
  });
}


