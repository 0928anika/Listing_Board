일반 분석 최신 A009: Code.gs + Layout.gs 교체 후 refreshListingAnalytics 실행.
2P·3P 분리 / 크롤링+리스팅 전체 효율 / 잔량 및 요청당 필요 d / 작업자 순위 탭.
투입 계획 설정은 필요 없습니다. 상세: A009-UPDATE.md.

일반 분석 최신 A008: Code.gs + Layout.gs 교체 후 refreshListingAnalytics 실행.
비교·순위 VI/d 통일, h는 작업시간 탭에서 확인. CONFIG 근무비율 적용.
상세: A008-UPDATE.md. 일회용 11월 보고는 변경 없음.

11월 보고 최신: 전체 REQUESTS 평균 VI 고정, 비크롤링 1/5개·크롤링 1/5/10개 필요 d. 크롤링 방식은 크롤링+리스팅 d 합산. 주차 표는 완료 요청이 아닌 작업 기록 기준으로 명시.
LeadershipReport.gs만 전체 교체 후 buildNovemberReport 실행.

11월 보고 최신: 월~일 주차별 요청 비중·VI·VI/d, 양쪽 2기록일 C/N TOP5, 전체 합계 C VI/d·N VI/d, 통합 VI/d, 비크롤링 기준 요청당 d.
LeadershipReport.gs만 교체 후 buildNovemberReport 실행. 상세: NOVEMBER-REPORT.md

11월 보고: 월 전체 집계 차단 수정, 소표본 구분 제거, 중앙값 두 지표로 단순화, 크롤링 VI%/요청% 분리.
LeadershipReport.gs만 교체 후 buildNovemberReport 실행.

11월 보고 최신: 요청 1/5/10개당 d, 요청당 1일 추정 크롤링 d, TOP 5 아래 전체 통계, 11월 일별 심층.
LeadershipReport.gs만 교체 후 buildNovemberReport 실행. 자세한 기준은 NOVEMBER-REPORT.md.

일회용 11월 보고 최신 업데이트: NOVEMBER-REPORT.md
LeadershipReport.gs만 전체 교체 후 buildNovemberReport 실행.
기존 자동 생성 탭을 11월 보고로 변경합니다. 일반 새로고침과 분리됩니다.

Listing Analytics A007

최신 안내: A007-UPDATE.md
일반 분석: Code.gs + Layout.gs 교체 후 refreshListingAnalytics 실행.
일회용 리더 보고: LeadershipReport.gs 추가 후 buildCrawlerHubLeadershipReport 별도 실행.
리더 보고서는 일반 새로고침과 트리거에서 변경하지 않습니다. 개인별 비교 없이 월별 팀 성과만 표시합니다.

이하 이전 버전 안내:

Listing Analytics A006

최신 업데이트·계산 기준: A006-UPDATE.md
Code.gs와 Layout.gs를 전체 교체 후 refreshListingAnalytics 실행.
대표 지표: 환산 근무일당 완료 VI. 투입 계획 설정 탭에서 인원·목표 작업일을 입력하세요.

이하 이전 버전 안내:

Listing Analytics A005

기존 분석 파일 업데이트: A005-UPDATE.md를 먼저 확인하세요.
Code.gs와 Layout.gs를 전체 교체한 후 refreshListingAnalytics 실행.
운영 코드/CONFIG/매니페스트/기존 트리거는 변경하지 않습니다.

아래는 기존 A004 설정 안내입니다.

Listing Analytics A004 — 2026-10-30

既存 A003/H01~H04 업데이트
1. 분석 파일(운영 파일 아님)의 Apps Script에서 Code.gs와 Layout.gs를 각각 전체 교체합니다.
2. 저장 후 refreshListingAnalytics 실행. 매니페스트/기존 CONFIG A:C/트리거 유지.
3. CONFIG G1:I에 작업자/근무비율/설명이 자동 준비됩니다.
   G2부터 작업자 이름, H열 근무비율 입력. 예: Amy | 0.5. 정상 1.
   근무비율은 0 초과 1 이하. 미입력 기본 1. 중복 이름/잘못된 수치는 갱신 중단.
   G:I에 이미 다른 내용이 있으면 코드가 덮어쓰지 않고 중단하므로 먼저 빈 영역으로 옮기세요.
4. 비율을 입력/변경한 후 refreshListingAnalytics 재실행.
   근무비율과 외부 목록은 과거 전체 기간에 소급 적용됩니다. 기간별 비율 이력 기능은 없습니다.
5. 새 내부·외부 비교 탭과 CONFIG 성공 시각을 확인하세요.
초기 설정/트리거 재설치/운영 코드 교체가 필요하지 않습니다.

신규 설치
빈 Google 스프레드시트 > 확장 프로그램 > Apps Script.
Code.gs와 Layout.gs 전체 붙이기. 매니페스트 표시 후 appsscript.json도 동봉 내용 적용.
setupListingAnalytics 1회 실행/권한 승인. CONFIG B2 운영 파일 ID/URL 입력.
A열 EXTERNAL_WORKER, B열 이름으로 외부 목록 입력/행 추가. G:H에 근무비율 입력.
refreshListingAnalytics 실행. 필요 시 installListingAnalyticsDaily 1회 실행(새 분석 프로젝트에만).
표준 Cloud 프로젝트 사용 시 Google Sheets API 활성화 필요. 운영 파일 조회 권한 필요.

일별 기록 추가 열
기존 월별 날짜별 2P/3P VI 및 월합계 유지. 그 오른쪽에:
근무일수 | 근무비율 | 1일평균(2P/3P/TOTAL) | 보정1일평균(2P/3P/TOTAL).
근무일수 = 해당 작업자의 해당 월 중 2P+3P 기록 작업시간 합계 > 0인 날짜 수.
동일 날짜의 여러 세션/요청/단계는 하루로 계산. CRAWL 준비시간과 ACTIVE는 기존처럼 제외.
이 값은 시트에 기록된 업무일 수이며 출퇴근 기록 기반은 아닙니다.
일평균 = 월 VI / 근무일수. 보정 = 월 VI / (근무일수 × 근무비율).
예: 1000 VI / 10일 = 100 VI/일, 비율 0.5이면 보정 200 VI/일.
일수 0이면 평균 공란. 음수 수정 유지. 작업시간 없는 실적은 기존 규칙대로 VI에 남습니다.
TOTAL 행 근무일수는 인일 합계이며 비율은 여러 사람을 단순 합산하지 않고 공란.
TOTAL 일반 평균 = 총 VI / 총 인일. 보정 평균 = 총 VI / 환산 인일 합계.
작업시간표/작업자 월별 VI/h 순위는 변경하지 않습니다.

내부·외부 비교
월마다 내부와 외부 행을 나란히 비교. 2P와 3P를 별도 열로 유지하고 TOTAL 보조 열 제공.
작업자 수/근무일수 합계/환산 근무일수/2P·3P VI/단계별 일평균/보정평균 제공.
그룹 일평균 = 그룹 총 VI / 작업자별 근무일수 합계(인일).
그룹 보정 = 그룹 총 VI / sum(각 작업자의 일수 × 근무비율).
작업자별 평균을 단순 평균하지 않음. 근무일의 분모는 2P/3P 공통으로 사용.
외부 전담/내부 전담 단계 차이가 있으므로 VI/h 단계별 순위와 다른 지표입니다.

Node별 요청 완료 세분화
완료일의 달 기준으로 당월 접수 완료 / 이전 월 접수 완료 / 접수일 미확인·오류를 표시.
이월 완료율(딜레이율) = 이전 월 접수 완료 수 / 전체 당월 DONE 완료 수.
미확인/오류도 분모 전체 완료에는 포함하고 별도 건수를 표시.
월을 넘겨 완료했다는 지표이며 SLA/납기 위반이나 지연 시간 자체를 뜻하지 않습니다.
TOTAL도 전체 건수로 비율 재계산. Node는 현재 값, 외부 포함 전체 요청 기준 유지.

기존 기준
원본 REQUESTS + WORK_SESSIONS + COMMITTED LONG_TERM_SESSIONS를 조회만 합니다.
상세 아카이브 직접 합산 없음. 기존 중복/중첩 방지, 미래 세션 제외, 종료일 귀속 유지.
운영 파일 날짜는 KST 날짜 시리얼 전제. 문자열 날짜 추측 변환 없음.
주간 월요일 시작/주 시작 월로 표 분리. 일별 기록 월별 표/2P·3P 하위 열 유지.
외부 지원 순위/작업자 주별/데이터 점검 탭 없음. 리드타임은 전체 완료 요청 기준 유지.
출력 결과는 매번 재계산되므로 분석 결과 탭 직접 편집 내용은 대체됩니다.
출력 도중 서비스 오류가 나면 일부 탭만 갱신될 수 있어 CONFIG 성공 시각 확인 후 활용하세요.
별도 외부 서버 없음. 코드에서 source ID는 Sheets API GET 조회에만 사용합니다.
Google 파일 > 파일 > 다운로드 > Excel로 숫자 재사용 가능.

로컬 합성 집계/배치/근무일/비율/2P·3P 비교/이월 완료율 검사 68개 통과.
실제 회사 시트 권한/화면/성능/이력 완전성은 미검증. 운영 실행/배포 없음. B008 변경 없음.
