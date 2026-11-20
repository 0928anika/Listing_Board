// 시스템 전체에서 사용하는 시트명, 컬럼명, 상태값을 정의함.
const APP_CONFIG = Object.freeze({
  SHEETS: Object.freeze({
    REQUESTS: 'REQUESTS',
    TASKS: '3P_TASKS',
    CRAWL_TASKS: 'CRAWL_3P_TASKS',
    TASKS_2P: '2P_TASKS',
    ACTIVITY_LOG: 'ACTIVITY_LOG',
    CONFIG: 'CONFIG',
  }),

  STATUS: Object.freeze({
    TO_DO: 'TO DO',
    ONGOING: 'ONGOING',
    ON_HOLD: 'ON HOLD',
    DONE: 'DONE',
    DROP: 'DROP',
  }),

  PRIORITY: Object.freeze(['HIGH', 'NORMAL', 'LOW']),
  REQUEST_ID_PREFIX: 'REQ',
  RESERVED_DATA_ROWS: 1000,

  REQUEST_COLUMNS: Object.freeze([
    'Request ID', 'Node', 'Brand', 'Requestor', 'Request Type',
    'Priority', 'Active Assignees', 'Status', 'Requestor Notes', 'Total VI',
    '3P Completed VI', '2P Completed VI',
    'Started At', 'Blocked At', 'Resumed At', 'Completed At',
    'Blocked Time (Hours)', 'Lead Time (Hours)', 'Work Time (Hours)',
    // 계산 및 이력용 시스템 컬럼임. REQUESTS 화면에서는 숨김.
    'Requested At', 'Progress',
    // 바코드 생성은 작업단계가 아니라 2P 생산성 분석용 요청 속성임.
    'Barcode Provided',
  ]),

  TASK_COLUMNS: Object.freeze([
    'Request ID', '품번', '등록상품명', '사이즈', '색상', '판매가', '정상가', 'Brand',
    'Memo', '3P Assignee', '3P Completed', '2P Assignee', '2P Completed',
    'Work Status', 'Internal Task ID',
  ]),

  ACTIVITY_COLUMNS: Object.freeze([
    'Timestamp', 'User', 'Request ID', 'Action', 'Before', 'After',
  ]),
});
