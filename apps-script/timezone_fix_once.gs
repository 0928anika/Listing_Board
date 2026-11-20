// 운영자가 추가하여 이미 1회 실행한 함수. 기준본 보관용이며 재실행 지시가 아닙니다.
function FIX_TIMEZONE_ONCE() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Seoul');
  ss.setSpreadsheetLocale('ko_KR');
  SpreadsheetApp.flush();
  console.log('Timezone: ' + ss.getSpreadsheetTimeZone());
  console.log('Locale: ' + ss.getSpreadsheetLocale());
}
