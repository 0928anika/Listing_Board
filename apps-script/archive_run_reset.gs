// Manually verified archived request only. Does not change source/archive rows.
function resetArchiveRunOnce() {
  assertDashboardAdmin_();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) {
    throw new Error('백업 실행이 끝난 뒤 다시 실행하세요.');
  }
  try {
    const run = archiveResumeReadV2_(ARCHIVE_RESUME_V2.RUN);
    if (run && run.job) {
      const table = archiveResumeSourceV2_(APP_CONFIG.SHEETS.REQUESTS, run.job.id);
      const column = table.map['Archive Status'];
      if (table.rows.length !== 1 || column === undefined) {
        throw new Error('현재 요청의 백업 완료 상태를 확인할 수 없습니다.');
      }
      const status = table.sheet.getRange(table.rows[0], column + 1).getValue();
      if (String(status).trim().toUpperCase() !== 'ARCHIVED') {
        throw new Error('현재 요청이 미완료라 초기화하지 않았습니다.');
      }
    }
    archiveResumePropsV2_().deleteProperty(ARCHIVE_RESUME_V2.RUN);
    console.log('백업 실행 상태만 초기화했습니다.');
  } finally {
    lock.releaseLock();
  }
}
