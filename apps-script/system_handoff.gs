// 시스템 인수인계 전용. 권한 저장소와 백업 트리거를 함께 전환하는 교차 기능만 둠.
function transferDashboardPrimaryAdmin(email) {
  const current = dashboardRequireCurrentEmail_(), target = dashboardValidateEmail_(email);
  return withOpsDocumentLock_(function() {
    const primary = dashboardPrimaryAdminEmail_();
    if (current !== primary) throw new Error('현재 기본 관리자만 기본 권한을 이전할 수 있습니다.');
    if (target === primary) throw new Error('이미 기본 관리자인 계정입니다.');
    const props = dashboardAccessProperties_(), users = dashboardReadUsers_(), admins = dashboardReadAdmins_();
    if (!users[target] || users[target].status !== 'GRANTED' || !admins[target]) {
      throw new Error('대상 계정을 먼저 사용 중인 관리자로 승인해 주세요.');
    }

    const triggerOwner = dashboardNormalizeEmail_(props.getProperty(ARCHIVE_CONFIG_V1.TRIGGER_OWNER_KEY));
    if (triggerOwner && triggerOwner !== current) {
      throw new Error('백업 트리거 소유자가 현재 기본 관리자와 다릅니다: '+triggerOwner+' 계정에서 먼저 트리거를 정리해 주세요.');
    }
    const resumeInstall=archiveResumeReadV2_(ARCHIVE_RESUME_V2.INSTALL);
    if(resumeInstall&&resumeInstall.owner!==current)throw Error('백업 이어하기 트리거 설치 계정에서 먼저 인계를 진행하세요.');
    let removedArchiveTriggers = 0;
    ScriptApp.getProjectTriggers().forEach(function(trigger) {
      if (trigger.getHandlerFunction() === ARCHIVE_CONFIG_V1.TRIGGER_HANDLER || trigger.getHandlerFunction() === ARCHIVE_RESUME_V2.HANDLER) {
        ScriptApp.deleteTrigger(trigger);
        removedArchiveTriggers++;
      }
    });

    const now = new Date().toISOString();
    dashboardGrantUserRecord_(users, primary, current, now);
    admins[primary] = Object.assign({}, admins[primary] || {}, {promotedAt:(admins[primary] && admins[primary].promotedAt) || now, promotedBy:(admins[primary] && admins[primary].promotedBy) || current});
    props.setProperty(DASHBOARD_ACCESS_KEYS_V2.PRIMARY_ADMIN, target);
    props.setProperty(DASHBOARD_ACCESS_KEYS_V2.LEGACY_ADMIN, target);
    dashboardWriteUsers_(users);
    dashboardWriteAdmins_(admins);
    props.deleteProperty(ARCHIVE_RESUME_V2.INSTALL);
    props.deleteProperty(ARCHIVE_CONFIG_V1.TRIGGER_OWNER_KEY);
    props.deleteProperty(ARCHIVE_CONFIG_V1.TRIGGER_ID_KEY);
    props.deleteProperty(ARCHIVE_CONFIG_V1.TRIGGER_INSTALLED_AT_KEY);
    return {ok:true,previousPrimaryAdminEmail:primary,primaryAdminEmail:target,removedArchiveTriggers:removedArchiveTriggers,archiveTriggerReinstallRequired:true};
  });
}
