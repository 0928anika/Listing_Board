// Dashboard popup access control. Source sheets remain unchanged and accessible as before.
const DASHBOARD_ACCESS_KEYS_V2 = Object.freeze({
  PRIMARY_ADMIN:'DASHBOARD_ACCESS_PRIMARY_ADMIN_EMAIL_V2',
  ADMINS:'DASHBOARD_ACCESS_ADMINS_V2',
  USERS:'DASHBOARD_ACCESS_USERS_V1',
  SCHEMA:'DASHBOARD_ACCESS_SCHEMA_VERSION',
  LEGACY_ADMIN:'DASHBOARD_ACCESS_ADMIN_EMAIL_V1',
  LEGACY_PASSWORD:'DASHBOARD_ACCESS_PASSWORD_V1'
});

function dashboardAccessProperties_() {
  const props = PropertiesService.getDocumentProperties();
  if (!props) throw new Error('Dashboard 권한 저장소를 사용할 수 없습니다. 관리자에게 문의해 주세요.');
  return props;
}

function dashboardNormalizeEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function dashboardCurrentEmail_() {
  return dashboardNormalizeEmail_(Session.getActiveUser().getEmail());
}

function dashboardRequireCurrentEmail_() {
  const email = dashboardCurrentEmail_();
  if (!email) {
    throw new Error('현재 Google 계정 이메일을 확인할 수 없습니다. 이 시트를 로그인한 계정으로 다시 열고 권한을 승인해 주세요.');
  }
  return email;
}

function dashboardSpreadsheetOwnerEmail_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const owner = ss && typeof ss.getOwner === 'function' ? ss.getOwner() : null;
  return dashboardNormalizeEmail_(owner && owner.getEmail ? owner.getEmail() : '');
}

// During migration the installed V1 administrator wins; the spreadsheet owner
// is used only when no administrator has been stored. Later transfers are explicit.
function dashboardPrimaryAdminEmail_() {
  const props = dashboardAccessProperties_();
  const stored = dashboardNormalizeEmail_(props.getProperty(DASHBOARD_ACCESS_KEYS_V2.PRIMARY_ADMIN));
  const legacy = dashboardNormalizeEmail_(props.getProperty(DASHBOARD_ACCESS_KEYS_V2.LEGACY_ADMIN));
  const email = stored || legacy || dashboardSpreadsheetOwnerEmail_();
  if (!email) throw new Error('Dashboard 주 관리자를 확인할 수 없습니다. 문서 소유자가 권한 관리를 초기 설정해 주세요.');
  return email;
}

function dashboardReadUsers_() {
  const raw = dashboardAccessProperties_().getProperty(DASHBOARD_ACCESS_KEYS_V2.USERS);
  if (!raw) return {};
  let parsed;
  try { parsed = JSON.parse(raw); }
  catch (error) { throw new Error('Dashboard 허용 목록이 손상되었습니다. 관리자가 설정을 확인해 주세요.'); }
  if (!parsed || Array.isArray(parsed) || typeof parsed !== 'object') {
    throw new Error('Dashboard 허용 목록 형식이 올바르지 않습니다.');
  }
  const normalized = {};
  Object.keys(parsed).forEach(function(key) {
    const email = dashboardNormalizeEmail_(key), value = parsed[key];
    if (!email || !value || typeof value !== 'object' || Array.isArray(value)) return;
    normalized[email] = value;
  });
  return normalized;
}

function dashboardWriteUsers_(users) {
  dashboardAccessProperties_().setProperty(DASHBOARD_ACCESS_KEYS_V2.USERS, JSON.stringify(users));
}

function dashboardReadAdmins_() {
  const raw = dashboardAccessProperties_().getProperty(DASHBOARD_ACCESS_KEYS_V2.ADMINS);
  if (!raw) return {};
  let parsed;
  try { parsed = JSON.parse(raw); }
  catch (error) { throw new Error('Dashboard 관리자 목록이 손상되었습니다. 주 관리자가 설정을 확인해 주세요.'); }
  if (!parsed || Array.isArray(parsed) || typeof parsed !== 'object') {
    throw new Error('Dashboard 관리자 목록 형식이 올바르지 않습니다.');
  }
  const normalized = {};
  Object.keys(parsed).forEach(function(key) {
    const email = dashboardNormalizeEmail_(key), value = parsed[key];
    if (!email || !value || typeof value !== 'object' || Array.isArray(value)) return;
    normalized[email] = value;
  });
  return normalized;
}

function dashboardWriteAdmins_(admins) {
  dashboardAccessProperties_().setProperty(DASHBOARD_ACCESS_KEYS_V2.ADMINS, JSON.stringify(admins));
}

function dashboardGrantUserRecord_(users, email, actor, now) {
  const previous = users[email] || {}, alreadyGranted = previous.status === 'GRANTED';
  users[email] = Object.assign({}, previous, {
    status:'GRANTED',
    grantedAt:alreadyGranted && previous.grantedAt ? previous.grantedAt : now,
    grantedBy:alreadyGranted && previous.grantedBy ? previous.grantedBy : actor
  });
  delete users[email].revokedAt;
  delete users[email].revokedBy;
  return users[email];
}

function dashboardStateForEmail_(email) {
  const primaryAdminEmail = dashboardPrimaryAdminEmail_();
  const users = dashboardReadUsers_(), admins = dashboardReadAdmins_(), record = users[email] || null;
  const isPrimaryAdmin = !!email && email === primaryAdminEmail;
  const isSecondaryAdmin = !!email && !!admins[email] && !!record && record.status === 'GRANTED';
  const isAdmin = isPrimaryAdmin || isSecondaryAdmin;
  return {
    configured:!!primaryAdminEmail,
    accessModel:'ADMIN_APPROVAL_ONLY',
    email:email,
    identityAvailable:!!email,
    allowed:isAdmin || !!record && record.status === 'GRANTED',
    isAdmin:isAdmin,
    isPrimaryAdmin:isPrimaryAdmin,
    revoked:!isPrimaryAdmin && !!record && record.status === 'REVOKED'
  };
}

function getDashboardAccessState() {
  return dashboardStateForEmail_(dashboardCurrentEmail_());
}

function assertDashboardAccess_() {
  const email = dashboardRequireCurrentEmail_(), state = dashboardStateForEmail_(email);
  if (!state.allowed) {
    if (state.revoked) throw new Error('Dashboard 접근 권한이 회수되었습니다. 관리자에게 문의해 주세요.');
    throw new Error('Dashboard 접근이 승인되지 않은 Google 계정입니다. 관리자에게 문의해 주세요.');
  }
  return state;
}

function assertDashboardAdmin_() {
  const email = dashboardRequireCurrentEmail_(), state = dashboardStateForEmail_(email);
  if (!state.isAdmin) throw new Error('권한 관리자 전용 기능입니다.');
  return email;
}

function dashboardValidateEmail_(value) {
  const email = dashboardNormalizeEmail_(value);
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 254) {
    throw new Error('올바른 이메일을 입력해 주세요.');
  }
  return email;
}

// Migrates the installed V1 allow-list to approval-only V2 access control.
// Granted and revoked users are preserved, while the legacy password property
// is permanently removed. Only the immutable primary administrator may run it.
function installDashboardAccessControlV2() {
  const current = dashboardRequireCurrentEmail_();
  return withOpsDocumentLock_(function() {
    const props = dashboardAccessProperties_();
    const storedPrimary = dashboardNormalizeEmail_(props.getProperty(DASHBOARD_ACCESS_KEYS_V2.PRIMARY_ADMIN));
    const legacyAdmin = dashboardNormalizeEmail_(props.getProperty(DASHBOARD_ACCESS_KEYS_V2.LEGACY_ADMIN));
    const primaryAdminEmail = storedPrimary || legacyAdmin || dashboardSpreadsheetOwnerEmail_();
    if (!primaryAdminEmail || current !== primaryAdminEmail) {
      throw new Error('이 문서의 주 관리자만 권한 관리를 초기 설정할 수 있습니다.');
    }

    const users = dashboardReadUsers_(), admins = dashboardReadAdmins_(), now = new Date().toISOString();
    dashboardGrantUserRecord_(users, primaryAdminEmail, primaryAdminEmail, now);
    admins[primaryAdminEmail] = Object.assign({}, admins[primaryAdminEmail] || {}, {
      promotedAt:(admins[primaryAdminEmail] && admins[primaryAdminEmail].promotedAt) || now,
      promotedBy:(admins[primaryAdminEmail] && admins[primaryAdminEmail].promotedBy) || primaryAdminEmail
    });

    props.setProperty(DASHBOARD_ACCESS_KEYS_V2.PRIMARY_ADMIN, primaryAdminEmail);
    props.setProperty(DASHBOARD_ACCESS_KEYS_V2.LEGACY_ADMIN, primaryAdminEmail);
    props.setProperty(DASHBOARD_ACCESS_KEYS_V2.SCHEMA, '2');
    dashboardWriteUsers_(users);
    dashboardWriteAdmins_(admins);
    const legacyPasswordRemoved = props.getProperty(DASHBOARD_ACCESS_KEYS_V2.LEGACY_PASSWORD) !== null;
    props.deleteProperty(DASHBOARD_ACCESS_KEYS_V2.LEGACY_PASSWORD);

    const result = {
      ok:true,
      primaryAdminEmail:primaryAdminEmail,
      grantedUserCount:Object.keys(users).filter(function(email){ return users[email] && users[email].status === 'GRANTED'; }).length,
      adminCount:Object.keys(admins).filter(function(email){
        return email === primaryAdminEmail || users[email] && users[email].status === 'GRANTED';
      }).length,
      legacyPasswordRemoved:legacyPasswordRemoved
    };
    console.log(JSON.stringify(result));
    return result;
  });
}

function getDashboardAccessAdminData() {
  assertDashboardAdmin_();
  const primaryAdminEmail = dashboardPrimaryAdminEmail_(), users = dashboardReadUsers_(), admins = dashboardReadAdmins_();
  const emails = {};
  Object.keys(users).forEach(function(email){ emails[email] = true; });
  Object.keys(admins).forEach(function(email){ emails[email] = true; });
  emails[primaryAdminEmail] = true;

  const rows = Object.keys(emails).map(function(email) {
    const record = users[email] || {}, adminRecord = admins[email] || {};
    const isPrimaryAdmin = email === primaryAdminEmail;
    const isAdmin = isPrimaryAdmin || !!admins[email] && record.status === 'GRANTED';
    return {
      email:email,
      status:isPrimaryAdmin ? 'GRANTED' : (record.status || 'REVOKED'),
      role:isPrimaryAdmin ? 'PRIMARY_ADMIN' : (isAdmin ? 'ADMIN' : 'VIEWER'),
      isAdmin:isAdmin,
      isPrimaryAdmin:isPrimaryAdmin,
      grantedAt:record.grantedAt || '',
      grantedBy:record.grantedBy || '',
      revokedAt:record.revokedAt || '',
      revokedBy:record.revokedBy || '',
      promotedAt:adminRecord.promotedAt || '',
      promotedBy:adminRecord.promotedBy || ''
    };
  });
  rows.sort(function(left, right) {
    if (left.isPrimaryAdmin !== right.isPrimaryAdmin) return left.isPrimaryAdmin ? -1 : 1;
    if (left.isAdmin !== right.isAdmin) return left.isAdmin ? -1 : 1;
    return left.email.localeCompare(right.email);
  });

  return {
    primaryAdminEmail:primaryAdminEmail,
    adminEmail:primaryAdminEmail,
    canTransferPrimary:dashboardCurrentEmail_() === primaryAdminEmail,
    adminEmails:rows.filter(function(row){ return row.isAdmin; }).map(function(row){ return row.email; }),
    users:rows
  };
}

// Approves a new email as a Dashboard viewer. Revoked accounts require restore.
function grantDashboardAccess(email) {
  const current = dashboardRequireCurrentEmail_(), target = dashboardValidateEmail_(email);
  return withOpsDocumentLock_(function() {
    const admin = assertDashboardAdmin_();
    if (admin !== current) throw new Error('권한 관리자 계정을 확인할 수 없습니다.');
    const primaryAdminEmail = dashboardPrimaryAdminEmail_(), users = dashboardReadUsers_(), admins = dashboardReadAdmins_();
    if (users[target] && users[target].status === 'REVOKED') {
      throw new Error('권한이 회수된 계정입니다. 목록의 복구 기능을 사용해 주세요.');
    }
    dashboardGrantUserRecord_(users, target, admin, new Date().toISOString());
    dashboardWriteUsers_(users);
    const isPrimaryAdmin = target === primaryAdminEmail;
    const isAdmin = isPrimaryAdmin || !!admins[target];
    return {ok:true, email:target, status:'GRANTED', role:isPrimaryAdmin ? 'PRIMARY_ADMIN' : (isAdmin ? 'ADMIN' : 'VIEWER')};
  });
}

function restoreDashboardAccess(email) {
  const current = dashboardRequireCurrentEmail_(), target = dashboardValidateEmail_(email);
  return withOpsDocumentLock_(function() {
    const admin = assertDashboardAdmin_();
    if (admin !== current) throw new Error('권한 관리자 계정을 확인할 수 없습니다.');
    const primaryAdminEmail = dashboardPrimaryAdminEmail_(), users = dashboardReadUsers_(), admins = dashboardReadAdmins_();
    dashboardGrantUserRecord_(users, target, admin, new Date().toISOString());
    dashboardWriteUsers_(users);
    const isPrimaryAdmin = target === primaryAdminEmail;
    const isAdmin = isPrimaryAdmin || !!admins[target];
    return {ok:true, email:target, status:'GRANTED', role:isPrimaryAdmin ? 'PRIMARY_ADMIN' : (isAdmin ? 'ADMIN' : 'VIEWER')};
  });
}

// Revoking a secondary administrator also removes their administrator role.
function revokeDashboardAccess(email) {
  const current = dashboardRequireCurrentEmail_(), target = dashboardValidateEmail_(email);
  return withOpsDocumentLock_(function() {
    const admin = assertDashboardAdmin_();
    if (admin !== current) throw new Error('권한 관리자 계정을 확인할 수 없습니다.');
    const primaryAdminEmail = dashboardPrimaryAdminEmail_();
    if (target === primaryAdminEmail) throw new Error('주 관리자의 권한은 회수할 수 없습니다.');
    const users = dashboardReadUsers_(), admins = dashboardReadAdmins_(), previous = users[target] || {};
    const wasAdmin = !!admins[target];
    users[target] = Object.assign({}, previous, {
      status:'REVOKED',
      revokedAt:new Date().toISOString(),
      revokedBy:admin
    });
    delete admins[target];
    dashboardWriteUsers_(users);
    dashboardWriteAdmins_(admins);
    return {ok:true, email:target, status:'REVOKED', role:'VIEWER', wasAdmin:wasAdmin};
  });
}

// Promotes an approved or new email to administrator. Revoked accounts require restore first.
function grantDashboardAdminAccess(email) {
  const current = dashboardRequireCurrentEmail_(), target = dashboardValidateEmail_(email);
  return withOpsDocumentLock_(function() {
    const admin = assertDashboardAdmin_();
    if (admin !== current) throw new Error('권한 관리자 계정을 확인할 수 없습니다.');
    const primaryAdminEmail = dashboardPrimaryAdminEmail_(), users = dashboardReadUsers_(), admins = dashboardReadAdmins_(), now = new Date().toISOString();
    if (users[target] && users[target].status === 'REVOKED') {
      throw new Error('권한이 회수된 계정입니다. 먼저 Dashboard 접근 권한을 복구해 주세요.');
    }
    dashboardGrantUserRecord_(users, target, admin, now);
    if (target !== primaryAdminEmail) {
      admins[target] = Object.assign({}, admins[target] || {}, {
        promotedAt:(admins[target] && admins[target].promotedAt) || now,
        promotedBy:(admins[target] && admins[target].promotedBy) || admin
      });
    }
    dashboardWriteUsers_(users);
    dashboardWriteAdmins_(admins);
    return {ok:true, email:target, status:'GRANTED', role:target === primaryAdminEmail ? 'PRIMARY_ADMIN' : 'ADMIN'};
  });
}

// Demotes a non-primary administrator to an approved Dashboard viewer.
function demoteDashboardAdminAccess(email) {
  const current = dashboardRequireCurrentEmail_(), target = dashboardValidateEmail_(email);
  return withOpsDocumentLock_(function() {
    const admin = assertDashboardAdmin_();
    if (admin !== current) throw new Error('권한 관리자 계정을 확인할 수 없습니다.');
    const primaryAdminEmail = dashboardPrimaryAdminEmail_();
    if (target === primaryAdminEmail) throw new Error('주 관리자는 관리자 권한을 해제할 수 없습니다.');
    const users = dashboardReadUsers_(), admins = dashboardReadAdmins_();
    if (!admins[target]) throw new Error('해당 계정은 관리자가 아닙니다.');
    delete admins[target];
    dashboardGrantUserRecord_(users, target, admin, new Date().toISOString());
    dashboardWriteUsers_(users);
    dashboardWriteAdmins_(admins);
    return {ok:true, email:target, status:'GRANTED', role:'VIEWER'};
  });
}

function showDashboardAccessManagement() {
  assertDashboardAdmin_();
  const html = HtmlService.createHtmlOutputFromFile('DashboardAccessManagement').setWidth(940).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, '권한 관리');
}
