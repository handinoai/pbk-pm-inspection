function doGet(e) {
  var callback = e.parameter.callback;
  var action = e.parameter.action || 'daily_worker';
  var fromDate = e.parameter.from || '';
  var toDate = e.parameter.to || '';
  var toolId = e.parameter.tool_id || '';
  
  var result;
  
  try {
    if (action === 'daily_worker') {
      result = getDailyWorkerRecords(fromDate, toDate);
    } else if (action === 'daily_hipot') {
      result = getDailyHipotRecords(fromDate, toDate);
    } else if (action === 'tool_trend') {
      result = getToolTrend(toolId);
    } else if (action === 'all_tool_trends') {
      result = getAllToolTrends();
    } else if (action === 'tool_warnings') {
      result = getToolWarnings();
    } else if (action === 'monthly') {
      result = getMonthlyRecords(fromDate, toDate);
    } else if (action === 'save_monthly') {
      result = saveMonthlyRecord(e.parameter);
    } else if (action === 'biyearly') {
      result = getBiyearlyRecords(fromDate, toDate);
    } else if (action === 'save_biyearly') {
      result = saveBiyearlyRecord(e.parameter);
    } else if (action === 'login_user') {
      result = handleLoginUser(e.parameter);
    } else if (action === 'register_user') {
      result = handleRegisterUser(e.parameter);
    } else if (action === 'approve_user') {
      result = handleApproveUser(e.parameter);
    } else if (action === 'get_users') {
      result = handleGetUsers(e.parameter);
    } else {
      result = { error: 'Unknown action' };
    }
  } catch (error) {
    result = { error: error.toString() };
  }
  
  var output = JSON.stringify(result);
  
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + output + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  } else {
    return ContentService
      .createTextOutput(output)
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getDailyWorkerRecords(fromDate, toDate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('daily_worker_records');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var records = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var record = {};
    for (var j = 0; j < headers.length; j++) {
      record[headers[j]] = row[j];
    }
    if (record.date instanceof Date) {
      record.date = Utilities.formatDate(record.date, 'Asia/Seoul', 'yyyy-MM-dd');
    }
    if (fromDate && record.date < fromDate) continue;
    if (toDate && record.date > toDate) continue;
    records.push(record);
  }
  
  return { success: true, records: records };
}

function getDailyHipotRecords(fromDate, toDate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('daily_hipot_records');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var records = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var record = {};
    for (var j = 0; j < headers.length; j++) {
      record[headers[j]] = row[j];
    }
    if (record.date instanceof Date) {
      record.date = Utilities.formatDate(record.date, 'Asia/Seoul', 'yyyy-MM-dd');
    }
    if (fromDate && record.date < fromDate) continue;
    if (toDate && record.date > toDate) continue;
    records.push(record);
  }
  
  return { success: true, records: records };
}

// ============ Monthly Records ============

function getMonthlyRecords(fromDate, toDate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('monthly_records');
  if (!sheet) return { success: true, records: [] };
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, records: [] };
  
  var headers = data[0];
  var records = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var record = {};
    for (var j = 0; j < headers.length; j++) {
      record[headers[j]] = row[j];
    }
    
    // period 날짜 변환
    var period = record.period;
    if (period instanceof Date) {
      period = Utilities.formatDate(period, 'Asia/Seoul', 'yyyy-MM');
      record.period = period;
    } else {
      period = String(period || '').slice(0, 7);
      record.period = period;
    }
    if (!period) continue;
    
    // gloves 날짜 필드 변환
    if (record.gloves_open_date instanceof Date) {
      record.gloves_open_date = Utilities.formatDate(record.gloves_open_date, 'Asia/Seoul', 'yyyy-MM-dd');
    }
    if (record.gloves_expiry instanceof Date) {
      record.gloves_expiry = Utilities.formatDate(record.gloves_expiry, 'Asia/Seoul', 'yyyy-MM-dd');
    }
    
    if (fromDate) {
      var fromMonth = fromDate.slice(0, 7);
      if (period < fromMonth) continue;
    }
    if (toDate) {
      var toMonth = toDate.slice(0, 7);
      if (period > toMonth) continue;
    }
    
    records.push(record);
  }
  
  return { success: true, records: records };
}

function saveMonthlyRecord(params) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('monthly_records');
  if (!sheet) return { success: false, error: 'monthly_records 시트가 없습니다' };
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var periodIdx = headers.indexOf('period');
  var existingRow = -1;
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][periodIdx]) === String(params.period)) {
      existingRow = i + 1;
      break;
    }
  }
  
  var row = [
    params.period || '', params.inspector || '',
    params.load_pass_value || '', params.load_pass_result || '',
    params.load_fail_value || '', params.load_fail_result || '',
    params.gloves_open_date || '', params.gloves_expiry || '', params.gloves_result || '',
    params.mv_fixture1_check1 || '', params.mv_fixture1_check2 || '', params.mv_fixture1_result || '',
    params.mv_fixture2_check1 || '', params.mv_fixture2_check2 || '', params.mv_fixture2_result || '',
    params.mv_fixture3_check1 || '', params.mv_fixture3_check2 || '', params.mv_fixture3_result || '',
    params.mv_fixture_remarks || '',
    params.mv_tray_serial || '', params.mv_tray_check1 || '', params.mv_tray_check2 || '',
    params.mv_tray_result || '', params.mv_tray_remarks || '',
    params.overall_result || '', params.remarks || '',
    params.created_at || Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')
  ];
  
  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  
  return { success: true, message: 'Monthly 기록 저장 완료' };
}

// ============ Bi-yearly Records ============

function getBiyearlyRecords(fromDate, toDate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('biyearly_records');
  if (!sheet) return { success: true, records: [] };
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: true, records: [] };
  
  var headers = data[0];
  var records = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var record = {};
    for (var j = 0; j < headers.length; j++) {
      record[headers[j]] = row[j];
    }
    
    // period를 문자열로 확실하게 처리 (2026-H1 형식)
    var period = String(record.period || '');
    record.period = period;
    if (!period) continue;
    
    // 필터링: period에서 연도 추출하여 비교
    if (fromDate) {
      var fromYear = fromDate.slice(0, 4);
      var periodYear = period.slice(0, 4);
      if (periodYear < fromYear) continue;
    }
    if (toDate) {
      var toYear = toDate.slice(0, 4);
      var periodYear2 = period.slice(0, 4);
      if (periodYear2 > toYear) continue;
    }
    
    records.push(record);
  }
  
  return { success: true, records: records };
}

function saveBiyearlyRecord(params) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('biyearly_records');
  if (!sheet) return { success: false, error: 'biyearly_records 시트가 없습니다' };
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var periodIdx = headers.indexOf('period');
  var existingRow = -1;
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][periodIdx]) === String(params.period)) {
      existingRow = i + 1;
      break;
    }
  }
  
  var row = [
    params.period || '', params.inspector || '',
    params.surface_1_id || '', params.surface_1_value || '', params.surface_1_result || '',
    params.surface_2_id || '', params.surface_2_value || '', params.surface_2_result || '',
    params.surface_3_id || '', params.surface_3_value || '', params.surface_3_result || '',
    params.surface_4_id || '', params.surface_4_value || '', params.surface_4_result || '',
    params.surface_5_id || '', params.surface_5_value || '', params.surface_5_result || '',
    params.wrist_1_id || '', params.wrist_1_value || '', params.wrist_1_result || '',
    params.wrist_2_id || '', params.wrist_2_value || '', params.wrist_2_result || '',
    params.wrist_3_id || '', params.wrist_3_value || '', params.wrist_3_result || '',
    params.wrist_4_id || '', params.wrist_4_value || '', params.wrist_4_result || '',
    params.wrist_5_id || '', params.wrist_5_value || '', params.wrist_5_result || '',
    params.overall_result || '', params.remarks || '',
    params.created_at || Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss')
  ];
  
  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  
  return { success: true, message: 'Bi-yearly 기록 저장 완료' };
}

// ============ Tool Trend ============

function getAllToolTrends() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('daily_worker_records');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var dateIdx = headers.indexOf('date');
  var toolIdIdx = headers.indexOf('tool_id');
  var t1Idx = headers.indexOf('torque_1');
  var t2Idx = headers.indexOf('torque_2');
  var t3Idx = headers.indexOf('torque_3');
  
  var toolData = {};
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var toolId = row[toolIdIdx];
    var date = row[dateIdx];
    if (!toolId) continue;
    if (date instanceof Date) {
      date = Utilities.formatDate(date, 'Asia/Seoul', 'yyyy-MM-dd');
    }
    
    var t1 = parseFloat(row[t1Idx]) || 0;
    var t2 = parseFloat(row[t2Idx]) || 0;
    var t3 = parseFloat(row[t3Idx]) || 0;
    var values = [];
    if (t1 > 0) values.push(t1);
    if (t2 > 0) values.push(t2);
    if (t3 > 0) values.push(t3);
    if (values.length === 0) continue;
    
    var avg = values.reduce(function(a, b) { return a + b; }, 0) / values.length;
    if (!toolData[toolId]) toolData[toolId] = [];
    toolData[toolId].push({ date: date, avg: avg, values: values });
  }
  
  var SPEC_MIN = 12.32, SPEC_MAX = 16.67;
  var trends = [];
  
  for (var tid in toolData) {
    var records = toolData[tid];
    records.sort(function(a, b) { return b.date.localeCompare(a.date); });
    
    var recentDays = [], uniqueDates = [];
    for (var k = 0; k < records.length; k++) {
      if (uniqueDates.indexOf(records[k].date) === -1) {
        uniqueDates.push(records[k].date);
        if (uniqueDates.length > 5) break;
      }
      if (uniqueDates.length <= 5) recentDays.push(records[k]);
    }
    if (recentDays.length === 0) continue;
    
    var allAvgs = recentDays.map(function(r) { return r.avg; });
    var overallAvg = allAvgs.reduce(function(a, b) { return a + b; }, 0) / allAvgs.length;
    var minVal = Math.min.apply(null, allAvgs);
    var maxVal = Math.max.apply(null, allAvgs);
    var position = (overallAvg - SPEC_MIN) / (SPEC_MAX - SPEC_MIN) * 100;
    
    var WARNING_LOWER = 5, WARNING_UPPER = 95, CONSECUTIVE_DAYS_REQUIRED = 5;
    var warning = null;
    
    if (uniqueDates.length >= CONSECUTIVE_DAYS_REQUIRED) {
      var dateAvgs = {};
      for (var d = 0; d < recentDays.length; d++) {
        var dt = recentDays[d].date;
        if (!dateAvgs[dt]) dateAvgs[dt] = [];
        dateAvgs[dt].push(recentDays[d].avg);
      }
      var recentDates = uniqueDates.slice(0, CONSECUTIVE_DAYS_REQUIRED);
      var lowerCount = 0, upperCount = 0;
      for (var i = 0; i < recentDates.length; i++) {
        var dayValues = dateAvgs[recentDates[i]];
        var dayAvg = dayValues.reduce(function(a, b) { return a + b; }, 0) / dayValues.length;
        var dayPosition = (dayAvg - SPEC_MIN) / (SPEC_MAX - SPEC_MIN) * 100;
        if (dayPosition < WARNING_LOWER) lowerCount++;
        if (dayPosition > WARNING_UPPER) upperCount++;
      }
      if (lowerCount >= CONSECUTIVE_DAYS_REQUIRED) warning = '하한 근접 (5일 연속)';
      else if (upperCount >= CONSECUTIVE_DAYS_REQUIRED) warning = '상한 근접 (5일 연속)';
    }
    
    trends.push({
      tool_id: tid, sample_count: uniqueDates.length, total_measurements: recentDays.length,
      average: Math.round(overallAvg * 100) / 100, min: Math.round(minVal * 100) / 100,
      max: Math.round(maxVal * 100) / 100, position_percent: Math.round(position),
      warning: warning, last_date: recentDays[0].date, recent_data: recentDays.slice(0, 5)
    });
  }
  
  trends.sort(function(a, b) { return a.tool_id.localeCompare(b.tool_id); });
  return { success: true, trends: trends };
}

function getToolTrend(toolId) {
  var result = getAllToolTrends();
  if (!result.success) return result;
  for (var i = 0; i < result.trends.length; i++) {
    if (result.trends[i].tool_id === toolId) return { success: true, trend: result.trends[i] };
  }
  return { success: true, trend: null };
}

function getToolWarnings() {
  var result = getAllToolTrends();
  if (!result.success) return result;
  var warnings = [];
  for (var i = 0; i < result.trends.length; i++) {
    if (result.trends[i].warning !== null) warnings.push(result.trends[i]);
  }
  return { success: true, warnings: warnings };
}
// ============================================================
// USER AUTH FUNCTIONS
// ============================================================

var ADMIN_EMAIL = 'inho.son@promega.com';

function handleLoginUser(params) {
  try {
    var username = String(params.username || '').trim().toLowerCase();
    var passwordHash = String(params.password_hash || '').trim();
    if (!username || !passwordHash) {
      return { success: false, error: '아이디와 비밀번호를 입력해주세요.' };
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Users');
    if (!sheet) {
      return { success: false, error: '사용자 시트가 없습니다. 관리자에게 문의하세요.' };
    }
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (String(row[0] || '').trim().toLowerCase() !== username) continue;
      if (String(row[1] || '').trim() !== passwordHash) {
        return { success: false, error: '비밀번호가 올바르지 않습니다.' };
      }
      var status = String(row[4] || '').trim();
      if (status === 'pending') return { success: false, error: '가입 신청이 승인 대기 중입니다. 관리자 승인 후 로그인이 가능합니다.' };
      if (status === 'rejected') return { success: false, error: '가입 신청이 거절되었습니다. 관리자에게 문의하세요.' };
      if (status !== 'approved') return { success: false, error: '계정 상태를 확인해주세요.' };
      var userEmail = String(row[3] || '');
      var isAdmin = userEmail.toLowerCase() === ADMIN_EMAIL.toLowerCase();
      return {
        success: true,
        user: {
          username: String(row[0]).trim().toLowerCase(),
          name: String(row[2] || ''),
          email: userEmail,
          role: isAdmin ? 'admin' : 'user'
        }
      };
    }
    return { success: false, error: '존재하지 않는 아이디입니다.' };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function handleRegisterUser(params) {
  try {
    var username = String(params.username || '').trim().toLowerCase();
    var passwordHash = String(params.password_hash || '').trim();
    var name = String(params.name || '').trim();
    var email = String(params.email || '').trim();
    if (!username || !passwordHash || !name || !email) {
      return { success: false, error: '모든 항목을 입력해주세요.' };
    }
    if (!/^[a-z0-9_]+$/.test(username)) {
      return { success: false, error: '아이디는 영문 소문자, 숫자, 밑줄(_)만 사용 가능합니다.' };
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Users');
    if (!sheet) {
      sheet = ss.insertSheet('Users');
      sheet.appendRow(['username','password_hash','name','email','status','created_at','approved_at']);
    }
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim().toLowerCase() === username) {
        return { success: false, error: '이미 사용 중인 아이디입니다.' };
      }
    }
    var now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([username, passwordHash, name, email, 'pending', now, '']);
    try {
      var scriptUrl = ScriptApp.getService().getUrl();
      var approveUrl = scriptUrl + '?action=approve_user&username=' + encodeURIComponent(username) + '&action_type=approve';
      var rejectUrl  = scriptUrl + '?action=approve_user&username=' + encodeURIComponent(username) + '&action_type=reject';
      var subject = '[PBK PM Inspection] 회원가입 승인 요청: ' + name + ' (' + username + ')';
      var body = '안녕하세요,\n\nPBK PM Inspection 시스템에 새로운 가입 신청이 있습니다.\n\n'
        + '이름: ' + name + '\n아이디: ' + username + '\n이메일: ' + email + '\n신청일시: ' + now + '\n\n'
        + '아래 링크를 클릭하여 처리해주세요:\n\n'
        + '✅ 승인: ' + approveUrl + '\n\n'
        + '❌ 거절: ' + rejectUrl + '\n\n'
        + '또는 대시보드 [설정 > 사용자 관리] 메뉴에서 처리하실 수 있습니다.\n\n감사합니다.';
      MailApp.sendEmail(ADMIN_EMAIL, subject, body);
    } catch(mailErr) {
      Logger.log('이메일 발송 실패: ' + mailErr.toString());
    }
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function handleApproveUser(params) {
  try {
    var username = String(params.username || '').trim().toLowerCase();
    var actionType = String(params.action_type || '').trim();
    if (!username) return { success: false, error: '아이디가 필요합니다.' };
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Users');
    if (!sheet) return { success: false, error: 'Users 시트가 없습니다.' };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim().toLowerCase() !== username) continue;
      var newStatus = '';
      if (actionType === 'approve') {
        newStatus = 'approved';
        var approvedAt = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
        sheet.getRange(i + 1, 7).setValue(approvedAt);
        try {
          var userEmail = String(data[i][3] || '');
          var userName  = String(data[i][2] || '');
          if (userEmail) {
            MailApp.sendEmail(userEmail, '[PBK PM Inspection] 가입 승인 완료',
              '안녕하세요 ' + userName + '님,\n\nPBK PM Inspection 가입이 승인되었습니다.\n이제 로그인하여 사용하실 수 있습니다.\n\n감사합니다.');
          }
        } catch(e) {}
      } else if (actionType === 'reject') {
        newStatus = 'rejected';
      } else if (actionType === 'revoke') {
        newStatus = 'pending';
        sheet.getRange(i + 1, 7).setValue('');
      } else {
        return { success: false, error: '올바르지 않은 action_type' };
      }
      sheet.getRange(i + 1, 5).setValue(newStatus);
      return { success: true };
    }
    return { success: false, error: '사용자를 찾을 수 없습니다.' };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function handleGetUsers(params) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Users');
    if (!sheet) return { success: true, users: [] };
    var data = sheet.getDataRange().getValues();
    var users = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue;
      users.push({
        username:    String(row[0] || ''),
        name:        String(row[2] || ''),
        email:       String(row[3] || ''),
        status:      String(row[4] || ''),
        created_at:  String(row[5] || ''),
        approved_at: String(row[6] || '')
      });
    }
    return { success: true, users: users };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

// MailApp 권한 승인용 테스트 함수 - Apps Script 에디터에서 직접 실행하세요
function testEmailAuth() {
  MailApp.sendEmail(
    ADMIN_EMAIL,
    '[PBK PM Inspection] 이메일 권한 테스트',
    '이 메일이 수신되면 MailApp 권한이 정상적으로 승인된 것입니다.'
  );
  Logger.log('테스트 이메일 발송 완료: ' + ADMIN_EMAIL);
}
