const TIME_ZONE = 'Asia/Seoul';
const HEADER_ROW = 2;
const FIRST_DATA_ROW = 3;

// Set these in Apps Script > Project Settings > Script properties:
// SPREADSHEET_ID = your Google Spreadsheet ID
// ADMIN_PASSWORD = your chosen administrator password

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('출석 체크')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('관리자 설정 오류: SPREADSHEET_ID가 설정되지 않았습니다.');
  return SpreadsheetApp.openById(id);
}

function getSheet_() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName('Sheet1') || ss.getSheets()[0];
  if (!sheet) throw new Error('스프레드시트에서 시트를 찾을 수 없습니다.');
  return sheet;
}

function getRoster_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return [];
  // Based on the supplied workbook: B=조, C=공동체, D=이름, G=연락처, I=1주차.
  const values = sheet.getRange(FIRST_DATA_ROW, 2, lastRow - FIRST_DATA_ROW + 1, 8).getDisplayValues();
  let lastGroup = '';
  const roster = [];
  values.forEach((row, index) => {
    // B:I => [조, 공동체, 이름, 생년월일(실제 파일은 성별 값), 성별(실제 파일은 생년월일 값), 연락처, 빈 열, 1주차]
    if (String(row[0] || '').trim()) lastGroup = String(row[0]).trim();
    const community = String(row[1] || '').trim();
    const name = String(row[2] || '').trim();
    const phone = String(row[5] || '').trim();
    const attendance = String(row[7] || '').trim();
    if (!name || !phone) return;
    roster.push({
      rowNumber: FIRST_DATA_ROW + index,
      group: lastGroup,
      community: community,
      name: name,
      phone: phone,
      last4: phone.replace(/\D/g, '').slice(-4),
      attendance: attendance
    });
  });
  return roster;
}

function submitAttendance(name, phoneLast4) {
  name = String(name || '').trim();
  phoneLast4 = String(phoneLast4 || '').replace(/\D/g, '');
  if (!name || !/^\d{4}$/.test(phoneLast4)) {
    return { ok: false, message: '이름과 휴대폰 번호 뒤 4자리를 정확히 입력해 주세요.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const matches = getRoster_().filter(p => p.name === name && p.last4 === phoneLast4);
    if (matches.length === 0) {
      return { ok: false, message: '일치하는 참가자를 찾지 못했습니다. 이름과 번호를 확인해 주세요.' };
    }
    if (matches.length > 1) {
      return { ok: false, message: '동일한 이름과 번호 뒤 4자리를 가진 참가자가 여러 명입니다. 운영진에게 확인해 주세요.' };
    }

    const person = matches[0];
    const sheet = getSheet_();
    const now = new Date();
    // Repeated submissions intentionally overwrite the previous timestamp.
    sheet.getRange(person.rowNumber, 9).setValue(now).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    SpreadsheetApp.flush();

    return {
      ok: true,
      name: person.name,
      group: person.group,
      community: person.community,
      timestamp: Utilities.formatDate(now, TIME_ZONE, 'yyyy-MM-dd HH:mm:ss'),
      message: '출석이 등록되었습니다.'
    };
  } finally {
    lock.releaseLock();
  }
}

function verifyAdmin(password) {
  return isAdmin_(password);
}

function getAdminData(password) {
  requireAdmin_(password);
  return getRoster_().map(p => ({
    group: p.group,
    community: p.community,
    name: p.name,
    phone: p.phone,
    attendance: p.attendance
  }));
}

function exportWorkbook(password) {
  requireAdmin_(password);
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const url = 'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(id) + '/export?format=xlsx';
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    throw new Error('엑셀 파일을 생성하지 못했습니다. 스프레드시트 접근 권한과 배포 설정을 확인해 주세요.');
  }
  return {
    filename: '출석부_' + Utilities.formatDate(new Date(), TIME_ZONE, 'yyyyMMdd_HHmm') + '.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    base64: Utilities.base64Encode(response.getContent())
  };
}

function isAdmin_(password) {
  const saved = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  return !!saved && String(password || '') === saved;
}

function requireAdmin_(password) {
  if (!isAdmin_(password)) throw new Error('관리자 비밀번호가 올바르지 않습니다.');
}
