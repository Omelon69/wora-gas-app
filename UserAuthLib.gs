function getUserSheet() {
  const sheetId = PropertiesService.getScriptProperties().getProperty('APP_SHEET_ID');
  return SpreadsheetApp.openById(sheetId).getSheetByName('Users');
}

function getCurrentUserInfo() {
  const rawEmail = Session.getActiveUser().getEmail();
  const email = (rawEmail || '').trim().toLowerCase();
  if (!email) return { found: false, email: '' };

  const sheet = getUserSheet();
  const data = sheet.getDataRange().getValues(); // [header,...]

  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] + '').trim().toLowerCase() === email) {
      return {
        found: true,
        email,
        role: data[i][1],
        name: data[i][2],
        status: data[i][3] || '' // 'อนุมัติ' | 'อนุมัติแล้ว' | 'รอดำเนินการ' | 'ปฏิเสธ'
      };
    }
  }
  return { found: false, email };
}

function registerConfirmedUser(email) {
  if (!email) return false;
  const sheet = getUserSheet();
  const emails = sheet.getDataRange().getValues().map(r => (r[0] + '').trim().toLowerCase());
  if (!emails.includes(email.trim().toLowerCase())) {
    // Email, Role, Name, Status, Note
    sheet.appendRow([email, '', '', 'รอดำเนินการ', '']);
    return true;
  }
  return false;
}

function getAllUsers() {
  const sheet = getUserSheet();
  const data = sheet.getDataRange().getValues();
  return data.slice(1); // skip header
}

function updateUserInfo(updates) {
  // updates = [[role, name, status], ...] อ้างอิงลำดับปัจจุบันของชีต
  const sheet = getUserSheet();
  const data = sheet.getDataRange().getValues(); // [header,...]
  for (let i = 1; i < data.length && i <= updates.length; i++) {
    sheet.getRange(i + 1, 2).setValue(updates[i - 1][0]); // Role
    sheet.getRange(i + 1, 3).setValue(updates[i - 1][1]); // Name
    sheet.getRange(i + 1, 4).setValue(updates[i - 1][2]); // Status
  }
}

function getUserDetailsByEmail(email) {
  if (!email) return null;
  const sheet = getUserSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] + '').trim().toLowerCase() === email.trim().toLowerCase()) {
      return {
        email: data[i][0] || '',
        role: data[i][1] || '',
        name: data[i][2] || ''
      };
    }
  }
  return null;
}

function forceGoogleLogout() {
  return 'https://accounts.google.com/Logout';
}

function logoutUser() {
  return true;
}
