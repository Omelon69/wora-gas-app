/*************************************************
 * UserAuthLib — 0.836
 * - ใช้หัวตารางระบุคอลัมน์ (Email/Role/Name/Status/Note) -> ไม่ล็อคคอลัมน์
 * - รองรับ APP_SHEET_ID และ fallback ไป SpreadsheetApp.getActive()
 * - getCurrentUserInfo() ใช้ Session.getActiveUser().getEmail()
 * - forceGoogleLogout() / logoutUser() คืน URL ที่ถูกต้อง
 *   (สอดคล้องกับ _accountLinks() ใน Code.gs)
 **************************************************/

/** ===== Utilities: Sheet / Header Map ===== */
function getUserSheet() {
  // ใช้ APP_SHEET_ID เป็นหลัก; ถ้าไม่มีให้ fallback เป็นสเปรดชีตที่สคริปต์ผูกอยู่
  const sid = (PropertiesService.getScriptProperties().getProperty('APP_SHEET_ID') || '').trim();
  let ss = null;
  try {
    ss = sid ? SpreadsheetApp.openById(sid) : SpreadsheetApp.getActive();
  } catch (e) {
    ss = SpreadsheetApp.getActive();
  }
  const sh = ss.getSheetByName('Users');
  if (!sh) throw new Error('ไม่พบชีท Users');
  return sh;
}

/** อ่านทั้งตาราง + สร้างแผนที่หัวตาราง -> index ของคอลัมน์ */
function readUsersAll_() {
  const sh = getUserSheet();
  const rng = sh.getDataRange();
  const vals = rng.getValues();      // [header, ...rows]
  const header = (vals[0] || []).map(h => String(h || '').trim().toLowerCase());

  // จับคีย์หลักตามชื่อหัว (ยืดหยุ่นเล็กน้อย)
  const kEmail  = findCol_(header, ['email','อีเมล','อีเมล์']);
  const kRole   = findCol_(header, ['role','สิทธิ์','บทบาท']);
  const kName   = findCol_(header, ['name','ชื่อ','displayname']);
  const kStatus = findCol_(header, ['status','สถานะ']);
  const kNote   = findCol_(header, ['note','หมายเหตุ','notes']);

  return {
    headerIdx: { email:kEmail, role:kRole, name:kName, status:kStatus, note:kNote },
    rows: vals.slice(1), // ข้าม header
    range: rng,
    sheet: sh
  };
}

function findCol_(headerArr, candidates){
  for (let i=0;i<headerArr.length;i++){
    const h = headerArr[i];
    for (let c of candidates){
      if (h === c.toLowerCase()) return i;
    }
  }
  // ไม่เจอให้คืน -1 (ผู้ใช้บางคนอาจตั้งหัวตารางต่างออกไป)
  return -1;
}

/** ===== Current User ===== */
function getCurrentUserInfo() {
  // ต้องมี scope: https://www.googleapis.com/auth/userinfo.email
  // และ Deployment: Execute as => User accessing the web app
  const rawEmail = Session.getActiveUser().getEmail();
  const email = (rawEmail || '').trim().toLowerCase();
  if (!email) return { found: false, email: '' };

  const pack = readUsersAll_();
  const idx  = pack.headerIdx;

  // ถ้า schema users ไม่ครบ ก็ยังคืน found:false พร้อม email ให้ฟรอนต์แสดงสมัครได้
  if (idx.email < 0) {
    return { found:false, email: email };
  }

  // loop หาแถวตรงกับ email
  for (let i=0; i<pack.rows.length; i++) {
    const r = pack.rows[i];
    const em = String(r[idx.email] || '').trim().toLowerCase();
    if (em === email) {
      return {
        found: true,
        email: email,
        role:  idx.role   >=0 ? (r[idx.role]   || '') : '',
        name:  idx.name   >=0 ? (r[idx.name]   || '') : '',
        status:idx.status >=0 ? (r[idx.status] || '') : ''
      };
    }
  }
  // ยังไม่อยู่ใน Users
  return { found:false, email: email };
}

/** ===== Registration ===== */
function registerConfirmedUser(email) {
  if (!email) return false;
  const em = String(email).trim().toLowerCase();

  const pack = readUsersAll_();
  const { sheet, headerIdx } = pack;

  // ถ้ามีอยู่แล้วไม่ต้องเพิ่ม
  for (let i=0; i<pack.rows.length; i++){
    const r = pack.rows[i];
    const re = (headerIdx.email>=0) ? String(r[headerIdx.email] || '').trim().toLowerCase() : '';
    if (re && re === em) return false;
  }

  // เตรียมคอลัมน์ปลายทางตาม header ที่พบ (ถ้าไม่พบจะเติมท้ายๆ ไปตามเดิม)
  const out = [];
  const width = sheet.getLastColumn();
  for (let c=0;c<width;c++) out.push('');

  // เซตค่าโดยเคารพตำแหน่งคอลัมน์
  putCell_(out, headerIdx.email,  em);
  putCell_(out, headerIdx.role,   '');               // Role ว่าง
  putCell_(out, headerIdx.name,   '');               // Name ว่าง
  putCell_(out, headerIdx.status, 'รอดำเนินการ');   // Status ค่าเริ่มต้น
  putCell_(out, headerIdx.note,   '');               // Note ว่าง

  sheet.appendRow(out);
  return true;
}

function putCell_(rowArr, idx, val){
  if (idx >= 0 && idx < rowArr.length) rowArr[idx] = val;
}

/** ===== Admin: read/update ===== */
function getAllUsers() {
  const pack = readUsersAll_();
  return pack.rows; // ข้าม header แล้ว
}

// updates = [[role, name, status], ...] ตามลำดับ row ปัจจุบัน
function updateUserInfo(updates) {
  if (!updates || !updates.length) return;

  const pack = readUsersAll_();
  const { sheet, headerIdx } = pack;

  const startRow = 2; // แถวแรกของข้อมูลจริง (หลัง header)
  for (let i=0; i<updates.length && i<pack.rows.length; i++) {
    const r = startRow + i;
    const upd = updates[i];

    if (headerIdx.role   >= 0) sheet.getRange(r, headerIdx.role+1  ).setValue(upd[0] || '');
    if (headerIdx.name   >= 0) sheet.getRange(r, headerIdx.name+1  ).setValue(upd[1] || '');
    if (headerIdx.status >= 0) sheet.getRange(r, headerIdx.status+1).setValue(upd[2] || '');
  }
}

/** ===== Query by Email ===== */
function getUserDetailsByEmail(email) {
  if (!email) return null;
  const em = String(email).trim().toLowerCase();

  const pack = readUsersAll_();
  const idx  = pack.headerIdx;
  if (idx.email < 0) return null;

  for (let i=0; i<pack.rows.length; i++) {
    const r = pack.rows[i];
    const re = String(r[idx.email] || '').trim().toLowerCase();
    if (re === em) {
      return {
        email: r[idx.email]  || '',
        role:  idx.role >=0 ? (r[idx.role] || '') : '',
        name:  idx.name >=0 ? (r[idx.name] || '') : ''
      };
    }
  }
  return null;
}

/** ===== Account Switch / Logout (สอดคล้อง Code.gs) =====
 * - changeAccount(): ใช้ AccountChooser
 * - logout(): ออกจากระบบ Google แล้วพากลับไป chooser
 * ฟรอนต์ใหม่เรียก getAccountLinks() จาก Code.gs ได้เลย
 * แต่คงฟังก์ชันเดิมไว้ให้เข้ากันย้อนหลัง (backward compatible)
 */
function forceGoogleLogout() {
  // เดิมถูกใช้กับปุ่ม "เปลี่ยนบัญชี" → ควรพาไป AccountChooser
  var base = ScriptApp.getService().getUrl();
  return 'https://accounts.google.com/AccountChooser?continue=' + encodeURIComponent(base);
}

function logoutUser() {
  // เดิมฟรอนต์เรียกแล้ว reload; ตอนนี้ให้คืน true เฉยๆ เพื่อไม่พัง
  // (ฟรอนต์ใหม่ควรใช้ links.logout จาก getAccountLinks() แทน)
  return true;
}
