function doGet(e) {
  const page = (e && e.parameter.page) ? e.parameter.page : 'index';
  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle('WoraCRM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

/** ================================================
 *  WoraCRM v0.81 — Data & Access Base (Single file)
 *  - Sheet หลัก: 'Online' (fallback 'ช่องทางonline')
 *  - คีย์คอลัมน์อังกฤษตามคู่มือ (A..AI → key)
 *  - RBAC: เฉพาะ role 'sales' เห็นเฉพาะแถวที่ S:sales_owner ตรงชื่อ
 *  - พร้อมต่อยอด 0.82+ (Enumus), 0.83 (Compact), 0.84 (Full)
 *  ================================================ */

/** 👉 เปลี่ยนตามไฟล์จริงของคุณ (ตรงกับลิงก์ Google Sheet) */
const SHEET_DOC_ID = '1OPt1W2BuzIMqE6q4S5gCe4zxC4yQxKsUbTjumjsrV48';

/** ONLINE sheet + columns mapping (อิงคู่มือ) */
const ONLINE_CFG = {
  DATA_SHEET_CANDIDATES: ['Online', 'ช่องทางonline'], // ใช้ Online เป็นหลัก
  ENUMUS_SHEET_CANDIDATES: ['Enumus', '__enumus_col__X'],
  USERS_SHEET_CANDIDATES: ['Users'],
  RANGE_A1: 'A1:AI', // ครอบคลุมคอลัมน์ A..AI
  COL: { // A..AI → key อังกฤษ (ห้ามแก้ชื่อ key)
    A:'row_no', B:'prospect_code', C:'erp_code', D:'created_at', E:'date_text',
    F:'yyyymm', G:'year', H:'month', I:'day', J:'seq_in_month',
    K:'company', L:'contacts_count', M:'area_text', N:'lead_source',
    O:'is_new_customer', P:'added_line', Q:'admin_owner', R:'case_owner',
    S:'sales_owner', T:'handoff_date', U:'items', V:'amount', W:'needs',
    X:'status', Y:'quote_date', Z:'last_followup_date', AA:'last_follower',
    AB:'po_date', AC:'payment_term', AD:'so_number', AE:'next_followup_date',
    AF:'status_changed_at', AG:'owner_changed_at', AH:'updated_at', AI:'is_real_customer'
  }
};

/** RBAC: กติกาการมองเห็น */
const RBAC = {
  // เฉพาะ role 'sales' ที่จำกัดให้เห็นเฉพาะแถวที่ตนเป็นผู้ดูแล (คอลัมน์ S:sales_owner)
  VISIBILITY_MATCH_COLUMN: 'sales_owner',
  DEFAULT_ROLE: 'viewer'
};

/** ===== Utils: วันที่/ตัวเลข/ชื่อ ===== */
function pad2_(n){ return String(n).padStart(2,'0'); }

/** แปลงวันที่จากชีท → ISO 'YYYY-MM-DD' */
function parseSheetDate_(s){
  if (s == null) return null;
  var t = String(s).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(t)) return t; // yyyy-mm-dd
  var m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); // dd/mm/yyyy
  if (m){
    var d = new Date(+m[3], +m[2]-1, +m[1]);
    if (!isNaN(d.getTime())){
      return d.getFullYear() + '-' + pad2_(d.getMonth()+1) + '-' + pad2_(d.getDate());
    }
  }
  return null;
}

/** '#,##0.00' → number */
function parseAmount_(v){
  if (v == null) return 0;
  var s = String(v).replace(/[, ]/g,'');
  var n = Number(s);
  return isFinite(n) ? n : 0;
}

/** ทำให้ชื่อเปรียบเทียบได้เสมอ (ตัดเว้นวรรค/จุด, เป็นตัวพิมพ์เล็ก) */
function canonicalizeName_(s){
  return String(s || '')
    .toLowerCase()
    .replace(/[.\s]+/g,'')
    .trim();
}

/** ===== Data layer: อ่านชีท ===== */
function getDoc_(){
  return SpreadsheetApp.openById(SHEET_DOC_ID);
}
function getFirstSheetByNames_(names){
  var ss = getDoc_();
  for (var i=0;i<names.length;i++){
    var sh = ss.getSheetByName(names[i]);
    if (sh) return sh;
  }
  return null;
}

/** อ่านช่วง A1:AI ของชีท Online → 2D array */
function readOnlineValues_(){
  var sh = getFirstSheetByNames_(ONLINE_CFG.DATA_SHEET_CANDIDATES);
  if (!sh) throw new Error('ไม่พบชีท Online/ช่องทางonline');
  return { values: sh.getRange(ONLINE_CFG.RANGE_A1).getDisplayValues(), sheetName: sh.getName() };
}

/** map คอลัมน์ A..AI → key อังกฤษ ตาม ONLINE_CFG.COL */
function mapRowToObject_(rowArr){
  var letters = Object.keys(ONLINE_CFG.COL); // ['A','B',...,'AI']
  var keys = letters.map(function(L){ return ONLINE_CFG.COL[L]; });
  var obj = {};
  for (var i=0;i<keys.length;i++){
    obj[keys[i]] = (rowArr[i] != null ? rowArr[i] : '');
  }
  return obj;
}

/** คืน rows (raw objects) ไม่ parse */
function readOnlineRowsRaw_(){
  var pack = readOnlineValues_();
  var vals = pack.values;
  if (!vals || vals.length < 2) return [];
  var dataRows = vals.slice(1); // ข้ามแถวหัว
  return dataRows.map(mapRowToObject_);
}

/** ===== Users directory (แท็บ Users) ===== */
function readUsers_(){
  var sh = getFirstSheetByNames_(ONLINE_CFG.USERS_SHEET_CANDIDATES);
  if (!sh) return [];
  var vals = sh.getDataRange().getDisplayValues();
  if (!vals || vals.length < 2) return [];
  var header = vals[0].map(String);
  var idxEmail = header.findIndex(function(h){ return /email/i.test(h); });
  var idxRole  = header.findIndex(function(h){ return /role/i.test(h); });
  var idxName  = header.findIndex(function(h){ return /name/i.test(h); });
  return vals.slice(1).map(function(r){
    return {
      email: String(r[idxEmail] || '').trim().toLowerCase(),
      role: String(r[idxRole] || '').trim() || RBAC.DEFAULT_ROLE,
      name: String(r[idxName] || '').trim()
    };
  });
}

/** ใช้ email ปัจจุบัน (หรือ hint) → หา role/name จาก Users */
function resolveCurrentUser_(emailHint){
  var email = (emailHint || Session.getActiveUser().getEmail() || '').toLowerCase().trim();
  var list = readUsers_();
  var found = list.find(function(x){ return x.email === email; });
  if (!found) return null;
  return {
    id: email,
    email: email,
    displayName: found.name || email,
    role: found.role
  };
}

/** ===== Adapter: raw row → ระเบียนพร้อมใช้ ===== */
function adaptOnlineRow_(r){
  var created = parseSheetDate_(r['created_at']) || null;
  var yyyymm  = String(r['yyyymm'] || '').trim() || (created ? created.slice(0,7).replace('-','') : null);
  var isReal  = String(r['erp_code'] || '').trim() !== ''; // TRUE เมื่อมี ERP code

  // รองรับหัว 'sale_owner' แทน 'sales_owner' เผื่อไฟล์มีสะกดต่าง
  var salesOwner = r['sales_owner'] || r['sale_owner'] || '';

  return {
    row_no: Number(r['row_no']) || null,
    prospect_code: String(r['prospect_code'] || ''),
    erp_code: r['erp_code'] || null,
    created_at: created,
    date_text: r['date_text'] || null,
    yyyymm: yyyymm,
    year: Number(r['year']) || null,
    month: r['month'] || null,
    day: Number(r['day']) || null,
    seq_in_month: Number(r['seq_in_month']) || null,
    company: r['company'] || null,                     // K
    contacts_count: Number(r['contacts_count']) || null,
    area_text: r['area_text'] || null,
    lead_source: r['lead_source'] || null,
    is_new_customer: r['is_new_customer'] || null,
    added_line: r['added_line'] || null,
    admin_owner: r['admin_owner'] || null,
    case_owner: r['case_owner'] || null,
    sales_owner: salesOwner || null,                   // S
    handoff_date: parseSheetDate_(r['handoff_date']),
    items: r['items'] || null,
    amount: parseAmount_(r['amount']),                 // V
    needs: r['needs'] || null,
    status: r['status'] || null,                       // X
    quote_date: parseSheetDate_(r['quote_date']),      // Y
    last_followup_date: parseSheetDate_(r['last_followup_date']), // Z
    last_follower: r['last_follower'] || null,         // AA
    po_date: parseSheetDate_(r['po_date']),            // AB
    payment_term: r['payment_term'] || null,           // AC
    so_number: r['so_number'] || null,                 // AD
    next_followup_date: parseSheetDate_(r['next_followup_date']), // AE
    status_changed_at: r['status_changed_at'] || null, // AF
    owner_changed_at: r['owner_changed_at'] || null,   // AG
    updated_at: parseSheetDate_(r['updated_at']),      // AH
    is_real_customer: isReal                            // AI
  };
}

/** อ่านชีท Online → ระเบียนพร้อมใช้ทั้งหมด */
function fetchOnlineRecords_(){
  var raw = readOnlineRowsRaw_();
  return raw.map(adaptOnlineRow_);
}

/** ===== RBAC: จำกัดการเห็นเฉพาะเซลล์ ===== */
function filterRowsByRBAC_(rows, user){
  if (!user) return [];
  if (String(user.role).toLowerCase() !== 'sales') return rows; // เฉพาะเซลล์เท่านั้นที่ต้องจำกัด
  var me = canonicalizeName_(user.displayName);
  return rows.filter(function(r){
    var v = canonicalizeName_(r[RBAC.VISIBILITY_MATCH_COLUMN] || '');
    return v && (v === me);
  });
}

/** ===== Debug / Health check ===== */

/** 👇 ใส่อีเมลของคุณตรงนี้ถ้า Session ไม่ส่งอีเมลเข้ามา (เช่น 'sale01@yourco.com') */
const DEBUG_EMAIL_HINT = '';

/** ทดสอบ 0.81: อ่าน→แปลง→RBAC แล้วพิมพ์ผลใน Logs */
function test_081_DataAndAccess(){
  // 1) ระบุผู้ใช้ปัจจุบันจากแท็บ Users
  var user = resolveCurrentUser_(DEBUG_EMAIL_HINT);
  if (!user){
    Logger.log('⚠️ ไม่พบผู้ใช้งานในแท็บ Users (ตรวจอีเมล/Role/Name) หรือ Session ไม่ส่งอีเมล');
    Logger.log('👉 วิธีแก้: เติม DEBUG_EMAIL_HINT เป็นอีเมลที่อยู่ในแท็บ Users แล้วรันใหม่');
    return;
  }
  Logger.log('👤 Current user: %s (%s)', user.displayName, user.role);

  // 2) อ่านและแปลงข้อมูลทั้งหมด
  var allRows = fetchOnlineRecords_();
  Logger.log('📦 Total rows (all): %s', allRows.length);

  // 3) กรองตาม RBAC (เฉพาะ sales จะถูกจำกัด)
  var visible = filterRowsByRBAC_(allRows, user);
  Logger.log('🔎 Visible rows (after RBAC): %s', visible.length);

  // 4) แสดงตัวอย่าง 3 แถวแรก
  visible.slice(0,3).forEach(function(r, i){
    Logger.log('#%s prospect=%s | company=%s | owner=%s | amount=%s',
      i+1, r.prospect_code, r.company, r.sales_owner, r.amount);
  });

  Logger.log('✅ 0.81 OK (Data & Access Base)');
}

/** =================================================
 *  0.82 — ENUMUS LOADER (อ่านค่าจากแท็บ Enumus / __enumus_col__X)
 *  Flexible parser: รองรับทั้งหัวคอลัมน์แบบ Field|Key|Color
 *  หรือรูปแบบ A:กลุ่ม, B:ค่า, C:สี (กลุ่มซ้ำ/ค่าว่างได้)
 *  ================================================= */

/** อ่านค่าทั้งชีท Enumus → 2D array */
function readEnumusValues_(){
  var sh = getFirstSheetByNames_(ONLINE_CFG.ENUMUS_SHEET_CANDIDATES);
  if (!sh) return null;
  var vals = sh.getDataRange().getDisplayValues();
  if (!vals || vals.length < 2) return null;
  return vals;
}

/** แปลง Enumus 2D → โครงสร้าง groups: { [group]: { values:Set, colorByKey:{[key]:hex} } } */
function parseEnumus_(vals){
  var header = vals[0].map(function(s){ return String(s||'').trim().toLowerCase(); });
  var body = vals.slice(1);

  // พยายามหา header มาตรฐาน
  var iField = header.findIndex(function(h){ return /^(field|group)$/i.test(h); });
  var iKey   = header.findIndex(function(h){ return /^(key|value|text)$/i.test(h); });
  var iColor = header.findIndex(function(h){ return /^color$/i.test(h); });

  // ถ้าไม่เจอ header แบบชัดเจน ให้ fallback เป็น A:กลุ่ม, B:ค่า, C:สี
  if (iField < 0) iField = 0;
  if (iKey   < 0) iKey   = 1;
  if (iColor < 0) iColor = 2;

  var groups = {}; // { group: { values:Set, colorByKey:{} } }
  var currentGroup = '';

  body.forEach(function(r){
    var g = String(r[iField] || '').trim();
    var k = String(r[iKey]   || '').trim();
    var c = String(r[iColor] || '').trim();

    if (g) currentGroup = g;
    if (!currentGroup) return; // ยังไม่รู้กลุ่ม

    if (!groups[currentGroup]) groups[currentGroup] = { values: new Set(), colorByKey: {} };
    if (k) groups[currentGroup].values.add(k);
    if (k && c) groups[currentGroup].colorByKey[k] = normalizeColorHex_(c);
  });

  // แปลง Set → Array เพื่อ serializable
  var result = {};
  Object.keys(groups).forEach(function(g){
    result[g] = {
      values: Array.from(groups[g].values),
      colorByKey: groups[g].colorByKey
    };
  });
  return result;
}

/** ช่วย normalize สีให้เป็น #RRGGBB ถ้าเป็น rgb()/ชื่อสี จะปล่อยผ่าน */
function normalizeColorHex_(s){
  var t = String(s||'').trim();
  if (!t) return '';
  if (/^#([0-9a-f]{3}|[0-9a-f]{6})$/i.test(t)) {
    // 3 หรือ 6 หลักก็ได้ (ถ้า 3 จะคงไว้)
    if (t.length === 4) {
      // #abc → #aabbcc
      var r = t[1], g = t[2], b = t[3];
      return '#' + r + r + g + g + b + b;
    }
    return t.toUpperCase();
  }
  return t; // เผื่อเป็น 'rgb(…)' หรือชื่อสี
}

/** โหลด Enumus → object พร้อมใช้; ถ้าไม่มีชีท จะคืน {} */
function loadEnumus_(){
  var vals = readEnumusValues_();
  if (!vals) return {};
  return parseEnumus_(vals);
}

/** ดึง meta ของสถานะ: { color, label } (label=ค่าสถานะเดิม) */
function getStatusMeta_(status, enumus){
  enumus = enumus || loadEnumus_();
  var key = String(status||'').trim();
  var color = '';
  // กลุ่มสถานะอาจสะกด 'status' หรือไทย ให้ลองหลายแบบ
  var groupKeys = ['status', 'สถานะ', 'customer_status', 'case_status'];
  for (var i=0;i<groupKeys.length;i++){
    var g = enumus[groupKeys[i]];
    if (g && g.colorByKey && key && g.colorByKey[key]) {
      color = g.colorByKey[key];
      break;
    }
  }
  if (!color) color = '#E0E0E0'; // ค่าเริ่มต้น (เทาอ่อน) เมื่อไม่พบใน Enumus
  return { label: key, color: color };
}

/** คืนชุดค่าทั้งหมดของสถานะ (สำหรับ dropdown/filter ภายหลัง) */
function getAllStatusValues_(enumus){
  enumus = enumus || loadEnumus_();
  var groupKeys = ['status', 'สถานะ', 'customer_status', 'case_status'];
  for (var i=0;i<groupKeys.length;i++){
    var g = enumus[groupKeys[i]];
    if (g && g.values && g.values.length) return g.values;
  }
  return []; // ไม่พบ
}

/** รวมข้อมูลแบบย่อสำหรับ UI (เตรียมใช้ใน 0.83) */
function getCompactRowsForUI_(){
  var enumus = loadEnumus_();
  var rows = fetchOnlineRecords_();
  return rows.map(function(r){
    var owner = r.sales_owner || r.case_owner || '';
    var meta = getStatusMeta_((r.status||''), enumus);
    return {
      prospect_code: r.prospect_code || '',
      company: r.company || '',
      status: r.status || '',
      status_color: meta.color,
      owner: owner,
      updated_at: r.updated_at || ''
    };
  });
}

/** ====== TEST: 0.82 ====== */
function test_082_Enumus(){
  var enumus = loadEnumus_();
  var statuses = getAllStatusValues_(enumus);
  Logger.log('🎨 Enumus loaded. status values: %s', statuses.length);

  var rows = fetchOnlineRecords_();
  var seen = {};
  rows.forEach(function(r){
    var s = r.status || '(blank)';
    if (!seen[s]) {
      var meta = getStatusMeta_(s, enumus);
      Logger.log(' - status: "%s" → color=%s', s, meta.color);
      seen[s] = true;
    }
  });
  Logger.log('✅ 0.82 OK (Enumus Loader & Status meta)');
}
