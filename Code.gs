function doGet(e) {
  const page = (e && e.parameter.page) ? e.parameter.page : 'index';

  // intercept: /?page=switch → เปิด AccountChooser
  if (page === 'switch') {
    var url = _accountLinks().chooser;
    return HtmlService.createHtmlOutput(
      '<html><head><base target="_top"></head>' +
      '<body>กำลังสลับบัญชี…<script>top.location.replace("'+ url +'");</script></body></html>'
    );
  }

  // intercept: /?page=logout → ออกจากระบบทั้งหมด แล้วกลับมา chooser
  if (page === 'logout') {
    var url2 = _accountLinks().logout;
    return HtmlService.createHtmlOutput(
      '<html><head><base target="_top"></head>' +
      '<body>กำลังออกจากระบบ…<script>top.location.replace("'+ url2 +'");</script></body></html>'
    );
  }

  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle('WoraCRM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}
/** ---------- Auth URL Helpers ---------- */
function _accountLinks(){
  var base = ScriptApp.getService().getUrl(); // Web App URL ปัจจุบัน
  var chooser = 'https://accounts.google.com/AccountChooser?continue=' + encodeURIComponent(base);
  var addsession = 'https://accounts.google.com/AddSession?continue='     + encodeURIComponent(base);
  var logout  = 'https://accounts.google.com/Logout?continue=' + encodeURIComponent(chooser);
  return {
    webapp: base,
    chooser: chooser,            // เลือกบัญชีใหม่
    addsession: addsession,// เพิ่ม/สลับบัญชีแบบไม่ logout
    logout: logout,              // ออกจากระบบ Google แล้วกลับมาเลือกบัญชี
    // เผื่อมีหลายบัญชีอยู่แล้ว อยากจิ้มตามดัชนี
    authuser0: base + '?authuser=0',
    authuser1: base + '?authuser=1',
    authuser2: base + '?authuser=2'
  };
}

// ใช้เรียกจาก client เพื่อดึง URL จริง
function getAccountLinks(){ return _accountLinks(); }

/** ---------- READ Online (A1:AI) + header + สีเซลล์ ---------- */
function readOnlineValues_(){
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Online') || ss.getSheetByName('ช่องทางonline');
  if (!sh) throw new Error('ไม่พบชีท Online/ช่องทางonline');
  var rng = sh.getRange('A1:AI');
  return {
    values: rng.getDisplayValues(),
    bgs:    rng.getBackgrounds(),
    fonts:  rng.getFontColors(),
    sheetName: sh.getName()
  };
}

function mapRowToObject_(rowArr){
  var COL = {A:'row_no',B:'prospect_code',C:'erp_code',D:'created_at',E:'date_text',F:'yyyymm',G:'year',H:'month',I:'day',J:'seq_in_month',K:'company',L:'contacts_count',M:'area_text',N:'lead_source',O:'is_new_customer',P:'added_line',Q:'admin_owner',R:'case_owner',S:'sales_owner',T:'handoff_date',U:'items',V:'amount',W:'needs',X:'status',Y:'quote_date',Z:'last_followup_date',AA:'last_follower',AB:'po_date',AC:'payment_term',AD:'so_number',AE:'next_followup_date',AF:'status_changed_at',AG:'owner_changed_at',AH:'updated_at',AI:'is_real_customer'};
  var letters = Object.keys(COL);
  var keys = letters.map(function(L){ return COL[L]; });
  var obj = {};
  for (var i=0;i<keys.length;i++){ obj[keys[i]] = (rowArr[i] != null ? rowArr[i] : ''); }
  return obj;
}

function readOnlineRowsRaw_(){
  var pack = readOnlineValues_();
  var vals = pack.values, bgs = pack.bgs, fonts = pack.fonts;
  if (!vals || vals.length < 2) return [];
  var dataRows = vals.slice(1);

  var COL = {A:'row_no',B:'prospect_code',C:'erp_code',D:'created_at',E:'date_text',F:'yyyymm',G:'year',H:'month',I:'day',J:'seq_in_month',K:'company',L:'contacts_count',M:'area_text',N:'lead_source',O:'is_new_customer',P:'added_line',Q:'admin_owner',R:'case_owner',S:'sales_owner',T:'handoff_date',U:'items',V:'amount',W:'needs',X:'status',Y:'quote_date',Z:'last_followup_date',AA:'last_follower',AB:'po_date',AC:'payment_term',AD:'so_number',AE:'next_followup_date',AF:'status_changed_at',AG:'owner_changed_at',AH:'updated_at',AI:'is_real_customer'};
  var letters = Object.keys(COL);

  return dataRows.map(function(rowArr, i){
    var o = mapRowToObject_(rowArr);

    // สีทุกคอลัมน์ (K–AI ด้วย, แต่จริงๆครบ A–AI ไปเลย)
    var bgMap={}, fcMap={};
    for (var j=0;j<letters.length;j++){
      var key = COL[letters[j]];
      var bg = (bgs[i+1]   && bgs[i+1][j])   ? bgs[i+1][j]   : '';
      var fc = (fonts[i+1] && fonts[i+1][j]) ? fonts[i+1][j] : '';
      if (bg) bgMap[key] = bg;
      if (fc) fcMap[key] = fc;
    }
    o._bg = bgMap; o._fc = fcMap;
    o._status_bg = bgMap['status'] || '';
    o._status_fc = fcMap['status'] || '';

    return o;
  });
}

function getOnlineHeaderThaiMap_(){
  var pack = readOnlineValues_();
  var header = pack.values[0] || [];
  var COL = {A:'row_no',B:'prospect_code',C:'erp_code',D:'created_at',E:'date_text',F:'yyyymm',G:'year',H:'month',I:'day',J:'seq_in_month',K:'company',L:'contacts_count',M:'area_text',N:'lead_source',O:'is_new_customer',P:'added_line',Q:'admin_owner',R:'case_owner',S:'sales_owner',T:'handoff_date',U:'items',V:'amount',W:'needs',X:'status',Y:'quote_date',Z:'last_followup_date',AA:'last_follower',AB:'po_date',AC:'payment_term',AD:'so_number',AE:'next_followup_date',AF:'status_changed_at',AG:'owner_changed_at',AH:'updated_at',AI:'is_real_customer'};
  var letters = Object.keys(COL), map={};
  for (var i=0;i<letters.length;i++){
    var key = COL[letters[i]];
    map[key] = header[i] || key;
  }
  return map;
}

/** ---------- Utils / Adapter ---------- */
function parseSheetDate_(s){
  if (s == null) return null;
  var t = String(s).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(t)) return t;
  var m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m){
    var d = new Date(+m[3], +m[2]-1, +m[1]);
    if (!isNaN(d.getTime())){
      var mm = ('0'+(d.getMonth()+1)).slice(-2);
      var dd = ('0'+d.getDate()).slice(-2);
      return d.getFullYear() + '-' + mm + '-' + dd;
    }
  }
  return null;
}
function parseAmount_(v){
  if (v == null || v === '') return '';
  var s = String(v).replace(/[, ]/g,'');
  var n = Number(s);
  return isFinite(n) ? n : '';
}
function canonicalizeName_(s){
  return String(s||'').toLowerCase().replace(/[.\s]+/g,'').trim();
}

function adaptOnlineRow_(r){
  var created = parseSheetDate_(r['created_at']) || null;
  var yyyymm  = String(r['yyyymm']||'').trim() || (created ? created.slice(0,7).replace('-','') : null);
  var isReal  = String(r['erp_code']||'').trim() !== '';
  var owner   = r['sales_owner'] || r['sale_owner'] || '';
  return {
    prospect_code: String(r['prospect_code']||''),
    // คอลัมน์ K–AI (ค่า)
    company: r['company'] || '',
    contacts_count: r['contacts_count'] || '',
    area_text: r['area_text'] || '',
    lead_source: r['lead_source'] || '',
    is_new_customer: r['is_new_customer'] || '',
    added_line: r['added_line'] || '',
    admin_owner: r['admin_owner'] || '',
    case_owner: r['case_owner'] || '',
    sales_owner: owner || '',
    handoff_date: parseSheetDate_(r['handoff_date']),
    items: r['items'] || '',
    amount: parseAmount_(r['amount']),
    needs: r['needs'] || '',
    status: r['status'] || '',
    quote_date: parseSheetDate_(r['quote_date']),
    last_followup_date: parseSheetDate_(r['last_followup_date']),
    last_follower: r['last_follower'] || '',
    po_date: parseSheetDate_(r['po_date']),
    payment_term: r['payment_term'] || '',
    so_number: r['so_number'] || '',
    next_followup_date: parseSheetDate_(r['next_followup_date']),
    status_changed_at: parseSheetDate_(r['status_changed_at']),
    owner_changed_at: parseSheetDate_(r['owner_changed_at']),
    updated_at: parseSheetDate_(r['updated_at']),

    // คอลัมน์นอกเหนือ K–AI ที่ยังจำเป็นต่อระบบ
    created_at: created,
    yyyymm: yyyymm,
    // เก็บข้อความดิบจาก E และ Z ไว้ใช้โชว์ตรงๆ
    date_text_raw: r['date_text'] || '',                 // E
    last_followup_date_raw: r['last_followup_date'] || '', // Z

    // สี (ทุกคีย์)
    _bg: r._bg || {},
    _fc: r._fc || {},
    _status_bg: r._status_bg || '',
    _status_fc: r._status_fc || '',
    // ไว้ใช้กรอง
    is_real_customer: (r['is_real_customer'] || '').toString().trim() !== '' ? r['is_real_customer'] : (isReal ? 'TRUE' : '')
  };
}

function fetchOnlineRecords_(){ return readOnlineRowsRaw_().map(adaptOnlineRow_); }

/** ---------- RBAC ---------- */
function getCurrentUser(){
  var sh = SpreadsheetApp.getActive().getSheetByName('Users');
  if (!sh) return { id:'unknown', email:'', displayName:'Viewer', role:'viewer' };
  var vals = sh.getDataRange().getDisplayValues();
  if (!vals || vals.length < 2) return { id:'unknown', email:'', displayName:'Viewer', role:'viewer' };
  var email = (Session.getActiveUser().getEmail()||'').toLowerCase().trim();
  var header = vals[0].map(String);
  var iE = header.findIndex(h=>/email/i.test(h)), iR = header.findIndex(h=>/role/i.test(h)), iN = header.findIndex(h=>/name/i.test(h));
  var rec = vals.slice(1).map(r=>({email:String(r[iE]||'').toLowerCase().trim(), role:String(r[iR]||'').trim()||'viewer', name:String(r[iN]||'').trim()})).find(x=>x.email===email);
  if (!rec) return { id:'unknown', email:'', displayName:'Viewer', role:'viewer' };
  return { id: email, email: email, displayName: rec.name || email, role: rec.role };
}

function filterRowsByRBAC_(rows, user){
  if (!user) return [];
  if (String(user.role).toLowerCase() !== 'sales') return rows; // non-sales เห็นทั้งหมด
  var me = canonicalizeName_(user.displayName);
  return rows.filter(function(r){ return canonicalizeName_(r['sales_owner']||'') === me; });
}

/** ---------- Enumus ---------- */
function readEnumusValues_(){
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Enumus') || ss.getSheetByName('__enumus_col__X');
  if (!sh) return null;
  var vals = sh.getDataRange().getDisplayValues();
  if (!vals || vals.length < 2) return null;
  return vals;
}
function normalizeColorHex_(s){
  var t = String(s||'').trim();
  if (!t) return '';
  if (/^#([0-9a-f]{3}|[0-9a-f]{6})$/i.test(t)){
    if (t.length === 4){ var r=t[1],g=t[2],b=t[3]; return ('#'+r+r+g+g+b+b).toUpperCase(); }
    return t.toUpperCase();
  }
  return t;
}
function parseEnumus_(vals){
  var header = vals[0].map(function(s){ return String(s||'').trim().toLowerCase(); });
  var body = vals.slice(1);
  var iField = header.findIndex(function(h){ return /^(field|group)$/i.test(h); });
  var iKey   = header.findIndex(function(h){ return /^(key|value|text)$/i.test(h); });
  var iColor = header.findIndex(function(h){ return /^color$/i.test(h); });
  if (iField<0) iField=0; if (iKey<0) iKey=1; if (iColor<0) iColor=2;
  var groups={}, current='';
  body.forEach(function(r){
    var g=String(r[iField]||'').trim(), k=String(r[iKey]||'').trim(), c=String(r[iColor]||'').trim();
    if (g) current=g;
    if (!current) return;
    if (!groups[current]) groups[current]={ values:new Set(), colorByKey:{} };
    if (k) groups[current].values.add(k);
    if (k && c) groups[current].colorByKey[k] = normalizeColorHex_(c);
  });
  var out={}; Object.keys(groups).forEach(function(g){
    out[g]={ values:Array.from(groups[g].values), colorByKey:groups[g].colorByKey };
  });
  return out;
}
function loadEnumus_(){ var v=readEnumusValues_(); return v ? parseEnumus_(v) : {}; }

function getAllStatusValues_(enumus){
  enumus = enumus || loadEnumus_();
  var gk=['status','สถานะ','สถานะปัจจุบันของลูกค้า','customer_status','case_status'];
  for (var i=0;i<gk.length;i++){ var g=enumus[gk[i]]; if (g && g.values && g.values.length) return g.values; }
  return [];
}
function getStatusMeta_(status, enumus){
  enumus = enumus || loadEnumus_();
  var key = String(status||'').trim();
  var color = '';
  var groupKeys = ['status', 'สถานะ', 'สถานะปัจจุบันของลูกค้า', 'customer_status', 'case_status'];
  for (var i=0;i<groupKeys.length;i++){
    var g = enumus[groupKeys[i]];
    if (g && g.colorByKey && key && g.colorByKey[key]) { color = g.colorByKey[key]; break; }
  }
  if (!color) color = '#3a3a3a';
  return { label: key, color: color };
}

/** ---------- Sort / Filters ---------- */
function sortByBasis_(rows, basis, dir){
  var b = basis || 'updated_at';
  var asc = String(dir||'desc').toLowerCase() === 'asc';
  function padNum(s){ s = String(s||''); return ('00000000000000000000'+s).slice(-20); }
  function normDate(v, isAsc){
    if (!v) return isAsc ? '9999-12-31' : '0000-00-00';
    return String(v);
    }
  return rows.slice().sort(function(a,b2){
    var va, vb;
    if (b === 'prospect_code'){
      va = padNum(a.prospect_code); vb = padNum(b2.prospect_code);
    } else if (/_at$|date/i.test(b)) {
      va = normDate(a[b],  asc);
      vb = normDate(b2[b], asc);
    } else {
      va = String(a[b]||''); vb = String(b2[b]||'');
    }
    if (va === vb) return 0;
    return asc ? (va > vb ? 1 : -1) : (va > vb ? -1 : 1);
  });
}
function unique_(arr){ return Array.from(new Set(arr.filter(function(x){ return x!=null && x!==''; }))); }
function getAvailableDateFields_(rows){
  var fields=['updated_at','created_at','quote_date','next_followup_date'];
  return fields.filter(function(f){ return rows.some(function(r){ return r[f]; }); });
}

/** ---------- API: Compact ---------- */
function getCompactMeta(){
  var header = getOnlineHeaderThaiMap_();
  var enumus = loadEnumus_();
  var user = getCurrentUser();
  var visible = filterRowsByRBAC_(fetchOnlineRecords_(), user);

  var statuses = getAllStatusValues_(enumus);
  if (!statuses.length) statuses = unique_(visible.map(function(r){ return r.status; }));

  var owners = unique_(visible.map(function(r){ return r.sales_owner || r.case_owner || ''; })).sort();
  var yyyymmValues = unique_(visible.map(function(r){ return r.yyyymm; })).sort().reverse();
  var dateFields = getAvailableDateFields_(visible);

  return { header, statuses, owners, yyyymmValues, dateFields, user };
}

function getCompactRows(filters){
  var enumus = loadEnumus_();
  var user = getCurrentUser();
  var visible = filterRowsByRBAC_(fetchOnlineRecords_(), user);
  var f = filters || {};
  var basis=f.dateBasis||null, from=f.from||null, to=f.to||null;

  var filtered = visible.filter(function(r){
    if (f.yyyymm && r.yyyymm !== f.yyyymm) return false;

    // สถานะ: เลือกแล้วต้องตรง และตัดช่องว่าง
    if (f.status && f.status.length){
      if (!r.status) return false;
      if (f.status.indexOf(r.status) === -1) return false;
    }
    if (f.owner && f.owner.length){
      var own = r.sales_owner || r.case_owner || '';
      if (f.owner.indexOf(own) === -1) return false;
    }
    if (typeof f.isReal === 'boolean'){
      if (!!r.is_real_customer !== f.isReal) return false;
    }
    if (basis && (from||to)){
      var v = r[basis];
      if (!v) return false;
      if (from && v < from) return false;
      if (to && v > to) return false;
    }
    return true;
  });

  var sorted = sortByBasis_(filtered, f.sortBasis || 'updated_at', f.sortDir || 'desc');

  return sorted.map(function(r){
    var meta  = getStatusMeta_(r.status||'', enumus);
    var owner = r.sales_owner || r.case_owner || '';
    var bg    = r._status_bg || meta.color || '#3a3a3a';
    var fc    = r._status_fc || '';
    return {
      prospect_code: r.prospect_code||'',
      company:       r.company||'',
      status:        r.status||'',
      status_bg:     bg,
      status_fc:     fc,
      owner:         owner,
      updated_at:    r.updated_at||'',
      yyyymm:        r.yyyymm||''
    };
  });
}

/** ---------- API: Modal (รายละเอียดลูกค้า) ---------- */
function getLeadDetail(code){
  if (!code) throw new Error('missing prospect_code');
  var enumus = loadEnumus_(), header = getOnlineHeaderThaiMap_(), user = getCurrentUser();
  var r = filterRowsByRBAC_(fetchOnlineRecords_(), user).find(x => (x.prospect_code||'')===String(code));
  if (!r) throw new Error('ไม่พบรายการ');

  var meta = getStatusMeta_(r.status||'', enumus);
  var status_bg = r._status_bg || meta.color || '#3a3a3a';
  var status_fc = r._status_fc || '';

  // helper ใส่สีจากชีทจริงให้ field
  function kv(key, value){ return [key, value, (r._bg||{})[key]||'', (r._fc||{})[key]||'']; }

  // หมวดข้อมูล (แสดงครบถ้วนเท่าที่ไม่ใช่ field วันที่) — ตัด AI: ลูกค้าจริง? ออก
  var general = [
    kv('company', r.company),
    kv('contacts_count', r.contacts_count),
    kv('area_text', r.area_text),
    kv('lead_source', r.lead_source),
    kv('is_new_customer', r.is_new_customer),
    kv('added_line', r.added_line)
  ];
  var owners  = [
    kv('admin_owner', r.admin_owner),
    kv('case_owner',  r.case_owner),
    kv('sales_owner', r.sales_owner)
  ];
  var sales   = [
    kv('items',        r.items),
    kv('amount',       r.amount),
    kv('payment_term', r.payment_term),
    kv('so_number',    r.so_number)
  ];
  var tracking = [
    kv('last_follower', r.last_follower)
    // (วันอื่นๆ เก็บไว้ใน data แต่ "ไม่แสดง" ตามสเป็ก)
  ];
  var notes   = [ kv('needs', r.needs) ];

  // ใต้หัว: E และ Z เท่านั้น
  var created_display = r.date_text_raw || '';          // E
  var updated_display = r.last_followup_date_raw || ''; // Z

  return {
    header,
    prospect_code: r.prospect_code,
    status_label: r.status || '',
    status_bg, status_fc,
    created_at: created_display,
    updated_at: updated_display,
    sections: { general, owners, sales, tracking, notes },
    data: r
  };
}
