/******** WoraCRM – Apply dropdowns & colors from "Enumus" by targets ********
 * Sheet: 
 *   - Online: ตารางหลัก
 *   - Enumus: 
 *       A = enum_type (ชื่อหมวด)
 *       B = enum_value (ค่าที่จะแสดง และ "สีพื้น/ตัวอักษร" ที่ทาไว้จะถูกนำไปใช้)
 *       C = targets (คอลัมน์เป้าหมายใน Online; รับได้ทั้ง "ตัวอักษรคอลัมน์" เช่น X,R,S 
 *                    และ/หรือ "ชื่อหัวคอลัมน์" เช่น สถานะปัจจุบันของลูกค้า; คั่นด้วย , / | หรือช่องว่าง)
 *       แถวว่างคั่นหมวดได้ โค้ดจะข้ามให้
 *
 * Behavior:
 *   - รวมค่าจากหลายหมวดลงคอลัมน์เดียวได้ (ถ้าระบุ targets ซ้ำ)
 *   - สีพื้น/สีอักษร นำมาจากสีของช่อง B ใน Enumus (เฉพาะค่าที่พี่ทาสีไว้)
 *   - M:เขตพื้นที่ → ไม่ตั้ง dropdown (Text)
 *   - W: เปลี่ยนหัวเป็น "ความต้องการ" และเปิด wrap
 ***************************************************************************/

function applyFromEnumusByTargets(){
  var ss = SpreadsheetApp.getActive();
  var online = ss.getSheetByName('Online');
  var enumus = ss.getSheetByName('Enumus');
  if (!online) throw new Error('ไม่พบชีท "Online"');
  if (!enumus) throw new Error('ไม่พบชีท "Enumus"');

  // (0) ปรับหัว W และ M
  _renameHeader(online, 'W', 'ความต้องการ');
  _wrapColumn(online, 'W');
  _clearValidation(online, 'M');                // M เป็น Text
  online.getRange('M2:M').setNumberFormat('@');

  // (1) อ่าน Enumus → ได้รายการ {type -> [items]} และ mapping คอลัมน์เป้าหมายแบบรวม
  var em = _readEnumusWithTargets(enumus); // { types:{t:[{label,bg,fg}]}, byCol:{'X':[labels...]}, colorsByCol:{'X':{label:{bg,fg}}} }

  // (2) Resolve targets ด้วย "ตัวอักษรคอลัมน์" และ/หรือ "ชื่อหัวแถว1"
  var header = online.getRange(1,1,1,online.getLastColumn()).getValues()[0];
  var headToCol = _buildHeaderToColMap(header); // {'สถานะปัจจุบันของลูกค้า':'X', ...}
  var normalized = _normalizeTargets(em.byCol, headToCol); // {'X':[...], 'R':[...], ...} (เฉพาะคอลัมน์ที่มีอยู่จริง)

  // (3) สร้าง helper ต่อ "คอลัมน์" (ไม่ใช่ต่อหมวด) เพราะอาจรวมหลายหมวดเข้าคอลัมน์เดียว
  var helpers = _materializeHelpersPerColumn(ss, normalized); // {'X': Range, 'R': Range, ...}

  // (4) ใช้ Data Validation ตาม helper per column (ยกเว้น M)
  _applyValidationsPerColumn(online, helpers);

  // (5) สร้าง Conditional Formatting ตาม "สี" ที่กำหนดใน Enumus (ต่อคอลัมน์)
  _applyColorsPerColumn(online, normalized, em.colorsByCol);

  // (6) รูปแบบแสดงผลพื้นฐาน
  online.getRange('D2:D').setNumberFormat('yyyy-mm-dd hh:mm');
  online.getRange('E2:E').setNumberFormat('dd/mm/yyyy');
  online.getRange('F2:F').setNumberFormat('@');
  ['T','Y','Z','AB','AE'].forEach(function(c){ online.getRange(c+'2:'+c).setNumberFormat('yyyy-mm-dd'); });
  online.getRange('V2:V').setNumberFormat('#,##0.00');

  // (7) เตือนรหัสโพสเปก + ไฮไลต์ลูกค้าจริง
  var rules = online.getConditionalFormatRules() || [];
  // (option) ไม่ลบกฎเดิมของผู้ใช้ แต่เพิ่มของเรา
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($B2<>"",NOT(REGEXMATCH($B2,"^\\d{9}$")))')
    .setBackground('#F8CBAD').setRanges([ online.getRange('B2:B') ]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$AI2=TRUE')
    .setBackground('#E2F0D9').setRanges([ online.getRange('A2:AI') ]).build());
  // เติมกฎสีจากข้อ (5)
  for (var i=0;i<__ENUM_COL_RULES.length;i++) rules.push(__ENUM_COL_RULES[i]);
  online.setConditionalFormatRules(rules);

  SpreadsheetApp.flush();
}

/******** Helpers ********/

function _renameHeader(sh, col, newName){
  var r = sh.getRange(col+'1'); if (r.getValue() !== newName) r.setValue(newName);
}
function _wrapColumn(sh, col){
  sh.getRange(col+':'+col).setWrap(true).setVerticalAlignment('top');
}
function _clearValidation(sh, col){ sh.getRange(col+'2:'+col).setDataValidation(null); }

/** อ่าน Enumus พร้อม targets ในคอลัมน์ C */
function _readEnumusWithTargets(enumus){
  var rng = enumus.getDataRange();
  var vals = rng.getValues();          // A=type, B=label, C=targets
  var bgs  = rng.getBackgrounds();     // สีพื้นจาก B
  var fcs  = rng.getFontColors();      // สีตัวอักษรจาก B

  var types = {};         // {type: [{label,bg,fg}]}
  var byColRaw = {};      // {'X':[labels...], 'R':[...], 'สถานะปัจจุบันของลูกค้า':[...]} (ยังไม่ normalize)
  var colorsByColRaw = {};// {'X': {label:{bg,fg}} , 'หัวไทย': {...}}

  var currentType = null;
  for (var r=2; r<=rng.getLastRow(); r++){
    var t = (vals[r-1][0]||'').toString().trim();
    var label = (vals[r-1][1]||'').toString().trim();
    var targets = (vals[r-1][2]||'').toString().trim();

    // ข้ามแถวคั่นบล็อค
    if (!t && !label && !targets) continue;

    // ต่อหมวดเดิมถ้า A ว่าง
    if (t) currentType = t;
    if (!currentType) continue; // ยังไม่มีหมวด

    if (label){
      if (!types[currentType]) types[currentType] = [];
      var bg = bgs[r-1][1] || '';
      var fg = fcs[r-1][1] || '';
      types[currentType].push({label:label, bg:bg, fg:fg});
    }

    // targets: อาจว่างในบางแถว (ok) หรือใส่เฉพาะบางรายการก็ได้
    if (targets){
      var parts = _splitTargets(targets); // ['X','R','สถานะปัจจุบันของลูกค้า', ...]
      for (var i=0;i<parts.length;i++){
        var key = parts[i];
        if (!byColRaw[key]) byColRaw[key] = [];
        // รวมค่าของ "หมวด" ปัจจุบันทั้งหมด (ไม่ใช่เฉพาะแถวนี้) เพื่อให้ครบเซ็ต
        var arr = types[currentType] || [];
        for (var j=0;j<arr.length;j++){
          var L = arr[j].label;
          if (byColRaw[key].indexOf(L) < 0) byColRaw[key].push(L);
          // บันทึกสี
          if (!colorsByColRaw[key]) colorsByColRaw[key] = {};
          colorsByColRaw[key][L] = { bg: arr[j].bg, fg: arr[j].fg };
        }
      }
    }
  }
  return { types: types, byCol: byColRaw, colorsByCol: colorsByColRaw };
}

/** แยก targets ด้วย , / | เว้นวรรค */
function _splitTargets(str){
  var s = str.replace(/[|\/]/g, ',');
  var parts = s.split(/[\s,;、，]+/);
  var out=[]; for (var i=0;i<parts.length;i++){ var p = parts[i].trim(); if (p) out.push(p); }
  return out;
}

/** map ชื่อหัวออนไลน์ → คอลัมน์ เช่น {'สถานะปัจจุบันของลูกค้า':'X'} */
function _buildHeaderToColMap(headerArr){
  var map = {};
  for (var c=0;c<headerArr.length;c++){
    var name = (headerArr[c]||'').toString().trim();
    var col = _colLetterFromIndex(c+1);
    if (name) map[name] = col;
  }
  return map;
}

/** แปลง 1..n → A..Z..AA.. */
function _colLetterFromIndex(n){
  var s = '', t = n;
  while (t > 0){
    var m = (t - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    t = Math.floor((t - 1) / 26);
  }
  return s;
}

/** แปลง targets ให้เป็น "ตัวอักษรคอลัมน์ที่มีอยู่จริง" */
function _normalizeTargets(byColRaw, headToCol){
  var online = SpreadsheetApp.getActive().getSheetByName('Online');
  var maxCol = online.getLastColumn();
  var validCols = {};
  for (var i=1;i<=maxCol;i++){ validCols[_colLetterFromIndex(i)] = true; }

  var out = {}; // {'X':[labels...], 'R':[...]}
  for (var key in byColRaw){
    var col = key;
    // ถ้า key เป็นชื่อหัว → แปลงเป็นตัวอักษรคอลัมน์
    if (!validCols[col] && headToCol[key]) col = headToCol[key];
    // ข้ามคอลัมน์ที่ไม่มีจริง
    if (!validCols[col]) continue;
    if (col === 'M') continue; // ยืนยันไม่ map ให้ M

    // รวมค่าที่จะใช้ในคอลัมน์นี้
    var arr = byColRaw[key] || [];
    if (!out[col]) out[col] = [];
    for (var j=0;j<arr.length;j++){
      if (out[col].indexOf(arr[j]) < 0) out[col].push(arr[j]);
    }
  }
  return out;
}

/** helper per "column" (เพราะ 1 คอลัมน์อาจรวมหลายหมวด) */
function _materializeHelpersPerColumn(ss, byCol){
  var helpers = {};
  for (var col in byCol){
    var name = '__enumus_col__' + col;
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.clear(); sh.hideSheet();

    var labels = byCol[col].slice(0);
    if (labels.length === 0) labels = [''];
    var data = [];
    for (var i=0;i<labels.length;i++) data.push([labels[i]]);
    sh.getRange(1,1,data.length,1).setValues(data);
    helpers[col] = sh.getRange(1,1,data.length,1);
  }
  return helpers;
}

function _applyValidationsPerColumn(online, helpers){
  for (var col in helpers){
    var rule = SpreadsheetApp.newDataValidation().requireValueInRange(helpers[col], true).setAllowInvalid(false).build();
    online.getRange(col+'2:'+col).setDataValidation(rule);
  }
}

/** เก็บกฎสีไว้ชั่วคราวก่อน set */
var __ENUM_COL_RULES = [];

function _applyColorsPerColumn(online, byCol, colorsByCol){
  __ENUM_COL_RULES = [];
  for (var col in byCol){
    var labels = byCol[col] || [];
    var palette = colorsByCol[col] || {};
    for (var i=0;i<labels.length;i++){
      var L = labels[i];
      var style = palette[L] || {};
      var bg = (style.bg || '').toLowerCase();
      var fg = (style.fg || '').toLowerCase();

      // ถ้าเป็นสี default (#ffffff / #000000) จะไม่บังคับ
      var rule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(L)
        .setRanges([ online.getRange(col+'2:'+col) ]);
      var styled = false;
      if (bg && bg !== '#ffffff' && bg !== '#fff') { rule = rule.setBackground(style.bg); styled = true; }
      if (fg && fg !== '#000000' && fg !== '#000') { rule = rule.setFontColor(style.fg); styled = true; }
      if (styled) __ENUM_COL_RULES.push(rule.build());
    }
  }
}
