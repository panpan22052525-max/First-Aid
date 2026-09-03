/**
 * SSKRU Faculty of Nursing - Nursing Physical Examination Web App
 * Production backend for Google Apps Script + Google Sheets
 * Version 6.0
 *
 * วิธีใช้สั้น ๆ
 * 1) แก้ SPREADSHEET_ID หากต้องการใช้ไฟล์ Google Sheet อื่น
 * 2) รัน setupSheet() 1 ครั้งจาก Apps Script editor และอนุญาตสิทธิ์
 * 3) Deploy > New deployment > Web app
 */

const CONFIG = Object.freeze({
  SPREADSHEET_ID: '1ruMMFQOPZfbyuZmsk6W300ZWt95BoFK7QnrHErxmj84',
  SHEET_NAME: 'NursingRecords',
  LOG_NAME: 'AuditLog',
  APP_VERSION: 'v6.0',
  TZ: 'Asia/Bangkok'
});

const FIELDS = Object.freeze([
  'id','createdAt','updatedAt','examDate','hn','name','surname','gender','age',
  'chiefComplaint','presentIllness','pastHistory','familyHistory','healthBehavior',
  'drugAllergy','psychosocial','temp','pr','rr','bp','spo2','loc','weight','height','bmi','bmiClass',
  'heent','heart','lung','abdominal','extremities','impression','treatment','advice','examiner',
  'medicalSuppliesUsage','recordedByName','source','servicePoint','entryChannel','deviceInfo',
  'enteredBy','updatedBy','appVersion','symptomGroup','triageManual','triage','triageReason'
]);

const LOG_FIELDS = Object.freeze(['timestamp','action','recordId','user','payload']);

const TRIAGE_LEVELS = Object.freeze([
  { k:'red',    label:'แดง — วิกฤต/ฉุกเฉิน', en:'Critical / Emergency', color:'#d03b3b', icon:'●' },
  { k:'yellow', label:'เหลือง — เร่งด่วน', en:'Urgent', color:'#fab219', icon:'▲' },
  { k:'green',  label:'เขียว — ไม่รุนแรง/ไม่รีบด่วน', en:'Non-urgent', color:'#0ca30c', icon:'■' }
]);

const LOC_OPTIONS = Object.freeze([
  { v:'A', label:'A — รู้สึกตัวดี ตอบสนองปกติ' },
  { v:'V', label:'V — ซึมลง / สับสน (ตอบสนองต่อเสียง)' },
  { v:'P', label:'P — ตอบสนองต่อความเจ็บเท่านั้น' },
  { v:'U', label:'U — ไม่รู้สึกตัว ไม่ตอบสนอง' }
]);

const TRIAGE_RULES = Object.freeze([
  { lv:'red', f:'loc',  op:'in',  v:['P','U'], t:'ไม่รู้สึกตัว / ตอบสนองต่อความเจ็บเท่านั้น (AVPU = P หรือ U)' },
  { lv:'red', f:'text', op:'kw',  v:['หมดสติ','ไม่รู้สึกตัว','ไม่ตอบสนอง','ชักต่อเนื่อง','ชักซ้ำ','หายใจล้มเหลว','หยุดหายใจ','หายใจเฮือก','เขียวคล้ำ','ช็อก','shock','unconscious','arrest','apnea','cyanosis','gcs'], t:'พบคำบ่งชี้ภาวะคุกคามชีวิต' },
  { lv:'red', f:'rr',   op:'gte', v:30,  t:'RR ≥ 30 ครั้ง/นาที (หอบเหนื่อยรุนแรง)' },
  { lv:'red', f:'rr',   op:'lt',  v:8,   t:'RR < 8 ครั้ง/นาที (หายใจช้าผิดปกติ)' },
  { lv:'red', f:'spo2', op:'lt',  v:90,  t:'SpO₂ < 90% (พร่องออกซิเจน)' },
  { lv:'red', f:'sbp',  op:'lt',  v:90,  t:'SBP < 90 mmHg (สงสัยภาวะช็อก)' },
  { lv:'red', f:'pr',   op:'gt',  v:130, t:'PR > 130 ครั้ง/นาที' },
  { lv:'red', f:'pr',   op:'lt',  v:40,  t:'PR < 40 ครั้ง/นาที' },
  { lv:'red', f:'temp', op:'gte', v:40,  t:'T ≥ 40 °C' },
  { lv:'red', f:'temp', op:'lt',  v:35,  t:'T < 35 °C (อุณหภูมิต่ำผิดปกติ)' },
  { lv:'yellow', f:'loc',  op:'in',  v:['V'], t:'ซึมลง / สับสน (AVPU = V)' },
  { lv:'yellow', f:'text', op:'kw',  v:['เหนื่อย','หายใจลำบาก','หายใจเร็ว','แน่นหน้าอก','เจ็บหน้าอก','ใจสั่น','เวียนศีรษะ','หน้ามืด','อ่อนเพลีย','ซึม','ปลายมือปลายเท้าเย็น','ซีด','dyspnea','chest pain','palpitation','dizziness','pale'], t:'พบอาการที่ต้องประเมินโดยเร็ว' },
  { lv:'yellow', f:'rr',   op:'between', v:[21,29], t:'RR 21–29 ครั้ง/นาที (หายใจเร็ว)' },
  { lv:'yellow', f:'spo2', op:'between', v:[90,94], t:'SpO₂ 90–94%' },
  { lv:'yellow', f:'sbp',  op:'between', v:[90,99], t:'SBP 90–99 mmHg' },
  { lv:'yellow', f:'sbp',  op:'gte', v:180, t:'SBP ≥ 180 mmHg' },
  { lv:'yellow', f:'dbp',  op:'gte', v:110, t:'DBP ≥ 110 mmHg' },
  { lv:'yellow', f:'pr',   op:'between', v:[101,130], t:'PR 101–130 ครั้ง/นาที' },
  { lv:'yellow', f:'pr',   op:'between', v:[40,49],   t:'PR 40–49 ครั้ง/นาที' },
  { lv:'yellow', f:'temp', op:'between', v:[38.5,39.9], t:'T 38.5–39.9 °C (ไข้สูง)' },
  { lv:'yellow', f:'temp', op:'between', v:[35,35.9],   t:'T 35.0–35.9 °C' }
]);

const SYMPTOM_GROUPS = Object.freeze([
  { name:'ระบบทางเดินหายใจ', keywords:['ไอ','เจ็บคอ','น้ำมูก','คัดจมูก','หอบ','หายใจลำบาก','หายใจไม่ออก','wheeze','dyspnea','cough','sore throat'] },
  { name:'ไข้ / โรคติดเชื้อ', keywords:['ไข้','หนาวสั่น','ติดเชื้อ','fever','infection','sepsis'] },
  { name:'ระบบหัวใจ–หลอดเลือด', keywords:['เจ็บหน้าอก','แน่นหน้าอก','ใจสั่น','ความดัน','ความดันโลหิต','hypertension','chest pain','palpitation','heart'] },
  { name:'ระบบทางเดินอาหาร', keywords:['ปวดท้อง','ถ่ายเหลว','ท้องเสีย','อาเจียน','คลื่นไส้','ท้องผูก','abdominal','diarrhea','vomit','nausea'] },
  { name:'ระบบกล้ามเนื้อ–กระดูก', keywords:['ปวดเข่า','ปวดข้อ','ปวดหลัง','กล้ามเนื้อ','ข้อเท้า','sprain','muscle','joint','back pain'] },
  { name:'ระบบประสาท', keywords:['ปวดศีรษะ','เวียนศีรษะ','หน้ามืด','ชัก','ชา','อ่อนแรง','หมดสติ','headache','dizziness','seizure','weakness'] },
  { name:'ต่อมไร้ท่อ / เมตาบอลิก', keywords:['เบาหวาน','น้ำตาล','hypogly','hypergly','diabetes','thyroid'] },
  { name:'อุบัติเหตุ / บาดแผล', keywords:['อุบัติเหตุ','แผล','เลือดออก','ฟกช้ำ','ถลอก','บาด','ล้ม','accident','wound','bleeding','injury'] },
  { name:'ระบบทางเดินปัสสาวะ–ไต', keywords:['ปัสสาวะ','ไต','แสบขัด','บวม','urine','uti','renal','kidney'] },
  { name:'ผิวหนัง', keywords:['ผื่น','คัน','ผิวหนัง','ลมพิษ','rash','itch','urticaria'] },
  { name:'สุขภาพจิต', keywords:['เครียด','นอนไม่หลับ','วิตกกังวล','ซึมเศร้า','stress','insomnia','anxiety','depression'] },
  { name:'อนามัยแม่และเด็ก', keywords:['ตั้งครรภ์','ฝากครรภ์','เด็ก','ทารก','pregnan','antenatal','child','infant'] },
  { name:'ตรวจสุขภาพ / คัดกรอง', keywords:['ตรวจสุขภาพ','คัดกรอง','screening','checkup','check-up'] }
]);

const OTHER_GROUP = 'อื่น ๆ / ไม่ระบุ';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('แบบบันทึกการตรวจร่างกายทางการพยาบาล | SSKRU Faculty of Nursing')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupSheet() {
  const ss = getSpreadsheet_();
  const sh = getOrCreateSheet_(ss, CONFIG.SHEET_NAME);
  const log = getOrCreateSheet_(ss, CONFIG.LOG_NAME);
  ensureHeaders_(sh, FIELDS);
  ensureHeaders_(log, LOG_FIELDS);
  sh.setFrozenRows(1);
  log.setFrozenRows(1);
  try { sh.autoResizeColumns(1, sh.getLastColumn()); } catch (e) {}
  return { ok:true, message:'เชื่อมต่อและเตรียม Google Sheet เรียบร้อย', sheetName:sh.getName(), url:ss.getUrl() };
}

function getConnectionInfo() {
  try {
    const ss = getSpreadsheet_();
    const sh = getOrCreateSheet_(ss, CONFIG.SHEET_NAME);
    ensureHeaders_(sh, FIELDS);
    const rows = Math.max(0, sh.getLastRow() - 1);
    return {
      ok:true,
      fileName:ss.getName(),
      sheetName:sh.getName(),
      url:ss.getUrl(),
      rows:rows,
      user:activeUser_(),
      canEdit:true,
      version:CONFIG.APP_VERSION,
      sources:getDistinctSources_(sh),
      groups:SYMPTOM_GROUPS.map(g => g.name),
      otherGroup:OTHER_GROUP,
      primaryGroups:8,
      triageLevels:TRIAGE_LEVELS,
      triageRules:TRIAGE_RULES,
      locOptions:LOC_OPTIONS,
      serverTime:nowStr_()
    };
  } catch (err) {
    return { ok:false, message:err.message || String(err) };
  }
}

function getSheetUrl() {
  try { return { ok:true, url:getSpreadsheet_().getUrl() }; }
  catch (err) { return { ok:false, message:err.message || String(err) }; }
}

function listRecords() {
  try {
    const sh = getDataSheet_();
    const rows = readAllRecords_(sh);
    rows.sort((a,b) => String(b.createdAt || b.examDate || '').localeCompare(String(a.createdAt || a.examDate || '')));
    return { ok:true, rows:rows, stats:buildStats_(rows) };
  } catch (err) {
    return { ok:false, rows:[], stats:buildStats_([]), message:err.message || String(err) };
  }
}

function saveRecord(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    data = data || {};
    const sh = getDataSheet_();
    const user = activeUser_();
    const now = nowStr_();
    let rec = normalizeRecord_(data);
    rec.bp = normalizeBp_(rec.bp);

    if (!rec.source) return { ok:false, message:'กรุณาเลือกแหล่งที่มาของข้อมูลก่อนบันทึก' };
    if (!rec.name && !rec.surname && !rec.hn) return { ok:false, message:'กรุณากรอกชื่อ นามสกุล หรือ HN อย่างน้อย 1 ช่อง' };
    if (!rec.recordedByName) return { ok:false, message:'กรุณากรอกชื่อผู้บันทึกข้อมูลก่อนบันทึก' };

    rec.appVersion = CONFIG.APP_VERSION;
    rec.symptomGroup = classifySymptom_(rec);
    const tri = classifyTriage_(rec);
    rec.triage = tri.level;
    rec.triageReason = tri.reason;

    const map = headerMap_(sh);
    if (rec.id) {
      const row = findRowById_(sh, rec.id, map);
      if (row < 2) return { ok:false, message:'ไม่พบบันทึกรหัส ' + rec.id };
      const old = rowToObject_(sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0], map);
      rec.createdAt = old.createdAt || now;
      rec.updatedAt = now;
      rec.enteredBy = old.enteredBy || user;
      rec.updatedBy = user;
      writeObjectToRow_(sh, row, rec, map);
      SpreadsheetApp.flush();
      writeLog_('UPDATE', rec.id, rec);
      return { ok:true, mode:'update', id:rec.id, message:'แก้ไขบันทึกเรียบร้อย' };
    }

    rec.id = newId_();
    rec.createdAt = now;
    rec.updatedAt = now;
    rec.enteredBy = user;
    rec.updatedBy = user;
    appendObject_(sh, rec, map);
    SpreadsheetApp.flush();
    writeLog_('CREATE', rec.id, rec);
    return { ok:true, mode:'insert', id:rec.id, message:'บันทึกข้อมูลเรียบร้อย' };
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    return { ok:false, message:'บันทึกไม่สำเร็จ: ' + (err.message || String(err)) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function deleteRecord(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    id = clean_(id);
    if (!id) return { ok:false, message:'ไม่พบรหัสบันทึก' };
    const sh = getDataSheet_();
    const map = headerMap_(sh);
    const row = findRowById_(sh, id, map);
    if (row < 2) return { ok:false, message:'ไม่พบบันทึกรหัส ' + id };
    const old = rowToObject_(sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0], map);
    sh.deleteRow(row);
    SpreadsheetApp.flush();
    writeLog_('DELETE', id, old);
    return { ok:true, id:id, message:'ลบบันทึกเรียบร้อย' };
  } catch (err) {
    return { ok:false, message:'ลบไม่สำเร็จ: ' + (err.message || String(err)) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function systemCheck() {
  try {
    const ss = getSpreadsheet_();
    const sh = getDataSheet_();
    const map = headerMap_(sh);
    const missing = FIELDS.filter(k => !map[k]);
    return {
      ok: missing.length === 0,
      spreadsheet: ss.getName(),
      sheet: sh.getName(),
      rows: Math.max(0, sh.getLastRow() - 1),
      missingHeaders: missing,
      version: CONFIG.APP_VERSION,
      message: missing.length ? 'พบคอลัมน์ที่ยังขาด: ' + missing.join(', ') : 'ระบบพร้อมใช้งาน'
    };
  } catch (err) {
    return { ok:false, message:err.message || String(err), version:CONFIG.APP_VERSION };
  }
}

function testSaveRecord() {
  const stamp = Utilities.formatDate(new Date(), CONFIG.TZ, 'HHmmss');
  const result = saveRecord({
    examDate: Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd'),
    hn: 'TEST-' + stamp,
    name: 'ทดสอบ', surname: 'ระบบจริง', gender: 'หญิง / Female', age: '30',
    source: 'ศูนย์อำนวยการปฐมพยาบาล', servicePoint: 'ทดสอบระบบ',
    recordedByName: 'ผู้ทดสอบระบบ', temp:'36.8', pr:'80', rr:'18', bp:'120/80', spo2:'99', loc:'A',
    chiefComplaint:'ทดสอบการบันทึก Google Sheet', medicalSuppliesUsage:'Gauze 1 ชิ้น',
    entryChannel:'Apps Script Test', deviceInfo:'Server test'
  });
  console.log(JSON.stringify(result));
  return result;
}

function getSpreadsheet_() {
  if (!CONFIG.SPREADSHEET_ID || CONFIG.SPREADSHEET_ID.indexOf('PUT_') === 0) {
    throw new Error('กรุณากำหนด CONFIG.SPREADSHEET_ID ใน Code.gs');
  }
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function getDataSheet_() {
  const ss = getSpreadsheet_();
  const sh = getOrCreateSheet_(ss, CONFIG.SHEET_NAME);
  ensureHeaders_(sh, FIELDS);
  return sh;
}

function ensureHeaders_(sh, required) {
  const lastCol = sh.getLastColumn();
  if (sh.getLastRow() === 0 || lastCol === 0) {
    sh.getRange(1,1,1,required.length).setValues([required]);
    styleHeader_(sh, required.length);
    return;
  }
  const existing = sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(v => String(v).trim());
  let col = lastCol;
  required.forEach(key => {
    if (existing.indexOf(key) === -1) {
      col += 1;
      sh.getRange(1,col).setValue(key);
      existing.push(key);
    }
  });
  styleHeader_(sh, sh.getLastColumn());
}

function styleHeader_(sh, n) {
  try {
    sh.getRange(1,1,1,n).setFontWeight('bold').setBackground('#14663a').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  } catch (e) {}
}

function headerMap_(sh) {
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];
  const map = {};
  headers.forEach((h,i) => { if (h) map[String(h).trim()] = i + 1; });
  return map;
}

function normalizeRecord_(data) {
  const rec = {};
  FIELDS.forEach(k => { rec[k] = clean_(data[k]); });
  return rec;
}


function normalizeBp_(v) {
  const original = clean_(v);
  if (!original) return '';
  const compact = original.replace(/\s+/g, '').replace(/／/g, '/');
  const m = compact.match(/^(\d{2,3})\/(\d{2,3})$/);
  return m ? (m[1] + '/' + m[2]) : original;
}

function clean_(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

function appendObject_(sh, obj, map) {
  const width = sh.getLastColumn();
  const row = new Array(width).fill('');
  Object.keys(obj).forEach(k => { if (map[k]) row[map[k]-1] = obj[k]; });
  sh.appendRow(row);
}

function writeObjectToRow_(sh, rowNum, obj, map) {
  const width = sh.getLastColumn();
  const current = sh.getRange(rowNum,1,1,width).getValues()[0];
  Object.keys(obj).forEach(k => { if (map[k]) current[map[k]-1] = obj[k]; });
  sh.getRange(rowNum,1,1,width).setValues([current]);
}

function rowToObject_(row, map) {
  const o = {};
  Object.keys(map).forEach(k => {
    const v = row[map[k]-1];
    // <input type="date"> ต้องการรูปแบบ YYYY-MM-DD เท่านั้น
    if (k === 'examDate' && v instanceof Date) {
      o[k] = Utilities.formatDate(v, CONFIG.TZ, 'yyyy-MM-dd');
    } else {
      o[k] = displayCell_(v);
    }
  });
  return o;
}

function displayCell_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');
  return v === null || v === undefined ? '' : String(v);
}

function readAllRecords_(sh) {
  if (sh.getLastRow() < 2) return [];
  const map = headerMap_(sh);
  const values = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  return values.map(row => rowToObject_(row,map)).filter(r => clean_(r.id));
}

function findRowById_(sh, id, map) {
  const col = map.id;
  if (!col || sh.getLastRow() < 2) return -1;
  const finder = sh.getRange(2,col,sh.getLastRow()-1,1).createTextFinder(id).matchEntireCell(true).findNext();
  return finder ? finder.getRow() : -1;
}

function getDistinctSources_(sh) {
  if (sh.getLastRow() < 2) return [];
  const map = headerMap_(sh);
  if (!map.source) return [];
  const vals = sh.getRange(2,map.source,sh.getLastRow()-1,1).getDisplayValues().flat();
  const seen = {};
  return vals.map(v => v.trim()).filter(v => v && !seen[v] && (seen[v] = true)).sort();
}

function buildStats_(rows) {
  const src = {}, channel = {};
  let male=0, female=0, bmiSum=0, bmiN=0, abnormal=0, ageSum=0, ageN=0, today=0;
  const todayStr = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd');
  rows.forEach(r => {
    const s = clean_(r.source) || 'ไม่ระบุ'; src[s] = (src[s] || 0) + 1;
    const ch = clean_(r.entryChannel) || 'ไม่ระบุ'; channel[ch] = (channel[ch] || 0) + 1;
    const g = String(r.gender || '');
    if (g.indexOf('หญิง') > -1) female++; else if (g.indexOf('ชาย') > -1) male++;
    const b = parseFloat(r.bmi); if (isFinite(b) && b > 0) { bmiSum += b; bmiN++; if (b < 18.5 || b >= 25) abnormal++; }
    const a = parseFloat(r.age); if (isFinite(a) && a >= 0) { ageSum += a; ageN++; }
    if (String(r.examDate || '').slice(0,10) === todayStr) today++;
  });
  const toList = m => Object.keys(m).map(k => ({label:k,n:m[k]})).sort((a,b) => b.n-a.n || a.label.localeCompare(b.label));
  return {
    total:rows.length, today:today, male:male, female:female, other:rows.length-male-female,
    avgAge:ageN ? (ageSum/ageN).toFixed(0) : '-',
    avgBmi:bmiN ? (bmiSum/bmiN).toFixed(1) : '-', abnormalBmi:abnormal,
    bySource:toList(src), byChannel:toList(channel)
  };
}

function classifySymptom_(rec) {
  const text = [rec.chiefComplaint,rec.presentIllness,rec.impression,rec.treatment].join(' ').toLowerCase();
  for (let i=0; i<SYMPTOM_GROUPS.length; i++) {
    const g = SYMPTOM_GROUPS[i];
    if (g.keywords.some(k => text.indexOf(String(k).toLowerCase()) > -1)) return g.name;
  }
  return OTHER_GROUP;
}

function classifyTriage_(rec) {
  if (['red','yellow','green'].indexOf(rec.triageManual) > -1) {
    return { level:rec.triageManual, reason:'ผู้ตรวจ/พยาบาลกำหนดระดับเอง (Manual triage override)' };
  }
  const bp = String(rec.bp || '').split('/');
  const ctx = {
    loc:clean_(rec.loc).toUpperCase(),
    rr:num_(rec.rr), spo2:num_(rec.spo2), pr:num_(rec.pr), temp:num_(rec.temp),
    sbp:num_(bp[0]), dbp:num_(bp[1]),
    text:[rec.chiefComplaint,rec.presentIllness,rec.impression,rec.psychosocial].join(' ').toLowerCase()
  };
  for (const level of ['red','yellow']) {
    const reasons = [];
    TRIAGE_RULES.filter(r => r.lv === level).forEach(rule => { if (ruleMatch_(rule,ctx)) reasons.push(rule.t); });
    if (reasons.length) return { level:level, reason:reasons.join(' · ') };
  }
  return { level:'green', reason:'สัญญาณชีพและอาการอยู่ในเกณฑ์ไม่เร่งด่วน' };
}

function ruleMatch_(r,c) {
  const x = c[r.f];
  if (r.op === 'kw') return r.v.some(k => c.text.indexOf(String(k).toLowerCase()) > -1);
  if (r.op === 'in') return r.v.indexOf(String(x)) > -1;
  if (x === null || x === undefined || x === '' || !isFinite(Number(x))) return false;
  const n = Number(x);
  if (r.op === 'gte') return n >= Number(r.v);
  if (r.op === 'gt') return n > Number(r.v);
  if (r.op === 'lt') return n < Number(r.v);
  if (r.op === 'between') return n >= Number(r.v[0]) && n <= Number(r.v[1]);
  return false;
}

function num_(v) {
  const n = parseFloat(v);
  return isFinite(n) ? n : null;
}

function newId_() {
  return 'NR' + Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0,6).toUpperCase();
}

function nowStr_() {
  return Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');
}

function activeUser_() {
  try {
    return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'ไม่ระบุ';
  } catch (e) { return 'ไม่ระบุ'; }
}

function writeLog_(action, id, payload) {
  try {
    const ss = getSpreadsheet_();
    const sh = getOrCreateSheet_(ss, CONFIG.LOG_NAME);
    ensureHeaders_(sh, LOG_FIELDS);
    sh.appendRow([nowStr_(), action, id, activeUser_(), JSON.stringify(payload || {})]);
  } catch (e) { console.warn('AuditLog error: ' + e); }
}
