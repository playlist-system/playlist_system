/**
 * หน้ามือถือคนถ่ายรูป — ส่วนหนึ่งของ PHOTO CHECK STUDIO (ตั้งแต่ 5.5.85)
 *
 * - PHOTO CHECK STUDIO ส่งสรุปใบเช็คที่เติมแล้วมาที่หน้านี้ (หน้าผลการเติม → ปุ่ม “ส่งเข้าหน้ามือถือ” → doGet ?recv=studio)
 *   ส่งเฉพาะรายชื่อหนังต่อจอ: ถ่ายแล้ว / ยังไม่ถ่าย / วันที่หนังออก — ไม่ส่งไฟล์ใบเช็ค
 * - เก็บแยกตามคน (ชื่อไฟล์ใบเช็ค เช่น คุณ_A.xlsm → “คุณ_A”) เป็นไฟล์ .json ในโฟลเดอร์ Drive ของเจ้าของสคริปต์
 * - หน้าเว็บมือถือ (doGet): แต่ละคนเห็นเฉพาะใบเช็คที่มีอีเมลตัวเองในแท็บ “ผู้ใช้และสิทธิ์”
 *   เจ้าของชีต และอีเมลในช่อง “ผู้ดูแล / หัวหน้า” เห็นทุกคน และส่งข้อมูลจาก PHOTO CHECK STUDIO ได้
 * - แอปในมือถือ (5.5.86, เปิดได้แม้ไม่มีเน็ต): หน้าแอปอยู่ที่ playlist-system.github.io/playlist_system/phone/ (ไม่มีข้อมูลในหน้า)
 *   แอปขอข้อมูลจากสคริปต์นี้ด้วยลิงก์ส่วนตัว (doGet ?k=รหัส → JSON) แล้วเก็บไว้ในมือถือ
 *   ต้องทำให้ใช้งานได้อีกรายการหนึ่งแบบ “ผู้มีสิทธิ์เข้าถึง: ทุกคน” (ข้อมูลออกได้เฉพาะคนที่มีรหัสของใบนั้น)
 * ไม่มีการใช้ AI ในการทำงาน — เป็นสคริปต์ธรรมดา
 */
const CFG = {
  usersSheet: 'ผู้ใช้และสิทธิ์',
  folderName: 'หน้ามือถือคนถ่ายรูป (ข้อมูล)',   // โฟลเดอร์ใน Drive ที่เก็บข้อมูลของแต่ละคน
  firstRow: 6,                                    // แถวแรกของรายชื่อใบเช็คในแท็บผู้ใช้
  studioOrigin: 'https://playlist-system.github.io',   // เว็บ PHOTO CHECK STUDIO ที่ส่งข้อมูลมาได้
  tz: 'Asia/Bangkok',
  maxPeople: 80,
  appUrl: 'https://playlist-system.github.io/playlist_system/phone/',   // หน้าแอปในมือถือ (ไม่มีข้อมูลในหน้า)
  linkCol: 6,                                     // คอลัมน์ลิงก์แอปส่วนตัวในแท็บผู้ใช้ (F)
};
const P_ = () => PropertiesService.getScriptProperties();

// ===================== เมนูในชีต =====================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('หน้ามือถือ')
    .addItem('ตั้งค่าครั้งแรก', 'setup')
    .addItem('ลิงก์หน้ามือถือ', 'showPhoneLink')
    .addSeparator()
    .addItem('ตั้งค่าแอปในมือถือ (ลิงก์ส่วนตัว)', 'setupApp')
    .addItem('เปลี่ยนลิงก์ส่วนตัวของแถวที่เลือก', 'renewLink')
    .addToUi();
}

// แก้แท็บผู้ใช้ → ล้างแคช ให้สิทธิ์ใหม่มีผลทันที
function onEdit() {
  try { CacheService.getScriptCache().remove('users'); } catch (err) { /* ไม่เป็นไร แคชหมดอายุเองใน 2 นาที */ }
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  P_().setProperty('SHEET_ID', ss.getId());
  usersSheet_(ss);
  folder_();
  CacheService.getScriptCache().remove('users');
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert('หน้ามือถือคนถ่ายรูป', 'ตั้งค่าแล้ว\n\n' + (url
    ? 'ลิงก์หน้ามือถือ:\n' + url + '\n\nวางลิงก์นี้ใน PHOTO CHECK STUDIO (หน้าผลการเติม → ส่งเข้าหน้ามือถือ) แล้วกดส่ง'
    : 'ขั้นต่อไป: กด ทำให้ใช้งานได้ → การทำให้ใช้งานได้รายการใหม่ → เว็บแอป (ดูคู่มือ)'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function showPhoneLink() {
  const url = ScriptApp.getService().getUrl(), ui = SpreadsheetApp.getUi();
  ui.alert('หน้ามือถือคนถ่ายรูป', url
    ? url + '\n\n• วางลิงก์นี้ใน PHOTO CHECK STUDIO ครั้งแรกครั้งเดียว\n• ส่งใน LINE ให้ต่อท้ายด้วย ?openExternalBrowser=1 (เปิดใน Chrome ทันที ไม่ต้องล็อกอินซ้ำ):\n' + url + '?openExternalBrowser=1' +
      '\n• เปิดได้เฉพาะอีเมลบริษัท แต่ละคนเห็นเฉพาะของตัวเอง • อยากให้เปิดได้แม้ไม่มีเน็ต ใช้เมนู ตั้งค่าแอปในมือถือ'
    : 'ยังไม่ได้เผยแพร่เป็นเว็บแอป — ทำตามคู่มือข้อ "เผยแพร่หน้ามือถือ" ก่อน', ui.ButtonSet.OK);
}

function ssOf_() {
  const id = P_().getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('ยังไม่ได้ตั้งค่า — เปิดชีตหน้ามือถือ แล้วกดเมนู หน้ามือถือ → ตั้งค่าครั้งแรก');
  return ss;
}

// ===================== แท็บผู้ใช้และสิทธิ์ =====================
function usersSheet_(ss) {
  let sh = ss.getSheetByName(CFG.usersSheet);
  if (sh) return sh;
  const all = ss.getSheets();
  sh = all.length === 1 && all[0].getLastRow() === 0 ? all[0].setName(CFG.usersSheet) : ss.insertSheet(CFG.usersSheet, 0);
  sh.getRange('A1').setValue('ผู้ใช้และสิทธิ์ — หน้ามือถือคนถ่ายรูป').setFontSize(14).setFontWeight('bold');
  sh.getRange('A2:C2').setValues([['ผู้ดูแล / หัวหน้า', '', '← อีเมลบริษัท คั่นด้วย , • เห็นใบเช็คทุกคน และส่งข้อมูลจาก PHOTO CHECK STUDIO ได้']]);
  sh.getRange('A2').setFontWeight('bold');
  sh.getRange('B2').setBackground('#FFF2CC');
  sh.getRange('A3').setValue('ใส่อีเมลบริษัทของคนถ่ายรูปในช่องสีเหลือง (หลายคนคั่นด้วย ,) • ชื่อใบเช็คขึ้นเองเมื่อส่งข้อมูลจาก PHOTO CHECK STUDIO • เจ้าของชีตเห็นทุกคนเสมอ');
  sh.getRange('A3').setFontColor('#52525B');
  sh.getRange(5, 1, 1, 6).setValues([['ใบเช็ค', 'อีเมลคนถ่ายรูป', 'ส่งล่าสุด', 'จอ', 'ถ่ายแล้ว', 'ลิงก์แอปส่วนตัว (ส่งให้คนนั้นคนเดียว)']]).setFontWeight('bold').setBackground('#EEF2F8');
  [180, 320, 150, 60, 80, 420].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(5);
  return sh;
}

const emails_ = s => String(s || '').split(/[,\s;]+/).map(x => x.trim().toLowerCase()).filter(x => /@/.test(x));

function users_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('users');
  if (hit) { try { return JSON.parse(hit); } catch (err) { /* อ่านจากชีตใหม่ */ } }
  const sh = usersSheet_(ssOf_());
  const n = Math.max(0, sh.getLastRow() - CFG.firstRow + 1);
  const rows = n ? sh.getRange(CFG.firstRow, 1, n, 2).getValues()
    .filter(r => String(r[0]).trim()).map(r => ({ name: String(r[0]).trim(), emails: emails_(r[1]) })) : [];
  const U = { admins: emails_(sh.getRange(2, 2).getValue()), rows, owner: (Session.getEffectiveUser().getEmail() || '').toLowerCase() };
  cache.put('users', JSON.stringify(U), 120);
  return U;
}

/** สิทธิ์ของอีเมลนี้ → { me, admin, names: ใบเช็คที่เปิดดูได้ } */
function access_(email) {
  const U = users_(), me = String(email || '').trim().toLowerCase();
  const admin = !!me && (me === U.owner || U.admins.indexOf(me) >= 0);
  const names = U.rows.filter(r => admin || (me && r.emails.indexOf(me) >= 0)).map(r => r.name);
  return { me, admin, names };
}
const whoAmI_ = () => { try { return Session.getActiveUser().getEmail() || ''; } catch (err) { return ''; } };

// ===================== เก็บข้อมูล (ไฟล์ .json ใน Drive + แคช) =====================
function folder_() {
  const id = P_().getProperty('FOLDER_ID');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (err) { /* สร้างใหม่ */ } }
  const f = DriveApp.createFolder(CFG.folderName);
  P_().setProperty('FOLDER_ID', f.getId());
  return f;
}
const fileKey_ = name => 'F:' + name;
const cacheKey_ = name => 'P:' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, name, Utilities.Charset.UTF_8));

function cachePut_(name, json) {
  try {
    const z = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(json, 'application/json')).getBytes());
    if (z.length < 95000) CacheService.getScriptCache().put(cacheKey_(name), z, 21600);
  } catch (err) { /* ไม่มีแคชก็อ่านจาก Drive */ }
}

function personJson_(name) {
  const hit = CacheService.getScriptCache().get(cacheKey_(name));
  if (hit) {
    try { return Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(hit), 'application/x-gzip')).getDataAsString('UTF-8'); } catch (err) { /* อ่านจาก Drive */ }
  }
  const id = P_().getProperty(fileKey_(name));
  if (!id) return '';
  let json = '';
  try { json = DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8'); } catch (err) { return ''; }
  cachePut_(name, json);
  return json;
}

function savePerson_(p) {
  const json = JSON.stringify(p), key = fileKey_(p.name), id = P_().getProperty(key);
  let file = null;
  if (id) { try { file = DriveApp.getFileById(id); if (file.isTrashed()) file = null; } catch (err) { file = null; } }
  if (file) file.setContent(json);
  else {
    file = folder_().createFile(p.name.replace(/[\\/:*?"<>|]/g, '_') + '.json', json, 'application/json');
    P_().setProperty(key, file.getId());
  }
  cachePut_(p.name, json);
}

// ตรวจข้อมูลที่ส่งมา (มาจากหน้าเว็บ — ไม่เชื่อทั้งหมด ตัดให้อยู่ในรูปที่หน้ามือถือใช้)
function cleanPerson_(x) {
  const s = (v, n) => String(v == null ? '' : v).slice(0, n || 200);
  const num = v => (typeof v === 'number' && isFinite(v)) ? v : 0;
  const name = s(x && x.name, 80).trim();
  if (!name) return null;
  const screens = (Array.isArray(x.screens) ? x.screens : []).slice(0, 400).map(c => ({
    c: s(c.c, 40), p: s(c.p, 60), l: s(c.l, 120), n: num(c.n), s: num(c.s),
    t: (Array.isArray(c.t) ? c.t : []).slice(0, 3000).map(r => [s(r[0]), num(r[1])]),
    d: (Array.isArray(c.d) ? c.d : []).slice(0, 3000).map(r => [s(r[0]), num(r[1])]),
  }));
  return { name, file: s(x.file, 120), at: num(x.at) || Date.now(), mx: num(x.mx), screens };
}

// ===================== รับข้อมูลจาก PHOTO CHECK STUDIO =====================
// ปุ่ม “ส่งเข้าหน้ามือถือ” ในหน้าผลการเติม เปิดหน้านี้ (?recv=studio) → หน้านี้บอก ready → ระบบส่งข้อมูลมาทาง postMessage
// → recvSave เก็บทีละคน + อัปเดตแท็บผู้ใช้ → บอกผลกลับไปที่ระบบ

function recvPage_() {
  let ok = false, me = '';
  try { me = whoAmI_(); ok = access_(me).admin; } catch (err) { ok = false; }
  const d = JSON.stringify({ ok, studio: CFG.studioOrigin, me }).replace(/</g, '\\u003c');
  return HtmlService.createHtmlOutput(RECV_HTML.replace('__DATA__', () => d))
    .setTitle('รับข้อมูลจาก PHOTO CHECK STUDIO').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** เก็บข้อมูลที่ส่งมา {v, at, ver, people:[{name, file, at, mx, screens}]} → สรุปผล */
function recvSave(pay) {
  if (!access_(whoAmI_()).admin) throw new Error('ส่งได้เฉพาะเจ้าของชีต และอีเมลในช่อง “ผู้ดูแล / หัวหน้า” ของแท็บผู้ใช้และสิทธิ์');
  const people = (pay && Array.isArray(pay.people) ? pay.people : []).slice(0, CFG.maxPeople).map(cleanPerson_).filter(Boolean);
  if (!people.length) throw new Error('ไม่ได้รับข้อมูลใบเช็ค');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    people.forEach(savePerson_);
    const sh = usersSheet_(ssOf_());
    const n = Math.max(0, sh.getLastRow() - CFG.firstRow + 1);
    const cur = n ? sh.getRange(CFG.firstRow, 1, n, 2).getValues() : [];
    const rowOf = {};
    cur.forEach((r, i) => { const k = String(r[0]).trim(); if (k && rowOf[k] == null) rowOf[k] = CFG.firstRow + i; });
    const fresh = [];
    let next = CFG.firstRow + cur.length;
    for (const p of people) {
      const N = p.screens.reduce((a, c) => a + c.n, 0), S = p.screens.reduce((a, c) => a + c.s, 0);
      const vals = [new Date(p.at), p.screens.length, N ? S / N : 0];
      let r = rowOf[p.name];
      if (r == null) { r = next++; sh.getRange(r, 1, 1, 2).setValues([[p.name, '']]); sh.getRange(r, 2).setBackground('#FFF2CC'); fresh.push(p.name); }
      sh.getRange(r, 3, 1, 3).setValues([vals]);
      sh.getRange(r, 3).setNumberFormat('d mmm yyyy HH:mm');
      sh.getRange(r, 5).setNumberFormat('0%');
    }
    if (P_().getProperty('APP_ID')) writeLinks_(sh);
    CacheService.getScriptCache().remove('users');
    const U = users_();
    const noMail = U.rows.filter(r => !r.emails.length && people.some(p => p.name === r.name)).map(r => r.name);
    return { people: people.length, screens: people.reduce((a, p) => a + p.screens.length, 0), fresh, noMail };
  } finally {
    lock.releaseLock();
  }
}


// ===================== แอปในมือถือ (ลิงก์ส่วนตัว) =====================
// ลิงก์ส่วนตัว = หน้าแอปบน GitHub Pages + #k=รหัสสุ่ม 64 ตัว&a=รหัสเว็บแอป (อยู่หลัง # จึงไม่ถูกส่งไปที่ GitHub)
// รหัสเก็บใน Script Properties: K:<รหัส> → ชื่อใบเช็ค, T:<ชื่อใบเช็ค> → รหัส • เปลี่ยนลิงก์ = ลบรหัสเดิม (แอปที่ใช้รหัสเดิมลบข้อมูลในเครื่องเอง)

function newToken_() { return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').toLowerCase(); }

function tokenOf_(name, renew) {
  const P = P_(), old = P.getProperty('T:' + name);
  if (old && !renew) return old;
  if (old) P.deleteProperty('K:' + old);
  const t = newToken_();
  P.setProperty('K:' + t, name);
  P.setProperty('T:' + name, t);
  return t;
}

function appLink_(name, renew) {
  const id = P_().getProperty('APP_ID');
  return id ? CFG.appUrl + '?openExternalBrowser=1#k=' + tokenOf_(name, renew) + '&a=' + id : '';
}

/** เขียนลิงก์แอปส่วนตัวในคอลัมน์ F ของทุกแถวที่มีชื่อใบเช็ค */
function writeLinks_(sh) {
  sh.getRange(5, CFG.linkCol).setValue('ลิงก์แอปส่วนตัว (ส่งให้คนนั้นคนเดียว)').setFontWeight('bold').setBackground('#EEF2F8');
  const n = Math.max(0, sh.getLastRow() - CFG.firstRow + 1);
  if (!n) return 0;
  const names = sh.getRange(CFG.firstRow, 1, n, 1).getValues();
  const links = names.map(r => [String(r[0]).trim() ? appLink_(String(r[0]).trim()) : '']);
  sh.getRange(CFG.firstRow, CFG.linkCol, n, 1).setValues(links);
  return links.filter(r => r[0]).length;
}

function setupApp() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('ตั้งค่าแอปในมือถือ',
    'วาง URL ของเว็บแอปที่ทำให้ใช้งานได้แบบ “ผู้มีสิทธิ์เข้าถึง: ทุกคน”\n(คนละรายการกับลิงก์หน้ามือถือเดิม — ดูคู่มือข้อ “แอปในมือถือ”)', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const m = r.getResponseText().trim().match(/^https:\/\/script\.google\.com\/(?:a\/macros\/[^/]+|macros)\/s\/([A-Za-z0-9_-]{20,})\/exec/);
  if (!m) { ui.alert('ตั้งค่าแอปในมือถือ', 'ลิงก์ไม่ถูกต้อง — ต้องขึ้นต้นด้วย https://script.google.com/ และลงท้ายด้วย /exec', ui.ButtonSet.OK); return; }
  const ss = SpreadsheetApp.getActive();
  P_().setProperty('SHEET_ID', ss.getId());
  P_().setProperty('APP_ID', m[1]);
  const n = writeLinks_(usersSheet_(ss));
  ui.alert('ตั้งค่าแอปในมือถือ', 'เรียบร้อย — ลิงก์แอปส่วนตัว ' + n + ' ลิงก์ อยู่ในคอลัมน์ F ของแท็บ “' + CFG.usersSheet + '”\n\n' +
    '• ส่งลิงก์ของแต่ละคนให้คนนั้นทาง LINE แบบแชทส่วนตัว (ไม่ส่งในกลุ่ม)\n• ใครได้ลิงก์ไปก็เห็นใบเช็คนั้น — ถ้าลิงก์หลุด เลือกแถวแล้วกดเมนู เปลี่ยนลิงก์ส่วนตัวของแถวที่เลือก' +
    (n ? '' : '\n\nยังไม่มีชื่อใบเช็ค — ส่งข้อมูลจาก PHOTO CHECK STUDIO ก่อน แล้วลิงก์จะขึ้นเอง'), ui.ButtonSet.OK);
}

function renewLink() {
  const ui = SpreadsheetApp.getUi(), sh = SpreadsheetApp.getActiveSheet(), row = sh.getActiveCell().getRow();
  if (sh.getName() !== CFG.usersSheet || row < CFG.firstRow) { ui.alert('เปลี่ยนลิงก์ส่วนตัว', 'คลิกแถวของใบเช็คในแท็บ “' + CFG.usersSheet + '” ก่อน แล้วกดเมนูนี้อีกครั้ง', ui.ButtonSet.OK); return; }
  const name = String(sh.getRange(row, 1).getValue()).trim();
  if (!name) return;
  if (!P_().getProperty('APP_ID')) { ui.alert('เปลี่ยนลิงก์ส่วนตัว', 'ยังไม่ได้ตั้งค่าแอปในมือถือ — กดเมนู ตั้งค่าแอปในมือถือ ก่อน', ui.ButtonSet.OK); return; }
  if (ui.alert('เปลี่ยนลิงก์ส่วนตัวของ “' + name + '”?', 'ลิงก์เดิมจะใช้ไม่ได้ทันที (แอปที่ใช้ลิงก์เดิมลบข้อมูลในเครื่องเอง) แล้วต้องส่งลิงก์ใหม่ให้คนถ่ายรูป', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  sh.getRange(row, CFG.linkCol).setValue(appLink_(name, true));
  ui.alert('เปลี่ยนลิงก์ส่วนตัว', 'เปลี่ยนแล้ว — ส่งลิงก์ใหม่ในคอลัมน์ F ให้คนถ่ายรูป', ui.ButtonSet.OK);
}

/** แอปขอข้อมูล: ?k=รหัส [&since=เวลาข้อมูลที่มีในเครื่อง] → JSON */
function api_(prm) {
  const out = s => ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
  const bad = (code, err) => out(JSON.stringify({ ok: false, code, err }));
  const k = String(prm.k || '').toLowerCase();
  if (!/^[a-f0-9]{64}$/.test(k)) return bad('key', 'ลิงก์ไม่ถูกต้อง — ขอลิงก์ใหม่จากผู้ดูแล');
  let name, json;
  try {
    name = P_().getProperty('K:' + k);
    if (!name) return bad('key', 'ลิงก์นี้ใช้ไม่ได้แล้ว — ขอลิงก์ใหม่จากผู้ดูแล');
    json = personJson_(name);
  } catch (err) { return bad('err', 'อ่านข้อมูลไม่ได้ — ลองใหม่อีกครั้ง'); }
  if (!json) return out(JSON.stringify({ ok: true, name, data: null }));
  const at = (json.match(/"at":(\d+)/) || [])[1];    // เวลาข้อมูล (ฟิลด์แรกที่ชื่อ at คือของทั้งคน)
  if (prm.since && at && String(prm.since) === at) return out(JSON.stringify({ ok: true, same: true, name }));
  return out('{"ok":true,"name":' + JSON.stringify(name) + ',"data":' + json + '}');
}

// ===================== หน้ามือถือ =====================
function doGet(e) {
  const prm = (e && e.parameter) || {};
  if (prm.recv === 'studio') return recvPage_();
  if (prm.k) return api_(prm);
  let D;
  try {
    const A = access_(whoAmI_());
    const cur = A.names.indexOf(prm.p) >= 0 ? prm.p : (A.names[0] || '');
    const json = cur ? personJson_(cur) : '';
    D = { me: A.me, admin: A.admin, names: A.names, cur, data: json ? JSON.parse(json) : null };
  } catch (err) {
    D = { err: String((err && err.message) || err) };
  }
  const t = HtmlService.createTemplateFromFile('Phone');
  t.boot = JSON.stringify(D).replace(/</g, '\\u003c');
  return t.evaluate().setTitle('งานถ่ายรูป').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** หน้ามือถือเรียกเมื่อผู้ดูแลเลือกดูใบเช็คของคนอื่น */
function getPerson(name) {
  const A = access_(whoAmI_());
  if (A.names.indexOf(name) < 0) throw new Error('เปิดดูใบเช็คนี้ไม่ได้');
  const json = personJson_(name);
  return json ? JSON.parse(json) : null;
}

const RECV_HTML = `<!doctype html><html lang="th"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+Thai:wght@400;500&display=swap" rel="stylesheet">
<style>
:root{--bg:#F6F6F7;--sf:#fff;--fg:#18181B;--tx2:#52525B;--mu:#71717A;--ln:#E4E4E7;--ac:#1F4FD8;--ok:#1B7F4B;--oks:#E8F5EE;--er:#C62828;--ers:#FDECEC}
@media (prefers-color-scheme:dark){:root{--bg:#111113;--sf:#1C1C1F;--fg:#F4F4F5;--tx2:#C4C4CC;--mu:#9A9AA3;--ln:#2E2E33;--ac:#7FA2FF;--ok:#6FD39C;--oks:#16291F;--er:#FF8A80;--ers:#2E1A1A}}
*{box-sizing:border-box;margin:0;padding:0}html,body{background:var(--bg);color:var(--fg);font:15px/1.5 "Noto Sans Thai",system-ui,sans-serif}
.wrap{padding:20px 16px;display:grid;gap:12px}h1{font-size:22px;font-weight:500;line-height:1.3}.sub{font-size:13px;color:var(--mu)}
.step{display:flex;gap:12px;align-items:flex-start;background:var(--sf);border:1px solid var(--ln);border-radius:20px;padding:14px 16px}
.ck{width:22px;height:22px;border-radius:50%;background:var(--ok);color:#fff;font:500 13px/22px system-ui;text-align:center;flex:none;margin-top:1px}
.ck.w{background:var(--ln);color:var(--mu)}.ck.r{background:var(--er)}
.step b{display:block;font-weight:500}.step span{font-size:13px;color:var(--tx2);display:block;overflow-wrap:anywhere}
.msg{border-radius:20px;padding:14px 16px;font-weight:500}.msg small{display:block;font-weight:400;color:var(--tx2);font-size:13px;margin-top:2px}
.ok{background:var(--oks);color:var(--ok)}.bad{background:var(--ers);color:var(--er)}
.spin{width:22px;height:22px;border-radius:50%;border:3px solid var(--ln);border-top-color:var(--ac);animation:sp .9s linear infinite;flex:none;margin-top:1px}
@keyframes sp{to{transform:rotate(360deg)}}@media (prefers-reduced-motion:reduce){.spin{animation:none}}
</style></head><body><div class="wrap"><div><h1>รับข้อมูลจาก PHOTO CHECK STUDIO</h1><div class="sub">หน้ามือถือคนถ่ายรูป</div></div><div id="b" aria-live="polite" style="display:grid;gap:8px"></div></div>
<script>
var D = __DATA__, S = { pay: null, step: 'wait', res: null, err: '' };
function op() { try { return window.top.opener || null; } catch (e) { return null; } }
function tell(m) { var w = op(); if (!w) return; m.psPH = 1; try { w.postMessage(m, D.studio); } catch (e) {} }
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function step(st, t, sub) { var ic = st === 'ok' ? '<span class="ck">✓</span>' : st === 'run' ? '<span class="spin" role="img" aria-label="กำลังทำ"></span>' : st === 'bad' ? '<span class="ck r">!</span>' : '<span class="ck w">·</span>';
  return '<div class="step">' + ic + '<div><b>' + t + '</b>' + (sub ? '<span>' + sub + '</span>' : '') + '</div></div>'; }
function draw() {
  var h = '', p = S.pay && S.pay.people || [], ns = 0;
  p.forEach(function (x) { ns += (x.screens || []).length; });
  h += step(p.length ? 'ok' : S.step === 'err' ? 'bad' : 'run', p.length ? 'รับข้อมูล ' + p.length + ' คน ' + ns + ' จอ' : 'รอข้อมูลจาก PHOTO CHECK STUDIO', p.length ? esc(p.map(function (x) { return x.name; }).join(', ')) : '');
  if (p.length) h += step(S.res ? 'ok' : S.step === 'save' ? 'run' : S.step === 'err' ? 'bad' : 'wait', 'บันทึกให้หน้ามือถือ', S.step === 'save' ? 'อย่าเพิ่งปิดหน้านี้' : '');
  if (S.step === 'done') h += '<div class="msg ok">เสร็จแล้ว<small>คนถ่ายรูปเปิดลิงก์เดิมเห็นข้อมูลใหม่ทันที' +
    (S.res.noMail && S.res.noMail.length ? ' • ยังไม่มีอีเมลคนถ่ายรูปของ ' + esc(S.res.noMail.join(', ')) + ' — ใส่ในแท็บ “ผู้ใช้และสิทธิ์”' : '') + ' • หน้าต่างนี้ปิดเองใน 3 วินาที</small></div>';
  if (S.step === 'err') h += '<div class="msg bad">ส่งไม่สำเร็จ<small>' + esc(S.err) + '</small></div>';
  document.getElementById('b').innerHTML = h;
}
function fail(e) { S.step = 'err'; S.err = (e && e.message) || String(e || 'ไม่ทราบสาเหตุ'); draw(); tell({ t: 'done', ok: false, err: S.err }); }
window.addEventListener('message', function (e) {
  if (e.origin !== D.studio || !e.data || !e.data.psPH || e.data.t !== 'data' || S.pay) return;
  S.pay = e.data.snap || null;
  if (!S.pay || !S.pay.people || !S.pay.people.length) { fail({ message: 'ไม่ได้รับข้อมูลใบเช็ค' }); return; }
  S.step = 'save'; draw();
  google.script.run.withFailureHandler(fail).withSuccessHandler(function (r) {
    S.res = r; S.step = 'done'; draw(); tell({ t: 'done', ok: true, people: r.people, screens: r.screens, fresh: r.fresh, noMail: r.noMail });
  }).recvSave(S.pay);
});
if (!D.ok) { S.step = 'err'; S.err = (D.me ? D.me + ' ' : '') + 'ส่งข้อมูลเข้าหน้ามือถือไม่ได้ — ส่งได้เฉพาะเจ้าของชีต และอีเมลในช่อง “ผู้ดูแล / หัวหน้า” ของแท็บผู้ใช้และสิทธิ์'; draw(); tell({ t: 'done', ok: false, err: S.err }); }
else if (!op()) { S.step = 'err'; S.err = 'เปิดหน้านี้จากปุ่ม “ส่งเข้าหน้ามือถือ” ในหน้าผลการเติมของ PHOTO CHECK STUDIO'; draw(); }
else { draw(); tell({ t: 'ready' }); }
</script></body></html>`;
