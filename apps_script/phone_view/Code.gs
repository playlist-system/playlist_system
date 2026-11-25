/**
 * หน้ามือถือคนถ่ายรูป — ส่วนหนึ่งของ PHOTO CHECK STUDIO (ตั้งแต่ 5.5.85)
 *
 * - PHOTO CHECK STUDIO ส่งสรุปใบเช็คมาที่หน้านี้ (หน้างานโฟลเดอร์และรูป → ขั้นที่ 4 “ส่งเข้าหน้ามือถือ” → doGet ?recv=studio)
 *   ส่งเฉพาะรายชื่อหนังต่อจอ (รหัส • TYPE • ชื่อจอ): ถ่ายแล้ว / ยังไม่ถ่าย / วันที่หนังออก — ไม่ส่งไฟล์ใบเช็ค
 * - เก็บแยกตามคน (ชื่อไฟล์ใบเช็ค เช่น คุณ_A.xlsm → “คุณ_A”) เป็นไฟล์ .json ในโฟลเดอร์ Drive ของเจ้าของสคริปต์
 * - หน้าเว็บมือถือ (doGet) — สิทธิ์ตามแท็บ “ผู้ใช้และสิทธิ์”:
 *   หัวหน้า (แถว 2) และเจ้าของชีต = เห็นทุกใบ ส่งได้ทุกใบ • คนคัดรูป (คอลัมน์ C) = เห็นและส่งได้เฉพาะใบที่มีอีเมลตัวเอง
 *   คนถ่ายรูป (คอลัมน์ B) = เห็นเฉพาะใบของตัวเอง ดูอย่างเดียว
 * - 5.5.87: เมนูลิงก์หน้ามือถือให้ลิงก์สำหรับ LINE (?openExternalBrowser=1) • หน้ามือถือไม่โหลดฟอนต์จากเน็ต (เปิดเร็วขึ้น)
 * ไม่มีการใช้ AI ในการทำงาน — เป็นสคริปต์ธรรมดา
 */
const CFG = {
  usersSheet: 'ผู้ใช้และสิทธิ์',
  folderName: 'หน้ามือถือคนถ่ายรูป (ข้อมูล)',   // โฟลเดอร์ใน Drive ที่เก็บข้อมูลของแต่ละคน
  firstRow: 6,                                    // แถวแรกของรายชื่อใบเช็คในแท็บผู้ใช้ (A ใบเช็ค • B คนถ่ายรูป • C คนคัดรูป • D ส่งล่าสุด • E จอ • F ถ่ายแล้ว)
  studioOrigin: 'https://playlist-system.github.io',   // เว็บ PHOTO CHECK STUDIO ที่ส่งข้อมูลมาได้
  tz: 'Asia/Bangkok',
  maxPeople: 80,
};
const P_ = () => PropertiesService.getScriptProperties();

// ===================== เมนูในชีต =====================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('หน้ามือถือ')
    .addItem('ตั้งค่าครั้งแรก', 'setup')
    .addItem('ลิงก์หน้ามือถือ', 'showPhoneLink')
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
    ? 'ลิงก์หน้ามือถือ:\n' + url + '\n\nวางลิงก์นี้ใน PHOTO CHECK STUDIO (หน้างานโฟลเดอร์และรูป → ขั้นที่ 4 ส่งเข้าหน้ามือถือ) แล้วกดส่ง'
    : 'ขั้นต่อไป: กด ทำให้ใช้งานได้ → การทำให้ใช้งานได้รายการใหม่ → เว็บแอป (ดูคู่มือ)'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function showPhoneLink() {
  const url = ScriptApp.getService().getUrl(), ui = SpreadsheetApp.getUi();
  ui.alert('หน้ามือถือคนถ่ายรูป', url
    ? url + '\n\n• วางลิงก์นี้ใน PHOTO CHECK STUDIO ครั้งแรกครั้งเดียว\n• ส่งใน LINE ให้ใช้ลิงก์นี้ (ต่อท้ายด้วย ?openExternalBrowser=1 — LINE เปิดใน Chrome ทันที ไม่ต้องล็อกอินซ้ำ):\n' + url + '?openExternalBrowser=1' +
      '\n• เปิดได้เฉพาะอีเมลบริษัท แต่ละคนเห็นเฉพาะของตัวเอง'
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
  sh.getRange('A2:C2').setValues([['หัวหน้า', '', '← อีเมลบริษัท คั่นด้วย , • เห็นใบเช็คทุกคน และส่งข้อมูลได้ทุกใบ']]);
  sh.getRange('A2').setFontWeight('bold');
  sh.getRange('B2').setBackground('#FFF2CC');
  sh.getRange('A3').setValue('ใส่อีเมลบริษัทในช่องสีเหลือง (หลายคนคั่นด้วย ,) • คนถ่ายรูป = ดูได้อย่างเดียว • คนคัดรูป = ดูและส่งข้อมูลได้เฉพาะใบนั้น • ชื่อใบเช็คขึ้นเองเมื่อหัวหน้าส่งข้อมูลครั้งแรก • เจ้าของชีตเห็นทุกคนเสมอ');
  sh.getRange('A3').setFontColor('#52525B');
  sh.getRange(5, 1, 1, 6).setValues([['ใบเช็ค', 'อีเมลคนถ่ายรูป (ดูอย่างเดียว)', 'อีเมลคนคัดรูป (ดู + ส่งได้)', 'ส่งล่าสุด', 'จอ', 'ถ่ายแล้ว']]).setFontWeight('bold').setBackground('#EEF2F8');
  [180, 300, 300, 150, 60, 80].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(5);
  return sh;
}

const emails_ = s => String(s || '').split(/[,\s;]+/).map(x => x.trim().toLowerCase()).filter(x => /@/.test(x));

function users_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('users');
  if (hit) { try { return JSON.parse(hit); } catch (err) { /* อ่านจากชีตใหม่ */ } }
  const sh = usersSheet_(ssOf_());
  const n = Math.max(0, sh.getLastRow() - CFG.firstRow + 1);
  const rows = n ? sh.getRange(CFG.firstRow, 1, n, 3).getValues()
    .filter(r => String(r[0]).trim()).map(r => ({ name: String(r[0]).trim(), emails: emails_(r[1]), checkers: emails_(r[2]) })) : [];
  const U = { admins: emails_(sh.getRange(2, 2).getValue()), rows, owner: (Session.getEffectiveUser().getEmail() || '').toLowerCase() };
  cache.put('users', JSON.stringify(U), 120);
  return U;
}

/** สิทธิ์ของอีเมลนี้ → { me, admin, names: ใบเช็คที่เปิดดูได้, send: ใบเช็คที่ส่งข้อมูลได้ (หัวหน้า = ทุกใบ) } */
function access_(email) {
  const U = users_(), me = String(email || '').trim().toLowerCase();
  const admin = !!me && (me === U.owner || U.admins.indexOf(me) >= 0);
  const mine = r => !!me && (r.emails.indexOf(me) >= 0 || (r.checkers || []).indexOf(me) >= 0);
  const names = U.rows.filter(r => admin || mine(r)).map(r => r.name);
  const send = U.rows.filter(r => admin || (!!me && (r.checkers || []).indexOf(me) >= 0)).map(r => r.name);
  return { me, admin, names, send };
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
    c: s(c.c, 40), y: s(c.y, 20), p: s(c.p, 60), l: s(c.l, 120), n: num(c.n), s: num(c.s),
    t: (Array.isArray(c.t) ? c.t : []).slice(0, 3000).map(r => [s(r[0]), num(r[1])]),
    d: (Array.isArray(c.d) ? c.d : []).slice(0, 3000).map(r => [s(r[0]), num(r[1]), num(r[2])]),
  }));
  return { name, file: s(x.file, 120), at: num(x.at) || Date.now(), mx: num(x.mx), screens };
}

// ===================== รับข้อมูลจาก PHOTO CHECK STUDIO =====================
// ปุ่ม “ส่งเข้าหน้ามือถือ” ในหน้าผลการเติม เปิดหน้านี้ (?recv=studio) → หน้านี้บอก ready → ระบบส่งข้อมูลมาทาง postMessage
// → recvSave เก็บทีละคน + อัปเดตแท็บผู้ใช้ → บอกผลกลับไปที่ระบบ

function recvPage_() {
  let ok = false, me = '';
  try { me = whoAmI_(); const A = access_(me); ok = A.admin || A.send.length > 0; } catch (err) { ok = false; }
  const d = JSON.stringify({ ok, studio: CFG.studioOrigin, me }).replace(/</g, '\\u003c');
  return HtmlService.createHtmlOutput(RECV_HTML.replace('__DATA__', () => d))
    .setTitle('รับข้อมูลจาก PHOTO CHECK STUDIO').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** เก็บข้อมูลที่ส่งมา {v, at, ver, people:[{name, file, at, mx, screens}]} → สรุปผล */
function recvSave(pay) {
  const A = access_(whoAmI_());
  if (!A.admin && !A.send.length) throw new Error('ส่งได้เฉพาะเจ้าของชีต หัวหน้า และคนคัดรูป (อีเมลในแท็บผู้ใช้และสิทธิ์)');
  const all = (pay && Array.isArray(pay.people) ? pay.people : []).slice(0, CFG.maxPeople).map(cleanPerson_).filter(Boolean);
  if (!all.length) throw new Error('ไม่ได้รับข้อมูลใบเช็ค');
  // คนคัดรูปส่งได้เฉพาะใบที่มีอีเมลตัวเองในคอลัมน์ C • ใบที่ยังไม่มีในแท็บ (ชื่อใหม่) หัวหน้าส่งครั้งแรก
  const people = A.admin ? all : all.filter(p => A.send.indexOf(p.name) >= 0);
  const skipped = all.filter(p => people.indexOf(p) < 0).map(p => p.name);
  if (!people.length) throw new Error('ใบเช็คที่ส่งมา (' + skipped.join(', ') + ') ไม่มีอีเมลของคุณในช่องคนคัดรูป — ให้หัวหน้าใส่อีเมลในแท็บผู้ใช้และสิทธิ์ หรือให้หัวหน้าส่ง');
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
      if (r == null) { r = next++; sh.getRange(r, 1, 1, 3).setValues([[p.name, '', '']]); sh.getRange(r, 2, 1, 2).setBackground('#FFF2CC'); fresh.push(p.name); }
      sh.getRange(r, 4, 1, 3).setValues([vals]);
      sh.getRange(r, 4).setNumberFormat('d mmm yyyy HH:mm');
      sh.getRange(r, 6).setNumberFormat('0%');
    }
    CacheService.getScriptCache().remove('users');
    const U = users_();
    const noMail = U.rows.filter(r => !r.emails.length && people.some(p => p.name === r.name)).map(r => r.name);
    return { people: people.length, screens: people.reduce((a, p) => a + p.screens.length, 0), fresh, noMail, skipped };
  } finally {
    lock.releaseLock();
  }
}

// ===================== หน้ามือถือ =====================
function doGet(e) {
  const prm = (e && e.parameter) || {};
  if (prm.recv === 'studio') return recvPage_();
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
    (S.res.noMail && S.res.noMail.length ? ' • ยังไม่มีอีเมลคนถ่ายรูปของ ' + esc(S.res.noMail.join(', ')) + ' — ใส่ในแท็บ “ผู้ใช้และสิทธิ์”' : '') +
    (S.res.skipped && S.res.skipped.length ? ' • ไม่ได้ส่ง ' + esc(S.res.skipped.join(', ')) + ' (ไม่ใช่ใบที่คุณเป็นคนคัดรูป)' : '') + ' • หน้าต่างนี้ปิดเองใน 3 วินาที</small></div>';
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
    S.res = r; S.step = 'done'; draw(); tell({ t: 'done', ok: true, people: r.people, screens: r.screens, fresh: r.fresh, noMail: r.noMail, skipped: r.skipped });
  }).recvSave(S.pay);
});
if (!D.ok) { S.step = 'err'; S.err = (D.me ? D.me + ' ' : '') + 'ส่งข้อมูลเข้าหน้ามือถือไม่ได้ — ส่งได้เฉพาะเจ้าของชีต หัวหน้า และคนคัดรูป (อีเมลในแท็บผู้ใช้และสิทธิ์)'; draw(); tell({ t: 'done', ok: false, err: S.err }); }
else if (!op()) { S.step = 'err'; S.err = 'เปิดหน้านี้จากปุ่ม “ส่งเข้าหน้ามือถือ” (หน้างานโฟลเดอร์และรูป ขั้นที่ 4) ของ PHOTO CHECK STUDIO'; draw(); }
else { draw(); tell({ t: 'ready' }); }
</script></body></html>`;
