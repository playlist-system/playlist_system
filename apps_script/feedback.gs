/**
 * รับฟีดแบ็กจาก PHOTO CHECK STUDIO (Google Apps Script) — วางในบัญชี Google ของบริษัท
 *
 * ทีมกด "แจ้งปัญหา" ในระบบ → สคริปต์นี้เก็บให้ใน Google Drive ของบัญชีที่ deploy:
 *   - โฟลเดอร์ "ฟีดแบ็ก PHOTO CHECK STUDIO"
 *       ตาราง "คำแจ้งจากทีม"   1 แถวต่อเรื่อง (เลขที่ วันที่ ผู้แจ้ง เรื่อง ใบ/ชีต/จอ รายละเอียดย่อ จำนวนรูป ลิงก์โฟลเดอร์ สถานะ)
 *       โฟลเดอร์ย่อย "#0001 …" ต่อเรื่อง: ข้อความเต็ม.txt (ยาวเท่าไรก็ได้) รูปที่แนบ ไฟล์ใบเช็ค (ถ้าแนบ) ข้อมูลระบบ.txt
 *
 * ตั้งค่าครั้งเดียว (ดู apps_script/FEEDBACK.md):
 *   script.google.com → New project → วางโค้ดนี้ → Deploy → New deployment → Web app
 *   Execute as: Me • Who has access: Anyone → Deploy → อนุญาต → คัดลอก Web app URL ส่งให้คนดูแลระบบใส่ในระบบ
 * ไม่มีค่าใช้จ่าย ข้อมูลอยู่ใน Drive ของบริษัทเท่านั้น
 */

const ROOT_NAME = 'ฟีดแบ็ก PHOTO CHECK STUDIO';
const SHEET_NAME = 'คำแจ้งจากทีม';
const HEAD = ['เลขที่', 'วันที่ส่ง', 'ผู้แจ้ง', 'เรื่อง', 'ใบ / ชีต / จอ', 'รายละเอียด (ย่อ)', 'รูป', 'ไฟล์ใบเช็ค', 'โฟลเดอร์', 'สถานะ', 'เวอร์ชันระบบ'];

/** เปิด Web app URL ในเบราว์เซอร์ = ตรวจว่าพร้อมใช้ */
function doGet() {
  return json_({ ok: true, ready: true, name: 'PHOTO CHECK STUDIO feedback' });
}

/** ระบบส่งคำแจ้งมาที่นี่
 *  - แบบฟอร์ม (ช่อง p = JSON) เปิดในแท็บใหม่ → ตอบเป็นหน้าเว็บ “ส่งเรื่องแล้ว เลขที่ …” (ใช้ได้เมื่อ deploy แบบ “ทุกคนในบริษัท”)
 *  - แบบ JSON ล้วน → ตอบเป็น JSON (deploy แบบ “ทุกคน”) */
function doPost(e) {
  const form = !!(e && e.parameter && e.parameter.p);
  const out = function (o) { return form ? page_(o) : json_(o); };
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const p = JSON.parse(form ? e.parameter.p : ((e && e.postData && e.postData.contents) || '{}'));
    const text = String(p.text || '');
    if (!text.trim() && !(p.images || []).length) return out({ ok: false, err: 'ไม่มีรายละเอียด' });
    let mail = ''; try { mail = Session.getActiveUser().getEmail() || ''; } catch (_) {}
    if (mail) p.name = (p.name ? p.name + ' ' : '') + '(' + mail + ')';
    const root = root_(), sh = sheet_(root);
    const no = '#' + String(sh.getLastRow()).padStart(4, '0'); // แถว 1 = หัวตาราง → เรื่องแรก #0001
    const when = new Date();
    const dir = root.createFolder([no, Utilities.formatDate(when, 'Asia/Bangkok', 'yyyy-MM-dd HHmm'), clean_(p.name)].filter(String).join(' '));
    dir.createFile('ข้อความเต็ม.txt', ['ผู้แจ้ง: ' + (p.name || ''), 'เรื่อง: ' + (p.topic || ''), 'เกี่ยวกับ: ' + (p.where || ''), 'ส่งเมื่อ: ' + when, '', text].join('\n'), MimeType.PLAIN_TEXT);
    let n = 0;
    (p.images || []).forEach(function (im, i) {
      if (!im || !im.data) return;
      dir.createFile(Utilities.newBlob(Utilities.base64Decode(im.data), im.type || 'image/png', clean_(im.name) || ('รูป ' + (i + 1) + '.png')));
      n++;
    });
    const hasFile = !!(p.file && p.file.data);
    if (hasFile) dir.createFile(Utilities.newBlob(Utilities.base64Decode(p.file.data), 'application/vnd.ms-excel.sheet.macroEnabled.12', clean_(p.file.name) || 'ใบเช็ค.xlsm'));
    if (p.sys) dir.createFile('ข้อมูลระบบ.txt', JSON.stringify(p.sys, null, 1), MimeType.PLAIN_TEXT);
    sh.appendRow([no, when, p.name || '', p.topic || '', p.where || '', text.length > 500 ? text.slice(0, 500) + ' …' : text, n, hasFile ? 'มี' : '', dir.getUrl(), 'ใหม่', p.ver || '']);
    return out({ ok: true, no: no });
  } catch (err) {
    return out({ ok: false, err: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/** หน้าเว็บผลการส่ง (แท็บใหม่) • บอกหน้าระบบด้วย ถ้าเบราว์เซอร์ยอม */
function page_(o) {
  const esc = function (t) { return String(t || '').replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  const h = '<div style="font-family:Tahoma,sans-serif;max-width:560px;margin:64px auto;text-align:center;line-height:1.6">' +
    (o.ok ? '<h2 style="color:#1d7546">ส่งเรื่องแล้ว · เลขที่ ' + esc(o.no) + '</h2><p>ขอบคุณครับ คนดูแลระบบจะเห็นในตารางรวมคำแจ้ง<br>ปิดแท็บนี้ได้เลย</p>'
          : '<h2 style="color:#b3261e">ส่งไม่ได้</h2><p>' + esc(o.err) + '</p><p>กลับไปที่หน้าระบบ แล้วกดส่งอีกครั้ง</p>') +
    '</div><script>try{var w=window.top.opener;if(w)w.postMessage({psfb:1,ok:' + (o.ok ? 'true' : 'false') + ',no:' + JSON.stringify(o.no || '') + ',err:' + JSON.stringify(o.err || '') + '},"*")}catch(e){}</script>';
  return HtmlService.createHtmlOutput(h).setTitle(o.ok ? 'ส่งเรื่องแล้ว' : 'ส่งไม่ได้').setFaviconUrl('https://playlist-system.github.io/playlist_system/icons/studio-64.png');
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function clean_(s) { return String(s || '').replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 80); }
function root_() { const it = DriveApp.getFoldersByName(ROOT_NAME); return it.hasNext() ? it.next() : DriveApp.createFolder(ROOT_NAME); }
function sheet_(root) {
  const it = root.getFilesByName(SHEET_NAME);
  if (it.hasNext()) return SpreadsheetApp.open(it.next()).getSheets()[0];
  const ss = SpreadsheetApp.create(SHEET_NAME);
  DriveApp.getFileById(ss.getId()).moveTo(root);
  const sh = ss.getSheets()[0];
  sh.appendRow(HEAD);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold');
  return sh;
}

/** กด Run ฟังก์ชันนี้ครั้งแรก (ก่อน deploy) เพื่อให้ Google ขออนุญาต และสร้างโฟลเดอร์ + ตารางให้ */
function setup() {
  const root = root_(), sh = sheet_(root);
  Logger.log('พร้อมแล้ว: ' + root.getUrl() + ' • ตาราง ' + sh.getParent().getUrl());
}
