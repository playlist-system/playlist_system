// จำลอง apps_script/feedback.gs ด้วยบริการ Google ปลอม (ไม่ต่ออินเทอร์เน็ต)
//   node dev/tests/feedback_gs_sim.js apps_script/feedback.gs
// ตรวจ: ส่งแบบฟอร์ม → หน้าเว็บ "ส่งเรื่องแล้ว" + postMessage • ส่งแบบ JSON → {"ok":true} • ข้อความว่าง → "ส่งไม่ได้"
//       บันทึกแถวพร้อมอีเมลผู้แจ้ง และเก็บไฟล์ข้อความเต็ม + รูป
const fs = require('fs');
const src = fs.readFileSync(process.argv[2], 'utf8');
const files = [], rows = [];
const folder = name => ({ name, getUrl: () => 'https://drive/' + name, createFolder: n => folder(n),
  createFile: a => { files.push(typeof a === 'string' ? a : a.name); return {} }, getFilesByName: () => ({ hasNext: () => false }) });
const sheet = { getLastRow: () => rows.length + 1, appendRow: r => rows.push(r), setFrozenRows() {},
  getRange: () => ({ setFontWeight() {} }), getParent: () => ({ getUrl: () => 'u' }) };
const ctx = {
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  DriveApp: { getFoldersByName: () => ({ hasNext: () => false }), createFolder: n => folder(n), getFileById: () => ({ moveTo() {} }) },
  SpreadsheetApp: { create: () => ({ getId: () => 'x', getSheets: () => [sheet] }), open: () => ({ getSheets: () => [sheet] }) },
  Session: { getActiveUser: () => ({ getEmail: () => 'tester@example.com' }) },
  Utilities: { formatDate: () => '2026-01-01 0900', newBlob: (b, t, n) => ({ name: n }), base64Decode: s => Buffer.from(s, 'base64') },
  HtmlService: { createHtmlOutput: h => ({ h, setTitle(t) { this.t = t; return this }, setFaviconUrl(u) { this.icon = u; return this } }) },
  ContentService: { createTextOutput: t => ({ t, setMimeType() { return this } }), MimeType: { JSON: 'json' } },
  MimeType: { PLAIN_TEXT: 'text' },
};
const fn = new Function(...Object.keys(ctx), src + ';return {doPost,doGet};');
const { doPost } = fn(...Object.values(ctx));
const body = { name: 'คุณทดสอบ', topic: 'ไม่เติม', where: 'ตัวอย่าง.xlsm', text: 'ทดสอบ',
  images: [{ name: 'a.png', type: 'image/png', data: Buffer.from('x').toString('base64') }] };
const bad = [];
const r1 = doPost({ parameter: { p: JSON.stringify(body) }, postData: { contents: 'p=...' } });
if (!/ส่งเรื่องแล้ว/.test(r1.t || '') || !/postMessage/.test(r1.h || '')) bad.push('แบบฟอร์ม: ไม่ได้หน้า "ส่งเรื่องแล้ว" พร้อม postMessage');
const r2 = doPost({ parameter: {}, postData: { contents: JSON.stringify(body) } });
if (!/"ok":true/.test(r2.t || '')) bad.push('แบบ JSON: ไม่ได้ {"ok":true}');
const r3 = doPost({ parameter: { p: JSON.stringify({ text: '' }) } });
if (!/ส่งไม่ได้/.test((r3.t || '') + (r3.h || ''))) bad.push('ข้อความว่าง: ไม่ขึ้น "ส่งไม่ได้"');
const data = rows.filter(r => /^#\d+/.test(String(r[0])));
if (data.length !== 2) bad.push('จำนวนแถวที่บันทึก ' + data.length + ' (ควร 2)');
if (!data.every(r => String(r[2]).includes('tester@example.com'))) bad.push('ไม่ได้บันทึกอีเมลผู้แจ้ง');
if (!files.includes('a.png')) bad.push('ไม่ได้เก็บรูป');
console.log(bad.length ? 'ไม่ผ่าน\n- ' + bad.join('\n- ') : 'ผ่าน', '| แถว', data.map(r => [r[0], r[2]]), '| ไฟล์', files);
process.exit(bad.length ? 1 : 0);
