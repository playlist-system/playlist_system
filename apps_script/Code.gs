/**
 * ดึง Playlist ประจำวันจาก Gmail ลง Google Drive อัตโนมัติ (Google Apps Script)
 * ใช้กับ PHOTO CHECK STUDIO — สคริปต์นี้แค่รวบรวมไฟล์ การเติมใบเช็คยังทำในระบบเหมือนเดิม
 *
 * ทำงานในบัญชี Google ของคุณเอง (ไม่ผ่านบริการภายนอก ไม่มีค่าใช้จ่าย)
 *   - ทุกชั่วโมง 7:00–18:00 (เวลาไทย) ค้นอีเมลที่มีไฟล์ Excel แนบ
 *   - เก็บ Playlist ของวันนี้ลงโฟลเดอร์ Google Drive "Playlist_ปปปป-ดด-วว"
 *   - ส่งอีเมลสรุปว่าได้กี่แบบแล้ว ขาดแบบไหน (ส่งเฉพาะรอบที่มีไฟล์ใหม่)
 *
 * ตั้งค่าครั้งแรก: แก้ค่าในส่วน "ตั้งค่า" ด้านล่าง (ถ้าต้องการ) แล้วเลือกฟังก์ชัน setup กด Run
 * ดึงทันทีโดยไม่รอรอบชั่วโมง: เลือก collectNow กด Run
 * หยุดทำงาน: เลือก stopAll กด Run
 * หน้าเว็บ "Playlist รายวัน" (เช็กลิสต์ + ดาวน์โหลดทั้งหมดเป็น .zip): ทำให้ใช้ได้ด้วย Deploy → New deployment → Web app (ดู README)
 */

// ===== ตั้งค่า =====
const CONFIG = {
  // อีเมลที่จะรับสรุป — หลายคนคั่นด้วยจุลภาค เว้นว่าง = ส่งหาเจ้าของสคริปต์
  notifyEmail: '',
  // ID โฟลเดอร์ Google Drive ที่จะเก็บ Playlist (จากลิงก์ drive.google.com/drive/folders/<ID>)
  // เว้นว่าง = ตอนรัน setup ระบบสร้างโฟลเดอร์ "Playlist รายวัน" ใน My Drive ให้
  parentFolderId: '',
  // อีเมลผู้ส่ง Playlist (คั่นด้วย OR) เช่น 'a@company.com OR b@company.com' — เว้นว่าง = ดูทุกอีเมลที่มีไฟล์ Excel แนบ
  fromFilter: '',
  // ช่วงเวลาที่ทำงาน (ชั่วโมง เวลาไทย)
  startHour: 7,
  endHour: 18,
  // แบบที่ไม่ได้ส่งทุกวัน (latest: true) ค้นอีเมลย้อนหลังกี่วัน และใช้คำค้นอะไร
  latestDays: 14,
  latestSearch: 'showdc OR "show dc" OR "LA PlanB" OR LA_PlanB',
  // Playlist ที่ระบบเติมใบเช็ครูปใช้ (ลำดับมีผล: ไฟล์เข้าแบบแรกที่ชื่อตรง)
  //   pattern = คำในชื่อไฟล์ • latest = ถ้าวันนี้ไม่มีฉบับใหม่ ใช้ฉบับล่าสุด
  //   source: 'line' = ไม่ได้มาทางอีเมล (เตือนในอีเมลสรุป) • extra = Playlist เพิ่มเติม ไม่นับใน 10 แบบ
  kinds: [
    { name: 'แยกจอ 40Blue', pattern: 'แยกจอ.*40\\s*Blue' },
    { name: 'แยกจอ 40Red', pattern: 'แยกจอ.*40\\s*Red' },
    { name: 'แยกจอ 80Blue', pattern: 'แยกจอ.*80\\s*Blue' },
    { name: 'แยกจอ 80Red', pattern: 'แยกจอ.*80\\s*Red' },
    { name: 'แยกจอ CITY', pattern: 'แยกจอ.*CITY' },
    { name: 'แยกจอ Spare', pattern: 'แยกจอ.*Spare' },
    { name: 'Show DC', pattern: 'Show[\\s_]*DC', latest: true },
    { name: 'Signature Max', pattern: 'Signature' },
    { name: 'BKJ1', pattern: 'BKJ\\s*1' },
    { name: 'BKJ2', pattern: 'BKJ\\s*2' },
    { name: 'Metro LED', pattern: 'Metro', source: 'line' },
    { name: 'Rama9 Connect', pattern: 'Rama\\s*9[\\s_]*connect', extra: true },
    { name: 'Digital MRT', pattern: 'Digital[\\s_]*MRT', extra: true },
    { name: 'Gateway X', pattern: 'Gateway[\\s_]*X', extra: true },
    { name: 'Air Traveller', pattern: 'Air[\\s_]*Travell?er', extra: true },
    { name: 'The 20', pattern: 'The[\\s_]*20(?!\\d)', extra: true },
    { name: 'LA PlanB', pattern: 'LA[\\s_]*Plan[\\s_]*B', latest: true, extra: true },
  ],
};

const TZ = 'Asia/Bangkok';
const PROP_FOLDER = 'parentFolderId';
const PROP_WEB = 'webUrl';
const STUDIO_URL = 'https://playlist-system.github.io/playlist_system/';

/** ตั้งค่าครั้งแรก: สร้างโฟลเดอร์หลัก (ถ้ายังไม่มี) ตั้งให้ทำงานทุกชั่วโมง แล้วดึงครั้งแรก */
function setup() {
  const parent = parentFolder_(true);
  try {
    for (const t of ScriptApp.getProjectTriggers()) if (t.getHandlerFunction() === 'collectPlaylists') ScriptApp.deleteTrigger(t);
    ScriptApp.newTrigger('collectPlaylists').timeBased().everyHours(1).inTimezone(TZ).create();
    Logger.log('ตั้งค่าเสร็จแล้ว — ทำงานทุกชั่วโมง ' + CONFIG.startHour + ':00–' + CONFIG.endHour + ':59 เวลาไทย');
  } catch (e) {
    /* ตอนกดอนุญาต ไม่ได้ติ๊กสิทธิ์ "เรียกใช้เมื่อคุณไม่ได้อยู่" (script.scriptapp) หรือบริษัทไม่อนุญาตสิทธิ์นี้ */
    Logger.log('ยังตั้งให้ทำงานอัตโนมัติไม่ได้: ไม่มีสิทธิ์ "เรียกใช้เมื่อคุณไม่ได้อยู่"');
    Logger.log('วิธีแก้: กด Run setup อีกครั้ง ตอน Google ถามสิทธิ์ ให้ติ๊ก "เลือกทั้งหมด" (ทุกช่อง) แล้วกดดำเนินการต่อ');
    Logger.log('ระหว่างนี้ดึงเองได้ด้วย collectNow — ถ้าติ๊กครบแล้วยังขึ้นเรื่องนี้ แปลว่าบริษัทไม่อนุญาตสิทธิ์นี้ ต้องขอฝ่าย IT');
  }
  Logger.log('โฟลเดอร์หลัก: ' + parent.getUrl());
  const r = run_(true);
  Logger.log(r.text);
}

/** ดึงทันที (ไม่สนช่วงเวลา) */
function collectNow() {
  Logger.log(run_(true).text);
}

/** หยุดทำงานอัตโนมัติ */
function stopAll() {
  let n = 0;
  for (const t of ScriptApp.getProjectTriggers()) if (t.getHandlerFunction() === 'collectPlaylists') {
    ScriptApp.deleteTrigger(t);
    n++;
  }
  Logger.log(n ? 'หยุดทำงานอัตโนมัติแล้ว' : 'ไม่ได้ตั้งให้ทำงานอัตโนมัติไว้');
}

/** ตัวที่ทำงานทุกชั่วโมง */
function collectPlaylists() {
  const h = +Utilities.formatDate(new Date(), TZ, 'H');
  if (h < CONFIG.startHour || h > CONFIG.endHour) return;
  run_(false);
}

// ---------- ภายใน ----------
function parentFolder_(create) {
  const props = PropertiesService.getScriptProperties();
  const id = CONFIG.parentFolderId || props.getProperty(PROP_FOLDER);
  if (id) return DriveApp.getFolderById(id);
  if (!create) throw new Error('ยังไม่ได้ตั้งค่า — เลือกฟังก์ชัน setup แล้วกด Run');
  const f = DriveApp.getRootFolder().createFolder('Playlist รายวัน');
  props.setProperty(PROP_FOLDER, f.getId());
  return f;
}

function todayFolder_(parent, today, create) {
  const name = 'Playlist_' + today;
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return create ? parent.createFolder(name) : null;
}

const MON_ = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
const THM_ = { 'มค': 0, 'กพ': 1, 'มีค': 2, 'เมย': 3, 'พค': 4, 'มิย': 5, 'กค': 6, 'สค': 7, 'กย': 8, 'ตค': 9, 'พย': 10, 'ธค': 11 };
const yy_ = y => {
  y = +y;
  if (y < 100) y += y >= 60 ? 2500 : 2000;
  if (y > 2400) y -= 543;
  return y;
};
const ymd_ = (y, m, d) => (m >= 0 && m < 12 && d >= 1 && d <= 31 ? y + '-' + String(m + 1).padStart(2, '0') + '-' + String(d).padStart(2, '0') : null);

/** วันที่จากชื่อไฟล์ (กติกาเดียวกับ PHOTO CHECK STUDIO): 2026 10 09 • 20261009 • 2026_Oct_09 • 9-10-26 • 9 ต.ค. 69 • 9 Oct 26 */
function nameDate_(n, thisYear) {
  n = String(n).replace(/\.xls[xm]?$/i, '');
  let m = n.match(/(?:^|[^\d])(\d{4})[-_. ](\d{1,2})[-_. ](\d{1,2})(?!\d)/);
  if (m) return ymd_(yy_(m[1]), +m[2] - 1, +m[3]);
  m = n.match(/(?:^|[^\d])(20\d{2})(\d{2})(\d{2})(?!\d)/);
  if (m) return ymd_(+m[1], +m[2] - 1, +m[3]);
  m = n.match(/(?:^|[^\d])(\d{4})[-_. ]?(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[-_. ]?(\d{1,2})(?!\d)/i);
  if (m) return ymd_(yy_(m[1]), MON_[m[2].toLowerCase()], +m[3]);
  m = n.match(/(?:^|[^\d])(\d{1,2})[-_.](\d{1,2})[-_.](\d{2,4})(?!\d)/);
  if (m) return ymd_(yy_(m[3]), +m[2] - 1, +m[1]);
  m = n.replace(/\s+/g, '').match(/(\d{1,2})(ม\.?ค\.?|ก\.?พ\.?|มี\.?ค\.?|เม\.?ย\.?|พ\.?ค\.?|มิ\.?ย\.?|ก\.?ค\.?|ส\.?ค\.?|ก\.?ย\.?|ต\.?ค\.?|พ\.?ย\.?|ธ\.?ค\.?)(\d{2,4})?/);
  if (m) return ymd_(m[3] ? yy_(m[3]) : thisYear, THM_[m[2].replace(/\./g, '')], +m[1]);
  m = n.match(/(?:^|[^A-Za-z\d])(\d{1,2})[-_ ]*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[-_ ]*(\d{2,4})?(?![A-Za-z])/i);
  if (m) return ymd_(m[3] ? yy_(m[3]) : thisYear, MON_[m[2].toLowerCase()], +m[1]);
  return null;
}

/** คัดไฟล์ของวันนี้จากรายการไฟล์แนบ [{name, size, day(อีเมล), when, blob}] → {chosen, others} */
function pick_(atts, today) {
  const kinds = CONFIG.kinds.map(k => Object.assign({}, k, { re: new RegExp(k.pattern, 'i') }));
  const thisYear = +today.slice(0, 4);
  const byKind = {};
  const others = {};
  for (const a of atts) {
    const name = String(a.name || '').normalize('NFC');
    if (!/\.xlsx?$/i.test(name)) continue;
    const day = nameDate_(name, thisYear) || a.day;
    if (!day || day > today) continue; // ไฟล์ของวันข้างหน้า รอวันนั้น
    const kind = kinds.find(k => k.re.test(name));
    if (!kind) {
      if (day === today) others[name] = 1;
      continue;
    }
    if (day !== today && !kind.latest) continue; // ไฟล์ของวันก่อน
    const key = name + '|' + a.size;
    const list = byKind[kind.name] || (byKind[kind.name] = {});
    if (!list[key]) list[key] = { kind: kind.name, name: name, key: key, day: day, when: a.when, blob: a.blob, size: a.size };
  }
  const chosen = [];
  for (const k of kinds) {
    const list = Object.values(byKind[k.name] || {});
    const todays = list.filter(f => f.day === today);
    if (todays.length) chosen.push(...todays);
    else if (k.latest && list.length) {
      list.sort((a, b) => (a.day < b.day ? -1 : a.day > b.day ? 1 : a.when - b.when));
      chosen.push(list[list.length - 1]); // ฉบับล่าสุดที่มี
    }
  }
  return { chosen: chosen, others: Object.keys(others) };
}

/** ไฟล์แนบ Excel จาก Gmail (ย้อนหลัง 2 วัน + แบบที่ไม่ได้ส่งทุกวันย้อนหลัง latestDays วัน) */
function attachments_() {
  const from = CONFIG.fromFilter ? ' from:(' + CONFIG.fromFilter + ')' : '';
  const qs = ['has:attachment newer_than:2d {filename:xlsx filename:xls}' + from];
  if (CONFIG.latestSearch) qs.push('has:attachment newer_than:' + CONFIG.latestDays + 'd {filename:xlsx filename:xls} (' + CONFIG.latestSearch + ')' + from);
  const seen = {};
  const out = [];
  for (const q of qs)
    for (const th of GmailApp.search(q, 0, 200))
      for (const msg of th.getMessages()) {
        if (seen[msg.getId()]) continue;
        seen[msg.getId()] = 1;
        const d = msg.getDate();
        for (const at of msg.getAttachments({ includeInlineImages: false })) {
          const name = at.getName();
          if (!/\.xlsx?$/i.test(name || '')) continue;
          out.push({ name: name, size: at.getSize(), day: Utilities.formatDate(d, TZ, 'yyyy-MM-dd'), when: d.getTime(), blob: at });
        }
      }
  return out;
}

function run_(force) {
  // กันดึงซ้อนกัน (รอบชั่วโมงกับปุ่ม "ดึงตอนนี้" ในหน้าเว็บพร้อมกัน) — ไม่งั้นไฟล์อาจเข้าโฟลเดอร์ซ้ำ
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(120000)) return { added: [], subject: '', text: 'มีการดึงอีกรอบทำงานอยู่ ลองใหม่อีกครั้งในอีกสักครู่' };
  try {
    return collect_();
  } finally {
    lock.releaseLock();
  }
}

function collect_() {
  const now = new Date();
  const today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
  const parent = parentFolder_(false);
  const P = pick_(attachments_(), today);
  let folder = todayFolder_(parent, today, false);
  // ไฟล์ที่เก็บไว้แล้ว (ชื่อ + ขนาดเดียวกัน) ไม่เก็บซ้ำ
  const have = {};
  if (folder) {
    const it = folder.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      have[f.getName().normalize('NFC') + '|' + f.getSize()] = f.getName();
    }
  }
  const added = [];
  for (const f of P.chosen) {
    if (have[f.key]) continue;
    if (!folder) folder = todayFolder_(parent, today, true);
    folder.createFile(f.blob.copyBlob().setName(f.name));
    have[f.key] = f.name;
    added.push(f);
  }
  // แบบที่ไม่ได้ส่งทุกวัน (เช่น Show DC): ได้ฉบับของวันนี้แล้ว → ฉบับเก่าที่ใช้แทนไว้ก่อนหน้า ย้ายไปถังขยะของ Drive (กู้คืนได้)
  const trashed = [];
  if (folder && added.length) {
    const thisYear = +today.slice(0, 4);
    for (const k of CONFIG.kinds.filter(x => x.latest)) {
      const re = new RegExp(k.pattern, 'i');
      const L = [];
      const it = folder.getFiles();
      while (it.hasNext()) {
        const f = it.next();
        const n = f.getName().normalize('NFC');
        if (re.test(n)) L.push({ f: f, n: n, day: nameDate_(n, thisYear) });
      }
      if (!L.some(x => x.day === today)) continue;
      for (const x of L)
        if (x.day && x.day < today) {
          x.f.setTrashed(true);
          trashed.push(x.n);
        }
    }
  }
  const sum = summary_(folder, today, added, P.others, trashed);
  if (added.length) MailApp.sendEmail({ to: CONFIG.notifyEmail || Session.getEffectiveUser().getEmail(), subject: sum.subject, htmlBody: sum.html, name: 'Playlist อัตโนมัติ' });
  return { added: added.map(f => f.name), subject: sum.subject, text: sum.subject + (added.length ? ' — เพิ่ม ' + added.length + ' ไฟล์ ส่งอีเมลสรุปแล้ว' : ' — ไม่มีไฟล์ใหม่') };
}

function summary_(folder, today, added, others, trashed) {
  const TH = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  const thDay = s => +s.slice(8, 10) + ' ' + TH[+s.slice(5, 7) - 1];
  const day = thDay(today) + ' ' + (+today.slice(0, 4) + 543);
  const kinds = CONFIG.kinds.map(k => Object.assign({}, k, { re: new RegExp(k.pattern, 'i') }));
  const counts = {};
  const files = [];
  if (folder) {
    const it = folder.getFiles();
    while (it.hasNext()) {
      const n = it.next().getName().normalize('NFC');
      files.push(n);
      const k = kinds.find(x => x.re.test(n));
      if (k) counts[k.name] = (counts[k.name] || 0) + 1;
    }
  }
  const mail = kinds.filter(k => k.source !== 'line' && !k.extra).map(k => k.name);
  const line = kinds.filter(k => k.source === 'line' && !counts[k.name]).map(k => k.name);
  const extra = kinds.filter(k => k.extra && counts[k.name]).map(k => k.name);
  const haveN = mail.filter(k => counts[k]).length;
  const missing = mail.filter(k => !counts[k]);
  const multi = mail.filter(k => counts[k] > 1);
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
  const html =
    '<div style="font-family:Tahoma,sans-serif;font-size:15px;line-height:1.6">' +
    '<p>Playlist วันที่ ' + esc(day) + ' — ได้แล้ว <b>' + haveN + '/' + mail.length + '</b> แบบ (ที่มาทางอีเมล)</p>' +
    (added.length
      ? '<p>ไฟล์ที่เพิ่งเข้าโฟลเดอร์รอบนี้:</p><ul>' +
        added
          .map(f => '<li>' + esc(f.kind) + ': ' + esc(f.name) + (f.day !== today ? ' <span style="color:#b35a14">(ยังไม่มีฉบับใหม่ — ใช้ฉบับล่าสุด รอบ ' + esc(thDay(f.day)) + ')</span>' : '') + '</li>')
          .join('') +
        '</ul>'
      : '') +
    (missing.length
      ? '<p>ยังขาด <b>' + missing.length + '</b> แบบ: ' + missing.map(esc).join(', ') + '</p><p>ระบบเช็กอีเมลทุกชั่วโมงจนถึง ' + CONFIG.endHour + ':00</p>'
      : '<p><b>ครบ ' + mail.length + ' แบบที่มาทางอีเมลแล้ว</b> ดาวน์โหลดทั้งโฟลเดอร์ แล้วลากลง PHOTO CHECK STUDIO ได้เลย</p>') +
    (line.length ? '<p style="color:#1d4f91"><b>มาทาง LINE:</b> ' + line.map(esc).join(', ') + ' — ดาวน์โหลดจาก LINE แล้วใส่ในระบบเอง</p>' : '') +
    (extra.length ? '<p>Playlist เพิ่มเติม (สำหรับใบเช็ค CCTV / MRT / LA PlanB): ' + extra.map(esc).join(', ') + '</p>' : '') +
    (multi.length ? '<p style="color:#b35a14"><b>บางแบบมีมากกว่า 1 ไฟล์:</b> ' + multi.map(k => esc(k) + ' (' + counts[k] + ' ไฟล์)').join(', ') + ' — ตรวจว่าต้องใช้ฉบับไหนก่อนลากลงระบบ</p>' : '') +
    (trashed && trashed.length ? '<p style="color:#586274">ได้ฉบับของวันนี้แล้ว ย้ายฉบับเก่าไปถังขยะของ Drive: ' + trashed.map(esc).join(', ') + '</p>' : '') +
    (others.length ? '<p style="color:#586274">ไฟล์ Excel อื่นของวันนี้ (ไม่ได้เก็บ เพราะชื่อไม่ตรงแบบที่ใช้):</p><ul style="color:#586274">' + others.map(n => '<li>' + esc(n) + '</li>').join('') + '</ul>' : '') +
    (webUrl_() ? '<p><a href="' + webUrl_() + '"><b>เปิดหน้า Playlist รายวัน</b></a> — ดูเช็กลิสต์และดาวน์โหลดทั้งหมดเป็นไฟล์ .zip ไฟล์เดียว</p>' : '') +
    (folder ? '<p><a href="' + folder.getUrl() + '">เปิดโฟลเดอร์ Google Drive ของวันนี้</a></p>' : '') +
    '</div>';
  return { subject: 'Playlist ' + day + ': ' + haveN + '/' + mail.length + ' แบบ' + (missing.length ? ' (ยังไม่ครบ)' : ' — ครบแล้ว'), html: html };
}

// ---------- หน้าเว็บ "Playlist รายวัน" ----------
// ทำให้ใช้ได้: Deploy → New deployment → ประเภท Web app → Execute as: Me → Who has access: Only myself (หรือทุกคนในบริษัท) → Deploy
// หน้าเว็บอ่านเฉพาะโฟลเดอร์ "Playlist รายวัน" ของสคริปต์นี้ ไม่แก้ไข/ลบไฟล์

/** เปิดหน้าเว็บ */
function doGet() {
  try {
    const u = ScriptApp.getService().getUrl();
    if (/\/exec$/.test(u || '') && PropertiesService.getScriptProperties().getProperty(PROP_WEB) !== u) PropertiesService.getScriptProperties().setProperty(PROP_WEB, u);
  } catch (e) {}
  return HtmlService.createHtmlOutput(page_()).setTitle('Playlist รายวัน').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** ลิงก์หน้าเว็บ (จำไว้ตอนเปิดหน้าเว็บครั้งแรก) — ใช้ใส่ในอีเมลสรุป */
function webUrl_() {
  try {
    return PropertiesService.getScriptProperties().getProperty(PROP_WEB) || '';
  } catch (e) {
    return '';
  }
}

/** ข้อมูลของวันหนึ่ง (ปปปป-ดด-วว) ให้หน้าเว็บ */
function webDay(day) {
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(day || '') || day > today) day = today;
  const parent = parentFolder_(false);
  const folder = todayFolder_(parent, day, false);
  const TH = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  const thDay = s => +s.slice(8, 10) + ' ' + TH[+s.slice(5, 7) - 1];
  const kinds = CONFIG.kinds.map(k => Object.assign({}, k, { re: new RegExp(k.pattern, 'i') }));
  const thisYear = +day.slice(0, 4);
  const files = [];
  if (folder) {
    const it = folder.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      const n = f.getName().normalize('NFC');
      const k = kinds.find(x => x.re.test(n));
      const fd = nameDate_(n, thisYear);
      const at = f.getDateCreated();
      const atDay = Utilities.formatDate(at, TZ, 'yyyy-MM-dd');
      files.push({
        id: f.getId(),
        name: n,
        size: f.getSize(),
        kind: k ? k.name : '',
        at: at.getTime(),
        time: (atDay !== day ? thDay(atDay) + ' ' : '') + Utilities.formatDate(at, TZ, 'H:mm'),
        old: fd && fd < day ? thDay(fd) : '',
      });
    }
  }
  files.sort((a, b) => a.at - b.at);
  let auto = null;
  try {
    auto = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'collectPlaylists');
  } catch (e) {}
  return {
    day: day,
    today: today,
    folderUrl: folder ? folder.getUrl() : '',
    parentUrl: parent.getUrl(),
    kinds: CONFIG.kinds.map(k => ({ name: k.name, line: k.source === 'line', extra: !!k.extra, latest: !!k.latest })),
    files: files,
    auto: auto,
    startHour: CONFIG.startHour,
    endHour: CONFIG.endHour,
    studio: STUDIO_URL,
  };
}

/** ส่งไฟล์หนึ่งไฟล์ในโฟลเดอร์ของวันนั้นให้หน้าเว็บทีละส่วน (base64) — หน้าเว็บรวมไฟล์ที่เลือกเป็น .zip เอง
 *  (Playlist ทั้งวันรวมกันราว 20 MB ถ้าส่งก้อนเดียวอาจเกินขนาดที่ google.script.run รับได้ จึงส่งทีละไม่เกิน 8 ล้านตัวอักษร) */
function webPart(day, id, part) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(day || '')) throw new Error('วันที่ไม่ถูกต้อง');
  const folder = todayFolder_(parentFolder_(false), day, false);
  if (!folder) throw new Error('ยังไม่มีโฟลเดอร์ Playlist ของวันนี้');
  const it = folder.getFiles(); // เฉพาะไฟล์ในโฟลเดอร์ของวันนั้นเท่านั้น
  while (it.hasNext()) {
    const f = it.next();
    if (f.getId() !== id) continue;
    const b64 = Utilities.base64Encode(f.getBlob().getBytes());
    const size = 8 * 1024 * 1024; // หาร 4 ลงตัว แต่ละส่วนถอดรหัสแยกกันได้
    const n = Math.max(1, Math.ceil(b64.length / size));
    part = Math.min(Math.max(0, part | 0), n - 1);
    return { n: n, b64: b64.slice(part * size, (part + 1) * size) };
  }
  throw new Error('ไม่พบไฟล์นี้ในโฟลเดอร์แล้ว (อาจถูกย้ายหรือลบ) — โหลดหน้าใหม่แล้วลองอีกครั้ง');
}

/** ปุ่ม "ดึงตอนนี้" ในหน้าเว็บ */
function webCollect() {
  const r = run_(true);
  return { msg: r.text, added: r.added.length, data: webDay('') };
}

function page_() {
  return `<!doctype html>
<html lang="th"><head><base target="_top"><meta charset="utf-8">
<style>
:root{--bg:#f3f5f8;--card:#fff;--ink:#16202c;--sub:#5a6575;--line:#e1e6ed;--acc:#1d4f91;--acc-ink:#fff;--ok:#1d7a4b;--ok-bg:#e6f3eb;--warn:#9c4f0e;--warn-bg:#fcf0e2;--miss:#b0242d;--miss-bg:#fae9ea;--info:#1d4f91;--info-bg:#e8eff9;--bar:#e1e6ed;--sh:0 1px 2px rgba(16,24,40,.06)}
@media (prefers-color-scheme:dark){:root{--bg:#10151c;--card:#18202a;--ink:#e5eaf1;--sub:#9aa5b4;--line:#29323e;--acc:#7aaaf0;--acc-ink:#0c1420;--ok:#62c792;--ok-bg:#15301f;--warn:#e9a862;--warn-bg:#33240f;--miss:#f08389;--miss-bg:#3a1a1d;--info:#8cb6ef;--info-bg:#162a43;--bar:#29323e;--sh:none}}
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.6 Tahoma,"Leelawadee UI","Noto Sans Thai",sans-serif}
.wrap{max-width:760px;margin:0 auto;padding:18px 16px 110px}
a{color:var(--acc)}
.top{display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap}
h1{font-size:21px;margin:0;letter-spacing:.2px}
.lnk{font-size:14px;font-weight:bold;text-decoration:none;white-space:nowrap}
.lnk:hover{text-decoration:underline}
.date{display:flex;align-items:center;gap:8px;margin:14px 0 4px;flex-wrap:wrap}
.date .d{font-size:17px;font-weight:bold;margin-right:auto}
.ib{appearance:none;border:1px solid var(--line);background:var(--card);color:var(--ink);border-radius:9px;min-width:38px;height:38px;padding:0 10px;font:inherit;cursor:pointer}
.ib:disabled{opacity:.4;cursor:default}
input[type=date].ib{padding:0 8px}
.card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:16px;margin:12px 0;box-shadow:var(--sh)}
.card h2{font-size:15px;margin:0 0 2px}
.card .note{color:var(--sub);font-size:13.5px;margin:0}
.st{display:flex;align-items:flex-end;justify-content:space-between;gap:12px;flex-wrap:wrap}
.big{font-size:38px;font-weight:bold;line-height:1}
.big small{font-size:16px;color:var(--sub);font-weight:normal;margin-left:6px}
.seg{display:grid;grid-template-columns:repeat(10,1fr);gap:4px;margin:12px 0 10px}
.seg i{height:8px;border-radius:4px;background:var(--bar)}
.seg i.on{background:var(--ok)}
.msg{margin:0;font-weight:bold}
.msg.ok{color:var(--ok)}
.msg.miss{color:var(--miss)}
.meta{color:var(--sub);font-size:13.5px;margin:4px 0 0}
.btn{appearance:none;border:0;border-radius:10px;padding:10px 16px;font:inherit;font-weight:bold;cursor:pointer;background:var(--acc);color:var(--acc-ink);text-decoration:none;display:inline-flex;align-items:center;gap:6px;white-space:nowrap}
.btn.sec{background:transparent;color:var(--acc);border:1px solid var(--line)}
.btn:disabled{opacity:.5;cursor:default}
.list{margin-top:6px}
.row{display:grid;grid-template-columns:26px 132px 1fr;gap:4px 10px;align-items:start;padding:10px 0;border-top:1px solid var(--line)}
.row:first-child{border-top:0}
label.row{cursor:pointer}
.row input{width:18px;height:18px;margin:2px 0 0;accent-color:var(--acc);cursor:pointer}
.k{font-weight:bold}
.f{min-width:0;overflow-wrap:anywhere}
.f small{display:block;color:var(--sub);font-size:13px}
.ic{width:20px;height:20px;border-radius:50%;display:inline-grid;place-items:center;font-size:12px;font-weight:bold;margin-top:2px}
.ic.miss{background:var(--miss-bg);color:var(--miss)}
.ic.info{background:var(--info-bg);color:var(--info)}
.row.miss .f{color:var(--miss)}
.tag{display:inline-block;font-size:12px;font-weight:bold;border-radius:6px;padding:0 6px;margin-left:4px}
.tag.warn{background:var(--warn-bg);color:var(--warn)}
.tag.info{background:var(--info-bg);color:var(--info)}
.warnbox{background:var(--warn-bg);color:var(--warn);border-radius:10px;padding:8px 12px;font-size:13.5px;margin:8px 0 0}
.how{margin:6px 0 0;padding-left:20px}
.how li{margin:2px 0}
.bar{position:fixed;left:0;right:0;bottom:0;background:var(--card);border-top:1px solid var(--line);padding:12px 16px calc(12px + env(safe-area-inset-bottom))}
.bar .in{max-width:760px;margin:0 auto;display:flex;gap:10px;align-items:center;flex-wrap:wrap}
.bar .sum{color:var(--sub);font-size:13.5px;margin-left:auto}
.toast{position:fixed;left:50%;bottom:84px;transform:translateX(-50%);max-width:min(92vw,620px);background:var(--ink);color:var(--bg);padding:10px 14px;border-radius:10px;font-size:14px;opacity:0;pointer-events:none;transition:opacity .2s}
.toast.on{opacity:1}
.toast.err{background:var(--miss);color:#fff}
.empty{color:var(--sub);text-align:center;padding:30px 0}
@media (max-width:540px){.row{grid-template-columns:26px 1fr}.row .f{grid-column:2}.row .k:empty{display:none}.big{font-size:32px}.date .d{width:100%}.bar .in{gap:8px}.bar .sum{order:-1;margin:0;width:100%}.bar .btn{flex:1 1 auto;justify-content:center;padding:10px 12px}}
</style></head><body>
<div class="wrap" id="app"><p class="empty">กำลังโหลด…</p></div>
<div class="bar" id="bar" hidden></div>
<div class="toast" id="toast" role="status" aria-live="polite"></div>
<script>
var D = null, SEL = {}, BUSY = '', TT = 0, PROG = 0;
var TH = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
var WD = ['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'];
function $(s) { return document.querySelector(s); }
function esc(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function dt(d) { return new Date(d + 'T12:00:00Z'); }
function thLong(d) { var t = dt(d); return WD[t.getUTCDay()] + ' ' + t.getUTCDate() + ' ' + TH[t.getUTCMonth()] + ' ' + (t.getUTCFullYear() + 543); }
function shift(d, n) { var t = dt(d); t.setUTCDate(t.getUTCDate() + n); return t.toISOString().slice(0, 10); }
function kb(n) { return n >= 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB'; }
function toast(m, err) { var t = $('#toast'); t.textContent = m; t.className = 'toast on' + (err ? ' err' : ''); clearTimeout(TT); TT = setTimeout(function () { t.className = 'toast'; }, err ? 9000 : 6000); }
function call(fn, args, ok) {
  google.script.run.withSuccessHandler(ok).withFailureHandler(function (e) {
    var m = (e && e.message) || String(e); BUSY = '';
    if (!D) $('#app').innerHTML = '<div class="card"><h2>เปิดข้อมูลไม่ได้</h2><p class="note">' + esc(m) + '</p></div>';
    else draw();
    toast(m, 1);
  })[fn].apply(null, args);
}
function load(day) { BUSY = 'load'; draw(); call('webDay', [day || ''], function (d) { D = d; BUSY = ''; defaults(); draw(); }); }
function byKind() { var by = {}, oth = []; D.files.forEach(function (f) { if (f.kind) (by[f.kind] = by[f.kind] || []).push(f); else oth.push(f); }); return { by: by, oth: oth }; }
/* ค่าเริ่มต้น: ติ๊ก Playlist ที่ระบบใช้ (แบบละ 1 ไฟล์ = ฉบับที่เข้าโฟลเดอร์ล่าสุด) • Playlist เพิ่มเติมและไฟล์อื่นไม่ติ๊ก */
function defaults() {
  SEL = {}; var g = byKind();
  D.kinds.forEach(function (k) { var L = g.by[k.name]; if (!L || k.extra) return; SEL[L[L.length - 1].id] = true; });
}
function picked() { return D ? D.files.filter(function (f) { return SEL[f.id]; }) : []; }
function fileRow(f, label, first, multi) {
  var meta = ['เข้าโฟลเดอร์ ' + f.time, kb(f.size)];
  return '<label class="row"><input type="checkbox" data-id="' + esc(f.id) + '"' + (SEL[f.id] ? ' checked' : '') + '>' +
    '<span class="k">' + (first ? esc(label) + (multi ? '<span class="tag warn">' + multi + ' ไฟล์</span>' : '') : '') + '</span>' +
    '<span class="f">' + esc(f.name) + (f.old ? '<span class="tag info">ฉบับ ' + esc(f.old) + '</span>' : '') +
    '<small>' + meta.join(' · ') + (f.old ? ' · วันนี้ยังไม่มีฉบับใหม่ ใช้ฉบับล่าสุดไปก่อน' : '') + '</small></span></label>';
}
function draw() {
  var app = $('#app'), bar = $('#bar');
  if (!D) { app.innerHTML = '<p class="empty">กำลังโหลด…</p>'; bar.hidden = true; return; }
  var g = byKind(), day = D.day, isToday = day === D.today;
  var main = D.kinds.filter(function (k) { return !k.extra && !k.line; });
  var line = D.kinds.filter(function (k) { return k.line; });
  var extra = D.kinds.filter(function (k) { return k.extra && g.by[k.name]; });
  var have = main.filter(function (k) { return g.by[k.name]; });
  var miss = main.filter(function (k) { return !g.by[k.name]; });
  var multi = main.concat(line).filter(function (k) { return g.by[k.name] && g.by[k.name].length > 1; });
  var last = D.files.length ? D.files[D.files.length - 1] : null;
  var h = '<div class="top"><h1>Playlist รายวัน</h1><a class="lnk" href="' + esc(D.studio) + '" target="_blank" rel="noopener">เปิด PHOTO CHECK STUDIO ↗</a></div>';
  h += '<div class="date"><span class="d">' + esc(thLong(day)) + (isToday ? ' (วันนี้)' : '') + '</span>' +
    '<button class="ib" data-go="' + shift(day, -1) + '" aria-label="วันก่อนหน้า" title="วันก่อนหน้า">‹</button>' +
    '<input class="ib" type="date" id="pick" value="' + day + '" max="' + D.today + '" aria-label="เลือกวันที่">' +
    '<button class="ib" data-go="' + shift(day, 1) + '" aria-label="วันถัดไป" title="วันถัดไป"' + (isToday ? ' disabled' : '') + '>›</button>' +
    (isToday ? '' : '<button class="ib" data-go="' + D.today + '">วันนี้</button>') + '</div>';
  // สถานะ
  h += '<section class="card"><div class="st"><div class="big">' + have.length + '/' + main.length + '<small>แบบ</small></div>' +
    (isToday ? '<button class="btn sec" data-a="collect"' + (BUSY ? ' disabled' : '') + '>' + (BUSY === 'collect' ? 'กำลังดึง…' : 'ดึงตอนนี้') + '</button>' : '') + '</div>' +
    '<div class="seg" aria-hidden="true">' + main.map(function (k) { return '<i class="' + (g.by[k.name] ? 'on' : '') + '"></i>'; }).join('') + '</div>';
  if (!D.files.length) h += '<p class="msg miss">' + (isToday ? 'ยังไม่มี Playlist ของวันนี้' : 'ไม่มี Playlist ของวันที่ ' + dt(day).getUTCDate() + ' ' + TH[dt(day).getUTCMonth()] + ' ในโฟลเดอร์') + '</p>';
  else if (!miss.length) h += '<p class="msg ok">ครบ ' + main.length + ' แบบแล้ว — ดาวน์โหลดแล้วลากลงระบบได้เลย</p>';
  else h += '<p class="msg miss">ยังขาด ' + miss.length + ' แบบ: ' + miss.map(function (k) { return esc(k.name); }).join(', ') + '</p>';
  var meta = [];
  if (last) meta.push('ไฟล์ล่าสุดเข้าโฟลเดอร์ ' + esc(last.time));
  if (isToday) meta.push(D.auto === false ? 'ดึงอัตโนมัติ: ปิดอยู่ (กด setup ในหน้าโค้ดเพื่อเปิด)' : 'ระบบเช็กอีเมลเองทุกชั่วโมง ' + D.startHour + ':00–' + D.endHour + ':59');
  if (meta.length) h += '<p class="meta">' + meta.join(' • ') + '</p>';
  h += '</section>';
  // รายการ 10 แบบ + Metro LED
  h += '<section class="card"><h2>Playlist ที่ใช้ในระบบเติมใบเช็ค</h2><p class="note">ติ๊กไฟล์ที่จะดาวน์โหลด</p><div class="list">';
  main.concat(line).forEach(function (k) {
    var L = g.by[k.name];
    if (L) { L.forEach(function (f, i) { h += fileRow(f, k.name, i === 0, L.length > 1 ? L.length : 0); }); return; }
    if (k.line) h += '<div class="row"><span class="ic info">i</span><span class="k">' + esc(k.name) + '</span><span class="f">มาทาง LINE — ดาวน์โหลดจาก LINE แล้วลากลงระบบพร้อมไฟล์ .zip</span></div>';
    else h += '<div class="row miss"><span class="ic miss">✕</span><span class="k">' + esc(k.name) + '</span><span class="f">ยังไม่มา</span></div>';
  });
  h += '</div>';
  if (multi.length) h += '<p class="warnbox">' + multi.map(function (k) { return esc(k.name); }).join(', ') + ' มีมากกว่า 1 ไฟล์ (มีฉบับแก้มา) — ติ๊กไว้ให้เฉพาะไฟล์ที่เข้ามาล่าสุด ตรวจก่อนดาวน์โหลด</p>';
  h += '</section>';
  if (extra.length) {
    h += '<section class="card"><h2>Playlist เพิ่มเติม</h2><p class="note">ระบบเติมใบเช็คไม่ได้ใช้ ติ๊กถ้าต้องการดาวน์โหลดไปด้วย</p><div class="list">';
    extra.forEach(function (k) { g.by[k.name].forEach(function (f, i) { h += fileRow(f, k.name, i === 0, 0); }); });
    h += '</div></section>';
  }
  if (g.oth.length) {
    h += '<section class="card"><h2>ไฟล์อื่นในโฟลเดอร์</h2><div class="list">';
    g.oth.forEach(function (f, i) { h += fileRow(f, 'อื่น ๆ', i === 0, 0); });
    h += '</div></section>';
  }
  h += '<section class="card"><h2>วิธีใช้</h2><ol class="how"><li>กด <b>ดาวน์โหลด</b> ด้านล่าง ได้ไฟล์ .zip ไฟล์เดียว</li><li>ดาวน์โหลด Metro LED จาก LINE</li><li>เปิด PHOTO CHECK STUDIO แล้วลากไฟล์ .zip กับไฟล์ Metro LED ลงขั้นที่ 1 พร้อมกัน — <b>ไม่ต้องแตกไฟล์</b></li></ol></section>';
  app.innerHTML = h;
  app.style.opacity = BUSY === 'load' ? '.55' : '';
  drawBar();
}
function drawBar() {
  var bar = $('#bar'), P = picked(), sz = 0;
  P.forEach(function (f) { sz += f.size; });
  var url = D.folderUrl || D.parentUrl;
  bar.innerHTML = '<div class="in"><button class="btn" data-a="zip"' + (!P.length || BUSY ? ' disabled' : '') + '>' +
    (BUSY === 'zip' ? 'กำลังเตรียมไฟล์ ' + PROG + '/' + P.length + '…' : 'ดาวน์โหลด ' + P.length + ' ไฟล์ (.zip)') + '</button>' +
    '<a class="btn sec" href="' + esc(url) + '" target="_blank" rel="noopener">เปิดใน Drive</a>' +
    '<span class="sum">' + (P.length ? 'รวม ' + kb(sz) : 'ยังไม่ได้เลือกไฟล์') + '</span></div>';
  bar.hidden = false;
}
/* รวมเป็น .zip ในเบราว์เซอร์ (แบบไม่บีบอัด — ไฟล์ Excel บีบอัดอยู่แล้ว) ชื่อไฟล์ภาษาไทยเปิดได้ทั้ง Windows และ PHOTO CHECK STUDIO */
var CRC = (function () { var t = [], c, n, k; for (n = 0; n < 256; n++) { c = n; for (k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc32(u) { var c = -1; for (var i = 0; i < u.length; i++) c = CRC[(c ^ u[i]) & 255] ^ (c >>> 8); return (c ^ -1) >>> 0; }
function zipBlob(list) {
  var enc = new TextEncoder(), parts = [], cen = [], off = 0, now = new Date(), cs = 0;
  var tm = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1), dd = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
  list.forEach(function (e) {
    var nm = enc.encode(e.name), crc = crc32(e.data), len = e.data.length, h = new DataView(new ArrayBuffer(30)), c = new DataView(new ArrayBuffer(46));
    h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x800, true); h.setUint16(10, tm, true); h.setUint16(12, dd, true);
    h.setUint32(14, crc, true); h.setUint32(18, len, true); h.setUint32(22, len, true); h.setUint16(26, nm.length, true);
    c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x800, true); c.setUint16(12, tm, true); c.setUint16(14, dd, true);
    c.setUint32(16, crc, true); c.setUint32(20, len, true); c.setUint32(24, len, true); c.setUint16(28, nm.length, true); c.setUint32(42, off, true);
    parts.push(h.buffer, nm, e.data); cen.push(c.buffer, nm); off += 30 + nm.length + len; cs += 46 + nm.length;
  });
  var end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true); end.setUint16(8, list.length, true); end.setUint16(10, list.length, true); end.setUint32(12, cs, true); end.setUint32(16, off, true);
  return new Blob(parts.concat(cen, [end.buffer]), { type: 'application/zip' });
}
function b64u8(s) { var bin = atob(s), u = new Uint8Array(bin.length); for (var i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; }
function join(L) { var n = 0, o = 0; L.forEach(function (u) { n += u.length; }); var r = new Uint8Array(n); L.forEach(function (u) { r.set(u, o); o += u.length; }); return r; }
function zip() {
  var P = picked(), out = [], next = 0, done = 0, failed = false;
  if (!P.length || BUSY) return;
  BUSY = 'zip'; PROG = 0; draw();
  function fail(e) { if (failed) return; failed = true; BUSY = ''; draw(); toast('ดาวน์โหลดไม่สำเร็จ: ' + ((e && e.message) || e) + ' — ลองอีกครั้ง หรือใช้ปุ่ม “เปิดใน Drive”', 1); }
  function one(i, part, acc) {
    google.script.run.withFailureHandler(fail).withSuccessHandler(function (r) {
      if (failed) return;
      acc.push(b64u8(r.b64));
      if (part + 1 < r.n) return one(i, part + 1, acc);
      out[i] = { name: P[i].name, data: join(acc) }; done++; PROG = done; drawBar();
      if (next < P.length) one(next++, 0, []);
      else if (done === P.length) finish();
    }).webPart(D.day, P[i].id, part);
  }
  function finish() {
    var used = {};
    out.forEach(function (e) {
      var n = e.name, d = n.lastIndexOf('.'), b = d > 0 ? n.slice(0, d) : n, x = d > 0 ? n.slice(d) : '', k = 2;
      while (used[n.toLowerCase()]) n = b + ' (' + k++ + ')' + x;
      used[n.toLowerCase()] = 1; e.name = n;
    });
    var name = 'Playlist_' + D.day + '.zip', url = URL.createObjectURL(zipBlob(out));
    var a = document.createElement('a'); a.href = url; a.download = name; a.target = '_self';
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 60000);
    BUSY = ''; draw();
    toast('ดาวน์โหลด ' + name + ' (' + out.length + ' ไฟล์) แล้ว — ลากไฟล์นี้ลงขั้นที่ 1 ของ PHOTO CHECK STUDIO ได้เลย ไม่ต้องแตกไฟล์');
  }
  for (var k = 0; k < Math.min(3, P.length); k++) one(next++, 0, []);
}
function collect() {
  BUSY = 'collect'; draw();
  call('webCollect', [], function (r) { D = r.data; BUSY = ''; defaults(); draw(); toast(r.msg); });
}
document.addEventListener('click', function (e) {
  var b = e.target.closest('[data-go],[data-a]');
  if (!b || b.disabled || BUSY) return;
  if (b.dataset.go) { if (b.dataset.go <= D.today) load(b.dataset.go); }
  else if (b.dataset.a === 'zip') zip();
  else if (b.dataset.a === 'collect') collect();
});
document.addEventListener('change', function (e) {
  var t = e.target;
  if (t.id === 'pick') { if (t.value && t.value <= D.today) load(t.value); else t.value = D.day; return; }
  if (t.dataset.id) { SEL[t.dataset.id] = t.checked; drawBar(); }
});
load('');
</script></body></html>`;
}
