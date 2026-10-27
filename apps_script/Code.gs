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
    (folder ? '<p><a href="' + folder.getUrl() + '">เปิดโฟลเดอร์ Google Drive ของวันนี้</a></p>' : '') +
    '</div>';
  return { subject: 'Playlist ' + day + ': ' + haveN + '/' + mail.length + ' แบบ' + (missing.length ? ' (ยังไม่ครบ)' : ' — ครบแล้ว'), html: html };
}
