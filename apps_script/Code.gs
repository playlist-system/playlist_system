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
 * หน้าเว็บ PLAYLIST INBOX (ดูว่ามาแล้วแบบไหน + ดาวน์โหลดทั้งหมดเป็น .zip): ทำให้ใช้ได้ด้วย Deploy → New deployment → Web app (ดู README)
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
  // เก็บ Playlist ทุกไฟล์ของวันนี้ แม้ไม่อยู่ในรายการ kinds ด้านล่าง (ชื่อไฟล์ต้องมีคำใน playlistWords)
  //   false = เก็บเฉพาะแบบที่อยู่ในรายการ • ไฟล์ Excel ที่ชื่อไม่มีคำเหล่านี้ (เช่น ใบแจ้งหนี้) ไม่เก็บ แค่แจ้งในอีเมลสรุป
  collectAll: true,
  // ไม่ดึงไฟล์ที่ชื่อตรงกับแบบนี้เลย (ไม่เก็บลง Drive ไม่แสดงในหน้าเว็บและอีเมลสรุป) — ว่าง = ดึงทุกไฟล์ตามด้านบน
  //   ตอนนี้: "Metro LED_… _Playlist" = ไฟล์ลำดับฉายของ Metro LED ที่ระบบเติมใบเช็คไม่ได้ใช้ (ตั้งแต่ 5.5.46)
  skipPattern: '^Metro[\\s_]*LED',
  // หน้าระบบเติมใบเช็คที่ส่ง Playlist ให้ได้ (ปุ่ม “ดึง Playlist วันนี้” ในระบบ เปิดหน้านี้ด้วย ?send=studio) — ส่งให้เฉพาะเว็บนี้เท่านั้น
  studioOrigin: 'https://playlist-system.github.io',
  // ไอคอนแท็บของหน้าเว็บนี้ (รูป .png บนเว็บของระบบ) — ว่าง = ไอคอนเดิมของ Apps Script
  iconUrl: 'https://playlist-system.github.io/playlist_system/icons/playlist-64.png',
  playlistWords: 'รายการหนัง|playlist',
  // Playlist ที่ระบบเติมใบเช็ครูปใช้ (ลำดับมีผล: ไฟล์เข้าแบบแรกที่ชื่อตรง)
  //   pattern = คำในชื่อไฟล์ • latest = ถ้าวันนี้ไม่มีฉบับใหม่ ใช้ฉบับล่าสุด
  //   source: 'line' = ไม่ได้มาทางอีเมล (เตือนในอีเมลสรุป) • extra = Playlist เพิ่มเติม ไม่นับในแบบหลัก (ใช้กับใบเช็คสื่ออื่น)
  //   unused = ระบบเติมใบเช็คไม่ได้ใช้ (เก็บไว้เฉย ๆ ไม่รวมในปุ่ม "เลือกทั้งหมด")
  //   tip = ข้อความใต้ชื่อในหน้า PLAYLIST INBOX ตอนที่วันนี้ยังไม่มา (ไม่ใส่ = "ยังไม่มีอีเมลวันนี้")
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
    // Metro LED มาทางอีเมลแล้ว (ตั้งแต่ พ.ย. 2569 จาก Media Scheduling Center มาก่อนวันออกอากาศ 1 วัน) — อีเมลแนบ 2 ไฟล์ Excel:
    //   "รายการหนังโฆษณา 7-Eleven Metro LED …" = ไฟล์ที่ระบบเติมใบเช็คใช้ • "Metro LED_… _Playlist" = ไฟล์ลำดับฉาย (ไม่ดึง — ดู skipPattern ด้านบน)
    { name: 'Metro LED', pattern: 'รายการหนัง.*Metro', tip: 'ปกติมาทางอีเมลก่อนวันฉาย 1 วัน' },
    { name: 'Rama9 Connect', pattern: 'Rama\\s*9[\\s_]*connect', extra: true },
    { name: 'Digital MRT', pattern: 'Digital[\\s_]*MRT', extra: true },
    { name: 'Gateway X', pattern: 'Gateway[\\s_]*X', extra: true },
    { name: 'Air Traveller', pattern: 'Air[\\s_]*Travell?er', extra: true },
    { name: 'The 20', pattern: 'The[\\s_]*20(?!\\d)', extra: true },
    { name: 'LA PlanB', pattern: 'LA[\\s_]*Plan[\\s_]*B', latest: true, extra: true },

    { name: 'NTW 57Blue_1', pattern: 'NTW[\\s_]*57[\\s_]*Blue[\\s_]*1(?!\\d)', extra: true },
    { name: 'NTW 57Blue_2', pattern: 'NTW[\\s_]*57[\\s_]*Blue[\\s_]*2(?!\\d)', extra: true },
    { name: 'NTW 57Red_1', pattern: 'NTW[\\s_]*57[\\s_]*Red[\\s_]*1(?!\\d)', extra: true },
    { name: 'NTW 57Red_2', pattern: 'NTW[\\s_]*57[\\s_]*Red[\\s_]*2(?!\\d)', extra: true },
    { name: 'NTW Full Only', pattern: 'NTW[\\s_]*Full', extra: true },
    { name: 'NTW Spare', pattern: 'NTW[\\s_]*Spare', extra: true },
    { name: 'Central Network', pattern: 'Central[\\s_]*Network', extra: true },
    { name: 'Parc Paragon', pattern: 'Parc[\\s_]*Paragon', extra: true },
    { name: 'Paragon Max', pattern: 'Paragon[\\s_]*Max', extra: true },
    { name: 'Siam Paragon in mall', pattern: 'Siam[\\s_]*Paragon[\\s_]*in[\\s_]*mall', extra: true },
    { name: 'Siam Paragon Facade', pattern: 'Siam[\\s_]*Paragon[\\s_]*Facade', extra: true },
    { name: 'The District EM', pattern: 'District[\\s_]*EM', extra: true },
    { name: 'The Panoramix', pattern: 'Panoramix', extra: true },
    { name: 'LIDO Connect', pattern: 'LIDO', extra: true },
  ],
};

const TZ = 'Asia/Bangkok';
const OTHER_ = 'Playlist อื่น';
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

/** แบบชื่อไฟล์ที่ไม่ดึง (CONFIG.skipPattern) */
function skipRe_() {
  return CONFIG.skipPattern ? new RegExp(CONFIG.skipPattern, 'i') : null;
}

/** คัดไฟล์ของวันนี้จากรายการไฟล์แนบ [{name, size, day(อีเมล), when, blob}] → {chosen, others} */
function pick_(atts, today) {
  const kinds = CONFIG.kinds.map(k => Object.assign({}, k, { re: new RegExp(k.pattern, 'i') }));
  const skip = skipRe_();
  const thisYear = +today.slice(0, 4);
  const word = CONFIG.collectAll && CONFIG.playlistWords ? new RegExp(CONFIG.playlistWords, 'i') : null;
  const byKind = {};
  const others = {};
  const more = {}; // Playlist ของวันนี้ที่ไม่อยู่ในรายการ (collectAll)
  for (const a of atts) {
    const name = String(a.name || '').normalize('NFC');
    if (!/\.xlsx?$/i.test(name)) continue;
    if (skip && skip.test(name)) continue; // ไม่ดึงไฟล์ที่ตั้งไว้ใน skipPattern
    const day = nameDate_(name, thisYear) || a.day;
    if (!day || day > today) continue; // ไฟล์ของวันข้างหน้า รอวันนั้น
    const kind = kinds.find(k => k.re.test(name));
    if (!kind) {
      if (day !== today) continue;
      if (word && word.test(name)) {
        const k2 = name + '|' + a.size;
        if (!more[k2]) more[k2] = { kind: OTHER_, name: name, key: k2, day: day, when: a.when, blob: a.blob, size: a.size };
      } else others[name] = 1;
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
  chosen.push(...Object.values(more));
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
  const skip = skipRe_();
  const counts = {};
  const files = [];
  const more = [];
  if (folder) {
    const it = folder.getFiles();
    while (it.hasNext()) {
      const n = it.next().getName().normalize('NFC');
      if (skip && skip.test(n)) continue; // ไฟล์ที่ดึงไว้ก่อนตั้ง skipPattern
      files.push(n);
      const k = kinds.find(x => x.re.test(n));
      if (k) counts[k.name] = (counts[k.name] || 0) + 1;
      else more.push(n);
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
    (extra.length ? '<p>Playlist เพิ่มเติม (ไม่นับใน ' + mail.length + ' แบบ): ' + extra.map(esc).join(', ') + '</p>' : '') +
    (more.length ? '<p>Playlist อื่น ๆ ที่ไม่อยู่ในรายการ (เก็บไว้ให้ด้วย ' + more.length + ' ไฟล์): ' + more.map(esc).join(', ') + '</p>' : '') +
    (multi.length ? '<p style="color:#b35a14"><b>บางแบบมีมากกว่า 1 ไฟล์:</b> ' + multi.map(k => esc(k) + ' (' + counts[k] + ' ไฟล์)').join(', ') + ' — ตรวจว่าต้องใช้ฉบับไหนก่อนลากลงระบบ</p>' : '') +
    (trashed && trashed.length ? '<p style="color:#586274">ได้ฉบับของวันนี้แล้ว ย้ายฉบับเก่าไปถังขยะของ Drive: ' + trashed.map(esc).join(', ') + '</p>' : '') +
    (others.length ? '<p style="color:#586274">ไฟล์ Excel อื่นของวันนี้ (ไม่ได้เก็บ เพราะชื่อไม่เหมือน Playlist):</p><ul style="color:#586274">' + others.map(n => '<li>' + esc(n) + '</li>').join('') + '</ul>' : '') +
    (webUrl_() ? '<p><a href="' + webUrl_() + '"><b>เปิด PLAYLIST INBOX</b></a> — ดูว่ามาแล้วแบบไหน และดาวน์โหลดทั้งหมดเป็นไฟล์ .zip ไฟล์เดียว</p>' : '') +
    (folder ? '<p><a href="' + folder.getUrl() + '">เปิดโฟลเดอร์ Google Drive ของวันนี้</a></p>' : '') +
    '</div>';
  return { subject: 'Playlist ' + day + ': ' + haveN + '/' + mail.length + ' แบบ' + (missing.length ? ' (ยังไม่ครบ)' : ' — ครบแล้ว'), html: html };
}

// ---------- หน้าเว็บ PLAYLIST INBOX (เดิมชื่อ "Playlist รายวัน") ----------
// ทำให้ใช้ได้: Deploy → New deployment → ประเภท Web app → Execute as: Me → Who has access: Only myself (หรือทุกคนในบริษัท) → Deploy
// หน้าเว็บอ่านเฉพาะโฟลเดอร์ "Playlist รายวัน" ของสคริปต์นี้ ไม่แก้ไข/ลบไฟล์

/** เปิดหน้าเว็บ */
function doGet(e) {
  try {
    const u = ScriptApp.getService().getUrl();
    if (/\/exec$/.test(u || '') && PropertiesService.getScriptProperties().getProperty(PROP_WEB) !== u) PropertiesService.getScriptProperties().setProperty(PROP_WEB, u);
  } catch (e) {}
  const send = !!(e && e.parameter && e.parameter.send === 'studio'); // เปิดจากปุ่ม “ดึง Playlist วันนี้” ในระบบ → ส่งไฟล์กลับให้ระบบเอง
  return HtmlService.createHtmlOutput(page_(send)).setTitle(send ? 'ส่ง Playlist เข้า PHOTO CHECK STUDIO' : 'PLAYLIST INBOX').setFaviconUrl(CONFIG.iconUrl || 'https://playlist-system.github.io/playlist_system/icons/playlist-64.png').addMetaTag('viewport', 'width=device-width, initial-scale=1');
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
  const skip = skipRe_();
  const thisYear = +day.slice(0, 4);
  const files = [];
  if (folder) {
    const it = folder.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      const n = f.getName().normalize('NFC');
      if (skip && skip.test(n)) continue; // ไฟล์ที่ดึงไว้ก่อนตั้ง skipPattern
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
    kinds: CONFIG.kinds.map(k => ({ name: k.name, line: k.source === 'line', extra: !!k.extra, latest: !!k.latest, unused: !!k.unused, tip: k.tip || '' })),
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

function page_(send) {
  return `<!doctype html>
<html lang="th"><head><base target="_top"><meta charset="utf-8">
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@400;500;600;700&family=Prompt:wght@600;700&display=swap">
<style>
:root{color-scheme:light;--paper:#EDF1F5;--surface:#FFFFFF;--soft:#F1F4F8;--chip:#F7F9FB;--ink:#0E1A2F;--ink2:#33405A;--muted:#5A667E;--line:#DCE2EB;--line2:#E9EDF3;--line3:#B9C3D2;--box:#8C97A9;--dis:#B8C1CF;
 --ok:#047857;--okBtn:#047857;--okInk:#FFFFFF;--wait:#8A5A00;--waitSoft:#FFF5D9;--waitLine:#E2B655;--waitInk:#6B4A00;--info:#1D4F91;--infoSoft:#E8EFF9;--infoLine:#A9C2E6;--err:#B3261E;
 --sh:0 1px 2px rgba(14,26,47,.04),0 16px 40px -18px rgba(14,26,47,.18);--hd:"Prompt","IBM Plex Sans Thai","Leelawadee UI",Tahoma,sans-serif;--ui:"IBM Plex Sans Thai","Leelawadee UI",Tahoma,sans-serif}
@media (prefers-color-scheme:dark){:root{color-scheme:dark;--paper:#0B1220;--surface:#131C2E;--soft:#1B2740;--chip:#18233A;--ink:#E8EDF6;--ink2:#C3CCDD;--muted:#95A2BA;--line:#26334D;--line2:#1F2A41;--line3:#3A4A68;--box:#6B7A94;--dis:#4A5873;
 --ok:#4ADE9B;--okBtn:#34D399;--okInk:#052E1F;--wait:#F2C46B;--waitSoft:#2E240C;--waitLine:#8A6A1C;--waitInk:#F5D58E;--info:#8CB6EF;--infoSoft:#162A43;--infoLine:#2F4F7A;--err:#F28B82;--sh:none}}
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--paper);color:var(--ink);font:15px/1.6 var(--ui)}
button,input{font:inherit;color:inherit}
a{color:inherit}
svg{flex:none}
.top{background:var(--surface);border-bottom:1px solid var(--line)}
.top .in{max-width:1160px;margin:0 auto;padding:14px 24px;display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:12px 20px}
.brand{display:flex;align-items:center;gap:12px;min-width:0}
.logo{width:42px;height:42px;border-radius:12px;background:linear-gradient(135deg,#059669,#0EA5E9);display:grid;place-items:center;flex:none}
.brand b{display:block;font:700 19px/1.25 var(--hd);letter-spacing:.03em}
.brand small{display:block;font-size:12px;color:var(--muted)}
.tools{display:flex;flex-wrap:wrap;align-items:center;gap:10px}
.date{display:flex;align-items:center;gap:2px;padding:3px;border:1px solid var(--line);border-radius:999px;background:var(--surface)}
.nav{width:38px;height:38px;border:0;border-radius:999px;background:transparent;display:grid;place-items:center;cursor:pointer}
.nav:hover:not(:disabled){background:var(--soft)}
.nav:disabled{color:var(--dis);cursor:default}
.cur{position:relative;display:flex;align-items:center;gap:8px;height:38px;padding:0 14px;border-radius:999px;background:var(--soft);font-weight:600;font-size:14px;white-space:nowrap;cursor:pointer}
.cur input{position:absolute;inset:0;width:100%;height:100%;opacity:0;border:0;padding:0;margin:0;cursor:pointer}
.cur input::-webkit-calendar-picker-indicator{position:absolute;inset:0;width:100%;height:100%;margin:0;opacity:0;cursor:pointer}
.cur:focus-within{outline:3px solid var(--info);outline-offset:2px}
.badge{padding:0 8px;border-radius:999px;background:var(--ink);color:var(--surface);font-size:12px;line-height:20px;font-weight:500}
.pill{display:inline-flex;align-items:center;justify-content:center;gap:8px;min-height:44px;padding:0 18px;border:1px solid var(--line);border-radius:999px;background:var(--surface);font-weight:600;font-size:14px;text-decoration:none;cursor:pointer;white-space:nowrap}
.pill.dark{background:var(--ink);border-color:var(--ink);color:var(--surface)}
.pill.ok{background:var(--okBtn);border-color:var(--okBtn);color:var(--okInk);font-size:15px;min-height:50px;padding:0 24px}
.pill:disabled{opacity:.5;cursor:default}
.pill:focus-visible,.nav:focus-visible,.lk:focus-visible{outline:3px solid var(--info);outline-offset:2px}
main{max-width:1160px;margin:0 auto;padding:40px 24px 64px;display:flex;flex-direction:column;gap:24px;transition:opacity .15s}
main.dim{opacity:.55}
.banner{display:flex;gap:10px;align-items:flex-start;padding:12px 16px;border-radius:14px;background:var(--waitSoft);border:1.5px dashed var(--waitLine);color:var(--waitInk);font-size:14px}
.banner.info{background:var(--infoSoft);border:1px solid var(--infoLine);color:var(--ink2)}
.hero{display:flex;flex-wrap:wrap;align-items:flex-end;justify-content:space-between;gap:20px 32px}
.hero h1{margin:0;font:700 44px/1.15 var(--hd);letter-spacing:-.01em}
.hero p{margin:10px 0 0;color:var(--ink2);font-size:16px;max-width:58ch}
.hero p b{color:var(--wait)}
.acts{display:flex;flex-wrap:wrap;gap:10px;align-items:center}
.board{background:var(--surface);border:1px solid var(--line);border-radius:22px;padding:8px 24px;box-shadow:var(--sh)}
.grp{display:grid;grid-template-columns:112px minmax(0,1fr);gap:12px 20px;padding:18px 0;align-items:start}
.grp+.grp{border-top:1px solid var(--line2)}
.grp h2{margin:0;padding-top:10px;font:600 15px/1.3 var(--hd);color:var(--ink2)}
.grp h2 span{display:block;font:400 12px/1.5 var(--ui);color:var(--muted);margin-top:2px}
.tiles{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:10px}
.tile{position:relative;display:flex;flex-direction:column;gap:4px;min-height:80px;padding:12px 14px;border-radius:14px;min-width:0}
.tile.wide{grid-column:span 2}
.tile .nm{font:600 16px/1.3 var(--hd);padding-right:28px;overflow-wrap:anywhere}
.tile .st{display:flex;align-items:center;gap:6px;font-size:13px;font-weight:600}
.tile .hint{font-size:12px;line-height:1.45}
label.tile{background:var(--surface);border:1px solid var(--line);cursor:pointer}
label.tile:has(input:checked){border-color:var(--line3)}
label.tile .st{color:var(--ok)}
label.tile .st.info{color:var(--info)}
label.tile .hint{color:var(--muted)}
.tile input{position:absolute;opacity:0;width:1px;height:1px;margin:0}
.tile .box{position:absolute;top:12px;right:12px;width:20px;height:20px;border-radius:6px;border:1.5px solid var(--box);background:var(--surface);display:grid;place-items:center}
.tile .box svg{opacity:0}
.tile input:checked+.box{background:var(--ink);border-color:var(--ink)}
.tile input:checked+.box svg{opacity:1}
.tile input:focus-visible+.box{outline:3px solid var(--info);outline-offset:2px}
.tile.wait{background:var(--waitSoft);border:1.5px dashed var(--waitLine)}
.tile.wait .st{color:var(--wait)}
.tile.wait .hint{color:var(--waitInk)}
.tile.line{background:var(--infoSoft);border:1.5px dashed var(--infoLine)}
.tile.line .st{color:var(--info)}
.tile.line .hint{color:var(--ink2)}
.card{background:var(--surface);border:1px solid var(--line);border-radius:22px;padding:20px 24px;display:flex;flex-direction:column;gap:14px}
.card .hd{display:flex;flex-wrap:wrap;align-items:baseline;justify-content:space-between;gap:6px 16px}
.card h2{margin:0;font:600 17px/1.3 var(--hd)}
.card .hd p{margin:0;font-size:13px;color:var(--muted)}
.sub{display:flex;flex-wrap:wrap;align-items:baseline;justify-content:space-between;gap:4px 12px;font-size:13px;color:var(--muted);font-weight:600}
.sep{height:1px;background:var(--line2)}
.chips{display:flex;flex-wrap:wrap;gap:8px}
.chip{display:inline-flex;align-items:center;gap:8px;min-height:40px;max-width:100%;padding:4px 14px 4px 10px;border:1px solid var(--line);border-radius:999px;background:var(--chip);font-size:14px;line-height:1.4;cursor:pointer}
.chip input{width:16px;height:16px;margin:0;accent-color:var(--ink);flex:none;cursor:pointer}
.chip.dim{color:var(--muted)}
.chip small{color:var(--muted);font-size:12px}
.lk{border:0;background:none;padding:6px 2px;color:var(--ink);font-weight:600;font-size:13px;text-decoration:underline;text-underline-offset:3px;cursor:pointer}
.foot{display:flex;flex-wrap:wrap;justify-content:space-between;gap:8px 24px;font-size:13px;color:var(--muted)}
.dot{display:inline-block;width:8px;height:8px;border-radius:99px;background:#1FA463;margin-right:8px;vertical-align:1px}
.dot.off{background:var(--dis)}
.loading{max-width:1160px;margin:0 auto;padding:80px 24px;text-align:center;color:var(--muted)}
.errc{max-width:640px;margin:60px auto;padding:24px;background:var(--surface);border:1px solid var(--line);border-radius:18px}
.errc h2{margin:0 0 6px;font:600 18px/1.3 var(--hd)}
.errc p{margin:0;color:var(--muted)}
/* หน้าส่งเข้าระบบ (เปิดจากปุ่ม “ดึง Playlist วันนี้”) */
.sendp{min-height:100vh;display:grid;place-items:center;padding:24px 16px}
.pp{width:100%;max-width:480px;background:var(--surface);border:1px solid var(--line);border-radius:18px;padding:28px 32px 22px;display:flex;flex-direction:column;gap:24px;box-shadow:0 20px 60px -20px rgba(14,26,47,.35)}
.pp .brand b{font-size:17px}
.pp h1{margin:0;font:700 28px/1.25 var(--hd)}
.pp .when{margin:4px 0 16px;color:var(--muted);font-size:14px}
.prog{height:10px;border-radius:99px;background:var(--line2);overflow:hidden}
.prog span{display:block;height:100%;border-radius:99px;background:var(--okBtn);transition:width .3s}
.steps{margin:0;padding:0;list-style:none;display:flex;flex-direction:column;gap:12px}
.steps li{display:flex;align-items:center;gap:12px;color:var(--muted)}
.steps .i{width:24px;height:24px;border-radius:99px;display:grid;place-items:center;flex:none;border:3px solid var(--line)}
.steps .done{color:var(--ink)}
.steps .done .i{background:var(--okBtn);border-color:var(--okBtn)}
.steps .now{color:var(--ink);font-weight:600}
.steps .now .i{border-color:var(--okBtn)}
.ppf{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:center;gap:6px 12px;padding-top:14px;border-top:1px solid var(--line2);font-size:13px;color:var(--muted)}
.toast{position:fixed;left:50%;bottom:24px;transform:translateX(-50%);max-width:min(92vw,620px);background:var(--ink);color:var(--surface);padding:10px 16px;border-radius:12px;font-size:14px;opacity:0;pointer-events:none;transition:opacity .2s;z-index:9}
.toast.on{opacity:1}
.toast.err{background:var(--err);color:#fff}
@media (max-width:1000px){.tiles{grid-template-columns:repeat(4,minmax(0,1fr))}}
@media (max-width:640px){
 .top .in{padding:12px 16px}
 .tools,.date{width:100%}
 .date{justify-content:space-between}
 .tools .pill{flex:1 1 auto}
 main{padding:24px 16px 48px;gap:18px}
 .hero h1{font-size:32px}
 .acts{width:100%}
 .acts .pill{flex:1 1 0}
 .acts .pill.ok{flex:1 1 100%}
 .board{padding:4px 16px;border-radius:18px}
 .grp{grid-template-columns:1fr;gap:10px;padding:16px 0}
 .grp h2{padding-top:0}
 .grp h2 span{display:inline;margin-left:6px}
 .tiles{grid-template-columns:repeat(2,minmax(0,1fr))}
 .tile.wide{grid-column:auto}
 .card{padding:18px 16px;border-radius:18px}
 .pp{padding:24px 20px 18px}
}
@media (prefers-reduced-motion:reduce){*{transition:none!important}}
</style></head><body>
<div id="app"><p class="loading">กำลังโหลด…</p></div>
<div class="toast" id="toast" role="status" aria-live="polite"></div>
<script>
var D = null, SEL = {}, BUSY = '', TT = 0, PROG = 0;
var SEND = ${send ? 'true' : 'false'}, STUDIO = ${JSON.stringify(CONFIG.studioOrigin || '')}, SENT = false, SENDMSG = '', SENDOK = false, SENTN = 0, SENDINFO = false;
var TH = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
var WD = ['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'];
function sv(p, s, c, w) { return '<svg aria-hidden="true" width="' + s + '" height="' + s + '" viewBox="0 0 24 24" fill="none" stroke="currentColor"' + (c ? ' style="color:' + c + '"' : '') + ' stroke-width="' + (w || 2) + '" stroke-linecap="round" stroke-linejoin="round">' + p + '</svg>'; }
var P = { check: '<path d="M5 12.5l4.5 4.5L19 7.5"/>', clock: '<circle cx="12" cy="12" r="8.5"/><path d="M12 8v4.5l3 1.8"/>', left: '<path d="M15 6l-6 6 6 6"/>', right: '<path d="M9 6l6 6-6 6"/>',
  cal: '<rect x="3.5" y="5" width="17" height="15" rx="3"/><path d="M3.5 10h17M8 3v4M16 3v4"/>', ext: '<path d="M7 17L17 7M9 7h8v8"/>', info: '<circle cx="12" cy="12" r="8.5"/><path d="M12 11v5M12 8h.01"/>',
  refresh: '<path d="M20 11a8 8 0 0 0-14.6-4.5M4 4v4h4M4 13a8 8 0 0 0 14.6 4.5M20 20v-4h-4"/>', folder: '<path d="M3.5 7.5A2.5 2.5 0 0 1 6 5h4l2 2h6a2.5 2.5 0 0 1 2.5 2.5v7A2.5 2.5 0 0 1 18 19H6a2.5 2.5 0 0 1-2.5-2.5z"/>',
  down: '<path d="M12 4v11M7 10.5l5 5 5-5M5 20h14"/>', list: '<rect x="4" y="3.5" width="16" height="17" rx="3"/><path d="M8 9h8M8 12.5h8M8 16h5"/>' };
function $(s) { return document.querySelector(s); }
function esc(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function dt(d) { return new Date(d + 'T12:00:00Z'); }
function thLong(d) { var t = dt(d); return WD[t.getUTCDay()] + ' ' + t.getUTCDate() + ' ' + TH[t.getUTCMonth()] + ' ' + (t.getUTCFullYear() + 543); }
function thShort(d) { var t = dt(d); return t.getUTCDate() + ' ' + TH[t.getUTCMonth()]; }
function shift(d, n) { var t = dt(d); t.setUTCDate(t.getUTCDate() + n); return t.toISOString().slice(0, 10); }
function kb(n) { return n >= 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB'; }
function noext(n) { var d = n.lastIndexOf('.'); return d > 0 ? n.slice(0, d) : n; }
function andList(L) { return L.length < 2 ? L.join('') : L.slice(0, -1).join(', ') + ' กับ ' + L[L.length - 1]; }
function toast(m, err) { var t = $('#toast'); t.textContent = m; t.className = 'toast on' + (err ? ' err' : ''); clearTimeout(TT); TT = setTimeout(function () { t.className = 'toast'; }, err ? 9000 : 6000); }
function call(fn, args, ok) {
  google.script.run.withSuccessHandler(ok).withFailureHandler(function (e) {
    var m = (e && e.message) || String(e); BUSY = '';
    if (!D) $('#app').innerHTML = '<div class="errc"><h2>เปิดข้อมูลไม่ได้</h2><p>' + esc(m) + '</p></div>';
    else draw();
    toast(m, 1);
  })[fn].apply(null, args);
}
function load(day) { BUSY = 'load'; draw(); call('webDay', [day || ''], function (d) { D = d; BUSY = ''; defaults(); draw(); if (SEND && !SENT) { SENT = true; sendStudio(); } }); }
/* เปิดจากปุ่ม “ดึง Playlist วันนี้” ในระบบ: รวม Playlist ที่ติ๊กไว้ (แบบหลัก) เป็น .zip แล้วส่งกลับให้หน้าระบบ (ส่งให้เฉพาะ studioOrigin) */
function opener_() { try { return window.top.opener || null; } catch (e) { return null; } }
function sendStudio() {
  var w = opener_();
  if (!w || !STUDIO) { SENDINFO = true; SENDMSG = 'หน้านี้ไม่ได้เปิดจากปุ่ม “ดึง Playlist วันนี้” ในระบบ (หรือเบราว์เซอร์ไม่ให้ส่งกลับ) จึงยังไม่ได้ส่งไฟล์เข้าระบบ — ถ้าจะส่ง กดปุ่มนั้นใน PHOTO CHECK STUDIO หรือกด “ดาวน์โหลด” แล้วลากไฟล์ .zip ลงขั้นที่ 1'; draw(); return; }
  var g = byKind(), main = mainKinds(), miss = main.filter(function (k) { return !g.by[k.name]; }).map(function (k) { return k.name; });
  if (!picked().length) { try { w.postMessage({ psPL: 1, ok: false, day: D.day, err: 'ยังไม่มี Playlist ของวันนี้ในโฟลเดอร์' }, STUDIO); } catch (e) {} SENDMSG = 'ยังไม่มี Playlist ของวันนี้ — รอระบบดึงจากอีเมล (ทุกชั่วโมง) หรือกด “ดึงตอนนี้”'; draw(); return; }
  zip(function (name, blob, n) {
    try { w.postMessage({ psPL: 1, ok: true, day: D.day, name: name, blob: blob, n: n, have: main.length - miss.length, total: main.length, miss: miss, extra: picked().filter(isExtra).length }, STUDIO); }
    catch (e) { SENDMSG = 'ส่งเข้าระบบไม่ได้: ' + (e && e.message || e) + ' — กด “ดาวน์โหลด” แล้วลากไฟล์ .zip เองแทน'; draw(); return; }
    SENDOK = true; SENTN = n; SENDMSG = ''; draw();
    setTimeout(function () { try { window.top.close(); } catch (e) {} }, 800);
  });
}
function byKind() { var by = {}, oth = []; D.files.forEach(function (f) { if (f.kind) (by[f.kind] = by[f.kind] || []).push(f); else oth.push(f); }); return { by: by, oth: oth }; }
function mainKinds() { return D.kinds.filter(function (k) { return !k.extra && !k.line; }); }
/* ค่าเริ่มต้น: ติ๊ก Playlist ที่ระบบใช้ ทั้งแบบหลักและ Playlist เพิ่มเติม (แบบละ 1 ไฟล์ = ฉบับที่เข้าโฟลเดอร์ล่าสุด)
   ไม่ติ๊ก: แบบที่ระบบไม่ได้ใช้ (unused) • ฉบับเก่า • ไฟล์ที่ไม่อยู่ในรายการ — ปุ่ม “ดึง Playlist วันนี้” ในระบบจึงได้ Playlist เพิ่มเติมด้วยทุกครั้ง */
function defaults() {
  SEL = {}; var g = byKind();
  D.kinds.forEach(function (k) { var L = g.by[k.name]; if (!L || k.unused) return; SEL[L[L.length - 1].id] = true; });
}
function isExtra(f) { return D.kinds.some(function (k) { return k.extra && k.name === f.kind; }); }
function picked() { return D ? D.files.filter(function (f) { return SEL[f.id]; }) : []; }
function extraFiles(g) { var L = []; D.kinds.forEach(function (k) { if (k.extra && !k.unused && g.by[k.name]) L = L.concat(g.by[k.name]); }); return L; }
function selBtn(L, grp) { var all = L.length && L.every(function (f) { return SEL[f.id]; }); return '<button class="lk" data-a="sel" data-g="' + grp + '">' + (all ? 'ไม่เลือก' : 'เลือกทั้งหมด') + '</button>'; }
/* จัดกลุ่มแบบหลักตามคำแรกของชื่อ (เช่น “แยกจอ 40Blue” → แยกจอ, “BKJ1” → BKJ) กลุ่มที่มีแบบเดียวรวมเป็น “สื่ออื่น” */
function groups(kinds) {
  var key = function (n) { return n.indexOf(' ') > 0 ? n.split(' ')[0] : n.replace(/[\\s_-]*\\d+$/, ''); };
  var cnt = {}, order = [], G = {};
  kinds.forEach(function (k) { var x = key(k.name); cnt[x] = (cnt[x] || 0) + 1; });
  kinds.forEach(function (k) {
    var x = cnt[key(k.name)] > 1 ? key(k.name) : 'สื่ออื่น';
    if (!G[x]) { G[x] = { name: x, kinds: [] }; if (x !== 'สื่ออื่น') order.push(x); }
    G[x].kinds.push(k);
  });
  if (G['สื่ออื่น']) order.push('สื่ออื่น');
  return order.map(function (x) {
    var g = G[x], strip = x !== 'สื่ออื่น' && g.kinds.every(function (k) { return k.name.indexOf(x + ' ') === 0; });
    g.small = g.kinds.length >= 4;
    g.kinds.forEach(function (k) { k.short = strip ? k.name.slice(x.length + 1) : k.name; });
    return g;
  });
}
function tile(k, L, isToday, wide) {
  var w = wide ? ' wide' : '';
  if (L) {
    var f = L[L.length - 1];
    var st = f.old ? '<span class="st info">' + sv(P.info, 15, 0, 2.4) + 'ฉบับ ' + esc(f.old) + '</span>' : '<span class="st">' + sv(P.check, 15, 0, 3) + esc(f.time) + '</span>';
    var hint = f.old ? 'วันนี้ยังไม่มีฉบับใหม่ ใช้ฉบับล่าสุดไปก่อน' : L.length > 1 ? 'มี ' + L.length + ' ฉบับ ติ๊กไว้ให้ฉบับล่าสุด' : '';
    return '<label class="tile' + w + '" title="' + esc(f.name + ' · ' + kb(f.size)) + '"><input type="checkbox" data-id="' + esc(f.id) + '"' + (SEL[f.id] ? ' checked' : '') + ' aria-label="' + esc(k.name) + '">' +
      '<span class="box" aria-hidden="true">' + sv(P.check, 13, 'var(--surface)', 3.2) + '</span><span class="nm">' + esc(k.short) + '</span>' + st + (hint ? '<span class="hint">' + esc(hint) + '</span>' : '') + '</label>';
  }
  if (k.line) return '<div class="tile line' + w + '"><span class="nm">' + esc(k.short) + '</span><span class="st">' + sv(P.info, 15, 0, 2.4) + 'มาทาง LINE</span><span class="hint">ดาวน์โหลดจาก LINE แล้วลากลงระบบพร้อมไฟล์ .zip</span></div>';
  var h = !isToday ? 'ไม่มีไฟล์ของวันนั้นในโฟลเดอร์' : k.tip ? k.tip : k.latest ? 'ไม่ได้ส่งทุกวัน ยังไม่มีฉบับไหนในอีเมล' : 'ยังไม่มีอีเมลวันนี้';
  return '<div class="tile wait' + w + '"><span class="nm">' + esc(k.short) + '</span><span class="st">' + sv(P.clock, 15, 0, 2.4) + (isToday ? 'รออยู่' : 'ไม่มี') + '</span><span class="hint">' + esc(h) + '</span></div>';
}
function chip(f, label, extra) {
  return '<label class="chip' + (extra && extra.dim ? ' dim' : '') + '" title="' + esc(f.name + ' · ' + kb(f.size)) + '"><input type="checkbox" data-id="' + esc(f.id) + '"' + (SEL[f.id] ? ' checked' : '') + '><span>' + esc(label) + (extra && extra.note ? ' <small>' + esc(extra.note) + '</small>' : '') + '</span></label>';
}
function brand(sub) { return '<div class="brand"><div class="logo">' + sv(P.list, 23, '#fff', 2.2) + '</div><div><b>PLAYLIST INBOX</b><small>' + esc(sub) + '</small></div></div>'; }
function actsHtml() {
  var Pk = picked(), sz = 0, isToday = D.day === D.today, url = D.folderUrl || D.parentUrl;
  Pk.forEach(function (f) { sz += f.size; });
  return (isToday ? '<button class="pill" data-a="collect"' + (BUSY ? ' disabled' : '') + '>' + sv(P.refresh, 16) + (BUSY === 'collect' ? 'กำลังดึง…' : 'ดึงตอนนี้') + '</button>' : '') +
    '<a class="pill" href="' + esc(url) + '" target="_blank" rel="noopener">' + sv(P.folder, 16) + 'เปิดใน Drive</a>' +
    '<button class="pill ok" data-a="zip"' + (!Pk.length || BUSY ? ' disabled' : '') + ' title="' + (Pk.length ? 'รวม ' + kb(sz) : 'ยังไม่ได้เลือกไฟล์') + '">' + sv(P.down, 18, 0, 2.2) +
    (BUSY === 'zip' ? 'กำลังเตรียมไฟล์ ' + PROG + '/' + Pk.length + '…' : Pk.length ? 'ดาวน์โหลด ' + Pk.length + ' ไฟล์ (.zip)' : 'ยังไม่ได้เลือกไฟล์') + '</button>';
}
function sendView() { return SEND && !(SENDMSG && !SENDOK); }
function drawSend() {
  var Pk = D ? picked() : [], n = Pk.length, nx = D ? Pk.filter(isExtra).length : 0, sz = 0, g = D ? byKind() : null;
  Pk.forEach(function (f) { sz += f.size; });
  var miss = D ? mainKinds().filter(function (k) { return !g.by[k.name]; }).map(function (k) { return k.name; }) : [];
  var zipping = BUSY === 'zip', pct = !D ? 8 : SENDOK ? 100 : zipping ? 15 + Math.round(75 * PROG / Math.max(1, n)) : 12;
  var title = !D ? 'กำลังเตรียม Playlist ของวันนี้…' : SENDOK ? 'ส่ง ' + SENTN + ' ไฟล์เข้าระบบแล้ว' : 'กำลังส่ง ' + n + ' ไฟล์เข้าระบบ';
  var step = function (txt, s) { return '<li class="' + s + '"><span class="i">' + (s === 'done' ? sv(P.check, 13, 'var(--okInk)', 3.2) : '') + '</span>' + txt + '</li>'; };
  var h = '<div class="sendp"><div class="pp">' + brand('ส่ง Playlist เข้า PHOTO CHECK STUDIO') +
    '<div><h1>' + esc(title) + '</h1><p class="when">Playlist ' + (D ? 'วัน' + esc(thLong(D.day)) + (n ? ' — แบบหลัก ' + (n - nx) + ' ไฟล์' + (nx ? ' + เพิ่มเติม ' + nx + ' ไฟล์' : '') : '') : 'ของวันนี้') + '</p><div class="prog" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="' + pct + '"><span style="width:' + pct + '%"></span></div></div>' +
    '<ol class="steps">' + step('หาไฟล์ของวันนี้ใน Drive', D ? 'done' : 'now') +
    step('รวมเป็นไฟล์ .zip' + (n ? ' (' + kb(sz) + ')' : '') + (zipping ? ' ' + PROG + '/' + n : ''), SENDOK ? 'done' : zipping ? 'now' : '') +
    step('ส่งเข้า PHOTO CHECK STUDIO', SENDOK ? 'done' : '') + '</ol>';
  if (miss.length) h += '<div class="banner">' + sv(P.clock, 18, 0, 2.4) + '<span>ยังขาด ' + andList(miss.map(function (n) { return '<b>' + esc(n) + '</b>'; })) + ' ส่งที่มีไปก่อน มาครบแล้วกดดึงอีกครั้ง</span></div>';
  h += '<div class="ppf"><span>' + (SENDOK ? 'หน้านี้จะปิดเอง ถ้าไม่ปิด ปิดเองได้เลย' : 'หน้านี้ปิดเองเมื่อส่งเสร็จ') + '</span>' + (D && !SENDOK ? '<button class="lk" data-a="zip"' + (BUSY ? ' disabled' : '') + '>ดาวน์โหลด .zip แทน</button>' : '') + '</div></div></div>';
  $('#app').innerHTML = h;
}
function draw() {
  if (sendView()) return drawSend();
  var app = $('#app');
  if (!D) { app.innerHTML = '<p class="loading">กำลังโหลด…</p>'; return; }
  var g = byKind(), day = D.day, isToday = day === D.today;
  var main = mainKinds(), line = D.kinds.filter(function (k) { return k.line; });
  var have = main.filter(function (k) { return g.by[k.name]; }), miss = main.filter(function (k) { return !g.by[k.name]; });
  var last = D.files.length ? D.files[D.files.length - 1] : null;
  var h = '<header class="top"><div class="in">' + brand('ส่วนหนึ่งของ PHOTO CHECK STUDIO') + '<div class="tools"><div class="date">' +
    '<button class="nav" data-go="' + shift(day, -1) + '" aria-label="วันก่อนหน้า" title="วันก่อนหน้า">' + sv(P.left, 17, 0, 2.2) + '</button>' +
    '<label class="cur" title="เลือกวันที่">' + sv(P.cal, 16) + esc(thLong(day)) + (isToday ? '<span class="badge">วันนี้</span>' : '') +
    '<input type="date" id="pick" value="' + day + '" max="' + D.today + '" aria-label="เลือกวันที่"></label>' +
    '<button class="nav" data-go="' + shift(day, 1) + '" aria-label="วันถัดไป" title="วันถัดไป"' + (isToday ? ' disabled' : '') + '>' + sv(P.right, 17, 0, 2.2) + '</button></div>' +
    (isToday ? '' : '<button class="pill" data-go="' + D.today + '">กลับไปวันนี้</button>') +
    '<a class="pill dark" href="' + esc(D.studio) + '" target="_blank" rel="noopener">เปิด PHOTO CHECK STUDIO' + sv(P.ext, 15, 0, 2.2) + '</a></div></div></header>';
  h += '<main' + (BUSY === 'load' ? ' class="dim"' : '') + '>';
  if (SEND && SENDMSG) h += '<div class="banner' + (SENDINFO ? ' info' : '') + '" role="' + (SENDINFO ? 'status' : 'alert') + '">' + sv(P.info, 18, 0, 2.4) + '<span>' + esc(SENDMSG) + '</span></div>';
  // หัวข้อ + ปุ่ม
  var h1, sub;
  if (!D.files.length) { h1 = isToday ? 'ยังไม่มี Playlist ของวันนี้' : 'ไม่มี Playlist ของวันที่ ' + thShort(day); sub = isToday ? (D.auto === false ? 'ดึงอัตโนมัติปิดอยู่ กด “ดึงตอนนี้” หรือเปิดด้วย setup ในหน้าโค้ด' : 'ระบบดึงจากอีเมลให้เองทุกชั่วโมง ' + D.startHour + ':00–' + D.endHour + ':59 หรือกด “ดึงตอนนี้”') : 'ระบบไม่ได้เก็บ Playlist ของวันนั้นไว้ในโฟลเดอร์'; }
  else if (!miss.length) { h1 = 'ครบ ' + main.length + ' แบบแล้ว'; sub = 'ดาวน์โหลดแล้วลากไฟล์ .zip ลงขั้นที่ 1 ของ PHOTO CHECK STUDIO ได้เลย'; }
  else {
    h1 = 'มาแล้ว ' + have.length + ' จาก ' + main.length + ' แบบ';
    var names = miss.map(function (k) { return '<b>' + esc(k.name) + '</b>'; });
    sub = 'ยังขาด ' + (miss.length > 3 ? miss.length + ' แบบ: ' + names.join(', ') : andList(names)) + (isToday && D.auto !== false ? ' ระบบดึงจากอีเมลให้เองทุกชั่วโมงจนถึง ' + D.endHour + ':59' : '');
  }
  h += '<section class="hero"><div><h1>' + esc(h1) + '</h1><p>' + (sub.indexOf('<b>') >= 0 ? sub : esc(sub)) + '</p></div><div class="acts" id="acts">' + actsHtml() + '</div></section>';
  // กระดานแบบหลัก
  h += '<section class="board" aria-label="Playlist ที่ระบบเติมใบเช็คใช้">';
  groups(main.concat(line)).forEach(function (gr) {
    h += '<div class="grp"><h2>' + esc(gr.name) + (gr.small ? '<span>' + gr.kinds.length + ' แบบ</span>' : '') + '</h2><div class="tiles">';
    gr.kinds.forEach(function (k) { h += tile(k, g.by[k.name], isToday, !gr.small); });
    h += '</div></div>';
  });
  h += '</section>';
  // เพิ่มเติม / ฉบับเก่า / อื่น ๆ
  var extra = D.kinds.filter(function (k) { return k.extra && g.by[k.name]; });
  var older = []; main.concat(line).forEach(function (k) { var L = g.by[k.name]; if (L && L.length > 1) L.slice(0, -1).forEach(function (f) { older.push({ f: f, k: k }); }); });
  if (extra.length || older.length || g.oth.length) {
    h += '<section class="card">';
    if (extra.length) {
      h += '<div class="hd"><h2>Playlist เพิ่มเติม</h2><p>ใช้กับใบเช็คสื่ออื่น เช่น CCTV, MRT, NTW, Paragon — ติ๊กไว้ให้แล้ว ' + selBtn(extraFiles(g), 'extra') + '</p></div><div class="chips">';
      extra.forEach(function (k) { var L = g.by[k.name]; L.forEach(function (f, i) { h += chip(f, k.name, { dim: k.unused, note: [k.unused ? '(ระบบไม่ได้ใช้)' : '', L.length > 1 ? f.time : ''].filter(Boolean).join(' ') }); }); });
      h += '</div>';
    }
    if (older.length) {
      if (extra.length) h += '<div class="sep"></div>';
      h += '<div class="sub"><span>ฉบับเก่าของวันนี้ (มีฉบับใหม่กว่าแล้ว)</span></div><div class="chips">';
      older.forEach(function (o) { h += chip(o.f, o.k.name, { note: o.f.time }); });
      h += '</div>';
    }
    if (g.oth.length) {
      if (extra.length || older.length) h += '<div class="sep"></div>';
      h += '<div class="sub"><span>ไม่อยู่ในรายการ แต่ระบบเก็บไว้ให้</span>' + selBtn(g.oth, 'oth') + '</div><div class="chips">';
      g.oth.forEach(function (f) { h += chip(f, noext(f.name)); });
      h += '</div>';
    }
    h += '</section>';
  }
  // ท้ายหน้า
  var left = isToday ? (D.auto === false ? '<span class="dot off"></span>ดึงอัตโนมัติปิดอยู่ (เปิดด้วย setup ในหน้าโค้ด)' : '<span class="dot"></span>ดึงอัตโนมัติเปิดอยู่ ตรวจอีเมลทุกชั่วโมง ' + D.startHour + ':00–' + D.endHour + ':59') : '';
  if (last) left += (left ? ' • ' : '') + 'ไฟล์ล่าสุดเข้าเมื่อ ' + esc(last.time);
  h += '<div class="foot"><span>' + left + '</span><span>ลากไฟล์ .zip ลงขั้นที่ 1 ของ PHOTO CHECK STUDIO ได้เลย ไม่ต้องแตกไฟล์</span></div>';
  h += '</main>';
  app.innerHTML = h;
}
function drawBar() { if (sendView()) return drawSend(); var a = $('#acts'); if (a) a.innerHTML = actsHtml(); }
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
function zip(cb) {
  var Pk = picked(), out = [], next = 0, done = 0, failed = false;
  if (!Pk.length || BUSY) return;
  BUSY = 'zip'; PROG = 0; draw();
  function fail(e) { if (failed) return; failed = true; BUSY = ''; if (typeof cb === 'function') SENDMSG = 'รวมไฟล์ไม่สำเร็จ: ' + ((e && e.message) || e) + ' — กด “ดาวน์โหลด” อีกครั้ง หรือใช้ “เปิดใน Drive”'; draw(); toast('ดาวน์โหลดไม่สำเร็จ: ' + ((e && e.message) || e) + ' — ลองอีกครั้ง หรือใช้ปุ่ม “เปิดใน Drive”', 1); }
  function one(i, part, acc) {
    google.script.run.withFailureHandler(fail).withSuccessHandler(function (r) {
      if (failed) return;
      acc.push(b64u8(r.b64));
      if (part + 1 < r.n) return one(i, part + 1, acc);
      out[i] = { name: Pk[i].name, data: join(acc) }; done++; PROG = done; drawBar();
      if (next < Pk.length) one(next++, 0, []);
      else if (done === Pk.length) finish();
    }).webPart(D.day, Pk[i].id, part);
  }
  function finish() {
    var used = {};
    out.forEach(function (e) {
      var n = e.name, d = n.lastIndexOf('.'), b = d > 0 ? n.slice(0, d) : n, x = d > 0 ? n.slice(d) : '', k = 2;
      while (used[n.toLowerCase()]) n = b + ' (' + k++ + ')' + x;
      used[n.toLowerCase()] = 1; e.name = n;
    });
    var name = 'Playlist_' + D.day + '.zip';
    if (typeof cb === 'function') { BUSY = ''; draw(); return cb(name, zipBlob(out), out.length); } // ส่งเข้าระบบแทนดาวน์โหลด
    var url = URL.createObjectURL(zipBlob(out));
    var a = document.createElement('a'); a.href = url; a.download = name; a.target = '_self';
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 60000);
    BUSY = ''; draw();
    toast('ดาวน์โหลด ' + name + ' (' + out.length + ' ไฟล์) แล้ว — ลากไฟล์นี้ลงขั้นที่ 1 ของ PHOTO CHECK STUDIO ได้เลย ไม่ต้องแตกไฟล์');
  }
  for (var k = 0; k < Math.min(3, Pk.length); k++) one(next++, 0, []);
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
  else if (b.dataset.a === 'sel') {
    var g = byKind(), L = b.dataset.g === 'oth' ? g.oth : extraFiles(g), all = L.every(function (f) { return SEL[f.id]; });
    L.forEach(function (f) { SEL[f.id] = !all; }); draw();
  }
});
document.addEventListener('change', function (e) {
  var t = e.target;
  if (t.id === 'pick') { if (t.value && t.value <= D.today) load(t.value); else t.value = D.day; return; }
  if (t.dataset.id) { SEL[t.dataset.id] = t.checked; drawBar(); }
});
load('');
</script></body></html>`;
}
