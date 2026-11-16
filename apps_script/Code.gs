/**
 * ดึง Playlist ประจำวันจาก Gmail ลง Google Drive อัตโนมัติ (Google Apps Script)
 * ใช้กับ PHOTO CHECK STUDIO — สคริปต์นี้แค่รวบรวมไฟล์ การเติมใบเช็คยังทำในระบบเหมือนเดิม
 *
 * ทำงานในบัญชี Google ของคุณเอง (ไม่ผ่านบริการภายนอก ไม่มีค่าใช้จ่าย)
 *   - ทุกชั่วโมง 7:00–18:00 (เวลาไทย) ค้นอีเมลที่มีไฟล์ Excel แนบ
 *   - เก็บ Playlist ของวันนี้ลงโฟลเดอร์ Google Drive "Playlist_ปปปป-ดด-วว"
 *   - Playlist ที่ส่งมาก่อนวันฉาย (เช่น รอบเสาร์–อาทิตย์ส่งวันศุกร์) เก็บเข้าโฟลเดอร์ของวันนั้นทันที
 *   - ฉบับแก้ไข (แบบเดียวกัน วันเดียวกัน ส่งมาใหม่ในอีเมลฉบับหลัง) ใช้ฉบับใหม่ ฉบับเก่าย้ายไปถังขยะของ Drive (กู้คืนได้ 30 วัน)
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
  // Playlist ที่ส่งมาก่อนวันฉาย (วันที่ในชื่อไฟล์เป็นวันข้างหน้า เช่น รอบเสาร์–อาทิตย์ที่ส่งวันศุกร์):
  //   เก็บเข้าโฟลเดอร์ของวันนั้นทันทีที่อีเมลมา ไม่ต้องรอถึงวัน — รับล่วงหน้าได้ไม่เกินกี่วัน (0 = ไม่เก็บล่วงหน้า)
  aheadDays: 4,
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
  // กลุ่มในหน้า PLAYLIST INBOX เรียงตามนี้ (กลุ่มที่ไม่อยู่ในรายการต่อท้าย)
  groups: ['แยกจอ', 'BKJ', 'สื่ออื่น', 'NTW', 'ห้าง / อาคาร', 'MRT และอื่น ๆ'],
  // Playlist ที่ระบบเติมใบเช็ครูปใช้ (ลำดับมีผล: ไฟล์เข้าแบบแรกที่ชื่อตรง)
  //   pattern = คำในชื่อไฟล์ • group = กลุ่มในหน้า PLAYLIST INBOX • latest = ถ้าวันนี้ไม่มีฉบับใหม่ ใช้ฉบับล่าสุด
  //   daily: false = ไม่ได้ส่งทุกวัน — ไม่นับในตัวเลข “ต้องมาทุกวัน” และไม่ขึ้นว่าขาด (ถ้าบริษัทเปลี่ยนรอบการส่ง แก้ตรงนี้ได้เลย)
  //   source: 'line' = ไม่ได้มาทางอีเมล (เตือนในอีเมลสรุป)
  //   unused = ระบบเติมใบเช็คไม่ได้ใช้ (เก็บไว้เฉย ๆ ไม่ติ๊กไว้ให้)
  //   tip = ข้อความใต้ชื่อในหน้า PLAYLIST INBOX ตอนที่วันนี้ยังไม่มา (ไม่ใส่ = "ยังไม่มีอีเมลวันนี้" • แบบที่ไม่ได้ส่งทุกวัน ระบบบอกรอบการส่งจากวันที่เคยส่งมาให้เอง)
  kinds: [
    { name: 'แยกจอ 40Blue', pattern: 'แยกจอ.*40\\s*Blue', group: 'แยกจอ' },
    { name: 'แยกจอ 40Red', pattern: 'แยกจอ.*40\\s*Red', group: 'แยกจอ' },
    { name: 'แยกจอ 80Blue', pattern: 'แยกจอ.*80\\s*Blue', group: 'แยกจอ' },
    { name: 'แยกจอ 80Red', pattern: 'แยกจอ.*80\\s*Red', group: 'แยกจอ' },
    { name: 'แยกจอ CITY', pattern: 'แยกจอ.*CITY', group: 'แยกจอ' },
    { name: 'แยกจอ Spare', pattern: 'แยกจอ.*Spare', group: 'แยกจอ' },
    { name: 'Show DC', pattern: 'Show[\\s_]*DC', latest: true, group: 'สื่ออื่น' },
    { name: 'Signature Max', pattern: 'Signature', group: 'สื่ออื่น' },
    { name: 'BKJ1', pattern: 'BKJ\\s*1', group: 'BKJ' },
    { name: 'BKJ2', pattern: 'BKJ\\s*2', group: 'BKJ' },
    // Metro LED มาทางอีเมลแล้ว (ตั้งแต่ พ.ย. 2569 จาก Media Scheduling Center มาก่อนวันออกอากาศ 1 วัน) — อีเมลแนบ 2 ไฟล์ Excel:
    //   "รายการหนังโฆษณา 7-Eleven Metro LED …" = ไฟล์ที่ระบบเติมใบเช็คใช้ • "Metro LED_… _Playlist" = ไฟล์ลำดับฉาย (ไม่ดึง — ดู skipPattern ด้านบน)
    { name: 'Metro LED', pattern: 'รายการหนัง.*Metro', tip: 'ปกติมาทางอีเมลก่อนวันฉาย 1 วัน', group: 'สื่ออื่น' },
    { name: 'Rama9 Connect', pattern: 'Rama\\s*9[\\s_]*connect', group: 'MRT และอื่น ๆ', daily: false },
    { name: 'Digital MRT', pattern: 'Digital[\\s_]*MRT', group: 'MRT และอื่น ๆ', daily: false },
    { name: 'Gateway X', pattern: 'Gateway[\\s_]*X', group: 'MRT และอื่น ๆ', daily: false },
    // ไม่รวม “Playlist So_Ploenchit_At Narong_Air Traveler Gateway_BIG Y …” = ไฟล์ลำดับฉายของจอ Signature Max (เก็บไว้ในไฟล์อื่น ไม่ติ๊กให้)
    { name: 'Air Traveller', pattern: 'Air[\\s_]*Travell?er(?![\\s_]*Gateway)', group: 'MRT และอื่น ๆ', daily: false },
    { name: 'The 20', pattern: 'The[\\s_]*20(?!\\d)', group: 'MRT และอื่น ๆ', daily: false },
    { name: 'LA PlanB', pattern: 'LA[\\s_]*Plan[\\s_]*B', latest: true, group: 'MRT และอื่น ๆ', daily: false },

    { name: 'NTW 57Blue_1', pattern: 'NTW[\\s_]*57[\\s_]*Blue[\\s_]*1(?!\\d)', group: 'NTW' },
    { name: 'NTW 57Blue_2', pattern: 'NTW[\\s_]*57[\\s_]*Blue[\\s_]*2(?!\\d)', group: 'NTW' },
    { name: 'NTW 57Red_1', pattern: 'NTW[\\s_]*57[\\s_]*Red[\\s_]*1(?!\\d)', group: 'NTW' },
    { name: 'NTW 57Red_2', pattern: 'NTW[\\s_]*57[\\s_]*Red[\\s_]*2(?!\\d)', group: 'NTW' },
    { name: 'NTW Full Only', pattern: 'NTW[\\s_]*Full', group: 'NTW' },
    { name: 'NTW Spare', pattern: 'NTW[\\s_]*Spare', group: 'NTW' },
    { name: 'Central Network', pattern: 'Central[\\s_]*Network', group: 'ห้าง / อาคาร' },
    { name: 'Parc Paragon', pattern: 'Parc[\\s_]*Paragon', group: 'ห้าง / อาคาร' },
    { name: 'Paragon Max', pattern: 'Paragon[\\s_]*Max', group: 'ห้าง / อาคาร', daily: false },
    { name: 'Siam Paragon in mall', pattern: 'Siam[\\s_]*Paragon[\\s_]*in[\\s_]*mall', group: 'ห้าง / อาคาร', daily: false },
    { name: 'Siam Paragon Facade', pattern: 'Siam[\\s_]*Paragon[\\s_]*Facade', group: 'ห้าง / อาคาร', daily: false },
    { name: 'The District EM', pattern: 'District[\\s_]*EM', group: 'ห้าง / อาคาร', daily: false },
    { name: 'The Panoramix', pattern: 'Panoramix', group: 'ห้าง / อาคาร', daily: false },
    { name: 'LIDO Connect', pattern: 'LIDO', group: 'ห้าง / อาคาร', daily: false },
  ],
};

const TZ = 'Asia/Bangkok';
const OTHER_ = 'Playlist อื่น';
const PROP_FOLDER = 'parentFolderId';
const PROP_WEB = 'webUrl';
const PROP_DL = 'downloadedDays'; // วันที่ดาวน์โหลด / ส่งเข้าระบบแล้ว (ใช้ทำกล่อง “ตามงานวันหยุด”)
const PROP_DLSINCE = 'downloadedSince';
const PROP_REV = 'revisions'; // ฉบับแก้ไขที่ส่งมาแทนฉบับเดิม { 'ปปปป-ดด-วว': [{ k: แบบ, n: ชื่อไฟล์ใหม่, t: เวลาอีเมลใหม่, o: [ชื่อฉบับเก่า], ot: เวลาฉบับเก่า, a: ดึงวันนั้นไปแล้วก่อนฉบับแก้ไขมา }] }
const PROP_SEEN = 'kindDays'; // วันที่แต่ละแบบเคยส่งมา (ล่าสุด 10 วัน) ใช้บอก “ล่าสุด …” และรอบการส่งของแบบที่ไม่ได้ส่งทุกวัน
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

/** ไฟล์ในโฟลเดอร์ (ไม่รวมไฟล์ในถังขยะ) */
function files_(folder) {
  return folder.searchFiles('trashed = false');
}

/** แบบของไฟล์ (ตามชื่อ) → ชื่อแบบใน CONFIG.kinds หรือ '' */
function kindOf_(name, kinds) {
  const k = (kinds || kindsRe_()).find(x => x.re.test(name));
  return k ? k.name : '';
}
function kindsRe_() {
  return CONFIG.kinds.map(k => Object.assign({}, k, { re: new RegExp(k.pattern, 'i') }));
}
const isDaily_ = k => k.daily !== false;

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
function pick_(atts, today, ahead) {
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
        if (!more[k2]) more[k2] = { kind: OTHER_, name: name, key: k2, day: day, when: a.when, msg: a.msg, blob: a.blob, size: a.size };
      } else if (!ahead) others[name] = 1;
      continue;
    }
    if (day !== today && !kind.latest) continue; // ไฟล์ของวันก่อน
    const key = name + '|' + a.size;
    const list = byKind[kind.name] || (byKind[kind.name] = {});
    if (!list[key]) list[key] = { kind: kind.name, name: name, key: key, day: day, when: a.when, msg: a.msg, blob: a.blob, size: a.size };
  }
  const chosen = [];
  for (const k of kinds) {
    const list = Object.values(byKind[k.name] || {});
    let todays = list.filter(f => f.day === today);
    // แบบเดียวกันของวันเดียวกันมาจากหลายอีเมล = ฉบับแก้ไข → ใช้เฉพาะไฟล์จากอีเมลฉบับล่าสุด (หลายไฟล์ในอีเมลเดียวกันเก็บไว้ทั้งหมด)
    if (todays.length > 1) {
      const last = todays.reduce((a, b) => (b.when > a.when ? b : a));
      todays = todays.filter(f => f.msg === last.msg || f.when === last.when);
    }
    if (todays.length) chosen.push(...todays);
    else if (k.latest && list.length && !ahead) {
      list.sort((a, b) => (a.day < b.day ? -1 : a.day > b.day ? 1 : a.when - b.when));
      chosen.push(list[list.length - 1]); // ฉบับล่าสุดที่มี
    }
  }
  chosen.push(...Object.values(more));
  return { chosen: chosen, others: Object.keys(others) };
}

/** ไฟล์แนบ Excel จาก Gmail (ย้อนหลัง 2 วัน + แบบที่ไม่ได้ส่งทุกวันย้อนหลัง latestDays วัน)
 *  ไฟล์ของวันข้างหน้าในอีเมลชุดนี้ เก็บล่วงหน้าไว้แล้ว (collectAhead_) จึงไม่ต้องค้นย้อนหลังนานกว่านี้ */
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
          out.push({ name: name, size: at.getSize(), day: Utilities.formatDate(d, TZ, 'yyyy-MM-dd'), when: d.getTime(), msg: msg.getId(), blob: at });
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
  const atts = attachments_();
  const P = pick_(atts, today);
  if (!PropertiesService.getScriptProperties().getProperty(PROP_SEEN)) seenSeed_(parent, today);
  const S = saveDay_(parent, today, P.chosen);
  let folder = S.folder;
  const added = S.added;
  const revs = S.revs;
  // แบบที่ไม่ได้ส่งทุกวัน (เช่น Show DC): ได้ฉบับของวันนี้แล้ว → ฉบับเก่าที่ใช้แทนไว้ก่อนหน้า ย้ายไปถังขยะของ Drive (กู้คืนได้)
  const trashed = [];
  if (folder && added.length) {
    const thisYear = +today.slice(0, 4);
    for (const k of CONFIG.kinds.filter(x => x.latest)) {
      const re = new RegExp(k.pattern, 'i');
      const L = [];
      const it = files_(folder);
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
  const A = collectAhead_(parent, atts, today);
  const ahead = A.added;
  const allRevs = revs.map(r => Object.assign({ day: today }, r)).concat(A.revs);
  const sum = summary_(folder, today, added, P.others, trashed, ahead, allRevs);
  if (added.length || ahead.length) MailApp.sendEmail({ to: CONFIG.notifyEmail || Session.getEffectiveUser().getEmail(), subject: sum.subject, htmlBody: sum.html, name: 'Playlist อัตโนมัติ' });
  const aheadTxt = ahead.length ? ' • เก็บล่วงหน้า ' + ahead.length + ' ไฟล์ (' + aheadDays_(ahead).join(', ') + ')' : '';
  const revTxt = allRevs.length ? ' • ฉบับแก้ไข ' + allRevs.map(r => r.k).join(', ') + ' (ใช้ฉบับใหม่แทนแล้ว)' : '';
  return { added: added.map(f => f.name), ahead: ahead.map(f => f.name), revs: allRevs.length, subject: sum.subject, text: sum.subject + (added.length ? ' — เพิ่ม ' + added.length + ' ไฟล์' : ' — ไม่มีไฟล์ใหม่ของวันนี้') + aheadTxt + revTxt + (added.length || ahead.length ? ' ส่งอีเมลสรุปแล้ว' : '') };
}

/** ข้อความเก็บไว้ในรายละเอียดไฟล์ใน Drive: อีเมลไหน เข้ากี่โมง (ใช้แยกฉบับแก้ไขจากอีเมลฉบับหลัง) */
function mailNote_(f) {
  return 'PLAYLIST INBOX · อีเมลเข้า ' + Utilities.formatDate(new Date(f.when), TZ, 'yyyy-MM-dd H:mm') + ' · id ' + (f.msg || '') + ' · t ' + f.when;
}
function mailOf_(file) {
  let d = '';
  try {
    d = file.getDescription() || '';
  } catch (e) {}
  const m = d.match(/id (\S*) · t (\d+)/);
  return m ? { msg: m[1], when: +m[2] } : null;
}

/** เก็บไฟล์ที่คัดแล้วเข้าโฟลเดอร์ของวันนั้น — ไฟล์เดิม (ชื่อ + ขนาดเดียวกัน) ไม่เก็บซ้ำ
 *  ฉบับแก้ไข: แบบเดียวกัน วันที่เดียวกัน แต่มาจากอีเมลฉบับหลัง → เก็บฉบับใหม่ ฉบับเก่าย้ายไปถังขยะของ Drive (กู้คืนได้ 30 วัน)
 *  ไฟล์แบบเดียวกันหลายไฟล์ในอีเมลฉบับเดียวกัน (หรือบอกไม่ได้ว่าฉบับไหนมาก่อน) เก็บไว้ทั้งหมด หน้าเว็บติ๊กไว้ให้ไฟล์ล่าสุด */
function saveDay_(parent, day, chosen) {
  let folder = todayFolder_(parent, day, false);
  const kinds = kindsRe_();
  const thisYear = +day.slice(0, 4);
  const have = {};
  const ex = [];
  if (folder) {
    const it = files_(folder);
    while (it.hasNext()) {
      const f = it.next();
      const n = f.getName().normalize('NFC');
      have[n + '|' + f.getSize()] = 1;
      ex.push({ f: f, n: n, k: kindOf_(n, kinds), pd: nameDate_(n, thisYear) || day });
    }
  }
  const added = [];
  const revs = [];
  const sorted = chosen.slice().sort((a, b) => a.when - b.when);
  for (const c of sorted) {
    if (have[c.key]) continue;
    if (!folder) folder = todayFolder_(parent, day, true);
    const old = [];
    if (c.kind !== OTHER_ && c.day === day)
      for (const x of ex) {
        if (x.gone || x.k !== c.kind || x.pd !== day) continue;
        const m = mailOf_(x.f);
        const t = m ? m.when : x.f.getDateCreated().getTime();
        // ฉบับเดิมมาจากอีเมลฉบับก่อน (ไฟล์ที่เก็บก่อนมีระบบนี้: เก็บเข้า Drive ก่อนอีเมลใหม่มาถึง)
        if (m ? m.msg !== c.msg && m.when < c.when : t < c.when) old.push({ x: x, t: t });
      }
    const file = folder.createFile(c.blob.copyBlob().setName(c.name));
    try {
      file.setDescription(mailNote_(c));
    } catch (e) {}
    have[c.key] = 1;
    ex.push({ f: file, n: c.name, k: c.kind, pd: c.day });
    added.push(c);
    if (old.length) {
      old.forEach(o => {
        o.x.f.setTrashed(true);
        o.x.gone = true;
      });
      revs.push({ k: c.kind, n: c.name, t: c.when, o: old.map(o => o.x.n), ot: Math.max.apply(null, old.map(o => o.t)) });
    }
  }
  if (revs.length) revAdd_(day, revs);
  if (added.length) seenAdd_(added);
  return { folder: folder, added: added, revs: revs };
}

/** จำฉบับแก้ไขไว้ให้หน้าเว็บ (เก็บ 21 วัน) — ถ้าวันนั้นดึงไปแล้ว ติดป้าย a = ต้องดึงใหม่ (วันนั้นกลับมาในกล่อง “ตามงาน”) */
function revGet_() {
  try {
    return JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP_REV) || '{}');
  } catch (e) {
    return {};
  }
}
function revSet_(m) {
  const keep = addDays_(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), -21);
  const o = {};
  Object.keys(m)
    .sort()
    .forEach(d => {
      if (d >= keep && m[d] && m[d].length) o[d] = m[d].slice(-20);
    });
  PropertiesService.getScriptProperties().setProperty(PROP_REV, JSON.stringify(o));
}
function revAdd_(day, revs) {
  const m = revGet_();
  const done = !!dlGet_()[day];
  m[day] = (m[day] || []).filter(r => !revs.some(x => x.k === r.k)).concat(revs.map(r => Object.assign({}, r, { a: done })));
  revSet_(m);
}

/** วันที่แต่ละแบบเคยส่งมา (Playlist ของวันไหน) — 10 วันล่าสุด */
function seenGet_() {
  try {
    return JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP_SEEN) || '{}');
  } catch (e) {
    return {};
  }
}
function seenAdd_(files) {
  const m = seenGet_();
  let ch = false;
  files.forEach(f => {
    if (!f.kind || f.kind === OTHER_ || !f.day) return;
    const L = m[f.kind] || [];
    if (L.indexOf(f.day) >= 0) return;
    L.push(f.day);
    L.sort();
    m[f.kind] = L.slice(-10);
    ch = true;
  });
  if (ch) PropertiesService.getScriptProperties().setProperty(PROP_SEEN, JSON.stringify(m));
}
/** ครั้งแรก: ดูโฟลเดอร์ 21 วันย้อนหลังว่าแต่ละแบบเคยส่งมาวันไหน */
function seenSeed_(parent, today) {
  const m = {};
  const kinds = kindsRe_();
  for (let i = 21; i >= 0; i--) {
    const d = addDays_(today, -i);
    const f = todayFolder_(parent, d, false);
    if (!f) continue;
    const it = files_(f);
    while (it.hasNext()) {
      const n = it.next().getName().normalize('NFC');
      const k = kindOf_(n, kinds);
      const pd = nameDate_(n, +d.slice(0, 4)) || d;
      if (!k || pd !== d) continue;
      const L = m[k] || (m[k] = []);
      if (L.indexOf(d) < 0) L.push(d);
    }
  }
  Object.keys(m).forEach(k => (m[k] = m[k].sort().slice(-10)));
  PropertiesService.getScriptProperties().setProperty(PROP_SEEN, JSON.stringify(m));
}

/** วันที่ (ปปปป-ดด-วว) บวก n วัน */
function addDays_(day, n) {
  const d = new Date(Date.UTC(+day.slice(0, 4), +day.slice(5, 7) - 1, +day.slice(8, 10) + n));
  return d.toISOString().slice(0, 10);
}

/** วันในรายการไฟล์ล่วงหน้า เป็นข้อความ เช่น "ส. 21 พ.ย. 6 ไฟล์" */
function aheadDays_(ahead) {
  const TH = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  const WD = ['อา.', 'จ.', 'อ.', 'พ.', 'พฤ.', 'ศ.', 'ส.'];
  const n = {};
  for (const f of ahead) n[f.day] = (n[f.day] || 0) + 1;
  return Object.keys(n)
    .sort()
    .map(d => WD[new Date(Date.UTC(+d.slice(0, 4), +d.slice(5, 7) - 1, +d.slice(8, 10))).getUTCDay()] + ' ' + +d.slice(8, 10) + ' ' + TH[+d.slice(5, 7) - 1] + ' ' + n[d] + ' ไฟล์');
}

/** Playlist ที่ส่งมาก่อนวันฉาย (วันที่ในชื่อไฟล์เป็นวันข้างหน้า ไม่เกิน CONFIG.aheadDays วัน) → เก็บเข้าโฟลเดอร์ของวันนั้นเลย
 *  ถึงวันนั้นไม่ต้องหาอีเมลเก่าอีก (เดิมค้นอีเมลย้อนหลังแค่ 2 วัน รอบวันอาทิตย์ที่ส่งวันศุกร์จึงอาจหลุด)
 *  เฉพาะไฟล์ที่ระบุวันนั้นจริง — แบบ latest ไม่เอาฉบับเก่ามาใส่แทน และไม่ย้ายไฟล์ใดไปถังขยะ */
function collectAhead_(parent, atts, today) {
  const out = [];
  const revs = [];
  for (let i = 1; i <= (CONFIG.aheadDays | 0); i++) {
    const day = addDays_(today, i);
    const P = pick_(atts, day, true);
    if (!P.chosen.length) continue;
    const S = saveDay_(parent, day, P.chosen);
    S.added.forEach(f => out.push(Object.assign({}, f, { day: day })));
    S.revs.forEach(r => revs.push(Object.assign({ day: day }, r)));
  }
  return { added: out, revs: revs };
}

function summary_(folder, today, added, others, trashed, ahead, revs) {
  const TH = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  const thDay = s => +s.slice(8, 10) + ' ' + TH[+s.slice(5, 7) - 1];
  const day = thDay(today) + ' ' + (+today.slice(0, 4) + 543);
  const kinds = kindsRe_();
  const skip = skipRe_();
  const counts = {};
  const files = [];
  const more = [];
  if (folder) {
    const it = files_(folder);
    while (it.hasNext()) {
      const n = it.next().getName().normalize('NFC');
      if (skip && skip.test(n)) continue; // ไฟล์ที่ดึงไว้ก่อนตั้ง skipPattern
      files.push(n);
      const k = kinds.find(x => x.re.test(n));
      if (k) counts[k.name] = (counts[k.name] || 0) + 1;
      else more.push(n);
    }
  }
  const mail = kinds.filter(k => k.source !== 'line' && isDaily_(k)).map(k => k.name);
  const line = kinds.filter(k => k.source === 'line' && !counts[k.name]).map(k => k.name);
  const extra = kinds.filter(k => !isDaily_(k) && counts[k.name]).map(k => k.name);
  const haveN = mail.filter(k => counts[k]).length;
  const missing = mail.filter(k => !counts[k]);
  const multi = mail.filter(k => counts[k] > 1);
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
  const html =
    '<div style="font-family:Tahoma,sans-serif;font-size:15px;line-height:1.6">' +
    '<p>Playlist วันที่ ' + esc(day) + ' — แบบที่ต้องมาทุกวัน ได้แล้ว <b>' + haveN + '/' + mail.length + '</b> แบบ</p>' +
    (revs && revs.length
      ? '<p style="color:#1F5FE0"><b>มีฉบับแก้ไขส่งมาใหม่:</b> ' +
        revs.map(r => esc(r.k) + (r.day !== today ? ' (Playlist ' + esc(thDay(r.day)) + ')' : '')).join(', ') +
        ' — ระบบใช้ฉบับใหม่แทนให้แล้ว ฉบับเก่าอยู่ในถังขยะของ Drive (กู้คืนได้ 30 วัน)' +
        (revs.some(r => dlGet_()[r.day]) ? ' • <b>วันนั้นดึงเข้าระบบไปแล้ว ให้ดึงใหม่อีกครั้ง</b>' : '') +
        '</p>'
      : '') +
    (added.length
      ? '<p>ไฟล์ที่เพิ่งเข้าโฟลเดอร์รอบนี้:</p><ul>' +
        added
          .map(f => '<li>' + esc(f.kind) + ': ' + esc(f.name) + (f.day !== today ? ' <span style="color:#b35a14">(ยังไม่มีฉบับใหม่ — ใช้ฉบับล่าสุด รอบ ' + esc(thDay(f.day)) + ')</span>' : '') + '</li>')
          .join('') +
        '</ul>'
      : '') +
    (missing.length
      ? '<p>ยังขาด <b>' + missing.length + '</b> แบบ: ' + missing.map(esc).join(', ') + '</p><p>ระบบเช็กอีเมลทุกชั่วโมงจนถึง ' + CONFIG.endHour + ':00</p>'
      : '<p><b>ครบ ' + mail.length + ' แบบที่ต้องมาทุกวันแล้ว</b> ดาวน์โหลดทั้งโฟลเดอร์ แล้วลากลง PHOTO CHECK STUDIO ได้เลย</p>') +
    (line.length ? '<p style="color:#1d4f91"><b>มาทาง LINE:</b> ' + line.map(esc).join(', ') + ' — ดาวน์โหลดจาก LINE แล้วใส่ในระบบเอง</p>' : '') +
    (extra.length ? '<p>แบบที่ไม่ได้ส่งทุกวัน วันนี้มา (ไม่นับใน ' + mail.length + ' แบบ): ' + extra.map(esc).join(', ') + '</p>' : '') +
    (more.length ? '<p>Playlist อื่น ๆ ที่ไม่อยู่ในรายการ (เก็บไว้ให้ด้วย ' + more.length + ' ไฟล์): ' + more.map(esc).join(', ') + '</p>' : '') +
    (multi.length ? '<p style="color:#b35a14"><b>บางแบบมีมากกว่า 1 ไฟล์:</b> ' + multi.map(k => esc(k) + ' (' + counts[k] + ' ไฟล์)').join(', ') + ' — ตรวจว่าต้องใช้ฉบับไหนก่อนลากลงระบบ</p>' : '') +
    (ahead && ahead.length
      ? '<p style="color:#1d4f91"><b>Playlist ล่วงหน้า (ส่งมาก่อนวันฉาย)</b> — เก็บเข้าโฟลเดอร์ของวันนั้นแล้ว ถึงวันเปิด PLAYLIST INBOX ได้ตามปกติ: ' + aheadDays_(ahead).map(esc).join(' • ') + '</p>'
      : '') +
    (trashed && trashed.length ? '<p style="color:#586274">ได้ฉบับของวันนี้แล้ว ย้ายฉบับเก่าไปถังขยะของ Drive: ' + trashed.map(esc).join(', ') + '</p>' : '') +
    (others.length ? '<p style="color:#586274">ไฟล์ Excel อื่นของวันนี้ (ไม่ได้เก็บ เพราะชื่อไม่เหมือน Playlist):</p><ul style="color:#586274">' + others.map(n => '<li>' + esc(n) + '</li>').join('') + '</ul>' : '') +
    (webUrl_() ? '<p><a href="' + webUrl_() + '"><b>เปิด PLAYLIST INBOX</b></a> — ดูว่ามาแล้วแบบไหน และดาวน์โหลดทั้งหมดเป็นไฟล์ .zip ไฟล์เดียว</p>' : '') +
    (folder ? '<p><a href="' + folder.getUrl() + '">เปิดโฟลเดอร์ Google Drive ของวันนี้</a></p>' : '') +
    '</div>';
  return { subject: 'Playlist ' + day + ': ' + haveN + '/' + mail.length + ' แบบ' + (missing.length ? ' (ยังไม่ครบ)' : ' — ครบแล้ว') + (revs && revs.length ? ' • มีฉบับแก้ไข' : ''), html: html };
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

/** ไฟล์ในโฟลเดอร์ของวันหนึ่ง เรียงตามเวลาที่เข้า Drive */
function dayList_(folder, day) {
  const TH = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  const thDay = s => +s.slice(8, 10) + ' ' + TH[+s.slice(5, 7) - 1];
  const kinds = kindsRe_();
  const skip = skipRe_();
  const thisYear = +day.slice(0, 4);
  const files = [];
  if (folder) {
    const it = files_(folder);
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
        ad: atDay, // วันที่เข้าโฟลเดอร์
        hm: Utilities.formatDate(at, TZ, 'H:mm'),
        time: (atDay !== day ? thDay(atDay) + ' ' : '') + Utilities.formatDate(at, TZ, 'H:mm'),
        old: fd && fd < day ? thDay(fd) : '',
        pd: fd || '', // วันที่ของ Playlist ตามชื่อไฟล์ (ปปปป-ดด-วว)
      });
    }
  }
  files.sort((a, b) => a.at - b.at);
  return files;
}

/** สรุปวันหนึ่งสำหรับกล่อง “Playlist ล่วงหน้า” / “ตามงานวันหยุด”: มาแล้วกี่แบบ ขาดแบบไหน และไฟล์ที่ติ๊กไว้ให้
 *  (แบบละ 1 ไฟล์ = ฉบับล่าสุด เหมือนค่าเริ่มของหน้าเว็บ) • วันข้างหน้า: แบบ latest นับว่ามี เพราะถึงวันใช้ฉบับล่าสุดให้เอง */
function daySum_(day, files, today, revs) {
  const main = CONFIG.kinds.filter(k => isDaily_(k) && k.source !== 'line');
  const by = {};
  files.forEach(f => {
    if (f.kind) (by[f.kind] = by[f.kind] || []).push(f);
  });
  const miss = main.filter(k => !by[k.name] && !(k.latest && day > today)).map(k => k.name);
  const pick = [];
  let extra = 0;
  CONFIG.kinds.forEach(k => {
    const L = by[k.name];
    if (!L || k.unused) return;
    pick.push(L[L.length - 1]);
    if (!isDaily_(k)) extra++;
  });
  let size = 0;
  pick.forEach(f => (size += f.size));
  const R = revs || [];
  return {
    day: day,
    have: main.length - miss.length,
    total: main.length,
    miss: miss,
    extra: extra,
    n: pick.length,
    size: size,
    files: pick.map(f => ({ id: f.id, name: f.name, size: f.size })),
    rev: R.map(r => r.k),
    redo: R.filter(r => r.a).map(r => r.k), // ดึงวันนั้นไปแล้วก่อนฉบับแก้ไขเหล่านี้มา
  };
}

/** วันที่ดาวน์โหลด / ส่งเข้าระบบแล้ว { 'ปปปป-ดด-วว': 'd' = ดึงแล้ว, 's' = กดไม่ต้อง } */
function dlGet_() {
  try {
    return JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP_DL) || '{}');
  } catch (e) {
    return {};
  }
}

/** วันแรกที่เริ่มจำการดาวน์โหลด (วันที่ติดตั้งโค้ดนี้) — วันก่อนหน้านั้นถือว่าดึงไปแล้ว ไม่ขึ้นในกล่อง “ตามงานวันหยุด” */
function dlSince_(today) {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty(PROP_DLSINCE);
  if (!s) {
    s = today;
    p.setProperty(PROP_DLSINCE, s);
  }
  return s;
}

/** หน้าเว็บบอกว่าดึงวันไหนแล้ว — ดาวน์โหลด .zip / ส่งเข้าระบบ = 'd' • กด “ไม่ต้อง” = 's' */
function webMark(days, how) {
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const m = dlGet_();
  (days || []).forEach(d => {
    if (/^\d{4}-\d{2}-\d{2}$/.test(d || '') && d <= today) m[d] = how === 's' ? 's' : 'd';
  });
  const o = {};
  Object.keys(m)
    .sort()
    .slice(-60)
    .forEach(k => (o[k] = m[k]));
  PropertiesService.getScriptProperties().setProperty(PROP_DL, JSON.stringify(o));
  // ดึงวันนั้นแล้ว (หรือกดไม่ต้อง) → ฉบับแก้ไขของวันนั้นไม่ต้องเตือนให้ดึงใหม่อีก
  const R = revGet_();
  let ch = false;
  (days || []).forEach(d => (R[d] || []).forEach(r => {
    if (r.a) {
      r.a = false;
      ch = true;
    }
  }));
  if (ch) revSet_(R);
  return true;
}

/** วันก่อนหน้าที่ยังไม่ได้ดึง (วันหยุด / ลา / วันที่ไม่ได้เปิดระบบ) ย้อนจากเมื่อวาน จนถึงวันที่ดึงแล้ว (ไม่เกิน 10 วัน)
 *  วันที่ไม่มีโฟลเดอร์ (ไม่มี Playlist ส่งมา) ข้ามไปดูวันก่อนหน้าต่อ */
function catchup_(parent, today) {
  const dl = dlGet_();
  const since = dlSince_(today);
  const R = revGet_();
  const days = [];
  for (let i = 1; i <= 10; i++) {
    const d = addDays_(today, -i);
    if (d < since || dl[d]) break;
    days.push(d);
  }
  // วันที่ดึงไปแล้ว แต่มีฉบับแก้ไขส่งมาทีหลัง → กลับมาในกล่องนี้ให้ดึงใหม่
  Object.keys(R).forEach(d => {
    if (d < today && d >= addDays_(today, -10) && days.indexOf(d) < 0 && R[d].some(r => r.a)) days.push(d);
  });
  days.sort();
  const out = [];
  days.forEach(d => {
    const f = todayFolder_(parent, d, false);
    if (!f) return;
    const L = dayList_(f, d);
    if (L.length) out.push(daySum_(d, L, today, R[d]));
  });
  return out;
}

/** ข้อมูลของวันหนึ่ง (ปปปป-ดด-วว) ให้หน้าเว็บ — ดูวันข้างหน้าได้ไม่เกิน aheadDays วัน (Playlist ที่ส่งมาก่อนวันฉาย) */
function webDay(day) {
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const maxDay = addDays_(today, CONFIG.aheadDays | 0);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(day || '') || day > maxDay) day = today;
  const parent = parentFolder_(false);
  const folder = todayFolder_(parent, day, false);
  const files = dayList_(folder, day);
  const R = revGet_();
  const fmt = t => {
    const d = new Date(t);
    const dd = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
    const TH = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
    return (dd !== today ? +dd.slice(8, 10) + ' ' + TH[+dd.slice(5, 7) - 1] + ' ' : '') + Utilities.formatDate(d, TZ, 'H:mm');
  };
  const revs = (R[day] || []).map(r => ({ kind: r.k, name: r.n, time: fmt(r.t), oldTime: r.ot ? fmt(r.ot) : '', old: r.o || [], after: !!r.a }));
  const seen = seenGet_();
  let auto = null;
  try {
    auto = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'collectPlaylists');
  } catch (e) {}
  const ahead = [];
  let catchup = null;
  if (day === today) {
    for (let i = 1; i <= (CONFIG.aheadDays | 0); i++) {
      const d = addDays_(today, i);
      const f = todayFolder_(parent, d, false);
      const L = f ? dayList_(f, d) : [];
      if (L.length) ahead.push(daySum_(d, L, today, R[d]));
    }
    const prev = catchup_(parent, today);
    if (prev.length) catchup = { days: prev.concat([daySum_(today, files, today, R[today])]) };
  }
  const groups = (CONFIG.groups || []).slice();
  CONFIG.kinds.forEach(k => {
    const g = k.group || 'อื่น ๆ';
    if (groups.indexOf(g) < 0) groups.push(g);
  });
  return {
    day: day,
    today: today,
    maxDay: maxDay,
    folderUrl: folder ? folder.getUrl() : '',
    parentUrl: parent.getUrl(),
    now: Utilities.formatDate(new Date(), TZ, 'H:mm'),
    groups: groups,
    kinds: CONFIG.kinds.map(k => ({ name: k.name, group: k.group || 'อื่น ๆ', daily: isDaily_(k), line: k.source === 'line', latest: !!k.latest, unused: !!k.unused, tip: k.tip || '', days: seen[k.name] || [] })),
    files: files,
    revs: revs,
    auto: auto,
    startHour: CONFIG.startHour,
    endHour: CONFIG.endHour,
    studio: STUDIO_URL,
    ahead: ahead,
    catchup: catchup,
  };
}

/** ส่งไฟล์หนึ่งไฟล์ในโฟลเดอร์ของวันนั้นให้หน้าเว็บทีละส่วน (base64) — หน้าเว็บรวมไฟล์ที่เลือกเป็น .zip เอง
 *  (Playlist ทั้งวันรวมกันราว 20 MB ถ้าส่งก้อนเดียวอาจเกินขนาดที่ google.script.run รับได้ จึงส่งทีละไม่เกิน 8 ล้านตัวอักษร) */
function webPart(day, id, part) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(day || '')) throw new Error('วันที่ไม่ถูกต้อง');
  const folder = todayFolder_(parentFolder_(false), day, false);
  if (!folder) throw new Error('ยังไม่มีโฟลเดอร์ Playlist ของวันนี้');
  const it = files_(folder); // เฉพาะไฟล์ในโฟลเดอร์ของวันนั้นเท่านั้น
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
:root{color-scheme:light;
 --paper:#EEF2F8;--surface:#FFFFFF;--rail:#E4EAF3;--slotIn:#F6F8FC;--ink:#0B1E46;--ink2:#36435F;--muted:#5B6782;--line:#D6DEEA;--line2:#E6EBF3;--box:#8C97A9;--dis:#B8C1CF;
 --blue:#1F5FE0;--blueInk:#FFFFFF;--blueSoft:#E7EEFD;--blueLine:#9DB8F2;--ok:#0E7A54;--okSoft:#E2F3EB;--wait:#8F5200;--waitSoft:#FFF4DC;--waitLine:#E3B65B;--err:#B3261E;
 --now:#EE7D22;--board:#0B1E46;--board2:#16306A;--bInk:#FFFFFF;--bMuted:#A9B6D3;--bLine:rgba(255,255,255,.16);--sheet:#1E8E5A;--pin:#34D399;
 --sh:0 1px 2px rgba(11,30,70,.05),0 10px 30px -14px rgba(11,30,70,.22);--slotSh:0 1px 0 rgba(11,30,70,.06),0 2px 6px -2px rgba(11,30,70,.12);
 --hd:"Prompt","IBM Plex Sans Thai","Leelawadee UI",Tahoma,sans-serif;--ui:"IBM Plex Sans Thai","Leelawadee UI",Tahoma,sans-serif}
@media (prefers-color-scheme:dark){:root{color-scheme:dark;
 --paper:#0A1224;--surface:#111C33;--rail:#0D1730;--slotIn:#0F1A31;--ink:#E7ECF7;--ink2:#C2CBDD;--muted:#97A3BC;--line:#26344F;--line2:#1C2944;--box:#6B7A94;--dis:#4A5873;
 --blue:#7CA4FF;--blueInk:#06122E;--blueSoft:#172A4F;--blueLine:#34528A;--ok:#4CD69A;--okSoft:#10302A;--wait:#F2C46B;--waitSoft:#2C2310;--waitLine:#7D6220;--err:#F28B82;
 --board:#0E1B3B;--board2:#1A3270;--sheet:#3BBF80;--sh:none;--slotSh:none}}
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{margin:0;background:var(--paper);color:var(--ink);font:15px/1.6 var(--ui)}
button,input{font:inherit;color:inherit}
a{color:inherit}
svg{flex:none}
.in{max-width:1180px;margin:0 auto;padding:0 28px}
.top{background:var(--surface);border-bottom:1px solid var(--line)}
.top .in{display:flex;align-items:center;justify-content:space-between;gap:12px 16px;min-height:72px;flex-wrap:wrap;padding-top:10px;padding-bottom:10px}
.brand{display:flex;align-items:center;gap:12px;min-width:0}
.brand svg{display:block}
.brand b{display:block;font:700 19px/1.2 var(--hd);letter-spacing:.04em}
.brand small{display:block;font-size:12px;color:var(--muted)}
.tools{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.date{display:flex;align-items:center;gap:2px;padding:3px;border:1px solid var(--line);border-radius:999px;background:var(--surface)}
.nav{width:38px;height:38px;border:0;border-radius:999px;background:transparent;display:grid;place-items:center;color:var(--ink2);cursor:pointer}
.nav:hover:not(:disabled){background:var(--paper)}
.nav:disabled{color:var(--dis);cursor:default}
.cur{position:relative;display:flex;align-items:center;gap:8px;height:38px;padding:0 14px;border-radius:999px;background:var(--paper);font-weight:600;font-size:14px;white-space:nowrap;cursor:pointer}
.cur input{position:absolute;inset:0;width:100%;height:100%;opacity:0;border:0;padding:0;margin:0;cursor:pointer}
.cur input::-webkit-calendar-picker-indicator{position:absolute;inset:0;width:100%;height:100%;margin:0;opacity:0;cursor:pointer}
.cur:focus-within{outline:3px solid var(--blue);outline-offset:2px}
.tag{padding:0 8px;border-radius:99px;background:var(--ink);color:var(--surface);font-size:12px;line-height:20px;font-weight:500}
.btn{display:inline-flex;align-items:center;justify-content:center;gap:8px;min-height:42px;padding:0 18px;border-radius:12px;border:1px solid var(--line);background:var(--surface);font-weight:600;font-size:14px;color:var(--ink);white-space:nowrap;text-decoration:none;cursor:pointer}
.btn:hover:not(:disabled){border-color:var(--box)}
.btn.dark{background:var(--ink);border-color:var(--ink);color:var(--surface)}
.btn.pri{background:var(--blue);border-color:var(--blue);color:var(--blueInk);min-height:48px;padding:0 22px;font-size:15px}
.btn:disabled{opacity:.5;cursor:default}
.btn:focus-visible,.nav:focus-visible,.lk:focus-visible,.go:focus-visible,.tchip:focus-visible,.ico:focus-visible,button.day:focus-visible{outline:3px solid var(--blue);outline-offset:2px}
.lk{border:0;background:none;padding:4px 2px;color:var(--ink);font-weight:600;font-size:13px;text-decoration:underline;text-underline-offset:3px;cursor:pointer}
main{padding:28px 0 128px;transition:opacity .15s}
main.dim{opacity:.55}
main .in{display:flex;flex-direction:column;gap:22px}
.banner{display:flex;gap:10px;align-items:flex-start;padding:12px 16px;border-radius:14px;background:var(--waitSoft);border:1.5px dashed var(--waitLine);color:var(--ink2);font-size:14px}
.banner.info{background:var(--blueSoft);border:1px solid var(--blueLine)}
/* ---- กระดานสรุป ---- */
.board{position:relative;background:var(--board);color:var(--bInk);border-radius:22px;padding:26px 30px 22px;overflow:hidden}
.board::after{content:"";position:absolute;inset:0;background:radial-gradient(800px 260px at 85% -40%,var(--board2),transparent 70%);pointer-events:none}
.board>*{position:relative;z-index:1}
.bhead{display:flex;justify-content:space-between;align-items:flex-start;gap:14px 18px;flex-wrap:wrap}
.kicker{font-size:13px;color:var(--bMuted)}
.count{display:flex;align-items:baseline;gap:6px 16px;margin-top:2px;flex-wrap:wrap}
.count .n{font:700 64px/1 var(--hd);letter-spacing:-.01em;white-space:nowrap}
.count .n span{color:var(--bMuted);font-weight:600}
.count .t{font:600 20px/1.35 var(--hd);min-width:0}
.count .t small{display:block;font:400 14px/1.55 var(--ui);color:var(--bMuted);margin-top:4px;max-width:72ch}
.count .t small b{color:#FFD08A;font-weight:600}
.auto{display:flex;align-items:center;gap:8px;font-size:13px;color:var(--bMuted);padding:7px 12px;border:1px solid var(--bLine);border-radius:99px;white-space:nowrap}
.auto i{width:8px;height:8px;border-radius:9px;background:var(--pin);box-shadow:0 0 0 4px rgba(52,211,153,.18)}
.auto.off i{background:#8C97A9;box-shadow:none}
.tl{margin-top:18px}
.track{position:relative;height:80px}
.hours{position:absolute;left:0;right:0;bottom:0;display:flex;justify-content:space-between;font:500 12px/1 var(--hd);color:var(--bMuted)}
.hours span{width:0;display:flex;justify-content:center}
.rail{position:absolute;left:0;right:0;bottom:24px;height:2px;background:var(--bLine)}
.rail .past{position:absolute;left:0;top:0;bottom:0;background:rgba(255,255,255,.55)}
.rail .fut{position:absolute;top:0;bottom:0;right:0;background:repeating-linear-gradient(90deg,rgba(255,255,255,.28) 0 6px,transparent 6px 12px)}
.tick{position:absolute;bottom:20px;width:2px;height:10px;margin-left:-1px;background:rgba(255,255,255,.22);border-radius:2px}
.drop{--lv:0;position:absolute;bottom:20px;width:0;z-index:2}
.drop .pin{position:absolute;left:-5px;bottom:0;width:10px;height:10px;border-radius:10px;background:var(--pin);border:2px solid var(--board)}
.drop.pre .pin{background:#9DB8F2}
.drop .stem{position:absolute;left:-.5px;bottom:8px;width:1px;height:calc(10px + var(--lv)*32px);background:rgba(255,255,255,.45)}
.drop .chip{position:absolute;left:0;transform:translateX(-50%);bottom:calc(18px + var(--lv)*32px);display:flex;align-items:center;gap:6px;padding:3px 9px 3px 7px;border-radius:8px;background:#fff;color:#0B1E46;font:600 12px/1.3 var(--ui);white-space:nowrap;box-shadow:0 6px 14px -6px rgba(0,0,0,.5)}
.drop.edge .chip{transform:none;left:-6px}
.drop.edgeR .chip{transform:none;left:auto;right:-6px}
.drop .chip em{font:600 11px/1 var(--hd);font-style:normal;color:#5C6B8F}
.nowm{position:absolute;bottom:14px;top:0;width:2px;background:var(--now);transform:translateX(-1px);z-index:1}
.nowm b{position:absolute;top:0;left:8px;font:600 12px/1 var(--hd);color:var(--now);white-space:nowrap}
.nowm.flip b{left:auto;right:8px}
/* ---- ฉบับแก้ไข ---- */
.revnote{display:flex;align-items:center;gap:12px 14px;flex-wrap:wrap;background:var(--blueSoft);border:1px solid var(--blue);border-radius:16px;padding:14px 18px}
.revnote .ri{width:36px;height:36px;border-radius:10px;background:var(--blue);color:var(--blueInk);display:grid;place-items:center;flex:none}
.revnote .rt{flex:1;min-width:240px;font-size:14px;line-height:1.55;color:var(--ink2)}
.revnote .rt b{font:600 16px/1.4 var(--hd);display:block;color:var(--ink)}
.revnote .btn{min-height:40px}
/* ---- กำลังจะมา / ตามงาน ---- */
.strip{background:var(--surface);border:1px solid var(--line);border-radius:18px;padding:16px 18px;display:grid;grid-template-columns:160px minmax(0,1fr);gap:12px 18px;align-items:start}
.strip h3{margin:2px 0 0;font:600 16px/1.3 var(--hd)}
.strip h3 small{display:block;font:400 12px/1.45 var(--ui);color:var(--muted);margin-top:4px}
.dayrow{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:10px}
.day{position:relative;display:flex;align-items:center;gap:12px;padding:10px 12px;border:1px solid var(--line2);border-radius:14px;background:var(--slotIn);color:inherit;text-align:left;min-width:0;font-size:inherit}
button.day{cursor:pointer}
button.day:hover{border-color:var(--box)}
label.day{cursor:pointer;padding-right:42px}
label.day:not(:has(input:checked)){opacity:.6}
.day input{position:absolute;opacity:0;width:1px;height:1px;margin:0}
.day .dt{display:flex;flex-direction:column;min-width:0}
.day .dn{font:600 15px/1.3 var(--hd)}
.day .dn .tag{margin-left:6px;vertical-align:2px}
.day .ds{font-size:12px;line-height:1.45;color:var(--muted)}
.day .ds.rv{color:var(--blue);font-weight:600}
.day .go{align-self:flex-start;padding:0;border:0;background:none;font-size:12px;font-weight:600;text-decoration:underline;text-underline-offset:3px;color:var(--ink2);cursor:pointer}
.ring{--p:0;width:42px;height:42px;border-radius:42px;background:conic-gradient(var(--ok) calc(var(--p)*1%),var(--line) 0);display:grid;place-items:center;flex:none}
.ring span{width:32px;height:32px;border-radius:32px;background:var(--slotIn);display:grid;place-items:center;font:700 11px/1 var(--hd);color:var(--ink2)}
.catch{background:var(--blueSoft);border-color:var(--blueLine)}
.catch .day{background:var(--surface)}
.catch .ring span{background:var(--surface)}
.catch .act{display:flex;align-items:center;justify-content:space-between;gap:10px 14px;flex-wrap:wrap;grid-column:2}
.catch .act .note{font-size:13px;color:var(--ink2)}
.box{position:absolute;width:20px;height:20px;border-radius:6px;border:1.5px solid var(--box);background:var(--surface);display:grid;place-items:center}
.box svg{opacity:0}
input:checked~.box{background:var(--ink);border-color:var(--ink)}
input:checked~.box svg{opacity:1}
input:focus-visible~.box{outline:3px solid var(--blue);outline-offset:2px}
.day .box{right:12px;top:50%;margin-top:-10px}
/* ---- ช่องรับ Playlist ---- */
.rack{background:var(--surface);border:1px solid var(--line);border-radius:22px;box-shadow:var(--sh);padding:6px 22px}
.row{display:grid;grid-template-columns:160px minmax(0,1fr);gap:12px 18px;padding:18px 0}
.row+.row{border-top:1px solid var(--line2)}
.row h2{margin:0;padding-top:8px;font:600 16px/1.3 var(--hd)}
.row h2 small{display:block;font:400 12px/1.5 var(--ui);color:var(--muted)}
.row h2 .gcs{display:flex;flex-wrap:wrap;gap:6px;margin-top:8px}
.gc{display:inline-flex;align-items:center;gap:4px;padding:1px 8px;border-radius:99px;font:600 12px/1.6 var(--hd);background:var(--okSoft);color:var(--ok);white-space:nowrap}
.gc.part{background:var(--waitSoft);color:var(--wait)}
.gc.nd{background:var(--line2);color:var(--muted)}
.slots{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:10px;padding:10px;border-radius:16px;background:var(--rail)}
.slot{position:relative;display:flex;flex-direction:column;gap:6px;min-height:118px;padding:12px;border-radius:12px;min-width:0}
.slot.w2{grid-column:span 2}
.slot .nm{font:600 16px/1.25 var(--hd);padding-right:26px;overflow-wrap:anywhere}
.slot input{position:absolute;opacity:0;width:1px;height:1px;margin:0}
label.slot{background:var(--surface);box-shadow:var(--slotSh);cursor:pointer;border:1px solid transparent}
label.slot:hover{border-color:var(--line)}
label.slot:not(:has(input:checked)){background:var(--slotIn)}
label.slot:not(:has(input:checked)) .sheet{opacity:.45}
.slot .sheet{position:absolute;right:12px;top:12px;display:block}
.slot .box{right:12px;bottom:12px}
.srow{display:flex;flex-wrap:wrap;align-items:center;gap:6px}
.stamp{display:inline-flex;align-items:center;gap:5px;padding:1px 8px;border:1.5px solid currentColor;border-radius:99px;font:600 12px/1.5 var(--hd);color:var(--ok);white-space:nowrap}
.stamp.info{color:var(--blue)}
.stamp.grey{color:var(--muted)}
.slot .mm{font-size:12px;line-height:1.45;color:var(--ink2);padding-right:26px}
.slot .fn{padding-right:26px;font-size:12px;line-height:1.4;color:var(--muted);overflow-wrap:anywhere;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;margin-top:auto}
.slot .hint{font-size:12px;line-height:1.45}
.slot.wait{background:var(--slotIn);border:1.5px dashed var(--waitLine)}
.slot.wait .stamp,.slot.wait .hint{color:var(--wait)}
.slot.nd{background:var(--slotIn);border:1px solid var(--line2)}
.slot.nd .nm{color:var(--ink2)}
.slot.nd .fn{padding-right:0}
.slot.line{background:var(--blueSoft);border:1.5px dashed var(--blueLine)}
.slot.line .stamp{color:var(--blue)}
.slot.line .hint{color:var(--ink2)}
label.slot.rev{border:2px solid var(--blue)}
.tagnd,.tagrev{font:600 11px/1.7 var(--hd);padding:0 8px;border-radius:99px;white-space:nowrap}
.tagnd{color:var(--muted);background:var(--line2)}
.tagrev{color:var(--blueInk);background:var(--blue)}
.legend{display:flex;flex-wrap:wrap;gap:6px 16px;font-size:12px;color:var(--muted);padding:12px 0 8px;border-top:1px solid var(--line2)}
.legend span{display:inline-flex;align-items:center;gap:6px}
.legend i{width:14px;height:14px;border-radius:4px;display:inline-block}
/* ---- ไฟล์อื่น ---- */
.more{background:var(--surface);border:1px solid var(--line);border-radius:18px;padding:16px 18px;display:grid;grid-template-columns:160px minmax(0,1fr);gap:12px 18px}
.more h3{margin:2px 0 0;font:600 16px/1.3 var(--hd)}
.more h3 small{display:block;font:400 12px/1.45 var(--ui);color:var(--muted);margin-top:4px}
.more .sub{display:flex;flex-wrap:wrap;align-items:baseline;justify-content:space-between;gap:4px 12px;font-size:13px;color:var(--muted);font-weight:600;margin:0 0 8px}
.more .sub+.minis,.minis+.sub{margin-top:0}
.minis+.sub{margin-top:14px}
.minis{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:8px}
.mini{display:flex;align-items:flex-start;gap:10px;padding:10px 12px;border:1px solid var(--line2);border-radius:12px;background:var(--slotIn);cursor:pointer;min-width:0}
.mini input{width:16px;height:16px;margin:3px 0 0;accent-color:var(--ink);flex:none;cursor:pointer}
.mini .mt{display:flex;flex-direction:column;min-width:0}
.mini b{font-weight:600;font-size:14px;line-height:1.4}
.mini .fn,.mini .mm{font-size:12px;line-height:1.4;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.foot{display:flex;flex-wrap:wrap;justify-content:space-between;gap:8px 24px;font-size:13px;color:var(--muted)}
/* ---- ถาดรับ (ด้านล่าง) ---- */
.dock{position:fixed;left:0;right:0;bottom:0;z-index:5;background:var(--surface);background:color-mix(in srgb,var(--surface) 94%,transparent);-webkit-backdrop-filter:blur(10px);backdrop-filter:blur(10px);border-top:1px solid var(--line);box-shadow:0 -10px 30px -18px rgba(11,30,70,.35)}
.dock>.in{display:flex;align-items:center;justify-content:space-between;gap:10px 16px;min-height:78px;padding-top:10px;padding-bottom:10px}
.tray{display:flex;align-items:center;gap:12px;min-width:0;flex:1}
.ico{position:relative;width:46px;height:46px;border-radius:13px;border:0;background:var(--blueSoft);color:var(--blue);display:grid;place-items:center;flex:none;cursor:pointer}
.ico em{position:absolute;top:-6px;right:-6px;min-width:22px;height:22px;padding:0 6px;border-radius:99px;background:var(--blue);color:var(--blueInk);font:700 12px/22px var(--hd);font-style:normal;text-align:center;border:2px solid var(--surface)}
.ico.bump{animation:bump .45s ease}
.tt{min-width:0;flex:1}
.tt b{display:block;font:600 16px/1.3 var(--hd)}
.tchips{display:flex;gap:6px;margin-top:5px;overflow:hidden;white-space:nowrap;min-height:26px;align-items:center}
.tchip{display:inline-flex;align-items:center;gap:4px;flex:none;height:26px;padding:0 6px 0 10px;border:1px solid var(--line);border-radius:99px;background:var(--slotIn);font:600 12px/1 var(--hd);color:var(--ink2);cursor:pointer}
.tchip svg{opacity:.55}
.tchip:hover{border-color:var(--box)}
.tchip:hover svg{opacity:1}
.tchip[hidden],.tmore[hidden]{display:none}
.tchip.new{animation:drop .4s cubic-bezier(.2,.8,.3,1.25)}
.tmore{flex:none;height:26px;padding:0 10px;border:0;border-radius:99px;background:var(--blueSoft);color:var(--blue);font:700 12px/1 var(--hd);cursor:pointer}
.tempty{font-size:12px;color:var(--muted)}
.dock .acts{display:flex;gap:10px;flex-wrap:wrap;justify-content:flex-end}
.tpanel{border-bottom:1px solid var(--line2)}
.dock.open .tchips{display:none}
.tpanel .in{padding-top:14px;padding-bottom:12px;max-height:46vh;overflow:auto}
.tpanel .ph{display:flex;flex-wrap:wrap;align-items:baseline;justify-content:space-between;gap:4px 12px;margin-bottom:10px}
.tpanel .ph b{font:600 15px/1.3 var(--hd)}
.tpanel .wrap{display:flex;flex-wrap:wrap;gap:6px}
@keyframes drop{from{transform:translateY(-16px) scale(.85);opacity:0}to{transform:none;opacity:1}}
@keyframes bump{0%,100%{transform:none}35%{transform:translateY(3px) scale(.94)}70%{transform:translateY(-2px)}}
.loading{max-width:1180px;margin:0 auto;padding:80px 24px;text-align:center;color:var(--muted)}
.errc{max-width:640px;margin:60px auto;padding:24px;background:var(--surface);border:1px solid var(--line);border-radius:18px}
.errc h2{margin:0 0 6px;font:600 18px/1.3 var(--hd)}
.errc p{margin:0;color:var(--muted)}
.toast{position:fixed;left:50%;bottom:96px;transform:translateX(-50%);max-width:min(92vw,620px);background:var(--ink);color:var(--surface);padding:10px 16px;border-radius:12px;font-size:14px;opacity:0;pointer-events:none;transition:opacity .2s;z-index:9}
.toast.on{opacity:1}
.toast.err{background:var(--err);color:#fff}
/* หน้าส่งเข้าระบบ (เปิดจากปุ่ม “ดึง Playlist วันนี้”) */
.sendp{min-height:100vh;display:grid;place-items:center;padding:24px 16px}
.pp{width:100%;max-width:480px;background:var(--surface);border:1px solid var(--line);border-radius:18px;padding:28px 32px 22px;display:flex;flex-direction:column;gap:22px;box-shadow:0 20px 60px -20px rgba(11,30,70,.35)}
.pp .brand b{font-size:17px}
.pp h1{margin:0;font:700 28px/1.25 var(--hd)}
.pp .when{margin:4px 0 16px;color:var(--muted);font-size:14px}
.prog{height:10px;border-radius:99px;background:var(--line2);overflow:hidden}
.prog span{display:block;height:100%;border-radius:99px;background:var(--blue);transition:width .3s}
.steps{margin:0;padding:0;list-style:none;display:flex;flex-direction:column;gap:12px}
.steps li{display:flex;align-items:center;gap:12px;color:var(--muted)}
.steps .i{width:24px;height:24px;border-radius:99px;display:grid;place-items:center;flex:none;border:3px solid var(--line)}
.steps .done{color:var(--ink)}
.steps .done .i{background:var(--blue);border-color:var(--blue)}
.steps .now{color:var(--ink);font-weight:600}
.steps .now .i{border-color:var(--blue)}
.ppf{display:flex;flex-wrap:wrap;justify-content:space-between;align-items:center;gap:6px 12px;padding-top:14px;border-top:1px solid var(--line2);font-size:13px;color:var(--muted)}
.pp .dayrow{grid-template-columns:1fr;gap:8px}
.pp .choice{display:flex;flex-wrap:wrap;gap:10px}
.pp .choice .btn{flex:1 1 auto}
@media (max-width:1000px){.slots{grid-template-columns:repeat(4,minmax(0,1fr))}}
@media (max-width:640px){
 .in{padding:0 16px}
 .top .in{min-height:0;padding-top:12px;padding-bottom:12px}
 .tools,.date{width:100%}
 .date{justify-content:space-between}
 .tools .btn{flex:1 1 auto}
 main{padding:18px 0 150px}
 main .in{gap:16px}
 .board{padding:20px 18px 16px;border-radius:18px}
 .count .n{font-size:48px}
 .count .t{font-size:17px}
 .auto{display:none}
 .hours span:nth-child(even){visibility:hidden}
 .drop .chip,.drop .stem{display:none}
 .track{height:56px}
 .nowm b{left:auto;right:8px}
 .drop .pin{width:12px;height:12px;left:-6px}
 .strip,.more{grid-template-columns:1fr;padding:14px}
 .catch .act{grid-column:1}
 .catch .act .btn{width:100%}
 .rack{padding:4px 14px;border-radius:18px}
 .row{grid-template-columns:1fr;gap:8px;padding:16px 0}
 .row h2{padding-top:0}
 .row h2 small{display:inline;margin-left:6px}
 .slots{grid-template-columns:repeat(2,minmax(0,1fr));padding:8px;gap:8px}
 .slot.w2{grid-column:auto}
 .dock>.in{flex-wrap:wrap;gap:8px 10px;min-height:0}
 .tchips{display:none}
 .tray{flex:1 1 100%}
 .dock .acts{width:100%;flex-wrap:nowrap}
 .dock .acts .btn{padding:0 12px}
 .dock .acts .btn.pri{flex:1}
 .dock .acts .btn .lbl{display:none}
 .toast{bottom:150px}
 .pp{padding:24px 20px 18px}
}
@media (prefers-reduced-motion:reduce){*{transition:none!important;animation:none!important}}
</style></head><body>
<div id="app"><p class="loading">กำลังโหลด…</p></div>
<div class="toast" id="toast" role="status" aria-live="polite"></div>
<script>
var D = null, SEL = {}, BUSY = '', TT = 0, PROG = 0, ZTOT = 0, CSEL = {}, SENDASK = false, SENDDAYS = [], TRAY = false, NEWID = '';
var SEND = ${send ? 'true' : 'false'}, STUDIO = ${JSON.stringify(CONFIG.studioOrigin || '')}, SENT = false, SENDMSG = '', SENDOK = false, SENTN = 0, SENDINFO = false;
var TH = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
var WD = ['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'];
function sv(p, s, c, w) { return '<svg aria-hidden="true" width="' + s + '" height="' + s + '" viewBox="0 0 24 24" fill="none" stroke="currentColor"' + (c ? ' style="color:' + c + '"' : '') + ' stroke-width="' + (w || 2) + '" stroke-linecap="round" stroke-linejoin="round">' + p + '</svg>'; }
var P = { check: '<path d="M5 12.5l4.5 4.5L19 7.5"/>', clock: '<circle cx="12" cy="12" r="8.5"/><path d="M12 8v4.5l3 1.8"/>', left: '<path d="M15 6l-6 6 6 6"/>', right: '<path d="M9 6l6 6-6 6"/>',
  cal: '<rect x="3.5" y="5" width="17" height="15" rx="3"/><path d="M3.5 10h17M8 3v4M16 3v4"/>', ext: '<path d="M7 17L17 7M9 7h8v8"/>', info: '<circle cx="12" cy="12" r="8.5"/><path d="M12 11v5M12 8h.01"/>',
  refresh: '<path d="M20 11a8 8 0 0 0-14.6-4.5M4 4v4h4M4 13a8 8 0 0 0 14.6 4.5M20 20v-4h-4"/>', folder: '<path d="M3.5 7.5A2.5 2.5 0 0 1 6 5h4l2 2h6a2.5 2.5 0 0 1 2.5 2.5v7A2.5 2.5 0 0 1 18 19H6a2.5 2.5 0 0 1-2.5-2.5z"/>',
  down: '<path d="M12 4v11M7 10.5l5 5 5-5M5 20h14"/>', x: '<path d="M7 7l10 10M17 7L7 17"/>', tray: '<path d="M3 13h5l2 3h4l2-3h5v5a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><path d="M12 3v9M8.5 8.5 12 12l3.5-3.5"/>',
  rev: '<path d="M20 12a8 8 0 1 1-2.3-5.6M20 4v5h-5"/>' };
var LOGO = '<svg width="44" height="44" viewBox="0 0 44 44" aria-hidden="true"><rect width="44" height="44" rx="12" fill="#0B1E46"/><path d="M15 8h11l5 5v10H15z" fill="#fff"/><path d="M26 8v5h5" fill="#C9D6F5"/><path d="M18 15h8M18 18.5h6" stroke="#1E8E5A" stroke-width="2" stroke-linecap="round"/><path d="M9 25h7l2.5 4h7l2.5-4h7v8a3 3 0 0 1-3 3H12a3 3 0 0 1-3-3z" fill="#EE7D22"/><path d="M22 20v6.5M19 24l3 3 3-3" stroke="#0B1E46" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" fill="none"/></svg>';
var SHEET = '<svg class="sheet" aria-hidden="true" width="22" height="26" viewBox="0 0 22 26"><path d="M2 3a2 2 0 0 1 2-2h10l6 6v16a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2z" fill="var(--sheet)"/><path d="M14 1v6h6" fill="#fff" opacity=".55"/><path d="M6 12h10M6 15.5h10M6 19h6" stroke="#fff" stroke-width="1.6" stroke-linecap="round"/></svg>';
var BOX = '<span class="box" aria-hidden="true">' + sv(P.check, 13, 'var(--surface)', 3.2) + '</span>';
function $(s) { return document.querySelector(s); }
function esc(s) { return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function dt(d) { return new Date(d + 'T12:00:00Z'); }
function thLong(d) { var t = dt(d); return WD[t.getUTCDay()] + ' ' + t.getUTCDate() + ' ' + TH[t.getUTCMonth()] + ' ' + (t.getUTCFullYear() + 543); }
function thShort(d) { var t = dt(d); return t.getUTCDate() + ' ' + TH[t.getUTCMonth()]; }
function dayName(d) { var t = dt(d); return WD[t.getUTCDay()] + ' ' + t.getUTCDate() + ' ' + TH[t.getUTCMonth()]; }
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
function load(day) {
  BUSY = 'load'; draw();
  call('webDay', [day || ''], function (d) {
    D = d; BUSY = ''; defaults(); CSEL = {}; TRAY = false;
    if (D.catchup) D.catchup.days.forEach(function (a) { CSEL[a.day] = true; });
    if (SEND && !SENT) { SENT = true; if (D.catchup && opener_()) { SENDASK = true; draw(); } else { draw(); sendStudio([D.day]); } }
    else draw();
  });
}
/* บอกสคริปต์ว่าดึงวันไหนแล้ว (กล่อง “ตามงานวันหยุด” จะไม่ขึ้นวันนั้นอีก) — เงียบ ๆ ไม่ต้องรอผล */
function mark(days, how) { days = days.filter(function (d) { return d <= D.today; }); if (days.length) try { google.script.run.withFailureHandler(function () {}).webMark(days, how || 'd'); } catch (e) {} }
function catchDays() { return D && D.catchup ? D.catchup.days.filter(function (a) { return CSEL[a.day]; }) : []; }
/* ---------- ข้อมูล ---------- */
function kindBy(name) { for (var i = 0; i < D.kinds.length; i++) if (D.kinds[i].name === name) return D.kinds[i]; return null; }
function shortName(k) { var g = k.group || ''; return g && k.name.indexOf(g + ' ') === 0 ? k.name.slice(g.length + 1) : k.name; }
function gShort(g) { return String(g || 'อื่น ๆ').split(' ')[0]; }
function byKind() { var by = {}, oth = []; D.files.forEach(function (f) { if (f.kind) (by[f.kind] = by[f.kind] || []).push(f); else oth.push(f); }); return { by: by, oth: oth }; }
function mainKinds() { return D.kinds.filter(function (k) { return k.daily && !k.line; }); }
function ndKinds() { return D.kinds.filter(function (k) { return !k.daily && !k.line; }); }
function isExtra(f) { var k = kindBy(f.kind); return !!(k && !k.daily); }
function revOf(f) { var R = D.revs || []; for (var i = 0; i < R.length; i++) if (R[i].name === f.name) return R[i]; return null; }
/* ค่าเริ่มต้น: ติ๊ก Playlist ทุกแบบที่มา (แบบละ 1 ไฟล์ = ฉบับที่เข้าโฟลเดอร์ล่าสุด) ไม่ติ๊ก: แบบที่ระบบไม่ได้ใช้ (unused) • ฉบับเก่า • ไฟล์ที่ไม่อยู่ในรายการ */
function defaults() {
  SEL = {}; var g = byKind();
  D.kinds.forEach(function (k) { var L = g.by[k.name]; if (!L || k.unused) return; SEL[L[L.length - 1].id] = true; });
}
function picked() { return D ? D.files.filter(function (f) { return SEL[f.id]; }) : []; }
function selBtn(L, grp) { var all = L.length && L.every(function (f) { return SEL[f.id]; }); return '<button class="lk" data-a="sel" data-g="' + grp + '">' + (all ? 'ไม่เลือก' : 'เลือกทั้งหมด') + '</button>'; }
function lastSeen(k, before) { var L = (k.days || []).filter(function (d) { return d < before; }); return L.length ? L[L.length - 1] : ''; }
/* รอบการส่งของแบบที่ไม่ได้ส่งทุกวัน: ดูจากวันที่เคยส่งมา (ระยะห่างที่พบบ่อย) */
function freq(k) {
  var L = k.days || [], G = [];
  for (var i = 1; i < L.length; i++) G.push(Math.round((dt(L[i]) - dt(L[i - 1])) / 864e5));
  if (G.length < 2) return '';
  G.sort(function (a, b) { return a - b; });
  var m = G[Math.floor(G.length / 2)];
  return m <= 1 ? 'ส่งมาเกือบทุกวัน' : m === 2 ? 'ส่งมาวันเว้นวัน' : 'ส่งมาราว ' + m + ' วันครั้ง';
}
/* ---------- ส่งเข้าระบบ ---------- */
/* เปิดจากปุ่ม “ดึง Playlist วันนี้” ในระบบ: รวม Playlist ที่ติ๊กไว้เป็น .zip แล้วส่งกลับให้หน้าระบบ (ส่งให้เฉพาะ studioOrigin) */
function opener_() { try { return window.top.opener || null; } catch (e) { return null; } }
function sendStudio(days) {
  var w = opener_();
  SENDASK = false; SENDDAYS = days || [D.day];
  if (!w || !STUDIO) { SENDINFO = true; SENDMSG = 'หน้านี้ไม่ได้เปิดจากปุ่ม “ดึง Playlist วันนี้” ในระบบ (หรือเบราว์เซอร์ไม่ให้ส่งกลับ) จึงยังไม่ได้ส่งไฟล์เข้าระบบ — ถ้าจะส่ง กดปุ่มนั้นใน PHOTO CHECK STUDIO หรือกด “ดาวน์โหลด” แล้วลากไฟล์ .zip ลงขั้นที่ 1'; draw(); return; }
  var g = byKind(), main = mainKinds(), miss = main.filter(function (k) { return !g.by[k.name]; }).map(function (k) { return k.name; });
  var multi = SENDDAYS.length > 1, rev = (D.revs || []).map(function (r) { return r.kind; });
  if (!multi && !picked().length) { try { w.postMessage({ psPL: 1, ok: false, day: D.day, err: 'ยังไม่มี Playlist ของวันนี้ในโฟลเดอร์' }, STUDIO); } catch (e) {} SENDMSG = 'ยังไม่มี Playlist ของวันนี้ — รอระบบดึงจากอีเมล (ทุกชั่วโมง) หรือกด “ดึงตอนนี้”'; draw(); return; }
  var info = multi ? D.catchup.days.filter(function (a) { return SENDDAYS.indexOf(a.day) >= 0; }).map(function (a) { return a.day === D.day ? { day: a.day, have: main.length - miss.length, total: main.length, miss: miss, rev: rev } : { day: a.day, have: a.have, total: a.total, miss: a.miss, rev: a.rev || [] }; }) : null;
  (multi ? zipDays : zip)(function (name, blob, n) {
    try { w.postMessage({ psPL: 1, ok: true, day: D.day, name: name, blob: blob, n: n, have: main.length - miss.length, total: main.length, miss: miss, extra: picked().filter(isExtra).length, rev: rev, days: info }, STUDIO); }
    catch (e) { SENDMSG = 'ส่งเข้าระบบไม่ได้: ' + (e && e.message || e) + ' — กด “ดาวน์โหลด” แล้วลากไฟล์ .zip เองแทน'; draw(); return; }
    SENDOK = true; SENTN = n; SENDMSG = ''; mark(SENDDAYS); draw();
    setTimeout(function () { try { window.top.close(); } catch (e) {} }, 800);
  });
}
/* ---------- ส่วนต่าง ๆ ของหน้า ---------- */
function brand(sub) { return '<div class="brand">' + LOGO + '<div><b>PLAYLIST INBOX</b><small>' + esc(sub) + '</small></div></div>'; }
function ring(h, t) { var p = t ? Math.round(Math.min(h, t) / t * 100) : 0; return '<span class="ring" style="--p:' + p + '" aria-hidden="true"><span>' + h + '/' + t + '</span></span>'; }
function dayCard(a, opt) {
  var ds = a.n + ' ไฟล์ · ' + (a.miss.length ? 'ขาด ' + (a.miss.length > 2 ? a.miss.length + ' แบบ' : a.miss.join(', ')) : 'ครบ');
  var rd = a.redo && a.redo.length ? a.redo : null;
  var rv = a.rev && a.rev.length ? '<span class="ds rv">' + (rd ? 'ดึงไปแล้ว แต่มีฉบับแก้ไข: ' + esc(rd.join(', ')) : 'มีฉบับแก้ไข: ' + esc(a.rev.join(', '))) + '</span>' : '';
  var body = ring(a.have, a.total) + '<span class="dt"><span class="dn">' + esc(dayName(a.day)) + (a.day === D.today ? '<span class="tag">วันนี้</span>' : '') + '</span><span class="ds">' + esc(ds) + '</span>' + rv +
    (opt.go ? '<button class="go" data-go="' + a.day + '">ดูรายการ</button>' : '') + '</span>';
  if (opt.pick) return '<label class="day"><input type="checkbox" data-cd="' + a.day + '"' + (CSEL[a.day] ? ' checked' : '') + ' aria-label="' + esc(dayName(a.day)) + '">' + BOX + body + '</label>';
  if (opt.link) return '<button class="day" data-go="' + a.day + '" title="ดู Playlist ของวันนั้น">' + body + '</button>';
  return '<div class="day">' + body + '</div>';
}
function catchLabel() {
  var L = catchDays(), n = 0;
  L.forEach(function (a) { n += a.day === D.day ? picked().length : a.n; });
  if (BUSY === 'zipall') return 'กำลังเตรียมไฟล์ ' + PROG + '/' + ZTOT + '…';
  return L.length ? 'ดาวน์โหลด ' + L.length + ' วัน · ' + n + ' ไฟล์ (.zip)' : 'ยังไม่ได้เลือกวัน';
}
function catchHtml() {
  var C = D.catchup.days, prev = C.filter(function (a) { return a.day !== D.today; }), onlyRedo = prev.length && prev.every(function (a) { return a.redo && a.redo.length; });
  var title = onlyRedo ? 'ดึงใหม่: มีฉบับแก้ไข' : 'ตามงานวันหยุด';
  var sub = onlyRedo ? 'วันที่ดึงไปแล้ว แต่มีฉบับแก้ไขส่งมาทีหลัง' : 'วันที่ยังไม่ได้ดึง รวมเป็น .zip เดียว แยกโฟลเดอร์ตามวัน';
  return '<section class="strip catch" aria-label="ตามงานวันหยุด"><h3>' + title + '<small>' + sub + '</small></h3>' +
    '<div class="dayrow">' + C.map(function (a) { return dayCard(a, { pick: 1, go: a.day !== D.today }); }).join('') + '</div>' +
    '<div class="act"><span class="note">ลากไฟล์ .zip ลงขั้นที่ 1 ทีเดียว ระบบเรียงวันให้เอง · <button class="lk" data-a="cskip" title="ดึง Playlist ของวันก่อนหน้าไปแล้วทางอื่น — ไม่ต้องแสดงกล่องนี้อีก">ดึงไปแล้ว ไม่ต้องแสดง</button></span>' +
    '<button class="btn pri" id="zipall" data-a="zipall"' + (!catchDays().length || BUSY ? ' disabled' : '') + '>' + sv(P.down, 18, 0, 2.2) + esc(catchLabel()) + '</button></div></section>';
}
function aheadHtml() {
  return '<section class="strip" aria-label="Playlist ล่วงหน้า"><h3>กำลังจะมา<small>ส่งมาก่อนวันฉาย เก็บเข้าโฟลเดอร์ของวันนั้นแล้ว</small></h3><div class="dayrow">' +
    D.ahead.map(function (a) { return dayCard(a, { link: 1 }); }).join('') + '</div></section>';
}
function hm2pos(hm) { var p = String(hm).split(':'), s = +D.startHour, e = +D.endHour + 1; return Math.max(0, Math.min(100, ((+p[0] + (+p[1] || 0) / 60) - s) / (e - s) * 100)); }
function dropLabel(F) {
  if (F.length === 1) { var k = kindBy(F[0].kind); return k ? k.name : 'ไฟล์อื่น'; }
  var cnt = {}, order = [];
  F.forEach(function (f) { var k = kindBy(f.kind), g = k ? gShort(k.group) : 'อื่น ๆ'; if (!cnt[g]) { cnt[g] = 0; order.push(g); } cnt[g]++; });
  order.sort(function (a, b) { return cnt[b] - cnt[a]; });
  return order.length === 1 ? order[0] : order.slice(0, 2).join(' + ') + (order.length > 2 ? ' ฯลฯ' : '');
}
function drops() {
  var L = D.files.filter(function (f) { return f.ad === D.day; }).sort(function (a, b) { return a.at - b.at; }), C = [];
  L.forEach(function (f) { var c = C[C.length - 1]; if (c && f.at - c.at0 <= 15 * 60000) c.f.push(f); else C.push({ at0: f.at, hm: f.hm, f: [f] }); });
  var out = C.map(function (c) { return { pos: hm2pos(c.hm), em: c.hm, label: dropLabel(c.f), n: c.f.length }; });
  var pre = D.files.filter(function (f) { return f.ad && f.ad < D.day; });
  if (pre.length) out.unshift({ pos: 0, em: 'ก่อนวันนี้', label: dropLabel(pre), n: pre.length, pre: 1 });
  return out;
}
function boardHtml(g, isToday, future) {
  var main = mainKinds(), nd = ndKinds();
  var have = main.filter(function (k) { return g.by[k.name] || (future && k.latest); }), miss = main.filter(function (k) { return !(g.by[k.name] || (future && k.latest)); });
  var ndh = nd.filter(function (k) { return g.by[k.name]; });
  var t, sub = [];
  if (!D.files.length) {
    t = future ? 'ยังไม่มีไฟล์ล่วงหน้าของวันนั้น' : isToday ? 'ยังไม่มี Playlist ของวันนี้' : 'ไม่มี Playlist ของวันนั้น';
    sub.push(future ? 'Playlist ที่ส่งมาก่อนวันฉาย ระบบเก็บเข้าโฟลเดอร์ของวันนั้นให้เอง' : isToday ? (D.auto === false ? 'ดึงอัตโนมัติปิดอยู่ กด “ดึงตอนนี้” หรือเปิดด้วย setup ในหน้าโค้ด' : 'ระบบดึงจากอีเมลให้เองทุกชั่วโมง ' + D.startHour + ':00–' + D.endHour + ':59 หรือกด “ดึงตอนนี้”') : 'ระบบไม่ได้เก็บ Playlist ของวันนั้นไว้ในโฟลเดอร์');
  } else if (!miss.length) {
    t = future ? 'ครบทุกแบบที่ต้องมา (ล่วงหน้า)' : isToday ? 'ครบทุกแบบที่ต้องมาวันนี้' : 'ครบทุกแบบที่ต้องมา';
    sub.push(future ? 'ส่งมาก่อนวันฉาย เก็บไว้ในโฟลเดอร์ของวันนั้นแล้ว ถึงวันดาวน์โหลดหรือกด “ดึง Playlist วันนี้” ในระบบได้ตามปกติ' : isToday ? 'ดาวน์โหลด หรือกด “ดึง Playlist วันนี้” ในระบบได้เลย' : 'ดาวน์โหลดแล้วลากไฟล์ .zip ลงขั้นที่ 1 ของ PHOTO CHECK STUDIO ได้เลย');
  } else {
    t = future ? 'แบบส่งมาล่วงหน้าแล้ว' : 'แบบที่ต้องมาทุกวัน มาแล้ว';
    var names = miss.map(function (k) { return '<b>' + esc(k.name) + '</b>'; });
    sub.push((future || !isToday ? 'ขาด ' : 'ยังขาด ') + (miss.length > 4 ? miss.length + ' แบบ: ' + names.slice(0, 4).join(', ') + ' และอีก ' + (miss.length - 4) + ' แบบ' : andList(names)) +
      (isToday && D.auto !== false ? ' · ระบบตรวจอีเมลให้ทุกชั่วโมงจนถึง ' + D.endHour + ':59' : '') +
      (future ? ' · ส่งมาก่อนวันฉาย ถึงวันระบบดึงที่เหลือให้ตามปกติ' : ''));
  }
  if (D.files.length && nd.length) sub.push('แบบที่ไม่ได้ส่งทุกวัน ' + (future ? 'ส่งมาล่วงหน้าแล้ว ' : isToday ? 'วันนี้มาแล้ว ' : 'วันนั้นมา ') + ndh.length + ' จาก ' + nd.length + ' แบบ (ไม่นับว่าขาด)');
  var h = '<section class="board" aria-label="สรุป Playlist"><div class="bhead"><div><div class="kicker">Playlist ' + (future ? 'ล่วงหน้า ' : '') + esc(thLong(D.day)) + '</div>' +
    '<div class="count"><span class="n">' + have.length + '<span>/' + main.length + '</span></span><span class="t">' + esc(t) + '<small>' + sub.join('<br>') + '</small></span></div></div>' +
    (isToday ? '<span class="auto' + (D.auto === false ? ' off' : '') + '"><i></i>' + (D.auto === false ? 'ดึงอัตโนมัติปิดอยู่' : 'ตรวจอีเมลทุกชั่วโมง ' + D.startHour + ':00–' + D.endHour + ':59') + '</span>' : '') + '</div>';
  if (!future) {
    var nowp = isToday ? hm2pos(D.now || '0:00') : 100, s = +D.startHour, e = +D.endHour + 1, x;
    h += '<div class="tl"><div class="track" id="track"><div class="rail"><div class="past" style="width:' + nowp + '%"></div>' + (isToday ? '<div class="fut" style="left:' + nowp + '%"></div>' : '') + '</div>';
    for (x = s; x <= e; x++) h += '<span class="tick" style="left:' + ((x - s) / (e - s) * 100) + '%"></span>';
    drops().forEach(function (d) { h += '<span class="drop' + (d.pre ? ' pre edge' : '') + '" style="left:' + d.pos + '%"><span class="stem"></span><span class="chip"><em>' + esc(d.em) + '</em>' + esc(d.label) + (d.n > 1 ? ' ×' + d.n : '') + '</span><span class="pin"></span></span>'; });
    if (isToday) h += '<span class="nowm" id="nowm" style="left:' + nowp + '%"><b>ตอนนี้ ' + esc(D.now || '') + '</b></span>';
    h += '<div class="hours">'; for (x = s; x <= e; x++) h += '<span>' + (x < 10 ? '0' : '') + x + '</span>'; h += '</div></div></div>';
  }
  return h + '</section>';
}
function revHtml(isToday) {
  var R = D.revs || [];
  if (!R.length) return '';
  var after = R.some(function (r) { return r.after; }), names = andList(R.map(function (r) { return esc(r.kind); }));
  var when = R.map(function (r) { return r.time; }).join(', ');
  return '<section class="revnote" role="status" aria-label="ฉบับแก้ไข"><span class="ri">' + sv(P.rev, 18, 0, 2.4) + '</span><div class="rt"><b>มีฉบับแก้ไขส่งมาใหม่ ' + R.length + ' แบบ</b>' +
    names + ' ' + (isToday ? 'ของวันนี้' : 'ของวันนั้น') + ' ส่งมาใหม่เมื่อ ' + esc(when) + ' · ระบบใช้ฉบับใหม่แทนให้แล้ว ฉบับเก่าอยู่ในถังขยะของ Drive (กู้คืนได้ 30 วัน)' +
    (after ? ' · <b style="display:inline;font:inherit;font-weight:700">วันนั้นดึงเข้าระบบไปก่อนฉบับแก้ไขมา ให้ดึงใหม่อีกครั้ง</b>' : ' · ถ้าดึงเข้า PHOTO CHECK STUDIO ไปก่อนหน้านี้ ให้ดึงใหม่อีกครั้ง') + '</div>' +
    '<a class="btn" href="https://drive.google.com/drive/trash" target="_blank" rel="noopener">ดูฉบับเก่า</a>' + (after ? '<button class="btn dark" data-a="zip">' + sv(P.down, 16, 0, 2.2) + 'ดาวน์โหลดใหม่</button>' : '') + '</section>';
}
function trimName(f, wide) { return wide || f.name.length <= 26 ? f.name : '…' + f.name.slice(-24); } // ช่องเล็ก: แสดงท้ายชื่อไฟล์ (ส่วนที่ต่างกัน) ชื่อเต็มอยู่ใน title
function slot(k, L, ctx, wide) {
  var w = wide ? ' w2' : '', nm = '<span class="nm">' + esc(shortName(k)) + '</span>';
  var ndtag = !k.daily ? '<span class="tagnd">ไม่ทุกวัน</span>' : '';
  if (L) {
    var f = L[L.length - 1], rv = revOf(f), mm = '';
    var st = f.old ? '<span class="stamp info">' + sv(P.info, 13, 0, 2.6) + 'ฉบับ ' + esc(f.old) + '</span>' : '<span class="stamp">' + sv(P.check, 13, 0, 3) + esc(f.time) + '</span>';
    if (rv) mm = 'ส่งมาใหม่แทนฉบับ' + (rv.oldTime ? 'เมื่อ ' + rv.oldTime : 'เดิม') + ' · ฉบับเก่าอยู่ในถังขยะ Drive';
    else if (f.old) mm = k.daily ? 'วันนี้ยังไม่มีฉบับใหม่ ใช้ฉบับล่าสุดไปก่อน' : 'ไม่ได้ส่งทุกวัน ใช้ฉบับล่าสุด';
    else if (L.length > 1) mm = 'มี ' + L.length + ' ฉบับ ติ๊กไว้ให้ฉบับล่าสุด';
    return '<label class="slot' + w + (rv ? ' rev' : '') + '" title="' + esc(f.name + ' · ' + kb(f.size)) + '"><input type="checkbox" data-id="' + esc(f.id) + '"' + (SEL[f.id] ? ' checked' : '') + ' aria-label="' + esc(k.name) + '">' +
      SHEET + nm + '<span class="srow">' + st + (rv ? '<span class="tagrev">ฉบับแก้ไข</span>' : '') + ndtag + '</span>' + (mm ? '<span class="mm">' + esc(mm) + '</span>' : '') +
      (f.old ? '' : '<span class="fn">' + esc(trimName(f, wide)) + '</span>') + BOX + '</label>';
  }
  if (k.line) return '<div class="slot line' + w + '">' + nm + '<span class="srow"><span class="stamp">' + sv(P.info, 13, 0, 2.6) + 'มาทาง LINE</span></span><span class="hint">ดาวน์โหลดจาก LINE แล้วลากลงระบบพร้อมไฟล์ .zip</span></div>';
  if (ctx.future && k.latest) return '<div class="slot line' + w + '">' + nm + '<span class="srow"><span class="stamp">' + sv(P.info, 13, 0, 2.6) + 'ใช้ฉบับล่าสุด</span>' + ndtag + '</span><span class="hint">ไม่ได้ส่งทุกวัน ถึงวันระบบใช้ฉบับล่าสุดให้เอง</span></div>';
  if (!k.daily) {
    var ls = lastSeen(k, D.day), fq = k.tip || freq(k);
    var none = ctx.future ? 'ยังไม่มีไฟล์ล่วงหน้า' : ctx.isToday ? 'วันนี้ยังไม่มี' : 'วันนั้นไม่มี';
    return '<div class="slot nd' + w + '">' + nm + '<span class="srow">' + (ls ? '<span class="stamp grey">ล่าสุด ' + esc(thShort(ls)) + '</span>' : '') + ndtag + '</span><span class="fn">' + esc(none + (fq ? ' · ' + fq : '')) + '</span></div>';
  }
  var hint = ctx.future ? 'ยังไม่มีไฟล์ล่วงหน้าของวันนั้น' : !ctx.isToday ? 'ไม่มีไฟล์ของวันนั้นในโฟลเดอร์' : k.tip ? k.tip : k.latest ? 'ไม่ได้ส่งทุกวัน ยังไม่มีฉบับไหนในอีเมล' : 'ยังไม่มีอีเมลวันนี้';
  return '<div class="slot wait' + w + '">' + nm + '<span class="srow"><span class="stamp">' + sv(P.clock, 13, 0, 2.6) + (ctx.isToday ? 'รออยู่' : ctx.future ? 'ยังไม่มี' : 'ไม่มี') + '</span></span><span class="hint">' + esc(hint) + '</span></div>';
}
function rackHtml(g, ctx) {
  var h = '<section class="rack" aria-label="Playlist ทุกแบบ">';
  D.groups.forEach(function (gn) {
    var K = D.kinds.filter(function (k) { return (k.group || 'อื่น ๆ') === gn; });
    if (!K.length) return;
    var d = 0, dh = 0, nt = 0, ng = 0, small = K.length >= 4;
    K.forEach(function (k) { var got = !!g.by[k.name]; if (k.line) return; if (k.daily) { d++; if (got || (ctx.future && k.latest)) dh++; } else { nt++; if (got) ng++; } });
    var gc = (d ? '<span class="gc' + (dh < d ? ' part' : '') + '">' + sv(dh < d ? P.clock : P.check, 12, 0, 3) + dh + '/' + d + ' ทุกวัน</span>' : '') + (nt ? '<span class="gc nd">ไม่ทุกวัน มา ' + ng + '/' + nt + '</span>' : '');
    h += '<div class="row"><h2>' + esc(gn) + '<small>' + K.length + ' แบบ</small>' + (gc ? '<span class="gcs">' + gc + '</span>' : '') + '</h2><div class="slots">' +
      K.map(function (k) { return slot(k, g.by[k.name], ctx, !small); }).join('') + '</div></div>';
  });
  h += '<div class="legend"><span><i style="background:var(--surface);border:1px solid var(--line)"></i>มาแล้ว (ติ๊กไว้ให้)</span><span><i style="border:1.5px dashed var(--waitLine);background:var(--waitSoft)"></i>ต้องมาทุกวัน แต่ยังไม่มา</span>' +
    '<span><i style="background:var(--slotIn);border:1px solid var(--line2)"></i>ไม่ได้ส่งทุกวัน ยังไม่มี (ไม่นับว่าขาด)</span>' + ((D.revs || []).length ? '<span><i style="border:2px solid var(--blue)"></i>ฉบับแก้ไข</span>' : '') +
    (D.kinds.some(function (k) { return k.line; }) ? '<span><i style="background:var(--blueSoft);border:1.5px dashed var(--blueLine)"></i>มาทาง LINE</span>' : '') + '</div>';
  return h + '</section>';
}
function mini(f, label) {
  return '<label class="mini" title="' + esc(f.name) + '"><input type="checkbox" data-id="' + esc(f.id) + '"' + (SEL[f.id] ? ' checked' : '') + ' aria-label="' + esc(label) + '">' +
    '<span class="mt"><b>' + esc(label) + '</b>' + (label !== noext(f.name) ? '<span class="fn">' + esc(f.name) + '</span>' : '') +
    '<span class="mm">' + (f.pd ? 'วันที่ ' + esc(thShort(f.pd)) + ' · ' : '') + 'เข้า ' + esc(f.time) + ' · ' + esc(kb(f.size)) + '</span></span></label>';
}
function moreHtml(g) {
  var older = [];
  D.kinds.forEach(function (k) { var L = g.by[k.name]; if (L && L.length > 1) L.slice(0, -1).forEach(function (f) { older.push({ f: f, k: k }); }); });
  if (!older.length && !g.oth.length) return '';
  var h = '<section class="more" aria-label="ไฟล์อื่นในโฟลเดอร์"><h3>ไฟล์อื่นในโฟลเดอร์<small>ไม่ได้ติ๊กไว้ให้ ติ๊กเองได้ถ้าต้องการ</small></h3><div>';
  if (older.length) h += '<div class="sub"><span>แบบเดียวกันมีมากกว่า 1 ไฟล์ (ติ๊กไว้ให้ไฟล์ที่เข้าล่าสุดแล้ว)</span></div><div class="minis">' + older.map(function (o) { return mini(o.f, o.k.name); }).join('') + '</div>';
  if (g.oth.length) h += '<div class="sub"><span>ไม่อยู่ในรายการ แต่ระบบเก็บไว้ให้</span>' + selBtn(g.oth, 'oth') + '</div><div class="minis">' + g.oth.map(function (f) { return mini(f, noext(f.name)); }).join('') + '</div>';
  return h + '</div></section>';
}
/* ---------- ถาดรับ: ไฟล์ที่เลือกไว้ ---------- */
function chipLabel(f) {
  var k = kindBy(f.kind);
  if (!k) { var n = noext(f.name); return n.length > 18 ? n.slice(0, 17) + '…' : n; }
  var s = shortName(k), dup = D.kinds.some(function (x) { return x !== k && shortName(x) === s; });
  return dup ? k.name : s;
}
function tchip(f) { return '<button class="tchip' + (f.id === NEWID ? ' new' : '') + '" data-un="' + esc(f.id) + '" title="เอาออกจากถาด: ' + esc(f.name) + '">' + esc(chipLabel(f)) + sv(P.x, 12, 0, 2.6) + '</button>'; }
function trayOrder() {
  var Pk = picked(), idx = {};
  D.kinds.forEach(function (k, i) { idx[k.name] = i; });
  return Pk.sort(function (a, b) { var x = a.kind in idx ? idx[a.kind] : 999, y = b.kind in idx ? idx[b.kind] : 999; return x - y || a.at - b.at; });
}
function dockHtml() {
  var Pk = trayOrder(), sz = 0, isToday = D.day === D.today, url = D.folderUrl || D.parentUrl, catchup = isToday && D.catchup && !SEND;
  Pk.forEach(function (f) { sz += f.size; });
  var label = BUSY === 'zip' ? 'กำลังเตรียมไฟล์ ' + PROG + '/' + Pk.length + '…' : Pk.length ? (catchup ? 'เฉพาะวันนี้ ' : 'ดาวน์โหลด ') + Pk.length + ' ไฟล์ (.zip)' : 'ยังไม่ได้เลือกไฟล์';
  var h = '';
  if (TRAY) h += '<div class="tpanel" id="tpanel"><div class="in"><div class="ph"><b>ไฟล์ในถาด ' + Pk.length + ' ไฟล์</b><span><button class="lk" data-a="reset">ติ๊กตามค่าเริ่มต้น</button> · <button class="lk" data-a="clear">เอาออกทั้งหมด</button> · <button class="lk" data-a="tray">ปิด</button></span></div>' +
    (Pk.length ? '<div class="wrap">' + Pk.map(tchip).join('') + '</div>' : '<span class="tempty">ยังไม่มีไฟล์ในถาด ติ๊กที่ช่อง Playlist ด้านบนเพื่อใส่ไฟล์</span>') + '</div></div>';
  h += '<div class="in"><div class="tray"><button class="ico" id="trayico" data-a="tray" aria-expanded="' + (TRAY ? 'true' : 'false') + '" aria-label="ดูไฟล์ในถาดทั้งหมด" title="ดูไฟล์ในถาดทั้งหมด">' + sv(P.tray, 22) + (Pk.length ? '<em>' + Pk.length + '</em>' : '') + '</button>' +
    '<div class="tt"><b>' + (Pk.length ? 'เลือกไว้ ' + Pk.length + ' ไฟล์ · ' + kb(sz) : 'ถาดว่าง') + '</b><div class="tchips" id="tchips">' +
    (Pk.length ? Pk.map(tchip).join('') + '<button class="tmore" data-a="tray" hidden></button>' : '<span class="tempty">ติ๊กที่ช่อง Playlist เพื่อใส่ไฟล์ลงถาด</span>') + '</div></div></div>' +
    '<div class="acts">' + (isToday ? '<button class="btn" data-a="collect"' + (BUSY ? ' disabled' : '') + ' aria-label="ดึงตอนนี้" title="ดึงจากอีเมลตอนนี้">' + sv(P.refresh, 16) + '<span class="lbl">' + (BUSY === 'collect' ? 'กำลังดึง…' : 'ดึงตอนนี้') + '</span></button>' : '') +
    '<a class="btn" href="' + esc(url) + '" target="_blank" rel="noopener" aria-label="เปิดใน Drive" title="เปิดโฟลเดอร์ใน Google Drive">' + sv(P.folder, 16) + '<span class="lbl">เปิดใน Drive</span></a>' +
    '<button class="btn' + (catchup ? '' : ' pri') + '" data-a="zip"' + (!Pk.length || BUSY ? ' disabled' : '') + ' title="' + (Pk.length ? 'รวม ' + kb(sz) : 'ยังไม่ได้เลือกไฟล์') + '">' + sv(P.down, 18, 0, 2.2) + esc(label) + '</button></div></div>';
  return h;
}
/* ชิปในถาดแสดงเท่าที่พอดีบรรทัด ที่เหลือรวมเป็นปุ่ม +N (กดดูทั้งหมด) */
function fitChips() {
  var c = $('#tchips'); if (!c || !c.offsetWidth) return;
  var ch = [].slice.call(c.querySelectorAll('.tchip')), more = c.querySelector('.tmore'); if (!more) return;
  ch.forEach(function (x) { x.hidden = false; }); more.hidden = true;
  var W = c.clientWidth, cut = -1;
  for (var i = 0; i < ch.length; i++) if (ch[i].offsetLeft - c.offsetLeft + ch[i].offsetWidth > W) { cut = i; break; }
  if (cut < 0) return;
  more.hidden = false;
  while (cut > 0) { for (var j = cut; j < ch.length; j++) ch[j].hidden = true; more.textContent = '+' + (ch.length - cut); if (more.offsetLeft - c.offsetLeft + more.offsetWidth <= W) break; ch[cut - 1].hidden = true; cut--; }
  if (cut === 0) { ch.forEach(function (x) { x.hidden = true; }); more.textContent = '+' + ch.length; }
}
function drawDock(bump) {
  var d = $('#dock'); if (!d) return;
  d.className = 'dock' + (TRAY ? ' open' : ''); d.innerHTML = dockHtml(); fitChips(); NEWID = '';
  if (bump) { var i = $('#trayico'); if (i) { i.classList.remove('bump'); void i.offsetWidth; i.classList.add('bump'); } }
}
/* ป้ายเวลาที่ไฟล์เข้า: ชนกันหรือชนป้าย “ตอนนี้” → ยกขึ้นอีกชั้น */
function levelDrops() {
  var tr = $('#track'); if (!tr || innerWidth <= 640) return;
  var nowb = $('#nowm b'), lv = [], mx = 0, T = tr.getBoundingClientRect();
  /* ป้ายที่ล้นขอบซ้าย/ขวา ชิดขอบแทนการอยู่กึ่งกลางหมุด */
  [].slice.call(tr.querySelectorAll('.drop')).forEach(function (d) { var r = d.querySelector('.chip').getBoundingClientRect(); if (r.left < T.left) d.classList.add('edge'); else if (r.right > T.right) d.classList.add('edgeR'); });
  var D2 = [].slice.call(tr.querySelectorAll('.drop')).map(function (d) { var c = d.querySelector('.chip'); return { d: d, c: c, r: c.getBoundingClientRect() }; }).sort(function (a, b) { return a.r.left - b.r.left; });
  D2.forEach(function (o) { var k = 0; while (lv[k] !== undefined && o.r.left < lv[k] + 8) k++; o.d.style.setProperty('--lv', k); lv[k] = o.r.right; if (k > mx) mx = k; });
  tr.style.height = (80 + mx * 32) + 'px';
  if (nowb) {
    var hit = function () { var b = nowb.getBoundingClientRect(); return D2.some(function (o) { var r = o.c.getBoundingClientRect(); return r.left < b.right + 4 && r.right > b.left - 4 && r.top < b.bottom + 2 && r.bottom > b.top - 2; }); };
    if (hit()) { $('#nowm').classList.add('flip'); if (hit()) tr.style.height = (parseInt(tr.style.height, 10) + 24) + 'px'; }
  }
}
function sendView() { return SEND && !(SENDMSG && !SENDOK); }
function drawAsk() {
  var C = D.catchup.days, L = catchDays(), n = 0;
  L.forEach(function (a) { n += a.day === D.day ? picked().length : a.n; });
  var prev = C.filter(function (a) { return a.day !== D.today; });
  $('#app').innerHTML = '<div class="sendp"><div class="pp">' + brand('ส่ง Playlist เข้า PHOTO CHECK STUDIO') +
    '<div><h1>' + (prev.length && prev.every(function (a) { return a.redo && a.redo.length; }) ? 'มีฉบับแก้ไขของวันที่ดึงไปแล้ว ' + prev.length + ' วัน' : 'มี Playlist ที่ยังไม่ได้ดึง ' + prev.length + ' วัน') + '</h1><p class="when">ส่งพร้อมวันนี้ได้เลย ระบบเรียงวันและเติมต่อกันทีละวันให้เอง • ติ๊กออกได้ถ้าไม่ต้องการวันไหน</p></div>' +
    '<div class="dayrow">' + C.map(function (a) { return dayCard(a, { pick: 1 }); }).join('') + '</div>' +
    '<div class="choice"><button class="btn pri" data-a="sendsel"' + (L.length ? '' : ' disabled') + '>' + sv(P.check, 17, 0, 2.6) + (L.length ? 'ส่ง ' + L.length + ' วัน · ' + n + ' ไฟล์' : 'ยังไม่ได้เลือกวัน') + '</button>' +
    '<button class="btn" data-a="sendtoday">เฉพาะวันนี้</button></div>' +
    '<div class="ppf"><span>วันที่ไม่ได้ส่งจะไม่ถามอีก ดูย้อนหลังได้ที่ PLAYLIST INBOX</span></div></div></div>';
}
function drawSend() {
  if (SENDASK && D) return drawAsk();
  var multi = SENDDAYS.length > 1;
  var Pk = D ? picked() : [], n = multi ? ZTOT : Pk.length, nx = D && !multi ? Pk.filter(isExtra).length : 0, sz = 0, g = D ? byKind() : null;
  Pk.forEach(function (f) { sz += f.size; });
  var miss = D ? mainKinds().filter(function (k) { return !g.by[k.name]; }).map(function (k) { return k.name; }) : [];
  var zipping = BUSY === 'zip', pct = !D ? 8 : SENDOK ? 100 : zipping ? 15 + Math.round(75 * PROG / Math.max(1, n)) : 12;
  var title = !D ? 'กำลังเตรียม Playlist ของวันนี้…' : SENDOK ? 'ส่ง ' + SENTN + ' ไฟล์เข้าระบบแล้ว' : 'กำลังส่ง ' + n + ' ไฟล์' + (multi ? ' (' + SENDDAYS.length + ' วัน)' : '') + 'เข้าระบบ';
  if (multi) miss = [];
  var step = function (txt, s) { return '<li class="' + s + '"><span class="i">' + (s === 'done' ? sv(P.check, 13, 'var(--blueInk)', 3.2) : '') + '</span>' + txt + '</li>'; };
  var h = '<div class="sendp"><div class="pp">' + brand('ส่ง Playlist เข้า PHOTO CHECK STUDIO') +
    '<div><h1>' + esc(title) + '</h1><p class="when">Playlist ' + (multi ? esc(dayName(SENDDAYS[0])) + ' – ' + esc(dayName(SENDDAYS[SENDDAYS.length - 1])) : D ? 'วัน' + esc(thLong(D.day)) + (n ? ' — ต้องมาทุกวัน ' + (n - nx) + ' ไฟล์' + (nx ? ' + ไม่ได้ส่งทุกวัน ' + nx + ' ไฟล์' : '') : '') : 'ของวันนี้') + '</p><div class="prog" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="' + pct + '"><span style="width:' + pct + '%"></span></div></div>' +
    '<ol class="steps">' + step(multi ? 'หาไฟล์ ' + SENDDAYS.length + ' วันใน Drive' : 'หาไฟล์ของวันนี้ใน Drive', D ? 'done' : 'now') +
    step('รวมเป็นไฟล์ .zip' + (n ? ' (' + kb(sz) + ')' : '') + (zipping ? ' ' + PROG + '/' + n : ''), SENDOK ? 'done' : zipping ? 'now' : '') +
    step('ส่งเข้า PHOTO CHECK STUDIO', SENDOK ? 'done' : '') + '</ol>';
  if (miss.length) h += '<div class="banner">' + sv(P.clock, 18, 0, 2.4) + '<span>ยังขาด ' + andList(miss.map(function (n) { return '<b>' + esc(n) + '</b>'; })) + ' ส่งที่มีไปก่อน มาครบแล้วกดดึงอีกครั้ง</span></div>';
  if (D && (D.revs || []).length && !multi) h += '<div class="banner info">' + sv(P.rev, 18, 0, 2.4) + '<span>มีฉบับแก้ไข: ' + esc(D.revs.map(function (r) { return r.kind; }).join(', ')) + ' — ส่งฉบับใหม่ให้แล้ว</span></div>';
  h += '<div class="ppf"><span>' + (SENDOK ? 'หน้านี้จะปิดเอง ถ้าไม่ปิด ปิดเองได้เลย' : 'หน้านี้ปิดเองเมื่อส่งเสร็จ') + '</span>' + (D && !SENDOK ? '<button class="lk" data-a="zip"' + (BUSY ? ' disabled' : '') + '>ดาวน์โหลด .zip แทน</button>' : '') + '</div></div></div>';
  $('#app').innerHTML = h;
}
function draw() {
  if (sendView()) return drawSend();
  var app = $('#app');
  if (!D) { app.innerHTML = '<p class="loading">กำลังโหลด…</p>'; return; }
  var g = byKind(), day = D.day, isToday = day === D.today, future = day > D.today, maxDay = D.maxDay || D.today, ctx = { isToday: isToday, future: future };
  var last = D.files.length ? D.files[D.files.length - 1] : null;
  var h = '<header class="top"><div class="in">' + brand('ส่วนหนึ่งของ PHOTO CHECK STUDIO') + '<div class="tools"><div class="date">' +
    '<button class="nav" data-go="' + shift(day, -1) + '" aria-label="วันก่อนหน้า" title="วันก่อนหน้า">' + sv(P.left, 17, 0, 2.2) + '</button>' +
    '<label class="cur" title="เลือกวันที่">' + sv(P.cal, 16) + esc(thLong(day)) + (isToday ? '<span class="tag">วันนี้</span>' : future ? '<span class="tag">ล่วงหน้า</span>' : '') +
    '<input type="date" id="pick" value="' + day + '" max="' + maxDay + '" aria-label="เลือกวันที่"></label>' +
    '<button class="nav" data-go="' + shift(day, 1) + '" aria-label="วันถัดไป" title="วันถัดไป"' + (day >= maxDay ? ' disabled' : '') + '>' + sv(P.right, 17, 0, 2.2) + '</button></div>' +
    (isToday ? '' : '<button class="btn" data-go="' + D.today + '">กลับไปวันนี้</button>') +
    '<a class="btn dark" href="' + esc(D.studio) + '" target="_blank" rel="noopener">เปิด PHOTO CHECK STUDIO' + sv(P.ext, 15, 0, 2.2) + '</a></div></div></header>';
  h += '<main' + (BUSY === 'load' ? ' class="dim"' : '') + '><div class="in">';
  if (SEND && SENDMSG) h += '<div class="banner' + (SENDINFO ? ' info' : '') + '" role="' + (SENDINFO ? 'status' : 'alert') + '">' + sv(P.info, 18, 0, 2.4) + '<span>' + esc(SENDMSG) + '</span></div>';
  h += boardHtml(g, isToday, future);
  h += revHtml(isToday);
  if (isToday && D.catchup) h += catchHtml();
  if (isToday && D.ahead && D.ahead.length) h += aheadHtml();
  h += rackHtml(g, ctx);
  h += moreHtml(g);
  h += '<div class="foot"><span>' + (last ? 'ไฟล์ล่าสุดเข้าเมื่อ ' + esc(last.time) + ' · ' : '') + 'เก็บไว้ใน Google Drive โฟลเดอร์ Playlist รายวัน</span><span>ลากไฟล์ .zip ลงขั้นที่ 1 ของ PHOTO CHECK STUDIO ได้เลย ไม่ต้องแตกไฟล์</span></div>';
  h += '</div></main><div class="dock" id="dock" aria-label="ถาดรับ: ไฟล์ที่เลือกไว้"></div>';
  app.innerHTML = h;
  drawDock(false);
  levelDrops();
}
function drawBar(bump) {
  if (sendView()) return drawSend();
  drawDock(bump);
  var z = $('#zipall'); if (z) { z.innerHTML = sv(P.down, 18, 0, 2.2) + esc(catchLabel()); z.disabled = !catchDays().length || !!BUSY; }
}
var RT = 0; addEventListener('resize', function () { clearTimeout(RT); RT = setTimeout(function () { if (D && !sendView()) { fitChips(); levelDrops(); } }, 120); });
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
/* รวมไฟล์เป็น .zip: E = [{day, id, name, path}] (path = ชื่อใน .zip) • cb = ส่งเข้าระบบแทนดาวน์โหลด • after = หลังดาวน์โหลดเสร็จ */
function zipRun(E, zname, key, cb, after) {
  var out = [], next = 0, done = 0, failed = false;
  if (!E.length || BUSY) return;
  BUSY = cb ? 'zip' : key; PROG = 0; ZTOT = E.length; draw();
  function prog() { if (BUSY === 'zipall') { var b = $('#zipall'); if (b) b.innerHTML = sv(P.down, 18, 0, 2.2) + esc(catchLabel()); } else drawBar(); }
  function fail(e) { if (failed) return; failed = true; BUSY = ''; if (typeof cb === 'function') SENDMSG = 'รวมไฟล์ไม่สำเร็จ: ' + ((e && e.message) || e) + ' — กด “ดาวน์โหลด” อีกครั้ง หรือใช้ “เปิดใน Drive”'; draw(); toast('ดาวน์โหลดไม่สำเร็จ: ' + ((e && e.message) || e) + ' — ลองอีกครั้ง หรือใช้ปุ่ม “เปิดใน Drive”', 1); }
  function one(i, part, acc) {
    google.script.run.withFailureHandler(fail).withSuccessHandler(function (r) {
      if (failed) return;
      acc.push(b64u8(r.b64));
      if (part + 1 < r.n) return one(i, part + 1, acc);
      out[i] = { name: E[i].path, data: join(acc) }; done++; PROG = done; prog();
      if (next < E.length) one(next++, 0, []);
      else if (done === E.length) finish();
    }).webPart(E[i].day, E[i].id, part);
  }
  function finish() {
    var used = {};
    out.forEach(function (e) {
      var n = e.name, d = n.lastIndexOf('.'), b = d > 0 ? n.slice(0, d) : n, x = d > 0 ? n.slice(d) : '', k = 2;
      while (used[n.toLowerCase()]) n = b + ' (' + k++ + ')' + x;
      used[n.toLowerCase()] = 1; e.name = n;
    });
    if (typeof cb === 'function') { BUSY = ''; draw(); return cb(zname, zipBlob(out), out.length); } // ส่งเข้าระบบแทนดาวน์โหลด
    var url = URL.createObjectURL(zipBlob(out));
    var a = document.createElement('a'); a.href = url; a.download = zname; a.target = '_self';
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 60000);
    BUSY = '';
    if (after) after();
    draw();
    toast('ดาวน์โหลด ' + zname + ' (' + out.length + ' ไฟล์) แล้ว — ลากไฟล์นี้ลงขั้นที่ 1 ของ PHOTO CHECK STUDIO ได้เลย ไม่ต้องแตกไฟล์');
  }
  for (var k = 0; k < Math.min(3, E.length); k++) one(next++, 0, []);
}
function zip(cb) {
  var Pk = picked();
  zipRun(Pk.map(function (f) { return { day: D.day, id: f.id, name: f.name, path: f.name }; }), 'Playlist_' + D.day + '.zip', 'zip', cb, function () { mark([D.day]); if (D.revs) D.revs.forEach(function (r) { r.after = false; }); });
}
/* หลายวันใน .zip เดียว แยกโฟลเดอร์ตามวัน — วันที่กำลังดูใช้ไฟล์ที่ติ๊กบนหน้า วันอื่นใช้ไฟล์ที่ติ๊กไว้ให้ (แบบละ 1 ไฟล์) • ไฟล์เดียวกันหลายวัน (เช่น Show DC ฉบับเดิม) ใส่ครั้งเดียว */
function zipDays(cb) {
  var L = catchDays(), seen = {}, E = [];
  L.forEach(function (a) {
    (a.day === D.day ? picked() : a.files).forEach(function (f) {
      var k = f.name + '|' + f.size; if (seen[k]) return; seen[k] = 1;
      E.push({ day: a.day, id: f.id, name: f.name, path: 'Playlist_' + a.day + '/' + f.name });
    });
  });
  if (!L.length) return;
  var name = L.length > 1 ? 'Playlist_' + L[0].day + '_ถึง_' + L[L.length - 1].day + '.zip' : 'Playlist_' + L[0].day + '.zip';
  zipRun(E, name, 'zipall', cb, function () {
    mark(L.map(function (a) { return a.day; }));
    mark(D.catchup.days.filter(function (a) { return !CSEL[a.day]; }).map(function (a) { return a.day; }), 's');
    D.catchup = null;
  });
}
function collect() {
  BUSY = 'collect'; draw();
  call('webCollect', [], function (r) { D = r.data; BUSY = ''; defaults(); draw(); toast(r.msg); });
}
document.addEventListener('click', function (e) {
  var b = e.target.closest('[data-go],[data-a],[data-un]');
  if (!b || b.disabled || BUSY) return;
  if (b.dataset.un) { SEL[b.dataset.un] = false; var inp = document.querySelector('input[data-id="' + b.dataset.un + '"]'); if (inp) inp.checked = false; drawBar(); return; }
  if (b.dataset.go) { if (b.dataset.go <= (D.maxDay || D.today)) load(b.dataset.go); }
  else if (b.dataset.a === 'zip') zip();
  else if (b.dataset.a === 'zipall') zipDays();
  else if (b.dataset.a === 'tray') { TRAY = !TRAY; drawBar(); }
  else if (b.dataset.a === 'reset') { defaults(); draw(); }
  else if (b.dataset.a === 'clear') { SEL = {}; draw(); }
  else if (b.dataset.a === 'cskip') { mark(D.catchup.days.map(function (a) { return a.day; }).filter(function (d) { return d < D.today; }), 's'); D.catchup = null; draw(); toast('ไม่แสดงกล่องนี้แล้ว — ดู Playlist ย้อนหลังรายวันได้จากลูกศรด้านบน'); }
  else if (b.dataset.a === 'sendsel') { var sel = catchDays().map(function (a) { return a.day; }); mark(D.catchup.days.filter(function (a) { return !CSEL[a.day]; }).map(function (a) { return a.day; }), 's'); sendStudio(sel.length === 1 && sel[0] === D.day ? [D.day] : sel); }
  else if (b.dataset.a === 'sendtoday') { mark(D.catchup.days.map(function (a) { return a.day; }).filter(function (d) { return d < D.today; }), 's'); sendStudio([D.day]); }
  else if (b.dataset.a === 'collect') collect();
  else if (b.dataset.a === 'sel') {
    var g = byKind(), L = g.oth, all = L.every(function (f) { return SEL[f.id]; });
    L.forEach(function (f) { SEL[f.id] = !all; }); draw();
  }
});
document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && TRAY) { TRAY = false; drawBar(); } });
document.addEventListener('change', function (e) {
  var t = e.target;
  if (t.id === 'pick') { if (t.value && t.value <= (D.maxDay || D.today)) load(t.value); else t.value = D.day; return; }
  if (t.dataset.cd) { CSEL[t.dataset.cd] = t.checked; if (SENDASK) draw(); else drawBar(); return; }
  if (t.dataset.id) { SEL[t.dataset.id] = t.checked; if (t.checked) NEWID = t.dataset.id; drawBar(t.checked); }
});
load('');
</script></body></html>`;
}
