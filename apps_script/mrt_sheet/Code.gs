/**
 * ใบเช็ครูป MRT บน Google Sheets (ทดลอง) — ส่วนหนึ่งของ PHOTO CHECK STUDIO
 *
 * - เมนู "ใบเช็ครูป" ในชีต: นำเข้าใบเช็ค Excel (.xlsm) จากโฟลเดอร์ใน Drive → แท็บละสถานี
 * - คนคัดรูปพิมพ์เลขวันที่ถ่ายในช่องแบบใบเช็ค (ระบบจดว่าใครใส่ เมื่อไร ไว้ในโน้ตของช่อง)
 * - หน้าเว็บมือถือสำหรับคนถ่ายรูป (doGet) อ่านจากชีตนี้ ดูได้อย่างเดียว
 * - หน้ารับใบเช็คจาก PHOTO CHECK STUDIO (doGet ?recv=studio) — ปุ่ม “ส่งเข้าใบเช็ค MRT” ในหน้าผลการเติม ส่งไฟล์มาให้ เก็บลง Drive แล้วนำเข้าเหมือนเมนู
 * ไม่มีการใช้ AI ในการทำงาน — เป็นสูตรและสคริปต์ธรรมดา
 */
const CFG = {
  importFolderName: 'ใบเช็ค MRT นำเข้า',   // โฟลเดอร์ใน Drive ที่วางไฟล์ใบเช็ค Excel
  usersSheet: 'ผู้ใช้และสิทธิ์',
  summarySheet: 'สรุปคะแนน',
  firstRow: 6,          // แถวแรกของรายการหนังในแท็บสถานี
  firstCol: 3,          // คอลัมน์จอแรก (C)
  tz: 'Asia/Bangkok',
  none: 'ไม่ฉาย',
  studioOrigin: 'https://playlist-system.github.io',   // เว็บ PHOTO CHECK STUDIO ที่ส่งใบเช็คมาได้
};
const TH_MON = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
const COLORS = { navy: '#022866', cyan: '#1BB8EF', head: '#EEF2F8', ok: '#E6F5EC', todo: '#FFF1C9', none: '#E6E8EC',
  band: '#E2F4FC', bandText: '#0072A8', exit: '#B9312A', g: '#2FA866', a: '#E5A321', r: '#DC4F43' };

// ===================== เมนู =====================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('ใบเช็ครูป')
    .addItem('นำเข้าใบเช็ค Excel จาก Drive', 'importFromDrive')
    .addItem('ตั้งสิทธิ์ตามแท็บผู้ใช้', 'applyPermissions')
    .addItem('อัปเดตสีแท็บ', 'refreshTabColors')
    .addSeparator()
    .addItem('ลิงก์หน้ามือถือคนถ่ายรูป', 'showPhoneLink')
    .addItem('อัปเดตข้อมูลหน้ามือถือ', 'refreshPhone')
    .addToUi();
}

function showPhoneLink() {
  const url = ScriptApp.getService().getUrl();
  const ui = SpreadsheetApp.getUi();
  ui.alert('หน้ามือถือคนถ่ายรูป', url ? url + '\n\nส่งลิงก์นี้ใน LINE ได้เลย (เปิดได้เฉพาะอีเมลบริษัท)' :
    'ยังไม่ได้เผยแพร่เป็นเว็บแอป — ทำตามคู่มือข้อ "เผยแพร่หน้ามือถือ" ก่อน', ui.ButtonSet.OK);
}

// ===================== นำเข้า =====================
function importFromDrive() {
  const ss = SpreadsheetApp.getActive(), ui = SpreadsheetApp.getUi();
  // ขึ้นรอบใหม่: ถามก่อน → เก็บสำเนารอบเก่าไว้ แล้วเริ่มใหม่ (ไม่เอาวันที่ที่ใส่ในรอบก่อนมาปน)
  const ask = (label, prev) => {
    const a = ui.alert('ขึ้นรอบใหม่?', `ใบเช็คที่นำเข้าเป็นของ ${label}\nชีตตอนนี้เป็น ${prev}\n\nกด "ใช่" = เริ่ม ${label} (เก็บสำเนา ${prev} ไว้ใน Drive แล้วเริ่มใหม่)\nกด "ไม่" = นำเข้าใน ${prev} ต่อ`, ui.ButtonSet.YES_NO);
    return a === ui.Button.YES ? 'new' : a === ui.Button.NO ? 'keep' : 'cancel';
  };
  const r = withLock_(() => importCore_(ss, ask));
  if (r.none) { ui.alert('ยังไม่มีไฟล์', `วางไฟล์ใบเช็ค MRT (.xlsm) ในโฟลเดอร์ "${CFG.importFolderName}" ใน Google Drive แล้วกดใหม่อีกครั้ง\n\n${r.folderUrl}`, ui.ButtonSet.OK); return; }
  if (r.cancel) { ss.toast('ยังไม่ได้นำเข้า', 'ใบเช็ครูป', 5); return; }
  ss.toast(`นำเข้า ${r.files} ไฟล์ ${r.stations} สถานี เรียบร้อย${r.fresh ? ' (เริ่มรอบใหม่ — เก็บสำเนารอบก่อนไว้ใน Drive แล้ว)' : ` • หนังใหม่ ${r.newFilms} เรื่อง`}`, 'ใบเช็ครูป', 10);
}

/** กันนำเข้าซ้อนกัน (เมนู กับ ปุ่มส่งจาก PHOTO CHECK STUDIO พร้อมกัน) */
function withLock_(fn) {
  const L = LockService.getScriptLock();
  if (!L.tryLock(30000)) throw new Error('มีการนำเข้าอื่นกำลังทำอยู่ — รอสักครู่แล้วลองใหม่');
  try { return fn(); } finally { L.releaseLock(); }
}

/** นำเข้าใบเช็คทุกไฟล์ในโฟลเดอร์นำเข้า (คนละไฟล์ล่าสุด)
 *  decide(label, prev) ตอนต้องขึ้นรอบใหม่: 'new' = เริ่มรอบใหม่ / 'keep' = นำเข้าในรอบเดิม / 'cancel' = ไม่นำเข้า / undefined = คืน {ask} ไปถาม (ยังไม่แก้ชีต)
 *  คืน {none, folderUrl} | {ask, label, prev} | {cancel} | {ok, files, stations, fresh, label, prev, newFilms} */
function importCore_(ss, decide) {
  const folder = importFolder_();
  const it = folder.getFiles(), files = [];
  while (it.hasNext()) { const f = it.next(); if (/\.xls[xm]$/i.test(f.getName()) && !f.isTrashed()) files.push(f); }
  if (!files.length) return { none: true, folderUrl: folder.getUrl() };
  // คนละ 1 ไฟล์: ถ้ามีหลายไฟล์ของคนเดียวกัน ใช้ไฟล์ที่แก้ล่าสุด
  const latest = {};
  for (const f of files) { const k = (f.getName().match(/คุณ\s*([^\s.]+)/) || [, f.getName()])[1]; if (!latest[k] || f.getLastUpdated() > latest[k].getLastUpdated()) latest[k] = f; }
  const use = Object.keys(latest).map(k => latest[k]);
  const books = [];
  for (const f of use) {
    const meta = { name: '_นำเข้า_' + f.getName().replace(/\.\w+$/, ''), mimeType: MimeType.GOOGLE_SHEETS, parents: [folder.getId()] };
    let conv;
    try { conv = Drive.Files.create(meta, f.getBlob()); }
    catch (err) { conv = Drive.Files.create(meta, f.getBlob().setContentType(MimeType.MICROSOFT_EXCEL)); }  // บางบัญชีแปลง .xlsm ตรง ๆ ไม่ได้
    try {
      const tmp = SpreadsheetApp.openById(conv.id);
      const sheets = tmp.getSheets().map(sh => ({ name: sh.getName(), values: sh.getDataRange().getValues() }));
      books.push(parseBook_(f.getName(), sheets));
    } finally { DriveApp.getFileById(conv.id).setTrashed(true); }
  }
  const stations = [].concat(...books.map(b => b.stations));
  let round = roundOf_(stations, new Date());
  const props = PropertiesService.getScriptProperties(), prev = props.getProperty('ROUND');
  let fresh = !!prev && prev !== round.label;
  if (fresh) {
    const a = decide ? decide(round.label, prev) : undefined;
    if (a === undefined) return { ask: true, label: round.label, prev };
    if (a === 'cancel') return { cancel: true };
    if (a !== 'new') { fresh = false; round = Object.assign({}, round, { label: prev }); }
  }
  // หนังใหม่ = ชื่อที่ยังไม่มีในแท็บสถานีเดิม (บอกในผลนำเข้า)
  let newFilms = 0;
  if (!fresh) for (const st of stations) {
    const sh = ss.getSheetByName(st.station), n = sh ? sh.getLastRow() - CFG.firstRow + 1 : 0;
    const have = new Set(n > 0 ? sh.getRange(CFG.firstRow, 2, n, 1).getValues().map(r => String(r[0]).trim()).filter(String) : []);
    newFilms += st.films.filter(f => !have.has(f.name)).length;
  }
  if (fresh) DriveApp.getFileById(ss.getId()).makeCopy(`ใบเช็ครูป MRT ${prev} (เก็บไว้)`);
  for (const st of stations) writeStation_(ss, st, round, fresh);
  props.setProperty('ROUND', round.label);
  writeUsers_(ss, stations);
  writeSummary_(ss, stations, round);
  refreshTabColors(ss);
  props.setProperty('SHEET_ID', ss.getId());
  try { snapBuild_(ss); } catch (err) { /* เตรียมไม่ได้ → หน้ามือถือเตรียมเองตอนเปิดครั้งแรก */ }   // เตรียมข้อมูลหน้ามือถือใหม่ (เปิดลิงก์แล้วไม่ต้องอ่านทุกแท็บ)
  return { ok: true, files: use.length, stations: stations.length, fresh, label: round.label, prev: prev || '', newFilms };
}

function importFolder_() {
  const it = DriveApp.getFoldersByName(CFG.importFolderName);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CFG.importFolderName);
}

/** อ่านใบเช็ค MRT 1 ไฟล์ (ค่าจากทุกแท็บ) → [{station, person, screens, films}] */
function parseBook_(fileName, sheets) {
  const pm = String(fileName).match(/คุณ\s*([^\s.]+)/);
  const person = pm ? pm[1] : String(fileName).replace(/\.\w+$/, '');
  const stations = [];
  for (const sh of sheets) {
    const v = sh.values;
    if (!v || v.length < 17 || /ผลวิเคราะห์|สรุป/.test(sh.name)) continue;
    const r14 = v[13] || [], screens = [];
    for (let c = 4; c < r14.length && String(r14[c] || '').trim(); c++) screens.push(String(r14[c]).trim());
    if (!screens.length) continue;
    const films = []; let cur = null;
    for (let r = 15; r < v.length; r++) {   // แถว 16 = วันที่รอบแรก, หนังเริ่มแถว 17
      const row = v[r], cv = row[2], nm = String(row[3] == null ? '' : row[3]).trim();
      if (cv instanceof Date) cur = cv;
      if (!nm) continue;
      let exit = null;
      if (typeof cv === 'string' && /ออก/.test(cv)) { const m = cv.match(/\d+/); exit = m ? +m[0] : null; }
      let day = 0;
      const marks = screens.map((_, i) => {   // ค่าในช่องแบบใบเช็ค: เลขวันที่ถ่าย / 'ไม่ฉาย' / ''
        let x = row[4 + i];
        if (x instanceof Date) x = +Utilities.formatDate(x, CFG.tz, 'd');
        if (typeof x === 'number') { if (x >= 1 && x <= 31) day = Math.max(day, Math.floor(x)); return x; }
        if (typeof x === 'string' && /^(none|ไม่ฉาย)$/i.test(x.trim())) return CFG.none;
        return '';
      });
      films.push({ name: nm, add: cur, exit, marks, day });
    }
    stations.push({ station: sh.name.trim(), person, screens, films });
  }
  return { person, stations };
}

// ใบเช็ค MRT เป็นไฟล์ทั้งเดือน: รอบ 16–31 ยังมีแถบ "ขึ้น 1 …" อยู่บนสุด จึงใช้แถบแรกบอกได้แค่เดือน
// รอบ: มีวันที่ถ่าย 16 ขึ้นไป = รอบ 16–31 • ไม่มี → ตามวันที่นำเข้า (เดือนของไฟล์ผ่านไปแล้ว = รอบ 16–31)
function roundOf_(stations, today) {
  today = today || new Date();
  let min = null, shot = 0;
  for (const s of stations) for (const f of s.films) { if (f.add && (!min || f.add < min)) min = f.add; if (f.day > shot) shot = f.day; }
  const ymd = (d) => [+Utilities.formatDate(d, CFG.tz, 'yyyy'), +Utilities.formatDate(d, CFG.tz, 'M'), +Utilities.formatDate(d, CFG.tz, 'd')];
  const [y, m] = ymd(min || today), [ty, tm, td] = ymd(today);
  const half = shot >= 16 ? 2 : ty * 12 + tm > y * 12 + m ? 2 : ty * 12 + tm < y * 12 + m ? 1 : td <= 15 ? 1 : 2;
  const last = new Date(y, m, 0).getDate();
  return { y, m, half, label: `รอบ ${half === 1 ? '1–15' : '16–' + last} ${TH_MON[m - 1]} ${(y + 543) % 100}` };
}
const thDay_ = (d) => `${+Utilities.formatDate(d, CFG.tz, 'd')} ${TH_MON[+Utilities.formatDate(d, CFG.tz, 'M') - 1]} ${(+Utilities.formatDate(d, CFG.tz, 'yyyy') + 543) % 100}`;

/** ตารางของแท็บสถานี (ไม่แตะ Google API — ทดสอบแยกได้) */
function stationGrid_(st, round, old) {
  const nS = st.screens.length, rows = [], kinds = [];
  let lastAdd = null, first = true;
  const key = (f, code) => f + '\u0001' + code;
  for (const f of st.films) {
    const ak = f.add ? f.add.getTime() : 0;
    if (first || ak !== lastAdd) {
      const label = f.add ? (first ? 'ขึ้น ' : 'ขึ้นใหม่ ') + thDay_(f.add) : 'รายการหนัง';
      rows.push([label, ''].concat(Array(nS).fill(''))); kinds.push('band');
      lastAdd = ak; first = false;
    }
    const cells = f.marks.map((m, i) => {   // ค่าที่คนคัดรูปพิมพ์ในชีตชนะ ถ้าไม่มี ใช้ค่าจากใบเช็ค Excel
      const o = old && old[key(f.name, st.screens[i])];
      if (typeof o === 'number') return o;
      if (isNone_(o)) return CFG.none;
      return m;
    });
    rows.push([f.exit ? `ออก ${f.exit} ${TH_MON[round.m - 1]}` : '', f.name].concat(cells)); kinds.push('film');
  }
  return { rows, kinds };
}

const isNone_ = (x) => typeof x === 'string' && /^(none|ไม่ฉาย)$/i.test(x.trim());
const isDone_ = (x) => typeof x === 'number' || x instanceof Date || x === true || isNone_(x);
function colA1_(c) { let s = ''; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = (c - m - 1) / 26; } return s; }

function writeStation_(ss, st, round, fresh) {
  let sh = ss.getSheetByName(st.station);
  const nS = st.screens.length, fr = CFG.firstRow, fc = CFG.firstCol;
  // เก็บค่าและโน้ตเดิม (วันที่ที่พิมพ์ในชีตไม่หาย เมื่อนำเข้าใหม่)
  const old = {}, oldNotes = {};
  if (sh && !fresh && sh.getLastRow() >= fr) {
    const hdr = sh.getRange(2, fc, 1, sh.getLastColumn() - fc + 1).getValues()[0].map(h => String(h).split('\n')[0]);
    const n = sh.getLastRow() - fr + 1, w = hdr.length;
    const vals = sh.getRange(fr, 2, n, w + 1).getValues(), notes = sh.getRange(fr, fc, n, w).getNotes();
    vals.forEach((r, i) => { if (!r[0]) return; hdr.forEach((h, j) => {
      const k = r[0] + '\u0001' + h, x = r[j + 1];
      // ช่องที่คนคัดรูปพิมพ์ในชีต (มีโน้ต "ใส่ … เมื่อ …") ชนะ • ช่องที่มาจาก Excel รอบก่อน ใช้ค่าจาก Excel ใหม่
      if (!notes[i][j]) return;
      if (typeof x === 'number' || isNone_(x)) old[k] = x;
      else if (x === true) { const m = String(notes[i][j]).match(/(\d{1,2})\/\d{1,2}\/\d{4}/); if (m) old[k] = +m[1]; }  // ติ๊กจากฉบับช่องติ๊ก → วันที่ในโน้ต
      if (notes[i][j]) oldNotes[k] = notes[i][j]; }); });
    // จับคู่ด้วยเลขจอ
    st.screens.forEach(code => { const num = code.split(/\s+/)[0]; for (const k in old) if (k.endsWith('\u0001' + num)) old[k.replace(/\u0001.*$/, '\u0001' + code)] = old[k]; });
    for (const k in oldNotes) { const num = k.split('\u0001')[1]; const code = st.screens.find(c => c.split(/\s+/)[0] === num); if (code) oldNotes[k.split('\u0001')[0] + '\u0001' + code] = oldNotes[k]; }
  }
  if (!sh) sh = ss.insertSheet(st.station); else { sh.clear(); sh.clearConditionalFormatRules(); sh.getDataRange().clearDataValidations(); sh.getDataRange().clearNote(); sh.getDataRange().breakApart(); }
  const g = stationGrid_(st, round, old);
  const W = 2 + nS, lastCol = colA1_(1 + W), nR = g.rows.length, lastRow = fr + Math.max(nR, 1) - 1;
  const need = Math.max(lastRow + 20, 100);
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  if (sh.getMaxColumns() < W + 2) sh.insertColumnsAfter(sh.getMaxColumns(), W + 2 - sh.getMaxColumns());
  // หัว
  sh.getRange(1, 1, 1, 2).merge().setValue('สถานี' + st.station);
  const top = sh.getRange(1, 3, 1, nS);
  if (nS > 1) top.merge();
  sh.getRange(1, 3).setFormula(stationScoreFormula_(fc, fc + nS - 1, fr));
  sh.getRange(1, 1, 1, W).setBackground(COLORS.navy).setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(13).setVerticalAlignment('middle');
  sh.getRange(1, 3).setHorizontalAlignment('right');
  const head = [['ขึ้น / ออก', 'หนังโฆษณา'].concat(st.screens.map(c => { const p = c.split(/\s+/); return p[0] + '\n' + p[p.length - 1]; }))];
  sh.getRange(2, 1, 1, W).setValues(head).setBackground(COLORS.head).setFontWeight('bold').setFontColor(COLORS.navy).setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(2, 1, 1, 2).setHorizontalAlignment('left');
  sh.getRange(3, 1, 3, 2).setValues([['คะแนนจอ', `คุณ${st.person}  ${round.label}`], ['ได้รูป / ทั้งหมด', ''], ['ยังไม่ได้รูป', '']]).setFontColor('#5A6987').setFontWeight('bold');
  const f3 = [], f4 = [], f5 = [];
  for (let i = 0; i < nS; i++) { const L = colA1_(fc + i), rng = `${L}$${fr}:${L}`, done = doneF_(rng), tot = `COUNTA($B$${fr}:$B)`;
    f3.push(`=IFERROR(${done}/${tot},0)`); f4.push(`=${done}&"/"&${tot}`); f5.push(`="เหลือ "&(${tot}-${done})`); }
  sh.getRange(3, fc, 1, nS).setFormulas([f3]).setNumberFormat('0%').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(4, fc, 1, nS).setFormulas([f4]).setHorizontalAlignment('center').setFontColor('#5A6987');
  sh.getRange(5, fc, 1, nS).setFormulas([f5]).setHorizontalAlignment('center').setFontColor('#5A6987');
  sh.getRange(5, 1, 1, W).setBorder(null, null, true, null, null, null, '#9AA3B2', SpreadsheetApp.BorderStyle.SOLID_THICK);
  // รายการหนัง
  if (nR) {
    const body = sh.getRange(fr, 1, nR, W);
    body.setValues(g.rows).setVerticalAlignment('middle');
    const ticks = sh.getRange(fr, fc, nR, nS), C1 = `${colA1_(fc)}${fr}`;
    // ช่องแบบใบเช็ค: พิมพ์เลขวันที่ถ่าย 1–31 หรือ ไม่ฉาย (พิมพ์อย่างอื่นขึ้นเตือน ไม่บล็อก)
    ticks.setDataValidation(SpreadsheetApp.newDataValidation().requireFormulaSatisfied(`=OR(${C1}="",AND(ISNUMBER(${C1}),${C1}>=1,${C1}<=31),${C1}="${CFG.none}",LOWER(${C1})="none")`)
      .setAllowInvalid(true).setHelpText('ใส่เลขวันที่ถ่าย (1–31) หรือ ไม่ฉาย').build());
    const bands = [], exits = [];
    g.kinds.forEach((k, i) => { const r = fr + i;
      if (k === 'band') bands.push(`A${r}:${lastCol}${r}`); else if (g.rows[i][0]) exits.push('A' + r); });
    if (bands.length) sh.getRangeList(bands).setBackground(COLORS.band).setFontColor(COLORS.bandText).setFontWeight('bold');
    if (exits.length) sh.getRangeList(exits).setFontColor(COLORS.exit).setFontWeight('bold');
    ticks.setHorizontalAlignment('center').setFontWeight('bold');
    // โน้ตเดิม (ใครใส่ เมื่อไร)
    const notes = g.rows.map((r, i) => r.slice(2).map((_, j) => (g.kinds[i] === 'film' && oldNotes[r[1] + '\u0001' + st.screens[j]]) || ''));
    ticks.setNotes(notes);
  }
  // สีตามสถานะ
  const tickA1 = `${colA1_(fc)}${fr}:${colA1_(fc + nS - 1)}${lastRow}`, rules = [];
  const tk = sh.getRange(tickA1), sc = sh.getRange(3, fc, 1, nS), C0 = colA1_(fc);
  // แบบใบเช็ค: มีเลขวันที่ถ่าย = พื้นดำตัวเลขขาว • ไม่ฉาย = เทา • ยังไม่มีรูป = เหลือง
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=ISNUMBER(${C0}${fr})`).setBackground('#000000').setFontColor('#FFFFFF').setRanges([tk]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=OR(${C0}${fr}="${CFG.none}",LOWER(${C0}${fr})="none")`).setBackground(COLORS.none).setFontColor('#5A6987').setRanges([tk]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(`=AND($B${fr}<>"",${C0}${fr}="")`).setBackground(COLORS.todo).setRanges([tk]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.9).setBackground('#E1F3E8').setFontColor('#1B7646').setRanges([sc]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.5).setBackground('#FFF2CC').setFontColor('#8F5C06').setRanges([sc]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.5).setBackground(COLORS.r).setFontColor('#FFFFFF').setRanges([sc]).build());
  sh.setConditionalFormatRules(rules);
  // ขนาด / ตรึงแถว
  sh.setColumnWidth(1, 96); sh.setColumnWidth(2, 320); sh.setColumnWidths(fc, nS, 84);
  sh.setRowHeight(1, 40); sh.setRowHeight(2, 44);
  sh.setFrozenRows(5); sh.setFrozenColumns(2);
  sh.getRange(1, 1, lastRow, W).setFontFamily('IBM Plex Sans Thai');
}

/** สูตรนับช่องที่ได้รูปแล้ว: เลขวันที่ถ่าย + ไม่ฉาย (+ ติ๊กจากฉบับก่อน) เหมือนสูตรในใบเช็ค Excel */
const doneF_ = (rng) => `(COUNT(${rng})+COUNTIF(${rng},"${CFG.none}")+COUNTIF(${rng},"none")+COUNTIF(${rng},TRUE))`;
function stationScoreFormula_(c1, c2, fr) {
  const a = colA1_(c1), b = colA1_(c2), tot = `COUNTA($B$${fr}:$B)*COLUMNS(${a}2:${b}2)`;
  const done = doneF_(`${a}${fr}:${b}`);
  return `="ได้รูป "&${done}&"/"&${tot}&"   เหลือ "&(${tot}-${done})&"   คะแนน "&TEXT(IFERROR(AVERAGE(${a}3:${b}3),0),"0.0%")`;
}

function writeUsers_(ss, stations) {
  let sh = ss.getSheetByName(CFG.usersSheet);
  const keep = {};
  let admins = '';
  if (sh) {
    const v = sh.getDataRange().getValues();
    admins = String((v[1] || [])[1] || '');
    v.slice(4).forEach(r => { if (r[0]) keep[String(r[0]).trim()] = [r[2] || '', r[3] || '']; });
    sh.clear();
  } else sh = ss.insertSheet(CFG.usersSheet, 0);
  sh.getRange(1, 1).setValue('ผู้ใช้และสิทธิ์').setFontSize(14).setFontWeight('bold').setFontColor(COLORS.navy);
  sh.getRange(2, 1, 1, 2).setValues([['ผู้ดูแลระบบ / หัวหน้า (อีเมล คั่นด้วย ,)', admins]]);
  sh.getRange(3, 1).setValue('ใส่อีเมลบริษัท แล้วกดเมนู ใบเช็ครูป → ตั้งสิทธิ์ตามแท็บผู้ใช้  •  คนถ่ายรูปไม่ต้องได้สิทธิ์ชีต ใช้หน้ามือถือ').setFontColor('#5A6987');
  const rows = [['สถานี', 'ใบเช็คของ', 'อีเมลคนคัดรูป (ใส่วันที่ได้)', 'อีเมลคนถ่ายรูป (เห็นในมือถือ คั่นด้วย ,)']];
  stations.forEach(s => rows.push([s.station, 'คุณ' + s.person].concat(keep[s.station] || ['', ''])));
  sh.getRange(4, 1, rows.length, 4).setValues(rows);
  sh.getRange(4, 1, 1, 4).setBackground(COLORS.navy).setFontColor('#FFFFFF').setFontWeight('bold');
  sh.getRange(2, 2).setBackground('#FFF8E1');
  sh.getRange(5, 3, rows.length - 1, 2).setBackground('#FFF8E1');
  sh.setColumnWidth(1, 260); sh.setColumnWidth(2, 120); sh.setColumnWidth(3, 260); sh.setColumnWidth(4, 360);
  sh.setFrozenRows(4); sh.setTabColor(COLORS.navy);
}

function writeSummary_(ss, stations, round) {
  let sh = ss.getSheetByName(CFG.summarySheet);
  if (sh) sh.clear(); else sh = ss.insertSheet(CFG.summarySheet, 0);
  const rows = [['สถานี', 'จอ', 'ใบเช็คของ', 'คะแนน']];
  stations.forEach(s => s.screens.forEach((c, i) => rows.push([s.station, c, 'คุณ' + s.person, `='${s.station.replace(/'/g, "''")}'!${colA1_(CFG.firstCol + i)}3`])));
  sh.getRange(1, 1).setValue('สรุปคะแนน MRT  ' + round.label).setFontSize(14).setFontWeight('bold').setFontColor(COLORS.navy);
  sh.getRange(2, 1, 1, 2).setValues([['คะแนนรวมทั้ง MRT', `=IFERROR(AVERAGE(D5:D${3 + rows.length}),0)`]]).setFontWeight('bold');
  sh.getRange(2, 2).setNumberFormat('0.0%').setFontSize(16).setFontColor(COLORS.navy);
  sh.getRange(4, 1, rows.length, 4).setValues(rows);
  sh.getRange(4, 1, 1, 4).setBackground(COLORS.navy).setFontColor('#FFFFFF').setFontWeight('bold');
  const sc = sh.getRange(5, 4, rows.length - 1, 1).setNumberFormat('0%');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.9).setBackground('#E1F3E8').setRanges([sc]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.5).setBackground('#FFF2CC').setRanges([sc]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.5).setBackground('#FCE8E6').setRanges([sc]).build()]);
  sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 260); sh.setColumnWidth(3, 110); sh.setColumnWidth(4, 90);
  sh.setFrozenRows(4); sh.setTabColor(COLORS.cyan);
}

function stationSheets_(ss) {
  return ss.getSheets().filter(sh => sh.getName() !== CFG.usersSheet && sh.getName() !== CFG.summarySheet && String(sh.getRange(3, 1).getValue()) === 'คะแนนจอ');
}

function refreshTabColors(ss) {
  for (const sh of stationSheets_(ss && ss.getSheets ? ss : SpreadsheetApp.getActive())) tabColor_(sh);
}
function tabColor_(sh) {
  const v = sh.getRange(3, CFG.firstCol, 1, Math.max(sh.getLastColumn() - CFG.firstCol + 1, 1)).getValues()[0].filter(x => typeof x === 'number');
  const p = v.length ? v.reduce((a, b) => a + b, 0) / v.length : 0;
  sh.setTabColor(p >= 0.9 ? COLORS.g : p >= 0.5 ? COLORS.a : COLORS.r);
}

// ===================== จดว่าใครใส่ เมื่อไร =====================
function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() === CFG.usersSheet) { try { snapUsers_(sh.getParent()); } catch (err) { /* อัปเดตตอนนำเข้า / เมนู */ } return; }
  if (e.range.getRow() < CFG.firstRow || e.range.getColumn() < CFG.firstCol) return;
  if (String(sh.getRange(3, 1).getValue()) !== 'คะแนนจอ') return;
  const vals = e.range.getValues(), now = Utilities.formatDate(new Date(), CFG.tz, 'd/M/yyyy HH:mm');
  let who = '';
  try { who = Session.getActiveUser().getEmail(); } catch (err) { who = ''; }
  const notes = vals.map(r => r.map(x => x === '' || x === false || x == null ? '' : `ใส่ ${x === true ? '✓' : x} เมื่อ ${now}${who ? ' โดย ' + who : ''}`));
  e.range.setNotes(notes);
  tabColor_(sh);
  try { snapStation_(sh); } catch (err) { /* หน้ามือถืออัปเดตตอนนำเข้า / เมนู “อัปเดตข้อมูลหน้ามือถือ” */ }
}

// ===================== สิทธิ์ =====================
function applyPermissions() {
  const ss = SpreadsheetApp.getActive(), us = ss.getSheetByName(CFG.usersSheet);
  if (!us) { SpreadsheetApp.getUi().alert('ยังไม่มีแท็บผู้ใช้ — นำเข้าใบเช็คก่อน'); return; }
  const v = us.getDataRange().getValues(), split = s => String(s || '').split(/[,\s]+/).map(x => x.trim().toLowerCase()).filter(x => /@/.test(x));
  const admins = split((v[1] || [])[1]), me = Session.getEffectiveUser().getEmail();
  const map = {}; v.slice(4).forEach(r => { if (r[0]) map[String(r[0]).trim()] = split(r[2]); });
  const editors = new Set(admins);
  Object.values(map).forEach(l => l.forEach(x => editors.add(x)));
  editors.delete(me);
  if (editors.size) ss.addEditors([...editors]);
  const lock = (p, allow) => { p.addEditor(me); const cur = p.getEditors().map(u => u.getEmail().toLowerCase()); const want = new Set([me.toLowerCase(), ...allow]);
    p.removeEditors(cur.filter(x => !want.has(x))); if (allow.length) p.addEditors(allow); if (p.canDomainEdit()) p.setDomainEdit(false); };
  for (const sh of ss.getSheets()) {
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
    const isStation = String(sh.getRange(3, 1).getValue()) === 'คะแนนจอ';
    const p = sh.protect().setDescription('ล็อก: แก้ได้เฉพาะผู้ดูแล');
    if (isStation) {
      const w = sh.getLastColumn() - CFG.firstCol + 1, n = Math.max(sh.getLastRow() - CFG.firstRow + 1, 1);
      const tick = sh.getRange(CFG.firstRow, CFG.firstCol, n, w);
      p.setUnprotectedRanges([tick]);
      lock(p, admins);
      const rp = tick.protect().setDescription('ช่องวันที่ถ่ายของสถานี' + sh.getName());
      lock(rp, admins.concat(map[sh.getName()] || []));
    } else lock(p, admins);
  }
  ss.toast(`ตั้งสิทธิ์แล้ว: แชร์ให้ ${editors.size} อีเมล`, 'ใบเช็ครูป', 8);
}

// ===================== หน้ามือถือ =====================
function doGet(e) {
  if (e && e.parameter && e.parameter.recv === 'studio') return recvPage_();
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  const t = HtmlService.createTemplateFromFile('Phone');
  t.data = JSON.stringify(Object.assign(phoneData_(email), { url: ScriptApp.getService().getUrl() })).replace(/</g, '\\u003c');
  return t.evaluate().setTitle('ใบงานถ่ายรูป MRT').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ===================== ข้อมูลหน้ามือถือ (เตรียมไว้ล่วงหน้า) =====================
// เดิมเปิดลิงก์ทีไรอ่านทุกแท็บใหม่ (39 สถานี ≈ 80 ครั้ง) จึงช้า → เตรียมข้อมูลไว้ตอนนำเข้า และอัปเดตทีละสถานีตอนคนคัดรูปพิมพ์วันที่ (onEdit)
// เก็บแบบบีบอัดใน Script Properties (ทนถาวร แบ่งชิ้นละไม่เกิน 9 KB) + แคช 6 ชั่วโมงให้อ่านเร็ว • เปิดลิงก์ = อ่านข้อมูลที่เตรียมไว้อย่างเดียว
const SNAP_CH = 8500;
const unz_ = z => JSON.parse(Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(z), 'application/x-gzip')).getDataAsString('UTF-8'));
function snapLoad_() {
  const cache = CacheService.getScriptCache(), hit = cache.get('mrt-all');
  if (hit) { try { return unz_(hit); } catch (err) { /* แคชเสีย → อ่านจาก Properties */ } }
  const P = PropertiesService.getScriptProperties().getProperties(), n = +P.SNAP_N || 0;
  if (!n) return null;
  let z = '';
  for (let i = 0; i < n; i++) { const c = P['SNAP_' + i]; if (c == null) return null; z += c; }
  try { const out = unz_(z); if (z.length < 100000) cache.put('mrt-all', z, 21600); return out; } catch (err) { return null; }
}
function snapSave_(out) {
  out.at = Date.now();
  const z = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(out), 'application/json')).getBytes());
  const props = PropertiesService.getScriptProperties(), old = +props.getProperty('SNAP_N') || 0, n = Math.ceil(z.length / SNAP_CH), kv = { SNAP_N: String(n) };
  for (let i = 0; i < n; i++) kv['SNAP_' + i] = z.slice(i * SNAP_CH, (i + 1) * SNAP_CH);
  try { props.setProperties(kv, false); for (let i = n; i < old; i++) props.deleteProperty('SNAP_' + i); }
  catch (err) { props.deleteProperty('SNAP_N'); }   // ใหญ่เกินที่เก็บได้ (500 KB) → ใช้แคชอย่างเดียว
  const cache = CacheService.getScriptCache();
  if (z.length < 100000) cache.put('mrt-all', z, 21600); else cache.remove('mrt-all');
  return out;
}
function usersOf_(ss) {
  const us = ss.getSheetByName(CFG.usersSheet), o = {};
  if (us) us.getDataRange().getValues().slice(4).forEach(r => { if (r[0]) o[String(r[0]).trim()] = { owner: String(r[1] || ''), photo: String(r[3] || '').toLowerCase() }; });
  return o;
}
/** อ่านทุกแท็บใหม่ทั้งหมด (ตอนนำเข้า / เมนู / ยังไม่เคยเตรียม) — อ่านแท็บละครั้งเดียว */
function snapBuild_(ss) {
  const sm = ss.getSheetByName(CFG.summarySheet), out = { stations: [], users: usersOf_(ss), round: sm ? String(sm.getRange(1, 1).getValue()).replace(/^สรุปคะแนน MRT\s*/, '') : '' };
  for (const sh of ss.getSheets()) {
    const nm = sh.getName();
    if (nm === CFG.usersSheet || nm === CFG.summarySheet) continue;
    const v = sh.getDataRange().getValues();
    if (v.length > 2 && String(v[2][0]) === 'คะแนนจอ') out.stations.push(sheetModel_(nm, v));
  }
  return snapSave_(out);
}
/** คนคัดรูปพิมพ์วันที่ → อ่านแท็บนั้นแท็บเดียว แล้วแทนข้อมูลของสถานีนั้น */
function snapStation_(sh) {
  const L = LockService.getScriptLock();
  if (!L.tryLock(10000)) return;
  try {
    const out = snapLoad_();
    if (!out) return;   // ยังไม่เคยเตรียม → เตรียมตอนเปิดลิงก์ครั้งแรก
    const m = sheetModel_(sh.getName(), sh.getDataRange().getValues()), i = out.stations.findIndex(x => x.station === m.station);
    if (i >= 0) out.stations[i] = m; else out.stations.push(m);
    snapSave_(out);
  } finally { L.releaseLock(); }
}
function snapUsers_(ss) {
  const L = LockService.getScriptLock();
  if (!L.tryLock(10000)) return;
  try { const out = snapLoad_(); if (out) { out.users = usersOf_(ss); snapSave_(out); } } finally { L.releaseLock(); }
}
function refreshPhone() {
  const ss = SpreadsheetApp.getActive(), out = snapBuild_(ss);
  ss.toast(`อัปเดตข้อมูลหน้ามือถือแล้ว (${out.stations.length} สถานี)`, 'ใบเช็ครูป', 6);
}
function readAll_() {
  return snapLoad_() || snapBuild_(ssOf_());
}

/** แปลงค่าทั้งแท็บ → {station, screens:[{code, type, n, done, miss:[[ชื่อหนัง, วันขึ้น]]}], newest, films} */
function sheetModel_(name, v) {
  const fr = CFG.firstRow - 1, fc = CFG.firstCol - 1, hdr = v[1].slice(fc).filter(x => String(x).trim());
  const screens = hdr.map(h => { const p = String(h).split('\n'); return { code: p[0], type: p[1] || '', n: 0, done: 0, miss: [] }; });
  let add = 0, newest = 0, first = 0; const films = [];
  for (let r = fr; r < v.length; r++) {
    const row = v[r], nm = String(row[1] || '').trim();
    if (!nm) { const m = String(row[0] || '').match(/ขึ้น(?:ใหม่)?\s*(\d+)/); if (m) { add = +m[1]; newest = Math.max(newest, add); if (!first) first = add; } continue; }
    films.push([nm, add]);
    screens.forEach((s, i) => { const x = row[fc + i]; s.n++; if (isDone_(x)) s.done++; else s.miss.push([nm, add]); });
  }
  return { station: name, screens, newest, first, films };
}

function phoneData_(email) {
  const all = readAll_();
  let mine = all.stations.filter(s => (all.users[s.station] || {}).photo && all.users[s.station].photo.split(/[,\s]+/).indexOf(email) >= 0);
  const assigned = mine.length > 0;
  if (!assigned) mine = all.stations;
  const pct = s => s.n ? s.done / s.n : 0, avg = a => a.length ? a.reduce((x, y) => x + y, 0) / a.length : 0;
  const allScr = [].concat(...all.stations.map(s => s.screens));
  const newest = Math.max(0, ...mine.map(s => s.newest)), first = Math.min(...mine.map(s => s.first || 99));
  const newFilms = []; mine.forEach(s => s.films.forEach(([nm, d]) => { if (d === newest && newFilms.indexOf(nm) < 0) newFilms.push(nm); }));
  return {
    email, assigned, round: all.round,
    updated: (() => { const d = new Date(all.at || Date.now()), f = p => Utilities.formatDate(d, CFG.tz, p); return f('yyyyMMdd') === Utilities.formatDate(new Date(), CFG.tz, 'yyyyMMdd') ? f('HH:mm') : `${f('d')} ${TH_MON[+f('M') - 1]} ${f('HH:mm')}`; })(),
    mrt: avg(allScr.map(pct)), mrtScreens: allScr.length,
    newest, newFilms: newest > first ? newFilms : [],
    stations: mine.map(s => ({ name: s.station, owner: (all.users[s.station] || {}).owner || '', screens: s.screens })),
  };
}

// ===================== รับใบเช็คจาก PHOTO CHECK STUDIO =====================
// ปุ่ม “ส่งเข้าใบเช็ค MRT” ในหน้าผลการเติม เปิดหน้านี้ (?recv=studio) → หน้านี้บอก ready → ระบบส่งไฟล์ .xlsm ที่ชื่อมี MRT มาทาง postMessage
// → recvSave เก็บลงโฟลเดอร์นำเข้า (ทับไฟล์ชื่อเดิม) → recvImport นำเข้าเหมือนเมนู (ขึ้นรอบใหม่ถามในหน้านี้) → บอกผลกลับไปที่ระบบ
// ส่งได้เฉพาะเจ้าของสคริปต์ และอีเมลในช่อง “ผู้ดูแลระบบ / หัวหน้า” ของแท็บผู้ใช้และสิทธิ์
function ssOf_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
}

function recvAllowed_() {
  const me = (Session.getActiveUser().getEmail() || '').toLowerCase(), owner = (Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (!me) return false;
  if (me === owner) return true;
  const us = ssOf_().getSheetByName(CFG.usersSheet);
  const admins = us ? String(us.getRange(2, 2).getValue()).split(/[,\s]+/).map(x => x.trim().toLowerCase()).filter(x => /@/.test(x)) : [];
  return admins.indexOf(me) >= 0;
}

function recvPage_() {
  let ok = false, round = '';
  try { ok = recvAllowed_(); round = PropertiesService.getScriptProperties().getProperty('ROUND') || ''; } catch (err) { ok = false; }
  const d = JSON.stringify({ ok, round, studio: CFG.studioOrigin, me: Session.getActiveUser().getEmail() || '' }).replace(/</g, '\\u003c');
  return HtmlService.createHtmlOutput(RECV_HTML.replace('__DATA__', () => d)).setTitle('รับใบเช็คจาก PHOTO CHECK STUDIO').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** เก็บไฟล์ที่ส่งมาลงโฟลเดอร์นำเข้า [{name, b64}] → จำนวนไฟล์ */
function recvSave(files) {
  if (!recvAllowed_()) throw new Error('ส่งได้เฉพาะผู้ดูแลระบบ / หัวหน้า (ใส่อีเมลในแท็บผู้ใช้และสิทธิ์)');
  const folder = importFolder_();
  let n = 0;
  for (const f of files || []) {
    const name = String(f && f.name || '').replace(/[\\/]/g, '_').trim();
    if (!/\.xls[xm]$/i.test(name) || !f.b64) continue;
    const old = folder.getFilesByName(name);
    while (old.hasNext()) old.next().setTrashed(true);   // ทับไฟล์ชื่อเดิม (ฉบับเก่าอยู่ในถังขยะของ Drive)
    folder.createFile(Utilities.newBlob(Utilities.base64Decode(f.b64), /\.xlsm$/i.test(name) ? 'application/vnd.ms-excel.sheet.macroEnabled.12' : MimeType.MICROSOFT_EXCEL, name));
    n++;
  }
  if (!n) throw new Error('ไม่ได้รับไฟล์ใบเช็ค (.xlsm)');
  return n;
}

/** นำเข้าเหมือนเมนู • choice: ไม่ส่ง = ถ้าขึ้นรอบใหม่ คืน {ask} / 'new' / 'keep' / 'cancel' */
function recvImport(choice) {
  if (!recvAllowed_()) throw new Error('ส่งได้เฉพาะผู้ดูแลระบบ / หัวหน้า (ใส่อีเมลในแท็บผู้ใช้และสิทธิ์)');
  const ss = ssOf_();
  return withLock_(() => importCore_(ss, () => (choice === 'new' || choice === 'keep' || choice === 'cancel') ? choice : undefined));
}

const RECV_HTML = `<!doctype html><html lang="th"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@400;500;600&family=Prompt:wght@600&display=swap" rel="stylesheet">
<style>
:root{--f-ui:"IBM Plex Sans Thai",Tahoma,sans-serif;--f-hd:"Prompt","IBM Plex Sans Thai",sans-serif;--navy:#022866;--cyan:#1BB8EF;--bg:#F4F7FB;--sf:#fff;--ln:#D4DCE8;--tx:#0B1F45;--tx2:#34446A;--mu:#5A6987;--ok:#1B7646;--okb:#2FA866;--wr:#8F5C06;--wrs:#FFF2CC;--er:#B9312A;--ers:#FCE8E6}
*{box-sizing:border-box;margin:0;padding:0}html,body{background:var(--bg);color:var(--tx);font:15px/1.5 var(--f-ui)}
.app{background:var(--navy);color:#fff;padding:14px 18px 16px}.pb{font:600 15px/1 var(--f-hd)}.pb i{font-style:normal;color:var(--cyan)}
.t{font:600 19px/1.35 var(--f-hd);margin-top:8px}.s{font-size:13px;color:#BFD0EE}
.body{padding:14px}.step{display:flex;gap:10px;background:var(--sf);border:1px solid var(--ln);border-radius:14px;padding:10px 12px;margin-bottom:8px}
.ck{width:22px;height:22px;border-radius:50%;background:var(--okb);color:#fff;font:600 13px/22px var(--f-ui);text-align:center;flex:none;margin-top:1px}
.ck.w{background:var(--ln);color:var(--mu)}.ck.r{background:var(--er)}
.step b{display:block;font-weight:600}.step span{font-size:13px;color:var(--tx2);display:block;line-height:1.45;overflow-wrap:anywhere}
.done{background:#E1F3E8;border:1px solid #A9DDBE;border-radius:14px;padding:12px 14px;color:var(--ok);font-weight:600;margin-top:4px}.done small{display:block;font-weight:500;color:var(--tx2);font-size:13px;margin-top:2px}
.bad{background:var(--ers);border:1px solid #F1C3BE;border-radius:14px;padding:12px 14px;color:var(--er);font-weight:600}.bad small{display:block;font-weight:500;color:var(--tx2);font-size:13px;margin-top:2px}
.ask{background:var(--wrs);border:1px solid #F1D98A;border-radius:14px;padding:12px 14px}.ask b{display:block;font:600 16px var(--f-hd);color:var(--wr)}.ask p{font-size:13.5px;color:var(--tx2);margin:4px 0 10px;line-height:1.5}.ask p em{font-style:normal;font-weight:600;color:var(--tx)}
.btns{display:grid;gap:8px}.btns button{font:600 14px var(--f-ui);padding:10px 12px;border-radius:12px;border:1px solid var(--ln);background:var(--sf);color:var(--tx);cursor:pointer}.btns .p{background:var(--navy);border-color:var(--navy);color:#fff}
.btns small{display:block;font-weight:500;font-size:12px;opacity:.8}
.spin{width:22px;height:22px;border-radius:50%;border:3px solid var(--ln);border-top-color:var(--navy);animation:sp 0.9s linear infinite;flex:none;margin-top:1px}
@keyframes sp{to{transform:rotate(360deg)}}@media (prefers-reduced-motion:reduce){.spin{animation:none}}
</style></head><body>
<div class="app"><div class="pb">ใบเช็ครูป <i>MRT</i></div><div class="t">รับใบเช็คจาก PHOTO CHECK STUDIO</div><div class="s" id="rd"></div></div>
<div class="body" id="b" aria-live="polite"></div>
<script>
var D = __DATA__, S = { files: null, saved: 0, step: 'wait', ask: null, res: null, err: '' };
function op() { try { return window.top.opener || null; } catch (e) { return null; } }
function tell(m) { var w = op(); if (!w) return; m.psMRT = 1; try { w.postMessage(m, D.studio); } catch (e) {} }
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function step(st, t, sub) { var ic = st === 'ok' ? '<span class="ck">✓</span>' : st === 'run' ? '<span class="spin" role="img" aria-label="กำลังทำ"></span>' : st === 'bad' ? '<span class="ck r">!</span>' : '<span class="ck w">·</span>';
  return '<div class="step">' + ic + '<div><b>' + t + '</b>' + (sub ? '<span>' + sub + '</span>' : '') + '</div></div>'; }
function draw() {
  document.getElementById('rd').textContent = D.round ? 'ชีตตอนนี้: ' + D.round : '';
  var h = '', n = S.files ? S.files.length : 0, at = ['wait', 'save', 'imp', 'ask', 'done', 'err'].indexOf(S.step);
  h += step(n ? 'ok' : S.step === 'err' ? 'bad' : 'run', n ? 'รับไฟล์ ' + n + ' ไฟล์' : 'รอไฟล์จาก PHOTO CHECK STUDIO', n ? esc(S.files.map(function (f) { return f.name; }).join(', ')) : '');
  if (n) h += step(S.saved ? 'ok' : S.step === 'save' ? 'run' : S.step === 'err' ? 'bad' : 'wait', 'เก็บลงโฟลเดอร์ “ใบเช็ค MRT นำเข้า”', 'ทับไฟล์ชื่อเดิม');
  if (S.saved && S.step !== 'ask') h += step(S.res && S.res.ok ? 'ok' : S.step === 'imp' ? 'run' : S.step === 'err' ? 'bad' : 'wait', 'นำเข้าเข้าชีต',
    S.res && S.res.ok ? S.res.stations + ' สถานี • ' + (S.res.fresh ? 'เริ่ม' + esc(S.res.label) : 'หนังใหม่ ' + S.res.newFilms + ' เรื่อง') : S.step === 'imp' ? 'อาจใช้เวลา 1–3 นาที อย่าเพิ่งปิดหน้านี้' : '');
  if (S.step === 'ask') h += '<div class="ask"><b>ขึ้นรอบใหม่?</b><p>ใบเช็คที่ส่งมาเป็น <em>' + esc(S.ask.label) + '</em> • ชีตตอนนี้เป็น ' + esc(S.ask.prev) + '</p><div class="btns">' +
    '<button class="p" data-c="new">ขึ้นรอบใหม่<small>เก็บสำเนา ' + esc(S.ask.prev) + ' ไว้ใน Drive แล้วเริ่มชีตใหม่</small></button>' +
    '<button data-c="keep">นำเข้าในรอบเดิม<small>ยังเป็น ' + esc(S.ask.prev) + '</small></button><button data-c="cancel">ยังไม่นำเข้า</button></div></div>';
  if (S.step === 'done') h += S.res && S.res.ok ? '<div class="done">เสร็จแล้ว<small>วันที่ที่คนคัดรูปพิมพ์ไว้ในชีตไม่หาย • หน้ามือถือของคนถ่ายรูปเห็นหนังใหม่แล้ว • หน้าต่างนี้ปิดเองใน 3 วินาที</small></div>'
    : '<div class="bad">ยังไม่ได้นำเข้า<small>ไฟล์อยู่ในโฟลเดอร์ “ใบเช็ค MRT นำเข้า” แล้ว — นำเข้าทีหลังได้จากเมนู ใบเช็ครูป → นำเข้าใบเช็ค Excel จาก Drive • ปิดหน้านี้ได้</small></div>';
  if (S.step === 'err') h += '<div class="bad">ทำไม่สำเร็จ<small>' + esc(S.err) + '</small></div>';
  document.getElementById('b').innerHTML = h;
}
function fail(e) { S.step = 'err'; S.err = (e && e.message) || String(e || 'ไม่ทราบสาเหตุ'); draw(); tell({ t: 'done', ok: false, err: S.err }); }
function b64(buf) { var u = new Uint8Array(buf), s = '', CH = 0x8000; for (var i = 0; i < u.length; i += CH) s += String.fromCharCode.apply(null, u.subarray(i, i + CH)); return btoa(s); }
function imp(choice) {
  S.step = 'imp'; S.ask = null; draw();
  google.script.run.withFailureHandler(fail).withSuccessHandler(function (r) {
    if (r && r.ask) { S.step = 'ask'; S.ask = r; draw(); tell({ t: 'ask', label: r.label, prev: r.prev }); return; }
    if (r && r.cancel) { S.step = 'done'; S.res = r; draw(); tell({ t: 'done', ok: false, cancel: true }); return; }
    if (!r || !r.ok) { fail({ message: r && r.none ? 'ไม่พบไฟล์ในโฟลเดอร์นำเข้า' : 'นำเข้าไม่ได้' }); return; }
    S.step = 'done'; S.res = r; draw(); tell({ t: 'done', ok: true, files: r.files, stations: r.stations, newFilms: r.newFilms, fresh: r.fresh, label: r.label });
  }).recvImport(choice || null);
}
window.addEventListener('message', function (e) {
  if (e.origin !== D.studio || !e.data || !e.data.psMRT || e.data.t !== 'files' || S.files) return;
  S.files = (e.data.files || []).filter(function (f) { return f && f.name && f.buf; });
  if (!S.files.length) { fail({ message: 'ไม่ได้รับไฟล์ใบเช็ค MRT' }); return; }
  S.step = 'save'; draw();
  var pay; try { pay = S.files.map(function (f) { return { name: f.name, b64: b64(f.buf) }; }); } catch (x) { fail(x); return; }
  google.script.run.withFailureHandler(fail).withSuccessHandler(function (n) { S.saved = n; imp(null); }).recvSave(pay);
});
document.addEventListener('click', function (e) { var b = e.target.closest && e.target.closest('[data-c]'); if (!b || S.step !== 'ask') return; var c = b.getAttribute('data-c');
  if (c === 'cancel') { S.step = 'done'; S.res = { cancel: true }; draw(); tell({ t: 'done', ok: false, cancel: true }); return; } imp(c); });
if (!D.ok) { S.step = 'err'; S.err = (D.me ? D.me + ' ' : '') + 'ส่งใบเช็คเข้าชีตนี้ไม่ได้ — ส่งได้เฉพาะเจ้าของชีต และอีเมลในช่อง “ผู้ดูแลระบบ / หัวหน้า” ของแท็บผู้ใช้และสิทธิ์'; draw(); tell({ t: 'done', ok: false, err: S.err }); }
else if (!op()) { S.step = 'err'; S.err = 'เปิดหน้านี้จากปุ่ม “ส่งเข้าใบเช็ค MRT” ในหน้าผลการเติมของ PHOTO CHECK STUDIO'; draw(); }
else { draw(); tell({ t: 'ready' }); }
</script></body></html>`;
