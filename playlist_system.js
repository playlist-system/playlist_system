/*
 * ======================================================================
 *  ระบบเติมใบเช็ครูปอัตโนมัติ — โค้ด JavaScript ของระบบ
 *  เวอร์ชัน 3.8.1 • 20 ต.ค. 2569
 *  ระบบ : https://playlist-system.github.io/playlist_system/
 *  โค้ด : https://github.com/playlist-system/playlist_system
 * ======================================================================
 *
 *  ไฟล์นี้คือโค้ดทั้งหมดที่เขียนขึ้นสำหรับระบบ ดึงออกมาจาก playlist_system.html
 *  (ในไฟล์ HTML โค้ดชุดนี้อยู่ในแท็ก <script> อันสุดท้าย) จัดย่อหน้าใหม่ให้อ่านง่าย
 *  การทำงานเหมือนเดิมทุกอย่าง
 *
 *  ต้องใช้ร่วมกับ
 *    - SheetJS (xlsx) 0.18.5  → ตัวแปร XLSX   ใช้อ่านไฟล์ Excel
 *    - JSZip 3.10.1           → ตัวแปร JSZip  ใช้เปิด/บันทึกไฟล์ .xlsm โดยมาโครอยู่ครบ
 *      (ไลบรารี open source ทั้งสองตัวฝังอยู่ใน playlist_system.html แล้ว ไม่ได้รวมในไฟล์นี้)
 *    - โครง HTML + CSS ของ playlist_system.html  (#steps #app #stat #ovl #msg)
 *
 *  ถ้าจะใช้ไฟล์นี้แยกจาก HTML ให้โหลดตามลำดับนี้ ไว้ท้าย <body>
 *    <script src="xlsx.full.min.js"></script>
 *    <script src="jszip.min.js"></script>
 *    <script src="playlist_system.js"></script>
 *
 *  สารบัญ (ค้นด้วยชื่อฟังก์ชัน)
 *    1. ค่าพื้นฐาน รหัสจอ คอลัมน์ Excel สื่อ ......... codeOf, col2n, n2col, mediaOf
 *    2. แก้ XML ของชีตในไฟล์ใบเช็คโดยตรง ............. class Sheet, mkStyles
 *    3. จับคู่จอ Signature Max / Show DC ........... namedScreen, SIGAL, noteSkip
 *    4. วันที่ (ชื่อไฟล์ / ข้อมูล / วันทำงาน) ......... parseD, fullNameDate, nameDate, workDate
 *    5. จัดใบเช็คแบบที่ทำด้วยมือ (กฎทั้งหมด) ........ layoutChk
 *    6. เปิดใบเช็ค จับคู่ชีตกับจอ แล้วเติม ............ fillWb, avCols, screen10
 *    7. อ่าน Playlist 5 รูปแบบ + ด่านตรวจ Playlist .... parseSig, parseMetro, parseSplit, parsePlaylist,
 *                                                     parseSect, parseMco, slimXlsx, plCheck, noteQ, addFiles
 *    8. เช็กลิสต์ Playlist 11 แบบ / ความพร้อมใบเช็ค .... KINDS, kState, sheetNeeds, readyOf
 *    9. หน้าจอ 4 ขั้นตอน ............................. vData, vWall, vPl, vEd, vTl, vFl, vRes
 *   10. วิธีใช้ / ต้อนรับ / ระบบทำงานอย่างไร ............ vHelp, vWelcome, vHow
 *   11. ทัวร์แนะนำ ................................... TOURS, tourStart, tourDraw
 *   12. ปุ่มและเหตุการณ์ .............................. view, ACT, document.addEventListener
 *  (เวอร์ชัน 2.4 ปรับกฎในส่วน 3, 5 และ 7 จากการเทียบกับใบเช็คที่ทีมทำมือเดือน ก.ย.–ต.ค. 2569
 *   เช่น วันที่ 16 ทำต่อในรายการเดิม หมายเหตุ BKJ / Metro LED — ดูประวัติเวอร์ชันใน README.md
 *   เวอร์ชัน 2.5 แก้ส่วน 2, 5 และ 6: จอที่ไม่มีใน Playlist ของวันนั้นคงรายการเดิมไว้ในชีต
 *   ให้แถบวันที่ตรงกันทั้งชีต อ่านแถบวันที่ที่เป็นสูตร =F45 ได้ และลบ calcChain กัน Excel ซ่อมไฟล์
 *   เวอร์ชัน 2.6 แก้ส่วน 5: หนังเรื่องเดียวมีหลายแถวใน Playlist ไม่ลงซ้ำใต้วันเริ่มสัญญาที่เก่ากว่า
 *   และหนังที่ออกแล้วกลับมาในครึ่งเดือนเดียวกันใช้แถวเดิม
 *   เวอร์ชัน 2.7 แก้ส่วน 5, 6 และ 9: แถบวันที่ทุกวัน, ไม่เตือนหมายเหตุ 16:9 / 4:3, จอ Spare ใช้รายการสื่ออื่นได้
 *   และ 2.7.1: จอ CITY ที่ไม่มีใน Playlist ใช้รายการ THE CITY เป็นเรื่องปกติ
 *   เวอร์ชัน 2.8 แก้ส่วน 5: ต่อสัญญาวันที่ 16 พอหนังถอดจริง ใส่วันถอดทั้งแถวเดิมและแถวใต้ 16
 *   เวอร์ชัน 2.9 เพิ่มในส่วน 9 และ 12: บันทึกทับไฟล์เดิม (Chrome / Edge) เปิดใบเช็คทั้งโฟลเดอร์ และใช้ผลเติมวันถัดไป
 *   2.9.1: วันที่ไม่มีหนังใหม่ไม่ต้องมีแถบวันที่ (ยกเลิกแถบทุกวันของ 2.7)
 *   2.9.2: ไม่เขียนลงชีต Analysis / XXXX / ชีตสรุป ไม่ว่าจะจัดหน้าแบบไหน
 *   เวอร์ชัน 3.0 เพิ่มในส่วน 6, 7, 9 และ 12: ด่านตรวจ Playlist — อ่านคอลัมน์ตามชื่อหัวตาราง
 *   เตือนสถานะ / วันที่ / สื่อที่ไม่รู้จักก่อนเติม และหมายเหตุที่อ่านไม่ออกใน “ควรตรวจ”
 *   3.0.1 แก้ส่วน 5: วันที่ 16 สถานะต่อสัญญาที่ Playlist ยังลงวันที่สัญญาเดิม (หมด 15) ถือเป็นสัญญาใหม่เริ่ม 16
 *   3.0.2 แก้ส่วน 5: แถวว่างใต้แถบสุดท้ายในช่องชื่อหนังมีเส้นตารางตามช่องนับ
 *   3.0.3 แก้ส่วน 5: สีแถบรอบ 2 แบบ “สีตามชีต” ใช้สีแถบรอบ 2 เดิมของชีต (อ่านก่อนล้างแถว) ถ้ายังไม่มีใช้ฟ้า
 *   เวอร์ชัน 3.1 แก้ส่วน 1, 5, 7 และ 9: Metro LED อ่านสถานะจากสีแถวและอ่านระยะสัญญา (parseMetro)
 *   แถว #FAST_ADS / FastAd ไม่ใส่จาก Playlist และลบออกจากใบเช็คตอนเติม (FASTN)
 *   เวอร์ชัน 3.2 แก้ส่วน 5 และ 9: วันที่ถ่าย (ช่องหลังชื่อหนัง) ของบล็อกที่มีหลายจอ เช่น BKJ / Metro LED
 *   ย้ายตามชื่อหนังครบทุกคอลัมน์ และขึ้นเดือนใหม่ล้างทุกคอลัมน์ (b.pcs ใน layoutChk)
 *   3.2.1 แก้ส่วน 5: ครึ่งเดือนหลัง หนังใหม่ที่หมายเหตุ “ขึ้นระหว่างวัน วันที่ …” ลงวันที่ 1–15 ลงใต้แถบ 16
 *   เวอร์ชัน 3.3 แก้ส่วน 7: อ่าน Playlist เร็วขึ้น — ตัดช่องว่างที่มีแต่รูปแบบออกก่อนให้ SheetJS อ่าน (slimXlsx)
 *   เวอร์ชัน 3.4 แก้ส่วน 4, 5, 6, 7, 8 และ 9: Playlist เพิ่มเติม 5 แบบ (Rama9 Connect / Digital MRT / Gateway X: parseSect,
 *   Air Traveller / The 20: parseMco) และใบเช็คแบบใหม่ที่แถวสูตรวิเคราะห์อยู่แถว 10 (CCTV / MRT / Gateway X / Rama9 Connect /
 *   Air Traveller / THE 20: avCols, screen10)
 *   3.4.1 แก้ส่วน 3: หมายเหตุ “ทุกจอยกเว้น <ชื่อย่อจอ>” ของ Signature Max / Show DC (noteSkip)
 *   เวอร์ชัน 3.5 แก้ส่วน 5, 6, 7 และ 8: Playlist LA PlanB (parseSig, laDate — ไม่ได้ส่งทุกวัน ใช้ฉบับล่าสุด) และชีต LA PLANB
 *   (kind10, screen10) • คำในช่องออกตามแบบที่ใบเช็คใช้ทั้งไฟล์ (OSW, OSH) • รหัสจอที่ขีดต่างกัน (DGT-174 = DGT174)
 *   • หนัง Content ที่ใบเช็คเขียนไว้แบบไม่มี # ใช้แถวเดิม
 *   เวอร์ชัน 3.6 แก้ส่วน 5 และ 9: แถบวันที่ตามแบบของแต่ละใบ — แถบวันถัดไปท้ายตาราง (TRAIL) และวันที่ในแถบของจอที่ 2
 *   เป็นตัวเลขหรือสูตร (BVAL) • ตัวเลือกรายไฟล์ “ต่อสัญญาวันที่ 16 แถวเดิมไม่ใส่ออก” (k16 ใน fcol)
 *   เวอร์ชัน 3.7 เพิ่มในส่วน 5, 6, 9 และ 12: สร้างโฟลเดอร์แยกรูปของรอบปัจจุบันจากใบเช็คที่เติมแล้ว (picId, picPlan, picMake, vPic)
 *   เวอร์ชัน 3.8 เพิ่มในส่วนเดียวกัน: ตรวจโฟลเดอร์เทียบจอ — เทียบใบเช็คที่เติมแล้วกับโฟลเดอร์รอบ → แพค → จอในเครื่อง (cmpCK, cmpCheck, vCmp))
 * ======================================================================
 */

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
  RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CODEX = /^(DGT[-A-Za-z0-9]*|ID-\d+[A-Za-z0-9-]*)(?![A-Za-z0-9])/i,
  FIRST = 22;
const codeOf = v => {
  const m = cl(v).replace(/^\s+/, '').match(CODEX);
  return m ? m[1].toUpperCase() : '';
};
const cl = v => {
  if (v == null) return '';
  v = String(v).trim();
  return v === '#N/A' || v === '#REF!' || v === '#VALUE!' || v === '0' ? '' : v;
};
const col2n = s => [...s].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
const n2col = n => {
  let s = '';
  while (n) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = (n - r - 1) / 26;
  }
  return s;
};
function mediaOf(t) {
  const n = t.replace(/[^A-Za-z0-9]/g, '').toUpperCase();
  if (n.includes('CITY') || n.includes('CYTY')) return 'CITY';
  if (n.includes('SPARE')) return 'Spare';
  for (const z of ['40', '80']) {
    if (n.includes(z) && n.includes('BLUE')) return z + 'Blue';
    if (n.includes(z) && n.includes('RED')) return z + 'Red';
  }
  return null;
}
/* ชีตสรุป/ชีตว่างของแม่แบบ ไม่ใช่ชีตจอ: ไม่อ่านและไม่เขียนเลย ไม่ว่าจะจัดหน้าแบบไหน (ใบเช็คแต่ละคนจัดชีต Analysis ไม่เหมือนกัน) */
const notChk = sn => {
  const u = String(sn)
    .toUpperCase()
    .replace(/[^A-Z]/g, '');
  return /^X+$/.test(u) || /ANALYS|SUMMARY|DASHBOARD/.test(u) || /^\s*(สรุป|วิเคราะห์|ผลวิเคราะห์)/.test(sn);
};
/* ---------- check sheet xml ---------- */
class Sheet {
  constructor(xml) {
    this.doc = new DOMParser().parseFromString(xml, 'application/xml');
    this.sd = this.doc.getElementsByTagNameNS(NS, 'sheetData')[0];
    this.rows = new Map();
    for (const r of this.sd.children) this.rows.set(+r.getAttribute('r'), r);
  }
  row(n, mk = true) {
    if (this.rows.has(n)) return this.rows.get(n);
    if (!mk) return null;
    const el = this.doc.createElementNS(NS, 'row');
    el.setAttribute('r', n);
    let nxt = null;
    for (const r of this.sd.children) {
      if (+r.getAttribute('r') > n) {
        nxt = r;
        break;
      }
    }
    this.sd.insertBefore(el, nxt);
    this.rows.set(n, el);
    return el;
  }
  cell(col, n, mk = true) {
    const row = this.row(n, mk);
    if (!row) return null;
    row.removeAttribute('spans');
    const ref = col + n,
      cn = col2n(col);
    let nxt = null;
    for (const c of row.children) {
      const cc = col2n(c.getAttribute('r').match(/[A-Z]+/)[0]);
      if (c.getAttribute('r') === ref) return c;
      if (cc > cn) {
        nxt = c;
        break;
      }
    }
    if (!mk) return null;
    const el = this.doc.createElementNS(NS, 'c');
    el.setAttribute('r', ref);
    row.insertBefore(el, nxt);
    return el;
  }
  setText(col, n, t) {
    const c = this.cell(col, n);
    while (c.firstChild) c.removeChild(c.firstChild);
    c.setAttribute('t', 'inlineStr');
    const is = this.doc.createElementNS(NS, 'is'),
      te = this.doc.createElementNS(NS, 't');
    te.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    te.textContent = t;
    is.appendChild(te);
    c.appendChild(is);
  }
  clear(col, n) {
    const c = this.cell(col, n, false);
    if (!c) return;
    while (c.firstChild) c.removeChild(c.firstChild);
    c.removeAttribute('t');
  }
  getS(col, n) {
    const c = this.cell(col, n, false);
    return c ? c.getAttribute('s') : null;
  }
  take(col, n) {
    const c = this.cell(col, n, false);
    if (
      !c ||
      c.getElementsByTagNameNS(NS, 'f').length ||
      !(c.getElementsByTagNameNS(NS, 'v').length || c.getElementsByTagNameNS(NS, 'is').length)
    )
      return null;
    const k = c.cloneNode(true);
    while (c.firstChild) c.removeChild(c.firstChild);
    c.removeAttribute('t');
    return k;
  }
  give(col, n, k) {
    const c = this.cell(col, n);
    while (c.firstChild) c.removeChild(c.firstChild);
    c.removeAttribute('t');
    if (k.getAttribute('t')) c.setAttribute('t', k.getAttribute('t'));
    if (k.getAttribute('s') != null) c.setAttribute('s', k.getAttribute('s'));
    for (const ch of [...k.childNodes]) c.appendChild(ch.cloneNode(true));
  }
  has(col, n) {
    const c = this.cell(col, n, false);
    return (
      !!c && (c.getElementsByTagNameNS(NS, 'v').length > 0 || c.getElementsByTagNameNS(NS, 'is').length > 0)
    );
  }
  setS(col, n, s) {
    const c = this.cell(col, n);
    if (s != null) c.setAttribute('s', s);
    else c.removeAttribute('s');
  }
  setNum(col, n, v) {
    const c = this.cell(col, n);
    while (c.firstChild) c.removeChild(c.firstChild);
    c.removeAttribute('t');
    const e = this.doc.createElementNS(NS, 'v');
    e.textContent = String(v);
    c.appendChild(e);
  }
  setF(col, n, f, v) {
    const c = this.cell(col, n);
    while (c.firstChild) c.removeChild(c.firstChild);
    c.removeAttribute('t');
    const e = this.doc.createElementNS(NS, 'f');
    e.textContent = f;
    c.appendChild(e);
    if (v != null) {
      const x = this.doc.createElementNS(NS, 'v');
      x.textContent = String(v);
      c.appendChild(x);
    }
  }
  numRow(col, a, b) {
    for (let r = a; r <= b; r++) {
      const c = this.cell(col, r, false);
      if (
        c &&
        !c.getAttribute('t') &&
        c.getElementsByTagNameNS(NS, 'v').length &&
        !c.getElementsByTagNameNS(NS, 'f').length
      )
        return r;
    }
    return 0;
  }
  xml() {
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      new XMLSerializer().serializeToString(this.doc.documentElement)
    );
  }
}
const nrm = t =>
  String(t == null ? '' : t)
    .replace(/[^A-Za-z0-9\u0E00-\u0E7F]/g, '')
    .toUpperCase();
function namedScreen(label, code, sn) {
  if (/^BASE$/.test(nrm(sn || ''))) {
    const h = DB.filter(s => !/^(40|80)/.test(s.m) && /BASE/.test(nrm(s.c + ' ' + s.n)));
    return h.length === 1 ? h[0] : h.find(s => nrm(s.c) === 'BASE') || null;
  }
  const n = nrm(label + code),
    by = nm => DB.find(s => (s.m === 'Signature Max' || s.m === 'Show DC / Prudential') && s.c === nm);
  if (/BIGY/.test(n)) return by('The Big Y');
  if (/SOBANGKOK|SOBKK/.test(n)) return by('So BKK -H-V-');
  if (/TWINTUBE|TWT/.test(n)) return by('Twin Tube');
  if (/PLOENCHIT/.test(n)) return by('Ploenchit Interchange');
  if (/NARONG/.test(n)) return by('At Narong Interchange');
  if (/AIRTRAVELL/.test(n)) return by('Air Traveller Gateway');
  if (/SHOWDC/.test(n)) return by('ShowDC');
  if (/PRUDENTIAL/.test(n)) return by('PrudentialSathorn');
  if (/DUSITCENTRAL/.test(n)) return by('Dusit Central Park');
  if (/EXCHANGETOWER/.test(n))
    return by(
      /SM.*2|2.*SM/.test(n)
        ? 'Exchange Tower (SM) 2'
        : /SM/.test(n)
          ? 'Exchange Tower (SM)'
          : 'Exchange Tower',
    );
  const L = nrm(label);
  if (L.length >= 4) {
    const h = DB.filter(
      s =>
        (s.m === 'Signature Max' || s.m === 'Show DC / Prudential') &&
        (nrm(s.c).includes(L) || L.includes(nrm(s.c))),
    );
    if (h.length === 1) return h[0];
  }
  return null;
}
const jamList = (s, o) =>
  listOf(s, o).filter(a => !/^BJ\d*_\d+_/i.test(a) && !/weather/i.test(a) && !/^เวลารวม|^รวมทุกแพค/.test(a));
/* ---------- แบบใบเช็คที่ทำด้วยมือ: # ก่อน → A-Z, แยกตามวันเริ่ม, ออกวันที่ ---------- */
const SCR = new Map(),
  MON = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
function parseD(t) {
  if (!t) return null;
  {
    const x = String(t).trim();
    let q = x.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (q) return new Date(Date.UTC(+q[1], +q[2] - 1, +q[3]));
    q = x.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})$/);
    if (q) {
      let y = +q[3];
      if (y < 100) y += 2000;
      if (y > 2400) y -= 543;
      if (+q[2] >= 1 && +q[2] <= 12) return new Date(Date.UTC(y, +q[2] - 1, +q[1]));
    }
  }
  const m = String(t)
    .trim()
    .match(/^(\d{1,2})[-\/ ]([A-Za-z]{3})[A-Za-z]*[-\/ ](\d{2,4})$/);
  if (!m || !(m[2].toLowerCase() in MON)) return null;
  let y = +m[3];
  if (y < 100) y += 2000;
  if (y > 2400) y -= 543;
  return new Date(Date.UTC(y, MON[m[2].toLowerCase()], +m[1]));
}
const serial = d => Math.round((d - Date.UTC(1899, 11, 30)) / 864e5),
  addD = (d, n) => new Date(d.getTime() + n * 864e5);
const JUNK =
    /^(เวลา|รวม|หนังโฆษณา|Version$)|^\d+(\.\d+)?$|^Programm?atic|^ไม่มีไฟล์|เมื่อมีไฟล์|^เริ่มออน|^รอ\s*Memo|^Content\s.*นาที|^On\s*\d{1,2}\s*[A-Za-z]{3}/i,
  MTP = /^(Traffic MTP|PLACE MTP|ETA MTP|EDA)\s*$/i,
  FASTN = /(^|[^a-z])Fast[\s_-]*Ads?/i; /* FastAd4-3_V2, #FAST_ADS: ไม่ใช่หนังโฆษณา ไม่ลงใบเช็ค */
/* เรียงแบบ Excel (A→Z): ไม่สนตัวพิมพ์, ข้าม - และ ', เครื่องหมายอื่น < ตัวเลข < A-Z < ภาษาไทย */
const xk = s => {
  s = s.toLowerCase().replace(/[-']/g, '');
  let k = '';
  for (const ch of s) {
    const c = ch.codePointAt(0),
      r = /[0-9]/.test(ch) ? 1 : /[a-z]/.test(ch) ? 2 : /\p{L}/u.test(ch) ? 3 : 0;
    k += String.fromCharCode(48 + r) + String.fromCharCode(c < 0xffff ? c : 0xffff);
  }
  return k;
};
const cmpN = (a, b) => {
  const x = xk(a),
    y = xk(b);
  return x < y ? -1 : x > y ? 1 : a < b ? -1 : a > b ? 1 : 0;
};
/* วันที่ทำงาน = วันของ Playlist: จากชื่อไฟล์ (…แยกจอ_DD_…, …_1-10-26) หรือที่พิมพ์เอง */
let PLD = null,
  PLM = {},
  PLX = [],
  PLI =
    []; /* PLI: ด่านตรวจ Playlist {f,lv:'stop'|'info',t} */ /* PLM: สื่อ → {f:ชื่อไฟล์, d:วันที่} • PLX: ไฟล์ที่อ่านไม่ได้ */
const THM = {
  'ม.ค.': 0,
  'ก.พ.': 1,
  'มี.ค.': 2,
  'เม.ย.': 3,
  'พ.ค.': 4,
  'มิ.ย.': 5,
  'ก.ค.': 6,
  'ส.ค.': 7,
  'ก.ย.': 8,
  'ต.ค.': 9,
  'พ.ย.': 10,
  'ธ.ค.': 11,
};
const yy = y => {
  y = +y;
  if (y < 100) y += y >= 60 ? 2500 : 2000;
  if (y > 2400) y -= 543;
  return y;
};
/* วันที่ที่เขียนไว้ชัดในชื่อไฟล์ เช่น 2-10-26, 02.10.2026, 2 ต.ค. 69, 2Oct26 */
function fullNameDate(n) {
  const d0 = fullNameDate0(n);
  if (d0) return d0;
  const q = String(n).match(/(?:^|[^\d])(20\d{2})(\d{2})(\d{2})(?!\d)/);
  return q && +q[2] >= 1 && +q[2] <= 12 && +q[3] >= 1 && +q[3] <= 31
    ? new Date(Date.UTC(+q[1], +q[2] - 1, +q[3]))
    : null;
}
function fullNameDate0(n) {
  n = String(n).replace(/\.xls[xm]?$/i, '');
  const now = new Date().getFullYear();
  let m = n.match(/(?:^|[^\d])(\d{4})[-_. ](\d{1,2})[-_. ](\d{1,2})(?!\d)/);
  if (m) {
    const y = yy(m[1]),
      mo = +m[2] - 1,
      d = +m[3];
    if (mo >= 0 && mo < 12 && d >= 1 && d <= 31) return new Date(Date.UTC(y, mo, d));
  }
  m = n.match(
    /(?:^|[^\d])(\d{4})[-_. ]?(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[-_. ]?(\d{1,2})(?!\d)/i,
  );
  if (m) return new Date(Date.UTC(yy(m[1]), MON[m[2].toLowerCase()], +m[3]));
  m = n.match(/(?:^|[^\d])(\d{1,2})[-_.](\d{1,2})[-_.](\d{2,4})(?!\d)/);
  if (m) {
    const mo = +m[2] - 1,
      d = +m[1];
    if (mo >= 0 && mo < 12 && d >= 1 && d <= 31) return new Date(Date.UTC(yy(m[3]), mo, d));
  }
  m = n
    .replace(/\s+/g, '')
    .match(
      /(\d{1,2})(ม\.?ค\.?|ก\.?พ\.?|มี\.?ค\.?|เม\.?ย\.?|พ\.?ค\.?|มิ\.?ย\.?|ก\.?ค\.?|ส\.?ค\.?|ก\.?ย\.?|ต\.?ค\.?|พ\.?ย\.?|ธ\.?ค\.?)(\d{2,4})?/,
    );
  if (m) {
    const k = Object.keys(THM).find(x => x.replace(/\./g, '') === m[2].replace(/\./g, ''));
    if (k) return new Date(Date.UTC(m[3] ? yy(m[3]) : now, THM[k], +m[1]));
  }
  m = n.match(
    /(?:^|[^A-Za-z\d])(\d{1,2})[-_ ]?(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[-_ ]?(\d{2,4})?(?![A-Za-z])/i,
  );
  if (m) return new Date(Date.UTC(m[3] ? yy(m[3]) : now, MON[m[2].toLowerCase()], +m[1]));
  return null;
}
const today0 = () => {
  const n = new Date();
  return Date.UTC(n.getFullYear(), n.getMonth(), n.getDate());
};
/* วันที่จากข้อมูลในไฟล์: วันเริ่มล่าสุดของหนังที่ออนอยู่ (ไม่เกินวันนี้) */
function dataDate(p) {
  let b = 0;
  const cap = today0();
  for (const s of p)
    for (const r of s.rows) {
      if (r.off || !r.ad || JUNK.test(r.ad.trim()) || /เมื่อมีไฟล์/.test(r.ad)) continue;
      const d = parseD(r.s);
      if (d && +d <= cap && +d > b) b = +d;
    }
  return b ? new Date(b) : null;
}
function nameDate(n, ref) {
  const d = fullNameDate(n);
  if (d) return d;
  const m = String(n).match(/แยกจอ_?(\d{1,2})_/);
  if (m && ref && +m[1] >= 1 && +m[1] <= 31)
    return new Date(Date.UTC(ref.getUTCFullYear(), ref.getUTCMonth(), +m[1]));
  return null;
}
const plDate = () => {
  let b = 0,
    a = 0;
  for (const s of DB) {
    if (s.pd > a) a = s.pd;
    if (s.m !== 'Show DC / Prudential' && s.m !== 'LA PlanB' && s.pd > b) b = s.pd;
  }
  b = b || a;
  return b ? new Date(b) : null;
}; /* Show DC / LA PlanB ไม่ได้ส่งทุกวัน (อาจเป็นฉบับเก่าหรือของรอบถัดไป) จึงไม่ใช้กำหนดวันทำงาน */
const workDate = () => parseD(OP.base) || PLD || autoBase();
function autoBase() {
  const c = new Map();
  for (const s of DB)
    for (const r of s.rows)
      if (!r.off && isNew(r)) {
        const d = parseD(r.s);
        if (d) c.set(+d, (c.get(+d) || 0) + 1);
      }
  let best = 0,
    bn = 0;
  for (const [k, v] of c)
    if (v > bn) {
      bn = v;
      best = k;
    }
  return best ? new Date(best) : null;
}
function chkRows(scr, o, kind) {
  let L = scr.rows.filter(r => o.off || !r.off);
  if (o.newOnly) L = L.filter(isNew);
  return L.filter(
    r =>
      r.ad &&
      !JUNK.test(r.ad.trim()) &&
      r.ad !== scr.c &&
      !FASTN.test(r.ad) &&
      !skipRow(scr, r) &&
      !(SIGAL[scr.c] && r.cnt == null && /^Content/i.test(r.ty || '')) &&
      !(kind === 'jam' && (/^BJ\d*_\d+_/i.test(r.ad) || /weather/i.test(r.ad))),
  ).map(r => ({
    ad: (kind === 'jam' ? r.ad.replace(/^BKJ[12]?_/i, '') : r.ad).replace(
      /\s*\((?:O\/A|ออน)[^)]*\)\s*$/i,
      '',
    ),
    hash: /^Content/i.test(r.ty || ''),
    sp: kind === 'metro' && MTP.test(r.ad.trim()),
    sd: parseD(r.s),
    ed: parseD(r.e),
    st: r.st || '',
    off: r.off,
    note: r.note || '',
    loc: r.loc || '',
  }));
}
function mkStyles(doc) {
  const g = n => doc.getElementsByTagNameNS(NS, n)[0],
    cache = {},
    XS = new XMLSerializer(),
    ser = e => XS.serializeToString(e);
  const addU = (par, el) => {
    const t = ser(el),
      ch = par.children;
    for (let i = 0; i < ch.length; i++) if (ser(ch[i]) === t) return i;
    par.appendChild(el);
    par.setAttribute('count', ch.length);
    return ch.length - 1;
  };
  return {
    doc,
    fillOf(st) {
      const xf = g('cellXfs').children[+st || 0];
      if (!xf) return '';
      const fl = g('fills').children[+xf.getAttribute('fillId') || 0];
      const fg = fl && fl.getElementsByTagNameNS(NS, 'fgColor')[0];
      return fg ? (fg.getAttribute('rgb') || '').toUpperCase() : '';
    },
    hasBorder(st) {
      const xf = g('cellXfs').children[+st || 0];
      if (!xf) return false;
      const bo = g('borders').children[+xf.getAttribute('borderId') || 0];
      return !!bo && [...bo.children].some(e => e.getAttribute('style'));
    },
    derive(base, key, opt) {
      const k = base + '|' + key;
      if (k in cache) return cache[k];
      const xfs = g('cellXfs'),
        xf = xfs.children[+base || 0].cloneNode(true);
      if (opt.fill) {
        const fl = g('fills'),
          f = doc.createElementNS(NS, 'fill'),
          pf = doc.createElementNS(NS, 'patternFill');
        pf.setAttribute('patternType', 'solid');
        const fg = doc.createElementNS(NS, 'fgColor');
        fg.setAttribute('rgb', opt.fill);
        const bg = doc.createElementNS(NS, 'bgColor');
        bg.setAttribute('indexed', '64');
        pf.appendChild(fg);
        pf.appendChild(bg);
        f.appendChild(pf);
        xf.setAttribute('fillId', addU(fl, f));
        xf.setAttribute('applyFill', '1');
      }
      if (opt.color) {
        const fo = g('fonts'),
          bf = fo.children[+xf.getAttribute('fontId') || 0].cloneNode(true);
        for (const c of [...bf.getElementsByTagNameNS(NS, 'color')]) bf.removeChild(c);
        const c = doc.createElementNS(NS, 'color');
        c.setAttribute('rgb', opt.color);
        const sz = bf.getElementsByTagNameNS(NS, 'sz')[0];
        if (sz) bf.insertBefore(c, sz.nextSibling);
        else bf.appendChild(c);
        xf.setAttribute('fontId', addU(fo, bf));
        xf.setAttribute('applyFont', '1');
      }
      if (opt.border) {
        const bs = g('borders'),
          bo = doc.createElementNS(NS, 'border');
        for (const k of ['left', 'right', 'top', 'bottom']) {
          const e = doc.createElementNS(NS, k);
          e.setAttribute('style', 'thin');
          const co = doc.createElementNS(NS, 'color');
          co.setAttribute('indexed', '64');
          e.appendChild(co);
          bo.appendChild(e);
        }
        bo.appendChild(doc.createElementNS(NS, 'diagonal'));
        xf.setAttribute('borderId', addU(bs, bo));
        xf.setAttribute('applyBorder', '1');
      }
      if (opt.right) {
        let al = xf.getElementsByTagNameNS(NS, 'alignment')[0];
        if (!al) {
          al = doc.createElementNS(NS, 'alignment');
          xf.appendChild(al);
        }
        al.setAttribute('horizontal', 'right');
        xf.setAttribute('applyAlignment', '1');
      }
      return (cache[k] = addU(xfs, xf));
    },
    xml() {
      return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        new XMLSerializer().serializeToString(doc.documentElement)
      );
    },
  };
}
const modeS = (sh, col, a, b) => {
  const c = new Map();
  for (let r = a; r <= b; r++) {
    const v = sh.getS(col, r);
    if (v != null) c.set(v, (c.get(v) || 0) + 1);
  }
  let best = null,
    bn = 0;
  for (const [k, v] of c)
    if (v > bn) {
      bn = v;
      best = k;
    }
  return best;
};
/* สไตล์ของช่องที่มีชื่อหนัง (มีเส้นตาราง) กับช่องว่าง แยกกัน — กันเส้นตารางหายเมื่อรายการสั้น */
const MARK = {
  x: new Set(),
  test(v) {
    v = String(v || '').toUpperCase();
    const h = v.length === 8 ? v.slice(2) : v;
    return /^(8AEDFE|82ECFE|002060|375623|FFFF00|4472C4|2F5597|203864|1F4E79)$/.test(h) || this.x.has(h);
  },
}; /* สีที่ถือว่าเป็นแถบ/ไฮไลต์ รวมสีที่ผู้ใช้เลือก */
const modeF = (sh, col, a, b, f, ST, inc) => {
  const c = new Map();
  for (let r = a; r <= b; r++) {
    if (f != null && sh.has(col, r) !== f) continue;
    const v = sh.getS(col, r);
    if (v == null) {
      if (inc) c.set('N', (c.get('N') || 0) + 1);
      continue;
    }
    if (ST && MARK.test(ST.fillOf(v))) continue;
    c.set(v, (c.get(v) || 0) + 1);
  }
  let best = null,
    bn = 0;
  for (const [k, v] of c)
    if (v > bn) {
      bn = v;
      best = k;
    }
  return best === 'N' ? null : best;
};
/* ---------- รอบครึ่งเดือน: 1–15 แถบวันที่สีน้ำเงิน / 16–สิ้นเดือน แถบสีเขียว ต่อท้ายรอบแรกในชีตเดียวกัน ---------- */
const NAVY = 'FF002060',
  GRN = 'FF375623',
  DKB = /^(FF)?(002060|375623)$/i,
  ISG = /375623$/i;
const per = W => {
  const y = W.getUTCFullYear(),
    m = W.getUTCMonth(),
    two = W.getUTCDate() >= 16;
  return {
    two,
    P1: new Date(Date.UTC(y, m, 1)),
    P: new Date(Date.UTC(y, m, two ? 16 : 1)),
    E: two ? new Date(Date.UTC(y, m + 1, 0)) : new Date(Date.UTC(y, m, 15)),
  };
};
const s2d = v => new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 864e5);
/* สถานะที่แปลว่าเป็นไฟล์เดิมต่อเนื่อง (ขึ้นใหม่ไฟล์เดิม / ต่อสัญญา / อัปเดตเอกสาร) → ไม่ใส่ออกวันที่ ไม่เพิ่มแถวซ้ำ */
const CONT = /ไฟล์เดิม|ต่อสัญญา|อัปเดต/,
  NEWST = /ขึ้นใหม่/;
function layoutChk(sh, jobs, o, ST, W, wsv) {
  jobs.sort((x, y) => col2n(x[1]) - col2n(y[1]));
  const pv = c => n2col(col2n(c) - 1),
    nx = c => n2col(col2n(c) + 1);
  let hr = 0;
  for (const j of jobs) {
    hr = sh.numRow(pv(j[1]), 14, 40);
    if (hr) break;
  }
  if (!hr) hr = FIRST - 1;
  let { two, P1, P, E } = per(W);
  const P2 = two ? P : null,
    SCAN = 1500,
    dd = +W,
    noNotes = W.getUTCDate() === 1;
  const B = jobs.map(j => {
    const kind = j[4],
      all = kind === 'keep' ? [] : chkRows(j[3], { ...o, off: true }, kind),
      lc = j[1],
      nc = pv(lc),
      pc = nx(lc);
    return {
      lc,
      nc,
      pc,
      kind,
      n: j[5] || 1,
      all,
      scr: j[3],
      rows: o.off ? all : all.filter(r => !r.off),
      bl: modeF(sh, lc, hr + 1, hr + 400, true, ST),
      be: modeF(sh, lc, hr + 1, hr + 400, false, ST),
      bn: modeF(sh, nc, hr + 1, hr + 400, false, ST, true),
      bp: modeF(sh, pc, hr + 1, hr + 400, null, ST),
      hs: [sh.getS(nc, hr), sh.getS(lc, hr), sh.getS(pc, hr)],
    };
  });
  {
    const any = (B.find(b => b.bl != null) || {}).bl;
    for (const b of B) if (b.bl == null) b.bl = any != null ? any : b.be || sh.getS(b.lc, hr + 1) || '0';
    for (const b of B) if (!ST.hasBorder(b.bl)) b.bl = ST.derive(b.bl, 'bord', { border: true });
  }
  var PADC = 'FF8AEDFE';
  padscan: for (const b of B)
    for (let r = hr + 1; r <= hr + 400; r++)
      for (const L of [b.lc, b.pc]) {
        const st = sh.getS(L, r);
        if (st == null) continue;
        const fc = ST.fillOf(st);
        if (/^(FF)?(8AEDFE|82ECFE)$/i.test(fc)) {
          PADC = (fc.length === 6 ? 'FF' + fc : fc).toUpperCase();
          break padscan;
        }
      }
  /* อ่านของเดิมในชีต: แถบวันที่ (คอลัมน์วันที่ของบล็อกแรก) ชื่อหนัง และ “ออกวันที่” */
  const vOf = (c, r) => {
      const x = wsv && wsv[c + r];
      return x ? cl(x.v) : '';
    },
    b0 = B[0];
  /* ปีในหัววันที่ที่พิมพ์ผิด (เช่น 2025 แทน 2026) ให้ถือเป็นปีของรอบนี้ */
  const fOf = (c, r) => {
    const x = wsv && wsv[c + r];
    if (x && typeof x.f === 'string') return x.f;
    const e = sh.cell(c, r, false),
      f = e && e.getElementsByTagNameNS(NS, 'f')[0];
    return f ? f.textContent : null;
  }; /* SheetJS ไม่อ่านช่องสูตรที่ไม่มีค่าที่คำนวณไว้ → ดูสูตรจาก XML ของชีต */
  /* ช่องหลังชื่อหนัง (วันที่ถ่าย) ของจอที่มีหลายจอในบล็อกเดียว เช่น BKJ / Metro LED: คอลัมน์ติดกันที่แถว 9 เป็นสูตร AVERAGE เป็นของบล็อกเดียวกัน → ย้ายตามชื่อหนังทุกช่อง */
  let AR10 = false;
  {
    const fr = (c, rw) => {
        const x = wsv && wsv[c + rw];
        if (x && typeof x.f === 'string') return x.f;
        const R = sh.row(rw, false);
        if (R)
          for (const e of R.children)
            if (e.getAttribute('r') === c + rw) {
              const f = e.getElementsByTagNameNS(NS, 'f')[0];
              return f ? f.textContent : null;
            }
        return null;
      },
      AR = B.some(b => /^=?\s*AVERAGE/i.test(fr(b.pc, 9) || ''))
        ? 9
        : B.some(b => BLKAVG(fr(b.pc, 10)))
          ? 10
          : 9,
      avg = c => (AR === 9 ? /^=?\s*AVERAGE/i.test(fr(c, 9) || '') : BLKAVG(fr(c, 10))),
      ownC = new Set(B.flatMap(b => [b.nc, b.lc]));
    AR10 = AR === 10;
    for (const b of B) {
      b.pcs = [b.pc];
      if (avg(b.pc))
        for (let c = col2n(b.pc) + 1; c < col2n(b.pc) + 60; c++) {
          const L = n2col(c);
          if (!avg(L) || ownC.has(L)) break;
          b.pcs.push(L);
        }
      if (AR10) {
        const L = n2col(col2n(b.pcs[b.pcs.length - 1]) + 1);
        if (!ownC.has(L) && !fr(L, 10)) {
          let t = false;
          for (let r = hr + 1; r <= hr + 400 && !t; r++) {
            const x = wsv && wsv[L + r];
            if (x && typeof x.v === 'string' && cl(x.v) && !/^Folder/i.test(x.v)) t = true;
          }
          if (t) b.pcs.push(L);
        }
      }
      /* ใบเช็คแบบใหม่: ช่องหมายเหตุที่ทีมพิมพ์ต่อท้าย (เช่น “กลับมาออนใหม่วันที่ 12”) ย้ายตามชื่อหนังด้วย */ b.bps =
        b.pcs.map((L, i) => (i ? modeF(sh, L, hr + 1, hr + 400, null, ST) : b.bp));
      b.hss = b.pcs.map((L, i) => (i ? sh.getS(L, hr) : b.hs[2]));
    }
  }
  const PCZ = Math.max(
    Math.max(...B.map(b => col2n(b.pc))) + 12,
    ...B.map(b => col2n(b.pcs[b.pcs.length - 1])),
  );
  const numAt = (c, r, k = 0) => {
    const f = k < 4 && fOf(c, r),
      m = f && f.replace(/^=/, '').match(/^\$?([A-Z]{1,3})\$?(\d+)\s*(?:([+-])\s*(\d+))?$/);
    if (m) {
      const v = numAt(m[1], +m[2], k + 1);
      if (v != null) return v + (m[3] ? (m[3] === '-' ? -1 : 1) * +m[4] : 0);
    }
    const x = wsv && wsv[c + r];
    return x && typeof x.v === 'number' ? x.v : null;
  }; /* แถบวันที่ที่เป็นสูตรอ้างช่องอื่น (เช่น =F45) อ่านจากช่องต้นทาง */
  const inB = n => n != null && n > 20000 && n < 80000,
    barD = r => {
      let v = numAt(b0.nc, r);
      if (!inB(v)) {
        v = null;
        if (!B.some(b => vOf(b.lc, r)))
          for (const b of B.slice(1)) {
            const n = numAt(b.nc, r);
            if (inB(n)) {
              v = n;
              break;
            }
          }
      }
      if (v == null) return null;
      const d = s2d(v),
        y = P.getUTCFullYear();
      return Math.abs(d.getUTCFullYear() - y) === 1
        ? new Date(Date.UTC(y, d.getUTCMonth(), d.getUTCDate()))
        : d;
    };
  const isBar = r => {
    if (barD(r)) return true;
    const st = sh.getS(b0.lc, r);
    return st != null && DKB.test(ST.fillOf(st));
  };
  const bars = [];
  let last = hr;
  for (let r = hr; r <= hr + SCAN; r++) {
    if (isBar(r)) {
      bars.push([r, barD(r)]);
      last = r;
      continue;
    }
    if (B.some(b => vOf(b.lc, r) || /^ออก(วันที่)?\s*\d/.test(vOf(b.nc, r)))) last = r;
  }
  const TRAIL =
    !bars.length ||
    bars[bars.length - 1][0] >=
      last; /* แถบสุดท้ายของชีตไม่มีหนังข้างใต้ (ใส่แถบวันถัดไปรอไว้) หรือยังไม่มีแถบเลย → ใส่แถบรอแบบเดิม; แถบสุดท้ายมีหนังแล้ว → คนทำใบนี้ไม่ใส่แถบรอ */
  let BV = 0,
    BF = 0;
  for (const [r] of bars)
    for (const b of B.slice(1)) {
      const x = wsv && wsv[b.nc + r];
      if (!x) continue;
      if (typeof x.f === 'string') BF++;
      else if (typeof x.v === 'number') BV++;
    }
  const BVAL =
    BV >
    BF; /* วันที่ในแถบของจอที่ 2 เป็นต้นไป: ชีตพิมพ์เป็นตัวเลข → ใส่ตัวเลข ถ้าไม่ใช่ใช้สูตร =C25 แบบเดิม */
  const F = bars.length && bars[0][0] === hr ? bars[0][1] : null;
  let S = hr,
    hist = false,
    frozen = false;
  const CONTM = two && o.r2 !== 'new' && !!F && +F === +P1;
  if (CONTM) {
    P = P1;
    two = false;
  } /* วันที่ 16 เป็นต้นไป: ทำต่อในรายการเดิม (แบบที่ทีมทำ) */
  if (two && F && +F === +P1) {
    /* มีรอบ 1–15 ของเดือนนี้อยู่แล้ว → ไม่แตะ แล้วเริ่มรอบ 16 ต่อข้างล่าง */
    frozen = true;
    const b2 = bars.find(([r, d]) => r > hr && d && +d >= +P);
    if (b2) {
      S = b2[0];
      hist = true;
    } else {
      let ln = hr;
      for (let r = hr + 1; r <= last; r++) if (B.some(b => vOf(b.lc, r))) ln = r;
      const lb = bars.find(([r]) => r > ln);
      S = lb ? lb[0] : ln + 1;
    }
  } else if (!F || +F === +P) hist = true; /* อ่านวันที่หัวตารางไม่ได้ → ถือว่ารอบเดียวกัน (ไม่ล้างของเดิม) */
  if (
    !CONTM &&
    !two &&
    F &&
    +F === +P1 &&
    bars.some(
      ([r, d], i) =>
        r > hr &&
        d &&
        d.getUTCMonth() === P1.getUTCMonth() &&
        d.getUTCDate() >= 16 &&
        (i < bars.length - 1 || ISG.test(ST.fillOf(sh.getS(b0.lc, r) || '0'))),
    )
  )
    return {
      skip: 'ชีตนี้มีรอบ 16–สิ้นเดือนแล้ว แต่วันที่ทำงานอยู่ในรอบ 1–15 — ไม่ได้แก้ชีตนี้ ตรวจวันที่ทำงานอีกครั้ง',
    };
  const KB = B.filter(
    b => b.kind === 'keep',
  ); /* จอในชีตเดียวกันใช้แถวและแถบวันที่ร่วมกัน: จอที่ไม่มีใน Playlist ต้องคงรายการเดิมไว้ในแถวที่ตรงกัน */
  if (KB.length && !hist)
    return {
      skip:
        'ไม่พบจอ ' +
        KB.map(b => b.scr.c).join(', ') +
        ' ใน Playlist ที่ใส่มา และชีตนี้กำลังเริ่มรอบใหม่' +
        ' — ไม่ได้แก้ชีตนี้ (จอในชีตเดียวกันใช้แถบวันที่ร่วมกัน) ใส่ Playlist ที่มีจอนี้ด้วย เช่นไฟล์ล่าสุดของวันก่อน แล้วเติมใหม่',
    };
  const Z = Math.max(last + 2, S + 400),
    lim = Math.min(+addD(W, +o.days || 0), +E),
    H = lim; /* แบบที่ทำด้วยมือ: ใส่หนังที่เริ่มถึงวันนี้ และ “ออกวันที่” ถึงวันนี้ ไม่ใส่ล่วงหน้า */
  const exitOf = (n, ref) => {
    let d = new Date(Date.UTC(ref.getUTCFullYear(), ref.getUTCMonth(), n));
    if (n < ref.getUTCDate()) d = new Date(Date.UTC(ref.getUTCFullYear(), ref.getUTCMonth() + 1, n));
    return +addD(d, -1);
  };
  /* เก็บวันที่ถ่าย/ตัวเลขที่ใส่ไว้ในช่องหลังชื่อหนัง (ทุกจอของบล็อก) — ย้ายตามชื่อหนังเฉพาะในเดือนเดียวกัน ขึ้นเดือนใหม่ล้างทุกช่อง */
  let moved = 0,
    dropped = 0,
    cleared = 0,
    kept = 0,
    merged = 0,
    gone = 0,
    today = 0,
    nost = 0,
    renew = 0,
    back = 0,
    fdrop = 0;
  /* BKJ: หมายเหตุ “ขึ้นเฉพาะ BKJ2-24,…” / “ยกเว้นจอ BKJ-12,…” / “ไม่ออน BKJ2-35” เทียบกับเลขจอที่หัวคอลัมน์ของใบเช็ค */
  const jamNums = b => {
    const s = new Set(),
      c0 = col2n(b.lc);
    for (let r = 8; r <= 14; r++) {
      const x = wsv && wsv[b.lc + r];
      if (!(x && /BANGKOK|BKJ|JAM|METRO/i.test(cl(x.v)))) continue;
      for (let c = c0 + 1; c < c0 + 60; c++) {
        const y = wsv[n2col(c) + r],
          v = y ? cl(y.v) : '';
        if (!/^\d{1,3}$/.test(v)) break;
        s.add(+v);
      }
      break;
    }
    return s;
  };
  const nameSkip = (ad, nums) => {
    const m = String(ad || '').match(/Screen\s*(\d+(?:\s*(?:and|&|,|_)\s*\d+)*)/i);
    if (!m || !nums || !nums.size) return false;
    const s = new Set(m[1].split(/\D+/).filter(Boolean).map(Number));
    return ![...nums].some(x => s.has(x));
  };
  const locSkip = (loc, nums) => {
    if (!loc || !nums || !nums.size) return false;
    const inc = new Set(),
      exc = new Set();
    for (const m of String(loc).matchAll(/(!)?LEDMETRO_?(\d+)/gi)) (m[1] ? exc : inc).add(+m[2]);
    return (
      (inc.size > 0 && ![...nums].some(x => inc.has(x))) || (exc.size > 0 && [...nums].every(x => exc.has(x)))
    );
  };
  const jamSkip = (note, nums) => {
    if (!nums || !nums.size) return false;
    const n = String(note || ''),
      L = t => {
        const s = new Set();
        for (const m of t.matchAll(/BKJ\s*\d?\s*-\s*(\d{1,3})/gi)) s.add(+m[1]);
        return s;
      };
    let m = n.match(/(?:ขึ้น|ออน)เ?ฉพาะ(.*)/);
    if (m) {
      const s = L(m[1]);
      return s.size > 0 && ![...nums].some(x => s.has(x));
    }
    m = n.match(/(?:ยกเว้น(?:จอ)?|ไม่ออน(?:จอ)?)(.*)/);
    if (m) {
      const s = L(m[1].split(/ชดเชย/)[0]);
      return s.size > 0 && [...nums].every(x => s.has(x));
    }
    return false;
  };
  for (const b of B) {
    b.cnt = new Map();
    b.hist = [];
    let g = +P;
    for (let r = S + 1; r <= last; r++) {
      if (isBar(r)) {
        const d = barD(r);
        if (d) g = Math.max(+d, +P);
        continue;
      }
      const nm = vOf(b.lc, r);
      if (!nm || /^ออก(วันที่)?\s*\d/.test(nm)) continue;
      const fx = FASTN.test(nm);
      let t = null;
      for (let i = 0; i < b.pcs.length; i++) {
        const k = sh.take(b.pcs[i], r);
        if (k) {
          sh.setS(b.pcs[i], r, b.bps[i]);
          if (hist && !fx) (t || (t = b.pcs.map(() => null)))[i] = k;
          else if (!hist) cleared++;
        }
      }
      if (t) {
        if (!b.cnt.has(nm)) b.cnt.set(nm, []);
        b.cnt.get(nm).push(t);
      }
      if (fx) {
        if (hist) fdrop++;
        continue;
      } /* แถว #FAST_ADS ที่มีอยู่ในใบเช็ค → ลบออก */
      if (hist) {
        const m = vOf(b.nc, r).match(/^ออก(?:วันที่)?\s*(\d{1,2})/);
        b.hist.push({ nm, g, ned: m ? exitOf(+m[1], P) : 0, raw: m ? vOf(b.nc, r) : '' });
      }
    }
  }
  if (!hist && !frozen)
    for (const b of B)
      for (let r = Math.min(FIRST, S + 1); r <= last; r++)
        if (r > S || vOf(b.lc, r))
          for (const L of b.pcs)
            if (sh.take(L, r))
              cleared++; /* ขึ้นเดือนใหม่: ล้างวันที่ถ่ายที่เหลือด้วย — แถว # เหนือหัววันที่ (เช่น #BANGKOK JAM_ETA, #EDA) และตัวเลขที่ค้างในแถวที่ไม่มีชื่อหนัง (สูตรไม่แตะ) */
  /* รวมรายการ: ของเดิมในชีต + Playlist วันนี้ */
  for (const b of B) {
    const offEd = new Map(),
      pl = [],
      ent = [];
    b.later = [];
    b.ratio = [];
    b.sp = b.rows.filter(r => r.sp);
    {
      const cap = b.scr.pd
        ? b.scr.pd - 864e5
        : Infinity; /* อยู่ในรายการถอดของ Playlist วันไหน = ออกไม่เกินวันนั้น */
      for (const r of b.scr.rows)
        if (r.off && r.ad) {
          const e0 = parseD(r.e),
            ed = Math.min(e0 && +e0 >= +P ? +e0 : Infinity, cap);
          if (!isFinite(ed) || ed < +P) continue;
          const k =
            (/^Content/i.test(r.ty || '') ? '#' : '') +
            (b.kind === 'jam' ? r.ad.replace(/^BKJ[12]?_/i, '') : r.ad);
          offEd.set(k, Math.max(offEd.get(k) || 0, ed));
        }
    } /* วันหมดของหนังที่ถอดแล้ว (ใช้ทำ “ออกวันที่”) */
    const jn = b.kind === 'jam' || b.kind === 'metro' ? jamNums(b) : null,
      fam = n => String(n).replace(/^#/, '').trim().toUpperCase(),
      hb0 = b.hist.length > 0;
    for (const r of b.rows) {
      if (r.sp) continue;
      if (jn && (b.kind === 'jam' ? jamSkip(r.note, jn) || nameSkip(r.ad, jn) : locSkip(r.loc, jn)))
        continue; /* ยังอยู่ใน Playlist = ยังออนอยู่ แม้วันหมดในไฟล์ผ่านไปแล้ว */
      let g = r.sd && +r.sd > +P ? +r.sd : +P;
      const stl =
        CONTM &&
        dd >= +P2 &&
        /ต่อสัญญา/.test(r.st || '') &&
        r.ed &&
        +r.ed === +P2 - 864e5 &&
        !(r.sd && +r.sd >= +P2);
      if (stl)
        g = +P2; /* ต่อสัญญาวันที่ 16 แต่ Playlist ยังลงวันที่สัญญาเดิม (หมด 15) → ถือเป็นสัญญาใหม่เริ่ม 16 */
      const mid = String(r.note || '').match(
          /(?:ออนแอร์|ขึ้น)ระหว่างวัน\s*(?:วันที่)?\s*(\d{1,2})\/(\d{1,2})\/(\d{2,4})/,
        ),
        md = mid ? +new Date(Date.UTC(yy(mid[3]), +mid[2] - 1, +mid[1])) : 0;
      if (md >= +P && md <= dd)
        g =
          CONTM && dd >= +P2 && md < +P2 && !b.hist.some(h => fam(h.nm) === fam(r.ad))
            ? +P2
            : md; /* หมายเหตุ “ออนแอร์ระหว่างวันวันที่ …” → ลงใต้แถบวันนั้น — ครึ่งเดือนหลัง: หนังใหม่ที่หมายเหตุลงวันที่ 1–15 ลงใต้แถบ 16 (รอบ 1–15 ปิดแล้ว แบบที่ทีมทำ) */
      else if (NEWST.test(r.st || '') && g !== dd) {
        if (g > dd) g = Math.max(dd, +P);
        else if (hb0 && g > +P && !b.hist.some(h => h.g === g && fam(h.nm) === fam(r.ad))) g = dd;
      } /* “ขึ้นใหม่” = ขึ้นวันนี้ */
      if (g > lim) {
        if (!b.hist.some(h => fam(h.nm) === fam(r.ad))) b.later.push(r);
        continue;
      }
      const rnw =
        CONTM &&
        dd === +P2 &&
        (NEWST.test(r.st || '') || CONT.test(r.st || '')) &&
        ((r.sd && +r.sd === +P2) ||
          stl); /* วันที่ 16: ต่อสัญญา/ขึ้นใหม่ที่เริ่ม 16 → แถวใหม่ใต้ 16 แถวเดิมใส่ “ออก 16” */
      pl.push({
        nm: (r.hash ? '#' : '') + r.ad,
        hash: r.hash,
        g,
        ed: r.ed && +r.ed >= dd ? +r.ed : Infinity,
        st: r.st || '',
        rnw,
        note: r.note || '',
      });
    }
    pl.sort((x, y) => x.g - y.g);
    for (const h of b.hist) {
      let x =
        h.ned || offEd.get(h.nm) || (!h.nm.startsWith('#') && !offEd.has(h.nm) && offEd.get('#' + h.nm)) || 0;
      if (!h.ned && x && h.g >= +P2 && x < h.g)
        x = Math.max(
          h.g,
          (b.scr.pd || dd) - 864e5,
        ); /* ถอดแล้วแต่วันหมดยังเป็นของสัญญาเดิม (ก่อนแถบนี้) → ออกวันที่ของ Playlist ที่ถอด */
      ent.push({ nm: h.nm, hash: h.nm.startsWith('#'), g: h.g, ed: x, raw: h.raw, h: true, m: false });
    }
    /* หนังที่ยังไม่เคยลงในชีต: เริ่มก่อน/ในวันแรกของรอบ → ลงใต้แถบ “วันนี้” (ถ้าชีตมีรายการของรอบนี้อยู่แล้ว) ไม่เอาไปรวมกับวันแรก */
    const hb = b.hist.length > 0,
      gate =
        hb &&
        dd > +P &&
        b.all.some(
          r => r.st,
        ); /* ตั้งแต่วันที่ 2 ของรอบ: หนังที่ยังไม่มีในใบเช็ค เพิ่มเฉพาะสถานะ “ขึ้นใหม่…” (หรือเริ่มวันนี้/ล่วงหน้า) */
    for (const p of pl) {
      if (p.hash && !ent.some(e => e.nm === p.nm) && ent.some(e => e.h && e.nm === p.nm.slice(1))) {
        p.nm = p.nm.slice(1);
        p.hash = false;
      } /* หนัง Content ที่ใบเช็คเขียนไว้แล้วแบบไม่มี # (เช่น MD_…) → แถวเดิม ไม่เพิ่ม #MD_… ซ้ำ */
      const c = ent.filter(e => e.nm === p.nm),
        tg = p.g > +P ? p.g : hb ? dd : +P;
      {
        const hs = c.filter(e => e.h && !e.m);
        if (hb && hs.length && NEWST.test(p.st) && !p.rnw) {
          const last = hs.reduce((a, e) => (e.g > a.g ? e : a));
          if (
            CONTM &&
            dd >= +P2 &&
            last.g < +P2 &&
            last.ed &&
            last.ed < dd - 864e5 &&
            !c.some(e => e.g === dd)
          ) {
            ent.push({ nm: p.nm, hash: p.hash, g: dd, ed: p.ed, m: true });
            back++;
            continue;
          }
        } /* ครึ่งเดือนหลัง: หนังที่ลงไว้ในรอบ 1–15 ออกไปแล้วกลับมาขึ้นใหม่ → แถวใหม่ใต้วันนี้ (ในรอบ 1–15 ใช้แถวเดิม) */
        const op = hs.filter(e => !(e.ed && e.ed < dd));
        if (!p.rnw && op.length === 1 && hs.length > 1) {
          op[0].ed = p.ed;
          op[0].m = true;
          continue;
        }
      } /* มีหลายแถว: ใช้แถวล่าสุดที่ยังไม่ออก */
      if (p.rnw) {
        for (const e of c)
          if (e.g < +P2) {
            e.m = true;
            if (!o.k16 && !(e.ed && e.ed < +P2)) e.ed = +P2 - 864e5;
          }
        const t = c.find(e => e.g === +P2);
        if (t) {
          t.ed = p.ed;
          t.m = true;
        } else {
          ent.push({ nm: p.nm, hash: p.hash, g: +P2, ed: p.ed, m: true });
          renew++;
        }
        continue;
      }
      let t = c.find(e => e.g === p.g && !e.m) || (o.dedupe && c.find(e => e.g === p.g && e.m));
      if (t) {
        t.ed = t.m ? Math.max(t.ed, p.ed) : p.ed;
        t.m = true;
        continue;
      }
      if (p.g <= dd) {
        t = c.find(e => e.h && !e.m && e.g <= dd && e.g >= p.g); /* ลงไว้แล้วในวันก่อนหน้า → อยู่ที่เดิม */
        if (t) {
          t.ed = p.ed;
          t.m = true;
          continue;
        }
        if (o.dedupe) {
          t = c.find(e => e.m && e.g > p.g && e.g <= dd);
          if (t) {
            t.ed = Math.max(t.ed || 0, p.ed);
            continue;
          }
        }
      } /* Playlist มีหนังเรื่องเดียวกันหลายแถว (วันเริ่มสัญญาต่างกัน) และลงไว้แล้วใต้แถบที่หลังกว่า → ใช้แถวนั้น ไม่เพิ่มแถวซ้ำใต้วันเริ่มสัญญาที่เก่ากว่า */
      t = c.find(e => e.h && !e.m && e.g > dd); /* เคยวางล่วงหน้าไว้ แต่วันเริ่มเปลี่ยน → ย้ายตาม Playlist */
      if (t) {
        t.g = tg;
        t.ed = p.ed;
        t.m = true;
        continue;
      }
      t = c.filter(e => e.g <= tg).pop(); /* ไฟล์เดิมต่อเนื่อง หรือยังออนอยู่ → แถวเดิม ไม่เพิ่มซ้ำ */
      if (t && (CONT.test(p.st) || !t.ed || t.ed >= p.g || t.g >= +P2 === tg >= +P2)) {
        t.ed = t.ed && t.ed < p.g ? p.ed : Math.max(t.ed || 0, p.ed);
        t.m = true;
        merged++;
        continue;
      } /* ครึ่งเดือนเดียวกัน: หนังที่ออกไปแล้วกลับมาขึ้นอีก ใช้แถวเดิม (แบบที่ทีมทำ — ใบเช็คมือ ก.ย. ไม่มีหนังซ้ำ 2 แถวในครึ่งเดือนเดียวกัน) */
      if (gate && !NEWST.test(p.st) && p.g === +P) {
        nost++;
        continue;
      } /* สัญญาเริ่มก่อนรอบนี้ แต่ไม่เคยลงและไม่ได้ “ขึ้นใหม่” → ไม่เพิ่ม; ถ้าเริ่มในรอบนี้ ลงตามวันเริ่ม (กันพลาดเมื่อข้ามวัน) */
      if (tg !== p.g) today++;
      ent.push({ nm: p.nm, hash: p.hash, g: tg, ed: p.ed, m: true });
    }
    b.ent =
      b.kind === 'keep'
        ? ent
        : ent.filter(e => {
            if (e.m) return true;
            if (e.g > dd) {
              gone++;
              return false;
            }
            kept++;
            return true;
          });
  }
  if (CONTM)
    for (const b of B) {
      if (b.kind === 'keep') continue;
      const nw = new Map();
      for (const e of b.ent)
        if (e.g === +P2 && e.ed && isFinite(e.ed) && +addD(new Date(e.ed), 1) <= H) nw.set(e.nm, e.ed);
      for (const e of b.ent) if (e.g < +P2 && e.ed === +P2 - 864e5 && nw.has(e.nm)) e.ed = nw.get(e.nm);
    } /* ต่อสัญญาวันที่ 16: แถวเดิม “ออกวันที่ 16” ไว้ก่อน พอหนังถอดจริง ใส่วันถอดจริงทั้งแถวเดิมและแถวใต้ 16 (แบบที่ทีมทำ) */
  const seen = new Set();
  for (const b of B) for (const e of b.ent) if (e.g > +P) seen.add(e.g);
  const dates = [...seen].sort((x, y) => x - y);
  const srt = L => L.sort((x, y) => y.hash - x.hash || cmpN(x.nm, y.nm));
  for (const b of B) {
    b.main = srt(b.ent.filter(e => e.g === +P));
    b.grp = new Map(dates.map(k => [k, srt(b.ent.filter(e => e.g === k))]));
  }
  const OSH = { s: false, p: false };
  {
    let s = 0,
      l = 0,
      n1 = 0,
      p = 0;
    for (const b of B)
      for (let r = hr + 1; r <= last; r++) {
        const t = vOf(b.nc, r),
          m = t.match(/^ออก(วันที่)?\s*(\d{1,2})/);
        if (!m) continue;
        m[1] ? l++ : s++;
        if (+m[2] < 10) {
          n1++;
          if (/^0/.test(m[2])) p++;
        }
      }
    const lg = B.some(b => /^(Air Traveller|The 20|Rama9 Connect)$/.test((b.scr || {}).m || '')),
      la = B.every(b => ((b.scr || {}).m || '') === 'LA PlanB'),
      w = o.osw || { s: 0, l: 0 };
    if (AR10) {
      OSH.s = s + l ? s > l : la ? w.s > w.l : !lg;
      OSH.p = n1 ? p * 2 > n1 : lg && !OSH.s;
    } else OSH.s = s + l ? s > l : w.s > w.l;
  } /* ยังไม่มีหมายเหตุให้ดูแบบ (เช่นต้นเดือน): ใบ Air Traveller / THE 20 / Rama9 ทีมเขียน “ออกวันที่ 04” ใบ MRT / Gateway X / CCTV เขียน “ออก 4” */
  const note = (b, r, row) => {
    if (noNotes || !r.ed || !isFinite(r.ed)) return;
    const nd = +addD(new Date(r.ed), 1);
    if (nd > +P && nd <= H) {
      {
        const dn = new Date(nd).getUTCDate();
        sh.setText(
          b.nc,
          row,
          AR10
            ? (OSH.s ? 'ออก ' : 'ออกวันที่ ') + (OSH.p && dn < 10 ? '0' + dn : dn)
            : (OSH.s ? 'ออก ' : 'ออกวันที่ ') + dn,
        );
      }
      sh.setS(b.nc, row, ST.derive(b.bn, 'red', { color: 'FFFF0000', right: true }));
    }
  };
  const put = (b, r, row) => {
    const q = b.cnt.get(r.nm);
    if (q && q.length)
      q.shift().forEach((k, i) => {
        if (k) {
          sh.give(b.pcs[i], row, k);
          moved++;
        }
      });
    sh.setText(b.lc, row, r.nm);
    if (r.hash) sh.setS(b.lc, row, ST.derive(b.bl, 'yel', { fill: 'FFFFFF00' }));
    else sh.setS(b.lc, row, b.bl);
    if (b.kind !== 'keep') note(b, r, row);
    else if (r.raw) {
      sh.setText(b.nc, row, r.raw);
      sh.setS(b.nc, row, ST.derive(b.bn, 'red', { color: 'FFFF0000', right: true }));
    }
  };
  for (const c of [o.col1, o.col2]) if (c) MARK.x.add(String(c).toUpperCase().slice(-6));
  /* แถบรอบ 2 ที่ชีตมีอยู่แล้ว (สีต่างจากหัวตาราง): จำสไตล์ไว้ก่อนล้างแถว */
  const E2 =
      (two || CONTM) && !o.col2
        ? bars.find(
            ([r, d]) =>
              r > hr &&
              d &&
              +d >= +(CONTM ? P2 : P) &&
              +d <= +E &&
              sh.getS(b0.lc, r) != null &&
              sh.getS(b0.lc, r) !== b0.hs[1],
          )
        : null,
    SNAP = {};
  if (E2) {
    const a = Math.min(...B.map(b => col2n(b.nc))),
      z = PCZ;
    for (let c = a; c <= z; c++) {
      const L = n2col(c);
      SNAP[L] = sh.getS(L, E2[0]);
    }
  }
  const BS = new Set();
  for (const [br] of bars)
    for (const b of B)
      for (const c of [b.nc, ...b.pcs]) {
        const x = sh.getS(c, br);
        if (x != null && x !== b.bp && !b.bps.includes(x) && x !== b.bn && x !== b.be && x !== b.bl)
          BS.add(x);
      } /* สไตล์ของแถบเดิมทุกสี (รวมสีธีม) */
  for (const b of B) {
    const beNo = b.be == null || !ST.hasBorder(b.be);
    for (let r = S + 1; r <= Z; r++) {
      const ps0 = sh.getS(b.pc, r);
      {
        const ps = ps0;
        if (ps != null && (ps === b.hs[2] || MARK.test(ST.fillOf(ps)) || BS.has(ps))) sh.setS(b.pc, r, b.bp);
      }
      for (let i = 1; i < b.pcs.length; i++) {
        const ps = sh.getS(b.pcs[i], r);
        if (ps != null && (ps === b.hss[i] || MARK.test(ST.fillOf(ps)) || BS.has(ps)))
          sh.setS(b.pcs[i], r, b.bps[i]);
      }
      sh.clear(b.lc, r);
      sh.clear(b.nc, r);
      sh.setS(b.lc, r, beNo && ps0 != null && ST.hasBorder(ps0) && !ST.fillOf(ps0) ? b.bl : b.be);
      /* ช่องว่างไม่มีเส้น แต่ช่องนับข้างๆ เป็นตาราง (มีเส้น ไม่มีสี) → ใช้เส้นแบบแถวหนัง ให้ตารางต่อเนื่อง */ {
        const ns = sh.getS(b.nc, r);
        if (b.bn != null) sh.setS(b.nc, r, b.bn);
        else if (ns != null && (MARK.test(ST.fillOf(ns)) || BS.has(ns))) sh.setS(b.nc, r, null);
      }
    }
    if (b.kind === 'metro' && hr > FIRST && !frozen)
      for (let i = 0; i < hr - FIRST; i++) {
        const r = b.sp[i];
        if (r)
          (sh.setText(b.lc, FIRST + i, '#' + r.ad),
            sh.setS(b.lc, FIRST + i, ST.derive(b.bl, 'yel', { fill: 'FFFFFF00' })));
      }
  }
  /* แถบหัววันที่ให้ยาวเท่าแถบเดิมของชีต (เช่น Metro ที่มีช่องนับหลายจอ) */
  const c0 = Math.min(...B.map(b => col2n(b.nc))),
    c1 = PCZ,
    own = new Set(B.flatMap(b => [b.nc, b.lc, ...b.pcs])),
    HX = [];
  for (let c = c0; c <= c1; c++) {
    const L = n2col(c),
      st = sh.getS(L, hr);
    if (!own.has(L) && st != null && DKB.test(ST.fillOf(st))) HX.push([L, st]);
  }
  for (const [L] of HX) {
    const base = modeF(sh, L, hr + 1, hr + 400, null, ST);
    for (let r = S + 1; r <= Z; r++) {
      const st = sh.getS(L, r);
      if (st != null && (MARK.test(ST.fillOf(st)) || BS.has(st))) sh.setS(L, r, base);
    }
  }
  /* สีแถบ: เลือกเองได้ ถ้าไม่เลือก รอบ 1 ใช้สีหัวตารางเดิม รอบ 2 ใช้สีของแถบรอบ 2 ที่ชีตมีอยู่แล้ว ถ้ายังไม่มีใช้ฟ้า 4472C4 */
  const R2 = {},
    KEEP2 = two && !o.col2 && F && +F >= +P;
  if ((two || CONTM) && !o.col2 && !KEEP2) {
    const e = E2;
    if (e) {
      for (const b of B) for (const c of [b.nc, b.lc, ...b.pcs]) R2[c] = SNAP[c];
      for (const [L] of HX) R2[L] = SNAP[L];
    }
  }
  const C2 = o.col2 || 'FF4472C4',
    fx = (st, c, k) =>
      String(ST.fillOf(st) || '')
        .toUpperCase()
        .slice(-6) === String(c).toUpperCase().slice(-6)
        ? st
        : ST.derive(st, k + c, { fill: c });
  const hsx = (st, L, r2) => {
    if (st == null) return st;
    if (two || r2) {
      if (KEEP2) return st;
      return !o.col2 && L && R2[L] != null ? R2[L] : fx(st, C2, 'c2');
    }
    if (o.col1) return fx(st, o.col1, 'c1');
    return ISG.test(ST.fillOf(st)) ? ST.derive(st, 'nvy', { fill: NAVY }) : st;
  };
  const HXA = [...HX, ...B.flatMap(b => b.pcs.slice(1).map((L, i) => [L, b.hss[i + 1]]))].sort(
    (x, y) => col2n(x[0]) - col2n(y[0]),
  ); /* แถบวันที่คลุมทุกช่องวันที่ถ่ายของบล็อก (และคอลัมน์หัวสีเข้มอื่นที่เป็นส่วนของแถบ) */
  const hdr = (row, d) => {
    const r2 = CONTM && +d >= +P2;
    for (const [L, st] of HXA) {
      sh.clear(L, row);
      if (st != null) sh.setS(L, row, hsx(st, L, r2));
    }
    B.forEach((b, i) => {
      sh.clear(b.lc, row);
      if (i === 0 || BVAL) sh.setNum(b.nc, row, serial(d));
      else sh.setF(b.nc, row, B[i - 1].nc + row, serial(d));
      if (b.hs[0] != null) sh.setS(b.nc, row, hsx(b.hs[0], b.nc, r2));
      if (b.hs[1] != null) sh.setS(b.lc, row, hsx(b.hs[1], b.lc, r2));
      sh.clear(b.pc, row);
      if (b.hs[2] != null) sh.setS(b.pc, row, hsx(b.hs[2], b.pc, r2));
    });
  };
  hdr(S, P);
  const pad = (b, a, z) => {
    for (let row = a; row <= z; row++) {
      sh.setS(b.lc, row, ST.derive(b.bl, 'pad' + PADC, { fill: PADC, border: true }));
      b.pcs.forEach((L, i) =>
        sh.setS(
          L,
          row,
          ST.derive(b.bps[i] != null ? b.bps[i] : '0', 'pad' + PADC, { fill: PADC, border: true }),
        ),
      );
    }
  }; /* แถบสีฟ้าในช่องว่างต่อจากรายการที่สั้นกว่า */
  const mainLen = Math.max(0, ...B.map(b => b.main.length));
  for (const b of B) {
    b.main.forEach((r, i) => put(b, r, S + 1 + i));
    pad(b, S + 1 + b.main.length, S + mainLen);
  }
  let cur = S + 1 + mainLen;
  for (const k of dates) {
    const n = Math.max(...B.map(b => (b.grp.get(k) || []).length));
    hdr(cur, new Date(k));
    for (const b of B) {
      const g = b.grp.get(k) || [];
      g.forEach((r, i) => put(b, r, cur + 1 + i));
      pad(b, cur + 1 + g.length, cur + n);
    }
    cur += 1 + n;
  }
  if (TRAIL)
    hdr(
      cur,
      addD(new Date(Math.max(dates.length ? dates[dates.length - 1] : +P, dd)), 1),
    ); /* แถบว่างท้ายตาราง = วันถัดจากวันที่เติม (เติมวันที่ 2 → แถบ 3) */ /* ปิดท้ายด้วยแถบวันถัดไปเสมอ ไว้ใส่หนังของวันพรุ่งนี้ */
  if (o.off) {
    let row = cur + 2;
    for (const b of B) {
      let q = row;
      for (const r of b.rows.filter(x => x.off)) sh.setText(b.lc, q++, r.ad);
    }
  }
  for (const b of B) for (const q of b.cnt.values()) for (const t of q) dropped += t.filter(Boolean).length;
  return {
    hr,
    S,
    H,
    P,
    E,
    two,
    contm: CONTM,
    trail: TRAIL,
    bval: BVAL,
    fold: B.map(b => ({ lc: b.lc, items: (b.ent || []).map(e => [e.nm, e.g]) })),
    fdrop,
    keep: KB.map(b => b.scr.c).join(', '),
    renew,
    back,
    frozen,
    noNotes,
    hist,
    dates: dates.length,
    later: B.reduce((n, b) => n + b.later.length, 0),
    laterL: B.map(b => [b.scr.c, b.later.map(r => r.ad + (r.s ? ' (เริ่ม ' + r.s + ')' : ''))]).filter(
      x => x[1].length,
    ),
    ratioL: B.map(b => [b.scr.c, b.ratio]).filter(x => x[1].length),
    moved,
    dropped,
    cleared,
    kept,
    merged,
    gone,
    today,
    nost,
    W: dd,
    end: cur,
  };
}

/* ใบเช็คแบบใหม่ (3.4): แถวสูตรวิเคราะห์ (AVERAGE) แบบเดิมอยู่แถว 9 • แบบใหม่ (CCTV / MRT / Gateway X / Rama9 Connect / Air Traveller / THE 20) อยู่แถว 10 และตารางเริ่มแถว 17 หรือ 23
   ช่องสรุปรวมของสถานี MRT (=AVERAGE(E10:F10)) ไม่ใช่คอลัมน์ของจอ */
const BLKAVG = f => /^=?\s*AVERAGE/i.test(f || '') && !/AVERAGE\(\s*\$?[A-Z]+\$?10\s*:/i.test(f);
const avRowOf = ws => {
  let r10 = false;
  for (const k of Object.keys(ws)) {
    const m = k.match(/^[A-Z]+(9|10)$/);
    if (!m || !ws[k].f || !/^AVERAGE/.test(ws[k].f)) continue;
    if (m[1] === '9') return 9;
    if (BLKAVG(ws[k].f)) r10 = true;
  }
  return r10 ? 10 : 9;
};
const avCols = ws => {
  const R = avRowOf(ws),
    re = new RegExp('^([A-Z]+)' + R + '$'),
    av = [];
  for (const k of Object.keys(ws)) {
    const m = k.match(re);
    if (m && ws[k].f && (R === 9 ? /^AVERAGE/.test(ws[k].f) : BLKAVG(ws[k].f))) av.push(col2n(m[1]));
  }
  return av;
};
/* จับคู่ชีตแบบใหม่กับจอ: สถานี MRT ส่วนต่อขยาย 20 สถานี (สายสีน้ำเงินฝั่งท่าพระ–หลักสอง / เตาปูน–ท่าพระ) ใช้รายการ “ส่วนต่อขยาย” สถานีอื่นใช้ “18 สถานีเดิม” */
const MRTX = new Set(
  'ท่าพระ จรัญ13 จรัญฯ13 จรัญสนิทวงศ์13 ไฟฉาย บางขุนนนท์ บางยี่ขัน สิรินธร บางพลัด บางอ้อ บางโพ เตาปูน วัดมังกร สามยอด สนามไชย อิสรภาพ บางไผ่ บางหว้า เพชรเกษม48 ภาษีเจริญ บางแค หลักสอง'
    .split(' ')
    .map(s => nrm(s)),
);
const txt10 = (ws, cn) => {
    const g = a => cl((ws[a] || {}).v),
      T = [];
    for (const c of [cn - 1, cn]) for (let r = 11; r <= 14; r++) T.push(g(n2col(c) + r));
    return T;
  },
  ttl10 = (ws, sn) => nrm(sn + ' ' + ['A1', 'B1', 'C1', 'D1'].map(a => cl((ws[a] || {}).v)).join(' ')),
  lab10 = (ws, sn, cn) => txt10(ws, cn).filter(t => t && !/^CCTV$/i.test(t))[0] || sn,
  kind10 = (ws, sn, cn) => {
    const t = ttl10(ws, sn),
      b = nrm(txt10(ws, cn).join(' '));
    return /MRT/.test(t)
      ? 'Digital MRT'
      : /AIRTRAVELLERSERIES/.test(t + b)
        ? 'Air Traveller'
        : /THE20/.test(t)
          ? 'The 20'
          : /RAMA9CONNECT/.test(t)
            ? 'Rama9 Connect'
            : /GATEWAYX/.test(t)
              ? 'Gateway X'
              : /LAPLANB/.test(t)
                ? 'LA PlanB'
                : /AIRTRAVELL/.test(b)
                  ? 'Signature Max'
                  : '';
  };
function screen10(ws, sn, cn) {
  const b = nrm(txt10(ws, cn).join(' ')),
    k = kind10(ws, sn, cn),
    one = (m, re) => {
      const L = DB.filter(s => s.m === m && (!re || re.test(s.c)));
      return L.length === 1 || (!re && L.length) ? L[0] : null;
    };
  if (k === 'Digital MRT') return one(k, MRTX.has(nrm(sn)) ? /ต่อขยาย|ใหม่/ : /เดิม/);
  if (k === 'Air Traveller')
    return one(k, /AIRPORTIN|TOAIRPORT/.test(b) ? /IN$/ : /CITYOUT|TOCITY/.test(b) ? /OUT$/ : /^$/);
  if (k === 'The 20' || k === 'Rama9 Connect' || k === 'Gateway X' || k === 'LA PlanB') return one(k);
  return null;
}
/* ----- โฟลเดอร์แยกรูป (ขั้นที่ 4): แยกรูป\Pic @ สื่อ\จอ\10 OCT 26\1-15 หรือ 16-31\ชื่อหนัง
   รอบ 1–15 = หนังใต้แถบวันที่ 1–15 ของใบเช็ค • รอบ 16–สิ้นเดือน = หนังใต้แถบ 16 ขึ้นไป (แบบที่ทีมทำ) • สร้างเฉพาะที่ยังไม่มี ไม่ลบ ไม่ย้าย ไม่เปลี่ยนชื่อ
   จับคู่จอกับโฟลเดอร์: DGT / ID ดูรหัสหน้าชื่อโฟลเดอร์ • BKJ1_04 / BKJ2_20 ดูเลขจอในหัวคอลัมน์ • METRO_07 ดู 007 • จออื่น (So Bangkok, Twin Tube, The Big Y, BASE) ดูชื่อ ----- */
const PMON = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'],
  PMOD = ['#Module ขาว', '#Module ดำ', '#Module สี'],
  p2 = n => String(n).padStart(2, '0'),
  fsafe = n =>
    String(n)
      .replace(/[<>:"\/\\|?*\u0000-\u001f]/g, '_')
      .replace(/[\s.]+$/, '')
      .trim(),
  isMonD = n => /^\d{1,2}\s*[A-Za-z]{3}[A-Za-z]*\s*\d{2,4}$/.test(String(n).trim()),
  pck = c => String(c).replace(/-/g, '').toUpperCase(),
  pnm = t => nrm(t).replace(/^PIC/, '').replace(/^\d+/, '');
function picKey(n) {
  const s = String(n).trim(),
    c = codeOf(s);
  let m;
  if (c) return { code: pck(c) };
  if ((m = s.match(/^BKJ\s*([12])\s*_\s*(\d+)/i))) return { jam: m[1] + '_' + +m[2] };
  if ((m = s.match(/^METRO\s*_?\s*(\d+)/i))) return { metro: +m[1] };
  return { nm: pnm(s) };
}
function picId(ws, sn, lc) {
  const av = avCols(ws),
    cs = [];
  for (let c = col2n(lc) + 1; av.includes(c); c++) cs.push(c);
  const g = a => cl((ws[a] || {}).v),
    T = [],
    nums = [];
  let code = '';
  for (const c of cs) {
    const L = n2col(c);
    for (const r of [10, 11, 12]) {
      const t = g(L + r);
      if (!t) continue;
      T.push(t);
      if (!code) code = codeOf(t);
    }
    const n = g(L + '10') || g(L + '11');
    if (/^\d{1,3}$/.test(n)) nums.push(+n);
  }
  for (const r of [10, 11, 12]) {
    const t = g(lc + r);
    if (t) T.push(t);
  }
  const U = String(sn)
      .toUpperCase()
      .replace(/[^A-Z0-9]/g, ''),
    jm = !code && U.match(/(?:JAMS?|BKJ)([12])/);
  return {
    code: code ? pck(code) : '',
    jam: jm ? jm[1] : '',
    metro: !code && !jm && /METRO/.test(U),
    nums,
    labels: [...new Set(T.map(pnm))].filter(t => t.length >= 4),
    raw: T.find(t => codeOf(t)) || T[0] || '',
    media: mediaOf(sn) || '',
  };
}
const picDirs = async h => {
  const o = [];
  for await (const [n, x] of h.entries()) if (x.kind === 'directory') o.push([n, x]);
  return o;
};
async function picPlan(root, W, res) {
  const D = new Date(W),
    y = D.getUTCFullYear(),
    m = D.getUTCMonth(),
    two = D.getUTCDate() >= 16,
    cut = Date.UTC(y, m, 16),
    ld = new Date(Date.UTC(y, m + 1, 0)).getUTCDate(),
    monN = p2(m + 1) + ' ' + PMON[m] + ' ' + String(y).slice(-2),
    rndN = two ? '16-' + ld : '1-15',
    isRnd = n =>
      two ? /^16\s*-\s*(2[89]|3[01])$/.test(String(n).trim()) : /^0?1\s*-\s*15$/.test(String(n).trim());
  const S = [],
    shut = [];
  const walk = async (h, path, dep) => {
    for (const [n, x] of await picDirs(h)) {
      if (isMonD(n)) continue;
      if (dep === 0 && /ปิดถาวร/.test(n)) {
        for (const [c] of await picDirs(x)) shut.push(picKey(c));
        continue;
      }
      const k = picKey(n),
        kids = await picDirs(x),
        hm = kids.some(([c]) => isMonD(c));
      if (k.code || k.jam || k.metro != null || hm) S.push({ path: [...path, n], h: x, k, hm, kids });
      else if (dep < 2) await walk(x, [...path, n], dep + 1);
    }
  };
  await walk(root, [], 0);
  const best = (L, media) => {
      if (!L.length) return null;
      const w = L.filter(s => s.hm),
        C = w.length ? w : L;
      return C.find(s => media && nrm(s.path.join(' ')).includes(nrm(media))) || C[0];
    },
    same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
  const T = new Map(),
    miss = [],
    closed = [];
  for (const r of res || [])
    for (const b of r.fold || []) {
      const films = [
        ...new Set(
          (b.items || [])
            .filter(([n, g]) => n && (two ? g >= cut : g < cut))
            .map(([n]) => String(n).trim())
            .filter(Boolean),
        ),
      ];
      let tg;
      if (b.code)
        tg = [
          [
            b.code,
            { code: b.code },
            best(
              S.filter(s => s.k.code === b.code),
              b.media,
            ),
          ],
        ];
      else if (b.jam)
        tg = b.nums.map(n => [
          'BKJ' + b.jam + '_' + p2(n),
          { jam: b.jam + '_' + n },
          best(
            S.filter(s => s.k.jam === b.jam + '_' + n),
            '',
          ),
        ]);
      else if (b.metro)
        tg = b.nums.map(n => [
          'METRO_' + p2(n),
          { metro: n },
          best(
            S.filter(s => s.k.metro === n),
            '',
          ),
        ]);
      else {
        const L = S.filter(
          s =>
            s.k.nm &&
            s.k.nm.length >= 4 &&
            b.labels.some(
              t =>
                s.k.nm === t ||
                (t.length >= 6 && s.k.nm.includes(t)) ||
                (s.k.nm.length >= 6 && t.includes(s.k.nm)),
            ),
        );
        tg = [[b.labels[0] || b.sn, null, best(L, '')]];
      }
      for (const [id, k, s] of tg) {
        if (!s) {
          (k && shut.some(x => same(x, k)) ? closed : miss).push({ f: r.name, sn: b.sn, id });
          continue;
        }
        const key = s.path.join('/'),
          e = T.get(key) || { s, films: new Set(), src: new Set() };
        films.forEach(f => e.films.add(f));
        e.src.add(r.name + ' • ' + b.sn);
        T.set(key, e);
      }
    }
  const plan = [];
  for (const e of T.values()) {
    const mon = e.s.kids.find(([n]) => {
      const q = String(n)
        .trim()
        .match(/^(\d{1,2})\s*([A-Za-z]{3})[A-Za-z]*\s*(\d{2,4})$/);
      return q && PMON.indexOf(q[2].toUpperCase()) === m && +q[3] % 100 === y % 100;
    });
    const md = mon ? mon[1] : null;
    let rd = null;
    const have = new Set();
    if (md) {
      for (const [n, x] of await picDirs(md))
        if (isRnd(n)) {
          rd = x;
          break;
        }
      if (rd) for await (const [n] of rd.entries()) have.add(n.toLowerCase());
    }
    const need = [],
      nn = new Set();
    for (const n of [...PMOD, ...[...e.films].sort((a, b) => a.localeCompare(b, 'th'))]) {
      const f = fsafe(n),
        l = f.toLowerCase();
      if (!f || have.has(l) || nn.has(l)) continue;
      nn.add(l);
      need.push([n, f]);
    }
    plan.push({
      s: e.s,
      src: [...e.src],
      md,
      rd,
      mon: mon ? mon[0] : monN,
      rnd: rd ? rd.name : rndN,
      mkMon: !md,
      mkRnd: !rd,
      need,
      films: e.films.size,
    });
  }
  plan.sort((a, b) => a.s.path.join('/').localeCompare(b.s.path.join('/'), 'th'));
  return {
    plan,
    miss,
    closed,
    two,
    monN,
    rndN,
    dir: root.name,
    n: plan.reduce((t, p) => t + p.need.length + (p.mkMon ? 1 : 0) + (p.mkRnd ? 1 : 0), 0),
  };
}
async function picMake(P) {
  let n = 0;
  const bad = [];
  for (const p of P.plan) {
    if (!p.need.length && !p.mkMon && !p.mkRnd) continue;
    try {
      const md = p.md || (await p.s.h.getDirectoryHandle(p.mon, { create: true }));
      if (!p.md) n++;
      const rd = p.rd || (await md.getDirectoryHandle(p.rnd, { create: true }));
      if (!p.rd) n++;
      for (const [, f] of p.need) {
        try {
          await rd.getDirectoryHandle(f, { create: true });
          n++;
        } catch (_) {
          bad.push(p.s.path[p.s.path.length - 1] + ' / ' + f);
        }
      }
    } catch (_) {
      bad.push(p.s.path.join(' / '));
    }
  }
  return { n, bad };
}
/* ----- ตรวจโฟลเดอร์เทียบจอ (3.8) อ่านอย่างเดียว: โฟลเดอร์รอบ (เช่น 10 Oct 1-15 / 10 Oct 16-31) → แพค (Plan B TV 40 BLUE / 40 RED / 80 BLUE / 80 RED / 240, CITY TV 25, Spare) → โฟลเดอร์จอ
   เทียบกับใบเช็ครูปที่เติมแล้ว: จอในชีต PBTV 40 BLUE ต้องมีโฟลเดอร์ในโฟลเดอร์แพค 40 BLUE ฯลฯ • Plan B TV 240 = ทุกจอของ 4 แพค ป้าย (40B) (40R) (80B) (80R) ในชื่อต้องตรงกับโฟลเดอร์แพค
   • Playlist ของวันทำงานใช้บอกเพิ่มว่าฝั่งไหน (ใบเช็ค / โฟลเดอร์) ตรงกับรอบนี้ ----- */
const CPN = {
    '40Blue': '40 BLUE',
    '40Red': '40 RED',
    '80Blue': '80 BLUE',
    '80Red': '80 RED',
    CITY: 'CITY',
    Spare: 'Spare',
  },
  CPB = ['40Blue', '40Red', '80Blue', '80Red'],
  cpkOf = n => {
    const u = String(n)
      .toUpperCase()
      .replace(/[^A-Z0-9]/g, '');
    if (u.includes('240')) return 'ALL';
    if (/C[IY]TY/.test(u)) return 'CITY';
    if (u.includes('SPARE')) return 'Spare';
    for (const z of ['40', '80'])
      if (u.includes(z)) {
        if (u.includes('BLUE')) return z + 'Blue';
        if (u.includes('RED')) return z + 'Red';
      }
    return '';
  },
  cTag = n => {
    const m = String(n).match(/\(\s*(40|80)\s*([BR])\s*\)/i);
    return m ? m[1] + (/B/i.test(m[2]) ? 'Blue' : 'Red') : '';
  },
  cRnd = n => {
    const q = String(n)
      .trim()
      .match(/^(\d{1,2})\s*([A-Za-z]{3})[A-Za-z]*\.?\s*(?:(\d{2}|\d{4})\s+)?(\d{1,2})\s*[-–]\s*(\d{1,2})$/);
    if (!q) return null;
    const m = PMON.indexOf(q[2].toUpperCase());
    return m < 0 ? null : { m, y: q[3] ? +q[3] % 100 : null, a: +q[4], b: +q[5] };
  },
  cScr = n => {
    const c = codeOf(n);
    return { nm: n, key: c ? pck(c) : '', base: nrm(String(n).replace(/\([^)]*\)/g, ' ')), tag: cTag(n) };
  },
  cId = s => s.key || 'N:' + s.base,
  cLab = (a, b) => a === b || (a.length >= 6 && b.includes(a)) || (b.length >= 6 && a.includes(b)),
  cpl = k => [...k].map(x => CPN[x] || x).join(' / '),
  cFn = n => String(n).replace(/\.xls[xm]?$/i, '');
/* จอในใบเช็ค (ชีตของ 6 แพค) จากผลการเติม: รหัสจอ หรือชื่อจอในหัวคอลัมน์ (เช่น Eastin Makkasan) */
function cmpCK(res) {
  const K = [];
  for (const r of res || [])
    for (const b of r.fold || []) {
      if (!CPN[b.media]) continue;
      const lab = b.code ? [] : (b.labels || []).filter(t => !cpkOf(t));
      if (!b.code && !lab.length) continue;
      let e = b.code
        ? K.find(x => x.key === b.code)
        : K.find(x => !x.key && x.lab.some(t => lab.includes(t)));
      if (!e) {
        e = {
          key: b.code || '',
          lab: lab.slice(),
          nm: cl(b.raw) || b.code || lab[0],
          pk: new Set(),
          src: [],
        };
        K.push(e);
      }
      e.pk.add(b.media);
      e.src.push({ f: cFn(r.name), sn: b.sn, m: b.media });
    }
  return K;
}
/* แพคของจอใน Playlist วันนี้ (ไว้บอกว่าฝั่งไหนตรงกับรอบนี้) */
function cPl(key, lab) {
  const S = new Set();
  let city = false;
  for (const s of DB) {
    if (!CPN[s.m]) continue;
    if (s.m === 'CITY' && nrm(s.c) === 'THECITY') {
      city = true;
      continue;
    }
    const c = codeOf(s.c) || codeOf(s.n);
    if (
      key
        ? c && pck(c) === key
        : !c &&
          lab.some(t =>
            [
              nrm(s.c),
              ...String(s.n || '')
                .split(/_+/)
                .map(nrm),
            ].some(u => u.length >= 4 && cLab(u, t)),
          )
    )
      S.add(s.m);
  }
  if (!S.size && city && /^ID/.test(key || '')) S.add('CITY');
  return S;
}
async function cmpCheck(root, W, res) {
  const D = new Date(W),
    Y = D.getUTCFullYear() % 100,
    M = D.getUTCMonth(),
    d = D.getUTCDate(),
    two = d >= 16,
    ld = new Date(Date.UTC(D.getUTCFullYear(), M + 1, 0)).getUTCDate(),
    want =
      p2(M + 1) +
      ' ' +
      PMON[M].charAt(0) +
      PMON[M].slice(1).toLowerCase() +
      ' ' +
      (two ? '16-' + ld : '1-15');
  const R = [],
    top = cRnd(root.name);
  if (top) R.push({ h: root, path: [root.name], r: top });
  else {
    const walk = async (h, path, dep) => {
      for (const [n, x] of await picDirs(h)) {
        const r = cRnd(n);
        if (r) R.push({ h: x, path: [...path, n], r });
        else if (dep < 2) await walk(x, [...path, n], dep + 1);
      }
    };
    await walk(root, [], 0);
  }
  const yy = r => (r.y != null ? r.y : r.m <= M ? Y : Y - 1),
    rk = r => (yy(r) * 12 + r.m) * 2 + (r.a >= 16 ? 1 : 0),
    now = (Y * 12 + M) * 2 + (two ? 1 : 0);
  const cur = R.filter(x => rk(x.r) === now);
  let rd = cur.find(x => x.r.y === Y) || cur[0],
    prev = false;
  if (!rd) {
    const old = R.filter(x => rk(x.r) < now).sort((a, b) => rk(b.r) - rk(a.r));
    rd = old[0];
    prev = !!rd;
  }
  if (!rd) return { err: true, dir: root.name, want };
  const PK = [],
    other = [];
  for (const [n, x] of await picDirs(rd.h)) {
    const k = cpkOf(n);
    if (!k) {
      other.push(n);
      continue;
    }
    const S = [];
    for (const [sn] of await picDirs(x)) S.push(cScr(sn));
    S.sort((a, b) => a.nm.localeCompare(b.nm, 'th'));
    PK.push({ n, k, S });
  }
  PK.sort((a, b) => a.n.localeCompare(b.n));
  const K = cmpCK(res),
    pkF = k => PK.filter(p => p.k === k).map(p => p.n),
    ALL = PK.filter(p => p.k === 'ALL');
  /* โฟลเดอร์จอในโฟลเดอร์แพค (ไม่นับ 240) */
  const IX = new Map();
  for (const p of PK)
    if (p.k !== 'ALL')
      for (const sc of p.S) {
        const id = cId(sc);
        if (!IX.has(id)) IX.set(id, []);
        IX.get(id).push({ p, sc });
      }
  const fFind = e =>
    e.key
      ? IX.get(e.key) || []
      : [...IX.entries()]
          .filter(([id]) => id.startsWith('N:') && e.lab.some(t => cLab(t, id.slice(2))))
          .flatMap(([, v]) => v);
  const kOf = id =>
      K.find(x => (x.key ? x.key === id : id.startsWith('N:') && x.lab.some(t => cLab(t, id.slice(2))))),
    plOf = id => (id.startsWith('N:') ? cPl('', [id.slice(2)]) : cPl(id, []));
  /* แต่ละเรื่อง: g = ck (แก้ในใบเช็ค) / fd (แก้ในโฟลเดอร์) / chk (ให้ตรวจเอง) • act = สิ่งที่ต้องทำ • now = ตอนนี้เป็นอย่างไร */
  const dn = n =>
      String(n)
        .replace(/\([^)]*\)/g, ' ')
        .replace(/\s+/g, ' ')
        .trim(),
    tg = k => k.slice(0, 2) + k.charAt(2),
    who = L => [...new Set(L.map(s => '“' + s.f + '”'))].join(', '),
    shs = L => [...new Set(L.map(s => s.sn))].join(', '),
    fq = L => '“' + [...new Set(L)].join('” / “') + '”',
    plS = S => (S.size ? ' • Playlist วันนี้อยู่ ' + cpl(S) : '');
  const iss = [],
    inv = new Set(),
    rows = new Map(PK.map(p => [p.n, { n: p.n, k: p.k, nf: p.S.length, nk: 0, ok: 0, bad: 0 }])),
    add = (g, p, nm, act, now, id) => {
      iss.push({ g, p, nm: dn(nm), act, now: now || '' });
      if (id) inv.add(id);
      const r = rows.get(p);
      if (r) r.bad++;
    };
  let okK = 0;
  const nopk = new Map();
  /* 1) จอในใบเช็ค → โฟลเดอร์ */
  for (const e of K) {
    const H = fFind(e),
      pl = cPl(e.key, e.lab),
      eid = e.key || 'N:' + (e.lab[0] || ''),
      nm = H.length ? H[0].sc.nm : e.nm;
    for (const m of e.pk) {
      const cf = pkF(m);
      for (const f of cf) rows.get(f).nk++;
      if (!cf.length) nopk.set(m, (nopk.get(m) || 0) + 1);
      for (const x of H) if (x.p.k === m) rows.get(x.p.n).ok++;
    }
    if (e.pk.size > 1) {
      const P1 = pl.size === 1 ? [...pl][0] : '',
        Wr = P1 && e.pk.has(P1) ? e.src.filter(s => s.m !== P1) : [],
        now =
          'ใบเช็คแต่ละไฟล์ใส่จอนี้ไว้คนละแพค: ' + e.src.map(s => s.f + ' = ' + CPN[s.m]).join(', ') + plS(pl);
      if (Wr.length)
        add(
          'ck',
          '',
          nm,
          'ใบเช็ค ' + who(Wr) + ' ใส่จอนี้ไว้ชีต ' + shs(Wr) + ' แต่ควรอยู่แพค ' + CPN[P1],
          now,
          eid,
        );
      else add('chk', '', nm, 'ใบเช็คใส่จอนี้ไว้คนละแพค ดูว่าจอนี้อยู่แพคไหน แล้วแก้ใบเช็คที่ผิด', now, eid);
      continue;
    }
    const m = [...e.pk][0],
      cf = pkF(m);
    if (!cf.length) continue;
    const ok = H.filter(x => x.p.k === m),
      bad = H.filter(x => x.p.k !== m && x.p.k !== 'ALL'),
      src = 'ใบเช็ค ' + who(e.src) + ' ชีต ' + shs(e.src);
    if (!H.length) {
      if (pl.size && !pl.has(m))
        add(
          'chk',
          cf[0],
          nm,
          'ไม่มีโฟลเดอร์จอนี้ และ Playlist วันนี้ไม่ได้อยู่แพค ' + CPN[m] + ' ดูว่าจอนี้อยู่แพคไหน',
          src + ' (แพค ' + CPN[m] + ')' + plS(pl),
          eid,
        );
      else
        add(
          'fd',
          cf[0],
          nm,
          'สร้างโฟลเดอร์จอนี้ใน “' + cf[0] + '”',
          src + ' มีจอนี้ แต่ในเครื่องยังไม่มีโฟลเดอร์' + plS(pl),
          eid,
        );
      continue;
    }
    if (!ok.length) {
      const ks = [...new Set(bad.map(x => x.p.k))],
        fn = fq(bad.map(x => x.p.n)),
        hit = ks.filter(k => pl.has(k)),
        now = 'ใบเช็คอยู่ ' + CPN[m] + ' • โฟลเดอร์อยู่ ' + fn + plS(pl);
      for (const x of bad) inv.add(cId(x.sc));
      if (hit.length && !pl.has(m))
        add(
          'ck',
          bad[0].p.n,
          nm,
          'ใบเช็ค ' + who(e.src) + ' ใส่จอนี้ไว้ชีต ' + shs(e.src) + ' แต่ควรอยู่แพค ' + cpl(hit),
          now,
          eid,
        );
      else if (pl.has(m) && !hit.length)
        add('fd', bad[0].p.n, nm, 'ย้ายโฟลเดอร์จาก ' + fn + ' ไป “' + cf[0] + '”', now, eid);
      else
        add(
          'chk',
          bad[0].p.n,
          nm,
          'ใบเช็คกับโฟลเดอร์อยู่คนละแพค ดูว่าจอนี้อยู่แพคไหน แล้วแก้ฝั่งที่ผิด',
          now,
          eid,
        );
      continue;
    }
    if (bad.length) {
      for (const x of bad)
        add(
          'fd',
          x.p.n,
          x.sc.nm,
          'เอาโฟลเดอร์ใน “' + x.p.n + '” ออก',
          'จอนี้อยู่แพค ' + CPN[m] + ' ตามใบเช็ค และมีโฟลเดอร์ใน “' + ok[0].p.n + '” แล้ว' + plS(pl),
          cId(x.sc),
        );
      continue;
    }
    okK++;
  }
  for (const [k, n] of nopk)
    add(
      'fd',
      '',
      'แพค ' + CPN[k],
      'สร้างโฟลเดอร์แพค ' + CPN[k] + ' ในโฟลเดอร์รอบ “' + rd.path[rd.path.length - 1] + '”',
      'ใบเช็คมีจอของแพคนี้ ' + n + ' จอ แต่รอบนี้ยังไม่มีโฟลเดอร์แพค',
      '',
    );
  /* 2) จอเดียวกันมีโฟลเดอร์หลายแพค (ไม่นับ 240) ที่ไม่ได้บอกไว้ในข้อ 1 */
  for (const [id, L] of IX) {
    if (new Set(L.map(x => x.p.k)).size < 2 || inv.has(id)) continue;
    const e = kOf(id);
    if (e && e.pk.size === 1) continue;
    const pl = plOf(id),
      nm = L[0].sc.nm,
      now = 'มีโฟลเดอร์ทั้งใน ' + fq(L.map(x => x.p.n)).replace(/” \/ “/g, '” และ “') + plS(pl),
      bad = L.filter(x => !pl.has(x.p.k)),
      to = [...pl].flatMap(pkF);
    if (pl.size && bad.length && bad.length < L.length)
      for (const x of bad) add('fd', x.p.n, nm, 'เอาโฟลเดอร์ใน “' + x.p.n + '” ออก', now, id);
    else if (pl.size && bad.length === L.length && to.length)
      add(
        'fd',
        L[0].p.n,
        nm,
        'ย้ายโฟลเดอร์จาก ' + fq(L.map(x => x.p.n)) + ' ไปไว้ที่ “' + to[0] + '” ที่เดียว',
        now,
        id,
      );
    else add('chk', L[0].p.n, nm, 'จอนี้มีโฟลเดอร์หลายแพค ดูว่าควรอยู่แพคไหน', now, id);
  }
  /* 3) Plan B TV 240 เทียบกับโฟลเดอร์แพค 40 / 80 */
  for (const A of ALL) {
    const got = new Set();
    rows.get(A.n).nk = K.filter(e => [...e.pk].some(k => CPB.includes(k))).length;
    for (const sc of A.S) {
      const id = cId(sc),
        L = (IX.get(id) || []).filter(x => CPB.includes(x.p.k)),
        ks = [...new Set(L.map(x => x.p.k))],
        e = kOf(id);
      got.add(id);
      if (inv.has(id)) continue;
      if (!L.length) {
        const o = IX.get(id),
          pl = plOf(id),
          pb = CPB.filter(k => pl.has(k)),
          on = o ? ' • โฟลเดอร์แพคอยู่ ' + fq(o.map(x => x.p.n)) : '';
        if (pb.length === 1 && pkF(pb[0]).length)
          add(
            'fd',
            A.n,
            sc.nm,
            'สร้างโฟลเดอร์จอนี้ใน “' + pkF(pb[0])[0] + '”',
            'มีใน “' + A.n + '” แต่ในโฟลเดอร์แพคยังไม่มี' + on + plS(pl),
            id,
          );
        else if (pl.size && !pb.length)
          add(
            'fd',
            A.n,
            sc.nm,
            'เอาโฟลเดอร์ใน “' + A.n + '” ออก',
            'จอนี้ไม่ได้อยู่ใน 4 แพค Plan B TV' + on + plS(pl),
            id,
          );
        else
          add(
            'chk',
            A.n,
            sc.nm,
            'จอนี้มีใน “' + A.n + '” แต่ไม่มีในโฟลเดอร์แพค 40 / 80',
            on.replace(/^ • /, ''),
            id,
          );
        continue;
      }
      if (ks.length > 1) continue;
      const k = ks[0],
        pn = 'โฟลเดอร์แพคอยู่ “' + L[0].p.n + '”';
      if (!sc.tag) {
        add('fd', A.n, sc.nm, 'เติม (' + tg(k) + ') ในชื่อโฟลเดอร์ใน “' + A.n + '”', pn, id);
        continue;
      }
      if (sc.tag !== k) {
        add(
          'fd',
          A.n,
          sc.nm,
          'แก้ชื่อโฟลเดอร์ใน “' + A.n + '”: เปลี่ยน (' + tg(sc.tag) + ') เป็น (' + tg(k) + ')',
          pn,
          id,
        );
        continue;
      }
      if (e && [...e.pk].some(q => CPB.includes(q))) rows.get(A.n).ok++;
    }
    for (const [id, L] of IX)
      for (const x of L)
        if (CPB.includes(x.p.k) && !got.has(id)) {
          got.add(id);
          if (inv.has(id)) continue;
          add(
            'fd',
            A.n,
            x.sc.nm,
            'เพิ่มโฟลเดอร์จอนี้ใน “' + A.n + '”',
            'มีในโฟลเดอร์ “' + x.p.n + '” แต่ใน “' + A.n + '” ยังไม่มี',
            id,
          );
        }
  }
  const G = [
    ['ck', 'ต้องแก้ในใบเช็ค', 'แก้ในไฟล์ใบเช็คของคนที่ระบุ'],
    ['fd', 'ต้องแก้ในโฟลเดอร์', 'แก้ในโฟลเดอร์เทียบจอในเครื่อง'],
    ['chk', 'ให้ตรวจเอง', 'ระบบบอกไม่ได้ว่าฝั่งไหนถูก ดูแล้วแก้ฝั่งที่ผิด'],
  ];
  const RW = [...rows.values()],
    rdn = rd.path[rd.path.length - 1];
  const T = [
    'ตรวจโฟลเดอร์เทียบจอ รอบ “' + rdn + '”' + (prev ? ' (ยังไม่มีรอบ ' + want + ' จึงใช้รอบล่าสุด)' : ''),
    'ตรวจจอในใบเช็ค ' +
      K.length +
      ' จอ • ' +
      (iss.length ? 'ถูกต้อง ' + okK + ' จอ • ต้องแก้ ' + iss.length + ' เรื่อง' : 'ถูกต้องทั้งหมด'),
  ];
  for (const [g, lb] of G) {
    const I = iss.filter(x => x.g === g);
    if (!I.length) continue;
    T.push('', lb + ' (' + I.length + ')');
    I.forEach((x, i) => {
      T.push(i + 1 + '. ' + x.nm + ' — ' + x.act);
      if (x.now) T.push('   ' + x.now);
    });
  }
  return {
    dir: root.name,
    rd: rd.path.join(' / '),
    rdn,
    prev,
    want,
    rows: RW,
    iss,
    other,
    nk: K.length,
    okK,
    G,
    n: iss.length,
    txt: T.join('\n'),
  };
}
async function fillWb(f, on, metro, o) {
  const rep = [],
    FOLD = [];
  const buf = f.buf.slice(0);
  const wb = XLSX.read(f.buf, { type: 'array' }),
    zip = await JSZip.loadAsync(buf);
  const wx = new DOMParser().parseFromString(
    await zip.file('xl/workbook.xml').async('string'),
    'application/xml',
  );
  const rels = new DOMParser().parseFromString(
    await zip.file('xl/_rels/workbook.xml.rels').async('string'),
    'application/xml',
  );
  const rid = {};
  for (const r of rels.getElementsByTagName('Relationship'))
    rid[r.getAttribute('Id')] = r.getAttribute('Target');
  const path = {};
  for (const s of wx.getElementsByTagNameNS(NS, 'sheet')) {
    const t = rid[s.getAttributeNS(RNS, 'id')];
    path[s.getAttribute('name')] = t.startsWith('/') ? t.slice(1) : 'xl/' + t;
  }
  const jobs = [];
  for (const sn of wb.SheetNames) {
    if (notChk(sn)) continue;
    const ws = wb.Sheets[sn],
      media = mediaOf(sn),
      blocks = [];
    let brow = 0;
    for (const rw of [11, 10, 12]) {
      if (blocks.length) break;
      for (const k of Object.keys(ws)) {
        const m = k.match(new RegExp('^([A-Z]+)' + rw + '$')),
          c = m && codeOf(ws[k].v);
        if (c) blocks.push([col2n(m[1]), c]);
      }
      if (blocks.length) brow = rw;
    }
    if (blocks.length && !media && !avCols(ws).length) {
      const zx = path[sn] && zip.file(path[sn]),
        xs = zx ? await zx.async('string') : '';
      if (!/<c r="[A-Z]+(?:9|10)"[^>]*>\s*<f[^>]*>\s*AVERAGE/.test(xs)) {
        rep.push([sn, '-', 'ข้าม (มีรหัสจอแต่ไม่มีคอลัมน์ผลวิเคราะห์ในแถว 9 — ไม่ใช่ชีตจอ)', 0]);
        continue;
      }
    }
    if (media && brow === 11) {
      const av = avCols(ws);
      for (const cn of av) {
        if (blocks.some(x => x[0] === cn)) continue;
        const raw = cl((ws[n2col(cn) + '11'] || {}).v);
        if (!raw || nrm(raw).length < 3) continue;
        const hit = [...on.keys()].filter(k => nrm(k.split('|')[1]) === nrm(raw)),
          own = hit.filter(k => k.startsWith(media + '|')),
          pick = own.length === 1 ? own[0] : hit.length === 1 ? hit[0] : null;
        if (pick) blocks.push([cn, pick.split('|')[1]]);
        else
          rep.push([
            sn,
            raw,
            hit.length > 1
              ? 'ชื่อจอนี้อยู่ในหลายสื่อ (' +
                hit.map(k => k.split('|')[0]).join(', ') +
                ') — ตั้งชื่อชีตให้ระบุสื่อ'
              : 'ไม่พบจอนี้ใน Playlist (ยังไม่ได้เติม)',
            0,
          ]);
      }
      blocks.sort((x, y) => x[0] - y[0]);
    }
    const mediaB = cn =>
      media || (avRowOf(ws) === 10 ? mediaOf(cl((ws[n2col(cn - 1) + '11'] || {}).v)) : null) || media;
    for (const [cn, code] of blocks) {
      const media = mediaB(cn);
      let key = media + '|' + code,
        ads = on.get(key);
      let fb = '';
      if (!ads) {
        const nk = c => c.replace(/-/g, ''),
          h = media
            ? [...on.keys()].filter(k => k.startsWith(media + '|') && nk(k.split('|')[1]) === nk(code))
            : [];
        if (h.length === 1) {
          key = h[0];
          ads = on.get(key);
          fb = ' (ใบเช็คเขียน ' + code + ' = ' + key.split('|')[1] + ' ใน Playlist)';
        }
      } /* รหัสที่ขีดต่างกัน เช่น ใบเช็ค DGT-174 / Playlist DGT174 */
      if (!ads && media === 'CITY' && on.has('CITY|THE CITY')) {
        key = 'CITY|THE CITY';
        ads = on.get(key);
        fb = ' (จอ CITY ใช้รายการ THE CITY)';
      }
      if (!ads) {
        const hit = [...on.keys()].filter(k => k.split('|')[1] === code);
        if (hit.length === 1) {
          key = hit[0];
          ads = on.get(key);
          fb =
            media === 'Spare'
              ? ' (จอ Spare ใช้รายการของสื่อ ' + hit[0].split('|')[0] + ')'
              : ' (ใช้รายการของสื่อ ' +
                hit[0].split('|')[0] +
                (media ? ' เพราะสื่อ ' + media + ' ไม่มีจอนี้' : ' เพราะชื่อชีตไม่ระบุสื่อ') +
                ' — โปรดตรวจ)';
        }
      }
      if (!ads) {
        const hit = [...on.keys()].filter(k => k.split('|')[1] === code);
        rep.push([
          sn,
          code,
          hit.length > 1
            ? 'รหัสนี้อยู่ในหลายสื่อ (' +
              hit.map(k => k.split('|')[0]).join(', ') +
              ') แต่ชื่อชีตไม่ระบุสื่อ — ตั้งชื่อชีตให้มี 40/80 + Blue/Red, CITY หรือ Spare'
            : 'ไม่พบจอนี้ใน Playlist',
          0,
        ]);
        continue;
      }
      jobs.push([sn, n2col(cn - 1), ads, DB.find(s => s.m + '|' + s.c === key), '']);
      rep.push([sn, code, 'เติมแล้ว' + fb, ads.length]);
    }
    if (blocks.length && avRowOf(ws) === 10) {
      const av = avCols(ws);
      for (const cn of av.filter(c => !av.includes(c - 1))) {
        if (blocks.some(x => x[0] === cn)) continue;
        const lc = n2col(cn - 1),
          label = lab10(ws, sn, cn),
          scr = screen10(ws, sn, cn) || namedScreen(txt10(ws, cn).join(' '), '', sn);
        if (!scr) {
          rep.push([sn, label, 'ไม่พบจอนี้ใน Playlist', 0]);
          continue;
        }
        const ads = listOf(scr, o);
        jobs.push([sn, lc, ads, scr, '']);
        rep.push([sn, label + ' → ' + scr.c, 'เติมแล้ว', ads.length]);
      }
    } /* ใบเช็คแบบใหม่ที่มีทั้งจอ DGT และจอที่ไม่มีรหัส (เช่น CCTV: Signature Max – Air Traveller Gateway) */
    if (!blocks.length && /METRO/i.test(sn.replace(/[^A-Za-z]/g, '')) && metro.length) {
      let first = 0;
      for (const n of avCols(ws)) if (!first || n < first) first = n;
      if (first) {
        jobs.push([sn, n2col(first - 1), metro, DB.find(s => s.m === 'MetroLED'), 'metro']);
        rep.push([sn, 'ทั้งแพค', 'เติมแล้ว', metro.length]);
      }
    } else if (
      !blocks.length &&
      !media &&
      /(JAMS?|BKJ)[12]([AB]|(?![0-9]))/.test(sn.toUpperCase().replace(/[^A-Z0-9]/g, ''))
    ) {
      const jm = sn
          .toUpperCase()
          .replace(/[^A-Z0-9]/g, '')
          .match(/(?:JAMS?|BKJ)([12])([AB])?/),
        key = jm[1] === '1' ? 'BKJ1 ALL' : 'BKJ2-' + (jm[2] || 'A'),
        scr = DB.find(x => x.m === 'Bangkok Jams' && x.c === key);
      const av = avCols(ws);
      av.sort((x, y) => x - y);
      if (!scr) rep.push([sn, key, 'ไม่พบจอนี้ใน Playlist', 0]);
      else if (!av.length) rep.push([sn, key, 'ข้าม (ไม่พบคอลัมน์ผลวิเคราะห์ในแถว 9)', 0]);
      else {
        const base = jamList(scr, o);
        for (const cn of av.filter(c => !av.includes(c - 1))) {
          const lc = n2col(cn - 1),
            ph = [];
          for (let r = FIRST; r < FIRST + 10; r++) {
            const c = ws[lc + r];
            if (c && typeof c.v === 'string' && c.v.startsWith('#')) ph.push(c.v);
          }
          const ads = [...ph, ...base];
          jobs.push([sn, lc, ads, scr, 'jam']);
          rep.push([sn, key + ' (คอลัมน์ ' + lc + ')', 'เติมแล้ว', ads.length]);
        }
      }
    } else if (!blocks.length && !media && !/^(ANALYSIS|X+)$/.test(sn.toUpperCase().replace(/[^A-Z]/g, ''))) {
      const av = avCols(ws);
      av.sort((x, y) => x - y);
      if (!av.length) rep.push([sn, '-', 'ข้าม (ไม่พบรหัสจอ และไม่พบคอลัมน์ผลวิเคราะห์ในแถว 9)', 0]);
      const r10 = avRowOf(ws) === 10;
      for (const cn of av.filter(c => !av.includes(c - 1))) {
        const lc = n2col(cn - 1),
          label = r10 ? lab10(ws, sn, cn) : cl((ws[lc + '10'] || {}).v),
          code = cl((ws[n2col(cn) + '11'] || {}).v) + cl((ws[n2col(cn) + '12'] || {}).v),
          scr = (r10 && screen10(ws, sn, cn)) || namedScreen(label, code, sn);
        if (!scr) {
          rep.push([
            sn,
            label || code || n2col(cn),
            /^BASE$/.test(nrm(sn))
              ? 'ไม่พบจอ Base ใน Playlist — นำเข้า Playlist ที่มีจอ Base ก่อน แล้วเติมใหม่'
              : 'ไม่พบจอนี้ใน Playlist',
            0,
          ]);
          continue;
        }
        const ads = listOf(scr, o);
        jobs.push([sn, lc, ads, scr, '']);
        rep.push([sn, label + ' → ' + scr.c, 'เติมแล้ว', ads.length]);
      }
    } else if (!blocks.length && media) rep.push([sn, '-', 'ข้าม (ไม่พบรหัสจอ DGT/ID-xx ในแถว 10-11)', 0]);
  }
  for (const j of jobs) {
    if (!j[3] || !j[3].rows) continue;
    const nq = noteQ(j[3], j[4], j[4] === 'jam' ? jamNumsOf(wb.Sheets[j[0]], j[1]) : null);
    if (nq.length)
      rep.push([
        j[0],
        j[3].c,
        'ตรวจ: หมายเหตุใน Playlist ที่ระบบยังอ่านไม่ออก แต่พูดถึงจอนี้ → ' + nq.join(' • '),
        0,
      ]);
  }
  const by = {};
  for (const j of jobs) (by[j[0]] = by[j[0]] || []).push(j);
  const KEEP = {};
  for (const sn in by) {
    const ws = wb.Sheets[sn],
      av = avCols(ws);
    av.sort((x, y) => x - y);
    const own = by[sn].map(j => col2n(j[1]));
    for (const cs of av.filter(c => !av.includes(c - 1))) {
      let ce = cs;
      while (av.includes(ce + 1)) ce++;
      if (own.some(l => l >= cs - 1 && l <= ce)) continue;
      const lc = n2col(cs - 1);
      let has = false;
      for (let r = FIRST; r < FIRST + 400 && !has; r++) {
        const x = ws[lc + r];
        if (x && typeof x.v === 'string' && x.v.trim() && !/^ออก/.test(x.v.trim())) has = true;
      }
      if (!has) continue;
      const label = cl((ws[lc + '10'] || {}).v) || cl((ws[n2col(cs) + '11'] || {}).v) || lc;
      (KEEP[sn] = KEEP[sn] || []).push([sn, lc, [], { c: label, rows: [], pd: 0 }, 'keep', ce - cs + 1]);
    }
  } /* จอที่ไม่มีใน Playlist วันนี้ แต่อยู่ในชีตเดียวกับจอที่เติม */
  const OSW = { s: 0, l: 0 };
  for (const sn of wb.SheetNames) {
    if (notChk(sn)) continue;
    const ws = wb.Sheets[sn];
    for (const k in ws) {
      if (k[0] === '!') continue;
      const v = ws[k].v,
        m = typeof v === 'string' && v.trim().match(/^ออก(วันที่)?\s*\d{1,2}$/);
      if (m) m[1] ? OSW.l++ : OSW.s++;
    }
  } /* แบบหมายเหตุของทั้งไฟล์ (“ออก 5” หรือ “ออกวันที่ 5”) ใช้กับชีตที่ยังไม่มีหมายเหตุ */
  const ST = o.hand
      ? mkStyles(
          new DOMParser().parseFromString(await zip.file('xl/styles.xml').async('string'), 'application/xml'),
        )
      : null,
    W = o.hand ? o.W : null;
  for (const sn in by) {
    const sh = new Sheet(await zip.file(path[sn]).async('string'));
    if (ST && W && by[sn].every(j => j[3])) {
      const L = layoutChk(sh, [...by[sn], ...(KEEP[sn] || [])], { ...o, osw: OSW }, ST, W, wb.Sheets[sn]);
      if (L.skip) rep.push([sn, 'รูปแบบ', 'ตรวจ: ' + L.skip, 0]);
      else {
        for (const x of L.fold || [])
          FOLD.push({ sn, lc: x.lc, items: x.items, ...picId(wb.Sheets[sn], sn, x.lc) });
        if (L.dropped)
          rep.push([
            sn,
            'วันที่ถ่าย',
            'ตรวจ: มีวันที่ถ่าย/ตัวเลขที่ใส่ไว้ ' +
              L.dropped +
              ' ช่องของหนังที่ไม่อยู่ในรายการแล้ว จึงไม่ได้ย้ายมา',
            0,
          ]);
        const rd = d => new Date(d).getUTCDate();
        rep.push([
          sn,
          'รูปแบบ',
          (L.contm
            ? 'ทั้งเดือน (วันที่ 16 เป็นต้นไปทำต่อในรายการเดิม แถบวันที่สีรอบ 2) หัววันที่แถว ' +
              L.hr +
              (L.renew
                ? ' • ต่อสัญญา/ขึ้นใหม่วันที่ 16 ' +
                  L.renew +
                  ' เรื่อง ลงใหม่ใต้แถบ 16 ' +
                  (o.k16 ? '(แถวเดิมไม่ใส่ออก ตามที่ตั้งไว้ของไฟล์นี้)' : '(แถวเดิมใส่ออกวันที่ 16)')
                : '')
            : L.two
              ? 'รอบ 16–' +
                rd(L.E) +
                ' (แถบวันที่รอบ 2) เริ่มแถว ' +
                L.S +
                (L.frozen ? ' ต่อท้ายรอบ 1–15 (รอบแรกไม่ถูกแก้)' : '')
              : 'รอบ 1–15 หัววันที่แถว ' + L.hr) +
            (L.noNotes ? ' • วันที่ 1 ยังไม่ใส่ “ออกวันที่”' : '') +
            (L.moved ? ' • ย้ายวันที่ถ่ายตามชื่อหนัง ' + L.moved + ' ช่อง (ทุกจอ)' : '') +
            (L.today ? ' • หนังใหม่ของวันนี้ ' + L.today + ' เรื่อง ลงใต้แถบวันที่ ' + rd(L.W) : '') +
            (L.nost ? ' • ไม่เพิ่ม ' + L.nost + ' เรื่องที่ยังไม่มีในใบเช็คแต่สถานะไม่ใช่ “ขึ้นใหม่”' : '') +
            (L.dates ? ' • แยกวันเริ่ม ' + L.dates + ' กลุ่ม' : '') +
            (L.trail ? '' : ' • ไม่ใส่แถบวันถัดไปท้ายตาราง (ตามแบบของชีตนี้)') +
            (L.bval ? ' • วันที่ในแถบของจอที่ 2 เป็นต้นไปใส่เป็นตัวเลข (ตามแบบของชีตนี้)' : '') +
            (L.merged ? ' • ไฟล์เดิม/ต่อสัญญา ' + L.merged + ' แถว ใช้แถวเดิม ไม่ใส่ออกวันที่' : '') +
            (L.kept ? ' • เก็บหนังที่ลงไว้แล้วแต่ไม่มีใน Playlist วันนี้ ' + L.kept + ' แถว' : '') +
            (L.cleared ? ' • ขึ้นรอบใหม่ ล้างวันที่ถ่ายของรอบก่อน ' + L.cleared + ' ช่อง (ทุกจอ)' : '') +
            (L.back
              ? ' • หนังที่เคยออกแล้วกลับมาขึ้นใหม่ ' +
                L.back +
                ' เรื่อง ลงแถวใหม่ใต้แถบวันที่ ' +
                new Date(L.W).getUTCDate()
              : '') +
            (L.later ? ' • ยังไม่ใส่ ' + L.later + ' เรื่องที่เริ่มหลัง ' + fd(new Date(L.H)) : '') +
            (L.keep
              ? ' • คงรายการเดิมของ ' + L.keep + ' (ไม่มีใน Playlist ที่ใส่มา) แถบวันที่ยังตรงกันทั้งชีต'
              : '') +
            (L.fdrop ? ' • ลบแถว #FAST_ADS ' + L.fdrop + ' แถว (ไม่ใช่หนังโฆษณา)' : ''),
          0,
        ]);
        for (const [c, l] of L.laterL || [])
          rep.push([sn, c, 'ตรวจ: ยังไม่ใส่ เพราะวันเริ่มในไฟล์ยังไม่ถึง → ' + l.join(', '), 0]);
        for (const [c, l] of L.ratioL || [])
          rep.push([
            sn,
            c,
            'ตรวจ: หนังใหม่วันนี้มีหมายเหตุระบุชนิดจอ (16:9 / 4:3) ระบบไม่รู้ว่าจอนี้เป็นแบบไหน จึงใส่ไว้ก่อน → ' +
              l.join(', '),
            0,
          ]);
        for (const j of by[sn]) {
          const sig = !!SIGAL[j[3].c],
            z = [
              ...new Set(
                j[3].rows
                  .filter(
                    r =>
                      !r.off &&
                      skipRow(j[3], r) &&
                      r.ad &&
                      r.ad !== j[3].c &&
                      !/เมื่อมีไฟล์/.test(r.ad) &&
                      !JUNK.test(r.ad.trim()) &&
                      !MTP.test(r.ad.trim()) &&
                      !FASTN.test(r.ad),
                  )
                  .map(r => r.ad),
              ),
            ];
          if (z.length)
            rep.push([
              sn,
              j[3].c,
              (sig
                ? 'ตรวจ: ไม่ใส่ตามหมายเหตุ “ไม่ออน/ออนเฉพาะ/ยกเว้น” → '
                : 'ตรวจ: ไม่ใส่เพราะจำนวนครั้ง = 0 → ') + z.join(', '),
              0,
            ]);
        }
      }
    } else
      for (const [, col, ads] of by[sn]) {
        for (let r = FIRST; r < FIRST + 300; r++) sh.clear(col, r);
        ads.forEach((a, i) => sh.setText(col, FIRST + i, a));
      }
    zip.file(path[sn], sh.xml(), { createFolders: false });
  }
  if (ST) zip.file('xl/styles.xml', ST.xml(), { createFolders: false });
  let cp = wx.getElementsByTagNameNS(NS, 'calcPr')[0];
  if (!cp) {
    cp = wx.createElementNS(NS, 'calcPr');
    let after = null;
    for (const n of ['sheets', 'functionGroups', 'externalReferences', 'definedNames']) {
      const e = wx.getElementsByTagNameNS(NS, n)[0];
      if (e) after = e;
    }
    if (after) after.parentNode.insertBefore(cp, after.nextSibling);
    else wx.documentElement.appendChild(cp);
  }
  cp.setAttribute('fullCalcOnLoad', '1');
  if (zip.file('xl/calcChain.xml')) {
    zip.remove('xl/calcChain.xml');
    for (const r of [...rels.getElementsByTagName('Relationship')])
      if (/calcChain$/i.test(r.getAttribute('Type') || '')) r.parentNode.removeChild(r);
    zip.file(
      'xl/_rels/workbook.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        new XMLSerializer().serializeToString(rels.documentElement),
      { createFolders: false },
    );
    const ct = zip.file('[Content_Types].xml');
    if (ct)
      zip.file(
        '[Content_Types].xml',
        (await ct.async('string')).replace(/<Override[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/gi, ''),
        { createFolders: false },
      );
  } /* ลบลำดับการคำนวณเดิม (Excel สร้างใหม่เอง) กัน Excel ขึ้น “ซ่อมแซมไฟล์” */
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      new XMLSerializer().serializeToString(wx.documentElement),
    { createFolders: false },
  );
  return { rep, fold: FOLD, blob: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }) };
}

const $ = s => document.querySelector(s),
  esc = s =>
    String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
const TABS = [
  ['dat', '① ข้อมูล'],
  ['pl', '② แก้ไขรายจอ'],
  ['tl', '③ แก้ทุกจอ'],
  ['fl', '④ เติมใบเช็ครูป'],
];
const OP = {
  media: 'ทั้งหมด',
  sort: 'orig',
  newOnly: false,
  dedupe: true,
  off: false,
  hand: true,
  base: '',
  r2: 'cont',
};
let DB = [],
  tab = 'dat',
  fm = 'ทั้งหมด',
  q = '',
  cs = null,
  C = [],
  last = null,
  note = '',
  dirty = 0;
const FSA =
  typeof window.showDirectoryPicker === 'function'; /* บันทึกทับไฟล์เดิมได้เฉพาะ Chrome / Edge บนคอม */
let CKDIR = null;
const HDB = {
  p: null,
  db() {
    return (
      this.p ||
      (this.p = new Promise((ok, no) => {
        const r = indexedDB.open('ps_fs', 1);
        r.onupgradeneeded = () => r.result.createObjectStore('h');
        r.onsuccess = () => ok(r.result);
        r.onerror = () => no(r.error);
      }))
    );
  },
  async get(k) {
    const d = await this.db();
    return new Promise(ok => {
      const q = d.transaction('h').objectStore('h').get(k);
      q.onsuccess = () => ok(q.result || null);
      q.onerror = () => ok(null);
    });
  },
  async set(k, v) {
    const d = await this.db();
    return new Promise(ok => {
      const t = d.transaction('h', 'readwrite');
      t.objectStore('h').put(v, k);
      t.oncomplete = () => ok();
      t.onerror = () => ok();
    });
  },
}; /* จำโฟลเดอร์ใบเช็คไว้ในเครื่องนี้ */
const isoD = d =>
  d.getUTCFullYear() +
  '-' +
  String(d.getUTCMonth() + 1).padStart(2, '0') +
  '-' +
  String(d.getUTCDate()).padStart(2, '0');
const saveBlob = (n, d) => {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([d]));
  a.download = n;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    URL.revokeObjectURL(a.href);
    a.remove();
  }, 1500);
};
/* วันที่จาก Excel: แปลงแบบ UTC กันวันเลื่อน 1 วันในเขตเวลาไทย */
const fdU = d =>
  String(d.getUTCDate()).padStart(2, '0') +
  '-' +
  'JanFebMarAprMayJunJulAugSepOctNovDec'.substr(d.getUTCMonth() * 3, 3) +
  '-' +
  String(d.getUTCFullYear()).slice(2);
const fd = v => {
  if (typeof v === 'number' && v > 20000 && v < 80000)
    return fdU(new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 864e5));
  if (v instanceof Date) {
    const d = new Date(v.getTime() + 432e5);
    return fdU(new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())));
  }
  return cl(v);
};
const isNew = r => /ขึ้นใหม่/.test(r.st || '');
const stable = rows => [...rows.filter(r => !r.off), ...rows.filter(r => r.off)];
function fixMetro(s) {
  if (s.m !== 'MetroLED') return s;
  let v = 0,
    ps = '',
    pe = '';
  for (const r of s.rows) {
    if (/^Version$/i.test(r.ad.trim())) {
      v++;
      continue;
    }
    if (v >= 2) r.off = true;
    if (r.s || r.e) {
      ps = r.s;
      pe = r.e;
    } else {
      r.s = ps;
      r.e = pe;
    }
  }
  s.rows = stable(s.rows.filter(r => !/^Version$/i.test(r.ad.trim())));
  return s;
}
function fromData(D) {
  const T = D.t;
  DB = D.s.map(s => ({
    m: s[0],
    c: s[1],
    n: s[2],
    f: s[3],
    ed: 0,
    rows: D.r
      .slice(s[4], s[4] + s[5] + s[6])
      .map((r, i) => ({
        ad: T[r[0]],
        ty: T[r[1]],
        len: r[2],
        cnt: r[3],
        st: T[r[4]],
        s: T[r[5]],
        e: T[r[6]],
        note: T[r[7]],
        off: i >= s[5],
      })),
  }));
  DB = DB.map(fixMetro);
}
/* ---------- Signature Max / Show DC: กติกาจากหมายเหตุ “ไม่ออน …” และ “ออนเฉพาะ …” ---------- */
const SIGAL = {
  'So BKK -H-V-': ['SO', 'SOBKK', 'SOBANGKOK'],
  'Twin Tube': ['TWT', 'TWINTUBE'],
  PrudentialSathorn: ['สาทร', 'สาธร', 'TWTPLUS', 'PRUDENTIAL', 'PRUDENTIALSATHORN'],
  'Air Traveller Gateway': ['AOT'],
  'The Big Y': ['BIGY', 'THEBIGY'],
  'Ploenchit Interchange': ['PJ', 'PLOENCHIT'],
  'At Narong Interchange': ['AN', 'NARONG', 'ATNARONG'],
  'Exchange Tower': ['EXT'],
  'Exchange Tower (SM)': ['EXT'],
  'Exchange Tower (SM) 2': ['EXT'],
  Base: ['BASE'],
  ShowDC: ['SDC', 'SHOWDC'],
  'Dusit Central Park': ['DS', 'DUSIT', 'CENTRALPARKDUSIT'],
};
const ntk = t =>
  String(t)
    .toUpperCase()
    .replace(/^\s*(จอ|แอร์)+/, '')
    .replace(/[^A-Z0-9฀-๿]/g, '');
const ntoks = seg => {
  const out = [];
  for (const p of seg.split(/[,\/]+/)) {
    const w = p.trim();
    if (!w) continue;
    out.push(ntk(w));
    for (const x of w.split(/\s+/)) out.push(ntk(x));
  }
  return out.filter(Boolean);
};
function noteSkip(scr, r) {
  const al = SIGAL[scr.c];
  if (!al) return false;
  const n = String(r.note || '');
  let m = n.match(/ไม่ออน(?:แอร์|จอ)*\s*(.*)/);
  if (m) return ntoks(m[1].split('//')[0]).some(t => al.includes(t));
  m = n.match(/ยกเว\u0e49+น(?:จอ)?\s*(.*)/);
  if (m)
    return ntoks(m[1].split('//')[0].split(/ชดเชย/)[0]).some(t =>
      al.includes(t),
    ); /* “ทุกจอยกเว้น SO TWT สาทร” / “ออนทุกจอยกเว้น THE BIGY” (ชื่อย่อจอ) → ไม่ใส่ในจอที่ยกเว้น (เขียนผิดเป็น ยกเว้้น ก็อ่านได้) */
  m = n.match(/(?:ออน|ขึ้น)?เ?ฉพาะ(?:จอ)?\s*(.*)/);
  if (m) {
    const t = ntoks(m[1].split('//')[0]);
    return t.some(x => ALLAL.has(x)) && !t.some(x => al.includes(x));
  }
  return false;
}
const ALLAL = new Set(Object.values(SIGAL).flat());
const skipRow = (scr, r) => (SIGAL[scr.c] ? noteSkip(scr, r) : r.cnt === 0 && !MTP.test(r.ad.trim()));
/* ไฟล์ Playlist แบบ Signature Max: 1 ชีต = 1 จอ หัวตารางมีคำว่า หนังโฆษณา และ “รายการปลด” แยกส่วนที่ถอดแล้ว */
const SIGCODE = { 'Exchange Tower': 'Exchange Tower (SM)', 'Exchange Tower (2)': 'Exchange Tower (SM) 2' };
/* ไฟล์แบบ Signature Max / Show DC / BKJ: 1 ชีต = 1 จอ หัวตารางมีคำว่า หนังโฆษณา และ “รายการปลด” แยกส่วนที่ถอดแล้ว */
function parseSig(wb, media, map) {
  media = media || 'Signature Max';
  const out = [];
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    if (!ws || !(media === 'LA PlanB' ? LAT : /^รายการหนังโฆษณา/).test(cl((ws.A1 || {}).v))) continue;
    const code = map ? map(sn) : media === 'Signature Max' ? SIGCODE[sn] || sn : sn;
    if (!code) continue;
    const R = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    R.e.c = Math.min(R.e.c, 30);
    R.e.r = Math.min(R.e.r, 3000);
    const A = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: true, range: R });
    const h = A.slice(0, 12).findIndex(r => r.some(x => cl(x) === 'หนังโฆษณา'));
    if (h < 0) continue;
    const H = A[h].map(x => cl(x)),
      ix = re => H.findIndex(x => re.test(x));
    const c = {
      ad: ix(/^หนังโฆษณา$/),
      ty: ix(/^Type$/i),
      len: ix(/^ความยาว/),
      cnt: ix(/^จำนวน\s*ครั้ง/),
      s: ix(/^วันเริ่มสัญญา/),
      e: ix(/^วันหมดสัญญา/),
      os: ix(/^เริ่ม\s*On-?Air/i),
      oe: ix(/^สิ้นสุด\s*On-?Air/i),
      st: ix(/^สถานะ/),
      note: ix(/^หมายเหตุ/),
    };
    const g = (row, k) => (c[k] < 0 ? '' : row[c[k]]);
    let off = false;
    const rows = [];
    for (let r = h + 1; r < A.length; r++) {
      const row = A[r];
      if (row.some(x => /^รายการปลด/.test(cl(x)))) {
        off = true;
        continue;
      }
      const ad = cl(g(row, 'ad'));
      if (!ad || ad === 'หนังโฆษณา' || /^[\d.]+$/.test(ad) || /^เวลา\s/.test(ad)) continue;
      const st = cl(g(row, 'st')),
        cn = g(row, 'cnt');
      rows.push({
        ad,
        ty: cl(g(row, 'ty')),
        len: +g(row, 'len') || 0,
        cnt: cn === '' || cn == null ? null : +cn || 0,
        st,
        s: fd(g(row, 'os')) || fd(g(row, 's')),
        e: fd(g(row, 'oe')) || fd(g(row, 'e')),
        note: cl(g(row, 'note')),
        off: off || /^ถอด/.test(st),
      });
    }
    if (rows.length) out.push({ m: media, c: code, n: sn, f: '', ed: 0, rep: 'screen', rows: stable(rows) });
  }
  return out;
}
/* Bangkok Jams รายวัน: ชีต BKJ ALL (BKJ1) • BKJ2-A • BKJ2-B */
const bkjCode = sn => {
  if (/^BKJ\s*1?\s*ALL$/i.test(sn)) return 'BKJ1 ALL';
  const m = sn.match(/^BKJ\s*2\s*-\s*([AB])$/i);
  return m ? 'BKJ2-' + m[1].toUpperCase() : null;
};
/* 7-Eleven Metro LED: ใช้ชีตของวันล่าสุด (เช่น “2 Oct”) และหาคอลัมน์จากหัวตาราง (การออกอากาศ Start/End, ระยะสัญญา Start/End) */
/* สถานะของ Metro อยู่ที่สีแถว: อ่านคำอธิบายสีในหัวชีต (เช่น Q1–Q3) แล้วเทียบกับสีช่อง Version ของแต่ละแถว */
const MCOL = t =>
    /ไฟล์เก่า|กลับมา/.test(t)
      ? 'ขึ้นใหม่ไฟล์เดิม'
      : /ต่อสัญญา/.test(t)
        ? 'ต่อสัญญา'
        : /ไฟล์ใหม่/.test(t)
          ? 'ขึ้นใหม่'
          : '',
  fillKey = x => {
    const s = x && x.s,
      g = s && s.patternType && s.patternType !== 'none' && s.fgColor;
    return !g
      ? ''
      : g.rgb
        ? String(g.rgb).slice(-6).toUpperCase()
        : g.theme != null
          ? 't' + g.theme + '/' + (+g.tint || 0).toFixed(2)
          : g.indexed != null
            ? 'i' + g.indexed
            : '';
  },
  laterD = (a, b) => {
    const x = parseD(a),
      y = parseD(b);
    return !x ? b : !y ? a : +y > +x ? b : a;
  };
function parseMetro(wb, I) {
  let sn = wb.SheetNames[wb.SheetNames.length - 1],
    best = -1;
  for (const n of wb.SheetNames) {
    const d = fullNameDate(n);
    if (d && +d > best) {
      best = +d;
      sn = n;
    }
  }
  const ws = wb.Sheets[sn],
    A = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' }),
    R0 = XLSX.utils.decode_range(ws['!ref'] || 'A1').s,
    at = (r, c) => ws[XLSX.utils.encode_cell({ r: r + R0.r, c: c + R0.c })];
  const h = A.slice(0, 15).findIndex(r => r.some(x => cl(x) === 'Version'));
  let cV = 5,
    oS = 14,
    oE = 15,
    cS = -1,
    cE = -1;
  const LEG = {};
  for (let r = 0; r < (h >= 0 ? h : 5); r++)
    (A[r] || []).forEach((x, c) => {
      const t = cl(x),
        st = t && MCOL(t),
        k = st && fillKey(at(r, c));
      if (k) LEG[k] = st;
    });
  if (h >= 0) {
    cV = A[h].findIndex(x => cl(x) === 'Version');
    const st = [],
      en = [];
    (A[h + 1] || []).forEach((x, i) => {
      if (/^Start$/i.test(cl(x))) st.push(i);
      if (/^End$/i.test(cl(x))) en.push(i);
    });
    if (st.length) {
      oS = st[0];
      oE = en[0] != null ? en[0] : oS + 1;
      cS = st[1] != null ? st[1] : -1;
      cE = en[1] != null ? en[1] : -1;
    }
  }
  const g = (row, i) => (i < 0 ? '' : row[i]);
  let off = false,
    gs = '',
    ge = '',
    nocol = 0;
  const rows = [];
  const cT = h >= 0 ? A[h].findIndex(x => /^Time$/i.test(cl(x))) : -1,
    cG = h >= 0 ? A[h].findIndex(x => /^Tags?$/i.test(cl(x))) : -1;
  for (let r = h >= 0 ? h + 2 : 5; r < A.length; r++) {
    const row = A[r] || [];
    const c0 = cl(row[0]);
    if (/^รายการถอด/.test(c0)) {
      off = true;
      continue;
    }
    const ad = cl(row[cV]);
    if (!ad) continue;
    if (c0) {
      gs = fd(g(row, cS));
      ge = fd(g(row, cE));
    }
    const fk = fillKey(at(r, cV));
    if (fk && !LEG[fk] && ad !== 'Version') nocol++;
    const cs = fd(g(row, cS)) || gs,
      ce = fd(g(row, cE)) || ge; /* ระยะสัญญาของแถวนี้ (แถวแรกของสัญญามีวันที่ แถวถัดไปใช้ของสัญญาเดียวกัน) */
    rows.push({
      ad,
      ty: cl(row[3]),
      len: +row[cV + 1] || 0,
      cnt: +row[cV + 2] || 0,
      st: LEG[fk] || '',
      s: laterD(fd(g(row, oS)), cs),
      e: fd(g(row, oE)) || fd(g(row, cE)) || ge,
      cs,
      ce,
      note: cl(row[2]),
      loc: [cl(g(row, cT)), cl(g(row, cG))].filter(x => /LEDMETRO/i.test(x)).join(' '),
      off,
    });
  }
  if (I && !Object.keys(LEG).length && rows.some(r => r.ad !== 'Version'))
    I.push({
      lv: 'info',
      t:
        'Metro LED: ไม่พบคำอธิบายสี (ต่อสัญญา / ไฟล์ใหม่ / ไฟล์เก่ากลับมาออกอากาศใหม่) ในหัวชีต ' +
        sn +
        ' ระบบจึงไม่ได้อ่านสถานะจากสีแถว',
    });
  return [fixMetro({ m: 'MetroLED', c: 'METRO', n: '7-Eleven Metro LED', f: 'LED', ed: 0, rows })];
}
/* ด่านตรวจ Playlist (3.0): สรุปรายการหนังแยกจอ อ่านคอลัมน์ตามชื่อหัวตาราง ถ้าทีม Playlist แทรก/สลับคอลัมน์ก็ยังอ่านถูก หาไม่เจอ → ใช้ตำแหน่งเดิมและเตือน */
const SPLIT_DEF = { ty: 1, len: 2, cnt: 3, st: 4, s: 6, e: 7, note: 8 },
  SPLIT_L = {
    ty: [/^Type$/i, 'Type'],
    len: [/^ความยาว/, 'ความยาว'],
    cnt: [/^จำนวน/, 'จำนวนครั้ง'],
    st: [/^สถานะ/, 'สถานะ spot'],
    s: [/^วันเริ่ม/, 'วันเริ่มสัญญา'],
    e: [/^วันหมด/, 'วันหมดสัญญา'],
    note: [/^(ข้อมูลการออน|หมายเหตุ)/, 'ข้อมูลการออน'],
  },
  SPLIT_KNOWN = /^(หนังโฆษณา|Type|ความยาว|จำนวน|สถานะ|เวลา|วันเริ่ม|วันหมด|ข้อมูลการออน|หมายเหตุ)/i;
const splitSheet = sn => !(sn === 'Index' || sn.startsWith('data') || sn.startsWith('Blank'));
const splitMedia = wb => {
  const ms = [...new Set(wb.SheetNames.filter(splitSheet).map(mediaOf).filter(Boolean))];
  return ms.length === 1 ? ms[0] : null;
};
const isSplitWb = wb =>
  wb.SheetNames.some(sn => {
    const ws = wb.Sheets[sn];
    if (!ws || !splitSheet(sn) || /^รายการหนังโฆษณา/.test(cl((ws.A1 || {}).v))) return false;
    const A = XLSX.utils.sheet_to_json(ws, {
      header: 1,
      defval: '',
      range: { s: { r: 0, c: 0 }, e: { r: 1, c: 60 } },
    });
    return (A[1] || []).some((x, i) => cl(x) === 'หนังโฆษณา' && codeOf((A[0] || [])[i]));
  });
function parseSplit(wb, media, I) {
  const out = [],
    names = {},
    iss = (lv, t) => {
      if (!I.some(x => x.t === t)) I.push({ lv, t });
    };
  if (wb.Sheets.Index)
    for (const r of XLSX.utils.sheet_to_json(wb.Sheets.Index, { header: 1, defval: '' }).slice(2))
      for (const i of [1, 5, 9]) {
        const nm = cl(r[i]);
        if (nm) names[nm.split(/[_ ]/)[0]] = nm;
      }
  for (const sn of wb.SheetNames) {
    if (!splitSheet(sn)) continue;
    const A = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: '' });
    if (!A.length) continue;
    const hr = A.slice(0, 6).findIndex(r => r.some(x => cl(x) === 'หนังโฆษณา')),
      blocks = [];
    if (hr < 1) {
      if (A[0].length > 1)
        iss('stop', `ชีต ${sn}: ไม่พบหัวตาราง “หนังโฆษณา” จึงอ่านตามตำแหน่งคอลัมน์เดิม อาจอ่านผิด`);
      for (let c0 = 1; c0 < A[0].length; c0 += 9) blocks.push([c0, SPLIT_DEF, 1]);
    } else {
      const H = A[hr].map(x => cl(x)),
        st = H.map((x, i) => (x === 'หนังโฆษณา' ? i : -1)).filter(i => i >= 0);
      if (hr !== 1)
        iss('info', `ชีต ${sn}: หัวตารางย้ายไปอยู่แถว ${hr + 1} (เดิมแถว 2) ระบบอ่านตามหัวตารางแล้ว`);
      st.forEach((c0, k) => {
        const end = k + 1 < st.length ? st[k + 1] : H.length,
          o = {};
        for (const key in SPLIT_L) {
          let x = -1;
          for (let c = c0 + 1; c < end; c++)
            if (SPLIT_L[key][0].test(H[c])) {
              x = c - c0;
              break;
            }
          if (x < 0) {
            o[key] = SPLIT_DEF[key];
            iss('stop', `ชีต ${sn}: หาคอลัมน์ “${SPLIT_L[key][1]}” ไม่เจอ จึงอ่านตามตำแหน่งเดิม อาจอ่านผิด`);
          } else o[key] = x;
        }
        for (let c = c0 + 1; c < end; c++)
          if (H[c] && !SPLIT_KNOWN.test(H[c]))
            iss('info', `ชีต ${sn}: มีคอลัมน์ใหม่ “${H[c]}” ระบบไม่ได้ใช้`);
        if (Object.keys(SPLIT_DEF).some(key => o[key] !== SPLIT_DEF[key]))
          iss('info', `ชีต ${sn}: คอลัมน์ย้ายตำแหน่งจากแบบเดิม ระบบอ่านตามชื่อหัวตารางแล้ว`);
        blocks.push([c0, o, hr]);
      });
    }
    for (const [c0, o, h] of blocks) {
      const top = A[h - 1] || [],
        code = cl(top[c0]);
      if (!code) continue;
      let off = false;
      const rows = [];
      for (let r = h + 1; r < A.length; r++) {
        const R = A[r],
          ad = cl(R[c0]);
        if (!ad) continue;
        if (ad === 'หนังโฆษณา') {
          off = true;
          continue;
        }
        if (/^[\d.]+$/.test(ad) && !cl(R[c0 + o.ty])) continue;
        if (cl(R[c0 + o.st]) === 'ถอด') off = true;
        rows.push({
          ad,
          ty: cl(R[c0 + o.ty]),
          len: +R[c0 + o.len] || 0,
          cnt: +R[c0 + o.cnt] || 0,
          st: cl(R[c0 + o.st]),
          s: fd(R[c0 + o.s]),
          e: fd(R[c0 + o.e]),
          note: cl(R[c0 + o.note]),
          off,
        });
      }
      out.push({
        m: media,
        c: code,
        n: names[code] || code,
        f: cl(top[c0 + o.len]).replace(/[, ]+$/, ''),
        ed: 0,
        rows,
      });
    }
  }
  return out;
}
/* อ่าน Playlist ให้เร็วขึ้น (3.3): ชีตใน Playlist โดยเฉพาะ Signature Max มีช่องว่างที่มีแต่รูปแบบเป็นล้านช่อง (ข้อมูลจริงไม่ถึง 3,000 ช่อง) → ตัดช่องที่ไม่มีค่าออกก่อนให้ SheetJS อ่าน และคลายไฟล์ด้วยตัวคลายของเบราว์เซอร์ (เร็วกว่า) — ค่าที่อ่านได้เหมือนเดิมทุกช่อง ถ้ามีปัญหาใช้ไฟล์เดิม */
const EMPTYC = /<c r="[A-Z]{1,3}\d+"(?: s="\d+")?\/>/g,
  WSX = /^xl\/worksheets\/[^/]+\.xml$/;
async function slimXlsx(buf) {
  try {
    const u = new Uint8Array(buf),
      dv = new DataView(buf),
      raw = {},
      td = new TextDecoder();
    if (typeof DecompressionStream === 'function') {
      let e = u.length - 22;
      while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
      if (e >= 0) {
        const cnt = dv.getUint16(e + 10, true);
        let p = dv.getUint32(e + 16, true);
        for (let i = 0; i < cnt && p + 46 <= u.length && dv.getUint32(p, true) === 0x02014b50; i++) {
          const m = dv.getUint16(p + 10, true),
            cs = dv.getUint32(p + 20, true),
            nl = dv.getUint16(p + 28, true),
            xl = dv.getUint16(p + 30, true),
            kl = dv.getUint16(p + 32, true),
            lo = dv.getUint32(p + 42, true),
            name = td.decode(u.subarray(p + 46, p + 46 + nl));
          if (WSX.test(name) && (m === 8 || m === 0) && dv.getUint32(lo, true) === 0x04034b50) {
            const ds = lo + 30 + dv.getUint16(lo + 26, true) + dv.getUint16(lo + 28, true);
            raw[name] = [m, u.subarray(ds, ds + cs)];
          }
          p += 46 + nl + xl + kl;
        }
      }
    }
    const z = await JSZip.loadAsync(buf),
      inf = async d =>
        new Uint8Array(
          await new Response(
            new Blob([d]).stream().pipeThrough(new DecompressionStream('deflate-raw')),
          ).arrayBuffer(),
        );
    let cut = 0;
    for (const k of Object.keys(z.files)) {
      if (!WSX.test(k)) continue;
      const r = raw[k],
        x = r ? td.decode(r[0] === 8 ? await inf(r[1]) : r[1]) : await z.file(k).async('string'),
        y = x.replace(EMPTYC, '');
      cut += x.length - y.length;
      z.file(k, y);
    }
    return cut > 0 ? await z.generateAsync({ type: 'arraybuffer', compression: 'STORE' }) : buf;
  } catch (e) {
    return buf;
  }
}
/* Playlist แบบแบ่งกลุ่ม (3.4): Rama9 Connect / Digital MRT / Gateway X — ชีต “รายการหนัง…” มีหลายกลุ่ม (Plan B, BMN, Content …) แต่ละกลุ่มมีหัวตาราง “รายชื่อหนังโฆษณา”
   กลุ่ม “AD.ออก / Ads ถอด” = หนังที่ถอด • ไม่มีช่องสถานะ: ช่องชื่อหนังสีเหลือง = ขึ้นใหม่ สีชมพู = ขึ้นใหม่ไฟล์เดิม (ตามคำอธิบายสีท้ายตาราง) และรายชื่อใต้ “รายการขึ้นใหม่” (ชีต Playlist ของ Rama9)
   Content อาคาร / Content Plan B ไม่ลงใบเช็ค (ใบเช็ค MRT ไม่มี) • Content อื่น เช่น Content กทพ. ลงเป็นหนัง # (ใบเช็ค Gateway X มี) • แถวที่ยังไม่มีชื่อไฟล์ เช่น “Onair 19/10/2026” ไม่ใส่ */
const SECT_SKIP = /^Content\s*(อาคาร|Plan\s*B)/i,
  SECT_OFF = /^(AD\.?\s*ออก|Ads?\.?\s*ถอด|รายการ\s*(ปลด|ถอด))/i,
  SECT_JUNK = /^(on\s*-?\s*air|วันที่)\s*\d|\d{1,2}\/\d{1,2}\/\d{2,4}/i,
  SECT_NEW = /^(FFFF00)$/,
  SECT_BACK = /^(FF66CC|FF99CC|FF66FF|FF99FF)$/;
const sectHead = ws => {
  for (let r = 0; r < 8; r++)
    for (let c = 0; c < 14; c++) {
      const x = ws[XLSX.utils.encode_cell({ r, c })];
      if (x && cl(x.v) === 'รายชื่อหนังโฆษณา') return true;
    }
  return false;
};
const isSectWb = wb =>
  wb.SheetNames.some(n => {
    const ws = wb.Sheets[n];
    return !!ws && /^รายการหนังโฆษณา/.test(cl((ws.A1 || {}).v)) && sectHead(ws);
  });
const sectMedia = (wb, fn) => {
  const t = nrm(
    fn +
      ' ' +
      wb.SheetNames.map(
        n => cl(((wb.Sheets[n] || {}).A1 || {}).v) + ' ' + cl(((wb.Sheets[n] || {}).F1 || {}).v),
      ).join(' '),
  );
  return /RAMA9/.test(t)
    ? 'Rama9 Connect'
    : /MRT/.test(t)
      ? 'Digital MRT'
      : /GATEWAY/.test(t)
        ? 'Gateway X'
        : '';
};
function parseSect(wb, fn, I) {
  const media = sectMedia(wb, fn),
    out = [],
    NEWS = new Set(),
    E = XLSX.utils.encode_cell;
  if (!media)
    I.push({
      lv: 'stop',
      t: 'อ่านไฟล์แบบแบ่งกลุ่มได้ แต่ไม่รู้ว่าเป็นสื่อไหน (Rama9 Connect / Digital MRT / Gateway X) — หนังในไฟล์นี้ยังไม่ได้ใช้เติม ส่งไฟล์ให้ผู้ดูแลระบบ',
    });
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    if (!ws || !ws['!ref']) continue;
    const R = XLSX.utils.decode_range(ws['!ref']);
    for (let r = R.s.r; r <= Math.min(R.e.r, 12); r++)
      for (let c = R.s.c; c <= Math.min(R.e.c, 40); c++) {
        const x = ws[E({ r, c })];
        if (x && /^รายการขึ้นใหม่$/.test(cl(x.v)))
          for (let q = r + 1; q <= Math.min(R.e.r, r + 300); q++) {
            const y = ws[E({ r: q, c })];
            if (y && cl(y.v)) NEWS.add(cl(y.v));
          }
      }
  }
  const W = (wb.Workbook || {}).Sheets || [],
    hid = n => {
      const S = W[wb.SheetNames.indexOf(n)];
      return !!(S && S.Hidden);
    },
    lists = wb.SheetNames.filter(n => {
      const ws = wb.Sheets[n];
      return ws && !hid(n) && /^รายการหนังโฆษณา/.test(cl((ws.A1 || {}).v)) && sectHead(ws);
    });
  for (const sn of lists) {
    const ws = wb.Sheets[sn],
      R = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    R.e.c = Math.min(R.e.c, 30);
    R.e.r = Math.min(R.e.r, 800);
    const A = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: true, range: R }),
      at = (r, c) => ws[E({ r: r + R.s.r, c: c + R.s.c })];
    let lab = '';
    for (let r = 1; r < 4 && r < A.length && !lab; r++) {
      const t = A[r].map(cl).find(x => /สถานี|ส่วนต่อ/.test(x));
      if (t) lab = t.replace(/\s+/g, ' ').trim();
    }
    const code = lists.length > 1 && lab ? lab : media || sn;
    let H = null,
      sec = '',
      off = false,
      skip = false;
    const rows = [];
    for (let r = 0; r < A.length; r++) {
      const T = A[r].map(cl),
        hc = T.indexOf('รายชื่อหนังโฆษณา');
      if (hc >= 0) {
        let t = '';
        for (let q = r - 1; q >= Math.max(0, r - 2) && !t; q--) t = cl(A[q][0]) || cl(A[q][hc]);
        sec = t;
        off = SECT_OFF.test(t);
        skip = SECT_SKIP.test(t);
        const sub = (A[r + 1] || []).map(cl),
          hd = T.map((x, i) => (x + (sub[i] || '')).replace(/\s+/g, '')),
          ix = re => hd.findIndex(x => re.test(x));
        H = {
          ad: hc,
          ty: ix(/^Type$/i),
          len: ix(/^Spot/i),
          cnt: ix(/^จำนวนครั้ง/),
          os: ix(/^เริ่มการออนแอร์/),
          oe: ix(/^สิ้นสุดการออนแอร์/),
          s: ix(/^วันเริ่มสัญญา/),
          e: ix(/^วันหมดสัญญา/),
          note: ix(/^หมายเหตุ/),
        };
        r++;
        continue;
      }
      if (!H) continue;
      if (!/^\d+$/.test(T[0])) {
        if (T.some(x => /^Total$/i.test(x))) H = null;
        continue;
      }
      if (skip) continue;
      const ad = T[H.ad];
      if (!ad || SECT_JUNK.test(ad) || JUNK.test(ad)) continue;
      const g = k => (H[k] < 0 ? '' : A[r][H[k]]),
        fk = fillKey(at(r, H.ad)),
        cn = g('cnt');
      const st = off
        ? 'ถอด'
        : NEWS.has(ad) || SECT_NEW.test(fk)
          ? 'ขึ้นใหม่'
          : SECT_BACK.test(fk)
            ? 'ขึ้นใหม่ไฟล์เดิม'
            : '';
      rows.push({
        ad,
        ty: /^Content/i.test(sec) ? 'Content' : cl(g('ty')),
        len: +g('len') || 0,
        cnt: cn === '' || cn == null ? null : +cn || 0,
        st,
        s: fd(g('os')) || fd(g('s')),
        e: fd(g('oe')) || fd(g('e')),
        note: '',
        off,
      });
    }
    if (rows.length)
      out.push({
        m: media || 'อื่นๆ',
        c: code,
        n: media === 'Digital MRT' ? 'MRT ' + code : code,
        f: '',
        ed: 0,
        rows: stable(rows),
      });
  }
  return out;
}
/* Air Traveller / The 20 (3.4): หน้าตาแบบ Signature Max แต่ไฟล์ละสื่อเดียว — ใช้ชื่อสื่อของตัวเอง ไม่ปนกับ Signature Max • ข้ามชีตสรุป (ต้นฉบับรายการหนังแยกจอ, รายการหนังประจำวัน) และชีตที่ไม่มีหนังจริง */
const MCO_SKIP = /^(ต้นฉบับ|รายการหนังประจำวัน|รายการหนังรายจอ)/,
  mcoOf = t =>
    /AIRTRAVELLERSERIES|AIRTRAVELLER(?!GATEWAY)/.test(nrm(t))
      ? 'Air Traveller'
      : /^THE20|THE20/.test(nrm(t))
        ? 'The 20'
        : '';
const isMcoWb = wb => wb.SheetNames.includes('รายการหนังประจำวัน');
function parseMco(wb, fn) {
  const own = mcoOf(fn) || '',
    L = parseSig(wb, 'Media Co').filter(
      s => !MCO_SKIP.test(s.c) && s.rows.some(r => !r.off && r.ad && !JUNK.test(r.ad.trim())),
    );
  const by = new Map();
  for (const s of L) {
    const m = mcoOf(s.c);
    if (m) (by.get(m) || by.set(m, []).get(m)).push(s);
  }
  const m = own && by.has(own) ? own : [...by.keys()].sort((a, b) => by.get(b).length - by.get(a).length)[0];
  const prog = sn => {
    const ws = wb.Sheets[sn],
      A = ws
        ? XLSX.utils.sheet_to_json(ws, {
            header: 1,
            defval: '',
            range: { s: { r: 0, c: 0 }, e: { r: 400, c: 20 } },
          })
        : [],
      h = A.slice(0, 12).findIndex(r => r.some(x => cl(x) === 'หนังโฆษณา'));
    if (h < 0) return new Set();
    const H = A[h].map(cl),
      pa = H.indexOf('สินค้า'),
      aa = H.indexOf('หนังโฆษณา');
    return new Set(
      pa < 0
        ? []
        : A.slice(h + 1)
            .filter(r => /^Programm?atic/i.test(cl(r[pa])))
            .map(r => cl(r[aa])),
    );
  }; /* โฆษณาแบบ Programmatic (เช่น Vistar) ไม่ลงใบเช็ค (ใบเช็ค Air Traveller ของทีมไม่มี) */
  return m
    ? by.get(m).map(s => {
        const P = prog(s.c);
        return { ...s, m, n: s.c, rep: undefined, rows: s.rows.filter(r => !P.has(r.ad)) };
      })
    : [];
}
/* LA PlanB: ชีตเดียว “รายการโฆษณาประจำวันสื่อ LA PlanB เริ่มตั้งแต่ 20 ตุลาคม 2569” มีรายการที่ออน + “รายการปลด” ส่งเฉพาะวันที่รายการเปลี่ยน (ไม่ได้ส่งทุกวัน → ใช้ฉบับล่าสุด) วันของไฟล์ = วันที่แถว 2 / “เริ่มตั้งแต่” ไม่ใช่วันเริ่มของหนัง • ใบเช็ค: ชีต LA PLANB จอ LA PlanB 1 และ 2 ใช้รายการเดียวกัน */
const LAT = /^รายการ(?:หนัง)?โฆษณา.*LA\s*Plan\s*B/i,
  THMF = [
    'มกราคม',
    'กุมภาพันธ์',
    'มีนาคม',
    'เมษายน',
    'พฤษภาคม',
    'มิถุนายน',
    'กรกฎาคม',
    'สิงหาคม',
    'กันยายน',
    'ตุลาคม',
    'พฤศจิกายน',
    'ธันวาคม',
  ];
const isLaWb = wb => wb.SheetNames.some(n => LAT.test(cl(((wb.Sheets[n] || {}).A1 || {}).v)));
function laDate(wb) {
  for (const n of wb.SheetNames) {
    const ws = wb.Sheets[n];
    if (!ws || !LAT.test(cl((ws.A1 || {}).v))) continue;
    const v = (ws.A2 || {}).v;
    if (v instanceof Date) return new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()));
    if (typeof v === 'number' && v > 40000 && v < 80000) return s2d(v);
    const m = cl(ws.A1.v).match(/(\d{1,2})\s*([\u0E01-\u0E4E]+)\s*(\d{4})/),
      i = m ? THMF.indexOf(m[2]) : -1;
    if (i >= 0) return new Date(Date.UTC(yy(m[3]), i, +m[1]));
  }
  return null;
}
function parsePlaylist(f) {
  const I = [],
    ret = x => {
      if (x && x.length) {
        x.iss = I;
        return x;
      }
      return null;
    };
  if (/Metro.*LED/i.test(f.name))
    return ret(parseMetro(XLSX.read(f.buf, { type: 'array', cellDates: true, cellStyles: true }), I));
  if (!/แยกจอ/.test(f.name)) {
    const wb = XLSX.read(f.buf, { type: 'array' });
    if (wb.SheetNames.some(n => bkjCode(n))) return ret(parseSig(wb, 'Bangkok Jams', bkjCode));
    if (isLaWb(wb)) {
      const p = parseSig(wb, 'LA PlanB', () => 'LA PlanB');
      if (p.length) {
        p.dt = laDate(wb);
        return ret(p);
      }
    }
    if (isSectWb(wb)) return ret(parseSect(XLSX.read(f.buf, { type: 'array', cellStyles: true }), f.name, I));
    if (isMcoWb(wb)) {
      const p = parseMco(wb, f.name);
      if (p.length) return ret(p);
    }
    if (isSplitWb(wb)) {
      const m = splitMedia(wb);
      if (!m) return null;
      I.push({
        lv: 'info',
        t: `ชื่อไฟล์ไม่มีคำว่า “แยกจอ” แต่เนื้อในเป็นสรุปรายการหนังแยกจอ (${m}) ระบบอ่านให้แล้ว`,
      });
      return ret(parseSplit(wb, m, I));
    }
    return ret(
      parseSig(wb, /show\s*_*\s*dc|prudential/i.test(f.name) ? 'Show DC / Prudential' : 'Signature Max'),
    );
  }
  /* สรุปรายการหนังแยกจอ: สื่ออยู่ในชื่อไฟล์ เช่น …แยกจอ_01_40Blue, …แยกจอ_2026_10_02_80Red ถ้าชื่อไฟล์ไม่บอก ใช้ชื่อชีต (40Blue, 80Red_1 …) */
  const bn = f.name.replace(/\.xlsx$/i, '').replace(/\s*\(\d+\)$/, ''),
    rest = bn
      .replace(/^.*แยกจอ/, '')
      .replace(/\d{4}[-_.]\d{1,2}[-_.]\d{1,2}/g, '')
      .replace(/\d{1,2}[-_.]\d{1,2}[-_.]\d{2,4}/g, '')
      .replace(/^[_\s]+|[_\s]+$/g, '')
      .replace(/^\d{1,2}[_\s]+(?=\D)/, '');
  const wb = XLSX.read(f.buf, { type: 'array' }),
    sm = splitMedia(wb);
  let media = mediaOf(rest);
  if (media && sm && sm !== media)
    I.push({
      lv: 'stop',
      t: `ชื่อไฟล์บอกสื่อ ${media} แต่ชีตในไฟล์เป็น ${sm} ระบบใช้ตามชื่อไฟล์ ตรวจว่าใส่ไฟล์ถูก`,
    });
  if (!media && sm) {
    media = sm;
    I.push({ lv: 'info', t: `ชื่อไฟล์ไม่บอกสื่อ ระบบใช้ชื่อชีตในไฟล์ (${sm}) แทน` });
  }
  if (!media) {
    if (!rest) return null;
    media = rest.replace(/^\d{4}[-_. ]\d{1,2}[-_. ]\d{1,2}[_\s]*/, '') || rest;
    I.push({
      lv: 'stop',
      t: `สื่อ “${media}” ระบบยังไม่รู้จัก หนังในไฟล์นี้จึงยังไม่ได้ใช้เติม ส่งไฟล์ให้ผู้ดูแลระบบ`,
    });
  }
  return ret(parseSplit(wb, media, I));
}
/* ด่านตรวจ Playlist: สถานะที่ไม่รู้จัก และวันที่ที่อ่านไม่ได้ (ทุกแบบไฟล์) */
const KST = /ขึ้นใหม่|ไฟล์เดิม|ต่อสัญญา|อัปเดต|ถอด|ปรับความถี่|ชดเชย|On-?Air/i,
  NOEND = /กำหนด|จนกว่า|ปลด|ไม่มี/;
function plCheck(p, I) {
  const st = new Map(),
    bad = new Map(),
    all = new Set();
  for (const s of p)
    for (const r of s.rows) {
      if (!r.ad || r.ad === s.c || JUNK.test(String(r.ad).trim())) continue;
      all.add(r.ad);
      const x = String(r.st || '').trim();
      if (x && !KST.test(x)) {
        if (!st.has(x)) st.set(x, new Set());
        st.get(x).add(r.ad);
      }
      for (const k of ['s', 'e']) {
        const v = String(r[k] || '').trim();
        if (v && !parseD(v) && !NOEND.test(v)) {
          if (!bad.has(r.ad)) bad.set(r.ad, (k === 's' ? 'วันเริ่ม' : 'วันหมด') + ' “' + v + '”');
          break;
        }
      }
    }
  for (const [x, a] of st)
    I.push({
      lv: 'stop',
      t: `สถานะ “${x}” ระบบไม่รู้จัก (${a.size} เรื่อง เช่น ${[...a].slice(0, 2).join(', ')}) หนังเหล่านี้จะไม่ถูกนับว่า “ขึ้นใหม่” จึงไม่ลงใต้แถบของวันนี้`,
    });
  if (bad.size) {
    const many = bad.size >= 10 || bad.size >= all.size * 0.1,
      ex = [...bad]
        .slice(0, 3)
        .map(([a, v]) => `${a} (${v})`)
        .join(', ');
    I.push({
      lv: many ? 'stop' : 'info',
      t: `อ่านวันที่ไม่ได้ ${bad.size} เรื่อง${many ? ' รูปแบบวันที่อาจเปลี่ยน' : ''}: ${ex}${bad.size > 3 ? ' …' : ''}`,
    });
  }
}
const plRep = () =>
  PLI.length
    ? [
        '',
        'ด่านตรวจ Playlist',
        ...PLI.map(
          x =>
            `  ${x.lv === 'stop' ? 'ต้องตรวจ' : 'ข้อสังเกต'}${x.ack ? ' (กดตรวจแล้ว)' : ''} | ${x.f} | ${x.t}`,
        ),
      ]
    : [];
/* ตอนเติม: หมายเหตุที่พูดถึงจอที่กำลังเติม (ไม่ออน / ยกเว้น / เฉพาะ …) แต่ระบบอ่านกติกาไม่ออก → ขึ้นใน “ควรตรวจ” ไม่เปลี่ยนผลการเติม */
const RULEW = /ยกเว้น|ไม่ออน|ไม่ขึ้น|ไม่รับ|ฉพาะ|ห้าม|งด/,
  NORMN = t => String(t || '').replace(/([ัิ-ฺ็-๎])\1+/g, '$1');
const jamNumsOf = (ws, lc) => {
  const s = new Set(),
    c0 = col2n(lc);
  if (!ws) return s;
  for (let r = 8; r <= 14; r++) {
    const x = ws[lc + r];
    if (!(x && /BANGKOK|BKJ|JAM|METRO/i.test(cl(x.v)))) continue;
    for (let c = c0 + 1; c < c0 + 60; c++) {
      const y = ws[n2col(c) + r],
        v = y ? cl(y.v) : '';
      if (!/^\d{1,3}$/.test(v)) break;
      s.add(+v);
    }
    break;
  }
  return s;
};
function noteQ(scr, kind, nums) {
  const out = new Map(),
    al = SIGAL[scr.c];
  for (const r of scr.rows || []) {
    if (
      r.off ||
      !r.ad ||
      r.ad === scr.c ||
      JUNK.test(String(r.ad).trim()) ||
      FASTN.test(r.ad) ||
      MTP.test(String(r.ad).trim()) ||
      (kind === 'jam' && (/^BJ\d*_\d+_/i.test(r.ad) || /weather/i.test(r.ad)))
    )
      continue;
    const n0 = String(r.note || ''),
      n = NORMN(n0);
    if (!n.trim()) continue;
    let hit = false;
    if (kind === 'jam') {
      if (nums && nums.size) {
        const L = new Set([...n.matchAll(/BKJ\s*\d?\s*-\s*(\d{1,3})/gi)].map(m => +m[1]));
        hit =
          RULEW.test(n) &&
          [...nums].every(x => L.has(x)) &&
          !/(?:ขึ้น|ออน)เ?ฉพาะ/.test(n0) &&
          !/ยกเว้น|ไม่ออน/.test(n0);
      }
    } else if (al)
      hit =
        RULEW.test(n) &&
        ntoks(n).some(t => al.includes(t)) &&
        !/ไม่ออน/.test(n0) &&
        !/ยกเว\u0e49+น/.test(n0) &&
        !/(?:ออน|ขึ้น)?เ?ฉพาะ/.test(n0);
    if (
      !hit &&
      /(?:ออนแอร์|ขึ้น)ระหว่างวัน/.test(n) &&
      /\d/.test(n) &&
      !/(?:ออนแอร์|ขึ้น)ระหว่างวัน\s*(?:วันที่)?\s*(\d{1,2})\/(\d{1,2})\/(\d{2,4})/.test(n0)
    )
      hit = true;
    if (hit && !out.has(r.ad)) out.set(r.ad, n0.trim().replace(/\s+/g, ' '));
  }
  return [...out].map(([a, x]) => `${a}: “${x.length > 90 ? x.slice(0, 90) + '…' : x}”`);
}
function listOf(s, o) {
  let L = s.rows.filter(r => o.off || !r.off);
  if (o.newOnly) L = L.filter(isNew);
  if (o.sort === 'az') L = [...L].sort((a, b) => a.ad.localeCompare(b.ad, 'th'));
  if (o.sort === 'new') L = [...L].sort((a, b) => isNew(b) - isNew(a));
  const n = L.map(r => r.ad);
  return o.dedupe ? [...new Set(n)] : n;
}
function buildOn(o) {
  const on = new Map();
  for (const s of DB) {
    if (s.m === 'MetroLED' || !/^(40|80|CITY|Spare)/.test(s.m)) continue;
    if (o.media !== 'ทั้งหมด' && s.m !== o.media) continue;
    on.set(s.m + '|' + s.c, listOf(s, o));
  }
  return on;
}
async function addFiles(kind, files) {
  const arr = [];
  for (const f of files) arr.push({ name: f.name, buf: await f.arrayBuffer() });
  if (kind === 'ck') {
    C.push(...arr);
    note = `เพิ่มใบเช็ครูป ${arr.length} ไฟล์`;
    return view();
  }
  if (kind === 'pj') {
    PLI = [];
    try {
      const o = JSON.parse(new TextDecoder().decode(arr[0].buf));
      if (Array.isArray(o)) {
        DB = o;
        PLD = null;
        PLM = {};
      } else {
        DB = o.DB || [];
        PLD = plDate() || (o.PLD ? new Date(o.PLD) : null);
      }
      note = 'เปิดโปรเจกต์แล้ว';
    } catch (e) {
      note = 'ไฟล์โปรเจกต์ไม่ถูกต้อง';
    }
    return view();
  }
  note = 'กำลังอ่าน Playlist …';
  view();
  await new Promise(r => setTimeout(r, 30));
  let n = 0,
    skip = [];
  const ok = [],
    emp = [],
    got = [];
  let _i = 0;
  for (const f of arr) {
    _i++;
    if (typeof busyStep === 'function') await busyStep('กำลังอ่าน Playlist', _i, arr.length, f.name);
    let p;
    try {
      p = parsePlaylist(/Metro.*LED/i.test(f.name) ? f : { name: f.name, buf: await slimXlsx(f.buf) });
    } catch (e) {
      p = null;
    }
    PLX = PLX.filter(x => x.f !== f.name);
    PLI = PLI.filter(x => x.f !== f.name);
    if (!p) {
      skip.push(f.name);
      PLX.push({
        f: f.name,
        why: 'ไม่รู้จักรูปแบบไฟล์ (ไม่ใช่ สรุปรายการหนังแยกจอ / Signature Max / Show DC / BKJ / Metro LED)',
      });
      continue;
    }
    if (!p.some(s => s.rows.some(r => !r.off))) {
      emp.push(f.name);
      PLX.push({ f: f.name, why: 'อ่านได้แต่ไม่พบรายการหนัง (รูปแบบไฟล์อาจเปลี่ยน) — ใช้ข้อมูลเดิมไว้ก่อน' });
      continue;
    } /* อ่านได้แต่ไม่มีรายการหนังเลย → ไม่ลบข้อมูลเดิม */
    const fd0 = p.dt || fullNameDate(f.name) || dataDate(p) || nameDate(f.name, autoBase());
    for (const x of p) {
      x.pf = f.name;
      x.pd = fd0 ? +fd0 : 0;
    }
    {
      const I = p.iss || [];
      plCheck(p, I);
      for (const x of I) PLI.push({ f: f.name, ...x });
    }
    got.push({ f, p, d: fd0 ? +fd0 : 0 });
  }
  got.sort((a, b) => a.d - b.d); /* โยนไฟล์หลายวันพร้อมกัน → จอเดียวกันใช้ไฟล์ของวันล่าสุด */
  for (const { f, p } of got) {
    if (p.every(x => x.rep === 'screen')) {
      const ks = new Set(p.map(x => x.m + '|' + x.c));
      DB = DB.filter(x => !ks.has(x.m + '|' + x.c));
    } else {
      const ms = new Set(p.map(x => x.m));
      DB = DB.filter(x => !ms.has(x.m));
    }
    DB.push(...p);
    n += p.length;
    ok.push(f.name);
  }
  if (ok.length) {
    PLD = plDate();
    OP.base = '';
  }
  {
    const live = new Set(DB.map(s => s.pf));
    PLI = PLI.filter(x => live.has(x.f));
  }
  note =
    `นำเข้า ${new Set(got.flatMap(x => x.p.map(s => s.m + '|' + s.c))).size} จอ` +
    (skip.length ? ` • ข้าม ${skip.length} ไฟล์ที่ไม่รู้จัก` : '') +
    (emp.length
      ? ` • ${emp.join(', ')}: อ่านได้แต่ไม่พบรายการหนัง (รูปแบบไฟล์อาจเปลี่ยน) จึงใช้ข้อมูลเดิมไว้ก่อน`
      : '');
  if (typeof busyDone === 'function') busyDone();
  view();
}
/* ---------- หน้าจอ v2.1: ธีมสว่าง/มืด • เช็กลิสต์ Playlist 11 แบบ • ความพร้อมของใบเช็ค • ทัวร์แนะนำ • ระบบทำงานอย่างไร ---------- */
/* ส่วนนี้คือหน้าจอ ส่วนอ่าน Playlist และเติมใบเช็ค (ด้านบน) แก้เพียงเล็กน้อย (แถบความคืบหน้า และ Show DC ไม่กำหนดวันทำงาน) — ผลการเติมของไฟล์ชุดเดียวกันเหมือนเวอร์ชัน 1.0 ทุกไบต์ */
const VERTXT = 'ระบบเติมใบเช็ครูป เวอร์ชัน 3.8.1 (20 ต.ค. 2569)';
const IC = {
  up: '<svg viewBox="0 0 24 24"><path d="M12 15V4m0 0L7.5 8.5M12 4l4.5 4.5M4 15v3.5A1.5 1.5 0 0 0 5.5 20h13a1.5 1.5 0 0 0 1.5-1.5V15"/></svg>',
  folder:
    '<svg viewBox="0 0 24 24"><path d="M3 7.5A1.5 1.5 0 0 1 4.5 6H9l2 2h8.5A1.5 1.5 0 0 1 21 9.5v8a1.5 1.5 0 0 1-1.5 1.5h-15A1.5 1.5 0 0 1 3 17.5z"/></svg>',
  save: '<svg viewBox="0 0 24 24"><path d="M5 4h11l3 3v12a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1zM8 4v5h7V4M8 20v-6h8v6"/></svg>',
  dl: '<svg viewBox="0 0 24 24"><path d="M12 4v11m0 0-4.5-4.5M12 15l4.5-4.5M4 19h16"/></svg>',
  reset: '<svg viewBox="0 0 24 24"><path d="M4.5 12a7.5 7.5 0 1 0 2.2-5.3M4.5 4v3.5H8"/></svg>',
  search: '<svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="6.5"/><path d="m20 20-4.2-4.2"/></svg>',
  x: '<svg viewBox="0 0 24 24"><path d="M7 7l10 10M17 7 7 17"/></svg>',
  trash:
    '<svg viewBox="0 0 24 24"><path d="M4.5 7h15M10 11v6m4-6v6M6.5 7l.8 12.2a1 1 0 0 0 1 .8h7.4a1 1 0 0 0 1-.8L17.5 7M9.5 7V4.5h5V7"/></svg>',
  au: '<svg viewBox="0 0 24 24"><path d="m6.5 14.5 5.5-5.5 5.5 5.5"/></svg>',
  ad: '<svg viewBox="0 0 24 24"><path d="m6.5 9.5 5.5 5.5 5.5-5.5"/></svg>',
  chev: '<svg viewBox="0 0 24 24"><path d="m6.5 9.5 5.5 5.5 5.5-5.5"/></svg>',
  check: '<svg viewBox="0 0 24 24"><path d="m5.5 12.5 4.5 4.5 8.5-9.5"/></svg>',
  back: '<svg viewBox="0 0 24 24"><path d="M19 12H5m0 0 6-6m-6 6 6 6"/></svg>',
  sheet:
    '<svg viewBox="0 0 48 48"><path d="M28 5H13a3 3 0 0 0-3 3v32a3 3 0 0 0 3 3h22a3 3 0 0 0 3-3V15z" fill="#eef2f7" stroke="none"/><path d="M28 5v10h10" fill="#c8d2e0" stroke="none"/><rect x="4" y="17" width="20" height="18" rx="2.5" fill="#21a366" stroke="none"/><path d="m9 21 10 10m0-10L9 31" stroke="#fff" stroke-width="2.6"/><path d="M27 22h7M27 27h7M27 32h7" stroke="#7c8aa0" stroke-width="2"/></svg>',
  sun: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="4"/><path d="M12 2.5v2.2M12 19.3v2.2M2.5 12h2.2M19.3 12h2.2M5.3 5.3l1.6 1.6M17.1 17.1l1.6 1.6M5.3 18.7l1.6-1.6M17.1 6.9l1.6-1.6"/></svg>',
  moon: '<svg viewBox="0 0 24 24"><path d="M19.5 14.6A8 8 0 0 1 9.4 4.5a8 8 0 1 0 10.1 10.1z"/></svg>',
  tour: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="8.5"/><path d="m15.4 8.6-2.1 4.7-4.7 2.1 2.1-4.7z"/></svg>',
  book: '<svg viewBox="0 0 24 24"><path d="M4 5.5A1.5 1.5 0 0 1 5.5 4H11v16H5.5A1.5 1.5 0 0 1 4 18.5zM20 5.5A1.5 1.5 0 0 0 18.5 4H13v16h5.5a1.5 1.5 0 0 0 1.5-1.5z"/></svg>',
  how: '<svg viewBox="0 0 24 24"><rect x="3" y="4" width="7" height="6" rx="1.5"/><rect x="14" y="14" width="7" height="6" rx="1.5"/><path d="M10 7h3.5A2.5 2.5 0 0 1 16 9.5V14"/></svg>',
  warn: '<svg viewBox="0 0 24 24"><path d="M12 3.8 2.6 20h18.8z"/><path d="M12 10v4.6M12 17.3v.2"/></svg>',
  okc: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="8.5"/><path d="m8.2 12.3 2.6 2.6 5-5.4"/></svg>',
  info: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="8.5"/><path d="M12 11v5.5M12 7.8v.2"/></svg>',
};

/* ----- จำค่าในเครื่อง (ถ้าเบราว์เซอร์ไม่ให้เก็บ ใช้ค่าชั่วคราวแทน) ----- */
const MEM = {},
  store = {
    get: k => {
      try {
        const v = localStorage.getItem(k);
        return v == null ? (MEM[k] ?? null) : v;
      } catch (_) {
        return MEM[k] ?? null;
      }
    },
    set: (k, v) => {
      MEM[k] = v;
      try {
        localStorage.setItem(k, v);
      } catch (_) {}
    },
  };
let THEME = document.documentElement.dataset.theme === 'dark' ? 'dark' : 'light';
const setTheme = t => {
  THEME = t;
  document.documentElement.dataset.theme = t;
};
try {
  matchMedia('(prefers-color-scheme: dark)').addEventListener('change', e => {
    if (!store.get('ps_theme')) {
      setTheme(e.matches ? 'dark' : 'light');
      view();
    }
  });
} catch (_) {}

OP.col1 = store.get('ps_col1') || '';
OP.col2 = store.get('ps_col2') || '';
let adv = false,
  HELP = false,
  HOW = false,
  PASTE = false,
  WELCOME = !store.get('ps_welcomed'),
  TOUR = null,
  pendTour = null,
  SCANNING = false,
  LIT = new Set();
const RO = new Set();
let FOCUSSEL = null,
  TOURWAS = false,
  BUSY = null,
  FILLING = false,
  BASEERR = '',
  TQ = '',
  FLOPEN = false,
  RETF = null,
  FOCUSNEW = null,
  MDOWN = null,
  PTR = false,
  DEFER = false,
  TDRAWN = '',
  OVK = '';
const TH = [
  'ม.ค.',
  'ก.พ.',
  'มี.ค.',
  'เม.ย.',
  'พ.ค.',
  'มิ.ย.',
  'ก.ค.',
  'ส.ค.',
  'ก.ย.',
  'ต.ค.',
  'พ.ย.',
  'ธ.ค.',
];
const thd = d => (d ? `${d.getUTCDate()} ${TH[d.getUTCMonth()]} ${d.getUTCFullYear() + 543}` : '-');
const thdS = d => (d ? `${d.getUTCDate()} ${TH[d.getUTCMonth()]}` : '-');
const end = d =>
  d
    ? `${d.getUTCDate()}-${'JanFebMarAprMayJunJulAugSepOctNovDec'.substr(d.getUTCMonth() * 3, 3)}-${String(d.getUTCFullYear()).slice(2)}`
    : '';
function say(t) {
  const m = $('#msg');
  m.textContent = t;
  m.classList.add('show');
  clearTimeout(say.t);
  say.t = setTimeout(() => m.classList.remove('show'), 4500);
}
const esc2 = v => esc(v ?? '');
const OPDEF = {
  hand: true,
  media: 'ทั้งหมด',
  sort: 'orig',
  newOnly: false,
  dedupe: true,
  off: false,
  r2: 'cont',
};
const OPNAME = {
  hand: 'จัดรูปแบบตามใบเช็คจริง',
  media: 'สื่อ',
  sort: 'เรียงลำดับ',
  newOnly: 'เฉพาะหนัง “ขึ้นใหม่”',
  dedupe: 'ตัดชื่อซ้ำในจอเดียวกัน',
  off: 'รวมรายการปลดต่อท้าย',
  r2: 'วันที่ 16 เป็นต้นไป',
};
const advChg = () => optChanged().filter(k => k !== 'hand').length;
const optChanged = () =>
  Object.keys(OPDEF)
    .filter(k => (k !== 'sort' || !OP.hand) && (k !== 'r2' || OP.hand))
    .filter(k => OP[k] !== OPDEF[k]);
/* ----- สีแถบวันที่: ค่ากลาง (OP.col1/col2) และแยกรายไฟล์ (จำไว้ในเครื่องตามชื่อไฟล์) ----- */
const COLS = [
  ['', 'ตามชีต'],
  ['FF002060', 'น้ำเงินเข้ม'],
  ['FF4472C4', 'ฟ้า'],
  ['FF2F5597', 'ฟ้าเข้ม'],
  ['FF375623', 'เขียว'],
];
const swHex = c => (c ? '#' + String(c).slice(-6) : '');
const colName = c => (COLS.find(x => x[0] === c) || [0, 'สีที่เลือกเอง'])[1];
const fkey = n => 'ps_fcol:' + String(n).replace(/\.(xlsm|xlsx)$/i, '');
const fcol = n => {
  try {
    const o = JSON.parse(store.get(fkey(n)) || '{}');
    return { ...o, 1: o[1] || '', 2: o[2] || '' };
  } catch (_) {
    return { 1: '', 2: '' };
  }
}; /* k16 = ต่อสัญญาวันที่ 16 แถวเดิมไม่ใส่ออก */
const effCol = (r, f) => (f && fcol(f.name)[r]) || OP['col' + r] || '';
const shownCol = (r, f) => effCol(r, f) || (r === 2 ? 'FF4472C4' : 'FF002060');
function vColPick(k) {
  const cur = OP['col' + k],
    custom = cur && !COLS.some(x => x[0] === cur);
  return `<div class="swl" role="group" aria-label="สีแถบรอบ ${k === 2 ? '16–สิ้นเดือน' : '1–15'}">${COLS.map(([c, t]) => `<button class="swb ${cur === c ? 'on' : ''}" data-a="col" data-k="${k}|${c}" aria-pressed="${cur === c}" title="${t}">${c ? `<i style="background:${swHex(c)}"></i>` : '<i class="auto"></i>'}<span>${t}</span></button>`).join('')}<button class="swb ${custom ? 'on' : ''}" data-a="colc" data-k="${k}" aria-pressed="${!!custom}" title="เลือกสีเอง"><i class="${custom ? '' : 'rb'}" style="${custom ? 'background:' + swHex(cur) : ''}"></i><span>เลือกเอง</span></button></div>`;
}

/* ----- วันที่ทำงานที่พิมพ์เอง: ตรวจให้เข้มงวด รองรับปี พ.ศ. 2 หลัก (69 = 2569) และเดือนภาษาไทย ----- */
function parseUserDate(s) {
  s = String(s).trim().replace(/\s+/g, ' ');
  let d,
    m,
    y = null,
    q;
  if ((q = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/))) {
    y = yy(q[1]);
    m = +q[2] - 1;
    d = +q[3];
  } else if ((q = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2}|\d{4})$/))) {
    d = +q[1];
    m = +q[2] - 1;
    y = yy(q[3]);
  } else if ((q = s.match(/^(\d{1,2})[-\/ ]?([A-Za-z]{3,9})\.?(?:[-\/ ]?(\d{2}|\d{4}))?$/))) {
    const k = q[2].slice(0, 3).toLowerCase();
    if (!(k in MON)) return null;
    d = +q[1];
    m = MON[k];
    if (q[3]) y = yy(q[3]);
  } else if (
    (q = s
      .replace(/\s/g, '')
      .match(
        /^(\d{1,2})-?(ม\.?ค\.?|ก\.?พ\.?|มี\.?ค\.?|เม\.?ย\.?|พ\.?ค\.?|มิ\.?ย\.?|ก\.?ค\.?|ส\.?ค\.?|ก\.?ย\.?|ต\.?ค\.?|พ\.?ย\.?|ธ\.?ค\.?)-?(\d{2}|\d{4})?$/,
      ))
  ) {
    const k = Object.keys(THM).find(x => x.replace(/\./g, '') === q[2].replace(/\./g, ''));
    if (!k) return null;
    d = +q[1];
    m = THM[k];
    if (q[3]) y = yy(q[3]);
  } else return null;
  if (y == null) y = (PLD || new Date()).getUTCFullYear();
  if (!(m >= 0 && m <= 11) || d < 1 || y < 2000 || y > 2100) return null;
  const D = new Date(Date.UTC(y, m, d));
  return D.getUTCMonth() === m && D.getUTCDate() === d ? D : null;
}

/* ----- Playlist 11 แบบที่ระบบใช้ ----- */
const KINDS = [
  { k: '40Blue', c: 'b', g: 1, has: s => s.m === '40Blue' },
  { k: '40Red', c: 'r', g: 1, has: s => s.m === '40Red' },
  { k: '80Blue', c: 'b', g: 1, has: s => s.m === '80Blue' },
  { k: '80Red', c: 'r', g: 1, has: s => s.m === '80Red' },
  { k: 'CITY', c: 'c', g: 1, has: s => s.m === 'CITY' },
  { k: 'Spare', c: 's', g: 1, has: s => s.m === 'Spare' },
  { k: 'Signature Max', c: 'g', has: s => s.m === 'Signature Max' },
  { k: 'Show DC', c: 'p', latest: true, tip: 'ไม่ได้ส่งทุกวัน', has: s => s.m === 'Show DC / Prudential' },
  { k: 'BKJ1', c: 'j', has: s => s.m === 'Bangkok Jams' && s.c === 'BKJ1 ALL' },
  { k: 'BKJ2', c: 'j', has: s => s.m === 'Bangkok Jams' && /^BKJ2/.test(s.c) },
  { k: 'Metro LED', c: 'm', tip: 'มาทาง LINE', has: s => s.m === 'MetroLED' },
  { k: 'Rama9 Connect', c: 'g', x: 1, has: s => s.m === 'Rama9 Connect' },
  { k: 'Digital MRT', c: 'b', x: 1, has: s => s.m === 'Digital MRT' },
  { k: 'Gateway X', c: 'p', x: 1, has: s => s.m === 'Gateway X' },
  { k: 'Air Traveller', c: 'c', x: 1, has: s => s.m === 'Air Traveller' },
  { k: 'The 20', c: 'r', x: 1, has: s => s.m === 'The 20' },
  { k: 'LA PlanB', c: 'm', x: 1, latest: true, tip: 'ไม่ได้ส่งทุกวัน', has: s => s.m === 'LA PlanB' },
];
const kindOn = k => {
  const K = KINDS.find(x => x.k === k);
  return !!K && DB.some(K.has);
};
const kindOfS = s => (KINDS.find(K => K.has(s)) || {}).k || s.m;
function kState(K, wd) {
  const L = DB.filter(K.has);
  if (!L.length) return { on: false };
  const fm = new Map();
  for (const s of L) if (s.pf) fm.set(s.pf, Math.max(fm.get(s.pf) || 0, s.pd || 0));
  const d = Math.max(0, ...fm.values()),
    ads = L.reduce((n, s) => n + s.rows.filter(r => !r.off).length, 0);
  let st = 'ok';
  if (wd && d && d !== +wd) st = K.latest ? (d < +wd ? 'old' : 'next') : 'mis';
  return { on: true, d, scr: L.length, ads, st };
}

/* ----- ใบเช็คแต่ละไฟล์ใช้ Playlist แบบไหน: อ่านโครงไฟล์ครั้งเดียว แล้วจับคู่จอด้วยกติกาเดียวกับ fillWb ----- */
function sheetNeeds(buf, name) {
  const u = new Uint8Array(buf, 0, Math.min(4, buf.byteLength));
  if (!(u[0] === 0x50 && u[1] === 0x4b)) return { bad: 'ไม่ใช่ไฟล์ Excel (.xlsm / .xlsx) หรือไฟล์เสีย' };
  let wb;
  try {
    wb = XLSX.read(buf.slice(0), { type: 'array', sheetRows: 15 });
  } catch (e) {
    return { bad: 'ไฟล์เสียหรือเปิดไม่ได้' };
  }
  if (!wb || !wb.SheetNames || !wb.SheetNames.length) return { bad: 'ไม่พบชีตในไฟล์นี้' };
  if (
    wb.SheetNames.includes('Index') ||
    wb.SheetNames.some(n => /^รายการหนังโฆษณา/.test(cl(((wb.Sheets[n] || {}).A1 || {}).v))) ||
    isLaWb(wb) ||
    /สรุปรายการหนัง|รายการหนัง/.test(String(name || ''))
  )
    return { bad: 'ไฟล์นี้เป็นไฟล์ Playlist ไม่ใช่ใบเช็ครูป (ใส่ไว้ในขั้นที่ 1)' };
  const S = [],
    named = [],
    avOf = ws => avCols(ws).sort((x, y) => x - y);
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    if (!ws || notChk(sn)) continue;
    const media = mediaOf(sn),
      blocks = [];
    let brow = 0;
    for (const rw of [11, 10, 12]) {
      if (blocks.length) break;
      for (const k of Object.keys(ws)) {
        const m = k.match(new RegExp('^([A-Z]+)' + rw + '$')),
          c = m && codeOf(ws[k].v);
        if (c) blocks.push([col2n(m[1]), c]);
      }
      if (blocks.length) brow = rw;
    }
    if (blocks.length && !media && !avOf(ws).length) blocks.length = 0;
    const names = [];
    if (media && brow === 11)
      for (const cn of avOf(ws)) {
        if (blocks.some(x => x[0] === cn)) continue;
        const raw = cl((ws[n2col(cn) + '11'] || {}).v);
        if (raw && nrm(raw).length >= 3) names.push(raw);
      }
    if (blocks.length && avRowOf(ws) === 10) {
      const av = avOf(ws);
      for (const cn of av.filter(c => !av.includes(c - 1)))
        if (!blocks.some(x => x[0] === cn))
          named.push({ sn, label: lab10(ws, sn, cn), code: txt10(ws, cn).join(' '), k: kind10(ws, sn, cn) });
    }
    if (blocks.length && !media && avRowOf(ws) === 10) {
      const G = new Map();
      for (const [cn, c] of blocks) {
        const m = mediaOf(cl((ws[n2col(cn - 1) + '11'] || {}).v));
        (G.get(m) || G.set(m, []).get(m)).push(c);
      }
      for (const [m, cs] of G) S.push({ sn, media: m, codes: cs, names: [] });
      continue;
    }
    if (blocks.length || names.length) {
      S.push({ sn, media, codes: blocks.map(b => b[1]), names });
      continue;
    }
    if (media) continue; /* fillWb: ข้าม (ไม่พบรหัสจอ) */
    const U = sn.toUpperCase().replace(/[^A-Z0-9]/g, ''),
      av = avOf(ws);
    if (/METRO/i.test(sn.replace(/[^A-Za-z]/g, '')) && av.length && kindOn('Metro LED')) {
      S.push({ sn, kind: 'Metro LED' });
      continue;
    }
    const jm = U.match(/(?:JAMS?|BKJ)([12])([AB]|(?![0-9]))/);
    if (jm) {
      if (av.length) S.push({ sn, kind: 'BKJ' + jm[1] });
      continue;
    }
    if (/^(ANALYSIS|X+)$/.test(sn.toUpperCase().replace(/[^A-Z]/g, ''))) continue;
    if (/METRO/i.test(sn.replace(/[^A-Za-z]/g, '')) && av.length) {
      S.push({ sn, kind: 'Metro LED' });
      continue;
    }
    const r10 = avRowOf(ws) === 10;
    for (const cn of av.filter(c => !av.includes(c - 1))) {
      const lc = n2col(cn - 1);
      named.push({
        sn,
        label: r10 ? lab10(ws, sn, cn) : cl((ws[lc + '10'] || {}).v),
        code: r10
          ? txt10(ws, cn).join(' ')
          : cl((ws[n2col(cn) + '11'] || {}).v) + cl((ws[n2col(cn) + '12'] || {}).v),
        k: r10 ? kind10(ws, sn, cn) : '',
      });
    }
  }
  return { S, named, sheets: wb.SheetNames.length };
}
const guessNamed = b => {
  const n = nrm(b.label + ' ' + b.code + ' ' + b.sn);
  return /SHOWDC|SDC|PRUDENTIAL|สาทร|สาธร|TWINTUBE|TWT|DUSIT|CENTRALPARK/.test(n) ||
    (/EXCHANGETOWER/.test(n) && !/SM/.test(n))
    ? 'Show DC'
    : 'Signature Max';
};
function readyOf(f) {
  const nd = f.nd;
  if (!nd) return null;
  if (nd.bad) return { bad: nd.bad };
  const on = buildOn(OP),
    keys = [...on.keys()],
    all = OP.media === 'ทั้งหมด' ? on : buildOn({ ...OP, media: 'ทั้งหมด' }),
    M = new Map(),
    nf = [],
    warn = [],
    skip = [];
  const add = (k, sn, ok) => {
    const x = M.get(k) || { k, ok: true, sheets: new Set() };
    x.sheets.add(sn);
    x.ok = x.ok && ok;
    M.set(k, x);
  };
  for (const e of nd.S) {
    if (e.kind) {
      add(e.kind, e.sn, kindOn(e.kind));
      continue;
    }
    for (const code of e.codes) {
      if (on.has(e.media + '|' + code)) {
        add(e.media, e.sn, true);
        continue;
      }
      {
        const nk = c => c.replace(/-/g, ''),
          h = e.media
            ? keys.filter(k => k.startsWith(e.media + '|') && nk(k.split('|')[1]) === nk(code))
            : [];
        if (h.length === 1) {
          add(h[0].split('|')[0], e.sn, true);
          continue;
        }
      }
      if (e.media === 'CITY' && on.has('CITY|THE CITY')) {
        add('CITY', e.sn, true);
        continue;
      }
      const h = keys.filter(k => k.split('|')[1] === code);
      if (h.length === 1) {
        const m = h[0].split('|')[0];
        add(m, e.sn, true);
        if (e.media !== 'Spare') warn.push(`${code} (${e.sn}) ใช้รายการของ ${m}`);
        continue;
      }
      if (all !== on && [...all.keys()].some(k => k.split('|')[1] === code)) {
        skip.push(code);
        continue;
      }
      if (e.media && !kindOn(e.media)) add(e.media, e.sn, false);
      else nf.push(`${code} (${e.sn})`);
    }
    for (const raw of e.names) {
      const hit = keys.filter(k => nrm(k.split('|')[1]) === nrm(raw)),
        own = hit.filter(k => k.startsWith(e.media + '|')),
        pick = own.length === 1 ? own[0] : hit.length === 1 ? hit[0] : null;
      if (pick) add(pick.split('|')[0], e.sn, true);
      else if (!kindOn(e.media)) add(e.media, e.sn, false);
      else nf.push(`${raw} (${e.sn})`);
    }
  }
  for (const b of nd.named) {
    if (b.k && b.k !== 'Signature Max') {
      add(b.k, b.sn, kindOn(b.k));
      continue;
    }
    const s = namedScreen(b.label, b.code, b.sn);
    if (s) {
      add(kindOfS(s), b.sn, true);
      continue;
    }
    const g = guessNamed(b);
    if (!kindOn(g)) add(g, b.sn, false);
    else nf.push(`${b.label || b.code || b.sn} (${b.sn})`);
  }
  const ord = k => {
    const i = KINDS.findIndex(K => K.k === k);
    return i < 0 ? 99 : i;
  };
  const items = [...M.values()]
      .map(x => ({ ...x, sheets: [...x.sheets] }))
      .sort((a, b) => ord(a.k) - ord(b.k)),
    miss = items.filter(x => !x.ok).length;
  return {
    items,
    miss,
    nf,
    warn,
    skip,
    empty: !items.length && !nf.length && !skip.length,
    ok: (items.length > 0 || skip.length > 0) && !miss && !nf.length,
    sheets: nd.sheets,
  };
}
function scanLater() {
  if (SCANNING || BUSY) return;
  const f = C.find(x => !x.nd);
  if (!f) return;
  SCANNING = true;
  setTimeout(() => {
    try {
      f.nd = sheetNeeds(f.buf, f.name);
    } catch (e) {
      f.nd = { bad: 'ไฟล์เสียหรือเปิดไม่ได้' };
    }
    SCANNING = false;
    if (tab === 'fl') softView();
    else scanLater();
  }, 30);
}

/* ----- ชิ้นส่วนหน้าจอ ----- */
const STEPS = [
  ['dat', 'ใส่ Playlist'],
  ['pl', 'ตรวจ / แก้ไข'],
  ['fl', 'เติมใบเช็ครูป'],
  ['res', 'ผลการเติม'],
];
const stepKey = () => (tab === 'tl' ? 'pl' : tab);
const tourKey = () => (tab === 'tl' ? 'pl' : tab === 'pl' && cs != null && DB[cs] ? 'ed' : tab);
const stepIx = () => STEPS.findIndex(s => s[0] === stepKey());
function stepSub(k) {
  if (k === 'dat') {
    const n = KINDS.filter(K => !K.x && DB.some(K.has)).length;
    return DB.length ? `${n} จาก 11 แบบ` : 'ยังไม่ได้ใส่';
  }
  if (k === 'pl') {
    const e = DB.filter(s => s.ed).length;
    return e ? `แก้แล้ว ${e} จอ` : 'ไม่บังคับ';
  }
  if (k === 'fl') return C.length ? `ใบเช็ค ${C.length} ไฟล์` : 'ยังไม่ได้ใส่ใบเช็ค';
  return last ? `เติมแล้ว ${last.res.filter(r => r.blob).length} ไฟล์` : 'รอเติม';
}
const stepDone = (k, i, cur) =>
  i < cur &&
  (k === 'dat'
    ? DB.length > 0
    : k === 'pl'
      ? DB.some(s => s.ed)
      : k === 'fl'
        ? C.length > 0 || !!last
        : true);
const vSteps = () => {
  const cur = stepIx();
  return STEPS.map(([k, t], i) => {
    const dn = stepDone(k, i, cur);
    return `<button class="bs ${i === cur ? 'on' : dn ? 'done' : ''}" data-a="tab" data-k="${k}" ${k === 'res' && !last ? 'disabled' : ''} ${i === cur ? 'aria-current="step"' : ''}><span class="n">${dn ? IC.check : i + 1}</span><span><b>${t}</b><small>${stepSub(k)}</small></span></button>`;
  }).join('');
};
const vNav = () => {
  const tl = THEME === 'dark' ? 'เปลี่ยนเป็นโหมดสว่าง' : 'เปลี่ยนเป็นโหมดมืด';
  return `<button class="btn ghost" data-a="how" aria-label="ระบบทำงานอย่างไร" title="ระบบทำงานอย่างไร">${IC.how}<span class="lbl">ระบบทำงานอย่างไร</span></button><button class="btn ghost" data-a="tour" aria-label="แนะนำการใช้" title="แนะนำการใช้หน้านี้">${IC.tour}<span class="lbl">แนะนำการใช้</span></button><button class="btn ghost" data-a="help" aria-haspopup="dialog" aria-label="วิธีใช้" title="วิธีใช้">${IC.book}<span class="lbl">วิธีใช้</span></button><button class="btn secondary tg" data-a="theme" aria-label="${tl}" title="${tl}">${THEME === 'dark' ? IC.sun : IC.moon}</button>`;
};
const shead = (t, p, extra = '') =>
  `<div class="shead"><div><h1>${t}</h1>${p ? `<p>${p}</p>` : ''}</div>${extra}</div>`;
const PICK = { pl: ['.xlsx,.xlsm', true], ck: ['.xlsx,.xlsm', true], pj: ['.json', false] };
function openPicker(k) {
  const p = $('#picker');
  if (!p || !PICK[k]) return;
  p.dataset.k = k;
  p.accept = PICK[k][0];
  p.multiple = PICK[k][1];
  p.value = '';
  p.click();
}
const drop = (k, t, s, acc, tour, sm) =>
  `<div class="drop${sm ? ' sm' : ''}" data-dz="${k}" data-a="pick" data-k="${k}" data-tour="${tour}" tabindex="0" role="button" aria-label="${esc(t)} (กด Enter เพื่อเลือกไฟล์)"><span class="dic">${k === 'ck' ? IC.sheet : IC.up}</span><span class="dt"><b>${t}</b><span class="ds">${s}</span></span></div>`;
const alertHTML = (k, t, body) =>
  `<div class="alert ${k}" role="${k === 'er' ? 'alert' : 'status'}">${k === 'ok' ? IC.okc : k === 'info' ? IC.info : IC.warn}<b>${t}</b><div>${body}</div></div>`;
const rndHTML = wd => {
  const pr = per(wd),
    c = shownCol(pr.two ? 2 : 1);
  return `<span class="rnd"><i style="background:${swHex(c)}">${end(wd)}</i>${pr.two ? (OP.r2 === 'new' ? `รอบ 16–${pr.E.getUTCDate()} ${TH[pr.E.getUTCMonth()]} เริ่มรายการใหม่ ต่อท้ายรอบ 1–15 ในชีตเดิม` : `วันที่ 16–${pr.E.getUTCDate()} ${TH[pr.E.getUTCMonth()]} ทำต่อในรายการเดิม (แถบวันที่สีรอบ 2)`) : `รอบ 1–15 ${TH[pr.E.getUTCMonth()]}`}${effCol(pr.two ? 2 : 1) ? '' : ' (สีตามชีต)'}</span>`;
};
const vBusy = () =>
  BUSY
    ? `<div class="busy" role="status" aria-live="polite"><div class="bbox"><b>${BUSY.i ? `${BUSY.t} ${BUSY.i} จาก ${BUSY.n} ไฟล์` : `${BUSY.t} ${BUSY.n} ไฟล์ กำลังเตรียม…`}</b><div class="bar"><i style="width:${Math.round((100 * Math.max(0, BUSY.i - 1)) / Math.max(1, BUSY.n))}%"></i></div><small>${esc2(BUSY.f)}</small><small>ไฟล์ใหญ่อาจใช้เวลาหลายวินาที ระหว่างนี้หน้าจออาจขยับไม่ได้ ไม่ต้องกดซ้ำ</small></div></div>`
    : '';
function busyDone() {
  if (BUSY && !FILLING) {
    BUSY = null;
    $('.wrap').inert = false;
  }
}
async function busyStep(t, i, n, f) {
  BUSY = { t, i, n, f };
  const b = $('#busy');
  if (b) b.innerHTML = vBusy();
  $('.wrap').inert = true;
  await new Promise(r => requestAnimationFrame(() => setTimeout(r, 0)));
}

/* ----- ขั้นที่ 1 ----- */
function vDateCard(wd, has) {
  if (!has)
    return `<aside class="panel datecard dc0" data-tour="date"><span class="lbl">Playlist วันที่</span><b class="big">ยังไม่มี Playlist</b><span class="src">วันที่จะขึ้นเมื่อใส่ไฟล์ ระบบอ่านจากชื่อไฟล์ให้เอง</span></aside>`;
  return `<aside class="panel datecard" data-tour="date"><span class="lbl">Playlist วันที่</span><b class="big">${wd ? thd(wd) : 'ไม่พบวันที่'}</b><span class="src">${OP.base ? 'ตามที่พิมพ์ไว้ในขั้นที่ 3' : PLD ? 'จากไฟล์ Playlist ถ้าไม่ตรง แก้ได้ในขั้นที่ 3' : wd ? 'คาดจากวันเริ่มของหนังในไฟล์ ตรวจให้แน่ใจ หรือแก้ในขั้นที่ 3' : 'ใส่วันที่เองได้ในขั้นที่ 3'}</span>${BASEERR ? `<span class="src" style="color:var(--er);font-weight:600">วันที่ที่พิมพ์ในขั้นที่ 3 อ่านไม่ได้ ตอนนี้ใช้วันที่จาก Playlist</span>` : ''}${wd ? rndHTML(wd) : ''}</aside>`;
}
function vWall(wd) {
  const SX = KINDS.filter(K => K.x).map(K => ({ K, s: kState(K, wd) })),
    S = KINDS.filter(K => !K.x).map(K => ({ K, s: kState(K, wd) })),
    n = S.filter(x => x.s.on).length,
    miss = S.filter(x => !x.s.on).map(x => x.K.k),
    mis = S.filter(x => x.s.on && x.s.st === 'mis').map(x => x.K.k);
  const now = new Set(S.filter(x => x.s.on).map(x => x.K.k)),
    pow = new Set([...now].filter(k => !LIT.has(k)));
  LIT = now;
  const tile = ({ K, s }) => {
    let st,
      w = '';
    if (!s.on) st = K.tip ? `ยังไม่ได้ใส่ (${K.tip})` : 'ยังไม่ได้ใส่';
    else if (s.st === 'mis') {
      w = ' wr';
      st = 'วันที่ไม่ตรงวันทำงาน';
    } else if (s.st === 'next') {
      w = ' wr';
      st = 'ไฟล์ของรอบถัดไป ตรวจว่าตั้งใจใช้';
    } else if (s.st === 'old') st = 'ใช้ฉบับล่าสุดที่มี';
    else st = `${s.scr.toLocaleString()} จอ ${s.ads.toLocaleString()} รายการ`;
    return `<div class="tile k-${K.c}${s.on ? ' on' : ''}${w}${pow.has(K.k) ? ' pow' : ''}" role="listitem" aria-label="${K.k}: ${s.on && s.d ? thd(new Date(s.d)) + ' ' : ''}${st}"><span class="led" aria-hidden="true"></span><b>${K.k}</b>${s.on && s.d ? `<small>${thd(new Date(s.d))}</small>` : ''}<span class="st">${st}</span></div>`;
  };
  const extra = [...new Set(DB.filter(s => !KINDS.some(K => K.has(s))).map(s => s.m))];
  const sum =
    n === 11 && !mis.length
      ? `<div class="wsum ok">${IC.okc}<b>ครบ 11 แบบแล้ว</b><span>ไปเติมใบเช็ครูปได้เลย</span></div>`
      : n
        ? `<div class="wsum">${miss.length ? `<b>ยังขาด ${miss.length} แบบ</b><span>${miss.join(', ')}</span>` : '<b>ครบ 11 แบบ</b>'}${mis.length ? `<b>วันที่ไม่ตรง ${mis.length} แบบ</b><span>${mis.join(', ')} (ช่องไฟเหลือง)</span>` : ''}</div>`
        : `<div class="wsum"><span>ยังไม่ได้ใส่ Playlist ช่องจะสว่างทีละแบบเมื่อใส่ไฟล์</span></div>`;
  return `<section class="wall" data-tour="wall" aria-label="เช็กลิสต์ Playlist"><div class="wtop"><h2>เช็กลิสต์ Playlist</h2><span class="wcount"><b>${n}</b>จาก 11 แบบ</span></div>
 <div class="wg"><p>สรุปรายการหนังแยกจอ 6 ไฟล์</p><div class="wgrid" role="list">${S.filter(x => x.K.g)
   .map(tile)
   .join('')}</div></div>
 <div class="wg o"><p>Playlist อื่น 5 ไฟล์</p><div class="wgrid" role="list">${S.filter(x => !x.K.g)
   .map(tile)
   .join(
     '',
   )}</div></div>${SX.some(x => x.s.on) ? `<div class="wg o"><p>Playlist เพิ่มเติม (ใบเช็ค CCTV / MRT / LA PlanB — ไม่นับใน 11 แบบ)</p><div class="wgrid" role="list">${SX.map(tile).join('')}</div></div>` : ''}
 ${sum}
 ${n && n < 11 ? `<p class="wnote">ถ้าเติมเฉพาะใบเช็คของตัวเอง ใส่เฉพาะแบบที่ใบเช็คนั้นใช้ก็ได้ ขั้นที่ 3 จะบอกว่าใบเช็คแต่ละไฟล์ต้องใช้แบบไหน</p>` : ''}
 ${mis.length ? `<p class="wnote">ไฟล์ที่ใส่ทีหลังจะแทนที่ไฟล์เดิมของจอเดียวกัน ใส่ไฟล์ของวันที่ถูกต้องอีกครั้งเพื่อแทนที่</p>` : ''}
 ${extra.length ? `<p class="wnote">ไฟล์อื่นที่อ่านได้: ${extra.map(esc).join(', ')}</p>` : ''}</section>`;
}
const vPLI = () => {
  const P = PLI.filter(x => !x.ack);
  if (!P.length) return '';
  const st = P.filter(x => x.lv === 'stop'),
    inf = P.filter(x => x.lv !== 'stop'),
    grp = a => {
      const M = new Map();
      for (const x of a) {
        if (!M.has(x.f)) M.set(x.f, []);
        M.get(x.f).push(x.t);
      }
      return `<ul>${[...M].map(([f, t]) => `<li><b>${esc(f)}</b><ul>${t.map(y => `<li>${esc(y)}</li>`).join('')}</ul></li>`).join('')}</ul>`;
    };
  return (
    (st.length
      ? alertHTML(
          'wr',
          `ด่านตรวจ Playlist: ต้องตรวจ ${st.length} เรื่องก่อนเติม`,
          grp(st) +
            '<p>ระบบยังเติมได้ แต่ส่วนนี้อาจอ่านผิด ตรวจกับไฟล์ Playlist ก่อน ถ้ารูปแบบไฟล์เปลี่ยนจริง ส่งไฟล์ให้ผู้ดูแลระบบ ตอนกดเติม ระบบจะถามยืนยันอีกครั้ง</p><button class="btn quiet" data-a="clri">ตรวจแล้ว ซ่อนรายการนี้</button>',
        )
      : '') +
    (inf.length
      ? alertHTML(
          'info',
          `ด่านตรวจ Playlist: ข้อสังเกต ${inf.length} เรื่อง`,
          grp(inf) + '<p>ระบบจัดการให้แล้ว หรือไม่กระทบการเติม แจ้งไว้ให้รู้</p>',
        )
      : '')
  );
};
function vData() {
  const has = DB.length > 0,
    wd = has ? workDate() : null,
    sdc = KINDS.filter(K => K.latest);
  const fl = (() => {
    const M = new Map();
    for (const x of DB) if (x.pf && (!M.has(x.pf) || x.pd > M.get(x.pf))) M.set(x.pf, x.pd || 0);
    return [...M];
  })();
  const isSdc = ([f]) =>
    DB.some(
      s => s.pf === f && sdc.some(K => K.has(s)),
    ); /* Show DC / LA PlanB มีช่องบอกสถานะของตัวเอง ไม่นับเป็น “หลายวันปนกัน” */
  const ds = [
    ...new Set(
      fl
        .filter(x => !isSdc(x))
        .map(x => x[1])
        .filter(Boolean),
    ),
  ].sort((a, b) => a - b);
  return `${shead('ใส่ Playlist ของวันนี้', 'ลากไฟล์ Playlist ทุกไฟล์ของวันนี้มาวางพร้อมกัน ช่องในเช็กลิสต์จะสว่างเมื่อได้ไฟล์แบบนั้นแล้ว')}
 <div class="s1">${drop('pl', has ? 'เพิ่มหรือเปลี่ยนไฟล์ Playlist' : 'ลากไฟล์ Playlist มาวางที่นี่', 'หรือคลิกเพื่อเลือกไฟล์ .xlsx / .xls เลือกหลายไฟล์พร้อมกันได้ ไฟล์ใหม่ของจอเดียวกันจะแทนที่ของเดิม', '.xlsx,.xls,.xlsm', 'drop-pl', has)}${vDateCard(wd, has)}</div>
 ${vWall(wd)}
 <div class="stack">${PLX.length ? alertHTML('er', `อ่านไม่ได้ ${PLX.length} ไฟล์ ไฟล์เหล่านี้ยังไม่ได้นำเข้าระบบ`, `<ul>${PLX.map(x => `<li><b>${esc(x.f)}</b><br>${esc(x.why)}</li>`).join('')}</ul><p>ตรวจว่าเป็นไฟล์ Playlist .xlsx ที่ถูกต้อง ไม่ใช่ PDF หรือไฟล์อื่นที่แนบมาในอีเมลเดียวกัน ถ้าถูกแล้วแต่ยังอ่านไม่ได้ ส่งไฟล์ให้ผู้ดูแลระบบ</p><button class="btn quiet" data-a="clrx">ซ่อนรายการนี้</button>`) : ''}
 ${vPLI()}
 ${ds.length > 1 ? alertHTML('wr', `ไฟล์ Playlist มีหลายวันปนกัน (${ds.map(d => thd(new Date(d))).join(', ')})`, `<p>ระบบใช้ ${thd(wd)} เป็นวันทำงาน ช่องไฟเหลืองในเช็กลิสต์คือแบบที่ไฟล์เป็นวันอื่น ใส่ไฟล์ของวันที่ถูกต้องแทน หรือกด “ล้าง Playlist (เริ่มวันใหม่)” แล้วใส่เฉพาะไฟล์ของวันนั้น</p>`) : ''}</div>
 ${fl.length ? `<details class="files" ${FLOPEN ? 'open' : ''}><summary>ไฟล์ Playlist ที่ใส่แล้ว ${fl.length} ไฟล์</summary><div class="flist">${fl.map(([f, d]) => `<span class="fchip">${esc(f)}${d ? `<em>${thd(new Date(d))}</em>` : ''}</span>`).join('')}</div></details>` : ''}
 <div class="actbar"><div class="tools">${has ? `<button class="btn quiet" data-a="savepj">${IC.save}บันทึกโปรเจกต์</button>` : ''}<button class="btn quiet" data-a="pick" data-k="pj">${IC.folder}เปิดโปรเจกต์</button>${has ? `<button class="btn quiet" data-a="expx">${IC.dl}ส่งออก Playlist (.xlsx)</button><button class="btn quiet" data-a="reset">${IC.reset}ล้าง Playlist (เริ่มวันใหม่)</button>` : ''}</div>
 <div class="go"><button class="btn secondary" data-a="tab" data-k="pl" ${has ? '' : 'disabled'}>ตรวจ / แก้ไขก่อน</button><button class="btn primary" data-a="tab" data-k="fl" data-tour="next1" ${has ? '' : 'disabled'}>ไปเติมใบเช็ครูป</button></div></div>`;
}

/* ----- ขั้นที่ 2 ----- */
const mediaChips = () =>
  ['ทั้งหมด', ...new Set(DB.map(s => s.m))]
    .map(
      m =>
        `<button class="chip ${m === fm ? 'on' : ''}" data-a="media" data-k="${esc(m)}" aria-pressed="${m === fm}">${esc(m)}</button>`,
    )
    .join('');
const badge = m =>
  `<span class="bd ${/red/i.test(m) ? 'r' : /blue/i.test(m) ? 'b' : /city/i.test(m) ? 'c' : ''}">${esc(m)}</span>`;
const scrName = s =>
  (s.n || '').startsWith(s.c) ? s.n.slice(s.c.length).replace(/^[\s_\-–]+/, '') || s.n : s.n || '';
const sub = () =>
  `<div class="seg" role="tablist" data-tour="plseg"><button role="tab" aria-selected="${tab === 'pl'}" class="${tab === 'pl' ? 'on' : ''}" data-a="tab" data-k="pl">แก้รายจอ</button><button role="tab" aria-selected="${tab === 'tl'}" class="${tab === 'tl' ? 'on' : ''}" data-a="tab" data-k="tl">แก้ทุกจอพร้อมกัน</button></div>`;
const STS =
  '<datalist id="sts"><option>ขึ้นใหม่<option>ขึ้นใหม่ไฟล์เดิม<option>ต่อสัญญา + ขึ้นใหม่ไฟล์เดิม<option>ถอด</datalist>';
const nav2 = () =>
  `<div class="go" data-tour="plnext"><button class="btn secondary" data-a="tab" data-k="dat">${IC.back}ขั้นที่ 1</button><button class="btn primary" data-a="tab" data-k="fl">ไปเติมใบเช็ครูป</button></div>`;
function vPl() {
  if (cs != null && DB[cs]) return vEd(DB[cs]);
  if (fm !== 'ทั้งหมด' && !DB.some(s => s.m === fm)) fm = 'ทั้งหมด';
  const L = DB.map((s, i) => [s, i]).filter(
    ([s]) =>
      (fm === 'ทั้งหมด' || s.m === fm) &&
      (String(s.c) + String(s.n || '')).toLowerCase().includes(q.toLowerCase()),
  );
  return `${shead('ตรวจ / แก้ไข Playlist', 'ขั้นนี้ไม่บังคับ ใช้เมื่อต้องปลด เพิ่ม หรือเปลี่ยนชื่อหนังก่อนเติม ถ้า Playlist ถูกต้องแล้ว ข้ามไปเติมใบเช็ครูปได้เลย', nav2())}
 <section class="panel">${sub()}<div class="chips" data-tour="chips">${mediaChips()}</div>
 <label class="search">${IC.search}<input id="q" placeholder="ค้นหาจอ เช่น DGT009 หรือ สีลม" value="${esc(q)}" aria-label="ค้นหาจอ"></label><p class="count">${L.length.toLocaleString()} จอ</p>
 ${
   L.length
     ? `<div class="cards">${L.slice(0, 120)
         .map(([s, i]) => {
           const on = s.rows.filter(r => !r.off).length;
           return `<button class="scr" data-a="open" data-k="${i}"><span class="top"><b>${esc(s.c)}</b>${badge(s.m)}</span><span class="nm">${esc(scrName(s))}</span><span class="nums"><span><small>ออนแอร์</small><em>${on}</em></span><span><small>ปลด</small><i>${s.rows.length - on}</i></span></span>${s.ed ? '<span class="pill ok">แก้แล้ว</span>' : ''}</button>`;
         })
         .join('')}</div>`
     : `<p class="empty">${DB.length ? 'ไม่พบจอที่ตรงกับคำค้น ลองพิมพ์รหัสจอ เช่น DGT147' : 'ยังไม่ได้ใส่ Playlist กลับไปขั้นที่ 1 ก่อน'}</p>`
 }
 ${L.length > 120 ? '<p class="hint">แสดง 120 จอแรก พิมพ์ค้นหาเพื่อกรองให้แคบลง</p>' : ''}</section>`;
}
function vEd(s) {
  const stc = t => (/ถอด/.test(t || '') ? 'bad' : /ใหม่/.test(t || '') ? 'new' : ''),
    MC = s.rows.some(r => r.cs || r.ce);
  return `${shead('แก้หนังในจอ', 'สวิตช์เปิด = ออนแอร์ ปิด = ปลด (ใส่สถานะ “ถอด” ก็ปลดให้เช่นกัน) แก้ชื่อ สถานะ ความยาว และจำนวนครั้งได้ในตาราง ใช้ลูกศรเลื่อนลำดับ หรือถังขยะลบแถว', `<div class="go" data-tour="edback"><button class="btn secondary" data-a="back">${IC.back}กลับไปรายการจอ</button></div>`)}
 <section class="panel"><div class="edhead" data-tour="edhead"><h2>${esc(s.c)}</h2><span class="enm">${esc(scrName(s))}</span>${badge(s.m)}${s.f ? `<span class="fmt">Format ${esc(s.f)}</span>` : ''}</div>${s.m === 'MetroLED' ? '<p class="hint">สถานะอ่านจากสีแถวใน Playlist: เหลือง (ไฟล์ใหม่) = ขึ้นใหม่ • ส้ม (ไฟล์เก่ากลับมาออกอากาศใหม่) = ขึ้นใหม่ไฟล์เดิม • ฟ้าอ่อน = ต่อสัญญา — ระยะสัญญาอ่านจากคอลัมน์ “ระยะสัญญา” ถ้าสัญญาเริ่มหลังวันออกอากาศ ใช้วันเริ่มสัญญาจัดแถบวันที่</p>' : ''}
 <div class="tools" data-tour="edtools"><button class="btn primary" data-a="add">+ เพิ่มหนัง</button><button class="btn quiet" data-a="paste">วางรายชื่อหลายบรรทัด</button></div>
 <div class="tbl" data-tour="edtbl"><table><thead><tr><th>ออนแอร์</th><th>ชื่อหนัง</th><th>Type</th><th>สถานะ</th><th class="n">วิ</th><th class="n">ครั้ง</th><th>หมดสัญญา</th>${MC ? '<th>ระยะสัญญา</th>' : ''}<th>หมายเหตุ</th><th><span class="sr">จัดการ</span></th></tr></thead><tbody>
 ${s.rows
   .map(
     (
       r,
       i,
     ) => `<tr class="${r.off ? 'off' : ''}"><td><label class="sw"><input type="checkbox" data-f="on" data-i="${i}" ${r.off ? '' : 'checked'} aria-label="ออนแอร์ ${esc(r.ad)}"><span></span></label></td>
 <td><input class="w-ad" data-f="ad" data-i="${i}" value="${esc2(r.ad)}" placeholder="ชื่อหนังโฆษณา" aria-label="ชื่อหนัง"></td>
 <td><input class="w-ty" data-f="ty" data-i="${i}" value="${esc2(r.ty)}" placeholder="Type" aria-label="Type"></td>
 <td class="cst ${stc(r.st)}"><input class="w-st" data-f="st" data-i="${i}" value="${esc2(r.st)}" placeholder="-" list="sts" aria-label="สถานะ"></td>
 <td class="n"><input data-f="len" data-i="${i}" value="${esc2(r.len)}" inputmode="numeric" aria-label="ความยาว"></td><td class="n"><input data-f="cnt" data-i="${i}" value="${esc2(r.cnt)}" placeholder="-" inputmode="numeric" aria-label="จำนวนครั้ง"></td>
 <td><input class="w-e" data-f="e" data-i="${i}" value="${esc2(r.e)}" placeholder="-" aria-label="หมดสัญญา"></td>${MC ? `<td class="cper">${esc(r.cs || '-')} – ${esc(r.ce || '-')}</td>` : ''}<td><input class="w-nt" data-f="note" data-i="${i}" value="${esc2(r.note)}" placeholder="-" aria-label="หมายเหตุ"></td>
 <td class="acts"><button class="ib" data-a="up" data-k="${i}" aria-label="เลื่อนขึ้น">${IC.au}</button><button class="ib" data-a="dn" data-k="${i}" aria-label="เลื่อนลง">${IC.ad}</button><button class="ib del" data-a="del" data-k="${i}" aria-label="ลบแถว">${IC.trash}</button></td></tr>`,
   )
   .join('')}
 </tbody></table></div></section>${STS}`;
}
const tlHits = () => {
  const t = TQ.toLowerCase();
  if (!t) return { n: 0, sc: 0 };
  let n = 0,
    sc = 0;
  for (const s of DB) {
    const h = s.rows.filter(r => String(r.ad).toLowerCase().includes(t)).length;
    n += h;
    if (h) sc++;
  }
  return { n, sc };
};
function vTl() {
  const H = tlHits(),
    d = H.n ? '' : 'disabled';
  return `${shead('ตรวจ / แก้ไข Playlist', 'ขั้นนี้ไม่บังคับ ค้นชื่อหนังแล้วปลด คืน ลบ หรือเปลี่ยนชื่อในทุกจอที่พบ หรือเพิ่มหนังเรื่องใหม่เข้าทุกจอของสื่อในครั้งเดียว', nav2())}
 <section class="panel">${sub()}<div class="two">
 <div class="box"><h3>ค้นหาหนังแล้วแก้ทุกจอ</h3><label class="search">${IC.search}<input id="tq" placeholder="ค้นหาชื่อหนัง" value="${esc(TQ)}" aria-label="ค้นหาชื่อหนัง"></label>
 <p class="found">${TQ ? `พบ <em>${H.n.toLocaleString()}</em> แถว ใน <em>${H.sc}</em> จอ` : 'พิมพ์ส่วนหนึ่งของชื่อหนังเพื่อค้นหา'}</p>
 <label class="fld"><span>เปลี่ยนชื่อเป็น</span><input id="rn" placeholder="ใส่ชื่อใหม่"></label>
 <div class="row3"><button class="btn g" data-a="gOff" ${d}>ปลดทุกจอ</button><button class="btn b" data-a="gOn" ${d}>คืนเป็นออนแอร์</button><button class="btn r" data-a="gDel" ${d}>ลบทิ้ง</button></div>
 <button class="btn secondary" data-a="gRn" ${d}>เปลี่ยนชื่อทุกจอที่พบ</button></div>
 <div class="box"><h3>เพิ่มหนังเข้าทุกจอของสื่อ</h3>
 <label class="fld"><span>สื่อ</span><select id="am">${[...new Set(DB.map(s => s.m))].map(m => `<option>${esc(m)}</option>`).join('')}</select></label>
 <label class="fld"><span>ชื่อหนัง</span><input id="an" placeholder="ใส่ชื่อหนัง"></label>
 <label class="fld"><span>Type</span><input id="at" placeholder="เช่น PBTV40Blue"></label>
 <label class="fld"><span>สถานะ</span><input id="as" list="sts" placeholder="เช่น ขึ้นใหม่"></label>
 <button class="btn primary" data-a="gAdd">เพิ่มเข้าทุกจอ</button><p class="hint">หนังจะต่อท้ายรายการออนแอร์ของทุกจอในสื่อที่เลือก</p></div>
 </div></section>${STS}`;
}

/* ----- ขั้นที่ 3 ----- */
function vCk(f, r, i, rk) {
  const fc = fcol(f.name)[rk],
    cust = fc && !COLS.some(x => x[0] === fc),
    sel = `<label class="fcol"><span>สีแถบรอบ ${rk === 2 ? '16–สิ้นเดือน' : '1–15'}</span><select data-fcol="${esc(f.name)}" data-r="${rk}" aria-label="สีแถบของ ${esc(f.name)}"><option value="" ${fc ? '' : 'selected'}>ตามการตั้งค่า (${colName(OP['col' + rk])})</option>${COLS.slice(
      1,
    )
      .map(([c, t]) => `<option value="${c}" ${fc === c ? 'selected' : ''}>${t}</option>`)
      .join(
        '',
      )}${cust ? `<option value="${fc}" selected>สีที่เลือกเอง</option>` : ''}</select><i class="dot" style="background:${swHex(shownCol(rk, f))}"></i></label>`;
  const rm = `<button class="ib" data-a="rm" data-k="${i}" aria-label="เอา ${esc(f.name)} ออก" title="เอาไฟล์นี้ออก">${IC.x}</button>`;
  let pill,
    body = '';
  if (!r) pill = '<span class="pill">กำลังตรวจไฟล์…</span>';
  else if (r.bad) {
    pill = '<span class="pill er">จะไม่ถูกเติม</span>';
    body = `<p class="nfline">${esc(r.bad)} เอาไฟล์นี้ออกด้วยปุ่ม × แล้วใส่ใบเช็ครูปที่ถูกต้องแทน</p>`;
  } else if (r.empty) {
    pill = '<span class="pill er">ไม่พบชีตของใบเช็ครูป</span>';
    body =
      '<p class="nfline">ไฟล์นี้อาจไม่ใช่ใบเช็ครูป (เช่น เป็นไฟล์ Playlist) ระบบจะไม่เติมอะไรในไฟล์นี้ ตรวจว่าใส่ถูกไฟล์</p>';
  } else {
    pill = r.ok
      ? `<span class="pill ok">${IC.check}พร้อมเติม</span>`
      : (r.miss ? `<span class="pill wr">ยังขาด Playlist ${r.miss} แบบ</span>` : '') +
        (r.nf.length ? `<span class="pill wr">หาไม่พบ ${r.nf.length} จอ</span>` : '');
    body = `<div class="needs">${r.items.map(x => `<span class="need${x.ok ? '' : ' miss'}" title="ชีต: ${esc(x.sheets.join(', '))}">${esc(x.k)}${x.ok ? '' : ' ยังไม่ได้ใส่'}</span>`).join('')}${r.nf
      .slice(0, 12)
      .map(x => `<span class="need miss">ไม่พบจอ ${esc(x)}</span>`)
      .join(
        '',
      )}${r.nf.length > 12 ? `<span class="need miss">และอีก ${r.nf.length - 12} จอ</span>` : ''}</div>${r.warn.length ? `<p class="nfline">ใช้รายการแทน (จะขึ้นใน “ควรตรวจ”): ${esc(r.warn.join(', '))}</p>` : ''}${r.skip.length ? `<p class="hint" style="margin-top:6px">ข้าม ${r.skip.length} จอตามตัวเลือก “สื่อ” (จะขึ้นใน “ควรตรวจ” ว่าไม่พบจอ)</p>` : ''}`;
  }
  return `<div class="ck"><div class="ckh"><span class="nm">${esc(f.name)}${r && !r.bad && !r.empty ? `<small>${r.sheets} ชีต ใช้ Playlist ${r.items.length} แบบ</small>` : ''}</span>${pill}${rm}</div>${body}${r && !r.bad && !r.empty ? sel + (rk === 2 && OP.r2 !== 'new' ? `<label class="fcol"><input type="checkbox" data-f16="${esc(f.name)}" ${fcol(f.name).k16 ? 'checked' : ''}><span>หนังต่อสัญญาวันที่ 16: แถวเดิมไม่ใส่ “ออก 16” (ใส่ออกที่แถวใต้แถบ 16) — ระบบจำไว้สำหรับไฟล์นี้</span></label>` : '') : ''}</div>`;
}
function vFl() {
  const ms = ['ทั้งหมด', ...new Set(DB.filter(s => /^(40|80|CITY|Spare)/.test(s.m)).map(s => s.m))];
  if (!ms.includes(OP.media)) OP.media = 'ทั้งหมด';
  const ab = PLD || autoBase(),
    wd = workDate();
  const R = C.map(readyOf),
    busy = R.some(r => !r),
    nBad = R.filter(r => r && (r.bad || r.empty)).length,
    nMiss = R.filter(r => r && !r.bad && !r.empty && !r.ok).length,
    nOk = R.filter(r => r && r.ok).length;
  const lines = [];
  if (!DB.length)
    lines.push(
      alertHTML(
        'wr',
        'ยังไม่ได้ใส่ Playlist',
        '<p>กลับไปขั้นที่ 1 ใส่ Playlist ของวันนี้ก่อน แล้วปุ่มเติมจะกดได้</p>',
      ),
    );
  else if (C.length && busy)
    lines.push('<p class="hint">กำลังตรวจว่าใบเช็คแต่ละไฟล์ใช้ Playlist แบบไหน…</p>');
  else if (C.length) {
    if (nBad)
      lines.push(
        alertHTML(
          'er',
          `${nBad} ไฟล์เปิดไม่ได้หรือไม่ใช่ใบเช็ครูป`,
          '<p>ไฟล์เหล่านี้จะไม่ถูกเติม ไฟล์อื่นยังเติมได้ตามปกติ เอาออกด้วยปุ่ม × แล้วใส่ไฟล์ที่ถูกต้องแทน</p>',
        ),
      );
    if (nMiss)
      lines.push(
        alertHTML(
          'wr',
          `ใบเช็ค ${nMiss} ไฟล์ยังขาด Playlist หรือหาจอไม่พบ`,
          '<p>กดเติมได้เลย แต่จอที่ไม่มีใน Playlist จะไม่ถูกเติม และจะขึ้นในรายการ “ควรตรวจ” ถ้าต้องการให้ครบ กลับไปใส่ Playlist ที่ขาดในขั้นที่ 1 ก่อน</p>',
        ),
      );
    if (nOk && !nBad && !nMiss)
      lines.push(
        alertHTML('ok', `ใบเช็ค ${nOk} ไฟล์พร้อมเติม`, '<p>มี Playlist ครบทุกจอที่ใบเช็คเหล่านี้ใช้</p>'),
      );
  }
  {
    const ps = PLI.filter(x => x.lv === 'stop' && !x.ack).length;
    if (DB.length && ps)
      lines.push(
        alertHTML(
          'wr',
          `ด่านตรวจ Playlist: ต้องตรวจ ${ps} เรื่อง`,
          '<p>ดูรายละเอียดในขั้นที่ 1 ระบบเติมได้ แต่ส่วนนั้นอาจอ่านผิด ตอนกดเติมจะถามยืนยันอีกครั้ง</p>',
        ),
      );
  }
  {
    const ch = optChanged();
    if (ch.length)
      lines.push(
        alertHTML(
          'wr',
          `ตั้งค่าไม่ตรงกับค่าแนะนำ ${ch.length} ข้อ`,
          `<p>${ch.map(k => OPNAME[k]).join(', ')} ตรวจว่าตั้งใจเปลี่ยนจริง ถ้าไม่แน่ใจ กด “คืนค่าแนะนำ”</p><button class="btn quiet" data-a="optreset">คืนค่าแนะนำ</button>`,
        ),
      );
  }
  if (OP.media !== 'ทั้งหมด')
    lines.push(
      alertHTML(
        'info',
        `ตั้งค่าให้เติมจอแยกเฉพาะสื่อ ${esc(OP.media)}`,
        '<p>จอแยกของสื่ออื่น (40/80 Blue/Red, CITY, Spare) จะไม่ถูกเติม ส่วน Signature Max, Show DC, BKJ และ Metro LED ยังเติมตามปกติ เปลี่ยนได้ใน “ตัวเลือกเพิ่มเติม”</p>',
      ),
    );
  const can = C.length && DB.length && !BASEERR;
  let dmsg;
  if (BASEERR)
    dmsg = `<small class="err">อ่านวันที่ “${esc(BASEERR)}” ไม่ได้ พิมพ์แบบ 12-Oct-26, 12/10/2569 หรือ 12 ต.ค. 69 หรือลบออกเพื่อใช้วันที่จาก Playlist</small>`;
  else if (OP.base)
    dmsg = `<small>ใช้วันที่ ${thd(wd)} ตามที่พิมพ์ ลบออกเพื่อกลับไปใช้วันที่จาก Playlist</small>${PLD && Math.abs(wd - PLD) > 20 * 864e5 ? `<small class="wrn">ห่างจากวันของ Playlist (${thd(PLD)}) มาก ตรวจอีกครั้ง</small>` : ''}`;
  else
    dmsg = ab
      ? `<small>เว้นว่าง = ใช้ ${thd(ab)} จากไฟล์ Playlist ถ้าจะเปลี่ยน พิมพ์ เช่น 12-Oct-26 หรือ 12/10/2569</small>`
      : '<small>ไม่พบวันที่ใน Playlist กรุณาใส่เอง เช่น 12-Oct-26</small>';
  return `${shead('เติมใบเช็ครูป', 'ใส่ใบเช็ครูปของวันก่อนหน้า จะใส่ของทุกคนพร้อมกัน หรือเฉพาะของคุณก็ได้ ระบบจะเติมรายชื่อหนังของวันนี้ลงไป')}
 <div class="s3"><div class="col">${drop('ck', C.length ? 'เพิ่มใบเช็ครูป' : 'ลากไฟล์ใบเช็ครูปมาวางที่นี่', 'ใบเช็คของวันก่อนหน้า (.xlsm หรือ .xlsx) หรือคลิกเพื่อเลือกไฟล์ เลือกหลายไฟล์ได้ ไฟล์ชื่อซ้ำจะแทนที่ของเดิม', '.xlsx,.xlsm', 'drop-ck', C.length > 0)}${FSA ? `<button class="btn quiet ckdir" data-a="ckdir">${IC.folder}เปิดใบเช็คทั้งโฟลเดอร์${CKDIR ? ` (ล่าสุด: ${esc(CKDIR.name)})` : ''}</button>` : ''}
 <section class="panel" data-tour="ready"><h3 class="sec-t">ใบเช็ครูปที่ใส่แล้ว<small>${C.length} ไฟล์</small></h3>
 ${C.length ? `<div class="ckl">${C.map((f, i) => vCk(f, R[i], i, wd && per(wd).two ? 2 : 1)).join('')}</div>` : '<p class="empty">ยังไม่มีใบเช็ครูป ลากไฟล์มาวางในกล่องด้านบน แล้วระบบจะบอกว่าแต่ละไฟล์ต้องใช้ Playlist แบบไหน</p>'}</section></div>
 <aside class="panel set" data-tour="settings"><h3 class="sec-t">การตั้งค่า</h3>
 <label class="date"><span>วันที่ทำงาน (วันของ Playlist)</span><input id="o_base" class="${BASEERR ? 'bad' : ''}" placeholder="${end(ab) || 'เช่น 12-Oct-26'}" autocomplete="off" aria-invalid="${!!BASEERR}">${dmsg}${wd && !BASEERR ? rndHTML(wd) + (wd.getUTCDate() === 1 ? '<small>วันที่ 1 ยังไม่ใส่ “ออกวันที่”</small>' : '') : ''}</label>
 <label class="tgl"><span>จัดรูปแบบตามใบเช็คจริง <em class="rec">แนะนำ: เปิด</em><small><b>เปิด</b> = ทำแบบที่ทีมทำมือ: มีแถบวันที่ หนัง # ขึ้นก่อน (ไฮไลต์เหลือง) แล้วเรียง A–Z หนังใหม่ลงใต้วันที่ของวันนั้น ใส่ “ออกวันที่” และย้ายวันที่ถ่าย (ช่องหลังชื่อหนัง ทุกจอ) ตามชื่อหนังให้</small><small><b>ปิด</b> = ลบรายชื่อเดิม แล้ววางรายชื่อหนังเรียงลงมาอย่างเดียว ไม่เหมาะกับไฟล์ส่งงาน</small></span><span class="sw"><input type="checkbox" id="o_hand"${OP.hand ? '' : ' aria-describedby="o_hand_d"'}><span></span></span></label>
 ${OP.hand ? '' : alertHTML('wr', 'ปิดการจัดรูปแบบอยู่', '<p id="o_hand_d">ใบเช็คจะไม่มีแถบวันที่และ “ออกวันที่” และวันที่ถ่าย (ช่องหลังชื่อหนัง) จะไม่ย้ายตามชื่อหนัง ถ้าจะส่งงานจริง เปิดสวิตช์นี้ไว้</p>')}
 <div class="cols"><p class="cl">สีแถบวันที่</p><div class="crow"><span>รอบ 1–15</span>${vColPick(1)}</div><div class="crow"><span>รอบ 16–สิ้นเดือน</span>${vColPick(2)}</div><small>“ตามชีต” = ใช้สีที่ชีตนั้นใช้อยู่ (รอบ 16 ถ้าชีตยังไม่มีแถบรอบ 2 ใช้สีฟ้า) เลือกสีแยกรายไฟล์ได้ในรายการใบเช็คด้านซ้าย ระบบจำไว้ใช้ครั้งต่อไป</small></div>
 <button class="adv" data-a="adv" aria-expanded="${adv}"><span>ตัวเลือกเพิ่มเติม${advChg() ? ` <em class="chg">เปลี่ยนจากค่าแนะนำ ${advChg()} ข้อ</em>` : ''}</span>${IC.chev}</button>
 <div class="advb" ${adv ? '' : 'hidden'}>
 <div class="opt"><label class="fld"><span>สื่อ</span><select id="o_media" aria-describedby="d_media">${ms.map(m => `<option>${esc(m)}</option>`).join('')}</select></label><small id="d_media">เติมจอแยก (40/80 Blue/Red, CITY, Spare) เฉพาะสื่อที่เลือก ชีตจอแยกของสื่ออื่นจะไม่ถูกเติม ส่วน Signature Max, Show DC, BKJ และ Metro LED เติมตามปกติ ใช้เมื่ออยากเติมใหม่แค่สื่อเดียว <b>แนะนำ: ทั้งหมด</b></small></div>
 <div class="opt"><label class="fld"><span>เรียงลำดับ</span><select id="o_sort" aria-describedby="d_sort" ${OP.hand ? 'disabled' : ''}><option value="orig">ตามไฟล์ Playlist</option><option value="az">ก-ฮ / A-Z</option><option value="new">ขึ้นใหม่ก่อน</option></select></label><small id="d_sort">${OP.hand ? 'ใช้ได้เฉพาะตอนปิด “จัดรูปแบบตามใบเช็คจริง” ตอนเปิด ระบบเรียง # ก่อน แล้ว A–Z ให้อยู่แล้ว' : 'ลำดับรายชื่อตอนวางเรียงลงมา: ตามไฟล์ Playlist / ตามตัวอักษร / หนังขึ้นใหม่อยู่บนสุด'}</small></div>
 <div class="opt"><label class="cb"><input type="checkbox" id="o_new" aria-describedby="d_new"> เฉพาะหนัง “ขึ้นใหม่”</label><small id="d_new">ใส่เฉพาะหนังที่สถานะมีคำว่า “ขึ้นใหม่” หนังที่ออนต่อเนื่องจะถูกข้าม ใช้ดูว่าวันนี้มีหนังใหม่เรื่องไหน ไม่ใช่สำหรับไฟล์ส่งงาน <b>แนะนำ: ไม่ติ๊ก</b></small></div>
 <div class="opt"><label class="cb"><input type="checkbox" id="o_dd" aria-describedby="d_dd"> ตัดชื่อซ้ำในจอเดียวกัน</label><small id="d_dd">หนังชื่อเดียวกันที่อยู่ใน Playlist 2 แถวในจอเดียว (เช่น ออน 2 ช่วงเวลา) ลงใบเช็คแถวเดียว ถ้าไม่ติ๊กจะลง 2 แถวซ้ำกัน <b>แนะนำ: ติ๊ก</b></small></div>
 <div class="opt"><label class="fld"><span>วันที่ 16 เป็นต้นไป</span><select id="o_r2" aria-describedby="d_r2" ${OP.hand ? '' : 'disabled'}><option value="cont">ทำต่อในรายการเดิม</option><option value="new">เริ่มรายการใหม่ใต้แถบ 16</option></select></label><small id="d_r2"><b>ทำต่อในรายการเดิม</b> = แบบที่ทีมทำ: หนังเดิมอยู่ที่เดิม หนังที่ต่อสัญญา / ขึ้นใหม่วันที่ 16 ลงแถวใหม่ใต้แถบ 16 และแถวเดิมใส่ “ออกวันที่ 16” พอหนังถอดจริง ใส่วันถอดทั้งสองแถว <b>เริ่มรายการใหม่</b> = แบบเวอร์ชัน 2.3: ลงหนังที่ออนอยู่ทั้งหมดใหม่ใต้แถบ 16 ต่อท้ายรอบ 1–15 <b>แนะนำ: ทำต่อในรายการเดิม</b></small></div>
 <div class="opt"><label class="cb"><input type="checkbox" id="o_off" aria-describedby="d_off"> รวมรายการปลดต่อท้าย</label><small id="d_off">เขียนรายชื่อหนังที่ปลด/ถอดแล้วต่อท้ายตาราง ใช้ตรวจว่าเรื่องไหนถอดไปแล้ว อย่าใช้กับไฟล์ที่จะส่งงาน <b>แนะนำ: ไม่ติ๊ก</b></small></div>
 <button class="btn quiet" data-a="optreset" ${optChanged().length ? '' : 'disabled'}>คืนค่าแนะนำทั้งหมด</button></div></aside></div>
 <div class="fillbar">${lines.join('')}<button class="btn primary big" data-a="fill" data-tour="fill" ${can ? '' : 'disabled'}>เติมใบเช็ครูป${C.length ? ` ${C.length} ไฟล์` : ''}</button>${BASEERR ? '<p class="hint">แก้วันที่ทำงานให้ถูกต้องก่อน แล้วปุ่มนี้จะกดได้</p>' : DB.length && !C.length ? '<p class="hint">ใส่ใบเช็ครูปอย่างน้อย 1 ไฟล์ แล้วปุ่มนี้จะกดได้</p>' : ''}</div>`;
}

/* ----- ขั้นที่ 4 ----- */
const isFilled = x => x[2].startsWith('เติมแล้ว');
const isWarn = x =>
  isFilled(x)
    ? /ใช้รายการ/.test(x[2]) && !/จอ (Spare|CITY) ใช้รายการ/.test(x[2])
    : x[1] !== 'รูปแบบ' || /^ตรวจ/.test(x[2]);
const vPic = () => {
  if (!last || !last.res.some(r => r.fold && r.fold.length)) return '';
  const D = new Date(last.W),
    two = D.getUTCDate() >= 16,
    ld = new Date(Date.UTC(D.getUTCFullYear(), D.getUTCMonth() + 1, 0)).getUTCDate(),
    rn = (two ? '16–' + ld : '1–15') + ' ' + TH[D.getUTCMonth()],
    p = last.pic,
    P = p && p.P;
  let b = '';
  if (!FSA) b = '<p class="hint">ปุ่มนี้ใช้ได้ใน Chrome / Edge บนคอมพิวเตอร์</p>';
  else if (!P)
    b = `<p class="hint">สร้างโฟลเดอร์ที่ยังไม่มีในโฟลเดอร์แยกรูป ตามรายชื่อหนังในใบเช็คที่เติมแล้ว: เดือน → รอบ ${two ? '16-' + ld : '1-15'} → #Module ขาว / ดำ / สี และหนังทุกเรื่อง${two ? 'ที่อยู่ใต้แถบ 16 ขึ้นไป' : 'ของรอบ 1–15'} ระบบไม่ลบ ไม่ย้าย และไม่เปลี่ยนชื่อโฟลเดอร์หรือรูปที่มีอยู่ ก่อนสร้างจะให้ดูก่อนว่าจะสร้างอะไรบ้าง</p><div><button class="btn secondary" data-a="picscan">${IC.folder}เลือกโฟลเดอร์แยกรูปแล้วตรวจ</button></div>`;
  else {
    const W = P.plan.filter(x => x.need.length || x.mkMon || x.mkRnd),
      nm = x => esc(x.s.path[x.s.path.length - 1]);
    b = `${p.made ? (p.made.bad.length ? alertHTML('er', `สร้างแล้ว ${p.made.n.toLocaleString()} โฟลเดอร์ สร้างไม่ได้ ${p.made.bad.length} โฟลเดอร์`, `<p>${esc(p.made.bad.slice(0, 8).join(', '))}${p.made.bad.length > 8 ? ' …' : ''}</p>`) : alertHTML('ok', `สร้างแล้ว ${p.made.n.toLocaleString()} โฟลเดอร์`, '<p>เปิดโฟลเดอร์แยกรูปดูได้เลย ถ้ามีหนังเพิ่มวันพรุ่งนี้ เติมใบเช็คแล้วกดตรวจอีกครั้ง ระบบจะสร้างเฉพาะเรื่องที่ยังไม่มี</p>')) : ''}
 <p class="hint">โฟลเดอร์ “${esc(P.dir)}” • พบโฟลเดอร์จอ ${P.plan.length} จอ${P.n ? ` • จะสร้าง ${P.n.toLocaleString()} โฟลเดอร์ใน ${W.length} จอ` : ' • ครบแล้ว ไม่มีโฟลเดอร์ที่ต้องสร้าง'}</p>
 ${
   W.length
     ? `<details class="picd" open><summary>จะสร้างในจอเหล่านี้ (${W.length} จอ)</summary><ul>${W.map(
         x =>
           `<li><b>${nm(x)}</b> ${x.need.length ? x.need.length + ' โฟลเดอร์' : ''}${x.mkMon ? ` • โฟลเดอร์เดือน ${esc(x.mon)}` : ''}${x.mkRnd ? ` • โฟลเดอร์รอบ ${esc(x.rnd)}` : ''}<small>${esc(
             x.need
               .slice(0, 6)
               .map(q => q[1])
               .join(', '),
           )}${x.need.length > 6 ? ' …' : ''}</small></li>`,
       ).join('')}</ul></details>`
     : ''
 }
 ${P.miss.length ? `<details class="picd"><summary>หาโฟลเดอร์จอไม่เจอ ${P.miss.length} จอ (ระบบไม่สร้างโฟลเดอร์จอเอง)</summary><ul>${P.miss.map(x => `<li><b>${esc(x.id)}</b><small>${esc(x.f)} • ชีต ${esc(x.sn)}</small></li>`).join('')}</ul></details>` : ''}
 ${P.closed.length ? `<p class="hint">จอปิดถาวร ข้าม: ${esc(P.closed.map(x => x.id).join(', '))}</p>` : ''}
 <div class="picb">${P.n ? `<button class="btn primary" data-a="picmake">สร้าง ${P.n.toLocaleString()} โฟลเดอร์</button>` : ''}<button class="btn secondary" data-a="picscan" data-k="again">ตรวจอีกครั้ง</button><button class="linkb" data-a="picscan" data-k="pick">เลือกโฟลเดอร์อื่น</button></div>`;
  }
  return `<section class="pic"><h3>${IC.folder}สร้างโฟลเดอร์แยกรูป รอบ ${rn}</h3>${b}</section>`;
};
const vCmp = () => {
  if (!last || !last.res.some(r => (r.fold || []).some(b => CPN[b.media]))) return '';
  const c = last.cmp,
    R = c && c.R;
  let b = '';
  const btn = `<div class="picb"><button class="btn secondary" data-a="cmpscan" data-k="again">ตรวจอีกครั้ง</button>${R && !R.err ? '<button class="btn quiet" data-a="cmpcopy">คัดลอกผลตรวจ</button>' : ''}<button class="linkb" data-a="cmpscan" data-k="pick">เลือกโฟลเดอร์อื่น</button></div>`;
  if (!FSA) b = '<p class="hint">ปุ่มนี้ใช้ได้ใน Chrome / Edge บนคอมพิวเตอร์</p>';
  else if (!R)
    b = `<p class="hint">ตรวจว่าจอในใบเช็คอยู่ในโฟลเดอร์แพคที่ถูกต้องไหม เช่น จอในชีต 80 RED ต้องมีโฟลเดอร์อยู่ใน Plan B TV 80 RED • ระบบดูอย่างเดียว ไม่ย้ายหรือแก้อะไรในเครื่อง</p><div><button class="btn secondary" data-a="cmpscan">${IC.folder}เลือกโฟลเดอร์เทียบจอแล้วตรวจ</button></div><p class="hint">ครั้งแรกให้เลือกโฟลเดอร์ “เทียบจอ Plan B TV BKK” (ที่มีโฟลเดอร์รอบ เช่น 10 Oct 16-31 อยู่ข้างใน) ครั้งต่อไประบบจำไว้</p>`;
  else if (R.err)
    b =
      alertHTML(
        'wr',
        `ไม่พบโฟลเดอร์รอบในโฟลเดอร์ “${esc(R.dir)}”`,
        `<p>เลือกโฟลเดอร์ที่มีโฟลเดอร์รอบ เช่น “${esc(R.want)}” อยู่ข้างใน (หรือเลือกโฟลเดอร์รอบนั้นเลย)</p>`,
      ) + btn;
  else {
    const grp = ([g, lb, ds]) => {
      const I = R.iss.filter(x => x.g === g);
      return I.length
        ? `<div class="cmpg g-${g}"><h4>${lb}<span>${I.length}</span></h4><p class="hint">${ds}</p><ol class="cmpl">${I.map(x => `<li><b>${esc(x.nm)}</b><span class="do">${esc(x.act)}</span>${x.now ? `<small>${esc(x.now)}</small>` : ''}</li>`).join('')}</ol></div>`
        : '';
    };
    b = `${R.prev ? alertHTML('wr', `ยังไม่มีโฟลเดอร์รอบ ${esc(R.want)}`, `<p>ระบบตรวจโฟลเดอร์รอบล่าสุด “${esc(R.rdn)}” แทน รายการด้านล่างคือสิ่งที่ต้องแก้ตอนทำโฟลเดอร์รอบใหม่</p>`) : ''}
 ${R.n ? alertHTML('wr', `ต้องแก้ ${R.n.toLocaleString()} เรื่อง`, `<p>รอบ “${esc(R.rdn)}” • ตรวจจอในใบเช็ค ${R.nk} จอ ถูกต้อง ${R.okK} จอ</p>`) : alertHTML('ok', 'ถูกต้องทั้งหมด', `<p>รอบ “${esc(R.rdn)}” • จอในใบเช็ค ${R.nk} จออยู่ในโฟลเดอร์แพคที่ถูกต้อง และโฟลเดอร์ Plan B TV 240 ตรงกับโฟลเดอร์แพค</p>`)}
 ${R.G.map(grp).join('')}
 <details class="picd"><summary>ดูจำนวนแต่ละแพค</summary><div class="cmpw"><table class="cmpt"><thead><tr><th>โฟลเดอร์แพค</th><th>โฟลเดอร์ในเครื่อง</th><th>จอในใบเช็ค</th><th>ถูกต้อง</th><th>ต้องแก้</th></tr></thead><tbody>${R.rows.map(r => `<tr><td>${esc(r.n)}</td><td>${r.nf}</td><td>${r.nk}</td><td>${r.ok}</td><td class="${r.bad ? 'bad' : ''}">${r.bad}</td></tr>`).join('')}</tbody></table></div></details>
 <p class="hint">ระบบดูอย่างเดียว ไม่ได้ย้ายหรือแก้อะไรในเครื่อง • ตรวจเฉพาะจอที่มีในใบเช็คที่เติม • แก้แล้วกด “ตรวจอีกครั้ง”${R.other.length ? ` • โฟลเดอร์อื่นในโฟลเดอร์รอบ (ข้าม): ${esc(R.other.join(', '))}` : ''}</p>${btn}`;
  }
  return `<section class="pic cmp"><h3>${IC.folder}ตรวจโฟลเดอร์เทียบจอ Plan B TV BKK</h3>${b}</section>`;
};
function vRes() {
  if (!last) return vFl();
  const R = last.res.map(r => ({
    ok: r.rep.filter(isFilled).length,
    bad: r.rep.filter(isWarn),
    inf: r.rep.filter(x => x[1] === 'รูปแบบ' && !/^ตรวจ/.test(x[2])),
    scr: r.rep.filter(isFilled),
    sh: new Set(r.rep.filter(isFilled).map(x => x[0])).size,
  }));
  const tot = R.reduce((n, r) => n + r.ok, 0),
    nb = R.reduce((n, r) => n + r.bad.length, 0),
    nf = last.res.filter(r => r.blob).length,
    ne = last.res.length - nf;
  return `${shead('ผลการเติม', last.saved ? 'บันทึกใบเช็คที่เติมแล้ว ใส่ Playlist ของวันถัดไปต่อได้เลย (กด “ใช้ผลนี้เติมวันถัดไป” ด้านล่าง)' : FSA ? 'ไฟล์ที่เติมแล้วยังไม่ถูกบันทึกลงเครื่อง กด “บันทึกทับไฟล์เดิม” หรือดาวน์โหลดก่อนปิดหน้านี้' : 'ไฟล์ที่เติมแล้วยังไม่ถูกบันทึกลงเครื่อง กดดาวน์โหลดก่อนปิดหน้านี้', `<div class="go"><button class="btn secondary" data-a="tab" data-k="fl">${IC.back}กลับไปเติมใหม่</button></div>`)}
 <div class="sum" data-tour="sum"><span class="k">เติมเสร็จ</span><span><b>${nf}</b>ไฟล์</span><span><b>${tot.toLocaleString()}</b>จอ</span><span><b>${nb}</b>ควรตรวจ</span><span class="sp"></span>${FSA ? `<button class="btn secondary big" data-a="dlz" data-tour="dlz" ${nf ? '' : 'disabled'}>${IC.dl}ดาวน์โหลด (.zip)</button><button class="btn primary big" data-a="savedir" ${nf ? '' : 'disabled'}>${IC.save}บันทึกทับไฟล์เดิม</button>` : `<button class="btn primary big" data-a="dlz" data-tour="dlz" ${nf ? '' : 'disabled'}>${IC.dl}ดาวน์โหลดทั้งหมด (.zip)</button>`}</div>
 <div class="stack">${last.sv ? (last.sv.bad.length ? alertHTML('er', `บันทึกไม่ได้ ${last.sv.bad.length} ไฟล์`, `<p>${esc(last.sv.bad.join(', '))} อาจเปิดอยู่ใน Excel ปิดไฟล์ก่อน แล้วกด “บันทึกทับไฟล์เดิม” อีกครั้ง ไฟล์อื่นบันทึกแล้ว ${last.sv.ok} ไฟล์</p>`) : alertHTML('ok', `บันทึกทับ ${last.sv.ok} ไฟล์ในโฟลเดอร์ “${esc(last.sv.dir)}” แล้ว`, `<p>ไฟล์ก่อนเติมเก็บไว้ที่โฟลเดอร์ สำรองก่อนเติม/${last.sv.day} พร้อม รายงาน.txt${last.sv.neu ? ` • ${last.sv.neu} ไฟล์ไม่มีในโฟลเดอร์เดิม จึงสร้างใหม่` : ''} <button class="linkb" data-a="savedir" data-k="pick">บันทึกไปโฟลเดอร์อื่น</button></p>`)) : ''}
 ${ne ? alertHTML('er', `${ne} ไฟล์ไม่ได้เติม`, '<p>ไฟล์เหล่านี้เสีย หรือไม่ใช่ใบเช็ครูป และไม่อยู่ในไฟล์ zip ไฟล์อื่นเติมตามปกติ ดูเหตุผลในรายการของไฟล์นั้น แล้วใส่ไฟล์ที่ถูกต้องและเติมใหม่</p>') : ''}
 ${nb ? alertHTML('wr', '“ควรตรวจ” คืออะไร', '<p>จอที่เติมไม่ได้ จอที่เติมจากรายการอื่นแทน หรือหนังที่ระบบตั้งใจไม่ใส่ เปิดดูในแต่ละไฟล์แล้วตรวจในใบเช็คก่อนใช้งาน ความหมายของข้อความที่พบบ่อยอยู่ในปุ่ม “วิธีใช้”</p>') : alertHTML('ok', 'ไม่มีรายการที่ต้องตรวจเพิ่ม', '<p>ดาวน์โหลดแล้วเปิดใบเช็คใน Excel ดูความถูกต้องอีกครั้งก่อนใช้งาน</p>')}</div>
 <div class="rlist" data-tour="rlist">${last.res
   .map((r, i) => {
     const o = R[i],
       open = RO.has(i);
     return `<div class="rfile"><div class="rhead"><span class="fn"><b>${esc(r.name)}</b><small>${r.blob ? o.sh + ' ชีต' : 'ไม่ได้เติม'}</small></span>${r.blob ? `<span class="pill ok">เติมแล้ว ${o.ok} จอ</span>` : '<span class="pill er">ไม่ได้เติม</span>'}${o.bad.length ? `<span class="pill wr">ควรตรวจ ${o.bad.length}</span>` : ''}
 <button class="ib" data-a="dlf" data-k="${i}" aria-label="ดาวน์โหลด ${esc(r.name)}" title="ดาวน์โหลดไฟล์นี้" ${r.blob ? '' : 'disabled'}>${IC.dl}</button><button class="ib tog" data-a="rx" data-k="${i}" aria-expanded="${open}" aria-label="ดูรายละเอียด ${esc(r.name)}" title="ดูรายละเอียด">${IC.chev}</button></div>
 ${
   open
     ? `<div class="issues">${o.bad.map(x => `<div class="iss"><b>${esc(x[1])}</b><span>${esc(x[0])}: ${esc(x[2])}</span></div>`).join('')}${
         o.scr.length
           ? `<p class="ish">สรุปการจัดแต่ละชีต</p>${[...new Set(o.scr.map(x => x[0]))]
               .map(sn => {
                 const L = o.scr.filter(x => x[0] === sn),
                   f = o.inf.find(x => x[0] === sn);
                 return `<div class="iss inf"><b>${esc(sn)}</b><span><span class="scrs">เติม ${L.length} จอ: ${L.map(
                   x => {
                     const m = x[2].match(/จอ (?:Spare ใช้รายการของสื่อ|CITY ใช้รายการ) ([^)]+)/);
                     return (
                       esc(String(x[1]).split(' → ')[0]) +
                       ' (' +
                       x[3] +
                       ' เรื่อง' +
                       (m ? ' จากรายการ ' + esc(m[1]) : '') +
                       ')'
                     );
                   },
                 ).join(', ')}</span>${f ? esc(f[2]) : ''}</span></div>`;
               })
               .join('')}`
           : ''
       }</div>`
     : ''
 }</div>`;
   })
   .join('')}</div>
 ${vPic()}${vCmp()}<div class="nextday"><p>${FSA ? '“บันทึกทับไฟล์เดิม” เขียนใบเช็คที่เติมแล้วทับไฟล์ในโฟลเดอร์เดิม (เลือกโฟลเดอร์ครั้งแรกครั้งเดียว) และเก็บไฟล์ก่อนเติมไว้ใน สำรองก่อนเติม ส่วน zip มีใบเช็คทุกไฟล์และ รายงาน.txt' : 'ไฟล์ zip มีใบเช็ครูปทุกไฟล์ที่เติมได้ (ชื่อเดิม) และ รายงาน.txt'} ใบเช็คที่ได้วันนี้คือไฟล์ตั้งต้นของวันพรุ่งนี้</p><button class="btn secondary" data-a="next" ${nf ? '' : 'disabled'}>ใช้ผลนี้เติมวันถัดไป</button></div>`;
}

/* ----- หน้าต่าง: วิธีใช้ / ต้อนรับ / วางรายชื่อ / ระบบทำงานอย่างไร ----- */
const vHelp =
  () => `<div class="ov"><div class="dlg" role="dialog" aria-modal="true" aria-labelledby="hlt"><div class="dlgh"><h2 id="hlt">วิธีใช้ประจำวัน</h2><button class="ib" data-a="help" aria-label="ปิด" data-focus>${IC.x}</button></div>
<ol class="hs">
<li><b>เตรียมไฟล์</b><span>Playlist ของวันนี้ 11 แบบ (สรุปรายการหนังแยกจอ 40Blue, 40Red, 80Blue, 80Red, CITY, Spare และ Signature Max, Show DC, BKJ1, BKJ2, Metro LED) กับใบเช็ครูปที่เติมแล้วของวันก่อนหน้า • ถ้ามีใบเช็ค CCTV / MRT ใส่ Playlist เพิ่มเติม Rama9 Connect, Digital MRT, Gateway X, Air Traveller, The 20 ด้วย (ไม่นับใน 11 แบบ)</span></li>
<li><b>ขั้นที่ 1 ใส่ Playlist</b><span>ถ้ามีงานของวันก่อนค้างอยู่ กด “ล้าง Playlist (เริ่มวันใหม่)” แล้วลากไฟล์ทั้งหมดลงกล่อง ดูว่าช่องในเช็กลิสต์สว่างครบ และ “Playlist วันที่” ถูกต้อง</span></li>
<li><b>ขั้นที่ 2 ตรวจ / แก้ไข (ไม่บังคับ)</b><span>ใช้เมื่อต้องปลด เพิ่ม หรือเปลี่ยนชื่อหนังก่อนเติม</span></li>
<li><b>ขั้นที่ 3 เติมใบเช็ครูป</b><span>ลากใบเช็ครูปลงกล่อง ดูว่าแต่ละไฟล์ขึ้น “พร้อมเติม” วันที่ทำงานและรอบถูกต้อง แล้วกด “เติมใบเช็ครูป”</span></li>
<li><b>ขั้นที่ 4 ผลการเติม</b><span>เปิดดูรายการ “ควรตรวจ” ของแต่ละไฟล์ แล้วกด “ดาวน์โหลดทั้งหมด (.zip)” ใบเช็คที่ได้คือไฟล์ตั้งต้นของวันพรุ่งนี้</span></li></ol>
<h3>ระบบจัดให้อย่างไร</h3><ul class="hr2">
<li>วันที่ 1 ลงหนังที่ออนอยู่ทั้งหมด หนัง # ขึ้นก่อน (ไฮไลต์เหลือง) แล้วเรียง A–Z</li>
<li>ด่านตรวจ Playlist: ระบบอ่านคอลัมน์ตามชื่อหัวตาราง ถ้าเจอสิ่งที่ไม่รู้จัก (หาคอลัมน์ไม่เจอ สถานะใหม่ วันที่อ่านไม่ได้ สื่อใหม่) จะเตือนในขั้นที่ 1 และถามยืนยันก่อนเติม หมายเหตุที่พูดถึงจอที่กำลังเติมแต่ระบบอ่านไม่ออกจะขึ้นใน “ควรตรวจ”</li>
<li>วันถัดไป หนังที่ยังไม่มีในใบเช็คและสถานะ “ขึ้นใหม่” จะลงใต้แถบวันที่ของวันนั้น หนังที่มีอยู่แล้วอยู่ที่เดิม ไม่ซ้ำ</li>
<li>หนังที่ยังอยู่ใน Playlist ถือว่ายังออนอยู่ “ออกวันที่ N” ใส่เมื่อหนังย้ายไปอยู่ในรายการถอด (วันที่ 1 ยังไม่ใส่) ไฟล์เดิม / ต่อสัญญากลางรอบ ใช้แถวเดิม ไม่เพิ่มแถวซ้ำ</li>
<li>วันที่ถ่าย (ช่องหลังชื่อหนังที่ทีมใส่เอง) ย้ายตามชื่อหนังให้เอง ทุกคอลัมน์ — จอที่มีหลายจอในบล็อกเดียว (BKJ, Metro LED) ย้ายครบทุกจอ • ขึ้นเดือนใหม่ (วันที่ 1) ล้างวันที่ถ่ายของเดือนก่อนให้ • ระบบดูแลแค่รายชื่อหนังกับ “ออกวันที่” ไม่เขียนวันที่ถ่ายเอง</li>
<li>วันที่ 16 เป็นต้นไปทำต่อในรายการเดิม แถบวันที่เป็นอีกสี (เลือกสีได้ในขั้นที่ 3) หนังที่ต่อสัญญา / ขึ้นใหม่วันที่ 16 ลงแถวใหม่ใต้แถบ 16 แถวเดิมใส่ “ออกวันที่ 16” พอหนังถอดจริง ใส่วันถอดทั้งสองแถว (ถ้าต้องการเริ่มรายการใหม่ เปลี่ยนได้ใน “ตัวเลือกเพิ่มเติม”) แถบท้ายตารางเป็นวันถัดไปเสมอ</li>
<li>อ่านหมายเหตุใน Playlist ให้: ไม่ออน / ออนเฉพาะ / เฉพาะจอ / ทุกจอยกเว้น (Signature Max, Show DC — เทียบชื่อย่อจอ เช่น SO TWT สาทร AOT BIGY), ขึ้นเฉพาะ / ยกเว้นจอ (BKJ), Tags LEDMETRO (Metro LED)</li>
<li>Metro LED: สถานะอ่านจากสีแถว (เหลือง = ขึ้นใหม่, ส้ม = ขึ้นใหม่ไฟล์เดิม, ฟ้าอ่อน = ต่อสัญญา) และอ่านระยะสัญญา — วันเริ่มใช้วันที่ช้ากว่าระหว่างเริ่มออกอากาศกับเริ่มสัญญา</li>
<li>แถว #FAST_ADS / FastAd ไม่ใช่หนังโฆษณา: ไม่ใส่จาก Playlist และลบออกจากใบเช็คตอนเติม</li><li>Playlist Rama9 Connect / Digital MRT / Gateway X (ไม่มีช่องสถานะ): ช่องชื่อหนังสีเหลือง = ขึ้นใหม่ สีชมพู = ขึ้นใหม่ไฟล์เดิม กลุ่ม AD.ออก / Ads ถอด = หนังที่ถอด • Content อาคาร / Content Plan B ไม่ลงใบเช็ค Content อื่น (เช่น Content กทพ.) ลงเป็นหนัง # • MRT ส่วนต่อขยาย 20 สถานี (ท่าพระ–หลักสอง, เตาปูน–ท่าพระ) ใช้รายการส่วนต่อขยาย สถานีอื่นใช้ 18 สถานีเดิม</li><li>Air Traveller (IN = To Airport, OUT = To City) และ The 20 อ่านแบบ Signature Max แต่แยกเป็นสื่อของตัวเอง ไม่ใส่โฆษณา Programmatic (Vistar)</li><li>ใบเช็คแบบใหม่ (CCTV / MRT / Gateway X / Rama9 Connect / Air Traveller / THE 20): ช่องหมายเหตุต่อท้ายจอสุดท้ายย้ายตามชื่อหนัง และ “ออก N” เขียนแบบเดียวกับที่ชีตใช้อยู่</li><li>LA PlanB: Playlist ส่งเฉพาะวันที่รายการเปลี่ยน ใส่ฉบับล่าสุดที่มีได้ทุกวัน (เช็กลิสต์ขึ้น “ใช้ฉบับล่าสุดที่มี”) • ชีต LA PLANB จอ LA PlanB 1 และ 2 ใช้รายการเดียวกัน หนังใน “รายการปลด” ใส่ออกวันที่เป็นวันที่ไฟล์ใหม่เริ่ม หนังใหม่ลงใต้แถบวันนั้น</li><li>แถบวันที่ทำตามแบบของแต่ละใบ: ใบที่ใส่แถบวันถัดไปรอไว้ท้ายตาราง ระบบใส่ให้ ใบที่ไม่ใส่ ระบบก็ไม่ใส่ • วันที่ในแถบของจอที่ 2 เป็นต้นไปเป็นสูตร (=C25) หรือตัวเลขตามที่ใบนั้นใช้ • หนังต่อสัญญาวันที่ 16 ปกติแถวเดิมใส่ “ออก 16” ถ้าใบไหนไม่ใส่ ติ๊กตัวเลือกใต้ชื่อไฟล์ในขั้นที่ 3 (ระบบจำไว้ตามชื่อไฟล์)</li><li>ตรวจโฟลเดอร์เทียบจอ (ขั้นที่ 4, Chrome / Edge): หลังเติมใบเช็ค กด “เลือกโฟลเดอร์เทียบจอแล้วตรวจ” ระบบหาโฟลเดอร์รอบ (เช่น 10 Oct 16-31) แล้วเทียบว่าจอในชีต 40 BLUE / 40 RED / 80 BLUE / 80 RED / CITY / Spare ของใบเช็ค อยู่ในโฟลเดอร์แพคเดียวกันไหม จอไหนไม่มีโฟลเดอร์ จอไหนมีโฟลเดอร์หลายแพค และป้าย (40B) (80R) ในโฟลเดอร์ Plan B TV 240 ตรงกับโฟลเดอร์แพคไหม • ผลแยกเป็น “ต้องแก้ในใบเช็ค” / “ต้องแก้ในโฟลเดอร์” / “ให้ตรวจเอง” แต่ละจอบอกสิ่งที่ต้องทำ (ใช้ Playlist วันนี้ตัดสินว่าฝั่งไหนผิด) • ยังไม่มีโฟลเดอร์รอบใหม่ ระบบตรวจรอบล่าสุดแทน • อ่านอย่างเดียว ไม่ย้าย ไม่สร้าง ไม่เปลี่ยนชื่อ</li><li>โฟลเดอร์แยกรูป (ขั้นที่ 4, Chrome / Edge): เลือกโฟลเดอร์แยกรูปครั้งแรกครั้งเดียว ระบบสร้างโฟลเดอร์ที่ยังไม่มีของรอบปัจจุบัน — เดือน (เช่น 10 OCT 26) → รอบ 1-15 หรือ 16-30 / 16-31 → #Module ขาว / ดำ / สี และหนังทุกเรื่องของรอบนั้นตามใบเช็ค (รอบ 16–สิ้นเดือน = หนังใต้แถบ 16 ขึ้นไป) • จับคู่จอจากรหัส DGT / ID, เลขจอ BKJ / METRO หรือชื่อจอ • ไม่ลบ ไม่ย้าย ไม่เปลี่ยนชื่อ และไม่แตะจอปิดถาวร</li><li>คำในช่องออกเขียนแบบเดียวกับที่ใบเช็คใช้อยู่ (“ออก 5” หรือ “ออกวันที่ 5”) ชีตที่ยังไม่มีหมายเหตุใช้แบบของชีตอื่นในไฟล์เดียวกัน • รหัสจอที่ขีดต่างจาก Playlist (ใบเช็ค DGT-174 / Playlist DGT174) จับคู่ให้ในสื่อเดียวกัน • หนัง Content ที่ใบเช็คเขียนไว้แบบไม่มี # ใช้แถวเดิม ไม่เพิ่มแถว # ซ้ำ</li></ul>
<h3>ข้อความ “ควรตรวจ” ที่พบบ่อย</h3><ul class="hr2">
<li><b>ไม่พบจอนี้ใน Playlist</b> ยังไม่ได้ใส่ Playlist ของสื่อนั้น หรือรหัสจอไม่ตรง ใส่ Playlist ให้ครบแล้วเติมใหม่</li>
<li><b>ใช้รายการของสื่ออื่น</b> ระบบเติมจากรายการใกล้เคียงให้ เทียบกับ Playlist อีกครั้ง (จอ Spare ที่ใช้รายการสื่ออื่น และจอ CITY ที่ใช้รายการ THE CITY เป็นเรื่องปกติ ไม่ขึ้นใน “ควรตรวจ”)</li>
<li><b>รหัสนี้อยู่ในหลายสื่อ</b> ตั้งชื่อชีตในใบเช็คให้มี 40/80 + Blue/Red, CITY หรือ Spare แล้วเติมใหม่</li>
<li><b>ไม่ใส่ตามหมายเหตุ / จำนวนครั้ง = 0</b> บอกหนังที่ระบบตั้งใจไม่ใส่ ดูว่าตรงกับ Playlist</li>
<li><b>ยังไม่ใส่ เพราะวันเริ่มในไฟล์ยังไม่ถึง</b> หนังที่เริ่มหลังวันที่ทำงาน ระบบจะใส่ให้เองเมื่อถึงวันเริ่ม</li>

<li><b>มีรอบ 16–สิ้นเดือนแล้ว แต่วันที่ทำงานอยู่ในรอบ 1–15</b> วันที่ทำงานน่าจะผิด แก้ในขั้นที่ 3 แล้วเติมใหม่</li>
<li><b>ไม่ได้เติม / เติมไฟล์นี้ไม่ได้</b> ไฟล์เสีย หรือไม่ใช่ใบเช็ครูป (เช่น ใส่ไฟล์ Playlist ผิดขั้น) ไฟล์อื่นยังเติมตามปกติ ใส่ไฟล์ที่ถูกต้องแล้วเติมใหม่</li></ul>
<p class="hn">ถ้าผลไม่ตรง เก็บไฟล์ Playlist ใบเช็คก่อนเติม และไฟล์ zip ของวันนั้นไว้ส่งให้ผู้ดูแลระบบ</p>
<div class="dlgf"><button class="btn primary" data-a="htour">${IC.tour}ดูทัวร์แนะนำหน้านี้</button><button class="btn secondary" data-a="hhow">${IC.how}ระบบทำงานอย่างไร</button></div></div></div>`;
const vWelcome = () =>
  `<div class="ov"><div class="dlg welcome" role="dialog" aria-modal="true" aria-labelledby="wlt"><div class="wv" aria-hidden="true">${['l', 'l', 'l', '', 'l', 'l', 'l', '', 'l', 'a', 'l', ''].map(x => `<i class="${x}"></i>`).join('')}</div><div class="wb"><h2 id="wlt">ยินดีต้อนรับสู่ระบบเติมใบเช็ครูป</h2><p>ใส่ Playlist ของวันนี้กับใบเช็ครูปของเมื่อวาน แล้วระบบจะเติมรายชื่อหนังลงใบเช็คให้ ไฟล์ทั้งหมดประมวลผลในเครื่องนี้ ไม่ถูกอัปโหลด</p><p>ใช้ครั้งแรก แนะนำให้ดูทัวร์สั้น ๆ ก่อน ใช้เวลาประมาณ 1 นาที</p><div class="dlgf"><button class="btn primary" data-a="wtour" data-focus>${IC.tour}เริ่มทัวร์แนะนำ</button><button class="btn secondary" data-a="whow">${IC.how}ระบบทำงานอย่างไร</button><button class="btn ghost" data-a="wskip">ข้าม เริ่มใช้งานเลย</button></div></div></div></div>`;
const vPaste = () =>
  `<div class="ov"><div class="dlg" role="dialog" aria-modal="true" aria-labelledby="pst"><div class="dlgh"><h2 id="pst">วางรายชื่อหนัง</h2><button class="ib" data-a="pastecancel" aria-label="ปิด">${IC.x}</button></div><p class="hint" style="margin-bottom:10px">วางรายชื่อหนัง บรรทัดละ 1 เรื่อง (คัดลอกจาก Excel หรือ Notepad ได้) หนังจะต่อท้ายรายการออนแอร์ของจอนี้</p><textarea id="pastebox" class="pastebox" aria-label="รายชื่อหนัง บรรทัดละ 1 เรื่อง" data-focus></textarea><div class="dlgf"><button class="btn primary" data-a="pasteok">เพิ่มเข้าจอนี้</button><button class="btn ghost" data-a="pastecancel">ยกเลิก</button></div></div></div>`;
function vHow() {
  const row = (g, c0, c1, c2, cls = '', k1 = '', k0 = '') =>
    `<tr class="${cls}"><td class="g">${g ? `<span class="cn">${g}</span>` : ''}</td><td class="c0 ${k0}">${c0}</td><td class="c1 ${k1}">${c1}</td><td class="c2">${c2}</td></tr>`;
  const sheet = `<div><div class="xl" role="img" aria-label="ตัวอย่างใบเช็ครูปหลังเติมวันที่ 3"><table><tbody><tr class="hdrrow"><td class="g"></td><td>วันที่ / หมายเหตุ</td><td>ชื่อหนังโฆษณา</td><td>นับ</td></tr>
 ${row(1, '1-Oct-26', '', '', 'xb')}${row(2, '', '#ประชาสัมพันธ์จอ', '12', '', 'y')}${row('', '', '#ข่าวสารเมือง', '12', '', 'y')}${row(3, '', 'กาแฟ ก รสใหม่', '11')}${row(4, 'ออกวันที่ 3', 'ธนาคาร ข แอปใหม่', '12', '', '', 'rd')}${row(5, '', 'รถยนต์ ค รุ่นใหม่', '10')}
 ${row(6, '2-Oct-26', '', '', 'xb')}${row('', '', 'ประกัน ง แผนครอบครัว', '6')}${row('', '3-Oct-26', '', '', 'xb')}${row('', '', 'มือถือ จ รุ่นพับได้', '3')}${row(7, '4-Oct-26', '', '', 'xb')}</tbody></table></div>
 <div class="xl" style="margin-top:10px" role="img" aria-label="แถบวันที่ของรอบ 16"><table><tbody>${row(8, '16-Oct-26', '', '', 'xb2')}</tbody></table></div><p class="fine">ตัวอย่างใบเช็คหลังเติมวันที่ 3 ชื่อหนังเป็นชื่อสมมติ</p></div>`;
  return `<div class="ov page"><div class="pg" role="dialog" aria-modal="true" aria-labelledby="howt"><div class="pgbar"><b>ระบบทำงานอย่างไร</b><button class="btn secondary" data-a="how" data-focus>${IC.x}ปิด</button></div><div class="pgin">
 <header class="pghero"><h1 id="howt">จาก Playlist 11 ไฟล์ ถึงใบเช็ครูปของทุกคน ในครั้งเดียว</h1><p>ทุกวันผู้ตรวจต้องนำรายชื่อหนังโฆษณาจาก Playlist ไปลงใบเช็ครูปของแต่ละคน ระบบนี้ทำส่วนนั้นให้ตามกฎที่ทีมใช้ทำมือ แล้วให้คนตรวจผลก่อนใช้งานจริง</p></header>
 <section class="flow" aria-label="เส้นทางของไฟล์">
 <div class="fnode"><span class="fic">11</span><b>Playlist ของวันนี้</b><small>ไฟล์ Excel จากอีเมลและ LINE ไม่ต้องเปลี่ยนชื่อไฟล์</small></div><div class="farrow" aria-hidden="true"></div>
 <div class="fnode main"><b>ระบบในเบราว์เซอร์</b><ol><li>อ่านรายชื่อหนังของทุกจอ</li><li>จับคู่จอกับชีตในใบเช็ครูป</li><li>จัดเรียงและใส่ “ออกวันที่” ตามกฎ</li></ol><small>ไฟล์ไม่ออกจากเครื่องนี้</small></div><div class="farrow" aria-hidden="true"></div>
 <div class="fnode"><span class="fic">.xlsm</span><b>ใบเช็ครูปของทุกคน</b><small>มาโคร รูปแบบ และสีเดิมอยู่ครบ ใช้ต่อได้ทันที พร้อมรายการ “ควรตรวจ”</small></div></section>
 <section class="pgsec"><h2>ก่อนและหลังใช้ระบบ</h2><div class="ba">
 <div class="bcol"><h3>เดิม ทำมือ</h3><ol><li>เปิด Playlist ทีละไฟล์</li><li>คัดลอกรายชื่อหนังลง Notepad</li><li>วางลงใบเช็คทีละจอ แล้วจัดเรียงและใส่ “ออกวันที่” เอง</li></ol><p class="tm"><b>1–2 ชม.</b>ต่อวัน รวมทุกคน</p></div>
 <div class="bcol after"><h3>ตอนนี้ ใช้ระบบ</h3><ol><li>ใส่ Playlist ของวันนี้</li><li>ตรวจ / แก้ไข (ไม่บังคับ)</li><li>ใส่ใบเช็คของเมื่อวาน แล้วกดเติม</li><li>อ่าน “ควรตรวจ” แล้วดาวน์โหลด</li></ol><p class="tm"><b>ประมาณ 10 นาที</b>ต่อวัน</p></div></div>
 <p class="fine">เวลาเป็นค่าประมาณ จะจับเวลาจริงช่วงทดลองใช้คู่ขนานกับการทำมือ</p></section>
 <section class="pgsec"><h2>ในใบเช็ครูป ระบบจัดให้แบบนี้</h2><div class="demo">${sheet}<ol class="callouts">
 <li><span class="cn">1</span><span><b>แถบวันที่สีน้ำเงิน</b> วันแรกของรอบ 1–15 ลงหนังที่ออนอยู่ทั้งหมด</span></li>
 <li><span class="cn">2</span><span><b>หนัง # ขึ้นก่อน</b> และไฮไลต์สีเหลือง</span></li>
 <li><span class="cn">3</span><span><b>หนังอื่นเรียง A–Z และ ก–ฮ</b> แบบเดียวกับที่ Excel เรียง</span></li>
 <li><span class="cn">4</span><span><b>“ออกวันที่ N” ตัวแดง</b> ใส่เมื่อถึงวันที่หนังออก ไฟล์เดิม / ต่อสัญญา ไม่ใส่</span></li>
 <li><span class="cn">5</span><span><b>วันที่ถ่ายในช่องหลังชื่อหนัง</b> (ทุกจอ) ย้ายตามชื่อหนังให้เอง ขึ้นเดือนใหม่ล้างให้</span></li>
 <li><span class="cn">6</span><span><b>หนังใหม่ของแต่ละวัน</b> (สถานะ “ขึ้นใหม่”) ลงใต้แถบวันที่ของวันนั้น ไม่ซ้ำกับของเดิม</span></li>
 <li><span class="cn">7</span><span><b>แถบท้ายตารางเป็นวันถัดไป</b> ไว้ลงหนังของพรุ่งนี้</span></li>
 <li><span class="cn">8</span><span><b>วันที่ 16 ทำต่อในรายการเดิม</b> แถบวันที่ 16 เป็นต้นไปเป็นอีกสี (ค่าเริ่มต้นสีฟ้า เลือกสีได้) หนังที่ต่อสัญญาวันที่ 16 ลงใหม่ใต้แถบ 16 แถวเดิมใส่ “ออกวันที่ 16” พอถอดจริงใส่วันถอดทั้งสองแถว</span></li></ol></div></section>
 <section class="pgtwo"><div><h2>ข้อมูลลูกค้าปลอดภัย</h2><ul><li>ไฟล์ Playlist และใบเช็คประมวลผลในเบราว์เซอร์ของเครื่องนี้ ไม่ถูกอัปโหลดไปที่ใด แม้เปิดระบบจากลิงก์</li><li>ใช้แบบออฟไลน์ได้ โดยดาวน์โหลดไฟล์ระบบไปเปิดในเครื่อง</li><li>ทุกครั้งที่เติม ระบบสรุปรายการ “ควรตรวจ” ให้คนตรวจซ้ำ คนเป็นผู้ตัดสินสุดท้ายก่อนใช้งานจริง</li></ul></div>
 <div><h2>สร้างด้วยอะไร</h2><ul><li><b>Claude</b> (AI ของ Anthropic) ช่วยถอดกฎจากใบเช็คที่ทีมทำมือ เขียนระบบ และทดสอบกับไฟล์จริง</li><li><b>SheetJS</b> อ่านไฟล์ Excel และ <b>JSZip</b> แก้ไฟล์ .xlsm โดยมาโครอยู่ครบ</li><li><b>n8n</b> (ไม่บังคับ) รวบรวม Playlist จาก Gmail ลง Google Drive ทุกชั่วโมง พร้อมอีเมลบอกว่าได้ครบหรือยัง</li></ul></div></section>
 <div class="pgcta"><button class="btn primary big" data-a="how">เริ่มใช้งาน</button><button class="btn secondary big" data-a="howtour">${IC.tour}ดูทัวร์แนะนำ</button></div>
 </div></div></div>`;
}

/* ----- ทัวร์แนะนำ: ชี้ทีละจุดในหน้าที่เปิดอยู่ ----- */
const TOURS = {
  dat: [
    {
      t: 'steps',
      h: 'ทำงาน 4 ขั้น',
      p: 'ไล่จากซ้ายไปขวา แถบนี้บอกว่าอยู่ขั้นไหน และแต่ละขั้นทำไปถึงไหนแล้ว กดที่ขั้นเพื่อย้อนกลับได้',
    },
    {
      t: 'drop-pl',
      h: 'ใส่ Playlist ของวันนี้',
      p: 'ลากไฟล์ Playlist ทุกไฟล์มาวางพร้อมกัน หรือคลิกเพื่อเลือกไฟล์ ไม่ต้องเปลี่ยนชื่อไฟล์ ระบบอ่านวันที่จากชื่อไฟล์ให้',
    },
    {
      t: 'wall',
      h: 'เช็กลิสต์ 11 แบบ',
      p: 'แต่ละช่องคือ Playlist 1 แบบ ช่องจะสว่างพร้อมไฟเขียวเมื่อใส่ไฟล์แล้ว ช่องเส้นประคือแบบที่ยังขาด ไฟเหลืองคือวันที่ไม่ตรงกับวันทำงาน',
    },
    {
      t: 'date',
      h: 'วันที่ของ Playlist',
      p: 'ต้องเป็นวันที่จะเติม และบอกว่าเป็นรอบ 1–15 หรือ 16–สิ้นเดือน ถ้าไม่ตรง แก้ได้ในขั้นที่ 3',
    },
    {
      t: 'next1',
      h: 'ไปขั้นต่อไป',
      p: 'ใส่ครบแล้วกด “ไปเติมใบเช็ครูป” ได้เลย ถ้าต้องแก้รายการหนังก่อน กด “ตรวจ / แก้ไขก่อน”',
    },
    {
      t: 'hdr',
      h: 'ตัวช่วยมุมขวาบน',
      p: '“ระบบทำงานอย่างไร” อธิบายภาพรวมสำหรับคนที่ยังไม่รู้จักระบบ “แนะนำการใช้” เปิดทัวร์ของหน้าที่อยู่ “วิธีใช้” มีขั้นตอนและกฎทั้งหมด ปุ่มสุดท้ายสลับโหมดสว่างและมืด',
    },
  ],
  pl: [
    {
      t: 'plseg',
      h: 'แก้รายจอ หรือแก้ทุกจอพร้อมกัน',
      p: 'แก้รายจอ: เลือกจอแล้วแก้หนังในตาราง แก้ทุกจอพร้อมกัน: ค้นชื่อหนังแล้วปลด คืน ลบ หรือเปลี่ยนชื่อในทุกจอที่พบ',
    },
    {
      t: 'chips',
      h: 'กรองตามสื่อ',
      p: 'เลือกสื่อ หรือพิมพ์รหัสจอ เช่น DGT009 ในช่องค้นหา แล้วคลิกการ์ดจอเพื่อเปิดตาราง',
    },
    {
      t: 'plnext',
      h: 'ขั้นนี้ไม่บังคับ',
      p: 'ถ้า Playlist ถูกต้องแล้ว กด “ไปเติมใบเช็ครูป” ได้เลย ถ้าแก้ไว้และจะเก็บไว้ใช้ต่อ กด “บันทึกโปรเจกต์” ในขั้นที่ 1',
    },
  ],
  ed: [
    { t: 'edhead', h: 'จอที่กำลังแก้', p: 'รหัสจอ ชื่อจอ และสื่อของจอนี้ การแก้มีผลกับการเติมครั้งนี้ทันที' },
    {
      t: 'edtbl',
      h: 'ตารางหนังในจอ',
      p: 'สวิตช์เปิด = ออนแอร์ ปิด = ปลด แก้ชื่อหนังหรือสถานะได้ในช่อง ใส่สถานะ “ถอด” ก็ปลดให้ ลูกศรเลื่อนลำดับ ถังขยะลบแถว',
    },
    {
      t: 'edtools',
      h: 'เพิ่มหนัง',
      p: '“+ เพิ่มหนัง” เพิ่มทีละเรื่อง “วางรายชื่อหลายบรรทัด” วางรายชื่อที่คัดลอกจาก Excel หรือ Notepad ได้ทีเดียวหลายเรื่อง',
    },
    { t: 'edback', h: 'กลับไปรายการจอ', p: 'แก้เสร็จแล้วกลับไปเลือกจออื่น จอที่แก้แล้วจะมีป้าย “แก้แล้ว”' },
  ],
  fl: [
    {
      t: 'drop-ck',
      h: 'ใส่ใบเช็ครูปของเมื่อวาน',
      p: 'ลากใบเช็ครูปของวันก่อนหน้ามาวาง ใส่ของทุกคนพร้อมกัน หรือเฉพาะของคุณก็ได้',
    },
    {
      t: 'ready',
      h: 'ดูว่าแต่ละไฟล์พร้อมไหม',
      p: 'ระบบดูว่าใบเช็คแต่ละไฟล์ใช้ Playlist แบบไหน จุดเขียวคือมีแล้ว สีเหลืองคือยังขาดหรือหาจอไม่พบ ถ้าขาดยังกดเติมได้ แต่จอนั้นจะขึ้นใน “ควรตรวจ”',
    },
    {
      t: 'settings',
      h: 'วันที่และรอบ',
      p: 'ปกติเว้นว่างไว้ ระบบใช้วันที่จากชื่อไฟล์ Playlist ถ้าพิมพ์เอง ระบบจะบอกวันที่ที่อ่านได้ใต้ช่อง ดูว่ารอบถูกต้อง เปิด “จัดรูปแบบตามใบเช็คจริง” ไว้เสมอ และเลือกสีแถบวันที่ได้ (ปกติใช้ “ตามชีต”) ตัวเลือกเพิ่มเติมทุกตัวมีคำอธิบายใต้ตัวเลือก',
    },
    { t: 'fill', h: 'กดเติม', p: 'ระบบเติมทุกไฟล์ในครั้งเดียว ใช้เวลาไม่กี่วินาที แล้วพาไปหน้าผลการเติม' },
  ],
  res: [
    { t: 'sum', h: 'สรุปผล', p: 'จำนวนไฟล์ จำนวนจอที่เติมแล้ว และจำนวนรายการที่ควรตรวจ' },
    {
      t: 'rlist',
      h: 'ดูรายการ “ควรตรวจ”',
      p: 'กดลูกศรท้ายแต่ละไฟล์เพื่อดูรายละเอียด แล้วตรวจในใบเช็คให้เรียบร้อยก่อนใช้งาน',
    },
    {
      t: 'dlz',
      h: 'ดาวน์โหลด',
      p: 'กดเพื่อบันทึกไฟล์ zip ที่มีใบเช็คทุกไฟล์ ไฟล์ยังไม่ถูกบันทึกจนกว่าจะกดปุ่มนี้ ใบเช็คที่ได้คือไฟล์ตั้งต้นของพรุ่งนี้',
    },
  ],
};
const tq = t => document.querySelector(`[data-tour="${t}"]`);
function tourStart(key, auto) {
  const L = (TOURS[key] || []).filter(s => tq(s.t));
  if (auto) store.set('ps_tour_' + key, '1');
  if (!L.length) {
    if (!auto) say('หน้านี้ยังไม่มีทัวร์แนะนำ');
    return;
  }
  store.set('ps_tour_' + key, '1');
  TOUR = { key, L, i: 0 };
  TDRAWN = '';
}
function tourDraw() {
  const box = $('#tour');
  if (!TOUR) {
    if (box.innerHTML) box.innerHTML = '';
    TDRAWN = '';
    return;
  }
  const s = TOUR.L[TOUR.i],
    el = tq(s.t);
  if (!el) {
    TOUR = null;
    box.innerHTML = '';
    return;
  }
  const n = TOUR.L.length,
    i = TOUR.i;
  box.innerHTML = `<div class="tcatch"></div><div class="thole"></div><div class="tbub" role="dialog" aria-modal="true" aria-labelledby="tbt" aria-describedby="tbp"><p class="tc">${i + 1} จาก ${n}</p><h3 id="tbt">${s.h}</h3><p id="tbp">${s.p}</p><div class="tdots" aria-hidden="true">${TOUR.L.map((_, j) => `<i class="${j === i ? 'on' : ''}"></i>`).join('')}</div><div class="tbtn"><button class="btn ghost" data-a="tend">ปิดทัวร์</button><span class="sp"></span>${i ? '<button class="btn secondary" data-a="tprev">ย้อนกลับ</button>' : ''}<button class="btn primary" data-a="tnext">${i === n - 1 ? 'เข้าใจแล้ว' : 'ถัดไป'}</button></div></div>`;
  const k = TOUR.key + ':' + i;
  if (TDRAWN !== k) {
    TDRAWN = k;
    const r = el.getBoundingClientRect();
    if (r.top < 8 || r.bottom > innerHeight - 200)
      el.scrollIntoView({ block: r.height > innerHeight * 0.55 ? 'start' : 'center' });
  }
  tourPlace();
  const b = box.querySelector('[data-a=tnext]');
  if (b) b.focus({ preventScroll: true });
}
function tourPlace() {
  if (!TOUR) return;
  const el = tq(TOUR.L[TOUR.i].t),
    h = $('#tour .thole'),
    b = $('#tour .tbub');
  if (!el || !h || !b) return;
  const r = el.getBoundingClientRect(),
    p = 6,
    vw = document.documentElement.clientWidth,
    vh = innerHeight;
  const T = Math.max(4, r.top - p),
    L = Math.max(4, r.left - p),
    Rr = Math.min(vw - 4, r.right + p),
    B = Math.min(vh - 4, r.bottom + p);
  Object.assign(h.style, {
    top: T + 'px',
    left: L + 'px',
    width: Math.max(0, Rr - L) + 'px',
    height: Math.max(0, B - T) + 'px',
  });
  const bw = b.offsetWidth,
    bh = b.offsetHeight;
  let y = B + 12;
  if (y + bh > vh - 8) {
    y = T - 12 - bh;
    if (y < 8) y = Math.max(8, vh - bh - 12);
  }
  b.style.top = y + 'px';
  b.style.left = Math.min(Math.max(8, L), vw - bw - 8) + 'px';
}
addEventListener('resize', tourPlace);
addEventListener('scroll', tourPlace, { passive: true });

/* ----- วาดหน้าจอ (เก็บโฟกัส ตำแหน่งเคอร์เซอร์ และตำแหน่งเลื่อนของตารางไว้) ----- */
const focusKey = el =>
  !el || el === document.body
    ? null
    : el.id
      ? { id: el.id }
      : el.dataset && el.dataset.f
        ? { f: el.dataset.f, i: el.dataset.i }
        : el.dataset && el.dataset.a
          ? { a: el.dataset.a, k: el.dataset.k }
          : null;
const findKey = k => {
  if (!k) return null;
  try {
    if (k.id) return document.getElementById(k.id);
    if (k.f) return document.querySelector(`[data-f="${CSS.escape(k.f)}"][data-i="${CSS.escape(k.i)}"]`);
    return document.querySelector(
      `[data-a="${CSS.escape(k.a)}"]${k.k != null ? `[data-k="${CSS.escape(k.k)}"]` : ''}`,
    );
  } catch (_) {
    return null;
  }
};
function softView() {
  if (PTR) DEFER = true;
  else view();
}
function view() {
  const ae = document.activeElement,
    fk = focusKey(ae);
  let sel = null;
  try {
    if (ae && typeof ae.selectionStart === 'number') sel = [ae.selectionStart, ae.selectionEnd];
  } catch (_) {}
  const kv = {};
  for (const id of ['rn', 'am', 'an', 'at', 'as', 'o_base']) if ($('#' + id)) kv[id] = $('#' + id).value;
  const tb = $('.tbl'),
    tbs = tb ? [tb.scrollTop, tb.scrollLeft] : null;
  $('#hnav').innerHTML = vNav();
  $('#steps').innerHTML = vSteps();
  $('#ver').textContent = VERTXT;
  $('#app').innerHTML = { dat: vData, pl: vPl, tl: vTl, fl: vFl, res: vRes }[tab]();
  if (tab === 'fl') {
    $('#o_media').value = OP.media;
    $('#o_sort').value = OP.sort;
    $('#o_new').checked = OP.newOnly;
    $('#o_dd').checked = OP.dedupe;
    $('#o_off').checked = OP.off;
    $('#o_hand').checked = OP.hand;
    $('#o_r2').value = OP.r2;
    $('#o_base').value = BASEERR || OP.base;
  }
  for (const id in kv) {
    const el = $('#' + id);
    if (el && (id !== 'o_base' || (fk && fk.id === 'o_base'))) el.value = kv[id];
  }
  if (tbs) {
    const t2 = $('.tbl');
    if (t2) {
      t2.scrollTop = tbs[0];
      t2.scrollLeft = tbs[1];
    }
  }
  $('#busy').innerHTML = vBusy();
  const ok = WELCOME ? 'w' : PASTE ? 'v' : HOW ? 'h' : HELP ? 'p' : '',
    o = $('#ovl'),
    was = OVK;
  if (OVK !== ok) {
    OVK = ok;
    o.innerHTML =
      ok === 'w' ? vWelcome() : ok === 'v' ? vPaste() : ok === 'h' ? vHow() : ok === 'p' ? vHelp() : '';
    if (ok) {
      if (!was) RETF = RETF || fk;
      const f = o.querySelector('[data-focus]') || o.querySelector('button');
      if (f) f.focus({ preventScroll: true });
    }
  }
  document.body.classList.toggle('lock', !!ok || !!BUSY);
  if (note) {
    say(note);
    note = '';
  }
  scanLater();
  if (pendTour) {
    const k = pendTour;
    pendTour = null;
    if (!TOUR) RETF = RETF || fk;
    tourStart(k, false);
  } else if (!TOUR && !ok && !BUSY && store.get('ps_tours') === 'on' && !store.get('ps_tour_' + tourKey()))
    tourStart(tourKey(), true);
  $('.wrap').inert = !!(ok || TOUR || BUSY);
  tourDraw();
  const tw = TOURWAS;
  TOURWAS = !!TOUR;
  if (!ok && !TOUR && !BUSY) {
    if (FOCUSSEL) {
      const e = document.querySelector(FOCUSSEL);
      FOCUSSEL = null;
      if (e) e.focus({ preventScroll: true });
    } else if (FOCUSNEW != null) {
      const e = $(`[data-f="ad"][data-i="${FOCUSNEW}"]`);
      FOCUSNEW = null;
      if (e) {
        e.scrollIntoView({ block: 'center' });
        e.focus();
      }
    } else if (was || tw) {
      const e = findKey(RETF);
      RETF = null;
      if (e) e.focus({ preventScroll: true });
    } else if (fk) {
      const e = findKey(fk);
      if (e && e !== document.activeElement) {
        e.focus({ preventScroll: true });
        if (sel)
          try {
            e.setSelectionRange(sel[0], sel[1]);
          } catch (_) {}
      }
    }
  }
}
const closeOv = () => {
  if (WELCOME) {
    store.set('ps_welcomed', '1');
    WELCOME = false;
  }
  HELP = HOW = PASTE = false;
};
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') {
    if (TOUR) {
      TOUR = null;
      view();
    } else if (WELCOME || HELP || HOW || PASTE) {
      closeOv();
      view();
    }
    return;
  }
  if (
    TOUR &&
    (e.key === 'ArrowRight' || e.key === 'ArrowLeft') &&
    !/^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName)
  ) {
    e.key === 'ArrowRight' ? ACT.tnext() : ACT.tprev();
    view();
    return;
  }
  if (
    (e.key === 'Enter' || e.key === ' ') &&
    e.target.matches &&
    e.target.matches('div[role=button][data-a=pick]')
  ) {
    e.preventDefault();
    if (!BUSY && !FILLING) openPicker(e.target.dataset.k);
    return;
  }
  if (e.key === 'Tab') {
    const box = TOUR ? $('#tour .tbub') : WELCOME || HELP || HOW || PASTE ? $('#ovl [role=dialog]') : null;
    if (!box) return;
    const F = [...box.querySelectorAll('button:not([disabled]),textarea,input,[tabindex="0"]')];
    if (!F.length) return;
    const i = F.indexOf(document.activeElement);
    if (e.shiftKey && i <= 0) {
      e.preventDefault();
      F[F.length - 1].focus();
    } else if (!e.shiftKey && (i === -1 || i === F.length - 1)) {
      e.preventDefault();
      F[0].focus();
    }
  }
});
document.addEventListener(
  'mousedown',
  e => {
    MDOWN = e.target;
  },
  true,
);
document.addEventListener(
  'pointerdown',
  () => {
    PTR = true;
  },
  true,
);
document.addEventListener(
  'pointerup',
  () => {
    PTR = false;
    if (DEFER) {
      DEFER = false;
      setTimeout(view, 0);
    }
  },
  true,
);
document.addEventListener(
  'pointercancel',
  () => {
    PTR = false;
    if (DEFER) {
      DEFER = false;
      setTimeout(view, 0);
    }
  },
  true,
);
document.addEventListener('click', e => {
  if (e.target.classList && e.target.classList.contains('ov') && MDOWN === e.target && !PASTE) {
    closeOv();
    view();
  }
});
document.addEventListener(
  'toggle',
  e => {
    if (e.target.matches && e.target.matches('details.files')) FLOPEN = e.target.open;
  },
  true,
);
document.addEventListener('dragenter', e => {
  const z = e.target.closest && e.target.closest('[data-dz]');
  if (z) z.classList.add('over');
});
document.addEventListener('dragleave', e => {
  const z = e.target.closest && e.target.closest('[data-dz]');
  if (z && !z.contains(e.relatedTarget)) z.classList.remove('over');
});
document.addEventListener(
  'drop',
  () => document.querySelectorAll('.drop.over').forEach(z => z.classList.remove('over')),
  true,
);

/* ----- รับไฟล์: แยกไฟล์ให้ถูกที่ ป้องกันการกดซ้ำระหว่างทำงาน และตรวจไฟล์โปรเจกต์ก่อนเปิด ----- */
const okScreen = s =>
  s &&
  typeof s === 'object' &&
  typeof s.m === 'string' &&
  typeof s.c === 'string' &&
  Array.isArray(s.rows) &&
  s.rows.every(r => r && typeof r === 'object' && typeof r.ad === 'string');
async function loadProject(f) {
  let o;
  try {
    o = JSON.parse(new TextDecoder().decode(await f.arrayBuffer()));
  } catch (e) {
    note = 'เปิดโปรเจกต์ไม่ได้: ไฟล์นี้ไม่ใช่ไฟล์โปรเจกต์ (.json) ของระบบ';
    return view();
  }
  const L = Array.isArray(o) ? o : o && Array.isArray(o.DB) ? o.DB : null;
  if (!L || !L.length || !L.every(okScreen)) {
    note = 'ไฟล์นี้ไม่ใช่ไฟล์โปรเจกต์ของระบบ ใช้ไฟล์ที่ได้จากปุ่ม “บันทึกโปรเจกต์”';
    return view();
  }
  for (const s of L)
    for (const r of s.rows) {
      r.off = !!r.off;
      for (const k of ['ty', 'st', 's', 'e', 'note'])
        if (r[k] != null && typeof r[k] !== 'string') r[k] = String(r[k]);
      for (const k of ['len', 'cnt']) if (r[k] != null && typeof r[k] !== 'number') r[k] = +r[k] || 0;
    }
  if (Array.isArray(o)) {
    DB = L;
    PLD = null;
    PLM = {};
  } else {
    DB = L;
    PLD = plDate() || (o.PLD ? new Date(o.PLD) : null);
  } /* เหมือน addFiles('pj') ของเดิม แต่ใช้ข้อมูลที่ตรวจแล้ว */
  dirty = 0;
  BASEERR = '';
  OP.base = '';
  cs = null;
  LIT = new Set(L.map(s => kindOfS(s)));
  note = 'เปิดโปรเจกต์แล้ว';
  view();
}
async function addAny(kind, files) {
  files = [...files];
  if (!files.length) return;
  if (BUSY || FILLING) {
    say('กำลังทำงานอยู่ รอให้เสร็จก่อน');
    return;
  }
  if (TOUR) TOUR = null;
  const js = files.filter(f => /\.json$/i.test(f.name)),
    rest = files.filter(f => !/\.json$/i.test(f.name));
  if (kind === 'pj' || (js.length && !rest.length)) {
    if (js[0] || files[0]) await loadProject(js[0] || files[0]);
    return;
  }
  const ck = kind === 'ck' ? rest : rest.filter(f => /\.xlsm$/i.test(f.name)),
    pl = kind === 'ck' ? [] : rest.filter(f => !/\.xlsm$/i.test(f.name));
  if (pl.length) {
    BUSY = { t: 'กำลังอ่าน Playlist', i: 0, n: pl.length, f: '' };
    view();
    try {
      await addFiles('pl', pl);
    } finally {
      BUSY = null;
      BASEERR = '';
      $('.wrap').inert = false;
    }
    /* ไฟล์ที่อ่านเป็น Playlist ไม่ได้ แต่เป็นใบเช็ครูป (.xlsx) → ย้ายไปขั้นที่ 3 ให้ */
    const mv2 = [];
    for (const f of pl) {
      if (!PLX.some(x => x.f === f.name)) continue;
      try {
        const nd = sheetNeeds(await f.arrayBuffer(), f.name);
        if (!nd.bad && (nd.S.length || nd.named.length)) mv2.push(f);
      } catch (_) {}
    }
    if (mv2.length) {
      const nm = new Set(mv2.map(f => f.name));
      PLX = PLX.filter(x => !nm.has(x.f));
      await addChecks(mv2, `${mv2.length} ไฟล์เป็นใบเช็ครูป ใส่ไว้ในขั้นที่ 3 ให้แล้ว`);
    }
    FOCUSSEL = '[data-dz="pl"]';
    view();
  }
  if (ck.length)
    await addChecks(
      ck,
      kind !== 'ck' ? `ไฟล์ .xlsm ${ck.length} ไฟล์เป็นใบเช็ครูป ใส่ไว้ในขั้นที่ 3 ให้แล้ว` : '',
    );
}
async function addChecks(ck, msg) {
  const M = new Map();
  for (const f of ck) M.set(f.name, f);
  const inDrop = ck.length - M.size,
    uniq = [...M.values()],
    names = new Set(M.keys()),
    dup = C.filter(x => names.has(x.name)).length;
  C = C.filter(x => !names.has(x.name));
  await addFiles('ck', uniq);
  const extra = [
    dup ? `แทนที่ไฟล์ชื่อซ้ำ ${dup} ไฟล์` : '',
    inDrop ? `มีไฟล์ชื่อซ้ำกัน ${inDrop} ไฟล์ ใช้ไฟล์ที่เลือกท้ายสุด` : '',
  ]
    .filter(Boolean)
    .join(' ');
  if (msg || extra) {
    note = (msg || `เพิ่มใบเช็ครูป ${uniq.length} ไฟล์`) + (extra ? ` (${extra})` : '');
    view();
  }
}

/* ----- ปุ่มต่าง ๆ ----- */
const mv = (a, i, j) => {
  if (j < 0 || j >= a.length) return;
  [a[i], a[j]] = [a[j], a[i]];
};
const touch = s => {
  s.ed = 1;
  dirty = 1;
};
const curS = () => (cs != null && DB[cs] ? DB[cs] : null);
const ACT = {
  help: () => {
    HELP = !HELP;
    HOW = false;
  },
  how: () => {
    HOW = !HOW;
    HELP = false;
  },
  hhow: () => {
    HELP = false;
    HOW = true;
  },
  tour: () => {
    HELP = HOW = false;
    pendTour = tourKey();
  },
  htour: () => {
    HELP = false;
    pendTour = tourKey();
  },
  howtour: () => {
    HOW = false;
    store.set('ps_tours', 'on');
    pendTour = tourKey();
  },
  wtour: () => {
    store.set('ps_welcomed', '1');
    store.set('ps_tours', 'on');
    WELCOME = false;
    pendTour = tourKey();
  },
  whow: () => {
    store.set('ps_welcomed', '1');
    WELCOME = false;
    HOW = true;
  },
  wskip: () => {
    store.set('ps_welcomed', '1');
    store.set('ps_tours', 'off');
    WELCOME = false;
  },
  tnext: () => {
    if (!TOUR) return;
    if (TOUR.i < TOUR.L.length - 1) TOUR.i++;
    else TOUR = null;
  },
  tprev: () => {
    if (TOUR && TOUR.i > 0) TOUR.i--;
  },
  tend: () => {
    TOUR = null;
  },
  theme: () => {
    const t = THEME === 'dark' ? 'light' : 'dark';
    store.set('ps_theme', t);
    setTheme(t);
  },
  col: k => {
    const [r, c] = k.split('|');
    OP['col' + r] = c || '';
    store.set('ps_col' + r, c || '');
  },
  colc: k => {
    const p = $('#colpick');
    if (!p) return;
    p.dataset.k = k;
    p.value = swHex(OP['col' + k]) || '#4472c4';
    p.click();
  },
  optreset: () => {
    Object.assign(OP, OPDEF);
    note = 'คืนค่าแนะนำแล้ว';
    FOCUSSEL = '#o_hand';
  },
  clrx: () => {
    PLX = [];
    FOCUSSEL = '[data-dz="pl"]';
  },
  clri: () => {
    for (const x of PLI) x.ack = 1;
    note = 'ซ่อนรายการด่านตรวจแล้ว (ยังบันทึกไว้ในรายงาน)';
    FOCUSSEL = '[data-dz="pl"]';
  },
  pick: k => {
    if (!BUSY && !FILLING) openPicker(k);
  },
  adv: () => {
    adv = !adv;
  },
  rx: k => {
    RO.has(+k) ? RO.delete(+k) : RO.add(+k);
  },
  tab: k => {
    if (BUSY || FILLING) return;
    if (k === 'res' && !last) return;
    tab = k;
    cs = null;
    window.scrollTo(0, 0);
  },
  media: k => {
    fm = k;
  },
  open: k => {
    if (DB[+k]) cs = +k;
  },
  back: () => {
    cs = null;
  },
  reset: () => {
    if (confirm('ล้าง Playlist ใบเช็ครูป และผลการเติมที่ใส่ไว้ทั้งหมด เพื่อเริ่มวันใหม่?')) {
      DB = [];
      PLD = null;
      PLM = {};
      PLX = [];
      PLI = [];
      OP.base = '';
      OP.media = 'ทั้งหมด';
      BASEERR = '';
      C = [];
      last = null;
      RO.clear();
      fm = 'ทั้งหมด';
      q = '';
      TQ = '';
      cs = null;
      dirty = 0;
      LIT = new Set();
      note = 'ล้างแล้ว เริ่มใส่ Playlist ของวันใหม่ได้เลย';
    }
  },
  add: () => {
    const s = curS();
    if (!s) return;
    const i = s.rows.findIndex(r => r.off),
      at = i < 0 ? s.rows.length : i;
    s.rows.splice(at, 0, {
      ad: '',
      ty: '',
      len: 15,
      cnt: 1,
      st: 'ขึ้นใหม่',
      s: '',
      e: '',
      note: '',
      off: false,
    });
    touch(s);
    FOCUSNEW = at;
  },
  paste: () => {
    if (curS()) PASTE = true;
  },
  pastecancel: () => {
    PASTE = false;
  },
  pasteok: () => {
    const s = curS(),
      t = ($('#pastebox') || {}).value || '';
    PASTE = false;
    if (!s) return;
    const L = t
      .split(/\r?\n/)
      .map(x => x.trim())
      .filter(Boolean);
    if (!L.length) return;
    const i = s.rows.findIndex(r => r.off);
    s.rows.splice(
      i < 0 ? s.rows.length : i,
      0,
      ...L.map(ad => ({ ad, ty: '', len: 15, cnt: 1, st: '', s: '', e: '', note: '', off: false })),
    );
    touch(s);
    note = `เพิ่มแล้ว ${L.length} เรื่อง`;
  },
  up: k => {
    const s = curS();
    if (s) {
      mv(s.rows, +k, +k - 1);
      touch(s);
      FOCUSSEL = `[data-a="up"][data-k="${Math.max(0, +k - 1)}"]`;
    }
  },
  dn: k => {
    const s = curS();
    if (s) {
      mv(s.rows, +k, +k + 1);
      touch(s);
      FOCUSSEL = `[data-a="dn"][data-k="${Math.min(s.rows.length - 1, +k + 1)}"]`;
    }
  },
  del: k => {
    const s = curS();
    if (s) {
      const r = s.rows.splice(+k, 1)[0];
      touch(s);
      FOCUSSEL = '[data-a="add"]';
      note = `ลบ “${r ? r.ad : ''}” แล้ว`;
    }
  },
  rm: k => {
    if (!FILLING) C.splice(+k, 1);
  },
  gOff: () =>
    g(r => {
      r.off = true;
    }),
  gOn: () =>
    g(r => {
      r.off = false;
    }),
  gDel: () => {
    const H = tlHits();
    if (H.n && confirm(`ลบหนังที่มีคำว่า “${TQ}” ออก ${H.n} แถว ใน ${H.sc} จอ? ลบแล้วย้อนกลับไม่ได้`))
      g(null);
  },
  gRn: () => {
    const v = $('#rn').value.trim();
    if (v) {
      g(r => {
        r.ad = v;
      });
      TQ = v;
    }
  },
  gAdd: () => {
    const m = $('#am').value,
      n = $('#an').value.trim();
    if (!n) {
      note = 'ใส่ชื่อหนังก่อน';
      return;
    }
    let c = 0;
    for (const s of DB)
      if (s.m === m) {
        const i = s.rows.findIndex(r => r.off);
        s.rows.splice(i < 0 ? s.rows.length : i, 0, {
          ad: n,
          ty: $('#at').value,
          len: 15,
          cnt: 1,
          st: $('#as').value,
          s: '',
          e: '',
          note: '',
          off: false,
        });
        touch(s);
        c++;
      }
    note = `เพิ่มแล้ว ${c} จอ`;
  },
  savepj: () => {
    saveBlob('playlist_project.json', JSON.stringify({ v: 2, DB, PLD: PLD ? +PLD : null, PLM }));
    dirty = 0;
    note = 'บันทึกโปรเจกต์แล้ว';
  },
  ckdir: async () => {
    if (!FSA) return;
    let d;
    try {
      d = await showDirectoryPicker({ id: 'ck', mode: 'readwrite' });
    } catch (e) {
      if (e && e.name === 'AbortError') return;
      note = 'เปิดโฟลเดอร์ไม่ได้';
      return;
    }
    const fs = [];
    for await (const [n, x] of d.entries())
      if (x.kind === 'file' && /\.xls[xm]$/i.test(n) && !n.startsWith('~$')) fs.push(await x.getFile());
    if (!fs.length) {
      note = `ไม่พบไฟล์ใบเช็ค (.xlsm / .xlsx) ในโฟลเดอร์ “${d.name}”`;
      return;
    }
    CKDIR = d;
    HDB.set('ck', d).catch(() => {});
    fs.sort((a, b) => a.name.localeCompare(b.name, 'th'));
    await addChecks(fs, `เปิดใบเช็ค ${fs.length} ไฟล์จากโฟลเดอร์ “${d.name}”`);
  },
  cmpscan: async k => {
    if (!last || !FSA) return;
    let d = k === 'again' && last.cmp ? last.cmp.d : null;
    try {
      if (!d && k !== 'pick') {
        const x = await HDB.get('cmp').catch(() => null);
        if (x) d = x;
      }
      if (
        d &&
        (await d.queryPermission({ mode: 'read' })) !== 'granted' &&
        (await d.requestPermission({ mode: 'read' })) !== 'granted'
      )
        d = null;
      if (!d) d = await showDirectoryPicker({ id: 'cmp', mode: 'read' });
    } catch (e) {
      if (e && e.name === 'AbortError') return;
      note = 'เลือกโฟลเดอร์ไม่ได้';
      return;
    }
    HDB.set('cmp', d).catch(() => {});
    note = 'กำลังตรวจโฟลเดอร์เทียบจอ …';
    view();
    const R = await cmpCheck(d, last.W, last.res);
    last.cmp = { d, R };
    note = R.err
      ? `ไม่พบโฟลเดอร์รอบในโฟลเดอร์ “${d.name}”`
      : R.n
        ? `ตรวจแล้ว ต้องแก้ ${R.n.toLocaleString()} เรื่อง`
        : 'ตรวจแล้ว ถูกต้องทั้งหมด';
  },
  cmpcopy: async () => {
    const R = last && last.cmp && last.cmp.R;
    if (!R || !R.txt) return;
    try {
      await navigator.clipboard.writeText(R.txt);
      note = 'คัดลอกผลตรวจแล้ว วางในแชตหรืออีเมลได้เลย';
    } catch (_) {
      note = 'คัดลอกไม่ได้ ลองอีกครั้ง';
    }
  },
  picscan: async k => {
    if (!last || !FSA) return;
    let d = k === 'again' && last.pic ? last.pic.d : null;
    try {
      if (!d && k !== 'pick') {
        const x = await HDB.get('pic').catch(() => null);
        if (x) d = x;
      }
      if (
        d &&
        (await d.queryPermission({ mode: 'readwrite' })) !== 'granted' &&
        (await d.requestPermission({ mode: 'readwrite' })) !== 'granted'
      )
        d = null;
      if (!d) d = await showDirectoryPicker({ id: 'pic', mode: 'readwrite' });
    } catch (e) {
      if (e && e.name === 'AbortError') return;
      note = 'เลือกโฟลเดอร์ไม่ได้';
      return;
    }
    HDB.set('pic', d).catch(() => {});
    note = 'กำลังตรวจโฟลเดอร์แยกรูป …';
    view();
    const P = await picPlan(d, last.W, last.res);
    last.pic = { d, P, made: k === 'again' && last.pic ? last.pic.made : null };
    note = P.plan.length
      ? P.n
        ? `ตรวจแล้ว จะสร้าง ${P.n.toLocaleString()} โฟลเดอร์`
        : 'โฟลเดอร์แยกรูปครบแล้ว'
      : `ไม่พบโฟลเดอร์จอในโฟลเดอร์ “${d.name}” — เลือกโฟลเดอร์แยกรูปที่มีโฟลเดอร์ Pic @ …`;
  },
  picmake: async () => {
    if (!last || !last.pic || !last.pic.P) return;
    note = 'กำลังสร้างโฟลเดอร์ …';
    view();
    const r = await picMake(last.pic.P);
    last.pic.made = r;
    last.pic.P = await picPlan(last.pic.d, last.W, last.res);
    note = r.bad.length
      ? `สร้างแล้ว ${r.n} โฟลเดอร์ สร้างไม่ได้ ${r.bad.length} โฟลเดอร์`
      : `สร้างแล้ว ${r.n} โฟลเดอร์`;
  },
  savedir: async k => {
    if (!last || !FSA) return;
    let d = k === 'pick' ? null : CKDIR;
    try {
      if (!d && k !== 'pick') {
        const x = await HDB.get('ck').catch(() => null);
        if (x) d = x;
      }
      if (
        d &&
        (await d.queryPermission({ mode: 'readwrite' })) !== 'granted' &&
        (await d.requestPermission({ mode: 'readwrite' })) !== 'granted'
      )
        d = null;
      if (!d) d = await showDirectoryPicker({ id: 'ck', mode: 'readwrite' });
    } catch (e) {
      if (e && e.name === 'AbortError') return;
      note = 'เลือกโฟลเดอร์ไม่ได้';
      return;
    }
    CKDIR = d;
    HDB.set('ck', d).catch(() => {});
    const day = isoD(new Date(last.W));
    const bk = await (
      await d.getDirectoryHandle('สำรองก่อนเติม', { create: true })
    ).getDirectoryHandle(day, { create: true });
    const has = async (dir, n) => {
      try {
        return await (await dir.getFileHandle(n)).getFile();
      } catch (_) {
        return null;
      }
    };
    let ok = 0,
      neu = 0;
    const bad = [];
    for (const r of last.res) {
      if (!r.blob) continue;
      try {
        const old = await has(d, r.name);
        if (old) {
          if (!(await has(bk, r.name))) {
            const w = await (await bk.getFileHandle(r.name, { create: true })).createWritable();
            await w.write(await old.arrayBuffer());
            await w.close();
          }
        } else neu++;
        const w = await (await d.getFileHandle(r.name, { create: true })).createWritable();
        await w.write(r.blob);
        await w.close();
        ok++;
      } catch (e) {
        bad.push(r.name);
      }
    }
    try {
      const w = await (await bk.getFileHandle('รายงาน.txt', { create: true })).createWritable();
      await w.write(last.txt || '');
      await w.close();
    } catch (_) {}
    last.sv = { dir: d.name, ok, neu, bad, day };
    if (!bad.length) last.saved = true;
    note = bad.length
      ? `บันทึกไม่ได้ ${bad.length} ไฟล์ — ปิดไฟล์ใน Excel ก่อน แล้วกดบันทึกอีกครั้ง`
      : `บันทึกทับ ${ok} ไฟล์ในโฟลเดอร์ “${d.name}” แล้ว`;
  },
  next: () => {
    if (!last) return;
    if (
      !last.saved &&
      !confirm(
        'ผลการเติมนี้ยังไม่ได้บันทึกหรือดาวน์โหลด ใช้เป็นใบเช็คตั้งต้นของวันถัดไปเลยไหม? (ยังดาวน์โหลดผลนี้ได้จนกว่าจะเติมครั้งถัดไป)',
      )
    )
      return;
    const old = new Map(C.map(f => [f.name, f]));
    C = last.res
      .filter(r => r.blob)
      .map(r => {
        const o = old.get(r.name) || {};
        const x = {
          ...o,
          name: r.name,
          buf: r.blob.buffer.slice(r.blob.byteOffset, r.blob.byteOffset + r.blob.byteLength),
        };
        delete x.nd;
        return x;
      });
    DB = [];
    PLD = null;
    PLM = {};
    PLX = [];
    PLI = [];
    OP.base = '';
    BASEERR = '';
    LIT = new Set();
    RO.clear();
    cs = null;
    fm = 'ทั้งหมด';
    q = '';
    TQ = '';
    dirty = 0;
    tab = 'dat';
    window.scrollTo(0, 0);
    note = `ใบเช็คที่เติมแล้ว ${C.length} ไฟล์พร้อมเป็นไฟล์ตั้งต้น ใส่ Playlist ของวันถัดไปได้เลย`;
  },
  dlz: () => {
    if (!last) return;
    saveBlob('ใบเช็ครูป_เติมแล้ว.zip', last.zip);
    last.saved = true;
  },
  dlf: k => {
    const r = last && last.res[+k];
    if (r && r.blob) {
      saveBlob(r.name, r.blob);
      (last.dl = last.dl || new Set()).add(+k);
      if (last.res.every((x, i) => !x.blob || last.dl.has(i))) last.saved = true;
    }
  },
  expx: () => {
    const A = [
      [
        'สื่อ',
        'รหัสจอ',
        'ชื่อจอ',
        'รายการ',
        'ลำดับ',
        'หนังโฆษณา',
        'Type',
        'สถานะ spot',
        'ความยาว (วิ)',
        'จำนวนครั้ง',
        'วันเริ่ม',
        'วันสิ้นสุด',
        'ข้อมูลการออน / หมายเหตุ',
      ],
    ];
    for (const s of DB) {
      let a = 0,
        b = 0;
      for (const r of s.rows)
        A.push([
          s.m,
          s.c,
          s.n,
          r.off ? 'ปลด' : 'ออนแอร์',
          r.off ? ++b : ++a,
          r.ad,
          r.ty,
          r.st,
          r.len,
          r.cnt,
          r.s,
          r.e,
          r.note,
        ]);
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(A), 'Playlist');
    saveBlob('Playlist_ส่งออก.xlsx', XLSX.write(wb, { type: 'array', bookType: 'xlsx', compression: true }));
  },
  /* เติมใบเช็ค: เหมือนเวอร์ชัน 1.0 แต่ล็อกหน้าจอระหว่างเติม และไฟล์ที่เปิดไม่ได้ไม่ทำให้ไฟล์อื่นล้ม */
  fill: async () => {
    if (FILLING || BUSY) return;
    if (!DB.length) {
      note = 'ยังไม่ได้ใส่ Playlist ของวันนี้ — ใส่ไฟล์ที่ขั้นที่ 1 ก่อน';
      return;
    }
    if (BASEERR) {
      note = 'แก้วันที่ทำงานให้ถูกต้องก่อน';
      return;
    }
    const o = { ...OP, W: workDate() };
    if (o.hand && !o.W) {
      note = 'ไม่พบวันที่ Playlist — ใส่วันที่ทำงานเอง เช่น 2-Oct-26';
      return;
    }
    const list = [...C];
    if (!list.length) return;
    {
      const ps = PLI.filter(x => x.lv === 'stop' && !x.ack);
      if (
        ps.length &&
        !confirm(
          `ด่านตรวจ Playlist พบ ${ps.length} เรื่องที่ระบบอาจอ่านผิด:\n\n${ps
            .slice(0, 6)
            .map(x => '• ' + x.f + ': ' + x.t)
            .join('\n')}${ps.length > 6 ? '\n…' : ''}\n\nเติมต่อเลยไหม? (กด Cancel เพื่อกลับไปตรวจ)`,
        )
      )
        return;
    }
    if (
      last &&
      !last.saved &&
      !confirm(
        'ผลการเติมครั้งก่อนยังไม่ได้บันทึกหรือดาวน์โหลด ถ้าเติมใหม่ ผลเดิมจะหายไป ต้องการเติมใหม่เลยไหม?',
      )
    )
      return;
    FILLING = true;
    try {
      const on = buildOn(o),
        ms = DB.find(s => s.m === 'MetroLED'),
        metro = ms ? listOf(ms, o) : [],
        out = new JSZip(),
        res = [],
        L = ['รายงานการเติมใบเช็ครูป', ...plRep()];
      let i = 0;
      for (const f of list) {
        i++;
        await busyStep('กำลังเติมใบเช็ครูป', i, list.length, f.name);
        let r;
        const rd = f.nd ? readyOf(f) : null;
        if (rd && (rd.bad || rd.empty))
          r = {
            rep: [
              [
                'ทั้งไฟล์',
                'ไม่ได้เติม',
                'ตรวจ: ' + (rd.bad || 'ไม่พบชีตของใบเช็ครูปในไฟล์นี้') + ' จึงไม่ได้เติม',
                0,
              ],
            ],
            blob: null,
          };
        else
          try {
            r = await fillWb(f, on, metro, {
              ...o,
              col1: effCol(1, f),
              col2: effCol(2, f),
              k16: !!fcol(f.name).k16,
            });
          } catch (e) {
            r = {
              rep: [
                [
                  'ทั้งไฟล์',
                  'เติมไฟล์นี้ไม่ได้',
                  'ตรวจ: เติมไฟล์นี้ไม่ได้ ไฟล์เสีย หรือไม่ใช่ใบเช็ครูป (.xlsm / .xlsx)',
                  0,
                ],
              ],
              blob: null,
            };
          }
        if (r.blob) out.file(f.name, r.blob, { createFolders: false });
        res.push({ name: f.name, rep: r.rep, blob: r.blob, fold: r.fold || [] });
        L.push(
          '',
          f.name,
          ...r.rep.map(x => `  ${x[0]} | ${x[1]} | ${x[2]}${x[3] ? ' (' + x[3] + ')' : ''}`),
        );
      }
      out.file('รายงาน.txt', L.join('\n'));
      last = {
        res,
        zip: await out.generateAsync({ type: 'blob' }),
        saved: false,
        W: o.W ? +o.W : Date.now(),
        txt: L.join('\n'),
      };
      note = 'เติมเสร็จแล้ว';
      tab = 'res';
      cs = null;
      RO.clear();
      window.scrollTo(0, 0);
      const fi = res.findIndex(r => r.rep.some(isWarn));
      if (fi >= 0) RO.add(fi);
    } finally {
      FILLING = false;
      BUSY = null;
      $('.wrap').inert = false;
    }
  },
};
function g(fn) {
  const t = TQ.toLowerCase();
  if (!t) return;
  let c = 0;
  for (const s of DB) {
    let h = 0;
    s.rows = s.rows.filter(r => {
      if (!String(r.ad).toLowerCase().includes(t)) return true;
      c++;
      h++;
      if (fn) fn(r);
      return !!fn;
    });
    if (h) {
      s.rows = stable(s.rows);
      touch(s);
    }
  }
  note = `แก้ไข ${c} แถว`;
}
document.addEventListener('click', async e => {
  const a = e.target.closest('[data-a]');
  if (!a || a.tagName === 'INPUT') return;
  if (a.dataset.a === 'file') return;
  if (a.disabled) return;
  try {
    await ACT[a.dataset.a](a.dataset.k);
  } catch (x) {
    note = 'ผิดพลาด: ' + (x.message || x);
  }
  view();
});
document.addEventListener('change', e => {
  const t = e.target;
  if (t.id === 'o_base') {
    const raw = t.value.trim();
    if (!raw) {
      OP.base = '';
      BASEERR = '';
    } else {
      const d = parseUserDate(raw);
      if (d) {
        OP.base = end(d);
        BASEERR = '';
      } else {
        BASEERR = raw;
        OP.base = '';
      }
    }
    softView();
    return;
  }
  if (t.id && t.id.startsWith('o_')) {
    OP.media = $('#o_media').value;
    OP.sort = $('#o_sort').value;
    OP.newOnly = $('#o_new').checked;
    OP.dedupe = $('#o_dd').checked;
    OP.off = $('#o_off').checked;
    OP.hand = $('#o_hand').checked;
    OP.r2 = $('#o_r2').value;
    softView();
    return;
  }
  if (t.id === 'colpick') {
    OP['col' + t.dataset.k] = 'FF' + t.value.slice(1).toUpperCase();
    store.set('ps_col' + t.dataset.k, OP['col' + t.dataset.k]);
    view();
    return;
  }
  if (t.dataset.f16 != null) {
    const m = fcol(t.dataset.f16);
    m.k16 = t.checked ? 1 : 0;
    store.set(fkey(t.dataset.f16), JSON.stringify(m));
    softView();
    return;
  }
  if (t.dataset.fcol != null) {
    const m = fcol(t.dataset.fcol);
    m[t.dataset.r] = t.value;
    store.set(fkey(t.dataset.fcol), JSON.stringify(m));
    softView();
    return;
  }
  if (t.id === 'picker') {
    const fs = [...t.files];
    t.value = '';
    addAny(t.dataset.k, fs);
    return;
  }
  const s = curS();
  if (!s) return;
  if (t.dataset.f === 'on') {
    const r = s.rows[+t.dataset.i];
    r.off = !t.checked;
    s.rows = stable(s.rows);
    touch(s);
    FOCUSSEL = `[data-f="on"][data-i="${s.rows.indexOf(r)}"]`;
    view();
    return;
  }
  if (t.dataset.f === 'st' && /^ถอด/.test(t.value.trim())) {
    const r = s.rows[+t.dataset.i];
    if (r && !r.off) {
      r.off = true;
      s.rows = stable(s.rows);
      touch(s);
      FOCUSSEL = `[data-f="st"][data-i="${s.rows.indexOf(r)}"]`;
      note = `ปลด “${r.ad}” แล้ว (สถานะถอด)`;
      view();
    }
  }
});
document.addEventListener('input', e => {
  const t = e.target;
  if (t.dataset.f && t.dataset.f !== 'on') {
    const s = curS();
    if (!s) return;
    const r = s.rows[+t.dataset.i],
      f = t.dataset.f;
    if (!r) return;
    r[f] = f === 'len' || f === 'cnt' ? +t.value || 0 : t.value;
    touch(s);
    $('#steps').innerHTML = vSteps();
  } else if (t.id === 'q') {
    q = t.value;
    view();
  } else if (t.id === 'tq') {
    TQ = t.value;
    view();
  }
});
document.addEventListener('dragover', e => {
  e.preventDefault();
});
document.addEventListener('drop', e => {
  e.preventDefault();
  const fs = [...((e.dataTransfer && e.dataTransfer.files) || [])];
  if (!fs.length) return;
  const z = e.target.closest && e.target.closest('[data-dz]');
  let k = z ? z.dataset.dz : tab === 'fl' ? 'ck' : tab === 'dat' ? 'pl' : null;
  if (!k) {
    if (fs.every(f => /\.json$/i.test(f.name))) k = 'pj';
    else {
      say('ลากไฟล์ Playlist ไปวางในขั้นที่ 1 หรือใบเช็ครูปในขั้นที่ 3');
      return;
    }
  }
  if (TOUR) {
    TOUR = null;
    view();
  }
  addAny(k, fs);
});
window.addEventListener('beforeunload', e => {
  if (dirty || (last && !last.saved)) {
    e.preventDefault();
    e.returnValue = '';
  }
});
view();
