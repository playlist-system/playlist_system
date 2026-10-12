/*
 * ======================================================================
 *  ระบบเติมใบเช็ครูปอัตโนมัติ — โค้ด JavaScript ของระบบ
 *  เวอร์ชัน 2.1 • 12 ต.ค. 2569
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
 *    6. เปิดใบเช็ค จับคู่ชีตกับจอ แล้วเติม ............ fillWb
 *    7. อ่าน Playlist 5 รูปแบบ ....................... parseSig, parseMetro, parsePlaylist, addFiles
 *    8. เช็กลิสต์ Playlist 11 แบบ / ความพร้อมใบเช็ค .... KINDS, kState, sheetNeeds, readyOf
 *    9. หน้าจอ 4 ขั้นตอน ............................. vData, vWall, vPl, vEd, vTl, vFl, vRes
 *   10. วิธีใช้ / ต้อนรับ / ระบบทำงานอย่างไร ............ vHelp, vWelcome, vHow
 *   11. ทัวร์แนะนำ ................................... TOURS, tourStart, tourDraw
 *   12. ปุ่มและเหตุการณ์ .............................. view, ACT, document.addEventListener
 *  (ส่วน 1–7 คือส่วนอ่าน Playlist และเติมใบเช็ค แก้จาก 1.0 เพียงเล็กน้อย ผลการเติมเหมือนเดิมทุกไบต์)
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
  setF(col, n, f) {
    const c = this.cell(col, n);
    while (c.firstChild) c.removeChild(c.firstChild);
    c.removeAttribute('t');
    const e = this.doc.createElementNS(NS, 'f');
    e.textContent = f;
    c.appendChild(e);
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
const JUNK = /^(เวลา|รวม|หนังโฆษณา|Version$)|^\d+(\.\d+)?$|^Programm?atic|^ไม่มีไฟล์|เมื่อมีไฟล์/i,
  MTP = /^(Traffic MTP|PLACE MTP|ETA MTP|EDA)\s*$/i;
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
  PLX = []; /* PLM: สื่อ → {f:ชื่อไฟล์, d:วันที่} • PLX: ไฟล์ที่อ่านไม่ได้ */
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
    if (s.m !== 'Show DC / Prudential' && s.pd > b) b = s.pd;
  }
  b = b || a;
  return b ? new Date(b) : null;
}; /* Show DC อาจเป็นไฟล์ของรอบถัดไป จึงไม่ใช้กำหนดวันทำงาน */
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
      !/FastAds?/i.test(r.ad) &&
      !skipRow(scr, r) &&
      !(kind === 'jam' && (/^BJ\d*_\d+_/i.test(r.ad) || /weather/i.test(r.ad))),
  ).map(r => ({
    ad: kind === 'jam' ? r.ad.replace(/^BKJ[12]?_/i, '') : r.ad,
    hash: /^Content/i.test(r.ty || ''),
    sp: kind === 'metro' && MTP.test(r.ad.trim()),
    sd: parseD(r.s),
    ed: parseD(r.e),
    st: r.st || '',
    off: r.off,
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
const MARK = /^(FF)?(8AEDFE|82ECFE|002060|375623|FFFF00)$/i;
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
  const { two, P1, P, E } = per(W),
    SCAN = 1500,
    dd = +W,
    noNotes = W.getUTCDate() === 1;
  const B = jobs.map(j => {
    const kind = j[4],
      all = chkRows(j[3], { ...o, off: true }, kind),
      lc = j[1],
      nc = pv(lc),
      pc = nx(lc);
    return {
      lc,
      nc,
      pc,
      kind,
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
  const barD = r => {
    const x = wsv && wsv[b0.nc + r];
    if (!(x && typeof x.v === 'number' && x.v > 20000 && x.v < 80000)) return null;
    const d = s2d(x.v),
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
    if (B.some(b => vOf(b.lc, r) || /^ออกวันที่/.test(vOf(b.nc, r)))) last = r;
  }
  const F = bars.length && bars[0][0] === hr ? bars[0][1] : null;
  let S = hr,
    hist = false,
    frozen = false;
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
    !two &&
    F &&
    +F === +P1 &&
    bars.some(
      ([r, d]) => r > hr && d && d.getUTCDate() >= 16 && ISG.test(ST.fillOf(sh.getS(b0.lc, r) || '0')),
    )
  )
    return {
      skip: 'ชีตนี้มีรอบ 16–สิ้นเดือน (แถบเขียว) แล้ว แต่วันที่ทำงานอยู่ในรอบ 1–15 — ไม่ได้แก้ชีตนี้ ตรวจวันที่ทำงานอีกครั้ง',
    };
  const Z = Math.max(last + 2, S + 400),
    lim = Math.min(+addD(W, +o.days || 0), +E),
    H = lim; /* แบบที่ทำด้วยมือ: ใส่หนังที่เริ่มถึงวันนี้ และ “ออกวันที่” ถึงวันนี้ ไม่ใส่ล่วงหน้า */
  const exitOf = (n, ref) => {
    let d = new Date(Date.UTC(ref.getUTCFullYear(), ref.getUTCMonth(), n));
    if (n < ref.getUTCDate()) d = new Date(Date.UTC(ref.getUTCFullYear(), ref.getUTCMonth() + 1, n));
    return +addD(d, -1);
  };
  /* เก็บตัวเลขที่ผู้ตรวจนับไว้ (ช่องข้างชื่อหนัง) — ย้ายตามชื่อหนังเฉพาะในรอบเดียวกัน */
  let moved = 0,
    dropped = 0,
    cleared = 0,
    kept = 0,
    merged = 0,
    gone = 0,
    today = 0,
    nost = 0;
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
      if (!nm || /^ออกวันที่/.test(nm)) continue;
      const k = sh.take(b.pc, r);
      if (k) {
        sh.setS(b.pc, r, b.bp);
        if (hist) {
          if (!b.cnt.has(nm)) b.cnt.set(nm, []);
          b.cnt.get(nm).push(k);
        } else cleared++;
      }
      if (hist) {
        const m = vOf(b.nc, r).match(/^ออกวันที่\s*(\d{1,2})/);
        b.hist.push({ nm, g, ned: m ? exitOf(+m[1], P) : 0 });
      }
    }
  }
  /* รวมรายการ: ของเดิมในชีต + Playlist วันนี้ */
  for (const b of B) {
    const offEd = new Map(),
      pl = [],
      ent = [];
    b.later = [];
    b.sp = b.rows.filter(r => r.sp);
    {
      const cap = b.scr.pd
        ? b.scr.pd - 864e5
        : Infinity; /* อยู่ในรายการถอดของ Playlist วันไหน = ออกไม่เกินวันนั้น */
      for (const r of b.scr.rows)
        if (r.off && r.ad) {
          const e0 = parseD(r.e),
            ed = Math.min(e0 ? +e0 : Infinity, cap);
          if (!isFinite(ed) || ed < +P) continue;
          const k =
            (/^Content/i.test(r.ty || '') ? '#' : '') +
            (b.kind === 'jam' ? r.ad.replace(/^BKJ[12]?_/i, '') : r.ad);
          offEd.set(k, Math.max(offEd.get(k) || 0, ed));
        }
    } /* วันหมดของหนังที่ถอดแล้ว (ใช้ทำ “ออกวันที่”) */
    for (const r of b.rows) {
      if (r.sp) continue;
      if (r.ed && +r.ed < +P) continue;
      const g = r.sd && +r.sd > +P ? +r.sd : +P;
      if (g > lim) {
        b.later.push(r);
        continue;
      }
      pl.push({
        nm: (r.hash ? '#' : '') + r.ad,
        hash: r.hash,
        g,
        ed: r.ed ? +r.ed : Infinity,
        st: r.st || '',
      });
    }
    pl.sort((x, y) => x.g - y.g);
    for (const h of b.hist)
      ent.push({
        nm: h.nm,
        hash: h.nm.startsWith('#'),
        g: h.g,
        ed: h.ned || offEd.get(h.nm) || 0,
        h: true,
        m: false,
      });
    /* หนังที่ยังไม่เคยลงในชีต: เริ่มก่อน/ในวันแรกของรอบ → ลงใต้แถบ “วันนี้” (ถ้าชีตมีรายการของรอบนี้อยู่แล้ว) ไม่เอาไปรวมกับวันแรก */
    const hb = b.hist.length > 0,
      gate =
        hb &&
        dd > +P &&
        b.all.some(
          r => r.st,
        ); /* ตั้งแต่วันที่ 2 ของรอบ: หนังที่ยังไม่มีในใบเช็ค เพิ่มเฉพาะสถานะ “ขึ้นใหม่…” (หรือเริ่มวันนี้/ล่วงหน้า) */
    for (const p of pl) {
      const c = ent.filter(e => e.nm === p.nm),
        tg = p.g > +P ? p.g : hb ? dd : +P;
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
      }
      t = c.find(e => e.h && !e.m && e.g > dd); /* เคยวางล่วงหน้าไว้ แต่วันเริ่มเปลี่ยน → ย้ายตาม Playlist */
      if (t) {
        t.g = tg;
        t.ed = p.ed;
        t.m = true;
        continue;
      }
      t = c.filter(e => e.g <= tg).pop(); /* ไฟล์เดิมต่อเนื่อง หรือยังออนอยู่ → แถวเดิม ไม่เพิ่มซ้ำ */
      if (t && (CONT.test(p.st) || !t.ed || t.ed >= p.g)) {
        t.ed = Math.max(t.ed || 0, p.ed);
        t.m = true;
        merged++;
        continue;
      }
      if (gate && !NEWST.test(p.st) && p.g === +P) {
        nost++;
        continue;
      } /* สัญญาเริ่มก่อนรอบนี้ แต่ไม่เคยลงและไม่ได้ “ขึ้นใหม่” → ไม่เพิ่ม; ถ้าเริ่มในรอบนี้ ลงตามวันเริ่ม (กันพลาดเมื่อข้ามวัน) */
      if (tg !== p.g) today++;
      ent.push({ nm: p.nm, hash: p.hash, g: tg, ed: p.ed, m: true });
    }
    b.ent = ent.filter(e => {
      if (e.m) return true;
      if (e.g > dd) {
        gone++;
        return false;
      }
      kept++;
      return true;
    });
  }
  const seen = new Set();
  for (const b of B) for (const e of b.ent) if (e.g > +P) seen.add(e.g);
  const dates = [...seen].sort((x, y) => x - y);
  const srt = L => L.sort((x, y) => y.hash - x.hash || cmpN(x.nm, y.nm));
  for (const b of B) {
    b.main = srt(b.ent.filter(e => e.g === +P));
    b.grp = new Map(dates.map(k => [k, srt(b.ent.filter(e => e.g === k))]));
  }
  const note = (b, r, row) => {
    if (noNotes || !r.ed || !isFinite(r.ed)) return;
    const nd = +addD(new Date(r.ed), 1);
    if (nd > +P && nd <= H) {
      sh.setText(b.nc, row, 'ออกวันที่ ' + new Date(nd).getUTCDate());
      sh.setS(b.nc, row, ST.derive(b.bn, 'red', { color: 'FFFF0000', right: true }));
    }
  };
  const put = (b, r, row) => {
    const q = b.cnt.get(r.nm);
    if (q && q.length) {
      sh.give(b.pc, row, q.shift());
      moved++;
    }
    sh.setText(b.lc, row, r.nm);
    if (r.hash) sh.setS(b.lc, row, ST.derive(b.bl, 'yel', { fill: 'FFFFFF00' }));
    else sh.setS(b.lc, row, b.bl);
    note(b, r, row);
  };
  for (const b of B) {
    for (let r = S + 1; r <= Z; r++) {
      {
        const ps = sh.getS(b.pc, r);
        if (ps != null && (ps === b.hs[2] || MARK.test(ST.fillOf(ps)))) sh.setS(b.pc, r, b.bp);
      }
      sh.clear(b.lc, r);
      sh.clear(b.nc, r);
      sh.setS(b.lc, r, b.be);
      {
        const ns = sh.getS(b.nc, r);
        if (b.bn != null) sh.setS(b.nc, r, b.bn);
        else if (ns != null && MARK.test(ST.fillOf(ns))) sh.setS(b.nc, r, null);
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
    c1 = Math.max(...B.map(b => col2n(b.pc))) + 12,
    own = new Set(B.flatMap(b => [b.nc, b.lc, b.pc])),
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
      if (st != null && MARK.test(ST.fillOf(st))) sh.setS(L, r, base);
    }
  }
  const hsx = st => {
    if (st == null) return st;
    const g = ISG.test(ST.fillOf(st));
    return two
      ? g
        ? st
        : ST.derive(st, 'grn', { fill: GRN })
      : g
        ? ST.derive(st, 'nvy', { fill: NAVY })
        : st;
  };
  const hdr = (row, d) => {
    for (const [L, st] of HX) {
      sh.clear(L, row);
      sh.setS(L, row, hsx(st));
    }
    B.forEach((b, i) => {
      sh.clear(b.lc, row);
      if (i === 0) sh.setNum(b.nc, row, serial(d));
      else sh.setF(b.nc, row, B[i - 1].nc + row);
      if (b.hs[0] != null) sh.setS(b.nc, row, hsx(b.hs[0]));
      if (b.hs[1] != null) sh.setS(b.lc, row, hsx(b.hs[1]));
      sh.clear(b.pc, row);
      if (b.hs[2] != null) sh.setS(b.pc, row, hsx(b.hs[2]));
    });
  };
  hdr(S, P);
  const pad = (b, a, z) => {
    for (let row = a; row <= z; row++) {
      sh.setS(b.lc, row, ST.derive(b.bl, 'pad' + PADC, { fill: PADC, border: true }));
      sh.setS(b.pc, row, ST.derive(b.bp != null ? b.bp : '0', 'pad' + PADC, { fill: PADC, border: true }));
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
  for (const b of B) for (const q of b.cnt.values()) dropped += q.length;
  return {
    hr,
    S,
    H,
    P,
    E,
    two,
    frozen,
    noNotes,
    hist,
    dates: dates.length,
    later: B.reduce((n, b) => n + b.later.length, 0),
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

async function fillWb(f, on, metro, o) {
  const rep = [];
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
    if (media && brow === 11) {
      const av = [];
      for (const k of Object.keys(ws)) {
        const m = k.match(/^([A-Z]+)9$/);
        if (m && ws[k].f && /^AVERAGE/.test(ws[k].f)) av.push(col2n(m[1]));
      }
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
    for (const [cn, code] of blocks) {
      let key = media + '|' + code,
        ads = on.get(key);
      let fb = '';
      if (!ads && media === 'CITY' && on.has('CITY|THE CITY')) {
        key = 'CITY|THE CITY';
        ads = on.get(key);
        fb = ' (ไม่มีจอนี้ใน Playlist CITY — ใช้รายการ THE CITY แทน)';
      }
      if (!ads) {
        const hit = [...on.keys()].filter(k => k.split('|')[1] === code);
        if (hit.length === 1) {
          key = hit[0];
          ads = on.get(key);
          fb =
            ' (ใช้รายการของสื่อ ' +
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
    if (!blocks.length && /METRO/i.test(sn.replace(/[^A-Za-z]/g, '')) && metro.length) {
      let first = 0;
      for (const k of Object.keys(ws)) {
        const m = k.match(/^([A-Z]+)9$/);
        if (m && ws[k].f && /^AVERAGE/.test(ws[k].f)) {
          const n = col2n(m[1]);
          if (!first || n < first) first = n;
        }
      }
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
      const av = [];
      for (const k of Object.keys(ws)) {
        const m = k.match(/^([A-Z]+)9$/);
        if (m && ws[k].f && /^AVERAGE/.test(ws[k].f)) av.push(col2n(m[1]));
      }
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
      const av = [];
      for (const k of Object.keys(ws)) {
        const m = k.match(/^([A-Z]+)9$/);
        if (m && ws[k].f && /^AVERAGE/.test(ws[k].f)) av.push(col2n(m[1]));
      }
      av.sort((x, y) => x - y);
      if (!av.length) rep.push([sn, '-', 'ข้าม (ไม่พบรหัสจอ และไม่พบคอลัมน์ผลวิเคราะห์ในแถว 9)', 0]);
      for (const cn of av.filter(c => !av.includes(c - 1))) {
        const lc = n2col(cn - 1),
          label = cl((ws[lc + '10'] || {}).v),
          code = cl((ws[n2col(cn) + '11'] || {}).v) + cl((ws[n2col(cn) + '12'] || {}).v),
          scr = namedScreen(label, code, sn);
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
  const by = {};
  for (const j of jobs) (by[j[0]] = by[j[0]] || []).push(j);
  const ST = o.hand
      ? mkStyles(
          new DOMParser().parseFromString(await zip.file('xl/styles.xml').async('string'), 'application/xml'),
        )
      : null,
    W = o.hand ? o.W : null;
  for (const sn in by) {
    const sh = new Sheet(await zip.file(path[sn]).async('string'));
    if (ST && W && by[sn].every(j => j[3])) {
      const L = layoutChk(sh, by[sn], o, ST, W, wb.Sheets[sn]);
      if (L.skip) rep.push([sn, 'รูปแบบ', 'ตรวจ: ' + L.skip, 0]);
      else {
        if (L.dropped)
          rep.push([
            sn,
            'ตัวเลขนับ',
            'ตรวจ: มีตัวเลขที่นับไว้ ' + L.dropped + ' ช่องของหนังที่ไม่อยู่ในรายการแล้ว จึงไม่ได้ย้ายมา',
            0,
          ]);
        const rd = d => new Date(d).getUTCDate();
        rep.push([
          sn,
          'รูปแบบ',
          (L.two
            ? 'รอบ 16–' +
              rd(L.E) +
              ' (แถบวันที่สีเขียว) เริ่มแถว ' +
              L.S +
              (L.frozen ? ' ต่อท้ายรอบ 1–15 (รอบแรกไม่ถูกแก้)' : '')
            : 'รอบ 1–15 หัววันที่แถว ' + L.hr) +
            (L.noNotes ? ' • วันที่ 1 ยังไม่ใส่ “ออกวันที่”' : '') +
            (L.moved ? ' • ย้ายตัวเลขนับตามชื่อหนัง ' + L.moved + ' ช่อง' : '') +
            (L.today ? ' • หนังใหม่ของวันนี้ ' + L.today + ' เรื่อง ลงใต้แถบวันที่ ' + rd(L.W) : '') +
            (L.nost ? ' • ไม่เพิ่ม ' + L.nost + ' เรื่องที่ยังไม่มีในใบเช็คแต่สถานะไม่ใช่ “ขึ้นใหม่”' : '') +
            (L.dates ? ' • แยกวันเริ่ม ' + L.dates + ' กลุ่ม' : '') +
            (L.merged ? ' • ไฟล์เดิม/ต่อสัญญา ' + L.merged + ' แถว ใช้แถวเดิม ไม่ใส่ออกวันที่' : '') +
            (L.kept ? ' • เก็บหนังที่ลงไว้แล้วแต่ไม่มีใน Playlist วันนี้ ' + L.kept + ' แถว' : '') +
            (L.cleared ? ' • เริ่มรอบใหม่ ล้างตัวเลขนับของรอบก่อน ' + L.cleared + ' ช่อง' : '') +
            (L.later ? ' • ยังไม่ใส่ ' + L.later + ' เรื่องที่เริ่มหลัง ' + fd(new Date(L.H)) : ''),
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
                      !/FastAds?/i.test(r.ad),
                  )
                  .map(r => r.ad),
              ),
            ];
          if (z.length)
            rep.push([
              sn,
              j[3].c,
              (sig ? 'ตรวจ: ไม่ใส่ตามหมายเหตุ “ไม่ออน/ออนเฉพาะ” → ' : 'ตรวจ: ไม่ใส่เพราะจำนวนครั้ง = 0 → ') +
                z.join(', '),
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
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      new XMLSerializer().serializeToString(wx.documentElement),
    { createFolders: false },
  );
  return { rep, blob: await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }) };
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
const OP = { media: 'ทั้งหมด', sort: 'orig', newOnly: false, dedupe: true, off: false, hand: true, base: '' };
let DB = [],
  tab = 'dat',
  fm = 'ทั้งหมด',
  q = '',
  cs = null,
  C = [],
  last = null,
  note = '',
  dirty = 0;
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
  PrudentialSathorn: ['สาทร', 'สาธร', 'TWTPLUS', 'PRUDENTIAL'],
  'Air Traveller Gateway': ['AOT'],
  'The Big Y': ['AOT', 'BIGY', 'THEBIGY'],
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
  m = n.match(/ออนเ?ฉพาะ\s*(.*)/);
  if (m) {
    const t = ntoks(m[1].split('//')[0]);
    return t.length > 0 && !t.some(x => al.includes(x));
  }
  return false;
}
const skipRow = (scr, r) => (SIGAL[scr.c] ? noteSkip(scr, r) : r.cnt === 0 && !MTP.test(r.ad.trim()));
/* ไฟล์ Playlist แบบ Signature Max: 1 ชีต = 1 จอ หัวตารางมีคำว่า หนังโฆษณา และ “รายการปลด” แยกส่วนที่ถอดแล้ว */
const SIGCODE = { 'Exchange Tower': 'Exchange Tower (SM)', 'Exchange Tower (2)': 'Exchange Tower (SM) 2' };
/* ไฟล์แบบ Signature Max / Show DC / BKJ: 1 ชีต = 1 จอ หัวตารางมีคำว่า หนังโฆษณา และ “รายการปลด” แยกส่วนที่ถอดแล้ว */
function parseSig(wb, media, map) {
  media = media || 'Signature Max';
  const out = [];
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    if (!ws || !/^รายการหนังโฆษณา/.test(cl((ws.A1 || {}).v))) continue;
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
function parseMetro(wb) {
  let sn = wb.SheetNames[wb.SheetNames.length - 1],
    best = -1;
  for (const n of wb.SheetNames) {
    const d = fullNameDate(n);
    if (d && +d > best) {
      best = +d;
      sn = n;
    }
  }
  const A = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: '' });
  const h = A.slice(0, 15).findIndex(r => r.some(x => cl(x) === 'Version'));
  let cV = 5,
    oS = 14,
    oE = 15,
    cS = -1,
    cE = -1;
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
    ge = '';
  const rows = [];
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
    rows.push({
      ad,
      ty: cl(row[3]),
      len: +row[cV + 1] || 0,
      cnt: +row[cV + 2] || 0,
      st: '',
      s: fd(g(row, oS)) || fd(g(row, cS)) || gs,
      e: fd(g(row, oE)) || fd(g(row, cE)) || ge,
      note: cl(row[2]),
      off,
    });
  }
  return [fixMetro({ m: 'MetroLED', c: 'METRO', n: '7-Eleven Metro LED', f: 'LED', ed: 0, rows })];
}
function parsePlaylist(f) {
  const out = [];
  if (/Metro.*LED/i.test(f.name)) return parseMetro(XLSX.read(f.buf, { type: 'array', cellDates: true }));
  if (!/แยกจอ/.test(f.name)) {
    const wb = XLSX.read(f.buf, { type: 'array' });
    if (wb.SheetNames.some(n => bkjCode(n))) {
      const sg = parseSig(wb, 'Bangkok Jams', bkjCode);
      return sg.length ? sg : null;
    }
    const sg = parseSig(
      wb,
      /show\s*_*\s*dc|prudential/i.test(f.name) ? 'Show DC / Prudential' : 'Signature Max',
    );
    return sg.length ? sg : null;
  }
  /* สรุปรายการหนังแยกจอ: สื่ออยู่ในชื่อไฟล์ เช่น …แยกจอ_01_40Blue, …แยกจอ_2026_10_02_80Red */
  const bn = f.name.replace(/\.xlsx$/i, '').replace(/\s*\(\d+\)$/, ''),
    rest = bn
      .replace(/^.*แยกจอ/, '')
      .replace(/\d{4}[-_.]\d{1,2}[-_.]\d{1,2}/g, '')
      .replace(/\d{1,2}[-_.]\d{1,2}[-_.]\d{2,4}/g, '')
      .replace(/^[_\s]+|[_\s]+$/g, '')
      .replace(/^\d{1,2}[_\s]+(?=\D)/, ''),
    media = mediaOf(rest) || rest;
  if (!media) return null;
  const wb = XLSX.read(f.buf, { type: 'array' }),
    names = {};
  if (wb.Sheets.Index)
    for (const r of XLSX.utils.sheet_to_json(wb.Sheets.Index, { header: 1, defval: '' }).slice(2))
      for (const i of [1, 5, 9]) {
        const nm = cl(r[i]);
        if (nm) names[nm.split(/[_ ]/)[0]] = nm;
      }
  for (const sn of wb.SheetNames) {
    if (sn === 'Index' || sn.startsWith('data') || sn.startsWith('Blank')) continue;
    const A = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: '' });
    if (!A.length) continue;
    for (let c0 = 1; c0 < A[0].length; c0 += 9) {
      const code = cl(A[0][c0]);
      if (!code) continue;
      let off = false;
      const rows = [];
      for (let r = 2; r < A.length; r++) {
        const ad = cl(A[r][c0]);
        if (!ad) continue;
        if (ad === 'หนังโฆษณา') {
          off = true;
          continue;
        }
        if (/^[\d.]+$/.test(ad) && !cl(A[r][c0 + 1])) continue;
        if (cl(A[r][c0 + 4]) === 'ถอด') off = true;
        rows.push({
          ad,
          ty: cl(A[r][c0 + 1]),
          len: +A[r][c0 + 2] || 0,
          cnt: +A[r][c0 + 3] || 0,
          st: cl(A[r][c0 + 4]),
          s: fd(A[r][c0 + 6]),
          e: fd(A[r][c0 + 7]),
          note: cl(A[r][c0 + 8]),
          off,
        });
      }
      out.push({
        m: media,
        c: code,
        n: names[code] || code,
        f: cl(A[0][c0 + 2]).replace(/[, ]+$/, ''),
        ed: 0,
        rows,
      });
    }
  }
  return out.length ? out : null;
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
      p = parsePlaylist(f);
    } catch (e) {
      p = null;
    }
    PLX = PLX.filter(x => x.f !== f.name);
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
    const fd0 = fullNameDate(f.name) || dataDate(p) || nameDate(f.name, autoBase());
    for (const x of p) {
      x.pf = f.name;
      x.pd = fd0 ? +fd0 : 0;
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
const VERTXT = 'ระบบเติมใบเช็ครูป เวอร์ชัน 2.1 (12 ต.ค. 2569)';
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
    wb = XLSX.read(buf.slice(0), { type: 'array', sheetRows: 13 });
  } catch (e) {
    return { bad: 'ไฟล์เสียหรือเปิดไม่ได้' };
  }
  if (!wb || !wb.SheetNames || !wb.SheetNames.length) return { bad: 'ไม่พบชีตในไฟล์นี้' };
  if (
    wb.SheetNames.includes('Index') ||
    wb.SheetNames.some(n => /^รายการหนังโฆษณา/.test(cl(((wb.Sheets[n] || {}).A1 || {}).v))) ||
    /สรุปรายการหนัง|รายการหนัง/.test(String(name || ''))
  )
    return { bad: 'ไฟล์นี้เป็นไฟล์ Playlist ไม่ใช่ใบเช็ครูป (ใส่ไว้ในขั้นที่ 1)' };
  const S = [],
    named = [],
    avOf = ws => {
      const av = [];
      for (const key of Object.keys(ws)) {
        const m = key.match(/^([A-Z]+)9$/);
        if (m && ws[key].f && /^AVERAGE/.test(ws[key].f)) av.push(col2n(m[1]));
      }
      return av.sort((x, y) => x - y);
    };
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    if (!ws) continue;
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
    const names = [];
    if (media && brow === 11)
      for (const cn of avOf(ws)) {
        if (blocks.some(x => x[0] === cn)) continue;
        const raw = cl((ws[n2col(cn) + '11'] || {}).v);
        if (raw && nrm(raw).length >= 3) names.push(raw);
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
    for (const cn of av.filter(c => !av.includes(c - 1))) {
      const lc = n2col(cn - 1);
      named.push({
        sn,
        label: cl((ws[lc + '10'] || {}).v),
        code: cl((ws[n2col(cn) + '11'] || {}).v) + cl((ws[n2col(cn) + '12'] || {}).v),
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
      if (e.media === 'CITY' && on.has('CITY|THE CITY')) {
        add('CITY', e.sn, true);
        warn.push(`${code} ใช้รายการ THE CITY`);
        continue;
      }
      const h = keys.filter(k => k.split('|')[1] === code);
      if (h.length === 1) {
        const m = h[0].split('|')[0];
        add(m, e.sn, true);
        warn.push(`${code} (${e.sn}) ใช้รายการของ ${m}`);
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
    const n = KINDS.filter(K => DB.some(K.has)).length;
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
  const pr = per(wd);
  return `<span class="rnd ${pr.two ? 'p2' : ''}"><i>${end(wd)}</i>${pr.two ? `รอบ 16–${pr.E.getUTCDate()} ${TH[pr.E.getUTCMonth()]} แถบสีเขียว` : `รอบ 1–15 ${TH[pr.E.getUTCMonth()]} แถบสีน้ำเงิน`}</span>`;
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
  const S = KINDS.map(K => ({ K, s: kState(K, wd) })),
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
   .join('')}</div></div>
 ${sum}
 ${n && n < 11 ? `<p class="wnote">ถ้าเติมเฉพาะใบเช็คของตัวเอง ใส่เฉพาะแบบที่ใบเช็คนั้นใช้ก็ได้ ขั้นที่ 3 จะบอกว่าใบเช็คแต่ละไฟล์ต้องใช้แบบไหน</p>` : ''}
 ${mis.length ? `<p class="wnote">ไฟล์ที่ใส่ทีหลังจะแทนที่ไฟล์เดิมของจอเดียวกัน ใส่ไฟล์ของวันที่ถูกต้องอีกครั้งเพื่อแทนที่</p>` : ''}
 ${extra.length ? `<p class="wnote">ไฟล์อื่นที่อ่านได้: ${extra.map(esc).join(', ')}</p>` : ''}</section>`;
}
function vData() {
  const has = DB.length > 0,
    wd = has ? workDate() : null,
    sdc = KINDS.find(K => K.latest);
  const fl = (() => {
    const M = new Map();
    for (const x of DB) if (x.pf && (!M.has(x.pf) || x.pd > M.get(x.pf))) M.set(x.pf, x.pd || 0);
    return [...M];
  })();
  const isSdc = ([f]) =>
    DB.some(s => s.pf === f && sdc.has(s)); /* Show DC มีช่องบอกสถานะของตัวเอง ไม่นับเป็น “หลายวันปนกัน” */
  const ds = [
    ...new Set(
      fl
        .filter(x => !isSdc(x))
        .map(x => x[1])
        .filter(Boolean),
    ),
  ].sort((a, b) => a - b);
  return `${shead('ใส่ Playlist ของวันนี้', 'ลากไฟล์ Playlist ทุกไฟล์ของวันนี้มาวางพร้อมกัน ช่องในเช็กลิสต์จะสว่างเมื่อได้ไฟล์แบบนั้นแล้ว')}
 <div class="s1">${drop('pl', has ? 'เพิ่มหรือเปลี่ยนไฟล์ Playlist' : 'ลากไฟล์ Playlist มาวางที่นี่', 'หรือคลิกเพื่อเลือกไฟล์ .xlsx เลือกหลายไฟล์พร้อมกันได้ ไฟล์ใหม่ของจอเดียวกันจะแทนที่ของเดิม', '.xlsx,.xlsm', 'drop-pl', has)}${vDateCard(wd, has)}</div>
 ${vWall(wd)}
 <div class="stack">${PLX.length ? alertHTML('er', `อ่านไม่ได้ ${PLX.length} ไฟล์ ไฟล์เหล่านี้ยังไม่ได้นำเข้าระบบ`, `<ul>${PLX.map(x => `<li><b>${esc(x.f)}</b><br>${esc(x.why)}</li>`).join('')}</ul><p>ตรวจว่าเป็นไฟล์ Playlist .xlsx ที่ถูกต้อง ไม่ใช่ PDF หรือไฟล์อื่นที่แนบมาในอีเมลเดียวกัน ถ้าถูกแล้วแต่ยังอ่านไม่ได้ ส่งไฟล์ให้ผู้ดูแลระบบ</p><button class="btn quiet" data-a="clrx">ซ่อนรายการนี้</button>`) : ''}
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
  const stc = t => (/ถอด/.test(t || '') ? 'bad' : /ใหม่/.test(t || '') ? 'new' : '');
  return `${shead('แก้หนังในจอ', 'สวิตช์เปิด = ออนแอร์ ปิด = ปลด (ใส่สถานะ “ถอด” ก็ปลดให้เช่นกัน) แก้ชื่อ สถานะ ความยาว และจำนวนครั้งได้ในตาราง ใช้ลูกศรเลื่อนลำดับ หรือถังขยะลบแถว', `<div class="go" data-tour="edback"><button class="btn secondary" data-a="back">${IC.back}กลับไปรายการจอ</button></div>`)}
 <section class="panel"><div class="edhead" data-tour="edhead"><h2>${esc(s.c)}</h2><span class="enm">${esc(scrName(s))}</span>${badge(s.m)}${s.f ? `<span class="fmt">Format ${esc(s.f)}</span>` : ''}</div>
 <div class="tools" data-tour="edtools"><button class="btn primary" data-a="add">+ เพิ่มหนัง</button><button class="btn quiet" data-a="paste">วางรายชื่อหลายบรรทัด</button></div>
 <div class="tbl" data-tour="edtbl"><table><thead><tr><th>ออนแอร์</th><th>ชื่อหนัง</th><th>Type</th><th>สถานะ</th><th class="n">วิ</th><th class="n">ครั้ง</th><th>หมดสัญญา</th><th>หมายเหตุ</th><th><span class="sr">จัดการ</span></th></tr></thead><tbody>
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
 <td><input class="w-e" data-f="e" data-i="${i}" value="${esc2(r.e)}" placeholder="-" aria-label="หมดสัญญา"></td><td><input class="w-nt" data-f="note" data-i="${i}" value="${esc2(r.note)}" placeholder="-" aria-label="หมายเหตุ"></td>
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
function vCk(f, r, i) {
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
  return `<div class="ck"><div class="ckh"><span class="nm">${esc(f.name)}${r && !r.bad && !r.empty ? `<small>${r.sheets} ชีต ใช้ Playlist ${r.items.length} แบบ</small>` : ''}</span>${pill}${rm}</div>${body}</div>`;
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
 <div class="s3"><div class="col">${drop('ck', C.length ? 'เพิ่มใบเช็ครูป' : 'ลากไฟล์ใบเช็ครูปมาวางที่นี่', 'ใบเช็คของวันก่อนหน้า (.xlsm หรือ .xlsx) หรือคลิกเพื่อเลือกไฟล์ เลือกหลายไฟล์ได้ ไฟล์ชื่อซ้ำจะแทนที่ของเดิม', '.xlsx,.xlsm', 'drop-ck', C.length > 0)}
 <section class="panel" data-tour="ready"><h3 class="sec-t">ใบเช็ครูปที่ใส่แล้ว<small>${C.length} ไฟล์</small></h3>
 ${C.length ? `<div class="ckl">${C.map((f, i) => vCk(f, R[i], i)).join('')}</div>` : '<p class="empty">ยังไม่มีใบเช็ครูป ลากไฟล์มาวางในกล่องด้านบน แล้วระบบจะบอกว่าแต่ละไฟล์ต้องใช้ Playlist แบบไหน</p>'}</section></div>
 <aside class="panel set" data-tour="settings"><h3 class="sec-t">การตั้งค่า</h3>
 <label class="date"><span>วันที่ทำงาน (วันของ Playlist)</span><input id="o_base" class="${BASEERR ? 'bad' : ''}" placeholder="${end(ab) || 'เช่น 12-Oct-26'}" autocomplete="off" aria-invalid="${!!BASEERR}">${dmsg}${wd && !BASEERR ? rndHTML(wd) + (wd.getUTCDate() === 1 ? '<small>วันที่ 1 ยังไม่ใส่ “ออกวันที่”</small>' : '') : ''}</label>
 <label class="tgl"><span>จัดแบบที่ทำด้วยมือ<small># ขึ้นก่อน (ไฮไลต์เหลือง) แล้ว A–Z แยกตามวันเริ่ม และใส่ “ออกวันที่” ควรเปิดไว้เสมอ</small></span><span class="sw"><input type="checkbox" id="o_hand"><span></span></span></label>
 <button class="adv" data-a="adv" aria-expanded="${adv}">ตัวเลือกเพิ่มเติม${IC.chev}</button>
 <div class="advb" ${adv ? '' : 'hidden'}>
 <label class="fld"><span>สื่อ</span><select id="o_media">${ms.map(m => `<option>${esc(m)}</option>`).join('')}</select></label>
 <label class="fld"><span>เรียงลำดับ</span><select id="o_sort"><option value="orig">ตามไฟล์ Playlist</option><option value="az">ก-ฮ / A-Z</option><option value="new">ขึ้นใหม่ก่อน</option></select></label>
 <label class="cb"><input type="checkbox" id="o_new"> เฉพาะหนัง “ขึ้นใหม่”</label>
 <label class="cb"><input type="checkbox" id="o_dd"> ตัดชื่อซ้ำในจอเดียวกัน</label>
 <label class="cb"><input type="checkbox" id="o_off"> รวมรายการปลดต่อท้าย (ใช้ตรวจว่าถอดแล้ว)</label>
 <p class="hint">“สื่อ” จำกัดให้เติมเฉพาะจอของสื่อที่เลือก ส่วน “เรียงลำดับ” มีผลเมื่อปิดการจัดแบบที่ทำด้วยมือ</p></div></aside></div>
 <div class="fillbar">${lines.join('')}<button class="btn primary big" data-a="fill" data-tour="fill" ${can ? '' : 'disabled'}>เติมใบเช็ครูป${C.length ? ` ${C.length} ไฟล์` : ''}</button>${BASEERR ? '<p class="hint">แก้วันที่ทำงานให้ถูกต้องก่อน แล้วปุ่มนี้จะกดได้</p>' : DB.length && !C.length ? '<p class="hint">ใส่ใบเช็ครูปอย่างน้อย 1 ไฟล์ แล้วปุ่มนี้จะกดได้</p>' : ''}</div>`;
}

/* ----- ขั้นที่ 4 ----- */
const isFilled = x => x[2].startsWith('เติมแล้ว');
const isWarn = x => (isFilled(x) ? /ใช้รายการ/.test(x[2]) : x[1] !== 'รูปแบบ' || /^ตรวจ/.test(x[2]));
function vRes() {
  if (!last) return vFl();
  const R = last.res.map(r => ({
    ok: r.rep.filter(isFilled).length,
    bad: r.rep.filter(isWarn),
    inf: r.rep.filter(x => x[1] === 'รูปแบบ' && !/^ตรวจ/.test(x[2])),
    sh: new Set(r.rep.filter(isFilled).map(x => x[0])).size,
  }));
  const tot = R.reduce((n, r) => n + r.ok, 0),
    nb = R.reduce((n, r) => n + r.bad.length, 0),
    nf = last.res.filter(r => r.blob).length,
    ne = last.res.length - nf;
  return `${shead('ผลการเติม', 'ไฟล์ที่เติมแล้วยังไม่ถูกบันทึกลงเครื่อง กดดาวน์โหลดก่อนปิดหน้านี้', `<div class="go"><button class="btn secondary" data-a="tab" data-k="fl">${IC.back}กลับไปเติมใหม่</button></div>`)}
 <div class="sum" data-tour="sum"><span class="k">เติมเสร็จ</span><span><b>${nf}</b>ไฟล์</span><span><b>${tot.toLocaleString()}</b>จอ</span><span><b>${nb}</b>ควรตรวจ</span><span class="sp"></span><button class="btn primary big" data-a="dlz" data-tour="dlz" ${nf ? '' : 'disabled'}>${IC.dl}ดาวน์โหลดทั้งหมด (.zip)</button></div>
 <div class="stack">${ne ? alertHTML('er', `${ne} ไฟล์ไม่ได้เติม`, '<p>ไฟล์เหล่านี้เสีย หรือไม่ใช่ใบเช็ครูป และไม่อยู่ในไฟล์ zip ไฟล์อื่นเติมตามปกติ ดูเหตุผลในรายการของไฟล์นั้น แล้วใส่ไฟล์ที่ถูกต้องและเติมใหม่</p>') : ''}
 ${nb ? alertHTML('wr', '“ควรตรวจ” คืออะไร', '<p>จอที่เติมไม่ได้ จอที่เติมจากรายการอื่นแทน หรือหนังที่ระบบตั้งใจไม่ใส่ เปิดดูในแต่ละไฟล์แล้วตรวจในใบเช็คก่อนใช้งาน ความหมายของข้อความที่พบบ่อยอยู่ในปุ่ม “วิธีใช้”</p>') : alertHTML('ok', 'ไม่มีรายการที่ต้องตรวจเพิ่ม', '<p>ดาวน์โหลดแล้วเปิดใบเช็คใน Excel ดูความถูกต้องอีกครั้งก่อนใช้งาน</p>')}</div>
 <div class="rlist" data-tour="rlist">${last.res
   .map((r, i) => {
     const o = R[i],
       open = RO.has(i);
     return `<div class="rfile"><div class="rhead"><span class="fn"><b>${esc(r.name)}</b><small>${r.blob ? o.sh + ' ชีต' : 'ไม่ได้เติม'}</small></span>${r.blob ? `<span class="pill ok">เติมแล้ว ${o.ok} จอ</span>` : '<span class="pill er">ไม่ได้เติม</span>'}${o.bad.length ? `<span class="pill wr">ควรตรวจ ${o.bad.length}</span>` : ''}
 <button class="ib" data-a="dlf" data-k="${i}" aria-label="ดาวน์โหลด ${esc(r.name)}" title="ดาวน์โหลดไฟล์นี้" ${r.blob ? '' : 'disabled'}>${IC.dl}</button><button class="ib tog" data-a="rx" data-k="${i}" aria-expanded="${open}" aria-label="ดูรายละเอียด ${esc(r.name)}" title="ดูรายละเอียด">${IC.chev}</button></div>
 ${open ? `<div class="issues">${o.bad.map(x => `<div class="iss"><b>${esc(x[1])}</b><span>${esc(x[0])}: ${esc(x[2])}</span></div>`).join('')}${o.inf.length ? `<p class="ish">สรุปการจัดแต่ละชีต</p>${o.inf.map(x => `<div class="iss inf"><b>${esc(x[0])}</b><span>${esc(x[2])}</span></div>`).join('')}` : ''}</div>` : ''}</div>`;
   })
   .join('')}</div>
 <p class="hint" style="margin-top:14px">ไฟล์ zip มีใบเช็ครูปทุกไฟล์ที่เติมได้ (ชื่อเดิม) และ รายงาน.txt ใบเช็คที่ได้วันนี้คือไฟล์ตั้งต้นของวันพรุ่งนี้</p>`;
}

/* ----- หน้าต่าง: วิธีใช้ / ต้อนรับ / วางรายชื่อ / ระบบทำงานอย่างไร ----- */
const vHelp =
  () => `<div class="ov"><div class="dlg" role="dialog" aria-modal="true" aria-labelledby="hlt"><div class="dlgh"><h2 id="hlt">วิธีใช้ประจำวัน</h2><button class="ib" data-a="help" aria-label="ปิด" data-focus>${IC.x}</button></div>
<ol class="hs">
<li><b>เตรียมไฟล์</b><span>Playlist ของวันนี้ 11 แบบ (สรุปรายการหนังแยกจอ 40Blue, 40Red, 80Blue, 80Red, CITY, Spare และ Signature Max, Show DC, BKJ1, BKJ2, Metro LED) กับใบเช็ครูปที่เติมแล้วของวันก่อนหน้า</span></li>
<li><b>ขั้นที่ 1 ใส่ Playlist</b><span>ถ้ามีงานของวันก่อนค้างอยู่ กด “ล้าง Playlist (เริ่มวันใหม่)” แล้วลากไฟล์ทั้งหมดลงกล่อง ดูว่าช่องในเช็กลิสต์สว่างครบ และ “Playlist วันที่” ถูกต้อง</span></li>
<li><b>ขั้นที่ 2 ตรวจ / แก้ไข (ไม่บังคับ)</b><span>ใช้เมื่อต้องปลด เพิ่ม หรือเปลี่ยนชื่อหนังก่อนเติม</span></li>
<li><b>ขั้นที่ 3 เติมใบเช็ครูป</b><span>ลากใบเช็ครูปลงกล่อง ดูว่าแต่ละไฟล์ขึ้น “พร้อมเติม” วันที่ทำงานและรอบถูกต้อง แล้วกด “เติมใบเช็ครูป”</span></li>
<li><b>ขั้นที่ 4 ผลการเติม</b><span>เปิดดูรายการ “ควรตรวจ” ของแต่ละไฟล์ แล้วกด “ดาวน์โหลดทั้งหมด (.zip)” ใบเช็คที่ได้คือไฟล์ตั้งต้นของวันพรุ่งนี้</span></li></ol>
<h3>ระบบจัดให้อย่างไร</h3><ul class="hr2">
<li>วันแรกของรอบ (วันที่ 1 และ 16) ลงหนังที่ออนอยู่ทั้งหมด หนัง # ขึ้นก่อน (ไฮไลต์เหลือง) แล้วเรียง A–Z</li>
<li>วันถัดไป หนังที่ยังไม่มีในใบเช็คและสถานะ “ขึ้นใหม่” จะลงใต้แถบวันที่ของวันนั้น หนังที่มีอยู่แล้วอยู่ที่เดิม ไม่ซ้ำ</li>
<li>“ออกวันที่ N” ใส่เมื่อถึงวันที่หนังออก (วันที่ 1 ยังไม่ใส่) ไฟล์เดิม / ต่อสัญญา ไม่ใส่ออกวันที่และไม่เพิ่มแถวซ้ำ</li>
<li>ตัวเลขที่ผู้ตรวจนับไว้ ย้ายตามชื่อหนังให้เอง</li>
<li>วันที่ 16 เริ่มรอบใหม่เป็นแถบสีเขียวต่อท้ายในชีตเดิม ไม่แตะรอบ 1–15 และแถบท้ายตารางเป็นวันถัดไปเสมอ</li></ul>
<h3>ข้อความ “ควรตรวจ” ที่พบบ่อย</h3><ul class="hr2">
<li><b>ไม่พบจอนี้ใน Playlist</b> ยังไม่ได้ใส่ Playlist ของสื่อนั้น หรือรหัสจอไม่ตรง ใส่ Playlist ให้ครบแล้วเติมใหม่</li>
<li><b>ใช้รายการของสื่ออื่น / ใช้รายการ THE CITY แทน</b> ระบบเติมจากรายการใกล้เคียงให้ เทียบกับ Playlist อีกครั้ง</li>
<li><b>รหัสนี้อยู่ในหลายสื่อ</b> ตั้งชื่อชีตในใบเช็คให้มี 40/80 + Blue/Red, CITY หรือ Spare แล้วเติมใหม่</li>
<li><b>ไม่ใส่ตามหมายเหตุ / จำนวนครั้ง = 0</b> บอกหนังที่ระบบตั้งใจไม่ใส่ ดูว่าตรงกับ Playlist</li>
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
 <div class="xl" style="margin-top:10px" role="img" aria-label="แถบวันที่สีเขียวของรอบ 16"><table><tbody>${row(8, '16-Oct-26', '', '', 'xb2')}</tbody></table></div><p class="fine">ตัวอย่างใบเช็คหลังเติมวันที่ 3 ชื่อหนังเป็นชื่อสมมติ</p></div>`;
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
 <li><span class="cn">5</span><span><b>ตัวเลขในคอลัมน์ “นับ”</b> ที่ผู้ตรวจนับไว้ ย้ายตามชื่อหนังให้เอง</span></li>
 <li><span class="cn">6</span><span><b>หนังใหม่ของแต่ละวัน</b> (สถานะ “ขึ้นใหม่”) ลงใต้แถบวันที่ของวันนั้น ไม่ซ้ำกับของเดิม</span></li>
 <li><span class="cn">7</span><span><b>แถบท้ายตารางเป็นวันถัดไป</b> ไว้ลงหนังของพรุ่งนี้</span></li>
 <li><span class="cn">8</span><span><b>วันที่ 16 เริ่มรอบใหม่</b> เป็นแถบสีเขียวต่อท้ายในชีตเดิม ไม่แตะรอบ 1–15</span></li></ol></div></section>
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
      p: 'ต้องเป็นวันที่จะเติม แถบสีบอกรอบ: น้ำเงิน = 1–15 เขียว = 16–สิ้นเดือน ถ้าไม่ตรง แก้ได้ในขั้นที่ 3',
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
      p: 'ปกติเว้นว่างไว้ ระบบใช้วันที่จากชื่อไฟล์ Playlist ถ้าพิมพ์เอง ระบบจะบอกวันที่ที่อ่านได้ใต้ช่อง ดูแถบสีว่ารอบถูกต้อง และเปิด “จัดแบบที่ทำด้วยมือ” ไว้เสมอ',
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
  clrx: () => {
    PLX = [];
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
    FILLING = true;
    try {
      const on = buildOn(o),
        ms = DB.find(s => s.m === 'MetroLED'),
        metro = ms ? listOf(ms, o) : [],
        out = new JSZip(),
        res = [],
        L = ['รายงานการเติมใบเช็ครูป'];
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
            r = await fillWb(f, on, metro, o);
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
        res.push({ name: f.name, rep: r.rep, blob: r.blob });
        L.push(
          '',
          f.name,
          ...r.rep.map(x => `  ${x[0]} | ${x[1]} | ${x[2]}${x[3] ? ' (' + x[3] + ')' : ''}`),
        );
      }
      out.file('รายงาน.txt', L.join('\n'));
      last = { res, zip: await out.generateAsync({ type: 'blob' }), saved: false };
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
