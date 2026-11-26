#!/usr/bin/env python3
"""ตรวจหนังตกหล่นทั้งวัน — เทียบ Playlist ของวันหนึ่งกับใบเช็ครูปทุกจอ ด้วยกติกาของระบบเอง

    python3 dev/day_audit.py --pl <โฟลเดอร์ Playlist ของวันนั้น> --ck <โฟลเดอร์ใบเช็ค .xlsm>
                             [--now 2026-11-26T09:00] [--nost] [--out <โฟลเดอร์ผล>]

ทำอะไร
  1. เติมด้วย Playlist + ใบเช็คที่ให้ (สำเนา playlist_system.html ที่จดรายการหนังของแต่ละจอไว้) → ผลอยู่ใน --out
  2. แต่ละจอในใบเช็ค: หนังใน Playlist (ไม่นับรายการปลด / ชื่อจอ / Programmatic / FastAd / รอไฟล์)
     ที่ไม่มีในคอลัมน์นั้นของใบเช็คที่เติมแล้ว → บอกเหตุผลตามกติกาของระบบ
     (หมายเหตุ / จำนวนครั้ง 0 หรือว่าง / ข้ามด้วยเหตุอื่น / "ขาด" = ระบบควรลงแต่ไม่มี → บั๊ก)
  3. สรุปจากรายงานการเติม: จอที่ไม่พบใน Playlist, กันตกหล่น, มาช้า
  ผลเต็มเขียนที่ <out>/audit.json • หน้าจอพิมพ์เฉพาะสรุป

--nost  = เปิด "เพิ่มหนังกันตกหล่นให้เอง" (ให้ตรงกับรอบที่ผู้ใช้เติม — ดูรายงาน: "เพิ่มให้แล้ว" = เปิด, "ระบบจึงไม่เพิ่ม" = ปิด)
ถ้าใบเช็คที่ให้เป็นใบที่เติมแล้ว ผลที่เติมซ้ำต้องเหมือนเดิม (ใช้ dev/tests/run_tests.py cmp_xlsm ตรวจได้)
ไฟล์ Playlist / ใบเช็คของผู้ใช้ห้ามขึ้น repo — เก็บไว้ใน scratchpad
"""
import sys, os, re, json, glob, argparse, tempfile, collections, datetime
HERE = os.path.dirname(os.path.abspath(__file__)); REPO = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(HERE, 'tests')); import run_tests as R

HOOK = " for(const j of jobs){if(!j[3]||!j[3].rows)continue;const nq=noteQ("
CAP = (" window.__JOBS=(window.__JOBS||[]).concat(jobs.map(j=>({f:f.name,sn:j[0],col:j[1],m:j[3]&&j[3].m,c:j[3]&&j[3].c,k:j[4]})));\n")
FLAGS = r'''()=>{const keys=new Set((window.__JOBS||[]).map(j=>j.m+'|'+j.c));const O={};
  for(const s of DB){const k=s.m+'|'+s.c;if(!keys.has(k))continue;
    O[k]=(s.rows||[]).filter(r=>!r.off&&r.ad).map(r=>{const a=String(r.ad).trim();let ns=false,sk=false;try{ns=!!noteSkip(s,r)}catch(e){}try{sk=!!skipRow(s,r)}catch(e){}
      return{ad:a,cnt:r.cnt==null?null:r.cnt,note:String(r.note||'').slice(0,90),s:r.s||'',st:r.st||'',junk:JUNK.test(a)||a===s.c,mtp:MTP.test(a),fast:FASTN.test(r.ad),wf:/เมื่อมีไฟล์/.test(r.ad),ns,sk}})}
  return O}'''


def nz(t):
    return re.sub(r'^bkj[12]_', '', re.sub(r'\s+', '', str(t or '')).lower().lstrip('#'))


def sheet_names(path, sn, col, cache):
    import openpyxl
    if path not in cache:
        cache[path] = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = cache[path][sn]; ci = openpyxl.utils.column_index_from_string(col)
    out = set()
    for row in ws.iter_rows(min_row=15, max_row=700, min_col=ci, max_col=ci, values_only=True):
        v = row[0]
        if isinstance(v, str) and v.strip():
            out.add(nz(v))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--pl', required=True); ap.add_argument('--ck', required=True)
    ap.add_argument('--now', default=None); ap.add_argument('--nost', action='store_true')
    ap.add_argument('--out', default=None); ap.add_argument('--html', default=os.path.join(REPO, 'playlist_system.html'))
    a = ap.parse_args()
    out = a.out or tempfile.mkdtemp(prefix='day_audit_')
    src = open(a.html, encoding='utf-8').read()
    assert src.count(HOOK) == 1, 'หาจุดจดรายการจอใน fillWb ไม่เจอ (โค้ดเปลี่ยน) — แก้ HOOK ใน day_audit.py'
    dbg = os.path.join(out, '_studio_dbg.html'); os.makedirs(out, exist_ok=True)
    open(dbg, 'w', encoding='utf-8').write(src.replace(HOOK, CAP + HOOK))
    if a.now:
        R.NOW = a.now if '+' in a.now else a.now + ':00+07:00'
    if a.nost:
        R.INIT = R.INIT.replace('}catch', ";localStorage.setItem('ps_nost','1')}catch")
    plf = os.path.abspath(a.pl); ckf = os.path.abspath(a.ck)
    got = {}
    orig = R.open_page
    def open_page(pw, h):
        br, ctx, p, errs = orig(pw, h); oc = br.close
        def close():
            got['jobs'] = p.evaluate('window.__JOBS||[]'); got['flags'] = p.evaluate(FLAGS); got['errs'] = list(errs); oc()
        br.close = close
        return br, ctx, p, errs
    R.open_page = open_page
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        R.run_fill(pw, dbg, '/', {'name': 'audit', 'pl': plf.lstrip('/'), 'ck': ckf.lstrip('/')}, os.path.join(out, 'filled'))
    jobs, flags = got['jobs'], got['flags']
    cache, seen, res = {}, set(), []
    for j in jobs:
        key = (j['f'], j['sn'], j['col'])
        if key in seen or not j.get('c'): continue
        seen.add(key)
        have = sheet_names(os.path.join(out, 'filled', j['f']), j['sn'], j['col'], cache)
        for r in flags.get(j['m'] + '|' + j['c'], []):
            if r['junk'] or r['mtp'] or r['fast'] or r['wf'] or nz(r['ad']) in have: continue
            why = 'หมายเหตุ' if r['ns'] else ('จำนวนครั้ง 0/ว่าง' if r['sk'] and not r['cnt'] else ('ข้าม(อื่น)' if r['sk'] else 'ขาด'))
            res.append({'file': j['f'], 'sheet': j['sn'], 'col': j['col'], 'screen': j['m'] + '|' + j['c'], 'ad': r['ad'], 'cnt': r['cnt'], 'start': r['s'], 'status': r['st'], 'note': r['note'], 'why': why})
    rep = open(os.path.join(out, 'filled', '_report.txt'), encoding='utf-8').read()
    lines = [x for ln in rep.splitlines() for x in ln.split(' || ')]
    pick = lambda pat: sorted({re.sub(r'\s*/\s*\d+$', '', x)[:220] for x in lines if re.search(pat, x) and ' / รูปแบบ / ' not in x})
    summ = {'jobs': len(seen), 'by_reason': dict(collections.Counter(x['why'] for x in res)),
            'not_found': pick(r'ไม่พบจอ'), 'nost': pick(r'กันตกหล่น'), 'late': pick(r'มาช้า'), 'errors': got['errs'][:5]}
    json.dump({'summary': summ, 'missing': res}, open(os.path.join(out, 'audit.json'), 'w'), ensure_ascii=False, indent=1)
    print('จอที่ตรวจ', summ['jobs'], '| ผล', summ['by_reason'] or 'ครบทุกจอ', '| ผลเต็ม', os.path.join(out, 'audit.json'))
    for x in res:
        if x['why'] in ('ขาด', 'ข้าม(อื่น)'):
            print('  [%s] %s / %s / %s → %s (ครั้ง %s) %s' % (x['why'], x['file'], x['sheet'], x['screen'], x['ad'], x['cnt'], x['note']))
    for t, L in (('ไม่พบจอใน Playlist', summ['not_found']), ('กันตกหล่น', summ['nost']), ('มาช้า', summ['late'])):
        if L:
            print(t + ':'); [print('  ' + x) for x in L[:20]]
    if summ['errors']: print('pageerror:', summ['errors'])


if __name__ == '__main__':
    main()
