#!/usr/bin/env python3
"""ชุดทดสอบถดถอย (regression) ของ PHOTO CHECK STUDIO

    python3 dev/tests/run_tests.py playlist_system.html FIXTURES [--only oct27,fx] [--out DIR] [--update]

FIXTURES = โฟลเดอร์ชุดทดสอบ (แตกจาก zip ที่เก็บใน Google Drive ของบริษัท — ห้ามขึ้น GitHub
เพราะมี Playlist / ใบเช็ครูปจริง) ข้างในมี
    cases.json      รายการเคส (ดูด้านล่าง)
    expected/       ผลที่ถูกต้องของแต่ละเคส (สร้างด้วย --update)
    ...             ไฟล์ Playlist / ใบเช็ค / รายชื่อโฟลเดอร์ ตามที่ cases.json ชี้

cases.json
    {"fill":    [{"name": "oct27", "pl": "fill/oct27/pl", "ck": "fill/oct27/ck", "about": "…"}],
     "folders": [{"name": "fx", "ck": "folders/fx/ck", "dirs": "folders/fx/dirs.json",
                  "rounds": "cur", "upto": 0, "about": "…"}],
     "feedback": {"sheet": "fill/oct27/ck"}}

- fill    : ใส่ Playlist + ใบเช็ค → กดเติม → เทียบไฟล์ .xlsm ที่ได้ (ค่า สี ตัวหนา ขนาดตัวอักษร) และรายงาน
- folders : อ่านใบเช็คแบบหน้า "งานโฟลเดอร์และรูป" → แผน zip (ไม่มี / มีโฟลเดอร์เดิม) → แผนสร้างโฟลเดอร์
            บนโฟลเดอร์จำลองตาม dirs.json → สร้างจริงใน OPFS → เทียบผลทั้งหมดกับ expected/<name>.json
- dates   : อ่านวันที่แบบต่าง ๆ ด้วย parseD (ไม่ใช้ชุดทดสอบ รันทุกครั้ง) — รวมวันที่ที่พิมพ์ตก / (ตั้งแต่ 5.5.43)
- feedback: ปุ่มแจ้งปัญหาส่งแบบฟอร์มไปที่รับจำลอง (ที่อยู่เดียวกับที่ตั้งในระบบ) + จำลอง feedback.gs ด้วย node
            (ไม่ใช้ expected — ตรวจตามเงื่อนไข)

--update   เขียนผลรอบนี้ทับ expected/ (ใช้หลังตรวจแล้วว่าที่ต่างไปเป็นการแก้ที่ตั้งใจ)
ผลรอบนี้เก็บที่ --out (ค่าเริ่ม: FIXTURES/_out) • จบด้วย exit code 1 ถ้ามีเคสไม่ผ่าน
"""
import sys, os, json, glob, base64, argparse, shutil, subprocess, io, urllib.parse

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(os.path.dirname(HERE))
INIT = "try{localStorage.setItem('ps_welcomed','1');localStorage.setItem('ps_mascot','off')}catch(e){}"
NOW = '2026-11-10T10:00:00+07:00'   # นาฬิกาในเบราว์เซอร์ทดสอบ (cases.json "now" เปลี่ยนได้) — ใบเปล่าไม่มีวันที่ ระบบใช้วันนี้


def open_page(pw, html):
    import datetime
    br = pw.chromium.launch()
    ctx = br.new_context(viewport={'width': 1400, 'height': 1000})
    ctx.add_init_script(INIT)
    ctx.clock.set_fixed_time(datetime.datetime.fromisoformat(NOW))
    p = ctx.new_page()
    p.set_default_timeout(900000)
    errs = []
    p.on('pageerror', lambda e: errs.append(str(e)))
    p.on('dialog', lambda d: d.accept())
    data = open(html, 'rb').read()
    ctx.route('https://local.test/**', lambda r: r.fulfill(status=200, body=data, content_type='text/html; charset=utf-8'))
    p.goto('https://local.test/index.html')
    p.wait_for_function("typeof parsePlaylist === 'function'")
    return br, ctx, p, errs


def b64files(paths):
    return [[os.path.basename(f), base64.b64encode(open(f, 'rb').read()).decode()] for f in paths]


# ---------------------------------------------------------------- dates
# [ข้อความ, วันที่ที่ควรได้ (YYYY-MM-DD) หรือ None = ต้องอ่านไม่ได้]
DATES = [
    ('10/11/2026', '2026-11-10'), ('10/11/2569', '2026-11-10'), ('10-Nov-26', '2026-11-10'), ('2026-11-10', '2026-11-10'),
    # 5.5.43: พิมพ์ตกเครื่องหมาย / — อ่านได้เมื่อได้วันที่จริงแบบเดียว
    ('30/112026', '2026-11-30'), ('02/92026', '2026-09-02'), ('2710/2026', '2026-10-27'), ('30/112569', '2026-11-30'),
    ('1/12026', '2026-01-01'),
    ('3112/2026', '2026-12-31'), ('213/2026', '2026-03-21'),
    ('31/22026', None), ('111/2026', None), ('211/2026', None), ('30/111999', None), ('ไม่มีกำหนด', None), ('', None),   # 211/2026 = 21/1 หรือ 2/11 ไม่รู้ → ไม่เดา
]


def run_dates(pw, html):
    br, ctx, p, errs = open_page(pw, html)
    got = p.evaluate("L => L.map(t => { const d = parseD(t); return d ? d.toISOString().slice(0, 10) : null })", [t for t, _ in DATES])
    br.close()
    return ['“%s” ได้ %s (ควรได้ %s)' % (t, g, w) for (t, w), g in zip(DATES, got) if g != w], errs


# ---------------------------------------------------------------- fill
def run_fill(pw, html, fx, case, out):
    pl = sorted(glob.glob(os.path.join(fx, case['pl'], '*.xls*')) + glob.glob(os.path.join(fx, case['pl'], '*.zip')))
    ck = sorted(glob.glob(os.path.join(fx, case['ck'], '*.xlsm')))
    if not pl or not ck:
        raise SystemExit('เคส %s: ไม่พบไฟล์ Playlist หรือใบเช็ค' % case['name'])
    shutil.rmtree(out, ignore_errors=True)
    os.makedirs(out)
    br, ctx, p, errs = open_page(pw, html)
    p.evaluate("k => { $('#picker').dataset.k = k }", 'pl')
    b0 = p.evaluate('PSEQ')
    p.set_input_files('#picker', pl)
    p.wait_for_function('b => PSEQ > b && !BUSY', arg=b0)
    p.evaluate("k => { $('#picker').dataset.k = k }", 'ck')
    p.set_input_files('#picker', ck)
    p.wait_for_function('C.length===%d && !BUSY' % len(ck))
    p.wait_for_timeout(1500)
    p.evaluate("tab='fl'; view()")
    p.locator('[data-a="fill"]').first.click()
    p.wait_for_function("tab==='res' && last && last.res && last.res.length && !FILLING")
    for n, b in p.evaluate("last.res.filter(r=>r.blob).map(r=>{let s='';const b=r.blob;for(let i=0;i<b.length;i+=32768)"
                           "s+=String.fromCharCode.apply(null,b.subarray(i,i+32768));return [r.name,btoa(s)]})"):
        open(os.path.join(out, n), 'wb').write(base64.b64decode(b))
    rep = p.evaluate("last.res.map(r=>r.name+' | '+r.rep.map(x=>x.join(' / ')).join(' || ')).join('\\n')")
    open(os.path.join(out, '_report.txt'), 'w', encoding='utf-8').write(rep)
    br.close()
    return errs


def _col(x):
    """สีแบบอ่านง่าย: rgb / theme:<เลข><tint> / idx:<เลข>"""
    if x is None:
        return None
    t = getattr(x, 'type', None)
    if t == 'rgb':
        return x.rgb
    if t == 'theme':
        return 'theme:%s%+.2f' % (x.theme, x.tint or 0)
    if t == 'indexed':
        return 'idx:%s' % x.indexed
    return str(t)


def _fc(c):
    f = c.fill
    return _col(f.fgColor) if f and f.fill_type else None


def _style(c):
    return (_fc(c), c.font.b, _col(c.font.color), c.font.sz, c.font.name)


def cmp_xlsm(fa, fb, limit=15):
    import openpyxl
    wa = openpyxl.load_workbook(fa, keep_vba=False)
    wb = openpyxl.load_workbook(fb, keep_vba=False)
    diffs = []
    for s in sorted(set(wa.sheetnames) | set(wb.sheetnames)):
        if s not in wa.sheetnames or s not in wb.sheetnames:
            diffs.append(('ชีตมีแค่ฝั่งเดียว', s))
            continue
        sa, sb = wa[s], wb[s]
        R = min(max(sa.max_row, sb.max_row), 3000)
        C = min(max(sa.max_column, sb.max_column), 100)
        for r in range(1, R + 1):
            for c in range(1, C + 1):
                ca, cb = sa.cell(r, c), sb.cell(r, c)
                if ca.value != cb.value:
                    diffs.append((s, ca.coordinate, 'ค่า', ca.value, cb.value))
                else:
                    ka, kb = _style(ca), _style(cb)
                    if ka != kb:
                        diffs.append((s, ca.coordinate, 'รูปแบบ', ka, kb))
    return len(diffs), diffs[:limit]


def check_fill(exp, got):
    probs = []
    fe = sorted(os.path.basename(f) for f in glob.glob(exp + '/*.xlsm'))
    fg = sorted(os.path.basename(f) for f in glob.glob(got + '/*.xlsm'))
    if fe != fg:
        probs.append('ไฟล์ที่ได้ไม่ตรง: คาด %s ได้ %s' % (fe, fg))
    ra = open(os.path.join(exp, '_report.txt'), encoding='utf-8').read()
    rb = open(os.path.join(got, '_report.txt'), encoding='utf-8').read()
    if ra != rb:
        la, lb = ra.split('\n'), rb.split('\n')
        for i in range(max(len(la), len(lb))):
            a = la[i] if i < len(la) else ''
            b = lb[i] if i < len(lb) else ''
            if a != b:
                probs.append('รายงานต่าง บรรทัด %d\n      คาด: %s\n      ได้ : %s' % (i + 1, a[:300], b[:300]))
                break
    for n in sorted(set(fe) & set(fg)):
        k, ex = cmp_xlsm(os.path.join(exp, n), os.path.join(got, n))
        if k:
            probs.append('%s: ต่าง %d เซลล์' % (n, k) + ''.join('\n      %r' % (e,) for e in ex))
    return probs


# ---------------------------------------------------------------- folders
JS_FOLD = r"""async ([fs, dirs, rs, up]) => {
  TC = fs.map(([n, b]) => { const bin = atob(b), u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return {name: n, buf: u.buffer} });
  await tcRead(); const L = PICX; if (!L) return {none: 1, note: String(typeof note === 'undefined' ? '' : note)};
  const fmt = x => [x.pack, x.rnd, ...x.path].join('/') + '  [' + (x.sub || []).join(' | ') + ']';
  const zip = pzPlan(L.W, L.res, rs, up, null).list.map(fmt);
  const r0 = await navigator.storage.getDirectory(); try { await r0.removeEntry('rt', {recursive: true}) } catch (e) {}
  const root = await r0.getDirectoryHandle('rt', {create: true});
  for (const d of dirs) { let h = root; for (const n of d.split('/')) h = await h.getDirectoryHandle(n, {create: true}) }
  const X = await picScreens(root); X.dir = 'rt'; X.fx = await fxIdx(X.RF, L.W);
  const zx = pzPlan(L.W, L.res, rs, up, X).list.map(fmt);
  const P = await picPlan(root, L.W, L.res, rs, up);
  const plan = P.plan.map(q => q.disp + '  mk=' + q.mk.map(c => c[1]).join('>') + '  need=' + q.need.map(n => n[1]).join(' | ') + '  have=' + q.have);
  const mk = await picMake(P); const tree = [];
  const walk = async (h, p) => { const ks = []; for await (const [n, x] of h.entries()) if (x.kind === 'directory') ks.push([n, x]);
    ks.sort((a, b) => a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : 0); for (const [n, x] of ks) { const q = p ? p + '/' + n : n; tree.push(q); await walk(x, q) } };
  await walk(root, '');
  return JSON.parse(JSON.stringify({W: new Date(L.W).toISOString().slice(0, 10), sheets: L.res.map(r => [r.name, r.fold.length, !!r.blank]),
    zip, zipWithFolders: zx, plan, miss: P.miss.map(x => (x && x.id) || x), made: {n: mk.n, bad: mk.bad, c: mk.c}, tree}))
}"""


def run_folders(pw, html, fx, case):
    ck = sorted(glob.glob(os.path.join(fx, case['ck'], '*.xlsm')))
    dirs = json.load(open(os.path.join(fx, case['dirs']), encoding='utf-8')) if case.get('dirs') else []
    br, ctx, p, errs = open_page(pw, html)
    r = p.evaluate(JS_FOLD, [b64files(ck), dirs, case.get('rounds', 'cur'), case.get('upto', 0)])
    br.close()
    return r, errs


def check_json(exp, got):
    probs = []
    for k in sorted(set(exp) | set(got)):
        a, b = exp.get(k), got.get(k)
        if a == b:
            continue
        if isinstance(a, list) and isinstance(b, list):
            sa, sb = set(map(json.dumps, a)), set(map(json.dumps, b))
            lost = [json.loads(x) for x in sa - sb][:4]
            new = [json.loads(x) for x in sb - sa][:4]
            probs.append('%s: คาด %d รายการ ได้ %d • หายไป %s • เพิ่มมา %s%s' % (
                k, len(a), len(b), lost, new, '' if (lost or new) else ' • ลำดับต่าง'))
        else:
            probs.append('%s: คาด %r ได้ %r' % (k, a, b))
    return probs


# ---------------------------------------------------------------- feedback
def run_feedback(pw, html, fx, cfg):
    from PIL import Image
    tmp = os.path.join(fx, '_out', '_fb')
    os.makedirs(tmp, exist_ok=True)
    big, small = os.path.join(tmp, 'big.png'), os.path.join(tmp, 'small.png')
    Image.effect_noise((3000, 2000), 80).convert('RGB').save(big)
    Image.new('RGB', (40, 30), (20, 120, 200)).save(small)
    sheet = sorted(glob.glob(os.path.join(fx, cfg['sheet'], '*.xlsm')))[0]
    br, ctx, p, errs = open_page(pw, html)
    url = p.evaluate('fbUrl()')
    probs = []
    if not url:
        probs.append('ยังไม่ได้ตั้ง FB_URL ในระบบ')
        br.close()
        return probs, errs
    got = {}

    def h(route):
        got['body'] = route.request.post_data
        got['method'] = route.request.method
        page = ('<!doctype html><meta charset="utf-8"><h2>ส่งเรื่องแล้ว · เลขที่ #0007</h2><script>try{var w=window.top.opener;'
                'if(w)w.postMessage({psfb:1,ok:true,no:"#0007",err:""},"*")}catch(e){}</script>')
        route.fulfill(status=200, body=page, content_type='text/html; charset=utf-8')
    ctx.route(url, h)
    p.evaluate("k => { $('#picker').dataset.k = k }", 'ck')
    p.set_input_files('#picker', [sheet])
    p.wait_for_function('C.length===1 && !BUSY')
    if p.evaluate("document.body.innerText.includes('ยังไม่ได้ตั้งที่รับฟีดแบ็ก')"):
        probs.append('ขึ้นเตือนว่ายังไม่ได้ตั้งที่รับ')
    p.locator('[data-a="fb"]').first.click()
    p.wait_for_timeout(400)
    p.fill('#fb_nm', 'คุณทดสอบ')
    p.fill('#fb_wh', os.path.basename(sheet))
    p.locator('.fbtp label').nth(1).click()
    p.fill('#fb_tx', 'ทดสอบข้อความ ' * 300)
    p.set_input_files('#fb_file', [big, small])
    p.wait_for_function('FBI.length===2', timeout=60000)
    p.check('#fb_ckon')
    with ctx.expect_page() as pi:
        p.locator('[data-a="fbsend"]').click()
    pop = pi.value
    pop.wait_for_load_state()
    p.wait_for_function("FBS.includes('เลขที่')||FBS.includes('ส่งไม่ได้')", timeout=60000)
    if 'ส่งเรื่องแล้ว' not in pop.evaluate('document.body.innerText'):
        probs.append('แท็บผลการส่งไม่ขึ้น "ส่งเรื่องแล้ว"')
    if 'เลขที่ #0007' not in p.evaluate('FBS'):
        probs.append('หน้าระบบไม่ขึ้นเลขที่: ' + p.evaluate('FBS')[:120])
    try:
        b = json.loads(urllib.parse.parse_qs(got['body'])['p'][0])
        if got['method'] != 'POST': probs.append('ไม่ได้ส่งแบบ POST')
        if b.get('name') != 'คุณทดสอบ': probs.append('ชื่อผู้แจ้งไม่ตรง')
        if len(b.get('text', '')) < 3000: probs.append('ข้อความไม่ครบ')
        if len(b.get('images', [])) != 2: probs.append('รูปไม่ครบ 2 รูป')
        if (b.get('file') or {}).get('name') != os.path.basename(sheet): probs.append('ไม่ได้แนบไฟล์ใบเช็ค')
        if 'sys' not in b: probs.append('ไม่มีข้อมูลระบบ')
    except Exception as e:
        probs.append('อ่านข้อมูลที่ส่งไม่ได้: %r' % e)
    if p.evaluate("document.querySelectorAll('#fb_th figure').length") or p.evaluate("$('#fb_tx').value.length"):
        probs.append('ส่งแล้วฟอร์มไม่ล้าง')
    br.close()
    gs = os.path.join(REPO, 'apps_script', 'feedback.gs')
    r = subprocess.run(['node', os.path.join(HERE, 'feedback_gs_sim.js'), gs], capture_output=True, text=True)
    if r.returncode:
        probs.append('จำลอง feedback.gs ไม่ผ่าน:\n' + (r.stdout + r.stderr)[-800:])
    return probs, errs


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('html')
    ap.add_argument('fixtures')
    ap.add_argument('--only', default='')
    ap.add_argument('--out')
    ap.add_argument('--update', action='store_true')
    a = ap.parse_args()
    fx = os.path.abspath(a.fixtures)
    out = os.path.abspath(a.out or os.path.join(fx, '_out'))
    expd = os.path.join(fx, 'expected')
    cases = json.load(open(os.path.join(fx, 'cases.json'), encoding='utf-8'))
    global NOW
    NOW = cases.get('now', NOW)
    only = set(x for x in a.only.split(',') if x)
    want = lambda n: not only or n in only
    from playwright.sync_api import sync_playwright
    ver = open(a.html, encoding='utf-8').read().split("const VERTXT='", 1)[-1].split("'", 1)[0]
    print('ทดสอบ', ver)
    fails = []
    with sync_playwright() as pw:
        if want('dates'):
            probs, errs = run_dates(pw, a.html)
            probs += ['JS error: ' + x for x in errs]
            print(('ผ่าน ' if not probs else 'ไม่ผ่าน ') + 'dates — อ่านวันที่แบบต่าง ๆ (%d แบบ)' % len(DATES))
            for x in probs: print('    ' + x)
            if probs: fails.append('dates')
        for c in cases.get('fill', []):
            if not want(c['name']): continue
            o = os.path.join(out, c['name'])
            errs = run_fill(pw, a.html, fx, c, o)
            e = os.path.join(expd, c['name'])
            if a.update:
                shutil.rmtree(e, ignore_errors=True); shutil.copytree(o, e); probs = []
            else:
                probs = check_fill(e, o) if os.path.isdir(e) else ['ยังไม่มี expected (รัน --update ก่อน)']
            probs += ['JS error: ' + x for x in errs]
            print(('ผ่าน ' if not probs else 'ไม่ผ่าน ') + 'fill/' + c['name'], ('— ' + c['about']) if c.get('about') else '')
            for x in probs: print('    ' + x)
            if probs: fails.append(c['name'])
        for c in cases.get('folders', []):
            if not want(c['name']): continue
            r, errs = run_folders(pw, a.html, fx, c)
            os.makedirs(out, exist_ok=True)
            json.dump(r, open(os.path.join(out, c['name'] + '.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=0)
            e = os.path.join(expd, c['name'] + '.json')
            if a.update:
                os.makedirs(expd, exist_ok=True); shutil.copy(os.path.join(out, c['name'] + '.json'), e); probs = []
            else:
                probs = check_json(json.load(open(e, encoding='utf-8')), r) if os.path.exists(e) else ['ยังไม่มี expected (รัน --update ก่อน)']
            probs += ['JS error: ' + x for x in errs]
            print(('ผ่าน ' if not probs else 'ไม่ผ่าน ') + 'folders/' + c['name'], ('— ' + c['about']) if c.get('about') else '')
            for x in probs: print('    ' + x)
            if probs: fails.append(c['name'])
        if cases.get('feedback') and want('feedback'):
            probs, errs = run_feedback(pw, a.html, fx, cases['feedback'])
            probs += ['JS error: ' + x for x in errs]
            print(('ผ่าน ' if not probs else 'ไม่ผ่าน ') + 'feedback — ปุ่มแจ้งปัญหา + feedback.gs')
            for x in probs: print('    ' + x)
            if probs: fails.append('feedback')
    print('สรุป:', 'ผ่านทุกเคส' if not fails else 'ไม่ผ่าน ' + ', '.join(fails), '(เขียน expected ใหม่แล้ว)' if a.update else '')
    sys.exit(1 if fails else 0)


if __name__ == '__main__':
    main()
