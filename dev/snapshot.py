#!/usr/bin/env python3
"""เก็บสำเนาเวอร์ชันที่กำลังใช้อยู่ไว้ใน versions/ (ใช้ตอนเวอร์ชันล่าสุดมีปัญหา) — รันจากโฟลเดอร์ repo

    python3 dev/snapshot.py [--rev HEAD] [--keep 5]

- ดึง playlist_system.html จาก git (ค่าเริ่ม HEAD = เวอร์ชันที่อยู่บนเว็บตอนนี้) → versions/<เลขเวอร์ชัน>.html
- เก็บไว้ --keep เวอร์ชันล่าสุด (ค่าเริ่ม 5) ที่เก่ากว่านั้นลบออก (ยังอยู่ในประวัติ git)
- สร้าง versions/index.html ใหม่ = หน้าเลือกเวอร์ชัน (ปุ่ม "ใช้เวอร์ชันก่อนหน้า" ท้ายหน้าระบบลิงก์มาที่นี่)
- เวอร์ชันตั้งแต่ 5.5.37 รู้ตัวเองว่าถูกเปิดจาก versions/ แล้วขึ้นแถบเตือน • เวอร์ชันก่อนหน้านั้นสคริปต์นี้ใส่แถบให้
dev/set_version.py เรียกสคริปต์นี้ให้อัตโนมัติก่อนเปลี่ยนเลขเวอร์ชัน
"""
import argparse, glob, html, os, re, subprocess, sys

SITE = 'https://playlist-system.github.io/playlist_system/'
OUT = 'versions'


def vtuple(v):
    return tuple(int(x) for x in v.split('.') if x.isdigit())


def ver_of(h):
    m = re.search(r"const VERTXT='[^']*?เวอร์ชัน ([0-9.]+) \(([^)]*)\)'", h)
    return (m.group(1), m.group(2)) if m else (None, '')


def banner(h, v):
    """ใส่แถบ "กำลังใช้เวอร์ชันก่อนหน้า" ให้เวอร์ชันที่ยังไม่มีโค้ด OLDVER (ก่อน 5.5.37)"""
    if 'OLDVER' in h:
        return h
    a = '</style></head><body>'
    if h.count(a) != 1:
        print('  (ใส่แถบเตือนไม่ได้ — ไม่พบตำแหน่ง <body>)')
        return h
    st = 'style="color:#7A4A00;font-weight:700;text-decoration:underline;margin:0 6px;white-space:nowrap"'
    b = ('<div role="status" style="background:#FFF4D6;color:#5A3B00;border-bottom:1px solid #E9B949;padding:10px 16px;'
         'font:14px/1.6 Tahoma,sans-serif;text-align:center"><b>กำลังใช้เวอร์ชันก่อนหน้า (%s)</b> — ใช้ชั่วคราวระหว่างรอแก้เวอร์ชันล่าสุด '
         '<a %s href="%splaylist_system.html">กลับไปเวอร์ชันล่าสุด</a><a %s href="%sversions/">เลือกเวอร์ชันอื่น</a></div>') % (
        html.escape(v), st, SITE, st, SITE)
    return h.replace(a, a + b)


def summary(v):
    """หัวข้อของเวอร์ชันจากประวัติใน README (ข้อความตัวหนาตัวแรกหลัง —)"""
    try:
        t = open('README.md', encoding='utf-8').read()
    except OSError:
        return ''
    m = re.search(r'^- \*\*' + re.escape(v) + r' \([^)]*\)\*\* — (.*)$', t, re.M)
    if not m:
        return ''
    b = re.search(r'\*\*(.+?)\*\*', m.group(1))
    s = b.group(1) if b else m.group(1)
    s = re.sub(r'[`*]', '', s)
    return s[:140]


PAGE = '''<!doctype html>
<html lang="th">
<head>
<meta charset="utf-8">
<meta name="robots" content="noindex,nofollow">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>เวอร์ชันก่อนหน้า — PHOTO CHECK STUDIO</title>
<style>
:root{--bg:#EEF3FB;--card:#FFFFFF;--tx:#0F1B33;--mu:#46526B;--ln:#D5DBE5;--ac:#1F5FE0;--acink:#FFFFFF;--wbg:#FFF4D6;--wtx:#5A3B00;--wln:#E9B949}
@media (prefers-color-scheme:dark){:root{--bg:#0B1430;--card:#121D3F;--tx:#E9EEF8;--mu:#A9B6D3;--ln:#24365A;--ac:#7AA7FF;--acink:#0B1430;--wbg:#3A2C08;--wtx:#FBE3A6;--wln:#8A6A1C}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--tx);font:16px/1.65 "Leelawadee UI","Noto Sans Thai",Tahoma,sans-serif}
.w{max-width:720px;margin:0 auto;padding:32px 16px 48px}
h1{font-size:26px;line-height:1.3;margin:0 0 4px}
.sub{color:var(--mu);margin:0 0 20px}
.note{background:var(--wbg);color:var(--wtx);border:1px solid var(--wln);border-radius:12px;padding:12px 16px;margin:0 0 20px}
.note ol{margin:0;padding-left:20px}
ol.v{list-style:none;padding:0;margin:0;display:grid;gap:12px}
.v li{background:var(--card);border:1px solid var(--ln);border-radius:14px;padding:14px 16px;display:flex;gap:12px 16px;align-items:center;flex-wrap:wrap}
.v .i{flex:1 1 240px;min-width:0}
.v b{font-size:18px}
.v .d{color:var(--mu);font-size:14px}
.v .s{font-size:14px;margin-top:2px;overflow-wrap:anywhere}
.tag{display:inline-block;font-size:12px;font-weight:700;color:var(--ac);border:1px solid var(--ac);border-radius:99px;padding:0 8px;margin-left:8px;vertical-align:3px}
.btn{display:inline-block;background:var(--ac);color:var(--acink);text-decoration:none;font-weight:700;padding:10px 18px;border-radius:10px;white-space:nowrap}
.btn:focus-visible{outline:3px solid var(--ac);outline-offset:2px}
.btn.o{background:transparent;color:var(--ac);border:1.5px solid var(--ac)}
.back{margin-top:24px}
.empty{color:var(--mu)}
</style>
</head>
<body>
<div class="w">
<h1>เวอร์ชันก่อนหน้า</h1>
<p class="sub">PHOTO CHECK STUDIO — ใช้เมื่อเวอร์ชันล่าสุดมีปัญหา เช่น เปิดไม่ขึ้น หรือเติมใบเช็คผิด</p>
<div class="note"><ol>
<li>เลือกเวอร์ชันบนสุดก่อน ถ้ายังมีปัญหา ค่อยเลือกอันถัดไป</li>
<li>แจ้งคนดูแลระบบด้วย (ปุ่ม “แจ้งปัญหา” หรือ LINE)</li>
<li>แก้เสร็จแล้ว กลับไปใช้เวอร์ชันล่าสุดตามปกติ</li>
</ol></div>
%s
<p class="back"><a class="btn o" href="../playlist_system.html">กลับไปเวอร์ชันล่าสุด</a></p>
</div>
</body>
</html>
'''


def index():
    items = []
    for f in glob.glob(os.path.join(OUT, '*.html')):
        n = os.path.basename(f)
        if n == 'index.html':
            continue
        v, d = ver_of(open(f, encoding='utf-8').read())
        if v:
            items.append((vtuple(v), v, d, n))
    items.sort(reverse=True)
    if items:
        li = []
        for k, (_, v, d, n) in enumerate(items):
            s = summary(v)
            li.append('<li><div class="i"><b>%s</b>%s<div class="d">%s</div>%s</div><a class="btn" href="%s">ใช้เวอร์ชันนี้</a></li>' % (
                html.escape(v), '<span class="tag">ก่อนหน้าล่าสุด</span>' if k == 0 else '', html.escape(d),
                '<div class="s">%s</div>' % html.escape(s) if s else '', html.escape(n)))
        body = '<ol class="v">\n' + '\n'.join(li) + '\n</ol>'
    else:
        body = '<p class="empty">ยังไม่มีเวอร์ชันก่อนหน้า</p>'
    open(os.path.join(OUT, 'index.html'), 'w', encoding='utf-8').write(PAGE % body)
    return [x[1] for x in items]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--rev', default='HEAD')
    ap.add_argument('--keep', type=int, default=5)
    a = ap.parse_args()
    h = subprocess.run(['git', 'show', a.rev + ':playlist_system.html'], capture_output=True, check=True).stdout.decode('utf-8')
    v, d = ver_of(h)
    if not v:
        raise SystemExit('ไม่พบ VERTXT ใน %s' % a.rev)
    os.makedirs(OUT, exist_ok=True)
    open(os.path.join(OUT, v + '.html'), 'w', encoding='utf-8').write(banner(h, v))
    have = sorted((vtuple(ver_of(open(f, encoding='utf-8').read())[0] or '0'), f)
                  for f in glob.glob(os.path.join(OUT, '*.html')) if not f.endswith('index.html'))
    for _, f in have[:-a.keep] if a.keep > 0 else []:
        os.remove(f)
        print('  ลบสำเนาเก่า', f)
    print('เก็บ', v, '(%s)' % d, '→ versions/ • มี:', ', '.join(index()))


if __name__ == '__main__':
    main()
