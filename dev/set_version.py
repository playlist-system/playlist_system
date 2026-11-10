#!/usr/bin/env python3
"""เปลี่ยนเลขเวอร์ชันทุกที่ในครั้งเดียว (รันจากโฟลเดอร์ repo)

    python3 dev/set_version.py 5.5.37 "11 พ.ย. 2569" [--history "ข้อความประวัติใน README"]

แก้ 3 ไฟล์:
- playlist_system.html  : VERTXT = '… เวอร์ชัน X (วันที่)'
- README.md             : **เวอร์ชัน X · วันที่**  (+ บรรทัดแรกใต้ "## ประวัติเวอร์ชัน" ถ้าใส่ --history)
- คู่มือ…txt             : เวอร์ชัน X  •  ปรับปรุงล่าสุด วันที่
ส่วนหัวของ playlist_system.js ใช้ dev/build_js.py --note
"""
import re, sys, os, argparse
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docs_edit import edit, MANUAL

ap = argparse.ArgumentParser()
ap.add_argument('ver'); ap.add_argument('date')
ap.add_argument('--history', help='ข้อความ (markdown) ต่อจาก "— " ในประวัติเวอร์ชัน')
a = ap.parse_args()


def sub1(path, pat, rep, crlf=False):
    raw = open(path, 'rb').read()
    bom = raw.startswith(b'\xef\xbb\xbf')
    t = raw.decode('utf-8-sig')
    t2, n = re.subn(pat, rep, t, count=1)
    if n != 1:
        raise SystemExit('%s: ไม่พบ %s' % (path, pat))
    open(path, 'wb').write((b'\xef\xbb\xbf' if bom else b'') + t2.encode('utf-8'))
    return re.search(pat, t).group(0)


old = sub1('playlist_system.html', r"เวอร์ชัน [0-9.]+ \([^)]*\)';", "เวอร์ชัน %s (%s)';" % (a.ver, a.date))
sub1('README.md', r'\*\*เวอร์ชัน [0-9.]+ · [^*]*\*\*', '**เวอร์ชัน %s · %s**' % (a.ver, a.date))
sub1(MANUAL, r'เวอร์ชัน [0-9.]+  •  ปรับปรุงล่าสุด [^\r\n]*', 'เวอร์ชัน %s  •  ปรับปรุงล่าสุด %s' % (a.ver, a.date))
if a.history:
    edit('README.md', [('## ประวัติเวอร์ชัน\n\n',
                        '## ประวัติเวอร์ชัน\n\n- **%s (%s)** — %s\n' % (a.ver, a.date, a.history.strip()))])
print('จาก', old, '→', a.ver, a.date)
