#!/usr/bin/env python3
"""สร้าง playlist_system.js (สำเนาโค้ดที่อ่านง่าย) จาก playlist_system.html

    python3 dev/build_js.py playlist_system.html OLD_JS OUT_JS [--note "5.5.37: ข้อความสั้น"]

- โค้ดของระบบอยู่ใน <script> ที่มี function layoutChk( — ดึงส่วนนั้นมาจัดย่อหน้าด้วย prettier
- หัวไฟล์ (คอมเมนต์ /* ... */ ด้านบน) เอามาจาก OLD_JS (มักเป็น `git show HEAD:playlist_system.js`)
- --note เพิ่มบรรทัด " *   เวอร์ชัน <note>" ต่อจากบรรทัดเวอร์ชันล่าสุดในหัวไฟล์
- ห้ามแก้ playlist_system.js ด้วยมือ ให้แก้ playlist_system.html แล้วสร้างใหม่ด้วยสคริปต์นี้
- ต้องมี prettier (npm i -g prettier@3.8.1)
"""
import sys, subprocess, tempfile, os, argparse

ap = argparse.ArgumentParser()
ap.add_argument('html'); ap.add_argument('old_js'); ap.add_argument('out_js')
ap.add_argument('--note', help='เช่น "5.5.37: แก้ …" (ไม่ต้องใส่คำว่า เวอร์ชัน)')
a = ap.parse_args()

h = open(a.html, encoding='utf-8').read()
i = h.find('function layoutChk(')
if i < 0:
    raise SystemExit('ไม่พบ function layoutChk( ใน ' + a.html)
s = h.rfind('<script', 0, i)
js = h[h.find('>', s) + 1:h.find('</script>', i)]
old = open(a.old_js, encoding='utf-8').read()
end = old.find(' */\n') + 4
hdr = old[:end]
if a.note:
    mark = ' *   เวอร์ชัน '
    k = hdr.rfind(mark)
    if k < 0:
        raise SystemExit('ไม่พบบรรทัด " *   เวอร์ชัน " ในหัวไฟล์')
    e = hdr.find('\n', k) + 1
    hdr = hdr[:e] + mark + a.note.strip() + '\n' + hdr[e:]
with tempfile.NamedTemporaryFile('w', suffix='.js', delete=False, encoding='utf-8') as t:
    t.write(js)
    tp = t.name
try:
    pr = subprocess.run(['prettier', '--single-quote', '--print-width', '110', '--arrow-parens', 'avoid', tp],
                        capture_output=True, text=True, check=True).stdout
finally:
    os.unlink(tp)
open(a.out_js, 'w', encoding='utf-8').write(hdr + '\n' + pr)
print('ok', len(hdr), len(pr))
