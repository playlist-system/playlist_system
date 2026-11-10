#!/usr/bin/env python3
"""ตรวจไวยากรณ์ JavaScript ของระบบ (เฉพาะ <script> ที่มี layoutChk) ด้วย node --check

    python3 dev/check_syntax.py playlist_system.html
"""
import sys, re, subprocess, tempfile, os

html = sys.argv[1] if len(sys.argv) > 1 else 'playlist_system.html'
s = open(html, encoding='utf-8').read()
code = None
for m in re.finditer(r'<script>(.*?)</script>', s, re.S):
    if 'layoutChk' in m.group(1):
        code = m.group(1)
if code is None:
    raise SystemExit('ไม่พบ <script> ที่มี layoutChk')
with tempfile.NamedTemporaryFile('w', suffix='.js', delete=False, encoding='utf-8') as t:
    t.write(code)
    tp = t.name
try:
    r = subprocess.run(['node', '--check', tp], capture_output=True, text=True)
finally:
    os.unlink(tp)
if r.returncode:
    print(r.stderr)
    raise SystemExit('ไวยากรณ์ผิด')
print('syntax ok', len(code))
