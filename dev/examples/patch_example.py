#!/usr/bin/env python3
"""ตัวอย่างสคริปต์แก้โค้ด (จากเวอร์ชัน 5.5.36 ที่ตั้งที่รับฟีดแบ็ก)

    python3 patch.py playlist_system.html playlist_system.html

วิธีที่ใช้ทุกเวอร์ชัน: แทนข้อความเดิม → ข้อความใหม่ และ assert ว่าข้อความเดิมพบตามจำนวนที่คาด
ถ้าโค้ดเปลี่ยนไปจนหาไม่เจอ สคริปต์หยุดทันที ไม่แก้ผิดที่
- คอมเมนต์ในโค้ดขึ้นต้นด้วยเลขเวอร์ชัน เช่น /* 5.5.36: … */ ให้ตามหาได้ว่าเปลี่ยนเมื่อไร
- เลขเวอร์ชัน (VERTXT) ใช้ dev/set_version.py แทนการแก้ในสคริปต์นี้
- ห้ามมีชื่อหนังโฆษณาในคอมเมนต์ / ข้อความ
"""
import sys
SRC, DST = sys.argv[1], sys.argv[2]
s = open(SRC, encoding='utf-8').read()


def rep(a, b, cnt=1):
    global s
    n = s.count(a)
    assert n == cnt, (n, cnt, a[:160])
    s = s.replace(a, b)


rep("const FB_URL='';",
    "/* 5.5.36: ที่รับของบริษัท (เปิดได้เฉพาะบัญชี Google ของบริษัท) */\n"
    "const FB_URL='https://script.google.com/a/macros/<โดเมนบริษัท>/s/<รหัส>/exec';")
open(DST, 'w', encoding='utf-8').write(s)
print('ok', len(s))
