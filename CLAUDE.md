# CLAUDE.md — PHOTO CHECK STUDIO

อ่าน [`dev/README.md`](dev/README.md) ก่อนแก้อะไร: โครงไฟล์ ขั้นตอนออกเวอร์ชัน ชุดทดสอบ และแผนที่โค้ด

## สรุปสั้น

- ระบบทั้งหมดคือ `playlist_system.html` ไฟล์เดียว (โค้ดอยู่ใน `<script>` ที่มี `function layoutChk(`) • `playlist_system.js` สร้างด้วย `dev/build_js.py` ห้ามแก้ด้วยมือ
- ทุกเวอร์ชัน: แก้ HTML → `dev/check_syntax.py` → `dev/tests/run_tests.py` → `dev/set_version.py` → `dev/build_js.py --note` → commit / PR / merge → รอ GitHub Pages deploy
- ชุดทดสอบ (`photo-check-tests`) อยู่ใน Google Drive ของบริษัท — ขอจากเจ้าของระบบ ห้ามขึ้น repo

## กฎ

- ห้ามเอา Playlist / ใบเช็ครูป / ภาพหน้าจอผู้ใช้ ขึ้น GitHub หรือเว็บสาธารณะ
- ห้ามมีชื่อหนังโฆษณาใน README / PR / commit / คอมเมนต์ / โค้ดทดสอบใน repo
- เครื่องผู้ใช้อ่านอย่างเดียว เขียนลงโฟลเดอร์ผู้ใช้เฉพาะเมื่อขอ
- ระบบไม่ได้ใช้ AI ตอนทำงาน — พูดตามจริง
- คุยกับเจ้าของระบบเป็นภาษาไทยง่าย ๆ • เขาติดตั้งโปรแกรมไม่ได้ • ให้ดูภาพตัวอย่างก่อน และบอกเมื่อเสร็จ
