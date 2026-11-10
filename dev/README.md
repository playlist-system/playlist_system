# คู่มือคนดูแลระบบ (สำหรับแก้และออกเวอร์ชันใหม่)

เอกสารนี้สำหรับคนที่แก้โค้ด (หรือ Claude ในแชตใหม่) — ผู้ใช้ทั่วไปอ่าน [README หลัก](../README.md) และคู่มือ .txt

## 1. ระบบประกอบด้วยอะไร

| ไฟล์ | คืออะไร |
|---|---|
| `playlist_system.html` | **ตัวระบบทั้งหมด** ไฟล์เดียว (ไลบรารี SheetJS / JSZip ฝังอยู่ข้างใน) — โค้ดของระบบอยู่ใน `<script>` ที่มี `function layoutChk(` |
| `playlist_system.js` | สำเนาโค้ดที่จัดย่อหน้าให้อ่านง่าย **สร้างจาก HTML ด้วย `dev/build_js.py` ห้ามแก้ด้วยมือ** |
| `index.html` | หน้าเปิด → ส่งต่อไป `playlist_system.html` |
| `versions/` | สำเนา 5 เวอร์ชันก่อนหน้า + `index.html` หน้าเลือกเวอร์ชัน (ลิงก์ “ใช้เวอร์ชันก่อนหน้า” ท้ายหน้าระบบ) — `dev/set_version.py` เก็บให้อัตโนมัติ ห้ามแก้ด้วยมือ |
| `README.md`, `คู่มือการใช้งาน_ระบบเติมใบเช็ครูป.txt` | เอกสารผู้ใช้ — อัปเดตทุกเวอร์ชัน (คู่มือ .txt เป็น CRLF + BOM) |
| `apps_script/Code.gs` | Apps Script เก็บ Playlist จากอีเมลเข้า Google Drive + หน้าเว็บ "Playlist รายวัน" |
| `apps_script/feedback.gs` | ที่รับปุ่ม "แจ้งปัญหา" (Web app ของบริษัท เปิดได้เฉพาะบัญชีบริษัท) — วิธีตั้งดู `apps_script/FEEDBACK.md` |
| `n8n/` | ทางเลือกเก็บ Playlist ด้วย n8n (ไม่มีค่าตั้งส่วนตัว) |
| `dev/` | เครื่องมือของคนดูแลระบบ (ไฟล์นี้) |

- เว็บจริง: https://playlist-system.github.io/playlist_system/ — GitHub Pages deploy จาก `main` อัตโนมัติเมื่อ merge (ประมาณ 1 นาที)
- ระบบทำงานในเบราว์เซอร์ของผู้ใช้ทั้งหมด ไฟล์ไม่ถูกส่งออก **ตอนใช้งานไม่ได้ใช้ AI** (AI ช่วยเขียนโค้ดเท่านั้น) — พูดตามจริงเสมอ
- ปุ่มแจ้งปัญหาส่งแบบฟอร์มในแท็บใหม่ไปที่ `FB_URL` (ในโค้ด) — ผู้ใช้ต้องล็อกอินบัญชี Google ของบริษัท

## 2. กฎที่ต้องทำตามเสมอ

1. **ห้ามเอา Playlist / ใบเช็ครูป / ภาพหน้าจอจากผู้ใช้ ขึ้น GitHub หรือเว็บสาธารณะ** (ข้อมูลแคมเปญของลูกค้า) — ชุดทดสอบเก็บใน Google Drive ของบริษัทเท่านั้น
2. **ห้ามใส่ชื่อหนังโฆษณา** ใน README / PR / commit / คอมเมนต์ในโค้ด / โค้ดทดสอบใน repo — ภาพตัวอย่างใช้ข้อมูลปลอม
3. ไฟล์ที่แนบมากับฟีดแบ็กอยู่ใน Google Drive ของบริษัท ห้ามย้ายไป GitHub Issues
4. ค่าตั้ง n8n ส่วนตัวไม่ขึ้น GitHub • อีเมลผู้ใช้ใช้แค่ระบุผู้เขียน
5. เครื่องของผู้ใช้: **อ่านอย่างเดียว** ห้ามเขียนลงโฟลเดอร์ของผู้ใช้ (รวมถึงไดรฟ์ N: Gallary) ถ้าไม่ได้ขอ
6. เจ้าของระบบ: อธิบายเป็น **ภาษาไทยง่าย ๆ** • ติดตั้งโปรแกรมในเครื่องไม่ได้ • อยากเห็นภาพตัวอย่างก่อน และให้บอกเมื่อเสร็จ

## 3. เตรียมเครื่อง (ครั้งเดียวต่อแชต / เครื่อง)

```bash
npm i -g prettier@3.8.1                      # dev/build_js.py ใช้
pip install openpyxl playwright pillow        # ชุดทดสอบ (+ playwright install chromium ถ้ายังไม่มี)
node --version                                # ใช้ node 22
```

## 4. ออกเวอร์ชันใหม่ (ทำทุกครั้ง)

ตัวอย่างจริงอยู่ใน [`dev/examples/`](examples/) — รันจากโฟลเดอร์ repo

```bash
git checkout main && git pull
git checkout -b <ชื่อสั้น>-5-5-37

# 1) แก้ playlist_system.html — ใช้สคริปต์แทนข้อความที่ assert จำนวนครั้ง (ดู examples/patch_example.py)
python3 /path/patch.py playlist_system.html playlist_system.html

# 2) ตรวจไวยากรณ์
python3 dev/check_syntax.py playlist_system.html

# 3) ทดสอบ (ดูหัวข้อ 5) — ต่างจาก expected ได้เฉพาะเรื่องที่ตั้งใจแก้
python3 dev/tests/run_tests.py playlist_system.html /path/photo-check-tests

# 4) เลขเวอร์ชัน + ประวัติ (html, README, คู่มือ) — เก็บสำเนาเวอร์ชันบนเว็บตอนนี้ (git HEAD) ไว้ใน versions/ ให้ก่อน
python3 dev/set_version.py 5.5.37 "11 พ.ย. 2569" --history "**หัวข้อ** • รายละเอียดสั้น ๆ"
#    ถ้าต้องแก้เนื้อหาคู่มือ/README เพิ่ม ใช้ dev/docs_edit.py (ดู examples/docs_example.py)

# 5) สร้าง playlist_system.js ใหม่
git show HEAD:playlist_system.js > /tmp/old.js
python3 dev/build_js.py playlist_system.html /tmp/old.js playlist_system.js --note "5.5.37: ข้อความสั้น"
node --check playlist_system.js

# 6) commit → push → PR → merge → รอ deploy
git add -A && git commit -m "5.5.37: …"
git push -u origin HEAD
gh pr create … && gh pr merge …   # หรือผ่าน gh api
```

- ข้อความ commit / PR: บอกว่าแก้อะไร ทดสอบอะไร — **ไม่มีชื่อหนัง**
- หลัง merge ตรวจว่า deploy เสร็จ (`gh api repos/playlist-system/playlist_system/deployments`) แล้วบอกผู้ใช้กด **Ctrl+F5**
- เลขเวอร์ชัน: แก้เล็ก = ขยับเลขท้าย (5.5.36 → 5.5.37)
- อย่าลืม `git add versions/` (set_version สร้างสำเนาเวอร์ชันเดิม + หน้าเลือกเวอร์ชันใหม่)

### ถ้าเวอร์ชันใหม่มีปัญหา

- ทีมกด “ใช้เวอร์ชันก่อนหน้า” ท้ายหน้าระบบ → `versions/` เลือกเวอร์ชันที่ใช้ได้ (ตั้งแต่ 5.5.37 หน้าเวอร์ชันเก่าขึ้นแถบเหลืองเอง เพราะโค้ดเช็ค `OLDVER` จาก path)
- คนดูแลระบบ: แก้แล้วออกเวอร์ชันใหม่ตามปกติ — ถ้าต้องถอยทั้งเว็บ ใช้ `git revert` ของ merge นั้นแล้วออกเป็นเวอร์ชันใหม่
- เก็บสำเนาเองได้: `python3 dev/snapshot.py --rev <commit>` (สร้างหน้าเลือกเวอร์ชันใหม่ให้)

## 5. ชุดทดสอบ (regression)

- ไฟล์ทดสอบ (Playlist / ใบเช็คจริง + ผลที่ถูกต้อง) อยู่ใน **zip `photo-check-tests` ใน Google Drive ของบริษัท** — ขอจากเจ้าของระบบ แนบเข้าแชต แล้วแตกไฟล์
- **ห้าม commit โฟลเดอร์นี้เข้า repo**

```bash
python3 dev/tests/run_tests.py playlist_system.html photo-check-tests             # ทดสอบทั้งหมด (~10 นาที)
python3 dev/tests/run_tests.py playlist_system.html photo-check-tests --only nov10 # เฉพาะบางเคส
python3 dev/tests/run_tests.py playlist_system.html photo-check-tests --update     # ยอมรับผลใหม่เป็น expected
```

| เคส | ตรวจอะไร |
|---|---|
| `fill/*` | ใส่ Playlist + ใบเช็ค → กดเติม → เทียบไฟล์ .xlsm ทุกเซลล์ (ค่า สี ตัวหนา ขนาด/ชื่อฟอนต์) + รายงาน |
| `folders/*` | อ่านใบเช็คแบบหน้า "งานโฟลเดอร์และรูป" → แผน zip → แผนสร้างโฟลเดอร์บนโครงจำลอง → โฟลเดอร์ที่สร้างจริง |
| `feedback` | ปุ่มแจ้งปัญหาส่งครบ (ข้อความ รูป ไฟล์ ข้อมูลระบบ) ไปที่รับจำลอง + จำลอง `feedback.gs` |

- นาฬิกาในเบราว์เซอร์ทดสอบถูกตั้งตายตัว (`now` ใน cases.json) ผลจึงเหมือนเดิมทุกครั้ง
- เพิ่มเคสใหม่: ใส่ไฟล์ในโฟลเดอร์ชุดทดสอบ เพิ่มใน `cases.json` รัน `--only <ชื่อ> --update` ด้วยเวอร์ชันที่ถูกต้อง แล้ว zip ขึ้น Drive ใหม่
- เมื่อแก้แล้วผลเปลี่ยนตามที่ตั้งใจ: ตรวจความต่างที่รายงาน → `--update` → zip ชุดทดสอบใหม่ส่งให้เจ้าของระบบเก็บใน Drive
- `dev/tests/feedback_gs_sim.js` รันเดี่ยวได้: `node dev/tests/feedback_gs_sim.js apps_script/feedback.gs`

## 6. แผนที่โค้ด (ค้นชื่อใน playlist_system.html)

| ส่วน | ฟังก์ชันหลัก |
|---|---|
| อ่าน Playlist | `parsePlaylist` (แยกตามแบบไฟล์: `parseSect` `parseSplit` `parseSig` `parseMetro` `parseMco` `parseVen`) • `plCheck` (ตรวจแถว วันที่ผิด แถวเลื่อนช่อง เลขเกิน `durTail`) • `plDays` `plSync` |
| อ่าน / วางแผนใบเช็ค | `layoutChk` (แผนต่อชีต: จอไหน แถวไหน เตือนอะไร) • `chkRows` (แถวที่จะใส่ต่อจอ, แถวหมายเหตุ `RMK`) • `sheetNeeds` • `namedScreen` `screen10` `picId` |
| เติม | `fillRun` → `fillDays` (หลายวัน) → `fillWb` (เขียน .xlsm ผ่าน JSZip เก็บมาโคร) • `oldHdrFix` (หัวแบบเก่า) • `autoMv` `colWb` `xHist` (ย้าย / เพิ่มคอลัมน์ BKJ, Metro กลางเดือน) |
| ใบเปล่า / ย้ายจอ | `sdbLoad` `sdbRun` `sdbApply` `sdbMoveApply` `sdbNewFile` (สร้างใบเปล่าให้คนใหม่) `blankWb` |
| โฟลเดอร์รูป | `tcRead` (อ่านใบเช็คในหน้าเครื่องมือ → `PICX`) • `pzPlan` `pzScr` `pzZip` (zip โฟลเดอร์) • `picScreens` `picPlan` `picMake` (สร้างในโฟลเดอร์ที่เลือก) • `fxKind` `fxIdx` `FXP` (MRT / The 20 / Air Traveller / Gateway X / Rama 9) • `bkjAB` (ต่อท้าย (A)/(R) ของ BKJ) |
| หน้าจอ | `view()` วาดใหม่ทั้งหน้า • `vHome` `vPl` `vCk` `vFl` `vRes` `vTools` `vSdb` `vFb` … • ปุ่มทั้งหมดอยู่ใน `ACT` (`data-a="ชื่อ"`) |
| แจ้งปัญหา | `FB_URL` `fbUrl` `vFb` `ACT.fbsend` (ฟอร์มในแท็บใหม่ + รับ postMessage) |

## 7. เรื่องที่รู้แล้วแต่ยังไม่ได้ทำ

- MRT / Gateway X: จอในสถานี / กลุ่มเดียวกันได้โฟลเดอร์หนังชุดเดียวกัน (ยังแยกไม่ได้จากใบเช็ค)
- เก็บ Playlist จากอีเมลเข้า Google Drive อัตโนมัติด้วย Apps Script ของบริษัท (ยังไม่เริ่ม)
