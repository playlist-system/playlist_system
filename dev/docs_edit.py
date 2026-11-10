"""แก้ไฟล์เอกสาร (README / คู่มือ) แบบแทนข้อความ — ข้อความเดิมต้องพบ 1 ครั้งพอดี ไม่งั้นหยุด

ใช้ในสคริปต์ออกเวอร์ชัน:
    import sys; sys.path.insert(0, 'dev'); from docs_edit import edit, MANUAL
    edit('README.md', [(old, new), ...])
    edit(MANUAL, [(old, new)], crlf=True)     # คู่มือ .txt เป็น CRLF (+BOM) — เขียน \\n ได้เลย

รักษา BOM และ CRLF ของไฟล์ไว้เหมือนเดิม
"""
MANUAL = 'คู่มือการใช้งาน_ระบบเติมใบเช็ครูป.txt'


def edit(path, pairs, crlf=False):
    raw = open(path, 'rb').read()
    bom = raw.startswith(b'\xef\xbb\xbf')
    t = raw.decode('utf-8-sig')
    if crlf:
        t = t.replace('\r\n', '\n')
    for a, b in pairs:
        n = t.count(a)
        if n != 1:
            raise SystemExit('%s: พบ %d ครั้ง (ต้อง 1) สำหรับ %r' % (path, n, a[:70]))
        t = t.replace(a, b)
    if crlf:
        t = t.replace('\n', '\r\n')
    open(path, 'wb').write((b'\xef\xbb\xbf' if bom else b'') + t.encode('utf-8'))
