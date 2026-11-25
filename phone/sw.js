/* แอปงานถ่ายรูป: เก็บหน้าแอปไว้ในเครื่อง เปิดได้แม้ไม่มีเน็ต (ข้อมูลใบเช็คไม่ผ่านไฟล์นี้ — แอปเก็บเองใน localStorage)
   เปิดแอป = ใช้หน้าที่เก็บไว้ทันที แล้วโหลดหน้าใหม่เก็บไว้สำหรับครั้งหน้า • เปลี่ยน CACHE เมื่ออยากให้ทุกเครื่องล้างของเก่า */
const CACHE = 'pcs-phone-v1';
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icon-192.png', 'icon-512.png', 'icon-maskable-512.png'];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;   // ข้อมูลจาก Google ไม่ผ่านที่นี่
  e.respondWith(caches.open(CACHE).then(async c => {
    const key = e.request.mode === 'navigate' ? './' : e.request;
    const hit = await c.match(key, { ignoreSearch: true });
    const net = fetch(e.request).then(r => { if (r && r.ok) c.put(key, r.clone()); return r; }).catch(() => hit);
    if (hit) { e.waitUntil(net.catch(() => {})); return hit; }
    return net;
  }));
});
