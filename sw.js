/* GTPro Portfolio — service worker
   เปลือกแอปแคชไว้ให้เปิดได้แม้ไม่มีเน็ต ส่วนข้อมูลดึงสดเสมอ
   (ข้อมูลล่าสุดถูกเก็บใน localStorage โดยตัวหน้าเว็บเอง) */
const V = 'gtpro-v4';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(V).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;

  // อย่าแคชคำขอข้อมูล — ต้องได้ตัวเลขล่าสุดเสมอ
  if (req.url.includes('script.google.com') || req.url.includes('googleusercontent.com')) return;

  e.respondWith(
    caches.match(req).then(hit =>
      hit || fetch(req).then(res => {
        if (res && res.status === 200 && res.type === 'basic') {
          const copy = res.clone();
          caches.open(V).then(c => c.put(req, copy));
        }
        return res;
      }).catch(() => caches.match('./index.html'))
    )
  );
});
