/* GTPro Portfolio — service worker
   เปลือกแอปแคชไว้ให้เปิดได้แม้ไม่มีเน็ต ส่วนข้อมูลดึงสดเสมอ
   (ข้อมูลล่าสุดถูกเก็บใน localStorage โดยตัวหน้าเว็บเอง) */
const V = 'gtpro-v7';
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

/* แคชเฉพาะ "เปลือกแอป" เท่านั้น — ไฟล์นิ่ง ๆ ที่อยู่โดเมนเดียวกันและไม่มี query string
   อย่างอื่น (คำขอข้อมูล, version.json, รูปจากไดรฟ์) ปล่อยผ่านให้ดึงสดทุกครั้ง
   ไม่งั้นจะได้ตัวเลขเก่าค้างโดยไม่รู้ตัว */
function isShell(url) {
  if (url.origin !== self.location.origin) return false;
  if (url.search) return false;                       // มี ?… = คำขอข้อมูล
  if (url.pathname.indexOf('version.json') >= 0) return false;
  if (url.pathname.indexOf('/releases/') >= 0) return false;
  return /\.(html|js|css|png|svg|webmanifest|woff2?)$/.test(url.pathname) ||
         url.pathname.endsWith('/');
}

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;

  let url;
  try { url = new URL(req.url); } catch (err) { return; }
  if (!isShell(url)) return;

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
