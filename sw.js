/* GTPro Portfolio — service worker
 *
 * กลยุทธ์: network-first สำหรับเปลือกแอป
 *   เมื่อมีเน็ต → เอาของใหม่จากเซิร์ฟเวอร์เสมอ แล้วอัปเดตแคชไว้เผื่อออฟไลน์
 *   ไม่มีเน็ต / ช้าเกิน 4 วิ → ใช้ของในแคชแทน แอปยังเปิดได้
 *
 * ⚠️ ห้ามกลับไปใช้ cache-first เด็ดขาด
 * ของเดิม (ถึง gtpro-v7) เป็น `caches.match() || fetch()` ผลคือพอ index.html เข้าแคชแล้ว
 * ผู้ใช้จะติดอยู่กับเวอร์ชันเก่าตลอด แม้ push ของใหม่ขึ้น GitHub Pages แล้วก็ตาม
 * (ต้องรอ V เปลี่ยน + reload สองรอบถึงจะได้ของใหม่ ซึ่งผู้ใช้ไม่มีทางรู้)
 * ส่วนข้อมูล (คำขอที่มี query string) ไม่แตะเลย ปล่อยให้ดึงสดทุกครั้ง
 */
const V = 'gtpro-v8';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png'];
const NET_TIMEOUT = 4000;

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(V)
      .then(c => Promise.allSettled(SHELL.map(u => c.add(u))))  // ไฟล์ใดพลาดไม่ล้มทั้งชุด
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

/* หน้าเว็บสั่งให้ SW ตัวใหม่ทำงานทันทีได้ (ใช้ตอนกดปุ่ม "โหลดเวอร์ชันใหม่") */
self.addEventListener('message', e => {
  if (e.data === 'skipWaiting') self.skipWaiting();
});

/* เปลือกแอป = ไฟล์นิ่ง ๆ โดเมนเดียวกัน ไม่มี query string
   ไม่รวม: คำขอข้อมูล, version.json, ไฟล์ APK และตัว sw.js เอง */
function isShell(url) {
  if (url.origin !== self.location.origin) return false;
  if (url.search) return false;
  const p = url.pathname;
  if (p.indexOf('version.json') >= 0) return false;
  if (p.indexOf('/releases/') >= 0) return false;
  if (/\/sw\.js$/.test(p)) return false;
  return /\.(html|css|png|svg|webmanifest|woff2?)$/.test(p) || p.endsWith('/');
}

function fromNetwork(req) {
  return new Promise((resolve, reject) => {
    const t = setTimeout(() => reject(new Error('timeout')), NET_TIMEOUT);
    fetch(req).then(res => { clearTimeout(t); resolve(res); },
                    err => { clearTimeout(t); reject(err); });
  });
}

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;

  let url;
  try { url = new URL(req.url); } catch (err) { return; }
  if (!isShell(url)) return;

  e.respondWith(
    fromNetwork(req)
      .then(res => {
        if (res && res.status === 200 && res.type === 'basic') {
          const copy = res.clone();
          caches.open(V).then(c => c.put(req, copy)).catch(() => {});
        }
        return res;
      })
      .catch(() =>                                   // เน็ตล่ม/ช้า → ของในแคช
        caches.match(req).then(hit => hit || caches.match('./index.html'))
      )
  );
});
