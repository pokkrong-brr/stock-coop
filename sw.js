/* เปิดแอปได้เร็วและเปิดได้แม้เน็ตหลุด: เก็บไฟล์หน้าเว็บไว้ในเครื่อง
   ข้อมูลจริงยังวิ่งไป Google Sheets เสมอ (ไม่ cache การเรียก Apps Script) */
const V = 'stock-coop-2026-11-08.2';
const SHELL = ['./', 'index.html', 'manifest.webmanifest', 'icons/icon.svg', 'icons/icon-192.png', 'icons/apple-touch-icon.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(V).then(c => c.addAll(SHELL))); self.skipWaiting(); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const r = e.request, u = new URL(r.url);
  if (r.method !== 'GET' || /(^|\.)google(usercontent)?\.com$/.test(u.hostname) && !/fonts\./.test(u.hostname)) return;
  if (/fonts\.(googleapis|gstatic)\.com$/.test(u.hostname)) { // ฟอนต์: ใช้ของในเครื่องก่อน
    e.respondWith(caches.match(r).then(m => m || fetch(r).then(res => { const c = res.clone(); caches.open(V).then(x => x.put(r, c)); return res; })));
    return;
  }
  if (u.origin !== location.origin) return;
  // ไฟล์ของแอป: ดึงใหม่ก่อน (ได้เวอร์ชันล่าสุดเสมอ) ถ้าออฟไลน์ใช้ของในเครื่อง
  e.respondWith(fetch(r).then(res => { if (res.ok) { const c = res.clone(); caches.open(V).then(x => x.put(r, c)); } return res; })
    .catch(() => caches.match(r, { ignoreSearch: true }).then(m => m || caches.match('index.html'))));
});
