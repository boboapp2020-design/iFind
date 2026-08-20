const CACHE="ifind-v4";
const SHELL=["./","./index.html","./icon-192.png","./icon-512.png"];
self.addEventListener("install",e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(SHELL)).then(()=>self.skipWaiting()));});
self.addEventListener("activate",e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));});
self.addEventListener("fetch",e=>{
  const u=new URL(e.request.url);
  if(u.hostname.indexOf("script.google")>=0) return;                 // live API — network เสมอ
  if(u.pathname.endsWith("manifest.json")||u.pathname.endsWith("sw.js")) return; // อย่าแคช 2 ตัวนี้ (กันติดตั้งไม่ได้)
  // network-first สำหรับหน้าเว็บ (ได้เวอร์ชันใหม่ทันทีเมื่อออนไลน์, ออฟไลน์ใช้ cache)
  if(e.request.mode==="navigate"||u.pathname.endsWith("index.html")||u.pathname.endsWith("/")){
    e.respondWith(fetch(e.request).then(res=>{const cl=res.clone();caches.open(CACHE).then(c=>c.put(e.request,cl));return res;})
      .catch(()=>caches.match(e.request).then(r=>r||caches.match("./index.html"))));
    return;
  }
  e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request).then(res=>{
    if(e.request.method==="GET"&&u.origin===location.origin){const cl=res.clone();caches.open(CACHE).then(c=>c.put(e.request,cl));}
    return res;
  })));
});
