const CACHE="ifind-v1";
const SHELL=["./","./index.html","./manifest.json","./icon-192.png","./icon-512.png"];
self.addEventListener("install",e=>{e.waitUntil(caches.open(CACHE).then(c=>c.addAll(SHELL)).then(()=>self.skipWaiting()));});
self.addEventListener("activate",e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()));});
self.addEventListener("fetch",e=>{
  const u=new URL(e.request.url);
  if(u.hostname.indexOf("script.google")>=0) return; // live API — always network
  e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request).then(res=>{
    if(e.request.method==="GET"&&u.origin===location.origin){const cl=res.clone();caches.open(CACHE).then(c=>c.put(e.request,cl));}
    return res;
  }).catch(()=>caches.match("./index.html"))));
});