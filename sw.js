const V = 'wo-v4';
const S = ['./index.html','./manifest.json','./icon.svg'];
self.addEventListener('install', e => e.waitUntil(caches.open(V).then(c=>c.addAll(S)).then(()=>self.skipWaiting())));
self.addEventListener('activate', e => e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==V).map(k=>caches.delete(k)))).then(()=>self.clients.claim())));
self.addEventListener('fetch', e => {
  if (new URL(e.request.url).hostname === 'script.google.com') {
    e.respondWith(fetch(e.request).catch(()=>new Response(JSON.stringify({ok:false,error:'오프라인'}),{headers:{'Content-Type':'application/json'}})));
    return;
  }
  e.respondWith(caches.match(e.request).then(c=>c||fetch(e.request)));
});
