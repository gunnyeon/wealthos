// WEALTHOS Family — Service Worker
// 앱 껍데기(UI)는 오프라인에서도 로드, 데이터는 온라인 시 동기화

const CACHE = 'wealthos-v1';
const SHELL = ['./index.html', './manifest.json', './icon.svg'];

// 설치: 앱 껍데기 캐시
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())
  );
});

// 활성화: 이전 캐시 정리
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

// 요청 처리
self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);

  // Apps Script API 요청 → 네트워크 우선 (항상 최신 데이터)
  if (url.hostname === 'script.google.com') {
    e.respondWith(fetch(e.request).catch(() =>
      new Response(JSON.stringify({ ok: false, error: '오프라인 상태입니다' }), {
        headers: { 'Content-Type': 'application/json' }
      })
    ));
    return;
  }

  // 앱 UI 파일 → 캐시 우선 (오프라인에서도 UI 로드)
  e.respondWith(
    caches.match(e.request).then(cached => cached || fetch(e.request).then(res => {
      // 성공한 응답은 캐시에 저장
      if (res.ok && res.type !== 'opaque') {
        const clone = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, clone));
      }
      return res;
    }))
  );
});
