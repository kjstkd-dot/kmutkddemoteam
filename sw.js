const CACHE = 'kstar-v2';

self.addEventListener('install', function(){ self.skipWaiting(); });

self.addEventListener('activate', function(e){
  e.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(keys.filter(function(k){ return k !== CACHE; })
                            .map(function(k){ return caches.delete(k); }));
    }).then(function(){ return self.clients.claim(); })
  );
});

// 화면(HTML)은 항상 새로 받아온다 — 예전 화면이 남아 헷갈리지 않도록.
// 아이콘·이미지만 저장해두고, 인터넷이 끊겼을 때 대신 보여준다.
self.addEventListener('fetch', function(e){
  var req = e.request;
  if(req.method !== 'GET') return;
  if(new URL(req.url).origin !== self.location.origin) return;
  if(req.mode === 'navigate' || (req.headers.get('accept') || '').indexOf('text/html') > -1) return;

  e.respondWith(
    fetch(req).then(function(res){
      var copy = res.clone();
      caches.open(CACHE).then(function(c){ c.put(req, copy); }).catch(function(){});
      return res;
    }).catch(function(){
      return caches.match(req);
    })
  );
});
