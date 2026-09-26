const CACHE = 'jaseojeon-v1';

// 새 버전이 올라오면 기다리지 않고 바로 넘겨받는다.
self.addEventListener('install', function(){ self.skipWaiting(); });

self.addEventListener('activate', function(e){
  e.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(keys.filter(function(k){ return k !== CACHE; })
                            .map(function(k){ return caches.delete(k); }));
    }).then(function(){ return self.clients.claim(); })
  );
});

// 화면과 자료는 늘 새로 받아온다. 아이콘·이미지만 저장해두고
// 인터넷이 끊겼을 때 대신 보여준다.
self.addEventListener('fetch', function(e){
  var req = e.request;
  if(req.method !== 'GET') return;

  var url;
  try{ url = new URL(req.url); }catch(err){ return; }
  if(url.origin !== self.location.origin) return;
  if(!/\.(png|jpg|jpeg|svg|webp|ico)$/i.test(url.pathname)) return;

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
