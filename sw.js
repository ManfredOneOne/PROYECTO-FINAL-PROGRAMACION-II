/* Service worker: guarda la app en el dispositivo para que abra sin internet.
   Estrategia: primero la red (para recibir cambios) y, si no hay conexión, la copia guardada. */
var CACHE = "codebran-v2";
var ARCHIVOS = [
  "./index.html",
  "./proyecto_final_v3.html",
  "./logo_umg.png",
  "./manifest.webmanifest",
  "./iconos/icono-32.png",
  "./iconos/icono-180.png",
  "./iconos/icono-192.png",
  "./iconos/icono-512.png"
];

self.addEventListener("install", function(ev){
  ev.waitUntil(caches.open(CACHE).then(function(c){ return c.addAll(ARCHIVOS); }).then(function(){ return self.skipWaiting(); }));
});

self.addEventListener("activate", function(ev){
  ev.waitUntil(caches.keys().then(function(ks){
    return Promise.all(ks.filter(function(k){ return k !== CACHE; }).map(function(k){ return caches.delete(k); }));
  }).then(function(){ return self.clients.claim(); }));
});

self.addEventListener("fetch", function(ev){
  if(ev.request.method !== "GET" || new URL(ev.request.url).origin !== location.origin) return;
  ev.respondWith(
    fetch(ev.request).then(function(r){
      if(r && r.ok){ var copia = r.clone(); caches.open(CACHE).then(function(c){ c.put(ev.request, copia); }); }
      return r;
    }).catch(function(){
      return caches.match(ev.request, {ignoreSearch:true}).then(function(r){ return r || caches.match("./proyecto_final_v3.html"); });
    })
  );
});
