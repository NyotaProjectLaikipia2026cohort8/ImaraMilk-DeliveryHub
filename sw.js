const CACHE_NAME="imara-milk-shell-v1";
const APP_SHELL=["./","./index.html","./manifest.webmanifest","./icons/icon-192.png","./icons/icon-512.png","./icons/apple-touch-icon.png"];
self.addEventListener("install",event=>{
  event.waitUntil(caches.open(CACHE_NAME).then(cache=>cache.addAll(APP_SHELL)).then(()=>self.skipWaiting()));
});
self.addEventListener("activate",event=>{
  event.waitUntil(caches.keys().then(keys=>Promise.all(keys.filter(key=>key!==CACHE_NAME).map(key=>caches.delete(key)))).then(()=>self.clients.claim()));
});
self.addEventListener("fetch",event=>{
  const request=event.request;
  if(request.method!=="GET")return;
  const url=new URL(request.url);
  // Do not intercept remote API, Maps, fonts or other third-party requests.
  if(url.origin!==self.location.origin)return;
  if(url.pathname.endsWith("/sw.js"))return;
  event.respondWith(caches.match(request).then(cached=>{
    const network=fetch(request).then(response=>{
      if(response&&response.ok&&response.type==="basic"){
        const copy=response.clone();caches.open(CACHE_NAME).then(cache=>cache.put(request,copy));
      }
      return response;
    }).catch(()=>cached||caches.match("./index.html"));
    return cached||network;
  }));
});
