const CACHE='universe-invoice-v0.13.2';
const DEP_CACHE='universe-invoice-dependencies-v1';
const LOCAL_ASSETS=['./','./index.html','./styles.css','./app.js','./manifest.webmanifest','./icon.svg','./icon-192.png','./icon-512.png'];
const EXTERNAL_ASSETS=[
  'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
  'https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js',
  'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js',
  'https://cdn.jsdelivr.net/npm/sortablejs@1.15.2/Sortable.min.js',
  'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js'
];
self.addEventListener('install',e=>{
  self.skipWaiting();
  e.waitUntil((async()=>{
    const local=await caches.open(CACHE);await local.addAll(LOCAL_ASSETS);
    const deps=await caches.open(DEP_CACHE);
    // Cache every core library before activating this version. no-cors permits
    // caching opaque CDN script responses for later offline script loading.
    await Promise.all(EXTERNAL_ASSETS.map(async url=>{
      const req=new Request(url,{mode:'no-cors',cache:'reload'}),res=await fetch(req);
      await deps.put(url,res.clone());
    }));
  })());
});
self.addEventListener('activate',e=>e.waitUntil((async()=>{
  await Promise.all((await caches.keys()).filter(x=>x!==CACHE&&x!==DEP_CACHE).map(x=>caches.delete(x)));
  await self.clients.claim();
})()));
self.addEventListener('fetch',e=>{
  if(e.request.method!=='GET')return;
  const url=new URL(e.request.url),isLocal=url.origin===self.location.origin,isDependency=EXTERNAL_ASSETS.includes(url.href);
  if(!isLocal&&!isDependency)return;
  e.respondWith((async()=>{
    const cached=await caches.match(e.request,{ignoreSearch:isLocal})||await caches.match(url.href);
    if(cached){
      if(isLocal)fetch(e.request).then(async r=>{if(r?.ok){const c=await caches.open(CACHE);await c.put(e.request,r.clone())}}).catch(()=>{});
      return cached;
    }
    try{const response=await fetch(e.request);if(response){const c=await caches.open(isDependency?DEP_CACHE:CACHE);await c.put(url.href,response.clone())}return response}catch(err){if(isLocal&&e.request.mode==='navigate')return caches.match('./index.html');throw err}
  })());
});
