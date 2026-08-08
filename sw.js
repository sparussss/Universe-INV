const CACHE='universe-invoice-v0.14.8';
const DEP_CACHE='universe-invoice-dependencies-v1';
const LOCAL_ASSETS=['./','./index.html','./styles.css?v=0.14.8','./app-v0.14.8.js','./manifest.webmanifest?v=0.14.8','./icon.svg','./icon-192.png','./icon-512.png'];
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
    if(isDependency){
      const cached=await caches.match(url.href);
      if(cached)return cached;
      try{const response=await fetch(e.request);if(response){const c=await caches.open(DEP_CACHE);await c.put(url.href,response.clone())}return response}catch(err){throw err}
    }
    // Local PWA code is network-first so an updated index/app cannot be mixed with an old cached app.js.
    try{
      const response=await fetch(e.request,{cache:'no-store'});
      if(response?.ok){const c=await caches.open(CACHE);await c.put(e.request,response.clone())}
      return response;
    }catch(err){
      const cached=await caches.match(e.request)||await caches.match(url.pathname.endsWith('/')?'./index.html':e.request,{ignoreSearch:false});
      if(cached)return cached;
      if(e.request.mode==='navigate')return caches.match('./index.html');
      throw err;
    }
  })());
});
