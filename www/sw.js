// Koray - Rafta: internetsiz çalışma için önbellek
const C="depobul-v40",FILES=["./","index.html","manifest.json","icon-192.png","icon-512.png","fonts/fonts.css"];
self.addEventListener("install",e=>{e.waitUntil(caches.open(C).then(c=>c.addAll(FILES)));self.skipWaiting()});
self.addEventListener("activate",e=>{e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x!==C).map(x=>caches.delete(x)))));self.clients.claim()});
self.addEventListener("fetch",e=>{if(e.request.method!=="GET")return;
  e.respondWith(fetch(e.request,{cache:"no-cache"}).then(r=>{const cp=r.clone();caches.open(C).then(c=>c.put(e.request,cp)).catch(()=>{});return r}).catch(()=>caches.match(e.request,{ignoreSearch:true}).then(r=>r||caches.match("index.html"))))});
// Bildirimler: Firebase, uygulama kapalıyken gelen bildirimi gösterir
self.addEventListener("push",()=>{self.clients.matchAll({type:"window",includeUncontrolled:true}).then(l=>l.forEach(c=>c.postMessage({rafta:"esitle"})))});
try{importScripts("https://www.gstatic.com/firebasejs/10.14.1/firebase-app-compat.js","https://www.gstatic.com/firebasejs/10.14.1/firebase-messaging-compat.js");
  firebase.initializeApp({apiKey:"AIzaSyCJ5Yrff8yIy4XM1uc3Np-kG6iM2KumxAE",authDomain:"koray-rafta.firebaseapp.com",projectId:"koray-rafta",storageBucket:"koray-rafta.firebasestorage.app",messagingSenderId:"547988807383",appId:"1:547988807383:web:9d92c8c551ebb6ff555a92"});
  firebase.messaging()}catch(e){}
