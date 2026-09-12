// Service Worker - Sistema Medico Muratori (S.M.M.)
// Cachea el "cascaron" (todos los HTML de los modulos + iconos) para que
// la app abra al instante e instalada, funcione como una app nativa.
// Las llamadas a Google Sheets / Gemini / Mercado Pago SIEMPRE van a la red
// (nunca se cachean, porque son datos vivos de pacientes y pagos).

const CACHE_NAME = 'smm-shell-v1';
const SHELL_FILES = [
  './INDEX_SIM.html',
  './CAM.html',
  './OCM_Muratori_v13.html',
  './RIDM_Muratori_v1.html',
  './RSP_Muratori_v1.html',
  './KSM_Muratori_v1.html',
  './EVAL_Muratori_v1.html',
  './RCM_Muratori_v1.html',
  './IMDCYMT_Paciente_v1.html',
  './Estadisticas_SMM.html',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './agenda-familiar/index.html',
  './agenda-familiar/app.js',
  './agenda-familiar/style.css',
  './lector-cud/index.html'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES)).catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const url = event.request.url;
  // Nunca cachear llamadas a APIs externas o al backend propio (siempre a la red)
  if (
    url.includes('script.google.com') ||
    url.includes('generativelanguage.googleapis.com') ||
    url.includes('mercadopago.com') ||
    url.includes('nominatim.openstreetmap.org') ||
    url.includes('fonts.g')
  ) {
    return; // dejar pasar directo a la red
  }
  event.respondWith(
    caches.match(event.request).then((cached) => cached || fetch(event.request))
  );
});
