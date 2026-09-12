// Service Worker - Sistema Medico Muratori (S.M.M.)
// Cachea el "cascaron" (los módulos más usados + iconos) para que la app
// abra rápido si queda instalada, y sirve de respaldo si en algún momento
// no hay señal. Las llamadas a Google Sheets / Gemini / Mercado Pago
// SIEMPRE van a la red (nunca se cachean, porque son datos vivos de
// pacientes y pagos).
//
// REESCRITO — la lista de archivos estaba desactualizada (apuntaba a
// nombres viejos que ya no existen: CAM.html, OCM_Muratori_v13.html,
// RSP_Muratori_v1.html, KSM_Muratori_v1.html, EVAL_Muratori_v1.html,
// RCM_Muratori_v1.html, manifest.json, la carpeta icons/, agenda-familiar/
// y lector-cud/ — nada de eso está publicado). Con cache.addAll(), si UN
// solo archivo de la lista no existe, TODO el precacheo fallaba en
// silencio (estaba atrapado con .catch(()=>{})) y el Service Worker viejo
// se quedaba activo indefinidamente, sirviendo versiones viejas cacheadas
// sin que ninguna actualización de GitHub Pages se notara — esto explica
// varias veces que "sigue igual"/"no abre" después de actualizar.
//
// Dos cambios de fondo:
//  1) El precacheo ahora es archivo por archivo (no todo-o-nada), así un
//     nombre que cambie en el futuro no rompe a los demás.
//  2) El fetch pasó de "cache primero" a "red primero, caché como
//     respaldo" — así cualquier actualización que subas a GitHub Pages se
//     ve enseguida (con conexión), y el caché solo se usa si en el momento
//     no hay señal.

const CACHE_NAME = 'smm-shell-v2';
const SHELL_FILES = [
  './index.html',
  './OCM_Muratori_v14_0_26.html',
  './CAM_Herramientas_Clinicas.html',
  './INDEX_SIM.html',
  './RIDM_Muratori_v1.html',
  './RMYT_Autoevaluacion_Semanal_v1.html',
  './Recomendaciones_Medicas_Terapeuticas_v1.html',
  './Estadisticas_SMM.html',
  './favicon.ico',
  './apple-touch-icon.png',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) =>
      Promise.all(
        SHELL_FILES.map((f) => cache.add(f).catch(() => {
          // Si un archivo puntual no existe o falla, no tira abajo el
          // precacheo de los demás — solo ese queda sin cachear.
        }))
      )
    )
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
  // Red primero — si hay conexión, siempre trae la última versión
  // publicada. Solo si falla (sin señal) usa lo que haya en caché.
  event.respondWith(
    fetch(event.request)
      .then((resp) => {
        const copy = resp.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy)).catch(() => {});
        return resp;
      })
      .catch(() => caches.match(event.request))
  );
});
