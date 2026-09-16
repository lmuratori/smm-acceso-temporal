/* ============================================================
   EFECTO LUPA — Sistema Médico Muratori (SMM)
   ============================================================
   Qué hace:
   - Agranda (x3) el campo de texto que se está completando,
     mientras tiene el foco (cursor adentro), para que sea más
     fácil de leer mientras se escribe.
   - Funciona SOLO con tener este archivo incluido en la página —
     no hace falta marcar campo por campo, agarra automáticamente
     cualquier <input>, <textarea> o campo "contenteditable" del
     módulo. Es el mismo principio que el corrector ortográfico
     del navegador: anda en cualquier campo sin configurarlo.
   - Al salir del campo (Tab, click afuera, etc.) vuelve solo a su
     tamaño normal.
   - No agranda casillas de verificación, botones, radios, ni
     campos deshabilitados o de solo lectura — solo campos donde
     realmente se escribe texto.
   - Para dejar UN campo puntual sin este efecto (si en algún lugar
     molesta), agregarle el atributo data-lupa="off" en el HTML.

   CÓMO INSTALARLO EN CADA MÓDULO:
   Agregar esta línea justo antes de </body>:
     <script src="efecto-lupa.js" defer></script>

   LIMITACIÓN CONOCIDA:
   Si el campo está dentro de un panel/modal con scroll o con
   "overflow" recortado, el agrandado puede quedar parcialmente
   tapado por el borde de ese panel (es una limitación visual del
   truco de CSS usado, no un error — si se ve mal en algún panel
   puntual, avisar para ajustarlo ahí).
   ============================================================ */
(function(){
  if(window.__efectoLupaSMM) return; // evita duplicar si el script se incluye más de una vez
  window.__efectoLupaSMM = true;

  var ESCALA = 3;
  var SELECTOR_CAMPOS = 'input, textarea, [contenteditable="true"], [contenteditable=""]';
  var TIPOS_EXCLUIDOS = ['checkbox','radio','button','submit','reset','range','color','file','hidden','image'];

  var estilo = document.createElement('style');
  estilo.textContent =
    '.efecto-lupa-activo{' +
      'position:relative !important;' +
      'z-index:999999 !important;' +
      'transition:transform .12s ease-out !important;' +
      'box-shadow:0 10px 34px rgba(0,0,0,.38) !important;' +
      'background:#fff !important;' +
    '}';
  (document.head || document.documentElement).appendChild(estilo);

  function elegible(el){
    if(!el || !el.matches) return false;
    if(el.getAttribute && el.getAttribute('data-lupa')==='off') return false;
    if(!el.matches(SELECTOR_CAMPOS)) return false;
    if(el.tagName==='INPUT' && TIPOS_EXCLUIDOS.indexOf((el.type||'').toLowerCase())>-1) return false;
    if(el.disabled || el.readOnly) return false;
    return true;
  }

  function activar(el){
    if(!elegible(el)) return;
    var r = el.getBoundingClientRect();
    // Elige el origen de la ampliación según en qué parte de la pantalla
    // está el campo, para que crezca hacia el espacio libre en vez de
    // salirse de la pantalla.
    var origenX = r.left < window.innerWidth/3 ? 'left' : (r.right > window.innerWidth*2/3 ? 'right' : 'center');
    var origenY = r.top < window.innerHeight/3 ? 'top' : (r.bottom > window.innerHeight*2/3 ? 'bottom' : 'center');
    el.style.transformOrigin = origenX+' '+origenY;
    el.classList.add('efecto-lupa-activo');
    el.style.transform = 'scale('+ESCALA+')';
  }

  function desactivar(el){
    if(!el || !el.classList || !el.classList.contains('efecto-lupa-activo')) return;
    el.classList.remove('efecto-lupa-activo');
    el.style.transform = '';
    el.style.transformOrigin = '';
  }

  // focusin/focusout burbujean nativamente (a diferencia de focus/blur),
  // así que un solo listener en document alcanza para todo el documento,
  // incluyendo campos agregados dinámicamente después (modales, etc.).
  document.addEventListener('focusin', function(e){ activar(e.target); });
  document.addEventListener('focusout', function(e){ desactivar(e.target); });
})();
