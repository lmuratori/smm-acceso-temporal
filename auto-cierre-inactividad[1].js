/* ============================================================
   AUTO-CIERRE POR INACTIVIDAD — Sistema Médico Muratori (SMM)
   ============================================================
   Qué hace:
   - Al abrir un módulo (CAM, OCM, RIDM, etc.) pregunta cada
     cuánto tiempo de inactividad se debe cerrar solo.
   - Si no hay ningún click, tecla o toque en pantalla durante
     ese tiempo, muestra un aviso de 60 segundos y, si nadie
     responde, cierra la consulta (vuelve al menú principal).
   - Así se libera la conexión para poder entrar desde otro
     dispositivo (celular / computadora) sin quedar trabado.

   CÓMO INSTALARLO EN CADA MÓDULO:
   Agregar esta línea justo antes de </body> en cada archivo
   HTML (CAM.html, OCM_Muratori_v13.html, RIDM_Muratori_v1.html,
   RSP_Muratori_v1.html, KSM_Muratori_v1.html, EVAL_Muratori_v1.html,
   RCM_Muratori_v1.html, IMDCYMT_Paciente_v1.html):

     <script src="auto-cierre-inactividad.js"></script>

   Para los módulos que están en subcarpetas (agenda-familiar/,
   lector-cud/), usar en cambio:

     <script>window.SMM_INDEX_URL = '../INDEX_SIM.html';</script>
     <script src="../auto-cierre-inactividad.js"></script>

   No hace falta tocar nada más del código existente.
   ============================================================ */

(function () {
  'use strict';

  // A dónde volver cuando se cierra por inactividad
  var INDEX_URL = window.SMM_INDEX_URL || 'INDEX_SIM.html';

  // Opciones de tiempo disponibles, de 10 en 10, de 10 min a 2 hs
  var OPCIONES_MIN = [];
  for (var m = 10; m <= 120; m += 10) OPCIONES_MIN.push(m);

  var timerInactividad = null;
  var timerAviso = null;
  var minutosElegidos = null;

  function crearEstilos() {
    var css =
      '#smm-overlay-config,#smm-overlay-aviso{position:fixed;inset:0;background:rgba(11,31,58,.55);' +
      'display:flex;align-items:center;justify-content:center;z-index:999999;font-family:Inter,Arial,sans-serif}' +
      '#smm-caja{background:#fff;border-radius:14px;padding:26px 28px;max-width:360px;width:90%;' +
      'box-shadow:0 12px 40px rgba(0,0,0,.25);text-align:center}' +
      '#smm-caja h3{margin:0 0 10px;color:#0b1f3a;font-size:16px}' +
      '#smm-caja p{margin:0 0 16px;color:#475569;font-size:13px;line-height:1.5}' +
      '#smm-caja select{width:100%;padding:10px;border-radius:8px;border:1px solid #cbd5e1;font-size:14px;margin-bottom:16px}' +
      '#smm-caja button{background:#2563eb;color:#fff;border:none;padding:10px 18px;border-radius:8px;' +
      'font-size:14px;font-weight:600;cursor:pointer;width:100%}' +
      '#smm-caja button:hover{background:#1e4fc4}' +
      '#smm-caja .smm-cancelar{background:#e2e8f0;color:#334155;margin-top:8px}' +
      '#smm-caja .smm-cancelar:hover{background:#cbd5e1}' +
      '#smm-badge{position:fixed;bottom:14px;right:14px;background:#0b1f3a;color:#fff;font-size:11px;' +
      'padding:6px 12px;border-radius:20px;z-index:999998;opacity:.85;font-family:Inter,Arial,sans-serif}';
    var style = document.createElement('style');
    style.textContent = css;
    document.head.appendChild(style);
  }

  function pedirTiempo() {
    var overlay = document.createElement('div');
    overlay.id = 'smm-overlay-config';
    var opts = OPCIONES_MIN.map(function (m) {
      var etiqueta = m < 60 ? m + ' minutos' : (m === 60 ? '1 hora' : (m / 60).toFixed(1).replace('.0', '') + ' horas');
      var sel = m === 30 ? ' selected' : '';
      return '<option value="' + m + '"' + sel + '>' + etiqueta + '</option>';
    }).join('');

    overlay.innerHTML =
      '<div id="smm-caja">' +
      '<h3>⏱️ Cierre automático por inactividad</h3>' +
      '<p>Elegí cada cuánto tiempo sin uso se debe cerrar esta consulta automáticamente, para liberar el sistema y poder entrar desde otro dispositivo.</p>' +
      '<select id="smm-select-tiempo">' + opts + '</select>' +
      '<button id="smm-btn-confirmar">Confirmar</button>' +
      '<button class="smm-cancelar" id="smm-btn-desactivar">No cerrar automáticamente</button>' +
      '</div>';
    document.body.appendChild(overlay);

    document.getElementById('smm-btn-confirmar').onclick = function () {
      minutosElegidos = parseInt(document.getElementById('smm-select-tiempo').value, 10);
      overlay.remove();
      mostrarBadge();
      iniciarVigilancia();
    };
    document.getElementById('smm-btn-desactivar').onclick = function () {
      overlay.remove(); // no se activa el auto-cierre en esta sesión
    };
  }

  function mostrarBadge() {
    var badge = document.createElement('div');
    badge.id = 'smm-badge';
    badge.textContent = '🔒 Auto-cierre: ' + minutosElegidos + ' min de inactividad';
    document.body.appendChild(badge);
  }

  function reiniciarTimer() {
    clearTimeout(timerInactividad);
    clearTimeout(timerAviso);
    if (!minutosElegidos) return;
    var msEspera = minutosElegidos * 60 * 1000;
    // Avisa 60 segundos antes de cerrar
    timerInactividad = setTimeout(mostrarAvisoCierre, Math.max(msEspera - 60000, 0));
  }

  function mostrarAvisoCierre() {
    var overlay = document.createElement('div');
    overlay.id = 'smm-overlay-aviso';
    overlay.innerHTML =
      '<div id="smm-caja">' +
      '<h3>⚠️ Se va a cerrar por inactividad</h3>' +
      '<p>No se detectó actividad. Esta consulta se va a cerrar en <b id="smm-cuenta">60</b> segundos.</p>' +
      '<button id="smm-btn-seguir">Seguir usando el sistema</button>' +
      '</div>';
    document.body.appendChild(overlay);

    var restante = 60;
    var cuenta = document.getElementById('smm-cuenta');
    var intervalo = setInterval(function () {
      restante--;
      if (cuenta) cuenta.textContent = restante;
      if (restante <= 0) {
        clearInterval(intervalo);
        window.location.href = INDEX_URL;
      }
    }, 1000);

    document.getElementById('smm-btn-seguir').onclick = function () {
      clearInterval(intervalo);
      overlay.remove();
      reiniciarTimer();
    };
  }

  function iniciarVigilancia() {
    ['mousemove', 'mousedown', 'keydown', 'touchstart', 'scroll', 'click'].forEach(function (ev) {
      document.addEventListener(ev, reiniciarTimer, { passive: true });
    });
    reiniciarTimer();
  }

  function iniciar() {
    crearEstilos();
    pedirTiempo();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', iniciar);
  } else {
    iniciar();
  }
})();
