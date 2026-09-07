# Manual de Lógica y Circuitos de Funcionamiento — S.I.M-M

Compilado a partir de lectura directa de los archivos reales del sistema (no interpretación ni suposición): `INDEX_SIM.html`, `OCM_Muratori_v14_0_26.html`, `Codigo.gs`, `Panel_Especialistas.html`, `MANUAL_SIMM_v1_0.txt`, y el resto de módulos subidos (GONIO×5, KSM, Lector CUD, Lector Historias, RIDM, RMYT, Agenda de Mensajes, IMDCYMT, SSQ Legajo Integral).

---

## 1. Mapa de módulos

S.I.M-M es un conjunto de aplicaciones HTML de archivo único, todas indexadas desde `INDEX_SIM.html`, que abren una a la otra por link relativo (mismo directorio) o por menú "🔀 Ir a…" dentro de cada programa:

| Módulo | Función | Conecta al backend real |
|---|---|---|
| OCM (`OCM_Muratori_v14_0_26.html`) | Historia clínica, recetario, certificados, RMYT, **Libro Quirúrgico** | Sí |
| Panel de Especialistas | Autoriza especialista por especialista qué módulos comparte el sistema | Sí |
| KSM (Kinesiología) | Trabajo del kinesiólogo derivado | Sí |
| Lector CUD / Lector Historias | Lectura de certificados de discapacidad e historias | Sí |
| RIDM | Registro de pacientes con discapacidad | Sí |
| RMYT Autoevaluación (única y semanal) | El paciente autoevalúa su rehabilitación | Sí |
| GONIO ×5 (CostaBartani, Podoscopia, SeguimientoCientifico, ValgoVaro, y el genérico) | Mediciones/estudios específicos | Sí |
| IMDCYMT | Generador de indicaciones médicas | Sí |
| Agenda de Mensajes | Cola de preguntas/respuestas de pacientes (ver sección 5) | Sí |
| CAM | Herramientas clínicas | Sí |
| SSQ Legajo Integral | Evaluador de seguridad/calidad quirúrgica (en integración — ver documento de revisión aparte) | En diseño |

**Hallazgo importante y verificado**: **todos** estos módulos —sin excepción, lo confirmé línea por línea en cada archivo— apuntan a la **misma y única URL de backend**:

```
https://script.google.com/macros/s/AKfycbwYdoFuU8mQL4YKqufSBQ9AD_yqWEEesLtGzYQjNw2JG1BZIQF7ug80mHB5sLMXhaUpvQ/exec
```

No existe un backend separado por módulo, ni un backend separado "trial" y otro "global". Es un único Google Apps Script, bindeado a una única planilla de Google Sheets, desplegado desde la cuenta `consultas.sism@gmail.com`. Más detalle de esto en la sección 6.

---

## 2. Cómo detecta el sistema dónde está corriendo

`INDEX_SIM.html` (y el resto de los módulos) detectan automáticamente el entorno de red al cargar:

- `192.168.100.247:8080` → modo "🟢 Consultorio (red local)".
- `100.103.9.9:8080` (Tailscale) → acceso remoto seguro fuera del consultorio.
- Hosting en `github.io` (o cualquier host que no sea las dos IP de arriba ni `localhost`/`127.0.0.1`) → modo web público (GitHub Pages).
- `file://` → modo pendrive, sin servidor.

Esto **no cambia el backend de datos** (que siempre es el mismo Apps Script de la sección 1): solo cambia desde dónde se sirven los archivos HTML/JS/CSS del sistema. Es importante no confundir "modo de red" con "sistema trial vs. sistema global" — son dos cosas distintas.

---

## 3. El circuito de datos: cómo se garantiza que nada se cargue en falso

1. Cada módulo lee/escribe usando `fetch()` contra la URL única de la sección 1, con `GET` para lecturas y `POST` (`Content-Type: text/plain;charset=utf-8`, cuerpo JSON) para escrituras.
2. Un token opcional (`APP_TOKEN`, guardado en `localStorage`) se agrega a las llamadas protegidas. Mientras no se configure un token en el backend (`tokenValido_()` en Codigo.gs), el sistema queda en modo "bootstrap" y no exige token — es una decisión de diseño explícita del propio backend, no un descuido.
3. **Cola de sincronización offline** (`ocm_sync_queue` en `localStorage`, funciones `colaLeer/colaGuardar/colaAgregar/colaProcesar`): si una escritura falla por falta de conexión, se guarda en la cola y se reintenta solo al reconectar (evento `online`), al cargar la página, y cada 120.000 ms. Este es el mecanismo real que evita perder datos — no hay ningún campo que "simule" haberse guardado sin intentarlo contra el backend.
4. **DNI como clave universal**: no existe un ID de caso/cirugía separado en ninguna planilla real — todo se busca y guarda por DNI del paciente.
5. **Patrón append-only + alerta automática** (ejemplo real: `RMYT_Autoevaluaciones`, funciones `saveAutoevaluacionRMYT`/`getAutoevaluacionesRMYT`): nunca se pisa una fila anterior, se agrega una nueva y se compara con la anterior para generar alertas automáticas (ej. salto de dolor).

**Excepción real encontrada, y que NO debe tomarse como modelo para código nuevo**: el "Libro Quirúrgico" de OCM (equipo quirúrgico, estado, checklist) se guarda únicamente en `localStorage` del navegador (`libroQX_<dni>`) y nunca se sincroniza al backend. El propio comentario del código lo dice: "EN ESTE NAVEGADOR". Queda documentado acá como una limitación real y conocida del sistema actual, no como algo a replicar.

---

## 4. Circuito de mail

- Todo el correo saliente real se manda con `MailApp.sendEmail` desde la única cuenta que ejecuta el script (`consultas.sism@gmail.com`); la diferenciación por destinatario es solo cosmética vía `replyTo` + `name` (no hay múltiples remitentes reales).
- `enviarRecetaPorMail(d)` (Codigo.gs) es el motor real y genérico: recibe un mail de destino cualquiera, un asunto, un resumen de texto y opcionalmente uno o más PDF en base64, y los manda de verdad — es el que reutilizamos para el contacto por rol del sistema de Seguimiento Posquirúrgico (ver documento de revisión aparte).
- Direcciones reales confirmadas: `consultas.sism@gmail.com` (operativa, ejecuta el script y firma como remitente técnico), `drluismuratori@gmail.com` (replyTo de cara al paciente), `lmuratori@gmail.com` (admin/pruebas).

---

## 5. Circuito de preguntas y respuestas (Cola_Mensajes)

1. Un disparador revisa el Gmail del doctor, etiqueta los mails entrantes y arma una fila nueva en la hoja `Cola_Mensajes` por cada hilo.
2. Cada mail se clasifica automáticamente con IA (Gemini) en `sensible_publico` (discapacidad/CUD/hospital público, gratuito) o `privado` (requiere pago antes de responder), y se genera un borrador de respuesta.
3. El doctor revisa la lista de pendientes (`getMensajesPendientes()`), confirma el pago si corresponde (`confirmarPagoMensaje`), y aprueba el envío (`marcarRespondido`), que responde el **hilo real de Gmail** (`GmailApp.getThreadById`).
4. **Limitación real a tener en cuenta para cualquier integración nueva** (por ejemplo, preguntas desde un formulario web en vez de un mail): `marcarRespondido` necesita un `threadId` de Gmail real. Una pregunta que entra por web sin pasar por Gmail no tiene ese hilo, así que haría falta o bien generar un mail real que sí cree el hilo, o bien una acción hermana que responda por `MailApp.sendEmail` directo.

---

## 6. Circuito de especialistas, autorización y "trial"

Este es el punto que suele generar confusión entre "sistema trial" y "sistema global" — acá está lo que confirmé leyendo el código, no una interpretación:

- Cualquier especialista que abre KSM (o cualquier módulo compartible) y acepta el "peaje" (comisión %) queda registrado en la hoja de especialistas de la **misma planilla única** (sección 1). No hay una planilla ni un backend separado para "los que están en trial".
- Desde `Panel_Especialistas.html` el doctor decide, especialista por especialista: si está autorizado (`Autorizado: SI/PENDIENTE`), qué módulos ve (`investigacion`, `cud`, `grafica`, `plantillas` — nunca `podoscopia`, que está bloqueada de forma permanente y no debe removerse esa restricción bajo ninguna circunstancia) y qué % de peaje personalizado tiene.
- Casi todos los módulos (OCM, KSM, Panel de Especialistas, Lector CUD, Lector Historias, RIDM, Agenda de Mensajes) traen un bloque idéntico de "candado por vencimiento de prueba": al cargar, hacen `fetch(SHEETS_URL+'?action=check_trial')`, y si la respuesta trae `vencido:true`, tapan toda la pantalla con un aviso de "Período de prueba finalizado — contactá al Dr. Muratori".
- **Esto confirma que "trial" y "global" NO son dos sistemas ni dos links distintos** — son dos estados de autorización dentro del **mismo** backend único de la sección 1. "Sistema global" es, en los hechos, ese mismo backend cuando el `check_trial` de esa cuenta/instalación todavía no está vencido (o el doctor la marcó como autorizada de forma permanente).

**Punto que no puedo confirmar con lo que tengo, y te lo marco en vez de adivinar**: la acción `check_trial` (y la función `configurarTrial()` que el propio comentario del código menciona como su contraparte en el backend) **no aparece en el `Codigo.gs` que extraje** del paquete `SIMM_Paquete_v8_17ago2026.zip`. Es decir, todos los módulos cliente ya llaman a esa acción, pero mi copia del backend no la tiene implementada. Dos explicaciones posibles: (a) el backend realmente desplegado en producción es una versión más nueva que la que tengo, con `check_trial`/`configurarTrial` ya agregados, o (b) todavía no se implementó del lado del servidor y por eso hoy esas pantallas de candado simplemente nunca se activan (el `fetch` falla o no trae `vencido`, y el `catch` vacío lo deja pasar silenciosamente). Te lo pregunto directo en el documento de conectores, sección aparte.

---

## 7. Circuito de la Agenda de Pacientes y categorías

`_categoriaAgendaPaciente()` en OCM clasifica cada paciente en la agenda por color/categoría según: discapacidad (con gravedad 0-3), o estado del circuito quirúrgico (`tratamiento`, `decidido`, `observado`, `planificacion`, `programacion`, `seguimiento`), leyendo el estado guardado en `localStorage` (`libroQX_<dni>`, ver limitación de la sección 3). La categoría `'seguimiento'` ya existe hoy y es el gancho natural para colgar ahí el nuevo panel de Seguimiento Posquirúrgico.

---

Este manual se referencia junto con el documento de **Revisión de Integración SSQ Posquirúrgico** (arquitectura propuesta para el nuevo sistema) y con el documento de **Conectores del Sistema — Trial y Global** (los links reales, y la aclaración pendiente sobre `check_trial`).
