# Manual de Uso — S.I.M-M, Versión General
### Para el Dr. Luis Alberto Muratori (M.N. 100.540 / M.P. 9943)

Basado en lo verificado directamente en `INDEX_SIM.html`, `OCM_Muratori_v14_0_26.html`, `Panel_Especialistas.html` y el resto de los módulos reales del sistema.

---

## 1. Cómo entrar al sistema

`INDEX_SIM.html` detecta solo dónde estás parado y te muestra un aviso arriba de todo:

- **🟢 Consultorio (red local)** — estás conectado a la red del consultorio.
- **🟡 Remoto (Tailscale)** — estás afuera, accediendo por la red segura remota.
- **🌐 Acceso temporal (GitHub Pages)** — estás entrando desde la web pública (`https://lmuratori.github.io/smm-acceso-temporal/`).
- **💾 Modo Pen Drive** — estás abriendo los archivos directo desde un pendrive, sin servidor.

No hace falta que elijas nada — el sistema lo detecta solo y te muestra el badge correspondiente arriba a la derecha.

---

## 2. El circuito típico de una consulta

El propio sistema está pensado para este recorrido (así está documentado en el código de `INDEX_SIM.html`):

**CAM → OCM → RIMP**, con **RSP** como referencia constante durante toda la consulta, y las mediciones de **GONIO** según el estudio puntual que corresponda ese día.

- **CAM** (🧠 Consulta Asistida): antecedentes, examen físico, banderas rojas y diagnóstico — la antesala de la consulta.
- **OCM** (🩺 el módulo más usado): recetario, certificados, indicaciones y carrito de recetas — acá vivís la mayor parte del tiempo. Incluye también el Libro Quirúrgico, RMYT y Consultas Online.
- **RIMP** (📋): datos de filiación y antecedentes del paciente.
- **RSP** (📁): historial, evolución y ficha clínica — la referencia que consultás en paralelo.
- **GONIO** (📐 — cinco mediciones distintas): Podoscopia, Valgo/Varo, Costa-Bártani, Ángulos de Frente, y Seguimiento Científico — usás la que corresponda al estudio del día.

## 3. Resto de los módulos, uno por uno

| Módulo | Para qué sirve |
|---|---|
| **RCM** 💰 | Cobros rápidos por Mercado Pago y agenda de turnos |
| **KSM** 🏃 | Evolución de sesiones de kinesiología (este es el que suele abrir un especialista externo — ver el otro manual, "Versión Trial") |
| **EVAL** 📊 | Evaluación semanal del paciente |
| **IMDCYMT Paciente** 📄 | El documento completo que se le entrega al paciente: diagnóstico, medicación, ejercicios, evaluación y directorio |
| **Estadísticas SMM** 📊 | Localidad, exposición laboral OIT, rubros, antecedentes psiquiátricos y COVID — para análisis y presentaciones |
| **Agenda Familiar** 👨‍👩‍👧 | Contactos y fichas familiares, con envío directo a RSP |
| **Lector CUD** 🪪 | Lectura y análisis de Certificados Únicos de Discapacidad (clasificación Ley 22.431) |
| **Seguimiento Carpeta CUD** 📎 | Checklist de trámite + documentos (recetas, formularios de Junta, DNI, transporte) |
| **Isa — Asistente Discapacidad** 🤝 | Chat para que el paciente se oriente en el trámite y mande documentación |
| **Lector Historias** | Transcribe historias clínicas viejas (Word o papel tipeado) a los campos del sistema |
| **RIDM** | Registro de pacientes con discapacidad |
| **Agenda de Mensajes** ✉️ | Cola de preguntas de pacientes, con borrador de respuesta armado por IA para que apruebes antes de mandar |
| **Panel de Especialistas** 🔐 | Ver punto 4 |
| **QR Generador / QR Consulta Virtual** | Generan los links y QR para que un paciente pida una consulta virtual escaneando un cartel (guardia, farmacia, club, etc.) |

**Nota honesta**: encontré que `INDEX_SIM.html` todavía referencia el archivo como `OCM_Muratori_v13.html`, mientras que la versión real que tenés hoy es `OCM_Muratori_v14_0_26.html`. Puede que ya lo hayas corregido en tu copia publicada y esto sea solo un desfasaje del archivo que me pasaste — te lo marco para que lo revises, no lo dejé pasar por las dudas.

## 4. Panel de Especialistas — autorizar a otros profesionales

Desde acá administrás, uno por uno, a cualquier especialista (por ejemplo un kinesiólogo) que haya abierto algún módulo compartible del sistema:

- Ver quién está registrado, y si está **Autorizado**, **Pendiente** o **Bloqueado**.
- Elegir qué módulos ve cada uno: Investigación, Discapacidad/CUD, Gráfica de evolución del dolor, Plantillas — nunca más que eso.
- Poner un % de peaje personalizado por especialista (si no le ponés uno, usa el % general que configures arriba).
- **Podoscopía nunca aparece como opción para compartir** — es de tu uso exclusivo, el sistema no permite habilitarla a nadie más, ni por error.

## 5. Cómo actualizar el sistema

Hay dos lugares distintos según qué estés actualizando (el detalle completo está en el documento "Conectores del Sistema"):

- **Un módulo HTML** (una página del sistema): se sube por GitHub Desktop al repositorio `smm-acceso-temporal`, y se ve solo en `https://lmuratori.github.io/smm-acceso-temporal/` en un par de minutos.
- **El motor (`Codigo.gs`)**: se actualiza entrando a `script.google.com` con la cuenta `consultas.sism@gmail.com`, en el proyecto "S.M.I-M", Implementar → Administrar implementaciones → Nueva versión.

## 6. Buenas prácticas

- Guardá este manual y el paquete completo del sistema en tu USB de respaldo, y también en Drive u OneDrive.
- Cualquier especialista nuevo entra siempre en modo de prueba primero — nunca autorizado automáticamente (ver el manual de Versión Trial para lo que ellos ven).
- Ante cualquier duda sobre qué IA está usando una función puntual del sistema, se puede chequear la hoja `Log_Tokens_IA` de tu planilla — ahí queda el registro real de cada llamada a Gemini.
