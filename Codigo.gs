// === SISTEMA MÉDICO MURATORI — Apps Script API ===
// Esta versión conserva TODO lo que ya tenías (Pacientes, Historia_Clínica, Recetas,
// getAll/search/ping, addPaciente/addReceta/addHistoria) y agrega:
//   - el puente real con RIDM (para que los datos lleguen solos a OCM)
//   - el arreglo del backup, que venía fallando en silencio (mandaba 'backup'
//     y el script no tenía esa acción — por eso OCM decía "OK" sin haber guardado nada)
//   - la puerta lista para kinesiología, todavía sin usar

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName('Recetas') === null) ss.getSheets()[0].setName('Recetas');
  const recetasSheet = ss.getSheetByName('Recetas');
  recetasSheet.getRange(1,1,1,recetasSheet.getLastColumn())
    .setBackground('#6B3FA0').setFontColor('white').setFontWeight('bold');

  if (ss.getSheetByName('Pacientes') === null) {
    const ps = ss.insertSheet('Pacientes');
    const ph = ['ID','Apellido','Nombre','DNI','Fecha_Nac','Obra_Social',
                'N_Afiliado','Plan','Telefono','Email','Domicilio',
                'Lugar_Atencion','Diagnostico_Principal','Estado_Tratamiento',
                'Localidad','Trabajos_JSON','Vivienda_JSON'];
    ps.getRange(1,1,1,ph.length).setValues([ph])
      .setBackground('#1A73E8').setFontColor('white').setFontWeight('bold');
    ps.setFrozenRows(1);
  }
  if (ss.getSheetByName('Historia_Clinica') === null) {
    const hs = ss.insertSheet('Historia_Clinica');
    const hh = ['ID_HC','ID_Paciente','Fecha','Diagnostico','Medicamentos',
                'Estudios_Realizados','Estudios_Pendientes','Indicaciones',
                'Observaciones','Proximo_Control'];
    hs.getRange(1,1,1,hh.length).setValues([hh])
      .setBackground('#0F9D58').setFontColor('white').setFontWeight('bold');
    hs.setFrozenRows(1);
  }
  // NUEVO: hoja de espera para los formularios RIDM que completa el paciente
  if (ss.getSheetByName('RIDM_Pendientes') === null) {
    const rs = ss.insertSheet('RIDM_Pendientes');
    const rh = ['Timestamp','DNI','Apellido','Nombre','Telefono','Email','Datos_JSON','Procesado'];
    rs.getRange(1,1,1,rh.length).setValues([rh])
      .setBackground('#D97706').setFontColor('white').setFontWeight('bold');
    rs.setFrozenRows(1);
  }
  // NUEVO: registro de backups (antes se perdían porque la acción no existía)
  if (ss.getSheetByName('Backups') === null) {
    const bs = ss.insertSheet('Backups');
    bs.getRange(1,1,1,2).setValues([['Timestamp','Resumen_JSON']])
      .setBackground('#64748B').setFontColor('white').setFontWeight('bold');
    bs.setFrozenRows(1);
  }
  // NUEVO: evoluciones de kinesiología — todavía no se usa, queda lista para cuando se sume KSM
  if (ss.getSheetByName('Kinesiologia') === null) {
    const ks = ss.insertSheet('Kinesiologia');
    ks.getRange(1,1,1,5).setValues([['Timestamp','DNI','Paciente','Region','Evolucion_JSON']])
      .setBackground('#16A34A').setFontColor('white').setFontWeight('bold');
    ks.setFrozenRows(1);
  }
  // NUEVO: evaluaciones semanales que completa el PACIENTE (para el médico o el kinesiólogo)
  if (ss.getSheetByName('Evaluaciones_Pacientes') === null) {
    const es = ss.insertSheet('Evaluaciones_Pacientes');
    es.getRange(1,1,1,6).setValues([['Timestamp','DNI','Paciente','Profesional','Region','Valores_JSON']])
      .setBackground('#7C3AED').setFontColor('white').setFontWeight('bold');
    es.setFrozenRows(1);
  }
  // NUEVO: pagos de Mercado Pago confirmados de verdad contra la API de MP (no contra una imagen)
  if (ss.getSheetByName('Pagos_Confirmados') === null) {
    const pc = ss.insertSheet('Pagos_Confirmados');
    const pch = ['Timestamp','Payment_ID','DNI','Tipo_Consulta','Monto','Estado_MP','Atendido'];
    pc.getRange(1,1,1,pch.length).setValues([pch])
      .setBackground('#0EA5E9').setFontColor('white').setFontWeight('bold');
    pc.setFrozenRows(1);
  }
  // NUEVO: registros del Lector CUD, conectados con Pacientes (agenda familiar) e Historia_Clinica
  if (ss.getSheetByName('Discapacidad_CUD') === null) {
    const dc = ss.insertSheet('Discapacidad_CUD');
    const dch = ['Timestamp','DNI','Nombre','Telefono','Fecha_Nac','Diagnostico','Codigos_CIF','Tipo',
                 'Subtipo','Modalidad','Orientacion','Fecha_Emision','Fecha_Vencimiento','Junta',
                 'Firmantes_JSON','Efectores','Tipo_Pension','Trazabilidad'];
    dc.getRange(1,1,1,dch.length).setValues([dch])
      .setBackground('#244C46').setFontColor('white').setFontWeight('bold');
    dc.setFrozenRows(1);
  }
  recetasSheet.setFrozenRows(1);
  SpreadsheetApp.getUi().alert('✅ ¡Hojas configuradas! Ahora: Implementar > Nueva implementación');
}

// ══════════════════════════════════════════════════════════════════
// PARCHE — Token compartido para las acciones que devuelven datos
// masivos/sensibles de pacientes. Antes CUALQUIERA con la URL del Web App
// podía llamar get_vista_consolidada_todos, search, get_historial, etc.
// sin identificarse. No es autenticación por usuario (todas las apps
// comparten el mismo token), pero corta el riesgo de que la URL se filtre
// y cualquiera pueda leer datos de todos los pacientes.
//
// A PROPÓSITO queda "abierto" (sin exigir token) hasta que corras
// configurarAppToken() una vez — así no se rompe nada mientras vas
// actualizando cada app con el token nuevo. Una vez configurado, se exige
// en todas las acciones de la lista de abajo.
// ══════════════════════════════════════════════════════════════════
const ACCIONES_GET_PROTEGIDAS_ = [
  'getAll','search','get_pendientes','get_historial','get_pagos_pendientes',
  'getIMDCYMT','getAutoevaluacionesRMYT','getConsultas','get_pacientes_incompletos',
  'get_vista_consolidada','get_vista_consolidada_todos','getMensajesPendientes',
  'get_ridm_paciente',
  'get_preguntas_kine','get_resumen_facturacion_sistema',
  'get_permisos_especialista','get_lista_especialistas',
  'get_investigacion_paciente','get_discapacidad_paciente','get_plantillas_paciente'
];
function configurarAppToken(){
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('APP_TOKEN');
  if(!token){
    token = Utilities.getUuid();
    props.setProperty('APP_TOKEN', token);
  }
  MailApp.sendEmail('lmuratori@gmail.com', 'SIM-M — Token de acceso de las apps',
    'Este es el token que hay que pegar UNA vez en cada app (OCM, KSM, Panel de Especialistas, ' +
    'Agenda de Mensajes, Colegas) para que sigan funcionando una vez que el token quede exigido:\n\n' +
    token + '\n\nHasta que lo pegues en todas, el sistema sigue funcionando igual que antes (no se corta nada).');
  Logger.log('APP_TOKEN: ' + token);
}
function getAppToken_(){
  return PropertiesService.getScriptProperties().getProperty('APP_TOKEN') || '';
}
// Si todavía no se configuró ningún token (getAppToken_ vacío), no bloquea
// nada — es el modo "todavía no migrado". Una vez configurado, exige que
// coincida exactamente.
function tokenValido_(tokenRecibido){
  const real = getAppToken_();
  if(!real) return true;
  return String(tokenRecibido||'')===String(real);
}

// ══════════════════════════════════════════════════════════════════
// TRIAL — vencimiento controlado desde ESTE backend (no desde el HTML que
// se le entrega a un colega). Por default está DESACTIVADO: no cambia nada
// en el sistema del Dr. Muratori a menos que él mismo corra configurarTrial()
// desde el editor de Apps Script. Pensado para dar una copia de prueba: se
// clona este proyecto (Sheet + Apps Script) para el destinatario, y en ESA
// copia (no en la principal) se activa el trial con fecha de vencimiento.
// Una vez vencido, TODAS las acciones (menos check_trial) quedan bloqueadas
// server-side — no alcanza con editar el HTML del lado del cliente para
// esquivarlo, porque el backend es el que corta.
// ══════════════════════════════════════════════════════════════════
function configurarTrial(fechaVencimientoISO){
  // fechaVencimientoISO: 'AAAA-MM-DD', ej: '2026-09-15'
  const props = PropertiesService.getScriptProperties();
  props.setProperty('TRIAL_ACTIVO', 'true');
  props.setProperty('TRIAL_VENCIMIENTO', fechaVencimientoISO);
  Logger.log('Trial activado. Vence el ' + fechaVencimientoISO + '. Para desactivarlo, correr desactivarTrial().');
}
function desactivarTrial(){
  PropertiesService.getScriptProperties().deleteProperty('TRIAL_ACTIVO');
  PropertiesService.getScriptProperties().deleteProperty('TRIAL_VENCIMIENTO');
  Logger.log('Trial desactivado — el sistema vuelve a funcionar sin límite de fecha.');
}
function estadoTrial_(){
  const props = PropertiesService.getScriptProperties();
  const activo = props.getProperty('TRIAL_ACTIVO')==='true';
  const fechaVenc = props.getProperty('TRIAL_VENCIMIENTO') || '';
  let vencido = false;
  if(activo && fechaVenc){
    const hoy = new Date();
    const venc = new Date(fechaVenc + 'T23:59:59');
    vencido = hoy.getTime() > venc.getTime();
  }
  return {activo, vencido, fecha_vencimiento: fechaVenc};
}

function doGet(e) {
  const action=e.parameter.action||'',sheet=e.parameter.sheet||'',query=e.parameter.query||'';
  if(action==='check_trial') return jsonOut(Object.assign({success:true}, estadoTrial_()));
  const trial = estadoTrial_();
  if(trial.vencido){
    return jsonOut({success:false, error:'trial_vencido', vencido:true, mensaje:'El período de prueba finalizó el '+trial.fecha_vencimiento+'.'});
  }
  if(ACCIONES_GET_PROTEGIDAS_.indexOf(action)>-1 && !tokenValido_(e.parameter.token)){
    return jsonOut({success:false, error:'Token inválido o faltante — pedile al médico el token de acceso'});
  }
  let r;
  if(action==='getAll') r=getAllRows(sheet);
  else if(action==='search') r=searchPatient(query);
  else if(action==='ping') r={success:true,message:'API activa'};
  // ── NUEVO: puente RIDM → OCM ──
  else if(action==='get_pendientes') r=getRidmPendientes();
  else if(action==='get_ridm_paciente') r=getRidmPaciente(e.parameter.dni||'');
  else if(action==='get_historial') r=getHistorialPaciente(e.parameter.dni||'');
  else if(action==='marcar_procesado') r=marcarRidmProcesado(e.parameter.row||'');
  // ── NUEVO: cola de pagos de Mercado Pago confirmados, esperando que el médico los atienda ──
  else if(action==='get_pagos_pendientes') r=getPagosPendientes();
  else if(action==='getIMDCYMT') r=getIMDCYMT(e.parameter.dni||'');
  else if(action==='getAutoevaluacionesRMYT') r=getAutoevaluacionesRMYT(e.parameter.dni||'');
  else if(action==='getConsultas') r=getConsultas(e.parameter.estado||'');
  else if(action==='getTarifas') r=getTarifas();
  else if(action==='get_pacientes_incompletos') r=getPacientesCamposIncompletos();
  else if(action==='get_vista_consolidada') r=getVistaConsolidada(e.parameter.dni||'');
  else if(action==='get_vista_consolidada_todos') r=getVistaConsolidadaTodos();
  else if(action==='getMensajesPendientes') r=getMensajesPendientes();
  else if(action==='get_preguntas_kine') r=getPreguntasKine(e.parameter.dni||'');
  else if(action==='get_peaje_sistema_pct') r=getPeajeSistemaPct();
  else if(action==='get_resumen_facturacion_sistema') r=getResumenFacturacionSistema(e.parameter.profesionalEmail||'', e.parameter.mes||'', e.parameter.origen||'');
  else if(action==='get_lista_quirurgico') r=getListaQuirurgico(e.parameter.tipo||'');
  // ── NUEVO: permisos de especialistas + módulos compartidos (Investigación/CUD/Plantillas) ──
  else if(action==='get_permisos_especialista') r=getPermisosEspecialista(e.parameter.email||'');
  else if(action==='get_lista_especialistas') r=getListaEspecialistas();
  else if(action==='get_investigacion_paciente') r=getInvestigacionPacienteGateado_(e.parameter.dni||'', e.parameter.profesionalEmail||'');
  else if(action==='get_discapacidad_paciente') r=getDiscapacidadPacienteGateado_(e.parameter.dni||'', e.parameter.profesionalEmail||'');
  else if(action==='get_plantillas_paciente') r=getPlantillasPacienteGateado_(e.parameter.dni||'', e.parameter.profesionalEmail||'');
  else r={error:'Acción no reconocida'};
  return jsonOut(r);
}
// ── Igual que ACCIONES_GET_PROTEGIDAS_, pero para las acciones de ESCRITURA
// (guardar/enviar/configurar) que usan las apps de especialistas — KSM,
// Colegas, Panel de Especialistas, Lector CUD y Agenda de Mensajes. A
// PROPÓSITO se dejan afuera de esta lista:
//   - las acciones que usa el propio Dr. Muratori desde OCM (addPaciente,
//     saveConsulta, backup, etc.) — OCM tiene su propio control de acceso
//     físico (es la compu del consultorio) y gatear todo ahí sería
//     repetitivo sin agregar seguridad real.
//   - las acciones que llenan los pacientes desde un link que reciben por
//     mail (saveAutoevaluacionRMYT, evaluacion_submit, kine_pregunta_submit)
//     — el paciente nunca tiene ni va a tener el token, así que gatearlas
//     rompería el sistema para ellos. Por la misma razón queda afuera
//     qr_solicitud_consulta — la carga un paciente anónimo que entró desde
//     un código QR (guardia, comercios, clubes), tampoco tiene token.
//   - registrar_aceptacion_ksm — es el primer registro de un especialista
//     nuevo, que todavía no tiene el token (se lo damos recién después de
//     que el médico lo autorice desde el Panel de Especialistas). El
//     registro en sí no expone datos de pacientes, solo crea un pedido de
//     alta pendiente.
//   - el webhook de Mercado Pago (data.type==='payment') — lo autentica
//     Mercado Pago con su propio formato, no nuestro token, y se resuelve
//     antes de llegar a este chequeo.
const ACCIONES_POST_PROTEGIDAS_ = [
  'generarLinkMeet','kine_evolucion','kine_informe_enviar','kine_pregunta_responder',
  'registrar_facturacion_sistema','enviar_resumen_facturacion_sistema',
  'set_peaje_sistema_pct','set_permisos_especialista',
  'analizar_foto_cud','guardar_registro_cud',
  'analizar_foto_documento_discap',
  'crear_link_pago_mp',
  'asistente_consulta_ia',
  'analizar_texto_paciente','guardar_paciente_historial',
  'confirmarPagoMensaje','marcarRespondido'
];
function doPost(e) {
  const data=JSON.parse(e.postData.contents);

  // ── Notificación de Mercado Pago: no manda nuestro campo 'action', manda 'type' + 'data.id' ──
  if(data.type==='payment' && data.data && data.data.id){
    return jsonOut(confirmarPagoMP(String(data.data.id)));
  }

  const action = data.action || data.accion || '';
  const trial = estadoTrial_();
  if(trial.vencido){
    return jsonOut({success:false, error:'trial_vencido', vencido:true, mensaje:'El período de prueba finalizó el '+trial.fecha_vencimiento+'.'});
  }
  if(ACCIONES_POST_PROTEGIDAS_.indexOf(action)>-1 && !tokenValido_(data.token)){
    return jsonOut({success:false, error:'Token inválido o faltante — pedile al médico el token de acceso'});
  }
  let r;
  if(action==='addPaciente') r=addPaciente(data);
  else if(action==='get_pacientes_incompletos') r=getPacientesCamposIncompletos();
  else if(action==='enviar_recordatorios_campos') r=enviarRecordatorioCamposIncompletos(data);
  else if(action==='addReceta') r=addReceta(data);
  else if(action==='addHistoria') r=addHistoria(data);
  else if(action==='addHistoriaExtendida') r=addHistoriaExtendida(data);
  else if(action==='enviarRecetaPorMail') r=enviarRecetaPorMail(data);
  else if(action==='guardar_registro_cud') r=guardarRegistroCUD(data.registro||data);
  else if(action==='analizar_foto_cud') r=analizarFotoCUD(data);
  else if(action==='analizar_foto_documento_discap') r=analizarFotoDocumentoDiscap(data);
  else if(action==='crear_link_pago_mp') r=crearLinkPagoMP(data);
  else if(action==='asistente_consulta_ia') r=asistenteConsultaIA(data);
  else if(action==='analizar_texto_paciente') r=analizarTextoPaciente(data);
  else if(action==='analizar_foto_historia') r=analizarFotoHistoria(data);
  else if(action==='guardar_paciente_historial') r=guardarPacienteConHistorial(data);
  // ── NUEVO ──
  else if(action==='backup') r=guardarBackup(data);
  else if(action==='ridm_submit') r=guardarRidmSubmit(data);
  else if(action==='ridm_antecedentes') r=guardarRidmAntecedentes(data);
  else if(action==='qr_solicitud_consulta') r=registrarSolicitudQR(data);
  else if(action==='kine_evolucion') r=guardarKineEvolucion(data);
  else if(action==='evaluacion_submit') r=guardarEvaluacionPaciente(data);
  else if(action==='marcar_pago_atendido') r=marcarPagoAtendido(data.payment_id||'');
  else if(action==='saveIMDCYMT') r=saveIMDCYMT(data);
  else if(action==='saveAutoevaluacionRMYT') r=saveAutoevaluacionRMYT(data);
  else if(action==='saveConsulta') r=saveConsulta(data);
  else if(action==='actualizarConsulta') r=actualizarConsulta(data);
  else if(action==='generarRespuestaIA') r=generarRespuestaIA(data);
  else if(action==='saveTarifas') r=saveTarifas(data);
  else if(action==='marcarRespondido') r=marcarRespondido(data);
  else if(action==='generarLinkMeet') r=generarLinkMeet(data);
  else if(action==='confirmarPagoMensaje') r=confirmarPagoMensaje(data);
  else if(action==='kine_pregunta_submit') r=guardarPreguntaKine(data);
  else if(action==='kine_pregunta_responder') r=responderPreguntaKine(data);
  else if(action==='kine_informe_enviar') r=enviarInformeKine(data);
  else if(action==='set_peaje_sistema_pct') r=setPeajeSistemaPct(data);
  else if(action==='registrar_aceptacion_ksm') r=registrarAceptacionKSM(data);
  else if(action==='registrar_facturacion_sistema') r=registrarFacturacionSistema(data);
  else if(action==='enviar_resumen_facturacion_sistema') r=enviarResumenFacturacionSistema(data);
  else if(action==='guardar_item_quirurgico') r=guardarItemQuirurgico(data);
  // ── NUEVO: permisos de especialistas + sincronización de módulos antes locales ──
  else if(action==='set_permisos_especialista') r=setPermisosEspecialista(data);
  else if(action==='sync_empresa_catastro') r=sincronizarEmpresaCatastro(data);
  else if(action==='guardar_cud_libro_discapacidad') r=guardarLibroDiscapacidadCUD(data);
  else if(action==='sync_plantilla') r=sincronizarPlantilla(data);
  else r={error:'Acción no reconocida'};
  return jsonOut(r);
}
function jsonOut(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}

// ── Funciones que ya existían, sin tocar ──
function getAllRows(sheetName){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName(sheetName);
  if(!sheet) return{error:'Hoja no encontrada: '+sheetName};
  const data=sheet.getDataRange().getValues();
  if(data.length<2) return{success:true,data:[]};
  const headers=data[0];
  return{success:true,data:data.slice(1).map(row=>{const obj={};headers.forEach((h,i)=>{obj[h]=row[i];});return obj;})};
}
function searchPatient(query){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName('Pacientes');
  if(!sheet) return{error:'Hoja Pacientes no encontrada'};
  const data=sheet.getDataRange().getValues(),headers=data[0],q=query.toLowerCase();
  return{success:true,data:data.slice(1).filter(row=>row.some(c=>String(c).toLowerCase().includes(q))).map(row=>{const obj={};headers.forEach((h,i)=>{obj[h]=row[i];});return obj;})};
}
// Agrega las columnas nuevas a la hoja Pacientes si todavía no existen
// (para sistemas que ya estaban en uso antes de agosto 2026)
function asegurarColumnasPacientes(sheet){
  let headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const faltantes = ['Localidad','Trabajos_JSON','Vivienda_JSON','Centro_Derivador','RIMP_Enviado','Antecedentes_Generales_JSON'].filter(h=>headers.indexOf(h)===-1);
  faltantes.forEach(h=>{
    sheet.getRange(1, sheet.getLastColumn()+1).setValue(h).setBackground('#1A73E8').setFontColor('white').setFontWeight('bold');
  });
  return sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
}

// ── RIMP: link y envío del formulario de antecedentes (modo=clinico) ──
// Centralizado acá porque lo usan dos caminos distintos: el envío INMEDIATO
// al guardar un paciente marcado como "primera consulta" (addPaciente) y el
// recordatorio semanal a quienes todavía tienen campos vacíos (red de
// seguridad, por si el envío inmediato no se disparó o falló).
function getRimpUrl_(){
  // BUGFIX: tenía "/02_CELULAR/" de cuando se organizaban copias locales por
  // carpeta — pero el sitio publicado en GitHub Pages es plano (sin subcarpetas),
  // así que ese link SIEMPRE daba 404. Cualquier mail mandado antes de este
  // arreglo (RIMP inmediato al guardar "primera consulta", recordatorio semanal)
  // le llegó al paciente con un link roto.
  return 'https://lmuratori.github.io/smm-acceso-temporal/RIDM_Muratori_v1.html'; // GitHub Pages
}
// p.extenso=true → RIMP completo (infancia, vacunas, COVID, vivienda, jornada
// laboral) — se usa para trámites de discapacidad/CUD. Sin ese flag (default)
// se manda la versión breve (alergias, enfermedades, cirugías, medicación,
// deporte/federación, trabajo), pensada para una consulta habitual.
function construirLinkRimpAntecedentes_(p){
  const modo = p.extenso ? 'clinico' : 'clinico_breve';
  return `${getRimpUrl_()}?modo=${modo}&apellido=${encodeURIComponent(p.apellido||'')}&nombre=${encodeURIComponent(p.nombre||'')}&dni=${encodeURIComponent(p.dni||'')}&tel=${encodeURIComponent(p.telefono||'')}`;
}
function enviarLinkRimpAPaciente_(p){
  if(!p || !p.email) throw new Error('Falta el mail del paciente');
  const link = construirLinkRimpAntecedentes_(p);
  const cuerpo = p.extenso ?
    (`Hola ${p.nombre||''} ${p.apellido||''},\n\n` +
    `Para completar tu ficha médica te pedimos que cargues algunos datos adicionales (antecedentes, datos laborales y de vivienda).\n\n` +
    `Por favor, entrá al siguiente link cuando puedas:\n${link}\n\n` +
    `Es rápido y nos ayuda a llevar un mejor seguimiento de tu salud.\n\n` +
    `Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943`)
    :
    (`Hola ${p.nombre||''} ${p.apellido||''},\n\n` +
    `Antes de tu consulta te pedimos que completes algunos datos (alergias, enfermedades, cirugías, medicación habitual y actividad física).\n\n` +
    `Por favor, entrá al siguiente link cuando puedas:\n${link}\n\n` +
    `Es rápido y nos ayuda a prepararnos mejor para tu consulta.\n\n` +
    `Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943`);
  MailApp.sendEmail(p.email, 'Completá tu ficha médica — RIMP', cuerpo, {
    replyTo:'drluismuratori@gmail.com', name:'Dr. Luis Alberto Muratori'
  });
}
// Marca en la hoja Pacientes que a este paciente (por ID) ya se le mandó el
// RIMP, para no volver a mandárselo (ni desde el guardado inmediato ni desde
// el recordatorio semanal) hasta que alguien lo resetee a mano si hiciera falta.
function marcarRimpEnviado_(idPaciente){
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sheet=ss.getSheetByName('Pacientes');
  if(!sheet) return;
  const headers = asegurarColumnasPacientes(sheet);
  const colId = headers.indexOf('ID'), colRimp = headers.indexOf('RIMP_Enviado');
  if(colId===-1 || colRimp===-1) return;
  const data = sheet.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][colId])===String(idPaciente)){
      sheet.getRange(i+1, colRimp+1).setValue(new Date().toLocaleString('es-AR'));
      return;
    }
  }
}

function addPaciente(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName('Pacientes');
  const headers = asegurarColumnasPacientes(sheet);
  const data=sheet.getDataRange().getValues();
  const dniBuscado=String(d.DNI||'');
  const colDNI = headers.indexOf('DNI');
  const colRimpEnviado = headers.indexOf('RIMP_Enviado');

  const valoresNuevos = {
    Apellido:d.Apellido, Nombre:d.Nombre, DNI:dniBuscado, Fecha_Nac:d.Fecha_Nac,
    Obra_Social:d.Obra_Social, N_Afiliado:d.N_Afiliado, Plan:d.Plan, Telefono:d.Telefono,
    Email:d.Email, Domicilio:d.Domicilio, Lugar_Atencion:d.Lugar_Atencion,
    Diagnostico_Principal:d.Diagnostico_Principal, Estado_Tratamiento:d.Estado_Tratamiento||'Activo',
    Localidad:d.Localidad, Trabajos_JSON: d.Trabajos ? JSON.stringify(d.Trabajos) : undefined,
    Vivienda_JSON: d.Vivienda ? JSON.stringify(d.Vivienda) : undefined,
    Centro_Derivador: d.Centro_Derivador,
    // Antecedentes que carga el médico a mano la primera vez (motivo inicial,
    // enfermedades, cirugías, medicación, actividad laboral y deportiva) —
    // quedan en el paciente (no solo en la consulta del día) para que estén
    // disponibles siempre: en CAM, en certificados, en recetas, etc.
    Antecedentes_Generales_JSON: d.Antecedentes_Generales ? JSON.stringify(d.Antecedentes_Generales) : undefined
  };

  let idFinal = '', actualizado = false, yaTeniaRimp = false;

  if(dniBuscado){
    for(let i=1;i<data.length;i++){
      if(String(data[i][colDNI])===dniBuscado){
        const fila = data[i].slice();
        headers.forEach((h,idx)=>{
          if(valoresNuevos[h] !== undefined && valoresNuevos[h] !== ''){ fila[idx] = valoresNuevos[h]; }
        });
        sheet.getRange(i+1,1,1,fila.length).setValues([fila]);
        idFinal = data[i][headers.indexOf('ID')]; actualizado = true;
        yaTeniaRimp = colRimpEnviado>-1 ? !!data[i][colRimpEnviado] : false;
        break;
      }
    }
  }
  if(!actualizado){
    idFinal='PAC-'+String(sheet.getLastRow()).padStart(4,'0');
    const filaNueva = headers.map(h=> h==='ID' ? idFinal : (valoresNuevos[h] || ''));
    sheet.appendRow(filaNueva);
  }

  // ── "Primera consulta": si se marcó explícitamente y el paciente tiene
  // mail y todavía no se le había mandado el RIMP, se lo mandamos ahora
  // mismo — no hay que esperar al recordatorio semanal. Si ya se le había
  // mandado (a este mismo paciente, en un guardado anterior), no se repite,
  // así no se lo satura con el mismo link cada vez que se edita su ficha. ──
  let rimpEnviado = false;
  const mailDestino = d.Email || valoresNuevos.Email;
  if(d.Primera_Consulta && mailDestino && !yaTeniaRimp){
    try{
      enviarLinkRimpAPaciente_({
        apellido: d.Apellido||valoresNuevos.Apellido, nombre: d.Nombre||valoresNuevos.Nombre,
        dni: dniBuscado, telefono: d.Telefono||valoresNuevos.Telefono, email: mailDestino,
        // RIMP completo (vivienda, trabajos, infancia, vacunas, COVID) solo para
        // trámites de discapacidad/CUD — el resto recibe la versión breve.
        extenso: !!d.Discapacidad
      });
      rimpEnviado = true;
      marcarRimpEnviado_(idFinal);
    }catch(e){ /* si falla el envío del mail, no se interrumpe el guardado del paciente */ }
  }

  return{success:true, id:idFinal, actualizado:actualizado, rimpEnviado:rimpEnviado};
}

// ══════════════════════════════════════════════════════════════════
// FASE A (parte 2) — VISTA CONSOLIDADA POR DNI
// Cruza Pacientes + Historia_Clinica + Discapacidad_CUD + antecedentes
// de RIMP (RIDM_Pendientes) en un solo objeto por paciente. Es la base
// que va a usar el futuro Dashboard de estadísticas (Fase B).
// ══════════════════════════════════════════════════════════════════
function sheetToObjects(nombreHoja){
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sheet=ss.getSheetByName(nombreHoja);
  if(!sheet) return [];
  const data=sheet.getDataRange().getValues();
  if(data.length<2) return [];
  const headers=data[0];
  return data.slice(1).map(row=>{
    const obj={}; headers.forEach((h,i)=>{ obj[h]=row[i]; }); return obj;
  });
}

function construirVistaConsolidada(dni, cache){
  cache = cache || {
    pacientes: sheetToObjects('Pacientes'),
    historia: sheetToObjects('Historia_Clinica'),
    cud: sheetToObjects('Discapacidad_CUD'),
    ridm: sheetToObjects('RIDM_Pendientes')
  };
  const dniStr = String(dni||'');

  const pac = cache.pacientes.find(p=>String(p.DNI)===dniStr) || {};
  const historiaPac = cache.historia.filter(h=>String(h.DNI)===dniStr);
  const cudPac = cache.cud.filter(c=>String(c.DNI)===dniStr);
  // Último envío de RIMP para este DNI (donde viven los antecedentes detallados)
  const ridmPac = cache.ridm.filter(r=>String(r.DNI)===dniStr).sort((a,b)=>new Date(b.Timestamp)-new Date(a.Timestamp))[0];
  let antecedentes = {};
  try{ antecedentes = ridmPac ? JSON.parse(ridmPac.Datos_JSON||'{}') : {}; }catch(_){ antecedentes = {}; }

  let trabajos = [];
  try{ trabajos = pac.Trabajos_JSON ? JSON.parse(pac.Trabajos_JSON) : (antecedentes.trabajos||[]); }catch(_){ trabajos = antecedentes.trabajos||[]; }
  let vivienda = {};
  try{ vivienda = pac.Vivienda_JSON ? JSON.parse(pac.Vivienda_JSON) : {}; }catch(_){ vivienda = {}; }
  // Antecedentes que el médico cargó a mano en "+ Nuevo paciente" (motivo inicial,
  // enfermedades, cirugías, medicación, actividad laboral y deportiva) — el "pool"
  // que se guarda una sola vez y de acá en más pueden reusar CAM, certificados, etc.
  let antecedentesGenerales = {};
  try{ antecedentesGenerales = pac.Antecedentes_Generales_JSON ? JSON.parse(pac.Antecedentes_Generales_JSON) : {}; }catch(_){ antecedentesGenerales = {}; }

  return {
    dni: dniStr,
    apellido: pac.Apellido||antecedentes.apellido||'',
    nombre: pac.Nombre||antecedentes.nombre||'',
    fecha_nac: pac.Fecha_Nac||'',
    telefono: pac.Telefono||'',
    email: pac.Email||'',
    obra_social: pac.Obra_Social||'', plan: pac.Plan||'', nro_afiliado: pac.N_Afiliado||'',
    domicilio: pac.Domicilio||'',
    centro_derivador: pac.Centro_Derivador||'',
    localidad: pac.Localidad||antecedentes.localidad||'',
    lugar_atencion: pac.Lugar_Atencion||'',
    estado_tratamiento: pac.Estado_Tratamiento||'',
    trabajos: trabajos,                                    // [{rubro,empresa,puesto,exposicion}]
    vivienda: {
      agua: vivienda.agua || antecedentes.vivienda_agua || '',
      excretas: vivienda.excretas || antecedentes.vivienda_excretas || '',
      calefaccion: vivienda.calefaccion || antecedentes.vivienda_calefaccion || '',
      ventilacion: vivienda.ventilacion || antecedentes.vivienda_ventilacion || ''
    },
    antecedentes: {
      enfermedades: antecedentes.enfermedades || '',        // incluye psiquiátricos (Ansiedad, Depresión, TOC, Psicosis, Insomnio)
      otras_enf: antecedentes.otras_enf || '',
      infancia: antecedentes.infancia || '',
      vacunas: antecedentes.vacunas || '',
      covid: antecedentes.covid || '',
      covid_secuelas: antecedentes.covid_secuelas || '',
      cirugias: antecedentes.cirugias || '',
      medicacion: antecedentes.medicacion || '',
      alergia: antecedentes.alergia || '',                  // "Sí" / "No" / "" (no completado)
      alergia_det: antecedentes.alergia_det || '',
      deporte: antecedentes.deporte || '', deporte_det: antecedentes.deporte_det || '',
      federado: antecedentes.federado || '', federado_det: antecedentes.federado_det || ''
    },
    // Fuente distinta de "antecedentes" de arriba (esos vienen del RIMP que
    // completa el paciente en casa). Estos los carga el médico directamente.
    antecedentes_generales: {
      motivo_inicial: antecedentesGenerales.motivo_inicial || '',
      enfermedades: antecedentesGenerales.enfermedades || '',
      cirugias: antecedentesGenerales.cirugias || '',
      medicacion_habitual: antecedentesGenerales.medicacion_habitual || '',
      actividad_laboral: antecedentesGenerales.actividad_laboral || '',
      actividad_deportiva: antecedentesGenerales.actividad_deportiva || ''
    },
    diagnosticos: historiaPac.map(h=>h.Diagnostico).filter(Boolean),
    // Historia completa (no solo la lista de diagnósticos de arriba) — la usa
    // el Espacio de Consulta Virtual de Agenda de Mensajes para mostrar las
    // consultas anteriores del paciente durante la entrevista.
    historia: historiaPac.map(h=>({
      Fecha: h.Fecha||'', Diagnostico: h.Diagnostico||'', Observaciones: h.Observaciones||'',
      Medicamentos: h.Medicamentos||'', Proximo_Control: h.Proximo_Control||''
    })),
    discapacidad: cudPac.map(c=>({tipo:c.Tipo, subtipo:c.Subtipo, diagnostico:c.Diagnostico, codigos_cif:c.Codigos_CIF}))
  };
}

function getVistaConsolidada(dni){
  if(!dni) return{success:false, error:'Falta el DNI'};
  return{success:true, vista: construirVistaConsolidada(dni)};
}

function getVistaConsolidadaTodos(){
  const cache = {
    pacientes: sheetToObjects('Pacientes'),
    historia: sheetToObjects('Historia_Clinica'),
    cud: sheetToObjects('Discapacidad_CUD'),
    ridm: sheetToObjects('RIDM_Pendientes')
  };
  const dnis = [...new Set(cache.pacientes.map(p=>String(p.DNI)).filter(Boolean))];
  return{success:true, vistas: dnis.map(dni=>construirVistaConsolidada(dni, cache))};
}

// ── Recordatorio de campos esenciales incompletos (Localidad, Trabajos, Vivienda) ──
function getPacientesCamposIncompletos(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName('Pacientes');
  if(!sheet) return{success:true,pacientes:[]};
  const headers = asegurarColumnasPacientes(sheet);
  const data = sheet.getDataRange().getValues();
  const idx = (h)=>headers.indexOf(h);
  const pacientes = [];
  for(let i=1;i<data.length;i++){
    const row = data[i];
    const email = row[idx('Email')];
    if(!email) continue; // sin mail no se le puede mandar el recordatorio
    const faltaLocalidad = !row[idx('Localidad')];
    const faltaTrabajos  = !row[idx('Trabajos_JSON')];
    const faltaVivienda  = !row[idx('Vivienda_JSON')];
    if(faltaLocalidad || faltaTrabajos || faltaVivienda){
      pacientes.push({
        id: row[idx('ID')], apellido: row[idx('Apellido')], nombre: row[idx('Nombre')],
        dni: row[idx('DNI')], telefono: row[idx('Telefono')], email: email,
        falta: [faltaLocalidad?'Localidad':null, faltaTrabajos?'Datos laborales':null, faltaVivienda?'Vivienda y saneamiento':null].filter(Boolean),
        extenso: true // Localidad/Trabajos/Vivienda solo se piden en el RIMP completo
      });
    }
  }
  return{success:true,pacientes};
}

function enviarRecordatorioCamposIncompletos(d){
  const soloIds = Array.isArray(d && d.ids) && d.ids.length ? d.ids : null;
  const resultado = getPacientesCamposIncompletos();
  let pacientes = resultado.pacientes;
  if(soloIds) pacientes = pacientes.filter(p=>soloIds.indexOf(p.id)>-1);

  let enviados = 0, errores = 0;
  pacientes.forEach(p=>{
    try{
      const link = construirLinkRimpAntecedentes_(p);
      const cuerpo =
        `Hola ${p.nombre||''} ${p.apellido||''},\n\n` +
        `Para completar tu ficha médica nos falta cargar: ${p.falta.join(', ')}.\n\n` +
        `Por favor, completá estos datos entrando al siguiente link (te va a llevar directo a la sección que falta):\n${link}\n\n` +
        `Es rápido y nos ayuda a llevar un mejor seguimiento de tu salud.\n\n` +
        `Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943`;
      MailApp.sendEmail(p.email, 'Faltan completar algunos datos de tu ficha médica', cuerpo, {
        replyTo:'drluismuratori@gmail.com', name:'Dr. Luis Alberto Muratori'
      });
      enviados++;
      marcarRimpEnviado_(p.id); // mismo marcador que usa el envío inmediato de "primera consulta"
    }catch(e){ errores++; }
  });
  return{success:true, enviados, errores, total: pacientes.length};
}

// Ejecutar UNA VEZ manualmente desde el editor de Apps Script (▶ Ejecutar) para
// activar el envío automático semanal. No se llama desde el sistema (doPost).
function crearTriggerRecordatorioSemanal(){
  ScriptApp.getProjectTriggers().forEach(t=>{
    if(t.getHandlerFunction()==='enviarRecordatorioCamposIncompletos') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarRecordatorioCamposIncompletos')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
}
function addReceta(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName('Recetas');
  const nReceta='REC-'+new Date().getFullYear()+'-'+String(sheet.getLastRow()).padStart(4,'0');
  sheet.appendRow([nReceta,d.ID_Paciente||'',new Date().toLocaleDateString('es-AR'),d.Med1||'',d.Dosis1||'',d.Med2||'',d.Dosis2||'',d.Med3||'',d.Dosis3||'',d.Diagnostico||'',d.Obra_Social||'',d.N_Afiliado||'','Generada','No']);
  return{success:true,nReceta:nReceta};
}
function addHistoria(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName('Historia_Clinica');
  let headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  let dniCol=headers.indexOf('DNI');
  if(dniCol===-1){
    dniCol=headers.length;
    sheet.getRange(1,dniCol+1).setValue('DNI');
    headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  }
  const id='HC-'+String(sheet.getLastRow()).padStart(4,'0');
  const fila=new Array(headers.length).fill('');
  const set=(nombreCol,valor)=>{const i=headers.indexOf(nombreCol);if(i>-1)fila[i]=valor;};
  set('ID_HC',id);
  set('ID_Paciente',d.ID_Paciente||'');
  set('Fecha',new Date().toLocaleDateString('es-AR'));
  set('Diagnostico',d.Diagnostico||'');
  set('Medicamentos',d.Medicamentos||'');
  set('Estudios_Realizados',d.Estudios_Realizados||'');
  set('Estudios_Pendientes',d.Estudios_Pendientes||'');
  set('Indicaciones',d.Indicaciones||'');
  set('Observaciones',d.Observaciones||'');
  set('Proximo_Control',d.Proximo_Control||'');
  set('DNI',d.DNI||'');
  sheet.appendRow(fila);
  return{success:true,id:id};
}

// ── Funciones NUEVAS ──
function getRidmPendientes(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName('RIDM_Pendientes');
  if(!sheet) return{success:true,pendientes:[]};
  const data=sheet.getDataRange().getValues();
  if(data.length<2) return{success:true,pendientes:[]};
  const headers=data[0];
  const pendientes=data.slice(1).map((row,i)=>{
    const obj={_row:i+2};
    headers.forEach((h,j)=>{obj[h.toLowerCase()]=row[j];});
    return obj;
  }).filter(r=>String(r.procesado).toUpperCase()!=='SI');
  return{success:true,pendientes:pendientes};
}
function marcarRidmProcesado(rowNum){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName('RIDM_Pendientes');
  if(!sheet||!rowNum) return{success:false,error:'Falta hoja o número de fila'};
  const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const col=headers.indexOf('Procesado')+1;
  sheet.getRange(parseInt(rowNum),col).setValue('SI');
  return{success:true};
}
function guardarRidmSubmit(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName('RIDM_Pendientes');
  if(!sheet){
    sheet=ss.insertSheet('RIDM_Pendientes');
    sheet.appendRow(['Timestamp','DNI','Apellido','Nombre','Telefono','Email','Datos_JSON','Procesado']);
  }
  sheet.appendRow([new Date().toISOString(),d.dni||'',d.apellido||'',d.nombre||'',d.tel||'',d.mail||'',JSON.stringify(d.datos||{}),'NO']);
  // ── Filiación rápida: mandar automáticamente por MAIL el link para que el
  // paciente complete sus antecedentes (Parte 2) — antes esto dependía de que
  // el médico se lo copiara/mandara a mano por WhatsApp, que tiene límites de
  // uso. Si no cargó mail en Filiación, no se manda nada (no rompe el guardado,
  // sigue disponible el botón de WhatsApp/copiar link en pantalla). ──
  let mailEnviado = false;
  if(d.mail){
    try{
      enviarLinkRimpAPaciente_({ apellido:d.apellido||'', nombre:d.nombre||'', dni:d.dni||'', telefono:d.tel||'', email:d.mail, extenso:false });
      mailEnviado = true;
    }catch(e){ /* si falla el mail, el registro en RIDM_Pendientes ya quedó guardado igual */ }
  }
  return{success:true,message:'RIDM guardado, en espera de importar a OCM',mail_enviado:mailEnviado};
}
// ── Parte 2 del RIMP (antecedentes clínicos que el paciente completa en casa) ──
// Llega cuando el paciente completa el RIMP en modo=clinico.
// Se guarda igual que un RIDM_Pendientes para que OCM lo muestre con el badge
// y el médico pueda importarlo a la ficha del paciente con un toque.
// El campo datos.modo='clinico' distingue este envío de uno de primera vez.
function guardarRidmAntecedentes(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName('RIDM_Pendientes');
  if(!sheet){
    sheet=ss.insertSheet('RIDM_Pendientes');
    sheet.appendRow(['Timestamp','DNI','Apellido','Nombre','Telefono','Email','Datos_JSON','Procesado']);
  }
  // Buscar si ya hay una fila no procesada para este DNI y actualizarla
  // (para no acumular múltiples envíos de antecedentes del mismo paciente)
  const data=sheet.getDataRange().getValues();
  const headers=data[0];
  const dniColIdx=headers.indexOf('DNI');
  const procColIdx=headers.indexOf('Procesado');
  const datosColIdx=headers.indexOf('Datos_JSON');
  for(let i=1;i<data.length;i++){
    if(String(data[i][dniColIdx])===String(d.dni||'') && String(data[i][procColIdx]).toUpperCase()!=='SI'){
      // Mezclar los antecedentes con los datos que ya están en esa fila
      let prevDatos={};
      try{ prevDatos=JSON.parse(data[i][datosColIdx]||'{}'); }catch(_){}
      const merged=Object.assign({},prevDatos,d.datos||{},{_tipo:'antecedentes+merge'});
      sheet.getRange(i+1,datosColIdx+1).setValue(JSON.stringify(merged));
      return{success:true,message:'Antecedentes agregados al RIDM pendiente existente'};
    }
  }
  // Si no hay fila previa, crear una nueva
  sheet.appendRow([new Date().toISOString(),d.dni||'',d.apellido||'',d.nombre||'',d.tel||'',d.mail||'',JSON.stringify(Object.assign({},d.datos||{},{_tipo:'antecedentes'})),'NO']);
  return{success:true,message:'Antecedentes guardados como nuevo RIDM pendiente'};
}
// ══════════════════════════════════════════════════════════════════
// Trae los antecedentes RIMP de UN paciente puntual, tal cual los mandó
// desde su casa (mismo formato que guardarRidmSubmit/guardarRidmAntecedentes
// escriben en Datos_JSON) — para que OCM se los pueda mostrar al médico
// durante la consulta sin tener que ir a buscarlos a mano a la planilla.
// Antes esto solo se leía de localStorage del propio navegador del médico,
// por eso nunca aparecía: el paciente lo completa en SU celular, no en la
// compu del consultorio.
// ══════════════════════════════════════════════════════════════════
function getRidmPaciente(dni){
  if(!dni) return {success:false, error:'Falta el DNI'};
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sheet=ss.getSheetByName('RIDM_Pendientes');
  if(!sheet) return {success:true, encontrado:false};
  const data=sheet.getDataRange().getValues();
  if(data.length<2) return {success:true, encontrado:false};
  const headers=data[0];
  const filas = data.slice(1).map((row,i)=>{
    const obj={_row:i+2};
    headers.forEach((h,j)=>{obj[h]=row[j];});
    return obj;
  }).filter(r=>String(r.DNI)===String(dni));
  if(!filas.length) return {success:true, encontrado:false};
  filas.sort((a,b)=>new Date(b.Timestamp)-new Date(a.Timestamp));
  // Preferimos la más reciente que todavía no se importó; si ya están todas
  // importadas, igual mostramos la más reciente (por si el médico la quiere
  // volver a consultar).
  const fila = filas.find(r=>String(r.Procesado).toUpperCase()!=='SI') || filas[0];
  let datos={};
  try{ datos=JSON.parse(fila.Datos_JSON||'{}'); }catch(_){}
  const merged = Object.assign({
    dni: fila.DNI, apellido: fila.Apellido, nombre: fila.Nombre,
    tel: fila.Telefono, mail: fila.Email,
    fecha_envio: (fila.Timestamp instanceof Date) ? fila.Timestamp.toLocaleString('es-AR') : fila.Timestamp
  }, datos, { _row: fila._row });
  return {success:true, encontrado:true, procesado: String(fila.Procesado).toUpperCase()==='SI', datos: merged};
}
function guardarBackup(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName('Backups');
  if(!sheet){
    sheet=ss.insertSheet('Backups');
    sheet.appendRow(['Timestamp','Resumen_JSON']);
  }
  sheet.appendRow([d.timestamp||new Date().toISOString(),d.data||'']);
  return{success:true,message:'Backup guardado'};
}
function guardarKineEvolucion(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName('Kinesiologia');
  if(!sheet){
    sheet=ss.insertSheet('Kinesiologia');
    // Vals_JSON en vez de Evolucion_JSON para que RSP y get_historial lo lean igual
    sheet.appendRow(['Timestamp','DNI','Paciente','Region','Vals_JSON']);
  }
  // Si la hoja ya existe con el header viejo, agregar Vals_JSON si falta
  const headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  if(!headers.includes('DNI')){
    // Hoja vieja sin DNI — la reconvertimos (solo headers, no borra datos)
    sheet.getRange(1,1,1,5).setValues([['Timestamp','DNI','Paciente','Region','Vals_JSON']]);
  }
  sheet.appendRow([
    new Date().toISOString(),
    d.dni||'',
    d.paciente||'',
    d.region||'',
    JSON.stringify(d.vals||{})
  ]);
  return{success:true,message:'Evolución de kinesiología guardada'};
}
function guardarEvaluacionPaciente(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName('Evaluaciones_Pacientes');
  if(!sheet){
    sheet=ss.insertSheet('Evaluaciones_Pacientes');
    sheet.appendRow(['Timestamp','DNI','Paciente','Profesional','Region','Valores_JSON']);
  }
  sheet.appendRow([new Date().toISOString(),d.dni||'',d.paciente||'',d.profesional||'',d.region||'',JSON.stringify(d.vals||{})]);
  return{success:true,message:'Evaluación del paciente guardada'};
}
function getHistorialPaciente(dni){
  if(!dni) return{success:true,historial:[],evaluaciones:[],kinesiologia:[]};
  const ss=SpreadsheetApp.getActiveSpreadsheet();

  // ID del paciente, para encontrar también filas viejas que no tienen DNI propio
  let idPaciente=null;
  const pSheet=ss.getSheetByName('Pacientes');
  if(pSheet){
    const pData=pSheet.getDataRange().getValues();
    const pHeaders=pData[0];
    const dniCol=pHeaders.indexOf('DNI');
    const idCol=pHeaders.indexOf('ID');
    const paciente=pData.slice(1).find(row=>String(row[dniCol])===String(dni));
    if(paciente) idPaciente=paciente[idCol];
  }

  // 1) Historia clínica: por DNI directo (filas nuevas) o por ID_Paciente (filas viejas, sin DNI)
  let historial=[];
  const hSheet=ss.getSheetByName('Historia_Clinica');
  if(hSheet){
    const hData=hSheet.getDataRange().getValues();
    const hHeaders=hData[0];
    const dniColH=hHeaders.indexOf('DNI');
    const idPacCol=hHeaders.indexOf('ID_Paciente');
    historial=hData.slice(1).filter(row=>
      (dniColH>-1 && String(row[dniColH])===String(dni)) ||
      (idPaciente!==null && String(row[idPacCol])===String(idPaciente))
    ).map(row=>{const obj={};hHeaders.forEach((h,i)=>{obj[h]=row[i];});return obj;});
  }

  // 2) Evaluaciones que completa el paciente (médico o kinesiólogo) — vía DNI directo
  const evaluaciones = leerPorDni_('Evaluaciones_Pacientes', dni);

  // 3) Evoluciones que carga el kinesiólogo en KSM — vía DNI directo
  const kinesiologia = leerPorDni_('Kinesiologia', dni);

  return{success:true,historial:historial,evaluaciones:evaluaciones,kinesiologia:kinesiologia};
}
// Helper: lee una hoja que tiene columna DNI y devuelve las filas de ese paciente.
// Normaliza Evolucion_JSON → vals (para compatibilidad con filas guardadas antes del fix de header)
function leerPorDni_(nombreHoja, dni){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName(nombreHoja);
  if(!sheet) return[];
  const data=sheet.getDataRange().getValues();
  if(data.length<2) return[];
  const headers=data[0];
  const dniCol=headers.indexOf('DNI');
  if(dniCol===-1) return[];
  return data.slice(1).filter(row=>String(row[dniCol])===String(dni))
    .map(row=>{
      const obj={};
      headers.forEach((h,i)=>{obj[h]=row[i];});
      // Normalizar: si viene Evolucion_JSON pero no Vals_JSON, exponerlo como Valores_JSON también
      if(obj.Evolucion_JSON && !obj.Valores_JSON) obj.Valores_JSON=obj.Evolucion_JSON;
      if(obj.Vals_JSON && !obj.Valores_JSON) obj.Valores_JSON=obj.Vals_JSON;
      return obj;
    });
}

// ══════════════════════════════════════════════════════════════════
// PAGOS DE MERCADO PAGO — verificación real contra la API de MP
// (Apps Script no puede leer el header x-signature de la notificación,
// así que NO confiamos en el contenido del webhook: lo usamos solo como
// aviso de "revisá el pago X", y le preguntamos a Mercado Pago directo,
// con nuestro propio Access Token, si ese pago es real y está aprobado)
// ══════════════════════════════════════════════════════════════════
function confirmarPagoMP(paymentId){
  try{
    const token = PropertiesService.getScriptProperties().getProperty('MP_ACCESS_TOKEN');
    if(!token) return{success:false,error:'Falta configurar el Access Token de Mercado Pago — ejecutar configurarTokenMP() una vez desde el editor'};
    const resp = UrlFetchApp.fetch('https://api.mercadopago.com/v1/payments/'+encodeURIComponent(paymentId), {
      method: 'get',
      headers: {'Authorization': 'Bearer '+token},
      muteHttpExceptions: true
    });
    const pago = JSON.parse(resp.getContentText());
    if(!pago || !pago.id) return{success:false,error:'Mercado Pago no devolvió datos para ese pago'};
    const ref = parsearExternalReference_(pago.external_reference||'');
    guardarPagoConfirmado({
      payment_id: String(pago.id),
      dni: ref.dni,
      tipo: ref.tipo,
      monto: pago.transaction_amount||0,
      estado: pago.status||''
    });
    return{success:true, estado:pago.status, confirmado: pago.status==='approved'};
  }catch(e){
    return{success:false, error:'Error confirmando el pago: '+e.message};
  }
}
// Convención a futuro para RCM: el link de pago debe incluir external_reference
// con el formato "DNI|TIPO_DE_CONSULTA" (ej: "27444812|video15") para poder
// saber de quién es el pago y qué consulta corresponde. Hoy RCM todavía no
// lo manda — falta verlo y ajustarlo para que esto quede conectado de punta a punta.
// ══════════════════════════════════════════════════════════════════
// RCM — genera el link de pago real de Mercado Pago (Checkout Pro) para
// mandarle al paciente. Hasta ahora el sistema solo podía CONFIRMAR un
// pago ya hecho (confirmarPagoMP arriba) pero no existía nada que creara
// el link en sí — por eso a los pacientes nunca les aparecía la opción
// de pagar: el médico tenía que generarlo a mano afuera del sistema.
// external_reference queda como "DNI|tipo" para que, cuando MP avise el
// pago por webhook, parsearExternalReference_ lo pueda relacionar con el
// paciente correcto.
// ══════════════════════════════════════════════════════════════════
function crearLinkPagoMP(d){
  const token = (PropertiesService.getScriptProperties().getProperty('MP_ACCESS_TOKEN')||'').trim();
  if(!token) return {success:false, error:'Falta configurar el Access Token de Mercado Pago — ejecutar configurarTokenMP() una vez, o cargar MP_ACCESS_TOKEN en Propiedades del script'};
  const monto = parseFloat(d.monto);
  if(!monto || monto<=0) return {success:false, error:'Falta indicar el monto a cobrar'};
  const concepto = d.concepto || 'Consulta médica — Dr. Muratori';
  const urlBase = 'https://lmuratori.github.io/smm-acceso-temporal/';
  const payload = {
    items: [{ title: concepto, quantity: 1, unit_price: monto, currency_id: 'ARS' }],
    back_urls: { success: urlBase, failure: urlBase, pending: urlBase },
    auto_return: 'approved',
    notification_url: ScriptApp.getService().getUrl(),
    external_reference: (d.dni||'') + '|' + (d.tipo||'teleconsulta'),
    statement_descriptor: 'DR MURATORI'
  };
  if(d.mail){ payload.payer = { email: d.mail }; }
  try{
    const resp = UrlFetchApp.fetch('https://api.mercadopago.com/checkout/preferences', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    const raw = resp.getContentText();
    if(!raw){
      return {success:false, error:'Mercado Pago respondió vacío (código HTTP '+code+'). Puede ser un problema con el Access Token guardado en Propiedades del script — revisar que esté completo (formato APP_USR-...), sin espacios ni saltos de línea de más.'};
    }
    let data;
    try{ data = JSON.parse(raw); }catch(eParse){
      return {success:false, error:'Mercado Pago devolvió una respuesta no reconocible (código HTTP '+code+'): ' + raw.substring(0,300)};
    }
    if(!data.init_point){
      return {success:false, error: 'Código HTTP '+code+' — ' + (data.message || 'Mercado Pago no devolvió el link de pago') + (data.cause ? ' — '+JSON.stringify(data.cause) : '') };
    }
    return {success:true, link: data.init_point, preference_id: data.id};
  }catch(err){
    return {success:false, error:'No se pudo conectar con Mercado Pago: ' + err};
  }
}
// ══════════════════════════════════════════════════════════════════
// DIAGNÓSTICO — para correr UNA VEZ manualmente desde el editor (▶ Ejecutar)
// cuando crearLinkPagoMP da error 403. Prueba el mismo Access Token contra
// un endpoint simple (GET /users/me) para aislar si el bloqueo es del
// token/cuenta en general, o específico de crear preferencias de pago.
// Manda el resultado por mail, igual que testGeminiDirecto().
// ══════════════════════════════════════════════════════════════════
// Versión ampliada (agosto 2026) — Mercado Pago Soporte (ticket WCS-48270)
// pidió puntualmente: la traza HTTP completa de una llamada a GET
// /users/me (headers, URL y método, con el token oculto), la respuesta
// completa incluyendo HEADERS (no solo el body), confirmación de que el
// token se manda en el header Authorization: Bearer (no por URL), y si hay
// una lista de IPs permitidas configurada en la app. Esta función arma esa
// traza para las dos llamadas y la manda por mail, lista para pegar en la
// respuesta del ticket. Lo único que Apps Script NO puede confirmar es la
// lista de IPs — eso se revisa a mano en developers.mercadopago.com, dentro
// de la app "S-I-M-M" → Configuración de seguridad.
function diagnosticarTokenMP(){
  const token = (PropertiesService.getScriptProperties().getProperty('MP_ACCESS_TOKEN')||'').trim();
  if(!token){
    MailApp.sendEmail('lmuratori@gmail.com', 'Test MP — ERROR', 'Falta configurar MP_ACCESS_TOKEN en Propiedades del script.');
    return;
  }
  const tokenOculto = token.substring(0,12) + '...' + token.substring(token.length-4);
  let cuerpo =
    'Access Token usado (primeros/últimos caracteres): ' + tokenOculto + '\n' +
    'Confirmado: el token se envía SIEMPRE en el header "Authorization: Bearer <token>", nunca como parámetro en la URL (ver el código de crearLinkPagoMP/doPost, sección Mercado Pago).\n\n';

  // ── GET /users/me ──
  try{
    const url = 'https://api.mercadopago.com/users/me';
    const headersEnviados = { 'Authorization': 'Bearer ' + tokenOculto };
    const resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true
    });
    const codigo = resp.getResponseCode();
    const texto = resp.getContentText();
    let headersResp = {};
    try{ headersResp = resp.getAllHeaders(); }catch(eh){ headersResp = {error: String(eh)}; }
    cuerpo += '=== GET /users/me ===\n' +
      'cURL equivalente (token oculto):\n' +
      'curl -X GET "' + url + '" -H "Authorization: Bearer ' + tokenOculto + '"\n\n' +
      'Código HTTP: ' + codigo + '\n' +
      'Headers de la respuesta:\n' + JSON.stringify(headersResp, null, 2) + '\n' +
      'Body de la respuesta: ' + (texto || '(vacía)') + '\n\n';
  }catch(e){
    cuerpo += '=== GET /users/me ===\nNo se pudo ni llamar: ' + e + '\n\n';
  }

  // ── POST /checkout/preferences ──
  try{
    const url2 = 'https://api.mercadopago.com/checkout/preferences';
    const payload2 = JSON.stringify({ items: [{ title: 'Prueba diagnóstico', quantity: 1, unit_price: 100, currency_id: 'ARS' }] });
    const resp2 = UrlFetchApp.fetch(url2, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: payload2,
      muteHttpExceptions: true
    });
    const codigo2 = resp2.getResponseCode();
    const texto2 = resp2.getContentText();
    let headersResp2 = {};
    try{ headersResp2 = resp2.getAllHeaders(); }catch(eh){ headersResp2 = {error: String(eh)}; }
    cuerpo += '=== POST /checkout/preferences (prueba mínima) ===\n' +
      'cURL equivalente (token oculto):\n' +
      'curl -X POST "' + url2 + '" -H "Authorization: Bearer ' + tokenOculto + '" -H "Content-Type: application/json" -d \'' + payload2 + '\'\n\n' +
      'Código HTTP: ' + codigo2 + '\n' +
      'Headers de la respuesta:\n' + JSON.stringify(headersResp2, null, 2) + '\n' +
      'Body de la respuesta: ' + (texto2 || '(vacía)') + '\n\n';
  }catch(e){
    cuerpo += '=== POST /checkout/preferences ===\nNo se pudo ni llamar: ' + e + '\n\n';
  }

  cuerpo +=
    '── Pendiente de revisar A MANO (esto Apps Script no lo puede confirmar) ──\n' +
    'Mercado Pago Soporte también preguntó si hay una lista de IPs permitidas configurada en la aplicación. ' +
    'Revisar en developers.mercadopago.com → la app "S-I-M-M" → Configuración de seguridad. ' +
    'Si hay una lista cargada, la IP desde la que corre Google Apps Script cambia y no es fija — eso solo explicaría el 403 vacío en TODAS las llamadas, sin importar el endpoint, que es exactamente lo que está pasando.';

  MailApp.sendEmail('lmuratori@gmail.com', 'Test MP — diagnóstico Access Token (ampliado, para WCS-48270)', cuerpo);
}
function parsearExternalReference_(ref){
  const partes = String(ref).split('|');
  return{dni: partes[0]||'', tipo: partes[1]||''};
}
function guardarPagoConfirmado(info){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName('Pagos_Confirmados');
  if(!sheet){
    sheet=ss.insertSheet('Pagos_Confirmados');
    sheet.appendRow(['Timestamp','Payment_ID','DNI','Tipo_Consulta','Monto','Estado_MP','Atendido']);
  }
  // Si Mercado Pago reintenta la misma notificación, actualizamos en vez de duplicar
  const data=sheet.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][1])===String(info.payment_id)){
      sheet.getRange(i+1,6).setValue(info.estado);
      return;
    }
  }
  sheet.appendRow([new Date().toISOString(), info.payment_id, info.dni, info.tipo, info.monto, info.estado, 'NO']);
}
function getPagosPendientes(){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName('Pagos_Confirmados');
  if(!sheet) return{success:true,pendientes:[]};
  const data=sheet.getDataRange().getValues();
  if(data.length<2) return{success:true,pendientes:[]};
  const headers=data[0];
  const pendientes=data.slice(1).map((row,i)=>{
    const obj={_row:i+2};
    headers.forEach((h,j)=>{obj[h.toLowerCase()]=row[j];});
    return obj;
  }).filter(r=>r.estado_mp==='approved' && String(r.atendido).toUpperCase()!=='SI');
  return{success:true,pendientes:pendientes};
}
function marcarPagoAtendido(paymentId){
  const ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName('Pagos_Confirmados');
  if(!sheet||!paymentId) return{success:false,error:'Falta hoja o Payment_ID'};
  const data=sheet.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][1])===String(paymentId)){
      sheet.getRange(i+1,7).setValue('SI');
      return{success:true};
    }
  }
  return{success:false,error:'No se encontró ese Payment_ID'};
}
// ══════════════════════════════════════════════════════════════════
// IMDCYMT PACIENTE — Guarda y recupera el documento completo del
// paciente (para que funcione desde el celular, donde no hay localStorage de OCM)
// ══════════════════════════════════════════════════════════════════
function saveIMDCYMT(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName('IMDCYMT_Docs');
  if(!sheet){
    sheet=ss.insertSheet('IMDCYMT_Docs');
    sheet.appendRow(['Timestamp','DNI','Datos_JSON']);
    sheet.setFrozenRows(1);
  }
  const datos=d.datos||{};
  const dni=d.dni||datos.paciente&&datos.paciente.dni||'';
  // Actualizar si ya existe una fila para este DNI
  const rows=sheet.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][1])===String(dni)){
      sheet.getRange(i+1,1,1,3).setValues([[new Date().toISOString(),dni,JSON.stringify(datos)]]);
      return{success:true,message:'IMDCYMT actualizado'};
    }
  }
  sheet.appendRow([new Date().toISOString(),dni,JSON.stringify(datos)]);
  return{success:true,message:'IMDCYMT guardado'};
}
function getIMDCYMT(dni){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName('IMDCYMT_Docs');
  if(!sheet||!dni) return{success:false,datos:null};
  const rows=sheet.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][1])===String(dni)){
      try{
        const datos=JSON.parse(rows[i][2]||'{}');
        return{success:true,datos:datos};
      }catch(_){ return{success:false,datos:null}; }
    }
  }
  return{success:true,datos:null};
}
// Guarda el Access Token de Mercado Pago en un lugar seguro (las Propiedades
// del script) que nunca queda en este archivo ni se sube a la MyCloud.
// Pasos: 1) pegá tu token donde dice PEGÁ_AQUÍ  2) seleccioná esta función
// en el desplegable de arriba  3) tocá ▶ Ejecutar  4) borrá el token de aquí
// si querés (ya quedó guardado, no se vuelve a necesitar en el código).
function configurarTokenMP(){
  const token = 'PEGÁ_AQUÍ_TU_ACCESS_TOKEN_DE_MERCADO_PAGO';
  PropertiesService.getScriptProperties().setProperty('MP_ACCESS_TOKEN', token);
  Logger.log('Token guardado correctamente.');
}

// ══════════════════════════════════════════════════════════════════
// CONSULTAS ONLINE PAGAS — Cola de consultas + respuesta IA
// ══════════════════════════════════════════════════════════════════
const CONSULTAS_HEADERS = ['ID','DNI','Paciente','Tipo','Precio','Pregunta',
  'Estado','Fecha','Payment_ID','Respuesta_Draft','Respuesta_Final','Fecha_Respuesta'];

function sheetConsultas_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sh=ss.getSheetByName('Consultas_Online');
  if(!sh){
    sh=ss.insertSheet('Consultas_Online');
    sh.getRange(1,1,1,CONSULTAS_HEADERS.length).setValues([CONSULTAS_HEADERS])
      .setBackground('#7c3aed').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function saveConsulta(d){
  const sh=sheetConsultas_();
  const id='CONS-'+new Date().getTime();
  sh.appendRow([id,d.dni||'',d.paciente||'',d.tipo||'texto',d.precio||0,
    d.pregunta||'','pendiente_pago',new Date().toISOString(),'','','','']);
  return{success:true,id:id};
}

function getConsultas(estado){
  const sh=sheetConsultas_();
  const rows=sh.getDataRange().getValues();
  if(rows.length<2) return{success:true,consultas:[]};
  const headers=rows[0];
  let lista=rows.slice(1).map(row=>{
    const obj={};headers.forEach((h,i)=>{obj[h.toLowerCase()]=row[i];});return obj;
  });
  if(estado) lista=lista.filter(c=>c.estado===estado);
  return{success:true,consultas:lista.reverse()};
}

// ── REEMPLAZADA por el patch de peaje de sistema: ver más abajo el
//    bloque "PARCHE — Peaje de uso del sistema" para la versión final,
//    que agrega el registro automático del peaje al cerrar una consulta.

function generarRespuestaIA(d){
  const apiKey=PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if(!apiKey) return{success:false,error:'Falta configurar CLAUDE_API_KEY — ejecutar configurarClaudeAPI() una vez desde el editor'};
  const sistema=`Sos el Dr. Luis Alberto Muratori, especialista en Ortopedia, Traumatología y Cirugía de Columna (M.N. 100.540, M.P. 9943).
Respondé consultas médicas en español argentino, de forma clara, precisa y profesional.
Usá términos comprensibles para el paciente. Si la consulta requiere examen físico, aclaralo.
Si podés dar una respuesta útil sin verlo, hacelo.

Datos del paciente:
- Nombre: ${d.paciente||'No especificado'}
- Diagnóstico: ${d.diagnostico||'No especificado'}
- Medicación actual: ${d.medicacion||'No especificada'}
- Región tratada: ${d.region||'No especificada'}
- Historial relevante: ${d.historial||'Sin datos adicionales'}`;

  try{
    const resp=UrlFetchApp.fetch('https://api.anthropic.com/v1/messages',{
      method:'post',
      headers:{'Content-Type':'application/json','x-api-key':apiKey,'anthropic-version':'2023-06-01'},
      payload:JSON.stringify({model:'claude-sonnet-4-6',max_tokens:1000,
        system:sistema,messages:[{role:'user',content:d.pregunta}]}),
      muteHttpExceptions:true
    });
    const data=JSON.parse(resp.getContentText());
    if(data.content&&data.content[0]&&data.content[0].text){
      // Guardar el draft en la consulta
      if(d.id) actualizarConsulta({id:d.id,respuesta_draft:data.content[0].text,estado:'en_proceso'});
      return{success:true,respuesta:data.content[0].text};
    }
    return{success:false,error:'Respuesta inesperada de la API: '+JSON.stringify(data)};
  }catch(e){
    return{success:false,error:'Error de red: '+e.message};
  }
}

// ── TARIFAS ── Guardadas en una hoja de configuración
function getTarifas(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sh=ss.getSheetByName('Configuracion');
  if(!sh) return{success:true,tarifas:null}; // usa las default del frontend
  const rows=sh.getDataRange().getValues();
  for(const row of rows){
    if(row[0]==='tarifas_consulta'){
      try{ return{success:true,tarifas:JSON.parse(row[1]||'{}')}; }catch(_){}
    }
  }
  return{success:true,tarifas:null};
}

function saveTarifas(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sh=ss.getSheetByName('Configuracion');
  if(!sh){
    sh=ss.insertSheet('Configuracion');
    sh.appendRow(['clave','valor']);
    sh.setFrozenRows(1);
  }
  const rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(rows[i][0]==='tarifas_consulta'){
      sh.getRange(i+1,2).setValue(JSON.stringify(d.tarifas));
      return{success:true};
    }
  }
  sh.appendRow(['tarifas_consulta',JSON.stringify(d.tarifas)]);
  return{success:true};
}

// ── Configurar la clave de Claude (ejecutar una sola vez desde el editor) ──
function configurarClaudeAPI(){
  const key='PEGÁ_AQUÍ_TU_CLAUDE_API_KEY';
  PropertiesService.getScriptProperties().setProperty('CLAUDE_API_KEY',key);
  Logger.log('Claude API key guardada correctamente.');
}

// ═══════════════════════════════════════════════════════════════════
// NUEVO — Puente con CAM (Consulta Asistida Muratori, ex-Portada Excel)
// Agregado para integrar el razonamiento clinico (CIE-10, segmento,
// banderas rojas, EVA, etc.) sin romper nada de lo que ya funcionaba.
// ═══════════════════════════════════════════════════════════════════

// Guarda una entrada de Historia_Clinica con TODAS las columnas nuevas.
// Si una columna no existe todavia en la hoja, la crea sola (no hay que
// tocar la hoja a mano). Los campos viejos (Diagnostico, Medicamentos,
// etc.) siguen funcionando igual que con addHistoria().
function addHistoriaExtendida(d){
  const ss=SpreadsheetApp.getActiveSpreadsheet(), sheet=ss.getSheetByName('Historia_Clinica');
  let headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];

  // Columnas nuevas que puede traer CAM. Si falta alguna en la hoja, se agrega sola.
  const COLUMNAS_NUEVAS = [
    'Centro_Atencion','Medico_Tratante','Derivado_De','Area_Derivacion','Motivo_Eleccion',
    'Consentimiento_Datos','Consentimiento_Fecha',
    'Circunstancia_Tipo','Circunstancia_Cobertura',
    'Codigo_CIE10','Segmento_Corporal','Tipo_Patologia','Lateralidad',
    'Antecedentes_JSON','Antecedentes_Quirurgicos','Examen_Fisico',
    'Banderas_Rojas','Alerta_Urgente',
    'Peso','Talla','IMC','TA_Sistolica','TA_Diastolica','EVA_Dolor',
    'Farmaco','Dosis','Farmaco_Magistral','Farmacia_Derivada',
    'Kinesico','Kine_Centro','Planilla_Seguimiento','Terapias_Especiales',
    'Autocuidado','Modalidad_Entrenamiento',
    'Estudios_Solicitados',
    'Cirugia_Programada','Tipo_Procedimiento','Fecha_Cirugia_Programada',
    'Kinesiologo_Responsable','Tipo_Actividad_Fisica','Gym_Instructor',
    'Indica_Ortesis','Tipo_Ortesis','Indica_Plantillas','Material_Plantilla',
    'Perfil_Correccion','Discrepancia_Longitud','Lado','Largo_Plantilla',
    'Suplemento_Viscoelastico',
    'Se_Entrega_Certificado','Tipo_Certificado','Reposo_Desde','Reposo_Hasta','Dias_Reposo',
    'Estado_Caso','N_Control','Redaccion_Historia_Clinica'
  ];
  let huboColumnaNueva=false;
  COLUMNAS_NUEVAS.forEach(col=>{
    if(headers.indexOf(col)===-1){
      sheet.getRange(1,headers.length+1).setValue(col);
      headers.push(col);
      huboColumnaNueva=true;
    }
  });
  if(huboColumnaNueva){
    sheet.getRange(1,1,1,headers.length).setBackground('#0F9D58').setFontColor('white').setFontWeight('bold');
  }

  const id='HC-'+String(sheet.getLastRow()).padStart(4,'0');
  const fila=new Array(headers.length).fill('');
  const set=(nombreCol,valor)=>{const i=headers.indexOf(nombreCol);if(i>-1)fila[i]=valor||'';};

  // Campos clasicos (compatibilidad con lo que ya lee OCM/RSP)
  set('ID_HC', id);
  set('ID_Paciente', d.ID_Paciente||'');
  set('Fecha', new Date().toLocaleDateString('es-AR'));
  set('Diagnostico', d.Diagnostico||'');
  set('Medicamentos', d.Medicamentos||d.Farmaco||'');
  set('Estudios_Realizados', d.Estudios_Realizados||'');
  set('Estudios_Pendientes', d.Estudios_Solicitados||'');
  set('Indicaciones', d.Indicaciones||'');
  set('Observaciones', d.Observaciones||'');
  set('Proximo_Control', d.Proximo_Control||'');
  set('DNI', d.DNI||'');

  // Campos nuevos (razonamiento clinico de CAM)
  set('Centro_Atencion', d.Centro_Atencion);
  set('Medico_Tratante', d.Medico_Tratante);
  set('Derivado_De', d.Derivado_De);
  set('Area_Derivacion', d.Area_Derivacion);
  set('Motivo_Eleccion', d.Motivo_Eleccion);
  set('Consentimiento_Datos', d.Consentimiento_Datos);
  set('Consentimiento_Fecha', d.Consentimiento_Fecha);
  set('Circunstancia_Tipo', d.Circunstancia_Tipo);
  set('Circunstancia_Cobertura', d.Circunstancia_Cobertura);
  set('Codigo_CIE10', d.Codigo_CIE10);
  set('Segmento_Corporal', d.Segmento_Corporal);
  set('Tipo_Patologia', d.Tipo_Patologia);
  set('Lateralidad', d.Lateralidad);
  set('Antecedentes_JSON', d.Antecedentes_JSON ? JSON.stringify(d.Antecedentes_JSON) : '');
  set('Antecedentes_Quirurgicos', d.Antecedentes_Quirurgicos);
  set('Examen_Fisico', d.Examen_Fisico);
  set('Banderas_Rojas', d.Banderas_Rojas);
  set('Alerta_Urgente', d.Alerta_Urgente);
  set('Peso', d.Peso);
  set('Talla', d.Talla);
  set('IMC', d.IMC);
  set('TA_Sistolica', d.TA_Sistolica);
  set('TA_Diastolica', d.TA_Diastolica);
  set('EVA_Dolor', d.EVA_Dolor);
  set('Farmaco', d.Farmaco);
  set('Dosis', d.Dosis);
  set('Farmaco_Magistral', d.Farmaco_Magistral);
  set('Farmacia_Derivada', d.Farmacia_Derivada);
  set('Kinesico', d.Kinesico);
  set('Kine_Centro', d.Kine_Centro);
  set('Planilla_Seguimiento', d.Planilla_Seguimiento);
  set('Terapias_Especiales', d.Terapias_Especiales);
  set('Autocuidado', d.Autocuidado);
  set('Modalidad_Entrenamiento', d.Modalidad_Entrenamiento);
  set('Estudios_Solicitados', d.Estudios_Solicitados);
  set('Cirugia_Programada', d.Cirugia_Programada);
  set('Tipo_Procedimiento', d.Tipo_Procedimiento);
  set('Fecha_Cirugia_Programada', d.Fecha_Cirugia_Programada);
  set('Kinesiologo_Responsable', d.Kinesiologo_Responsable);
  set('Tipo_Actividad_Fisica', d.Tipo_Actividad_Fisica);
  set('Gym_Instructor', d.Gym_Instructor);
  set('Indica_Ortesis', d.Indica_Ortesis);
  set('Tipo_Ortesis', d.Tipo_Ortesis);
  set('Indica_Plantillas', d.Indica_Plantillas);
  set('Material_Plantilla', d.Material_Plantilla);
  set('Perfil_Correccion', d.Perfil_Correccion);
  set('Discrepancia_Longitud', d.Discrepancia_Longitud);
  set('Lado', d.Lado);
  set('Largo_Plantilla', d.Largo_Plantilla);
  set('Suplemento_Viscoelastico', d.Suplemento_Viscoelastico);
  set('Se_Entrega_Certificado', d.Se_Entrega_Certificado);
  set('Tipo_Certificado', d.Tipo_Certificado);
  set('Reposo_Desde', d.Reposo_Desde);
  set('Reposo_Hasta', d.Reposo_Hasta);
  set('Dias_Reposo', d.Dias_Reposo);
  set('Estado_Caso', d.Estado_Caso||'Activo');
  set('N_Control', d.N_Control);
  set('Redaccion_Historia_Clinica', d.Redaccion_Historia_Clinica);

  sheet.appendRow(fila);
  return {success:true, id:id};
}

// Envia por MAIL (no WhatsApp, a proposito: el medico eligio mail para no
// recibir una respuesta inmediata del paciente durante la consulta) la
// receta/documentacion + un mini-resumen para que el paciente sepa que
// tiene y que debe hacer. Requiere que el paciente tenga Email cargado.
//
// d = {
//   Email: 'paciente@mail.com',
//   Nombre_Paciente: 'Juan Perez',
//   Asunto: 'Resumen de tu consulta - Dr. Muratori',
//   Mini_Resumen: 'texto en lenguaje simple para el paciente',
//   Receta_Texto: 'texto completo de la receta (o vacio si va como PDF adjunto)',
//   Pdf_Base64: 'opcional, base64 de un PDF ya generado en OCM',
//   Pdf_Nombre: 'opcional, nombre del archivo adjunto'
// }
function enviarRecetaPorMail(d){
  if(!d.Email) return {success:false, error:'El paciente no tiene mail cargado'};

  const cuerpo =
    'Hola ' + (d.Nombre_Paciente||'') + ',\n\n' +
    'Este es un resumen de tu consulta de hoy con el Dr. Muratori:\n\n' +
    (d.Mini_Resumen||'') + '\n\n' +
    (d.Receta_Texto ? ('Indicaciones / Receta:\n' + d.Receta_Texto + '\n\n') : '') +
    'Ante cualquier duda, comunicate por los canales habituales del consultorio.\n\n' +
    'Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943';

  const opciones = { replyTo: 'drluismuratori@gmail.com', name: 'Dr. Luis Alberto Muratori' };
  // Soporta dos formatos: un solo PDF (Pdf_Base64/Pdf_Nombre, usado por la
  // Receta médica y la documentación quirúrgica) o varios PDF juntos en un
  // mismo mail (Pdfs: [{base64,nombre}, ...], usado por el envío de
  // Indicaciones Médicas tildadas). Antes solo se soportaba el primero, así
  // que el envío múltiple mandaba el mail SIN ningún PDF adjunto aunque
  // avisara "enviado" — corregido agosto 2026.
  if(d.Pdf_Base64){
    const bytes = Utilities.base64Decode(d.Pdf_Base64);
    const blob = Utilities.newBlob(bytes, 'application/pdf', d.Pdf_Nombre||'documentacion_consulta.pdf');
    opciones.attachments = [blob];
  } else if(d.Pdfs && d.Pdfs.length){
    opciones.attachments = d.Pdfs.map(function(p, i){
      const bytes = Utilities.base64Decode(p.base64);
      return Utilities.newBlob(bytes, 'application/pdf', p.nombre || ('documento_'+(i+1)+'.pdf'));
    });
  }

  try{
    MailApp.sendEmail(d.Email, d.Asunto||'Resumen de tu consulta - Dr. Muratori', cuerpo, opciones);
    return {success:true};
  }catch(e){
    return {success:false, error:String(e)};
  }
}

// ═══════════════════════════════════════════════════════════════════
// NUEVO — Lector CUD: guarda el registro de discapacidad, lo valida
// cruzado contra Pacientes (agenda familiar) y lo vuelca en Historia_Clinica.
// El Lector CUD manda el campo "accion" (no "action") y envuelve el
// registro dentro de {registro:{...}} - eso ya se resuelve en doPost.
// ═══════════════════════════════════════════════════════════════════
function guardarRegistroCUD(reg){
  if(!reg || !reg.dni) return {success:false, error:'Falta DNI en el registro del CUD'};
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) Guardar el registro completo en su propia hoja
  let dcSheet = ss.getSheetByName('Discapacidad_CUD');
  if(!dcSheet){
    dcSheet = ss.insertSheet('Discapacidad_CUD');
    dcSheet.appendRow(['Timestamp','DNI','Nombre','Telefono','Fecha_Nac','Diagnostico','Codigos_CIF','Tipo',
                        'Subtipo','Modalidad','Orientacion','Fecha_Emision','Fecha_Vencimiento','Junta',
                        'Firmantes_JSON','Efectores','Tipo_Pension','Trazabilidad']);
  }
  dcSheet.appendRow([
    new Date().toISOString(), reg.dni||'', reg.nombre||'', reg.telefono||'', reg.fnac||'',
    reg.diagnostico||'', reg.codigos||'', reg.tipo||'', reg.subtipo||'', reg.modalidad||'',
    reg.orientacion||'', reg.emision||'', reg.vencimiento||'', reg.junta||'',
    JSON.stringify(reg.firmantes||[]), reg.efectores||'', reg.tipoPension||'', reg.trazabilidad||''
  ]);

  // 2) Validacion cruzada con la Agenda familiar (Pacientes): si el DNI no existe
  //    todavia ahi, se crea un registro basico para que quede conectado; si ya
  //    existe, se actualiza el telefono si vino uno nuevo (sin pisar otros datos)
  const pSheet = ss.getSheetByName('Pacientes');
  let pacienteId = '';
  if(pSheet){
    const pData = pSheet.getDataRange().getValues();
    const pHeaders = pData[0];
    const dniCol = pHeaders.indexOf('DNI');
    const idCol = pHeaders.indexOf('ID');
    const telCol = pHeaders.indexOf('Telefono');
    let filaEncontrada = -1;
    for(let i=1;i<pData.length;i++){
      if(String(pData[i][dniCol])===String(reg.dni)){ filaEncontrada=i; break; }
    }
    if(filaEncontrada>-1){
      pacienteId = pData[filaEncontrada][idCol];
      if(reg.telefono && !pData[filaEncontrada][telCol] && telCol>-1){
        pSheet.getRange(filaEncontrada+1, telCol+1).setValue(reg.telefono);
      }
    } else {
      const nuevoId = addPaciente({
        Apellido:'', Nombre:reg.nombre||'', DNI:reg.dni, Fecha_Nac:reg.fnac||'',
        Telefono:reg.telefono||'', Diagnostico_Principal:reg.diagnostico||'',
        Estado_Tratamiento:'Discapacidad (CUD)'
      });
      pacienteId = nuevoId.id;
    }
  }

  // 3) Volcar tambien en Historia_Clinica, para que quede en la misma linea de
  //    tiempo que el resto de las consultas de ese paciente
  const hSheet = ss.getSheetByName('Historia_Clinica');
  if(hSheet){
    const idHc = 'HC-'+String(hSheet.getLastRow()).padStart(4,'0');
    const resumenCud = `Registro CUD: ${reg.tipo||''} ${reg.subtipo?'('+reg.subtipo+')':''}. `+
      `Vigencia: ${reg.emision||'?'} a ${reg.vencimiento||'?'}. Junta: ${reg.junta||'-'}.`;
    hSheet.appendRow([
      idHc, pacienteId||reg.dni, new Date().toLocaleDateString('es-AR'),
      reg.diagnostico||'', '', '', '', resumenCud, '', ''
    ]);
  }

  return {success:true, message:'Registro CUD guardado y conectado con Pacientes e Historia_Clinica', paciente_id: pacienteId};
}

// ══════════════════════════════════════════════════════════════════
// LECTOR CUD — Gemini Vision. El frontend (COL_Lector_CUD.html) manda
// la foto del certificado de discapacidad en base64; esto le pide a
// Gemini que extraiga los datos en el mismo formato que espera
// guardarRegistroCUD, para que el médico solo tenga que revisar y
// confirmar antes de guardar (nunca se guarda automáticamente).
// ══════════════════════════════════════════════════════════════════
function analizarFotoCUD(d){
  if(!d || !d.foto_base64) return {success:false, error:'Falta la foto'};
  const mime = d.mime_type || 'image/jpeg';
  const prompt = 'Sos un asistente que transcribe certificados únicos de discapacidad (CUD) argentinos a partir de una foto. ' +
    'Mirá la imagen y extraé EXACTAMENTE estos campos, en español, tal cual figuran impresos (si un campo no se ve o no figura, dejalo como cadena vacía "", no inventes datos):\n' +
    '- dni (solo números, sin puntos)\n' +
    '- nombre (Apellido, Nombre)\n' +
    '- telefono\n' +
    '- fnac (fecha de nacimiento, formato DD/MM/AAAA)\n' +
    '- diagnostico\n' +
    '- codigos (códigos CIF/CIE si figuran, separados por coma)\n' +
    '- tipo (ej: Motora, Visual, Auditiva, Mental, Visceral, Múltiple)\n' +
    '- subtipo\n' +
    '- modalidad (ej: Permanente, Transitorio)\n' +
    '- orientacion\n' +
    '- emision (fecha de emisión, DD/MM/AAAA)\n' +
    '- vencimiento (fecha de vencimiento, DD/MM/AAAA)\n' +
    '- junta (junta evaluadora / organismo emisor)\n' +
    '- efectores\n' +
    '- tipoPension (si figura pensión no contributiva u otra)\n' +
    '- trazabilidad (número de trámite/expediente si figura)\n\n' +
    'Devolvé SOLO un JSON válido, sin texto adicional ni markdown, con exactamente esas 15 claves (todas como texto).';

  try{
    const resp = UrlFetchApp.fetch(getGeminiUrl_(), {
      method:'post', contentType:'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [
          { text: prompt },
          { inline_data: { mime_type: mime, data: d.foto_base64 } }
        ]}],
        generationConfig: {
          thinkingConfig: { thinkingLevel: 'minimal' }
        }
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(resp.getContentText());
    registrarUsoTokens_('analizarFotoCUD', 'gemini-3.6-flash', data.usageMetadata);
    if(!data.candidates || !data.candidates.length){
      return {success:false, error:'Gemini no devolvió resultado — ' + (data.error ? data.error.message : 'respuesta vacía') };
    }
    let textoIA = data.candidates[0].content.parts[0].text;
    textoIA = textoIA.replace(/```json|```/g,'').trim();
    const registro = JSON.parse(textoIA);
    return {success:true, registro};
  }catch(err){
    return {success:false, error:'No se pudo leer la foto automáticamente: ' + err};
  }
}

// ═══════════════════════════════════════════════════════════════════
// NUEVO — Lector genérico de documentos de respaldo de discapacidad
// (estudios/informes que avalan el CUD, informes mensuales de
// prestadores). A diferencia de analizarFotoCUD (que espera ESPECÍFICAMENTE
// un CUD), este extrae solo lo básico que hace falta para cualquier
// papel: de qué es, cuándo y qué institución lo hizo — así se puede usar
// para fotografiar y cargar rápido cualquier informe de respaldo.
// ═══════════════════════════════════════════════════════════════════
function analizarFotoDocumentoDiscap(d){
  if(!d || !d.foto_base64) return {success:false, error:'Falta la foto'};
  const mime = d.mime_type || 'image/jpeg';
  const prompt = 'Sos un asistente que transcribe informes/estudios/certificados médicos argentinos (en el marco de un trámite ' +
    'de discapacidad — CUD, pensión no contributiva, etc.) a partir de una foto. ' +
    'Mirá la imagen y extraé EXACTAMENTE estos campos, en español (si un campo no se ve o no figura, dejalo como cadena vacía "", no inventes datos):\n' +
    '- tipo_documento (ej: "Informe kinésico", "Estudio de laboratorio", "Certificado neurológico", etc. — una frase corta describiendo qué es)\n' +
    '- fecha (fecha del documento, formato DD/MM/AAAA)\n' +
    '- institucion (institución, profesional o efector que lo emitió)\n' +
    '- resumen (resumen breve del contenido/hallazgo principal, 1-2 oraciones)\n\n' +
    'Devolvé SOLO un JSON válido, sin texto adicional ni markdown, con exactamente esas 4 claves (todas como texto).';

  try{
    const resp = UrlFetchApp.fetch(getGeminiUrl_(), {
      method:'post', contentType:'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [
          { text: prompt },
          { inline_data: { mime_type: mime, data: d.foto_base64 } }
        ]}],
        generationConfig: {
          thinkingConfig: { thinkingLevel: 'minimal' }
        }
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(resp.getContentText());
    registrarUsoTokens_('analizarFotoDocumentoDiscap', 'gemini-3.6-flash', data.usageMetadata);
    if(!data.candidates || !data.candidates.length){
      return {success:false, error:'Gemini no devolvió resultado — ' + (data.error ? data.error.message : 'respuesta vacía') };
    }
    let textoIA = data.candidates[0].content.parts[0].text;
    textoIA = textoIA.replace(/```json|```/g,'').trim();
    const registro = JSON.parse(textoIA);
    return {success:true, registro};
  }catch(err){
    return {success:false, error:'No se pudo leer la foto automáticamente: ' + err};
  }
}

// ═══════════════════════════════════════════════════════════════════
// NUEVO — Lector de Historias Clínicas viejas (Word/papel, texto pegado):
// el médico pega el texto tal cual estaba en el Word, la IA propone los
// datos de filiación + una lista de consultas con su fecha real, y recién
// después de revisar se guarda — mismo criterio que el Lector CUD, pero
// de texto en vez de foto (no hace falta Vision para esto).
// ═══════════════════════════════════════════════════════════════════
function analizarTextoPaciente(d){
  if(!d || !d.texto) return {success:false, error:'Falta el texto a analizar'};
  const prompt = 'Sos un asistente que transcribe historias clínicas argentinas viejas (escritas originalmente en Word o a mano y ' +
    'luego tipeadas) a partir del texto tal cual fue copiado del documento. El texto puede tener errores de tipeo, mayúsculas ' +
    'sueltas, o letras mal codificadas (por ejemplo "A±OS" en vez de "AÑOS", o "AG?ERO" en vez de "AGÜERO") — interpretalas ' +
    'igual, corrigiendo solo esos problemas de codificación, sin inventar ni completar datos que no estén.\n\n' +
    'Extraé EXACTAMENTE estos campos de FILIACIÓN (si un dato no está, dejalo como cadena vacía "", no inventes nada):\n' +
    '- apellido\n- nombre\n- dni (solo números)\n- edad (tal como figura, ej: "54 años")\n' +
    '- fecha_nac (SOLO si figura explícita en el texto, formato DD/MM/AAAA — NO calcules una fecha de nacimiento a partir de la edad, eso sería inventar un dato)\n' +
    '- obra_social (si el texto dice explícitamente que no tiene, particular, o no mutualizado, dejalo igual vacío — eso se marca aparte)\n' +
    '- plan\n- nro_afiliado\n- telefono\n' +
    '- email (si falta el signo @ antes del dominio, por ejemplo "nombreGMAIL.COM", corregilo asumiendo que se perdió el @, pero marcá "email_dudoso": true en ese caso)\n' +
    '- domicilio\n- localidad (si se puede inferir de la dirección)\n' +
    '- centro_derivador (institución, colega o médico que derivó al paciente, SOLO si se menciona explícitamente en el texto)\n\n' +
    'Devolvé SOLO un JSON válido, sin texto adicional ni markdown, con exactamente estas claves: apellido, nombre, dni, edad, ' +
    'fecha_nac, obra_social, plan, nro_afiliado, telefono, email, email_dudoso, domicilio, localidad, centro_derivador.\n\n' +
    'TEXTO A ANALIZAR:\n' + d.texto;

  try{
    const resp = UrlFetchApp.fetch(getGeminiUrl_(), {
      method:'post', contentType:'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [ { text: prompt } ]}],
        generationConfig: { thinkingConfig: { thinkingLevel: 'minimal' } }
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(resp.getContentText());
    registrarUsoTokens_('analizarTextoPaciente', 'gemini-3.6-flash', data.usageMetadata);
    if(!data.candidates || !data.candidates.length){
      return {success:false, error:'Gemini no devolvió resultado — ' + (data.error ? data.error.message : 'respuesta vacía') };
    }
    let textoIA = data.candidates[0].content.parts[0].text;
    textoIA = textoIA.replace(/```json|```/g,'').trim();
    const registro = JSON.parse(textoIA);
    return {success:true, registro};
  }catch(err){
    return {success:false, error:'No se pudo analizar el texto automáticamente: ' + err};
  }
}

// ══════════════════════════════════════════════════════════════════
// Asistente genérico de texto con IA — usado por la Consulta Guiada Beta
// (voz + IA): redactar en prosa clínica lo que el médico va dictando,
// sugerir maniobras de examen físico según el motivo/relato, y armar el
// resumen para el paciente. El prompt completo lo arma el front (cada
// tarea necesita instrucciones bien distintas), acá solo se lo pasamos
// a Gemini tal cual y devolvemos el texto de respuesta.
// ══════════════════════════════════════════════════════════════════
function asistenteConsultaIA(d){
  if(!d || !d.prompt) return {success:false, error:'Falta el texto a procesar'};
  try{
    const resp = UrlFetchApp.fetch(getGeminiUrl_(), {
      method:'post', contentType:'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: d.prompt }] }],
        generationConfig: { thinkingConfig: { thinkingLevel: 'minimal' } }
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(resp.getContentText());
    registrarUsoTokens_('asistenteConsultaIA', 'gemini-3.6-flash', data.usageMetadata);
    if(!data.candidates || !data.candidates.length){
      return {success:false, error:'Gemini no devolvió resultado — ' + (data.error ? data.error.message : 'respuesta vacía')};
    }
    const texto = data.candidates[0].content.parts[0].text || '';
    return {success:true, texto: texto.trim()};
  }catch(err){
    return {success:false, error:'No se pudo conectar con la IA: ' + err};
  }
}

// Misma idea que analizarTextoPaciente(), pero para cuando el médico tiene
// una FOTO de la ficha vieja (papel o Word impreso) en vez de poder pegar el
// texto — por ejemplo, sacada con la cámara web de la PC o el celular. Usa
// Vision (igual que analizarFotoCUD) y devuelve EXACTAMENTE el mismo esquema
// de campos que analizarTextoPaciente(), así el mismo formulario de revisión
// sirve para las dos vías de entrada.
function analizarFotoHistoria(d){
  if(!d || !d.foto_base64) return {success:false, error:'Falta la foto'};
  const mime = d.mime_type || 'image/jpeg';
  const prompt = 'Sos un asistente que transcribe historias clínicas argentinas viejas (fichas en papel o Word impreso) a partir ' +
    'de una foto. Puede haber letra manuscrita, mala calidad de escaneo o reflejos — hacé tu mejor lectura, pero si un dato no se ' +
    'llega a leer con confianza, dejalo como cadena vacía "" en vez de adivinar.\n\n' +
    'Extraé EXACTAMENTE estos campos de FILIACIÓN (si un dato no está o no se lee, dejalo como cadena vacía "", no inventes nada):\n' +
    '- apellido\n- nombre\n- dni (solo números)\n- edad (tal como figura, ej: "54 años")\n' +
    '- fecha_nac (SOLO si figura explícita en la foto, formato DD/MM/AAAA — NO calcules una fecha de nacimiento a partir de la edad, eso sería inventar un dato)\n' +
    '- obra_social (si el texto dice explícitamente que no tiene, particular, o no mutualizado, dejalo igual vacío — eso se marca aparte)\n' +
    '- plan\n- nro_afiliado\n- telefono\n' +
    '- email (si falta el signo @ antes del dominio, por ejemplo "nombreGMAIL.COM", corregilo asumiendo que se perdió el @, pero marcá "email_dudoso": true en ese caso)\n' +
    '- domicilio\n- localidad (si se puede inferir de la dirección)\n' +
    '- centro_derivador (institución, colega o médico que derivó al paciente, SOLO si se menciona explícitamente)\n\n' +
    'Devolvé SOLO un JSON válido, sin texto adicional ni markdown, con exactamente estas claves: apellido, nombre, dni, edad, ' +
    'fecha_nac, obra_social, plan, nro_afiliado, telefono, email, email_dudoso, domicilio, localidad, centro_derivador.';

  try{
    const resp = UrlFetchApp.fetch(getGeminiUrl_(), {
      method:'post', contentType:'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [
          { text: prompt },
          { inline_data: { mime_type: mime, data: d.foto_base64 } }
        ]}],
        generationConfig: { thinkingConfig: { thinkingLevel: 'minimal' } }
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(resp.getContentText());
    registrarUsoTokens_('analizarFotoHistoria', 'gemini-3.6-flash', data.usageMetadata);
    if(!data.candidates || !data.candidates.length){
      return {success:false, error:'Gemini no devolvió resultado — ' + (data.error ? data.error.message : 'respuesta vacía') };
    }
    let textoIA = data.candidates[0].content.parts[0].text;
    textoIA = textoIA.replace(/```json|```/g,'').trim();
    const registro = JSON.parse(textoIA);
    return {success:true, registro};
  }catch(err){
    return {success:false, error:'No se pudo leer la foto automáticamente: ' + err};
  }
}

// Guarda el paciente (upsert por DNI, reutilizando addPaciente) y, si vino
// una lista de historial, agrega una fila POR CADA entrada en Historia_Clinica
// con su FECHA REAL (a diferencia de addHistoria(), que siempre usa la fecha
// de hoy — acá el médico puede estar cargando una consulta de años atrás).
function guardarPacienteConHistorial(d){
  // Normaliza a las claves con mayúscula inicial que espera addPaciente(),
  // aceptando también las claves en minúscula que devuelve la IA (apellido,
  // nombre, dni, etc.) — así el front-end puede mandar cualquiera de las dos.
  const pick=(a,b)=> (d[a]!==undefined && d[a]!==null && d[a]!=='') ? d[a] : d[b];
  const dPaciente = {
    Apellido: pick('Apellido','apellido'), Nombre: pick('Nombre','nombre'), DNI: pick('DNI','dni'),
    Fecha_Nac: pick('Fecha_Nac','fecha_nac'), Obra_Social: pick('Obra_Social','obra_social'),
    N_Afiliado: pick('N_Afiliado','nro_afiliado'), Plan: pick('Plan','plan'),
    Telefono: pick('Telefono','telefono'), Email: pick('Email','email'),
    Domicilio: pick('Domicilio','domicilio'), Lugar_Atencion: pick('Lugar_Atencion','lugar_atencion'),
    Localidad: pick('Localidad','localidad'), Estado_Tratamiento: d.Estado_Tratamiento || 'Activo',
    Centro_Derivador: pick('Centro_Derivador','centro_derivador'),
    // Si se tilda "primera consulta" en el Lector de Fichas, addPaciente() manda
    // el RIMP de antecedentes ahora mismo (en vez de esperar al recordatorio semanal).
    Primera_Consulta: !!(d.Primera_Consulta || d.primera_consulta),
    Discapacidad: !!(d.Discapacidad || d.discapacidad)
  };
  const resultadoPaciente = addPaciente(dPaciente);
  let historiasGuardadas = 0;
  if(Array.isArray(d.historial) && d.historial.length){
    const ss=SpreadsheetApp.getActiveSpreadsheet(), sheet=ss.getSheetByName('Historia_Clinica');
    if(sheet){
      let headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
      let dniCol=headers.indexOf('DNI');
      if(dniCol===-1){
        dniCol=headers.length;
        sheet.getRange(1,dniCol+1).setValue('DNI');
        headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
      }
      d.historial.forEach(h=>{
        if(!h || !h.texto) return;
        const id='HC-'+String(sheet.getLastRow()).padStart(4,'0');
        const fila=new Array(headers.length).fill('');
        const set=(nombreCol,valor)=>{const i=headers.indexOf(nombreCol);if(i>-1)fila[i]=valor;};
        set('ID_HC',id);
        set('Fecha', h.fecha || new Date().toLocaleDateString('es-AR'));
        set('Observaciones', h.texto);
        set('DNI', d.DNI||d.dni||'');
        sheet.appendRow(fila);
        historiasGuardadas++;
      });
    }
  }
  return {success:true, paciente:resultadoPaciente, historiasGuardadas};
}

function probarPermisoMail(){
  MailApp.sendEmail('lmuratori@gmail.com', 'Prueba de permisos SMM', 'Si recibis esto, el permiso ya esta autorizado.');
}

// ══════════════════════════════════════════════════════════════════
// PARCHE — Autoevaluación Semanal RMYT
// A diferencia de IMDCYMT (que guarda un solo documento por paciente y lo
// pisa cada vez), acá CADA autoevaluación semanal se agrega como una fila
// nueva — necesitamos el historial completo, no solo la última, para poder
// graficar la evolución semana a semana.
// ══════════════════════════════════════════════════════════════════
function saveAutoevaluacionRMYT(d){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('RMYT_Autoevaluaciones');
  if(!sheet){
    sheet = ss.insertSheet('RMYT_Autoevaluaciones');
    sheet.appendRow(['Timestamp','DNI','Segmento','EVA_Reposo','EVA_Movimiento',
      'Neurodinamico','Incapacidad_Pct','Recuperacion_Pct','Escala_Nombre',
      'Escala_Pct_Discapacidad','Cumplimiento_Prom','Alerta','Tipo_Dolor',
      'Medicacion_Tipo','Medicacion_Detalle','Efectos_Adversos','Intolerancias',
      'Gastritis','Datos_JSON']);
    sheet.setFrozenRows(1);
  }
  const datos = d.datos || {};
  const dni = d.dni || '';
  if(!dni) return {success:false, error:'Falta DNI del paciente'};

  const cump = datos.cumplimiento || {};
  const valoresCump = [cump.medicacion, cump.ejercicios, cump.kinesio, cump.actividad_guiada]
    .filter(v => v !== null && v !== undefined);
  const cumplimientoProm = valoresCump.length
    ? Math.round(valoresCump.reduce((a,b)=>a+b,0) / valoresCump.length * 10) / 10
    : null;

  // ── Semáforo de alerta: EVA sube ≥3 puntos vs la semana anterior, reporta
  // corriente eléctrica constante, efecto adverso o gastritis → bandera roja
  // para revisión (búsqueda por nombre de columna, no por índice fijo, para
  // que funcione igual en hojas viejas y nuevas — mismo criterio que se usó
  // para la cola de mensajes del QR).
  let alerta = '';
  const evaMov = datos.dolor ? datos.dolor.eva_movimiento : null;
  const filasPrevias = sheet.getDataRange().getValues();
  const headersPrevios = filasPrevias[0] || [];
  const colDNIPrevio = headersPrevios.indexOf('DNI');
  const colEvaMovPrevio = headersPrevios.indexOf('EVA_Movimiento');
  let evaAnterior = null;
  if(colDNIPrevio > -1 && colEvaMovPrevio > -1){
    for(let i = filasPrevias.length - 1; i >= 1; i--){
      if(String(filasPrevias[i][colDNIPrevio]) === String(dni)){
        evaAnterior = filasPrevias[i][colEvaMovPrevio];
        break;
      }
    }
  }
  if(evaAnterior !== null && evaAnterior !== '' && evaMov !== null && (evaMov - evaAnterior) >= 3) alerta = 'DOLOR_EN_AUMENTO';
  if(datos.neurodinamico === 2) alerta = alerta ? alerta + ' + NEURO_CONSTANTE' : 'NEURO_CONSTANTE';
  const efAdv = datos.efectos_adversos || {};
  if(efAdv.tiene) alerta = alerta ? alerta + ' + EFECTO_ADVERSO' : 'EFECTO_ADVERSO';
  if(datos.gastritis) alerta = alerta ? alerta + ' + GASTRITIS' : 'GASTRITIS';

  // Asegura que existan todas las columnas (compatibilidad con hojas creadas
  // antes de agregar tipo de dolor / medicación / efectos adversos).
  const cols = ['Timestamp','DNI','Segmento','EVA_Reposo','EVA_Movimiento','Neurodinamico',
    'Incapacidad_Pct','Recuperacion_Pct','Escala_Nombre','Escala_Pct_Discapacidad',
    'Cumplimiento_Prom','Alerta','Tipo_Dolor','Medicacion_Tipo','Medicacion_Detalle',
    'Efectos_Adversos','Intolerancias','Gastritis','Datos_JSON'];
  const colIdx = {};
  cols.forEach(c => { colIdx[c] = asegurarColumnaColaMensajes_(sheet, c) - 1; }); // 0-based

  const med = datos.medicacion || {};
  const medicacionDetalle = (med.tipo === 'farmacopea' && Array.isArray(med.farmacopea))
    ? med.farmacopea.map(f => (String(f.farmaco||'') + ' ' + String(f.dosis||'')).trim()).filter(Boolean).join(' | ')
    : (med.tipo === 'magistral' ? ('Magistral — ' + (med.magistral_farmacia || 'sin especificar')) : '');
  const intol = datos.intolerancias || {};

  const fila = new Array(cols.length).fill('');
  fila[colIdx['Timestamp']] = new Date().toISOString();
  fila[colIdx['DNI']] = dni;
  fila[colIdx['Segmento']] = (datos.segmento && datos.segmento.titulo) || '';
  fila[colIdx['EVA_Reposo']] = datos.dolor ? datos.dolor.eva_reposo : '';
  fila[colIdx['EVA_Movimiento']] = datos.dolor ? datos.dolor.eva_movimiento : '';
  fila[colIdx['Neurodinamico']] = datos.neurodinamico != null ? datos.neurodinamico : '';
  fila[colIdx['Incapacidad_Pct']] = datos.incapacidad_relativa_pct != null ? datos.incapacidad_relativa_pct : '';
  fila[colIdx['Recuperacion_Pct']] = datos.recuperacion_percibida_pct != null ? datos.recuperacion_percibida_pct : '';
  fila[colIdx['Escala_Nombre']] = (datos.segmento && datos.segmento.escala_nombre) || '';
  fila[colIdx['Escala_Pct_Discapacidad']] = (datos.segmento && datos.segmento.escala_pct_discapacidad != null) ? datos.segmento.escala_pct_discapacidad : '';
  fila[colIdx['Cumplimiento_Prom']] = cumplimientoProm != null ? cumplimientoProm : '';
  fila[colIdx['Alerta']] = alerta;
  fila[colIdx['Tipo_Dolor']] = (datos.dolor && Array.isArray(datos.dolor.tipo)) ? datos.dolor.tipo.join(', ') : '';
  fila[colIdx['Medicacion_Tipo']] = med.tipo || '';
  fila[colIdx['Medicacion_Detalle']] = medicacionDetalle;
  fila[colIdx['Efectos_Adversos']] = efAdv.tiene ? (efAdv.detalle || 'Sí, sin detalle') : '';
  fila[colIdx['Intolerancias']] = intol.tiene ? (intol.detalle || 'Sí, sin detalle') : '';
  fila[colIdx['Gastritis']] = datos.gastritis ? 'Sí' : '';
  fila[colIdx['Datos_JSON']] = JSON.stringify(datos);

  sheet.getRange(sheet.getLastRow()+1, 1, 1, cols.length).setValues([fila]);

  return {success:true, message:'Autoevaluación registrada', alerta: alerta || null};
}

function getAutoevaluacionesRMYT(dni){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('RMYT_Autoevaluaciones');
  if(!sheet || !dni) return {success:true, data:[], tendencia:null};
  const rows = sheet.getDataRange().getValues();
  if(rows.length < 2) return {success:true, data:[], tendencia:null};
  const headers = rows[0];
  const idx = (nombre) => headers.indexOf(nombre);
  const iDNI=idx('DNI'), iFecha=idx('Timestamp'), iSeg=idx('Segmento'), iEvaR=idx('EVA_Reposo'),
    iEvaM=idx('EVA_Movimiento'), iNeuro=idx('Neurodinamico'), iIncap=idx('Incapacidad_Pct'),
    iRecup=idx('Recuperacion_Pct'), iEscN=idx('Escala_Nombre'), iEscPct=idx('Escala_Pct_Discapacidad'),
    iCump=idx('Cumplimiento_Prom'), iAlerta=idx('Alerta'), iTipoDolor=idx('Tipo_Dolor'),
    iMedTipo=idx('Medicacion_Tipo'), iMedDet=idx('Medicacion_Detalle'), iEfAdv=idx('Efectos_Adversos'),
    iIntol=idx('Intolerancias'), iGastritis=idx('Gastritis');

  const resultado = [];
  for(let i = 1; i < rows.length; i++){
    if(String(rows[i][iDNI]) === String(dni)){
      resultado.push({
        fecha: rows[i][iFecha],
        segmento: rows[i][iSeg],
        eva_reposo: rows[i][iEvaR],
        eva_movimiento: rows[i][iEvaM],
        neurodinamico: rows[i][iNeuro],
        incapacidad_pct: rows[i][iIncap],
        recuperacion_pct: rows[i][iRecup],
        escala_nombre: rows[i][iEscN],
        escala_pct_discapacidad: rows[i][iEscPct],
        cumplimiento_promedio: rows[i][iCump],
        alerta: rows[i][iAlerta],
        tipo_dolor: iTipoDolor>-1 ? rows[i][iTipoDolor] : '',
        medicacion_tipo: iMedTipo>-1 ? rows[i][iMedTipo] : '',
        medicacion_detalle: iMedDet>-1 ? rows[i][iMedDet] : '',
        efectos_adversos: iEfAdv>-1 ? rows[i][iEfAdv] : '',
        intolerancias: iIntol>-1 ? rows[i][iIntol] : '',
        gastritis: iGastritis>-1 ? rows[i][iGastritis] : ''
      });
    }
  }
  // Más reciente primero para mostrarlo en pantalla…
  resultado.sort((a,b) => new Date(b.fecha) - new Date(a.fecha));
  // …pero la tendencia se calcula en orden cronológico ascendente.
  const cronologico = resultado.slice().reverse();
  const tendencia = calcularTendenciaRMYT_(cronologico);
  return {success:true, data: resultado, tendencia};
}

// ══════════════════════════════════════════════════════════════════
// Motor predictivo simple del RMYT
// ══════════════════════════════════════════════════════════════════
// Compara la evolución OBJETIVA (% de discapacidad de la escala del
// segmento, o el EVA de movimiento si todavía no hay escala) contra la
// evolución SUBJETIVA (recuperación percibida que informa el propio
// paciente) semana a semana, mediante una regresión lineal simple
// (pendiente por semana). El objetivo es exactamente el que se pidió:
// poder distinguir, con el tiempo, una mejoría real y sostenida de una
// simple apreciación personal. NO es un modelo clínico ni reemplaza el
// criterio médico — por eso el resultado siempre se devuelve marcado
// como estimación (es_estimacion:true) y con lenguaje prudente.
function calcularTendenciaRMYT_(registrosCronologicos){
  const regs = (registrosCronologicos || []).filter(r => r.fecha);
  if(regs.length < 2){
    return {
      estado: 'datos_insuficientes',
      n_registros: regs.length,
      es_estimacion: true,
      mensaje: regs.length === 0
        ? 'Todavía no hay autoevaluaciones cargadas para este paciente.'
        : 'Hace falta al menos una segunda autoevaluación semanal para poder estimar una tendencia.'
    };
  }

  const t0 = new Date(regs[0].fecha).getTime();
  const semanas = regs.map(r => (new Date(r.fecha).getTime() - t0) / (1000*60*60*24*7));

  function pendientePorSemana(valores){
    const pares = [];
    for(let i=0; i<valores.length; i++){
      const v = valores[i];
      if(v !== '' && v !== null && v !== undefined && !isNaN(v)) pares.push([semanas[i], Number(v)]);
    }
    if(pares.length < 2) return null;
    const n = pares.length;
    const sx = pares.reduce((a,p)=>a+p[0],0), sy = pares.reduce((a,p)=>a+p[1],0);
    const sxy = pares.reduce((a,p)=>a+p[0]*p[1],0), sxx = pares.reduce((a,p)=>a+p[0]*p[0],0);
    const denom = (n*sxx - sx*sx);
    if(Math.abs(denom) < 1e-9) return 0; // todas las mediciones cayeron en la misma semana
    return (n*sxy - sx*sy) / denom;
  }

  const pendEva = pendientePorSemana(regs.map(r=>r.eva_movimiento));
  const pendDiscapacidad = pendientePorSemana(regs.map(r=>r.escala_pct_discapacidad));
  const pendRecuperacionPercibida = pendientePorSemana(regs.map(r=>r.recuperacion_pct));

  // Indicador objetivo preferido: % de discapacidad de la escala del
  // segmento; si todavía no hay escala cargada, se usa el EVA de
  // movimiento (llevado a una escala 0-100 para que sea comparable).
  let pendObjetiva = null, fuenteObjetiva = null;
  if(pendDiscapacidad !== null){ pendObjetiva = pendDiscapacidad; fuenteObjetiva = 'discapacidad'; }
  else if(pendEva !== null){ pendObjetiva = pendEva * 10; fuenteObjetiva = 'eva'; }

  let estado, mensaje;
  if(pendObjetiva === null){
    estado = 'datos_insuficientes';
    mensaje = 'No hay suficientes valores numéricos cargados todavía para estimar una tendencia objetiva.';
  } else if(pendObjetiva <= -2){
    estado = 'mejoria_sostenida';
    mensaje = 'Los indicadores objetivos (' + (fuenteObjetiva==='discapacidad' ? '% de discapacidad de la escala' : 'EVA de movimiento') + ') muestran una mejoría sostenida semana a semana.';
  } else if(pendObjetiva < 2){
    estado = 'estancado';
    mensaje = 'Los indicadores objetivos se mantienen prácticamente sin cambios entre evaluaciones — conviene revisar el plan terapéutico.';
  } else {
    estado = 'empeorando';
    mensaje = 'Los indicadores objetivos muestran una tendencia al empeoramiento — se recomienda revisión clínica prioritaria.';
  }

  // Discordancia entre lo objetivo y lo percibido: esto es justamente lo
  // que permite distinguir "mejoría evidente" de "apreciación personal".
  let discordancia = false, mensajeDiscordancia = null;
  if(pendRecuperacionPercibida !== null && pendObjetiva !== null){
    const percibidaMejorando = pendRecuperacionPercibida > 2;
    const objetivaMejorando = pendObjetiva <= -2;
    const objetivaEstancadaOPeor = pendObjetiva > -2;
    if(percibidaMejorando && objetivaEstancadaOPeor){
      discordancia = true;
      mensajeDiscordancia = 'El paciente refiere sentirse mejor, pero eso todavía no se refleja en los indicadores objetivos — podría ser una mejoría más subjetiva/anímica que funcional por ahora.';
    } else if(!percibidaMejorando && objetivaMejorando){
      discordancia = true;
      mensajeDiscordancia = 'Los indicadores objetivos mejoran pero el paciente no lo percibe así — vale la pena indagar dolor residual, expectativas u otros factores.';
    }
  }

  // Proyección simple: solo si hay mejoría objetiva clara en la escala de
  // discapacidad, estima en cuántas semanas llegaría a ≤20% (umbral
  // habitual de "discapacidad leve" en escalas tipo Oswestry) por
  // extrapolación lineal. Se omite si da un plazo demasiado largo/incierto.
  let proyeccionSemanas = null;
  if(fuenteObjetiva === 'discapacidad' && pendDiscapacidad < -0.5){
    const ultimo = regs[regs.length-1];
    const valorActual = Number(ultimo.escala_pct_discapacidad);
    if(!isNaN(valorActual) && valorActual > 20){
      const semanasEstim = Math.ceil((valorActual - 20) / Math.abs(pendDiscapacidad));
      if(semanasEstim <= 52) proyeccionSemanas = semanasEstim;
    }
  }

  return {
    estado, mensaje, discordancia,
    mensaje_discordancia: mensajeDiscordancia,
    pendiente_eva_semana: pendEva !== null ? Math.round(pendEva*100)/100 : null,
    pendiente_discapacidad_semana: pendDiscapacidad !== null ? Math.round(pendDiscapacidad*100)/100 : null,
    pendiente_recuperacion_percibida_semana: pendRecuperacionPercibida !== null ? Math.round(pendRecuperacionPercibida*100)/100 : null,
    proyeccion_semanas_a_leve: proyeccionSemanas,
    n_registros: regs.length,
    es_estimacion: true
  };
}

// ── RMYT: link y envío del formulario semanal de autoevaluación ──────
// Mismo patrón que el RIMP (construirLinkRimpAntecedentes_ / enviarLinkRimpAPaciente_):
// arma un link personalizado por DNI + segmento y lo manda por mail. Queda
// lista para que OCM la dispare desde el selector de RMYT/IMDCYMT — el
// cableado del botón en OCM es un paso aparte, todavía pendiente.
function getRmytUrl_(){
  return 'https://lmuratori.github.io/smm-acceso-temporal/RMYT_Autoevaluacion.html'; // GitHub Pages
}
function construirLinkRMYTAutoeval_(p){
  return getRmytUrl_() + '?dni=' + encodeURIComponent(p.dni||'') +
    '&segmento=' + encodeURIComponent(p.segmento||'') +
    '&apellido=' + encodeURIComponent(p.apellido||'') +
    '&nombre=' + encodeURIComponent(p.nombre||'');
}
function enviarLinkRMYTAPaciente_(p){
  if(!p || !p.email) throw new Error('Falta el mail del paciente');
  const link = construirLinkRMYTAutoeval_(p);
  const cuerpo =
    'Hola ' + (p.nombre||'') + ' ' + (p.apellido||'') + ',\n\n' +
    'Como parte de tu seguimiento' + (p.segmento ? (' (' + p.segmento + ')') : '') + ' te pedimos que completes esta breve autoevaluación semanal — te toma un par de minutos y nos ayuda a ver cómo vas evolucionando.\n\n' +
    'Por favor, entrá al siguiente link:\n' + link + '\n\n' +
    'Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943';
  MailApp.sendEmail(p.email, 'Autoevaluación semanal de seguimiento' + (p.segmento ? (' — ' + p.segmento) : ''), cuerpo, {
    replyTo:'drluismuratori@gmail.com', name:'Dr. Luis Alberto Muratori'
  });
}

// ══════════════════════════════════════════════════════════════════
// PARCHE — Canal de comunicación de Kinesiología (KSM)
// ══════════════════════════════════════════════════════════════════
// Mail del consultorio a donde llegan los mensajes/informes del sistema.
const MEDICO_MAIL_KINESIOLOGIA = 'consultas.sism@gmail.com';

function sheetPreguntasKine_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Preguntas_Kinesiologia');
  if(!sheet){
    sheet = ss.insertSheet('Preguntas_Kinesiologia');
    sheet.appendRow(['Timestamp','DNI','Paciente','Pregunta','Respuesta','Fecha_Respuesta','Estado']);
    sheet.getRange(1,1,1,7).setBackground('#166534').setFontColor('white').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// El paciente pregunta algo desde KSM_Autoevaluacion_Paciente.html
function guardarPreguntaKine(d){
  if(!d.dni || !d.pregunta) return{success:false, error:'Falta DNI o pregunta'};
  const sheet = sheetPreguntasKine_();
  sheet.appendRow([new Date().toISOString(), d.dni, d.paciente||'', d.pregunta, '', '', 'pendiente']);
  return{success:true, message:'Pregunta enviada al kinesiólogo/a'};
}

// KSM_Kinesiologia.html pide las preguntas de un paciente (o todas si no
// se manda DNI) para mostrarlas en el panel del profesional.
function getPreguntasKine(dni){
  const sheet = sheetPreguntasKine_();
  const data = sheet.getDataRange().getValues();
  if(data.length<2) return{success:true, preguntas:[]};
  const headers = data[0];
  let preguntas = data.slice(1).map((row,i)=>{
    const obj = {_row:i+2};
    headers.forEach((h,j)=>{ obj[h]=row[j]; });
    return obj;
  });
  if(dni) preguntas = preguntas.filter(p=> String(p.DNI)===String(dni));
  return{success:true, preguntas:preguntas};
}

// El kinesiólogo responde desde KSM_Kinesiologia.html — se guarda la
// respuesta y se le manda un mail al paciente (nunca WhatsApp, regla de
// oro del sistema). El replyTo queda con el mail del kinesiólogo, para
// que si el paciente contesta, le llegue a él/ella directo.
function responderPreguntaKine(d){
  const fila = parseInt(d.fila);
  if(!fila || !d.respuesta) return{success:false, error:'Falta la fila o la respuesta'};
  const sheet = sheetPreguntasKine_();
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const filaValores = sheet.getRange(fila,1,1,headers.length).getValues()[0];
  const dni = filaValores[headers.indexOf('DNI')];
  const paciente = filaValores[headers.indexOf('Paciente')];
  const pregunta = filaValores[headers.indexOf('Pregunta')];

  sheet.getRange(fila, headers.indexOf('Respuesta')+1).setValue(d.respuesta);
  sheet.getRange(fila, headers.indexOf('Fecha_Respuesta')+1).setValue(new Date().toISOString());
  sheet.getRange(fila, headers.indexOf('Estado')+1).setValue('respondida');

  // Buscar el mail del paciente en Pacientes por DNI
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = ss.getSheetByName('Pacientes');
  let emailPaciente = '';
  if(pSheet){
    const pData = pSheet.getDataRange().getValues();
    const pHeaders = pData[0];
    const dniCol = pHeaders.indexOf('DNI'), emailCol = pHeaders.indexOf('Email');
    const fila2 = pData.find(row=> String(row[dniCol])===String(dni));
    if(fila2) emailPaciente = fila2[emailCol];
  }

  if(emailPaciente){
    const cuerpo = 'Hola ' + (paciente||'') + ',\n\n' +
      'Tu kinesiólogo/a respondió tu consulta:\n\n' +
      '"' + pregunta + '"\n\n' +
      'Respuesta:\n' + d.respuesta + '\n\n' +
      'Consultorio Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943';
    try{
      MailApp.sendEmail(emailPaciente, 'Respuesta de tu kinesiólogo/a', cuerpo, {
        replyTo: d.kinesiologoEmail || MEDICO_MAIL_KINESIOLOGIA,
        name: d.kinesiologoNombre ? (d.kinesiologoNombre + ' — Consultorio Dr. Muratori') : 'Consultorio Dr. Muratori'
      });
    }catch(e){
      return{success:true, message:'Respuesta guardada, pero no se pudo enviar el mail: ' + e};
    }
  }
  return{success:true, message:'Respuesta guardada y enviada'};
}

// El kinesiólogo manda a KSM el resumen de adherencia/progreso ya armado
// (KSM_Kinesiologia.html arma el texto con las evoluciones + autoevaluaciones
// que ya trae get_historial) y esta función lo hace llegar al médico.
function enviarInformeKine(d){
  if(!d.resumen || !d.kinesiologoEmail) return{success:false, error:'Falta el resumen o el mail del kinesiólogo'};
  try{
    MailApp.sendEmail(MEDICO_MAIL_KINESIOLOGIA,
      '[KSM] Informe de kinesiología — ' + (d.paciente||d.dni||''),
      d.resumen,
      { replyTo: d.kinesiologoEmail, name: (d.kinesiologoNombre||'Kinesiología') + ' — KSM' }
    );
    return{success:true, message:'Informe enviado al médico'};
  }catch(e){
    return{success:false, error:'No se pudo enviar el informe: ' + e};
  }
}

// ══════════════════════════════════════════════════════════════════
// PARCHE — Peaje de uso del sistema (IA + desarrollo continuo)
// ══════════════════════════════════════════════════════════════════
function getPeajeSistemaPct(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Configuracion');
  if(!sh) return{success:true, porcentaje:5};
  const rows = sh.getDataRange().getValues();
  for(const row of rows){
    if(row[0]==='peaje_sistema_pct'){
      const n = Number(row[1]);
      return{success:true, porcentaje: isNaN(n) ? 5 : n};
    }
  }
  return{success:true, porcentaje:5};
}

function setPeajeSistemaPct(d){
  const pct = Number(d.porcentaje);
  if(isNaN(pct) || pct<0 || pct>100) return{success:false, error:'Porcentaje inválido'};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Configuracion');
  if(!sh){ sh = ss.insertSheet('Configuracion'); sh.appendRow(['clave','valor']); sh.setFrozenRows(1); }
  const rows = sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(rows[i][0]==='peaje_sistema_pct'){ sh.getRange(i+1,2).setValue(pct); return{success:true}; }
  }
  sh.appendRow(['peaje_sistema_pct', pct]);
  return{success:true};
}

// ══════════════════════════════════════════════════════════════════
// PARCHE — Autorización y permisos por especialista (generalizado a
// TODOS los profesionales, no solo kinesiología: "colegas" incluidos).
//
// Dos pasos separados, a propósito:
//   1) El PROFESIONAL acepta el % de peaje (registrarAceptacionKSM, como
//      ya existía) — su consentimiento a la comisión del sistema.
//   2) El MÉDICO autoriza y elige, por especialista, qué módulos ve
//      (setPermisosEspecialista) — su decisión de qué se comparte con
//      quién. Nadie ve nada del punto 2 hasta que el médico lo habilita,
//      aunque ya haya aceptado el peaje del punto 1.
//
// Podoscopía queda afuera a propósito: es de uso EXCLUSIVO del médico y
// nunca se ofrece como módulo compartible, ni aunque se autorice todo lo
// demás (ver MODULOS_COMPARTIBLES más abajo).
// ══════════════════════════════════════════════════════════════════
const MODULOS_COMPARTIBLES = ['investigacion','cud','grafica','plantillas'];

function sheetEspecialistas_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Kinesiologos_KSM');
  if(!sh){
    sh = ss.insertSheet('Kinesiologos_KSM');
    sh.appendRow(['Timestamp','Nombre','Matricula','Email','Porcentaje_Aceptado']);
    sh.getRange(1,1,1,5).setBackground('#166534').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  asegurarColumnasEspecialistas_(sh);
  return sh;
}
// Agrega las columnas nuevas del sistema de permisos a la hoja ya
// existente (Kinesiologos_KSM) sin tocar los registros que ya había.
function asegurarColumnasEspecialistas_(sh){
  let headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const faltantes = ['Tipo','Autorizado','Modulos_JSON','Porcentaje_Personalizado','Fecha_Actualizacion_Permisos'].filter(h=>headers.indexOf(h)===-1);
  faltantes.forEach(h=>{
    sh.getRange(1, sh.getLastColumn()+1).setValue(h).setBackground('#166534').setFontColor('white').setFontWeight('bold');
  });
  return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
}

// El profesional (kinesiólogo o colega) acepta el % de peaje vigente y
// queda registrado — "solo participa quien está de acuerdo". Si ya
// estaba registrado, actualiza sus datos sin pisar los permisos que el
// médico ya le haya otorgado.
function registrarAceptacionKSM(d){
  if(!d.email || !d.nombre) return{success:false, error:'Falta nombre o mail'};
  const sh = sheetEspecialistas_();
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  for(let i=1;i<data.length;i++){
    if(String(data[i][emailCol]).toLowerCase()===String(d.email).toLowerCase()){
      const fila = i+1;
      sh.getRange(fila, headers.indexOf('Timestamp')+1).setValue(new Date().toISOString());
      sh.getRange(fila, headers.indexOf('Nombre')+1).setValue(d.nombre);
      if(d.matricula) sh.getRange(fila, headers.indexOf('Matricula')+1).setValue(d.matricula);
      if(d.porcentaje) sh.getRange(fila, headers.indexOf('Porcentaje_Aceptado')+1).setValue(d.porcentaje);
      if(d.tipo) sh.getRange(fila, headers.indexOf('Tipo')+1).setValue(d.tipo);
      return{success:true, message:'Registro actualizado', yaExistia:true};
    }
  }
  const fila = new Array(headers.length).fill('');
  const set=(nombreCol,valor)=>{const i=headers.indexOf(nombreCol);if(i>-1)fila[i]=valor;};
  set('Timestamp', new Date().toISOString());
  set('Nombre', d.nombre);
  set('Matricula', d.matricula||'');
  set('Email', d.email);
  set('Porcentaje_Aceptado', d.porcentaje||'');
  set('Tipo', d.tipo||'kinesiologia');
  set('Autorizado', 'PENDIENTE');
  set('Modulos_JSON', '[]');
  sh.appendRow(fila);
  return{success:true, message:'Registro nuevo — queda pendiente de que el médico habilite sus módulos', yaExistia:false};
}

// Lo que consulta cada app (KSM, futuras apps de colegas) para saber qué
// puede mostrar de un profesional puntual.
function getPermisosEspecialista(email){
  if(!email) return{success:false, error:'Falta email'};
  const sh = sheetEspecialistas_();
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  for(let i=1;i<data.length;i++){
    if(String(data[i][emailCol]).toLowerCase()===String(email).toLowerCase()){
      let modulos = [];
      try{ modulos = JSON.parse(data[i][headers.indexOf('Modulos_JSON')]||'[]'); }catch(_){}
      const pctPersonal = data[i][headers.indexOf('Porcentaje_Personalizado')];
      return{
        success:true, registrado:true,
        autorizado: String(data[i][headers.indexOf('Autorizado')]).toUpperCase()==='SI',
        modulos: modulos,
        tipo: data[i][headers.indexOf('Tipo')]||'',
        porcentaje: (pctPersonal!=='' && pctPersonal!=null) ? Number(pctPersonal) : getPeajeSistemaPct().porcentaje
      };
    }
  }
  return{success:true, registrado:false, autorizado:false, modulos:[], tipo:'', porcentaje:getPeajeSistemaPct().porcentaje};
}

// Lo usa el MÉDICO (desde el panel de administración de especialistas)
// para autorizar y elegir módulos por profesional. 'podoscopia' se filtra
// siempre, aunque venga en la lista — no es negociable desde acá.
function setPermisosEspecialista(d){
  if(!d.email) return{success:false, error:'Falta email'};
  const sh = sheetEspecialistas_();
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  const modulos = (Array.isArray(d.modulos)?d.modulos:[]).filter(m=>MODULOS_COMPARTIBLES.indexOf(m)>-1);
  for(let i=1;i<data.length;i++){
    if(String(data[i][emailCol]).toLowerCase()===String(d.email).toLowerCase()){
      const fila = i+1;
      sh.getRange(fila, headers.indexOf('Autorizado')+1).setValue(d.autorizado?'SI':'NO');
      sh.getRange(fila, headers.indexOf('Modulos_JSON')+1).setValue(JSON.stringify(modulos));
      if(d.porcentajePersonalizado!==undefined && d.porcentajePersonalizado!==null && d.porcentajePersonalizado!==''){
        sh.getRange(fila, headers.indexOf('Porcentaje_Personalizado')+1).setValue(Number(d.porcentajePersonalizado));
      }
      if(d.tipo) sh.getRange(fila, headers.indexOf('Tipo')+1).setValue(d.tipo);
      sh.getRange(fila, headers.indexOf('Fecha_Actualizacion_Permisos')+1).setValue(new Date().toISOString());
      return{success:true};
    }
  }
  return{success:false, error:'No se encontró ningún especialista registrado con ese mail — primero tiene que entrar una vez a su app (KSM u otra) para registrarse y aceptar el peaje.'};
}

// Lista completa para el panel de administración del médico.
function getListaEspecialistas(){
  const sh = sheetEspecialistas_();
  const data = sh.getDataRange().getValues();
  if(data.length<2) return{success:true, especialistas:[]};
  const headers = data[0];
  const especialistas = data.slice(1).map(row=>{
    const obj={}; headers.forEach((h,i)=>{obj[h]=row[i];}); return obj;
  }).filter(e=>e.Email);
  return{success:true, especialistas};
}

// Compuerta única: todo endpoint que sirva un módulo compartible debe
// pasar por acá antes de devolver datos. 'podoscopia' siempre da false.
function tieneModuloAutorizado_(email, modulo){
  if(modulo==='podoscopia') return false;
  if(!email) return false;
  const p = getPermisosEspecialista(email);
  return !!(p.autorizado && p.modulos && p.modulos.indexOf(modulo)>-1);
}

function sheetFacturacionSistema_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Facturacion_Sistema');
  if(!sh){
    sh = ss.insertSheet('Facturacion_Sistema');
    sh.appendRow(['Timestamp','Origen','DNI','Paciente','Profesional_Nombre','Profesional_Email',
                  'Monto','Comision_Pct','Comision_Monto','Neto_Profesional','Mes_Referencia','Estado_Comision']);
    sh.getRange(1,1,1,12).setBackground('#0EA5E9').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// Función compartida — la usan tanto KSM (kinesiólogos) como Consultas
// Online (el propio médico) para registrar el peaje de una sesión/consulta
// paga. origen = 'kinesiologia' o 'consulta_medica'.
function registrarFacturacionSistema_(origen, dni, paciente, profesionalNombre, profesionalEmail, monto){
  if(!monto || monto<=0) return null;
  // Si el profesional tiene un % personalizado cargado (setPermisosEspecialista),
  // se usa ese; si no, el % global del sistema.
  let pct = getPeajeSistemaPct().porcentaje;
  if(profesionalEmail){
    const permisos = getPermisosEspecialista(profesionalEmail);
    if(permisos.registrado && !isNaN(Number(permisos.porcentaje))) pct = Number(permisos.porcentaje);
  }
  const comision = Math.round(monto * pct / 100 * 100) / 100;
  const neto = Math.round((monto - comision) * 100) / 100;
  const mesRef = new Date().toISOString().slice(0,7); // "2026-08"
  sheetFacturacionSistema_().appendRow([
    new Date().toISOString(), origen, dni||'', paciente||'',
    profesionalNombre||'', profesionalEmail||'',
    monto, pct, comision, neto, mesRef, 'pendiente'
  ]);
  return{comision, neto, porcentaje:pct};
}

// Endpoint que llama KSM_Kinesiologia.html cuando el kinesiólogo marca
// una sesión como paga.
function registrarFacturacionSistema(d){
  const r = registrarFacturacionSistema_('kinesiologia', d.dni, d.paciente, d.profesionalNombre||d.kinesiologoNombre, d.profesionalEmail||d.kinesiologoEmail, Number(d.monto));
  if(!r) return{success:false, error:'Monto inválido'};
  return Object.assign({success:true}, r);
}

// Resumen del mes — de un profesional puntual, o de todos si no se pasa
// mail (útil para vos: ver de un vistazo el peaje total del mes, sumando
// kinesiólogos + tus propias Consultas Online).
function getResumenFacturacionSistema(profesionalEmail, mes, origen){
  const sh = sheetFacturacionSistema_();
  const data = sh.getDataRange().getValues();
  if(data.length<2) return{success:true, filas:[], totalFacturado:0, totalComision:0, totalNeto:0};
  const headers = data[0];
  const mesRef = mes || new Date().toISOString().slice(0,7);
  let filas = data.slice(1).map(row=>{
    const obj={}; headers.forEach((h,i)=>{obj[h]=row[i];}); return obj;
  }).filter(f=> f.Mes_Referencia===mesRef);
  if(profesionalEmail) filas = filas.filter(f=> f.Profesional_Email===profesionalEmail);
  if(origen) filas = filas.filter(f=> f.Origen===origen);
  const totalFacturado = filas.reduce((s,f)=> s+Number(f.Monto||0), 0);
  const totalComision = filas.reduce((s,f)=> s+Number(f.Comision_Monto||0), 0);
  const totalNeto = filas.reduce((s,f)=> s+Number(f.Neto_Profesional||0), 0);
  return{success:true, filas:filas, mes:mesRef, totalFacturado:totalFacturado, totalComision:totalComision, totalNeto:totalNeto};
}

function enviarResumenFacturacionSistema(d){
  if(!d.profesionalEmail) return{success:false, error:'Falta el mail del profesional'};
  const resumen = getResumenFacturacionSistema(d.profesionalEmail, d.mes||'', '');
  const cuerpo =
    'Resumen de facturación del sistema — ' + (d.profesionalNombre||d.profesionalEmail) + '\n' +
    'Mes: ' + resumen.mes + '\n\n' +
    'Sesiones/consultas facturadas: ' + resumen.filas.length + '\n' +
    'Total facturado: $' + resumen.totalFacturado.toLocaleString('es-AR') + '\n' +
    'Peaje del sistema a transferir: $' + resumen.totalComision.toLocaleString('es-AR') + '\n' +
    'Neto del profesional: $' + resumen.totalNeto.toLocaleString('es-AR');
  try{
    MailApp.sendEmail(MEDICO_MAIL_KINESIOLOGIA,
      '[Sistema] Resumen de facturación — ' + (d.profesionalNombre||d.profesionalEmail) + ' — ' + resumen.mes,
      cuerpo,
      { replyTo: d.profesionalEmail, name: (d.profesionalNombre||'Sistema') + ' — SIM-M' }
    );
    return{success:true};
  }catch(e){
    return{success:false, error:'No se pudo enviar: ' + e};
  }
}

// ── actualizarConsulta (versión final — reemplaza la original) ──
// Único agregado respecto a la versión de siempre: cuando se marca una
// Consulta Online como respondida (respuesta_final), registra sola el
// peaje del sistema usando el precio que ya tenía cargado esa consulta.
function actualizarConsulta(d){
  const sh=sheetConsultas_();
  const rows=sh.getDataRange().getValues();
  const headers=rows[0];
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][0])===String(d.id)){
      if(d.estado!==undefined)   sh.getRange(i+1,headers.indexOf('Estado')+1).setValue(d.estado);
      if(d.payment_id)           sh.getRange(i+1,headers.indexOf('Payment_ID')+1).setValue(d.payment_id);
      if(d.respuesta_draft)      sh.getRange(i+1,headers.indexOf('Respuesta_Draft')+1).setValue(d.respuesta_draft);
      if(d.respuesta_final)      sh.getRange(i+1,headers.indexOf('Respuesta_Final')+1).setValue(d.respuesta_final);
      if(d.respuesta_final){
        sh.getRange(i+1,headers.indexOf('Fecha_Respuesta')+1).setValue(new Date().toISOString());
        try{
          registrarFacturacionSistema_(
            'consulta_medica',
            rows[i][headers.indexOf('DNI')],
            rows[i][headers.indexOf('Paciente')],
            'Dr. Luis Alberto Muratori',
            MEDICO_MAIL_KINESIOLOGIA,
            Number(rows[i][headers.indexOf('Precio')])
          );
        }catch(e){ Logger.log('No se pudo registrar el peaje de la consulta: ' + e); }
      }
      return{success:true};
    }
  }
  return{success:false,error:'Consulta no encontrada'};
}

// ══════════════════════════════════════════════════════════════════
// PARCHE — Procesador de Mensajes / Agenda de Correspondencia
// (versión optimizada en consumo de tokens — thinking al mínimo, cuerpo
// del mail limpio de citas antes de mandarlo a la IA, y log de tokens
// gastados en Log_Tokens_IA)
// ══════════════════════════════════════════════════════════════════
// SEGURIDAD (corregido en esta revisión): la clave de Gemini ya NO vive
// como texto plano acá — por eso había terminado expuesta antes. Ahora se
// guarda en las Propiedades del script, igual que ya se hacía con Mercado
// Pago y Claude. Para cargar la clave nueva (rotación):
//   1) Pegarla en configurarGeminiAPI() más abajo, donde dice PEGÁ_AQUÍ.
//   2) Elegir configurarGeminiAPI en el desplegable de arriba del editor.
//   3) Tocar ▶ Ejecutar UNA vez.
//   4) Borrar la clave de acá adentro si querés (ya quedó guardada, no hace falta de nuevo).
function configurarGeminiAPI(){
  const key = 'PEGÁ_AQUÍ_TU_GEMINI_API_KEY_NUEVA';
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
  Logger.log('Gemini API key guardada correctamente en las Propiedades del script.');
}
function getGeminiUrl_(){
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if(!key) throw new Error('Falta configurar GEMINI_API_KEY — ejecutar configurarGeminiAPI() una vez desde el editor.');
  return 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=' + key;
}
const LABEL_PROCESADO = 'SIM-M/Procesado';
const LABEL_ERROR = 'SIM-M/Error-procesamiento';
const PALABRAS_SENSIBLES = ['discapacidad','cud','certificado único','pnc','invalidez','hospital','pensión no contributiva'];

// Saca las cadenas de mails anteriores citados (que solo suman tokens de
// entrada sin aportar información nueva) antes de mandar el texto a la IA.
// No modifica el cuerpo original que se guarda en la planilla.
function limpiarCuerpoMail_(texto){
  if(!texto) return '';
  const marcadoresDeCita = [
    /El .{0,60}escribi[oó]:/i,
    /-----\s*Mensaje original\s*-----/i,
    /-----\s*Original Message\s*-----/i,
    /^>.*$/m,
    /De:\s*.+\n\s*Enviado:/i,
    /From:\s*.+\n\s*Sent:/i
  ];
  let corte = texto.length;
  marcadoresDeCita.forEach(regex=>{
    const m = texto.match(regex);
    if(m && typeof m.index === 'number' && m.index < corte) corte = m.index;
  });
  return texto.slice(0, corte).trim();
}

// Registra cuántos tokens gastó cada llamada a Gemini (input, pensamiento,
// output, total), para poder ver el consumo real con el tiempo.
function registrarUsoTokens_(funcion, modelo, usage){
  if(!usage) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Log_Tokens_IA');
  if(!sheet){
    sheet = ss.insertSheet('Log_Tokens_IA');
    sheet.appendRow(['Timestamp','Funcion','Modelo','Tokens_Input','Tokens_Pensamiento','Tokens_Output','Tokens_Total']);
    sheet.getRange(1,1,1,7).setBackground('#0c4a6e').setFontColor('white').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    new Date().toISOString(), funcion, modelo,
    usage.promptTokenCount||0, usage.thoughtsTokenCount||0,
    usage.candidatesTokenCount||0, usage.totalTokenCount||0
  ]);
}

function procesarBandejaEntrada(){
  let label = GmailApp.getUserLabelByName(LABEL_PROCESADO);
  if(!label) label = GmailApp.createLabel(LABEL_PROCESADO);
  let labelError = GmailApp.getUserLabelByName(LABEL_ERROR);
  if(!labelError) labelError = GmailApp.createLabel(LABEL_ERROR);

  const hilos = GmailApp.search('is:unread -label:"' + LABEL_PROCESADO + '" -label:"' + LABEL_ERROR + '"', 0, 20);
  const sheet = hojaColaMensajes_();

  hilos.forEach(hilo=>{
    try{
      const mensajes = hilo.getMessages();
      const ultimo = mensajes[mensajes.length-1];
      const remitente = ultimo.getFrom();
      const asunto = hilo.getFirstMessageSubject();
      const cuerpo = ultimo.getPlainBody().slice(0,3000);   // se guarda completo, como siempre
      const cuerpoParaIA = limpiarCuerpoMail_(cuerpo);       // versión recortada, solo para Gemini

      const textoCompleto = (asunto + ' ' + cuerpo).toLowerCase();
      const esSensible = PALABRAS_SENSIBLES.some(function(p){ return textoCompleto.includes(p); });

      const analisis = analizarConIA_(asunto, cuerpoParaIA, esSensible);

      const id = Utilities.getUuid();
      sheet.appendRow([
        id, new Date().toISOString(), remitente, asunto, cuerpo,
        analisis.clasificacion, analisis.cantPreguntas,
        analisis.requierePago, analisis.requierePago ? 'pendiente' : 'no_aplica',
        analisis.borrador, 'pendiente_revision', '', hilo.getId()
      ]);

      hilo.addLabel(label);
      hilo.markRead();
    }catch(err){
      hilo.addLabel(labelError);
      Logger.log('Error procesando hilo: ' + err);
    }
  });
}

function hojaColaMensajes_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Cola_Mensajes');
  if(!sheet){
    sheet = ss.insertSheet('Cola_Mensajes');
    sheet.appendRow(['id','fecha','remitente','asunto','cuerpo','clasificacion','cantPreguntas','requierePago','estadoPago','borrador','estado','fila_referencia','threadId']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function analizarConIA_(asunto, cuerpo, esSensible){
  const contextoSensible = esSensible ? 'El mensaje contiene palabras relacionadas a discapacidad/hospital público — probablemente corresponda al circuito gratuito y sensible.' : '';
  const prompt = 'Sos el asistente administrativo del Dr. Luis Alberto Muratori, médico traumatólogo (M.N. 100.540 / M.P. 9943, Mendoza, Argentina).\n' +
    'Te llegó este mail. Tu tarea es CLASIFICARLO y armar un BORRADOR de respuesta que el doctor va a revisar antes de enviar — nunca se envía sin su aprobación.\n\n' +
    contextoSensible + '\n\n' +
    'ASUNTO: ' + asunto + '\n' +
    'CUERPO:\n' + cuerpo + '\n\n' +
    'Reglas:\n' +
    '- "sensible_publico": discapacidad, CUD, PNC, hospital público → gratuito, respuesta mínima y clara.\n' +
    '- "privado": consulta de consultorio privado (cualquier cantidad de preguntas) → requiere pago antes de enviarse.\n\n' +
    'Devolvé SOLO un JSON válido, sin texto adicional, con esta forma exacta:\n' +
    '{ "clasificacion": "sensible_publico o privado", "cantPreguntas": numero entero, "requierePago": true o false, "borrador": "texto completo de la respuesta sugerida, en tono profesional y calido, en español, firmada como Dr. Luis Alberto Muratori" }';

  try{
    const resp = UrlFetchApp.fetch(getGeminiUrl_(), {
      method:'post', contentType:'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: {
          thinkingConfig: { thinkingLevel: 'minimal' }
        }
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(resp.getContentText());
    registrarUsoTokens_('analizarConIA_', 'gemini-3.6-flash', data.usageMetadata);
    let textoIA = data.candidates[0].content.parts[0].text;
    textoIA = textoIA.replace(/```json|```/g,'').trim();
    return JSON.parse(textoIA);
  }catch(err){
    return {
      clasificacion: esSensible ? 'sensible_publico' : 'privado',
      cantPreguntas: (cuerpo.match(/\?/g)||[]).length || 1,
      requierePago: !esSensible,
      borrador: '[COMPLETAR] no se pudo generar el borrador automatico: ' + err
    };
  }
}

function getMensajesPendientes(){
  const sheet = hojaColaMensajes_();
  const rows = sheet.getDataRange().getValues();
  // A partir de acá se leen las columnas por NOMBRE de encabezado, no por
  // índice fijo — así los campos nuevos del formulario del QR (Segmento,
  // Cuando_Paso, etc., ver registrarSolicitudQR) se pueden seguir agregando
  // sin romper esta función ni depender del orden exacto en que se crearon.
  const headers = rows.length ? rows[0] : [];
  const idx = (nombreCol) => headers.indexOf(nombreCol);
  const val = (row, nombreCol) => { const i = idx(nombreCol); return i>-1 ? row[i] : ''; };
  const data = [];
  for(let i=1;i<rows.length;i++){
    if(rows[i][10] !== 'pendiente_revision') continue;
    data.push({
      id: rows[i][0], fecha: rows[i][1], remitente: rows[i][2], asunto: rows[i][3],
      cuerpo: rows[i][4], clasificacion: rows[i][5], cantPreguntas: rows[i][6],
      requierePago: rows[i][7], estadoPago: rows[i][8], borrador: rows[i][9],
      estado: rows[i][10], fila: i+1,
      // De dónde vino: el nombre de lugar que trae el QR (ver
      // registrarSolicitudQR/asegurarColumnaOrigenColaMensajes_), o
      // vacío/'Mail' para lo que sigue llegando por el triage de Gmail.
      origen: val(rows[i],'Origen') || 'Mail',
      // Nivel de riesgo que estimó la IA a partir del relato del paciente
      // en el formulario del QR (ver evaluarTriageIA_/registrarSolicitudQR).
      // Vacío en lo que llega por mail o en solicitudes de antes de este agregado.
      riesgo: val(rows[i],'Riesgo_IA') || '',
      // 'guardia' o 'entrevista' — qué recomendó la IA (o el valor por
      // defecto si la IA no respondió). Vacío en mensajes que no vinieron por QR.
      recomendacion: val(rows[i],'Recomendacion_IA') || '',
      segmento: val(rows[i],'Segmento') || '',
      cuando: val(rows[i],'Cuando_Paso') || '',
      vioMedico: val(rows[i],'Vio_Medico') || '',
      diagnosticoPrevio: val(rows[i],'Diagnostico_Previo') || '',
      medicacionRecibida: val(rows[i],'Medicacion_Recibida') || '',
      fotoUrl: val(rows[i],'Foto_URL') || ''
    });
  }
  // Lo marcado como riesgo "alto" por la IA sube primero, sin importar la
  // fecha — para que el médico lo vea de entrada al abrir la lista.
  const rangoRiesgo_ = function(r){ return r==='alto' ? 0 : (r==='medio' ? 1 : 2); };
  data.sort(function(a,b){
    const diff = rangoRiesgo_(a.riesgo) - rangoRiesgo_(b.riesgo);
    if(diff !== 0) return diff;
    return new Date(b.fecha) - new Date(a.fecha);
  });
  return {success:true, data:data};
}

// ══════════════════════════════════════════════════════════════════
// QR — entrada pública de "consulta virtual de bajo costo", pensada para
// pacientes que llegan por un código QR (guardia, comercios, clubes, etc.)
// en vez de escribir por mail. Cae en la MISMA cola (Cola_Mensajes) que ya
// revisás en Agenda de Mensajes — no es un circuito aparte — sólo que acá
// no hace falta que la IA clasifique nada, porque el paciente ya cargó los
// datos en un formulario. Queda marcada como pago porque es la consulta
// de bajo costo (el cobro real se genera desde Agenda de Mensajes con el
// mismo botón de Mercado Pago que ya usa OCM).
// ══════════════════════════════════════════════════════════════════

// Agrega la columna 'Origen' a Cola_Mensajes si todavía no existe —
// retrocompatible con las filas viejas que llegaron por mail (van a quedar
// con esa celda vacía, y getMensajesPendientes las muestra como 'Mail').
function asegurarColumnaOrigenColaMensajes_(sheet){
  let headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  let col = headers.indexOf('Origen');
  if(col===-1){
    col = headers.length;
    sheet.getRange(1, col+1).setValue('Origen').setBackground('#0c4a6e').setFontColor('white').setFontWeight('bold');
  }
  return col+1; // 1-based, para getRange/appendRow
}

// Igual que la de Origen, pero para el nivel de riesgo que estima la IA
// (ver evaluarTriageIA_) a partir del relato que carga el paciente en el
// formulario del QR.
function asegurarColumnaRiesgoColaMensajes_(sheet){
  let headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  let col = headers.indexOf('Riesgo_IA');
  if(col===-1){
    col = headers.length;
    sheet.getRange(1, col+1).setValue('Riesgo_IA').setBackground('#0c4a6e').setFontColor('white').setFontWeight('bold');
  }
  return col+1;
}

// Agrega cualquier columna nueva de Cola_Mensajes por nombre — versión
// genérica de asegurarColumnaOrigenColaMensajes_/asegurarColumnaRiesgoColaMensajes_,
// para no repetir la misma función una y otra vez por cada campo nuevo del
// formulario del QR.
function asegurarColumnaColaMensajes_(sheet, nombreCol){
  let headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  let col = headers.indexOf(nombreCol);
  if(col===-1){
    col = headers.length;
    sheet.getRange(1, col+1).setValue(nombreCol).setBackground('#0c4a6e').setFontColor('white').setFontWeight('bold');
  }
  return col+1; // 1-based
}

// Guarda (si vino) la foto del estudio/informe que adjunta el paciente en el
// formulario del QR. Nunca bloquea el envío: si falla por lo que sea
// (imagen mal formada, sin permiso de Drive, etc.) devuelve '' y
// registrarSolicitudQR sigue como si no hubiese venido foto.
function guardarFotoEstudioQR_(fotoBase64, nombrePaciente){
  try{
    if(!fotoBase64) return '';
    const partes = String(fotoBase64).split(',');
    const meta = partes.length > 1 ? partes[0] : '';
    const b64 = partes.length > 1 ? partes[1] : partes[0];
    const mimeMatch = meta.match(/data:(.*?);base64/);
    const mime = mimeMatch ? mimeMatch[1] : 'image/jpeg';
    const bytes = Utilities.base64Decode(b64);
    const blob = Utilities.newBlob(bytes, mime, 'estudio_' + Date.now());
    const carpetas = DriveApp.getFoldersByName('SIMM - Fotos de Estudios (QR)');
    const carpeta = carpetas.hasNext() ? carpetas.next() : DriveApp.createFolder('SIMM - Fotos de Estudios (QR)');
    const archivo = carpeta.createFile(blob);
    archivo.setName((nombrePaciente || 'paciente') + ' — ' + archivo.getName());
    return archivo.getUrl();
  }catch(e){
    return '';
  }
}

// Mensaje de reserva para cuando la IA no está disponible o falla — nunca
// dejamos al paciente sin ninguna indicación de qué sigue.
function mensajePacienteReserva_(){
  return 'Recibimos tu consulta. Nos vamos a comunicar a la brevedad para coordinar la entrevista express de asesoramiento médico y explicarte cómo seguir.';
}

// Primera lectura que hace la IA sobre todo lo que cuenta el paciente en el
// formulario del QR (relato, segmento afectado, cuándo pasó, si ya vio un
// médico y qué le dijeron, qué medicación recibió) — pensada para dar, en
// TODOS los casos, una de dos recomendaciones claras: ir ya a una guardia
// traumatológica (si hay señales de alarma reales), o que corresponde
// coordinar la entrevista express de asesoramiento médico paga. NUNCA es un
// diagnóstico, y si la IA falla o no está configurada, no bloquea nada — la
// solicitud igual queda registrada para que el médico la revise a mano.
function evaluarTriageIA_(datos){
  try{
    const motivo = (datos && datos.motivo) || '';
    const prompt =
      'Sos un asistente de triage para el consultorio de traumatología/ortopedia del Dr. Luis Alberto Muratori. ' +
      'Un paciente completó este relato para pedir orientación por QR, en vez de ir directo a una guardia. ' +
      'Tu tarea es dar una PRIMERA LECTURA de riesgo — nunca un diagnóstico — para que el médico decida rápido cómo seguir, ' +
      'y para decirle al paciente, en TODOS los casos, cuál es el paso recomendado: ir a una guardia ahora, ' +
      'o que corresponde coordinar la entrevista express de asesoramiento médico (paga, de 5 minutos, por videollamada).\n\n' +
      'Relato del paciente: """' + motivo + '"""\n' +
      'Segmento del cuerpo afectado: ' + ((datos && datos.segmento) || 'no especificado') + '\n' +
      'Cuándo pasó: ' + ((datos && datos.cuando) || 'no especificado') + '\n' +
      '¿Ya vio a un médico?: ' + ((datos && datos.vioMedico) || 'no especificado') +
      ((datos && datos.diagnosticoPrevio) ? (' — diagnóstico que le dieron: ' + datos.diagnosticoPrevio) : '') + '\n' +
      'Medicación que ya recibió: ' + ((datos && datos.medicacionRecibida) || 'ninguna mencionada') + '\n\n' +
      'Devolvé SOLO un JSON válido, sin texto adicional ni markdown, con esta forma exacta:\n' +
      '{ "nivel_riesgo": "bajo, medio o alto", ' +
      '"recomendacion": "guardia o entrevista", ' +
      '"razonamiento": "1-2 frases para el médico, en términos clínicos, explicando por qué asignaste ese nivel", ' +
      '"estudios_habituales": "estudios que suelen pedirse en casos así, en 1 frase, o vacío si no aplica", ' +
      '"mensaje_paciente": "SIEMPRE completo (nunca vacío): 2-3 frases en lenguaje simple, sin nombrar ninguna enfermedad ni usar jerga médica, ' +
      'explicando el paso recomendado — si es guardia, decirle que vaya ahora a la guardia traumatológica más cercana; ' +
      'si es entrevista, explicarle que corresponde coordinar la entrevista express de asesoramiento médico, y que ahí se le va a indicar cómo seguir." }\n\n' +
      'Reglas: "guardia"/"alto" es solo para señales de alarma reales (déficit neurológico, deformidad evidente, imposibilidad total de apoyo o movimiento, signos de infección grave, dolor torácico, etc.) — no lo uses por precaución genérica. ' +
      'Si no hay señales de alarma claras, usá "entrevista" con nivel "bajo" o "medio". No inventes datos que el paciente no haya contado.';

    const resp = UrlFetchApp.fetch(getGeminiUrl_(), {
      method:'post', contentType:'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { thinkingConfig: { thinkingLevel: 'minimal' } }
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(resp.getContentText());
    registrarUsoTokens_('evaluarTriageIA_', 'gemini-3.6-flash', data.usageMetadata);
    if(!data.candidates || !data.candidates.length) return null;
    let texto = data.candidates[0].content.parts[0].text || '';
    texto = texto.replace(/```json|```/g,'').trim();
    const j = JSON.parse(texto);
    if(!j || ['bajo','medio','alto'].indexOf(j.nivel_riesgo)===-1) return null;
    if(['guardia','entrevista'].indexOf(j.recomendacion)===-1){
      j.recomendacion = j.nivel_riesgo==='alto' ? 'guardia' : 'entrevista';
    }
    if(!j.mensaje_paciente) j.mensaje_paciente = mensajePacienteReserva_();
    return j;
  }catch(e){
    return null;
  }
}

// d = {nombre, dni, telefono, email, obraSocial, motivo, segmento, cuando,
//      vioMedico, diagnosticoPrevio, medicacionRecibida, fotoBase64, origen}
function registrarSolicitudQR(d){
  if(!d || !String(d.nombre||'').trim()) return {success:false, error:'Falta el nombre'};
  if(!String(d.motivo||'').trim()) return {success:false, error:'Falta contar el motivo de consulta'};
  if(!String(d.telefono||'').trim() && !String(d.email||'').trim()){
    return {success:false, error:'Dejá al menos un teléfono o un mail de contacto'};
  }

  const sheet = hojaColaMensajes_();
  const colOrigen = asegurarColumnaOrigenColaMensajes_(sheet);
  const colRiesgo = asegurarColumnaRiesgoColaMensajes_(sheet);
  const colSegmento = asegurarColumnaColaMensajes_(sheet, 'Segmento');
  const colCuando = asegurarColumnaColaMensajes_(sheet, 'Cuando_Paso');
  const colVioMedico = asegurarColumnaColaMensajes_(sheet, 'Vio_Medico');
  const colDiagPrevio = asegurarColumnaColaMensajes_(sheet, 'Diagnostico_Previo');
  const colMedicacion = asegurarColumnaColaMensajes_(sheet, 'Medicacion_Recibida');
  const colFoto = asegurarColumnaColaMensajes_(sheet, 'Foto_URL');
  const colRecomendacion = asegurarColumnaColaMensajes_(sheet, 'Recomendacion_IA');
  const anchoFila = Math.max(colOrigen, colRiesgo, colSegmento, colCuando, colVioMedico, colDiagPrevio, colMedicacion, colFoto, colRecomendacion);

  // La foto nunca bloquea el envío — si falla, sigue todo igual sin ella.
  const fotoUrl = d.fotoBase64 ? guardarFotoEstudioQR_(d.fotoBase64, d.nombre) : '';

  const triage = evaluarTriageIA_(d);
  const recomendacion = triage ? triage.recomendacion : 'entrevista'; // sin IA, por defecto se coordina la entrevista (nunca se asume urgencia sola)
  const mensajePaciente = triage ? triage.mensaje_paciente : mensajePacienteReserva_();

  const asunto = 'Solicitud por QR — ' + d.nombre;
  let cuerpo =
    'Solicitud de orientación / entrevista express, ingresada por QR.\n\n' +
    'Nombre: ' + (d.nombre||'') + '\n' +
    'DNI: ' + (d.dni||'') + '\n' +
    'Teléfono: ' + (d.telefono||'') + '\n' +
    'Email: ' + (d.email||'') + '\n' +
    'Obra social / cobertura: ' + (d.obraSocial||'Particular') + '\n' +
    'Origen del QR: ' + (d.origen||'sin especificar') + '\n\n' +
    'Motivo de consulta:\n' + (d.motivo||'') + '\n\n' +
    'Segmento afectado: ' + (d.segmento||'—') + '\n' +
    'Cuándo pasó: ' + (d.cuando||'—') + '\n' +
    '¿Ya vio a un médico?: ' + (d.vioMedico||'—') + (d.diagnosticoPrevio ? (' — diagnóstico dado: ' + d.diagnosticoPrevio) : '') + '\n' +
    'Medicación/receta que le dieron: ' + (d.medicacionRecibida||'—') +
    (fotoUrl ? ('\nFoto de estudio/informe adjunta: ' + fotoUrl) : '');
  if(triage){
    cuerpo += '\n\n── Evaluación preliminar (IA — no reemplaza tu criterio clínico) ──\n' +
      'Nivel de riesgo estimado: ' + triage.nivel_riesgo + '\n' +
      'Recomendación: ' + (triage.recomendacion==='guardia' ? 'ir a guardia ahora' : 'coordinar entrevista express') + '\n' +
      'Razonamiento: ' + (triage.razonamiento||'—') +
      (triage.estudios_habituales ? ('\nEstudios habituales en casos así: ' + triage.estudios_habituales) : '');
  }
  const borrador =
    'Hola ' + (d.nombre||'') + ',\n\n' +
    (recomendacion==='guardia'
      ? 'Por lo que nos contaste, te recomendamos ir ahora a la guardia traumatológica más cercana para una evaluación en persona sin esperar.'
      : 'Recibimos tu consulta. Te vamos a contactar a la brevedad para coordinar el día y el horario de la entrevista express de asesoramiento médico.') + '\n\n' +
    'Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943';

  const id = Utilities.getUuid();
  const fila = new Array(anchoFila).fill('');
  fila[0] = id;
  fila[1] = new Date().toISOString();
  fila[2] = d.email || d.telefono || '';
  fila[3] = asunto;
  fila[4] = cuerpo;
  fila[5] = 'privado';
  fila[6] = 1;
  fila[7] = true;          // requierePago — la entrevista express es paga (salvo derivación directa a guardia)
  fila[8] = 'pendiente';   // estadoPago
  fila[9] = borrador;
  fila[10] = 'pendiente_revision';
  fila[11] = '';           // fila_referencia
  fila[12] = '';           // threadId — no viene de un hilo de Gmail
  fila[colOrigen-1] = d.origen || 'Sin especificar';
  fila[colRiesgo-1] = triage ? triage.nivel_riesgo : '';
  fila[colSegmento-1] = d.segmento || '';
  fila[colCuando-1] = d.cuando || '';
  fila[colVioMedico-1] = d.vioMedico || '';
  fila[colDiagPrevio-1] = d.diagnosticoPrevio || '';
  fila[colMedicacion-1] = d.medicacionRecibida || '';
  fila[colFoto-1] = fotoUrl;
  fila[colRecomendacion-1] = recomendacion;
  sheet.appendRow(fila);

  return {
    success: true,
    id: id,
    recomendacion: recomendacion,
    alerta_riesgo: recomendacion==='guardia',
    mensaje_paciente: mensajePaciente
  };
}

function confirmarPagoMensaje(d){
  const sheet = hojaColaMensajes_();
  sheet.getRange(d.fila, 9).setValue('confirmado');
  return {success:true};
}

function marcarRespondido(d){
  const sheet = hojaColaMensajes_();
  const threadId = sheet.getRange(d.fila, 13).getValue();
  try{
    const thread = GmailApp.getThreadById(threadId);
    thread.reply(d.respuestaFinal, { name: 'Dr. Luis Alberto Muratori' });
  }catch(e){
    GmailApp.sendEmail(d.destinatario, 'Re: ' + (d.asunto||'Su consulta'), d.respuestaFinal, { name: 'Dr. Luis Alberto Muratori' });
  }
  sheet.getRange(d.fila, 11).setValue('respondido');
  return {success:true};
}

function generarLinkMeet(d){
  const inicio = new Date(d.fechaHoraISO);
  const fin = new Date(inicio.getTime() + (d.duracionMin||5)*60000);
  const evento = CalendarApp.getDefaultCalendar().createEvent(
    d.titulo || 'Consulta virtual con el Dr. Muratori',
    inicio, fin,
    { description: d.descripcion||'', guests: d.mailPaciente||'', sendInvites: !!d.mailPaciente }
  );
  let link = '';
  try{
    link = 'https://calendar.google.com/calendar/event?eid=' + Utilities.base64Encode(evento.getId() + ' ' + Session.getActiveUser().getEmail()).replace(/=+$/,'');
  }catch(e){}
  return {success:true, link: link || 'Evento creado, revisar en Google Calendar', evento_id: evento.getId()};
}
function testGeminiDirecto(){
  try{
    const resp = UrlFetchApp.fetch(getGeminiUrl_(), {
      method:'post', contentType:'application/json',
      payload: JSON.stringify({
        contents: [{ parts: [{ text: 'Decime hola en una palabra' }] }],
        generationConfig: {
          thinkingConfig: { thinkingLevel: 'minimal' }
        }
      }),
      muteHttpExceptions: true
    });
    const codigo = resp.getResponseCode();
    const texto = resp.getContentText();
    MailApp.sendEmail('lmuratori@gmail.com', 'Test Gemini — resultado (optimizado)', 'Código HTTP: ' + codigo + '\n\nRespuesta completa:\n' + texto);
  }catch(e){
    MailApp.sendEmail('lmuratori@gmail.com', 'Test Gemini — ERROR', 'No se pudo ni siquiera llamar a Gemini: ' + e + '\n\nSi el error dice "Falta configurar GEMINI_API_KEY", hay que correr configurarGeminiAPI() primero.');
  }
}

// ══════════════════════════════════════════════════════════════════
// Disparador automático de procesarBandejaEntrada
// Ejecutar UNA SOLA VEZ manualmente desde el editor (▶ Ejecutar).
// ══════════════════════════════════════════════════════════════════
function crearTriggerProcesarBandejaEntrada(){
  ScriptApp.getProjectTriggers().forEach(t=>{
    if(t.getHandlerFunction()==='procesarBandejaEntrada') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('procesarBandejaEntrada')
    .timeBased()
    .everyMinutes(15)
    .create();
  Logger.log('Disparador creado: procesarBandejaEntrada va a correr cada 15 minutos.');
}

// ══════════════════════════════════════════════════════════════════
// PARCHE — Sincronización de Catastro de Empresas + cruce de Investigación
// Antes el Catastro vivía SOLO en IndexedDB del navegador de OCM: nunca
// llegaba a la nube, así que no había forma real de compartir el cruce de
// Investigación con nadie más (aunque el Dashboard de OCM lo mostrara).
// Esto agrega una copia del catastro en una hoja compartida y replica acá
// el mismo cruce por reglas que antes solo corría en el navegador del
// médico, para que KSM (y futuras apps de colegas) lo puedan consultar.
// ══════════════════════════════════════════════════════════════════
function sheetCatastroEmpresas_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Catastro_Empresas');
  if(!sh){
    sh = ss.insertSheet('Catastro_Empresas');
    sh.appendRow(['Timestamp','Nombre','Localidad','Agentes_Quimicos_JSON','Agentes_Fisicos_JSON','Agentes_Biologicos_JSON','Registrada']);
    sh.getRange(1,1,1,7).setBackground('#0f172a').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  asegurarColumnasCatastro_(sh);
  return sh;
}
// Columnas de vía hídrica (afluente/canal de riego) — se agregan a la hoja
// ya existente sin tocar los registros que ya había.
function asegurarColumnasCatastro_(sh){
  let headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const faltantes = ['Afluente','Afluente_Toxico','Afluente_Organico'].filter(h=>headers.indexOf(h)===-1);
  faltantes.forEach(h=>{
    sh.getRange(1, sh.getLastColumn()+1).setValue(h).setBackground('#0f172a').setFontColor('white').setFontWeight('bold');
  });
  return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
}
// La llama OCM cada vez que se agrega una empresa al Catastro (además de
// guardarla en su IndexedDB local, como ya hacía).
// Upsert por Nombre+Localidad (antes hacía appendRow siempre, así que
// guardar la misma empresa dos veces —o correr la migración de datos
// locales más de una vez— duplicaba filas en el Catastro).
function sincronizarEmpresaCatastro(d){
  if(!d || !d.nombre || !d.localidad) return{success:false, error:'Falta nombre o localidad de la empresa'};
  const sh = sheetCatastroEmpresas_();
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const fila = new Array(headers.length).fill('');
  const set=(nombreCol,valor,destino)=>{const i=headers.indexOf(nombreCol);if(i>-1)destino[i]=valor;};
  const valores = {};
  set('Timestamp', new Date().toISOString(), valores);
  set('Nombre', d.nombre, valores);
  set('Localidad', d.localidad, valores);
  set('Agentes_Quimicos_JSON', JSON.stringify(d.agentes_quimicos||[]), valores);
  set('Agentes_Fisicos_JSON', JSON.stringify(d.agentes_fisicos||[]), valores);
  set('Agentes_Biologicos_JSON', JSON.stringify(d.agentes_biologicos||[]), valores);
  set('Registrada', d.registrada!==false ? 'SI':'NO', valores);
  // Vía hídrica: descarga de efluentes a canal de riego. El contenido
  // orgánico y/o tóxico puede afectar a pacientes de OTRA localidad,
  // aguas abajo del mismo canal — por eso se cruza aparte de la OIT.
  set('Afluente', d.afluente||'', valores);
  set('Afluente_Toxico', d.afluente_toxico ? 'SI':'NO', valores);
  set('Afluente_Organico', d.afluente_organico ? 'SI':'NO', valores);
  const nombreCol = headers.indexOf('Nombre'), localidadCol = headers.indexOf('Localidad');
  const data = sh.getDataRange().getValues();
  const nombreNorm = quitarAcentosGs_(d.nombre), localidadNorm = quitarAcentosGs_(d.localidad);
  for(let i=1;i<data.length;i++){
    if(quitarAcentosGs_(data[i][nombreCol])===nombreNorm && quitarAcentosGs_(data[i][localidadCol])===localidadNorm){
      const filaExistente = data[i].slice();
      headers.forEach((h,idx)=>{ if(valores[idx]!==undefined) filaExistente[idx]=valores[idx]; });
      sh.getRange(i+1,1,1,filaExistente.length).setValues([filaExistente]);
      return{success:true, message:'Empresa actualizada en el catastro'};
    }
  }
  headers.forEach((h,idx)=>{ if(valores[idx]!==undefined) fila[idx]=valores[idx]; });
  sh.appendRow(fila);
  return{success:true, message:'Empresa agregada al catastro'};
}

// Misma tabla de reglas OIT que usa OCM (Dashboard → Investigación) —
// si se actualiza una, hay que actualizar la otra a mano; no comparten
// archivo porque uno corre en el navegador y el otro en Apps Script.
const DIAGNOSTICO_AGENTE_MAP_ = [
  {palabras:['asma','bronqui','epoc','respirator','neumon','pulmon'], agentes:['Polvos y partículas','Humos metálicos','Gases y vapores','Asbesto','Hidrocarburos'], categoria:'Respiratorio'},
  {palabras:['dermat','eczema','alergia cutanea','alergia de piel','piel'], agentes:['Ácidos y álcalis','Solventes orgánicos','Plaguicidas / agroquímicos'], categoria:'Dermatológico'},
  {palabras:['neuropat','temblor','convuls','encefalopat','neurolog'], agentes:['Solventes orgánicos','Metales pesados','Plaguicidas / agroquímicos'], categoria:'Neurológico'},
  {palabras:['hipoacusia','sordera','audit'], agentes:['Ruido'], categoria:'Auditivo'},
  {palabras:['vibracion','sindrome vibratorio','raynaud'], agentes:['Vibración'], categoria:'Osteoarticular por vibración'},
  {palabras:['cancer','oncolog','tumor','leucemia'], agentes:['Asbesto','Hidrocarburos','Plaguicidas / agroquímicos','Radiación ionizante'], categoria:'Oncológico'},
  {palabras:['parasit','infeccios','zoonosis','bacteri','viral'], agentes:['Bacterias','Virus','Parásitos','Riesgo zoonótico (contacto con animales)'], categoria:'Infeccioso / zoonótico'},
  {palabras:['malformacion','congenit'], agentes:['Plaguicidas / agroquímicos'], categoria:'Congénito'}
];
function quitarAcentosGs_(s){ return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase(); }

// Cruce de Investigación para UN paciente puntual — usa Pacientes.Localidad
// + el/los diagnóstico(s) ya sincronizados del Libro de Discapacidad
// (Discapacidad_CUD) + Catastro_Empresas.
function getInvestigacionPaciente(dni){
  if(!dni) return{success:false, error:'Falta DNI'};
  const cudRows = leerPorDni_('Discapacidad_CUD', dni);
  const diagnosticos = [...new Set(cudRows.map(r=>r.Diagnostico).filter(Boolean))];
  // Localidad y afluente vienen de la página 1 del Libro de Discapacidad
  // (mismo origen que usaba el cruce original en OCM). Si por algún motivo
  // no está cargada ahí, se usa Pacientes.Localidad como respaldo.
  let localidad = (cudRows.map(r=>r.Localidad).filter(Boolean)[0]) || '';
  const afluente = (cudRows.map(r=>r.Afluente).filter(Boolean)[0]) || '';
  if(!localidad){
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pSheet = ss.getSheetByName('Pacientes');
    if(pSheet){
      const pData = pSheet.getDataRange().getValues(), pHeaders = pData[0];
      const dniCol = pHeaders.indexOf('DNI'), locCol = pHeaders.indexOf('Localidad');
      const fila = pData.find(row=>String(row[dniCol])===String(dni));
      if(fila) localidad = fila[locCol]||'';
    }
  }
  if((!localidad && !afluente) || !diagnosticos.length){
    return{success:true, localidad, afluente, hallazgos:[], hallazgos_hidricos:[], mensaje: (!localidad && !afluente) ? 'El paciente no tiene localidad ni afluente cargados.' : 'El paciente todavía no tiene diagnóstico sincronizado desde el Libro de Discapacidad.'};
  }
  const empresas = sheetToObjects('Catastro_Empresas').map(e=>{
    let aq=[],af=[],ab=[];
    try{ aq = JSON.parse(e.Agentes_Quimicos_JSON||'[]'); }catch(_){}
    try{ af = JSON.parse(e.Agentes_Fisicos_JSON||'[]'); }catch(_){}
    try{ ab = JSON.parse(e.Agentes_Biologicos_JSON||'[]'); }catch(_){}
    return {
      nombre:e.Nombre, localidad:e.Localidad, agentes_quimicos:aq, agentes_fisicos:af, agentes_biologicos:ab,
      afluente:e.Afluente||'', afluente_toxico:String(e.Afluente_Toxico).toUpperCase()==='SI', afluente_organico:String(e.Afluente_Organico).toUpperCase()==='SI'
    };
  });

  // 1) Vía OIT clásica — misma localidad + agente asociado al diagnóstico.
  const localidadNorm = quitarAcentosGs_(localidad);
  const hallazgos = [];
  if(localidad){
    diagnosticos.forEach(diag=>{
      const diagNorm = quitarAcentosGs_(diag);
      DIAGNOSTICO_AGENTE_MAP_.forEach(regla=>{
        if(!regla.palabras.some(pal=>diagNorm.includes(pal))) return;
        empresas.forEach(e=>{
          if(quitarAcentosGs_(e.localidad)!==localidadNorm) return;
          const agentesEmpresa = [].concat(e.agentes_quimicos, e.agentes_fisicos, e.agentes_biologicos);
          const coincidentes = agentesEmpresa.filter(a=>regla.agentes.indexOf(a)>-1);
          if(coincidentes.length) hallazgos.push({categoria:regla.categoria, diagnostico:diag, empresa:e.nombre, agentes:coincidentes});
        });
      });
    });
  }

  // 2) Vía hídrica — mismo canal de riego/afluente, SIN importar la
  // localidad: los efluentes viajan por el canal y pueden llegar a otra
  // zona aguas abajo. No depende de la tabla de diagnósticos-OIT: alcanza
  // con compartir el afluente para que sea una alerta a revisar.
  const hallazgos_hidricos = [];
  if(afluente){
    const afluenteNorm = quitarAcentosGs_(afluente);
    empresas.forEach(e=>{
      if(!e.afluente || quitarAcentosGs_(e.afluente)!==afluenteNorm) return;
      const categorias = [];
      if(e.afluente_toxico) categorias.push('Tóxico (vía hídrica/riego)');
      if(e.afluente_organico) categorias.push('Orgánico (vía hídrica/riego)');
      if(!categorias.length) categorias.push('Vía hídrica/riego (sin clasificar)');
      hallazgos_hidricos.push({categoria:categorias.join(' + '), empresa:e.nombre, afluente:e.afluente});
    });
  }

  return{success:true, localidad, afluente, hallazgos, hallazgos_hidricos};
}
// Compuerta de permisos — la usa doGet (get_investigacion_paciente).
function getInvestigacionPacienteGateado_(dni, profesionalEmail){
  if(!tieneModuloAutorizado_(profesionalEmail, 'investigacion')) return{success:false, error:'No autorizado para el módulo de Investigación'};
  return getInvestigacionPaciente(dni);
}

// ══════════════════════════════════════════════════════════════════
// PARCHE — Sincronización del Libro de Discapacidad (CUD) a la nube
// Antes el Libro de Discapacidad (las 9 páginas, incluida la de CUD) se
// guardaba SOLO en localStorage del navegador de OCM — nunca llegaba a la
// hoja Discapacidad_CUD, así que la vista consolidada que ya usa KSM
// quedaba estructuralmente lista pero vacía en la práctica. Esto agrega
// el guardado real, con upsert por DNI (una fila por paciente — se
// actualiza en vez de duplicar cada vez que el médico toca "Guardar").
// ══════════════════════════════════════════════════════════════════
function guardarLibroDiscapacidadCUD(d){
  if(!d || !d.dni) return{success:false, error:'Falta DNI'};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Discapacidad_CUD');
  if(!sh){
    sh = ss.insertSheet('Discapacidad_CUD');
    sh.appendRow(['Timestamp','DNI','Nombre','Telefono','Fecha_Nac','Diagnostico','Codigos_CIF','Tipo',
                  'Subtipo','Modalidad','Orientacion','Fecha_Emision','Fecha_Vencimiento','Junta',
                  'Firmantes_JSON','Efectores','Tipo_Pension','Trazabilidad']);
  }
  let headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  // 'Origen' distingue este registro (viene del Libro de Discapacidad de OCM)
  // del que arma el Lector CUD. 'Localidad' y 'Afluente' vienen de la página 1
  // del Libro — son la base del cruce de Investigación (OIT + vía hídrica).
  const faltantes = ['Origen','Localidad','Afluente'].filter(h=>headers.indexOf(h)===-1);
  if(faltantes.length){
    faltantes.forEach(h=> sh.getRange(1, sh.getLastColumn()+1).setValue(h));
    headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  }
  const data = sh.getDataRange().getValues();
  const dniCol = headers.indexOf('DNI'), origenCol = headers.indexOf('Origen');
  const valores = {
    Timestamp: new Date().toISOString(), DNI: d.dni, Nombre: d.nombre||'',
    Diagnostico: d.diagnostico||'', Codigos_CIF: d.codigos_cif||'', Tipo: d.tipo||'',
    Subtipo: d.subtipo||'', Fecha_Emision: d.emision||'', Fecha_Vencimiento: d.vencimiento||'',
    Localidad: d.localidad||'', Afluente: d.afluente||'',
    Origen: 'libro_discapacidad_ocm'
  };
  for(let i=1;i<data.length;i++){
    if(String(data[i][dniCol])===String(d.dni) && String(data[i][origenCol])==='libro_discapacidad_ocm'){
      const fila = data[i].slice();
      headers.forEach((h,idx)=>{ if(valores[h]!==undefined) fila[idx]=valores[h]; });
      sh.getRange(i+1,1,1,fila.length).setValues([fila]);
      return{success:true, message:'Libro de Discapacidad actualizado en la nube'};
    }
  }
  const filaNueva = headers.map(h=> valores[h]!==undefined ? valores[h] : '');
  sh.appendRow(filaNueva);
  return{success:true, message:'Libro de Discapacidad sincronizado a la nube por primera vez'};
}
// Compuerta de permisos — la usa doGet (get_discapacidad_paciente). No
// reutiliza get_vista_consolidada tal cual porque esa acción la sigue
// usando OCM sin mail de profesional — esta es la puerta gateada para KSM.
function getDiscapacidadPacienteGateado_(dni, profesionalEmail){
  if(!tieneModuloAutorizado_(profesionalEmail, 'cud')) return{success:false, error:'No autorizado para el módulo de Discapacidad/CUD'};
  if(!dni) return{success:false, error:'Falta DNI'};
  return{success:true, discapacidad: construirVistaConsolidada(dni).discapacidad};
}

// ══════════════════════════════════════════════════════════════════
// PARCHE — Sincronización de Plantillas (para compartir SOLO estado y
// detalle clínico con especialistas autorizados — a propósito, sin datos
// de pago). Igual que Catastro y el Libro de Discapacidad, Plantillas
// vivía nada más que en IndexedDB local del navegador de OCM.
// ══════════════════════════════════════════════════════════════════
function sheetPlantillasSync_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Plantillas_Sync');
  if(!sh){
    sh = ss.insertSheet('Plantillas_Sync');
    sh.appendRow(['Timestamp','DNI','Paciente','NSerie','Estado','Detalle']);
    sh.getRange(1,1,1,6).setBackground('#334155').setFontColor('white').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
// La llama OCM al crear el pedido y al cambiar de estado (cambiarEstPlant).
function sincronizarPlantilla(d){
  if(!d || !d.dni || !d.nSerie) return{success:false, error:'Falta DNI o N° de pedido'};
  const sh = sheetPlantillasSync_();
  const data = sh.getDataRange().getValues();
  for(let i=1;i<data.length;i++){
    if(String(data[i][3])===String(d.nSerie)){
      sh.getRange(i+1,1,1,6).setValues([[new Date().toISOString(), d.dni, d.paciente||'', d.nSerie, d.estado||'', d.detalle||'']]);
      return{success:true};
    }
  }
  sh.appendRow([new Date().toISOString(), d.dni, d.paciente||'', d.nSerie, d.estado||'', d.detalle||'']);
  return{success:true};
}
function getPlantillasPaciente(dni){
  if(!dni) return{success:true, plantillas:[]};
  const sh = sheetPlantillasSync_();
  const data = sh.getDataRange().getValues();
  const plantillas = [];
  for(let i=1;i<data.length;i++){
    if(String(data[i][1])===String(dni)) plantillas.push({nSerie:data[i][3], estado:data[i][4], detalle:data[i][5], fecha:data[i][0]});
  }
  return{success:true, plantillas};
}
// Compuerta de permisos — la usa doGet (get_plantillas_paciente).
function getPlantillasPacienteGateado_(dni, profesionalEmail){
  if(!tieneModuloAutorizado_(profesionalEmail, 'plantillas')) return{success:false, error:'No autorizado para el módulo de Plantillas'};
  return getPlantillasPaciente(dni);
}

// ══════════════════════════════════════════════════════════════════
// Listas personalizadas — usado por CAM (Herramientas Clínicas):
// diccionario propio de síndromes/ítems que el médico va agregando,
// guardado en la nube (misma base que el resto de S.I.M-M) y disponible
// para siempre en cualquier dispositivo. Genérico por "tipo" para poder
// reutilizarse con otras listas futuras (instituciones, procedimientos,
// códigos, etc.) sin tener que tocar el backend de nuevo.
// ══════════════════════════════════════════════════════════════════
function hojaListasPersonalizadas_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Listas_Personalizadas');
  if(!sheet){
    sheet = ss.insertSheet('Listas_Personalizadas');
    sheet.appendRow(['tipo','valor','fecha']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
function getListaQuirurgico(tipo){
  const sheet = hojaListasPersonalizadas_();
  const rows = sheet.getDataRange().getValues();
  const data = [];
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][0])===String(tipo)) data.push(rows[i][1]);
  }
  return {success:true, data};
}
function guardarItemQuirurgico(d){
  if(!d || !d.tipo || !d.valor) return {success:false, error:'Falta tipo o valor'};
  const sheet = hojaListasPersonalizadas_();
  sheet.appendRow([d.tipo, d.valor, new Date().toISOString()]);
  return {success:true};
}
