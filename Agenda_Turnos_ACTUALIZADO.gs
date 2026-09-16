// ══════════════════════════════════════════════════════════════════
// AGENDA DE TURNOS — disponibilidad y reserva real de turnos presenciales,
// compartida entre la Consulta Virtual (Agenda_Mensajes.html, con token) y
// Turnos_Secretaria.html en cada sede (sin token — ver comentario en
// ACCIONES_POST_PROTEGIDAS_). Todo pasa por la misma hoja Turnos_Agenda,
// así nunca se pisan los turnos que carga el médico con los que cargan las
// secretarias.
// ══════════════════════════════════════════════════════════════════

// Duración de turno en minutos. DURACION_TURNO_MIN_ es el valor por
// defecto; DURACION_TURNO_MIN_POR_CONSULTORIO_ permite pisarlo para un
// consultorio puntual (hoy solo Palmares, que trabaja cada 15 minutos).
const DURACION_TURNO_MIN_ = 20;
const DURACION_TURNO_MIN_POR_CONSULTORIO_ = {
  'Centro Médico Palmares': 15
  // Rehabilitarte Instituto Médico y Consultorios Privados del Norte no
  // están acá, así que usan los 20 minutos por defecto de arriba.
};
function duracionTurno_(consultorio){
  return DURACION_TURNO_MIN_POR_CONSULTORIO_[consultorio] || DURACION_TURNO_MIN_;
}

const SEMANAS_AGENDA_ = 4;

// dia: 0=domingo … 6=sábado, igual que Date.getDay()
const CONSULTORIOS_HORARIOS_ = {
  'Centro Médico Palmares': [
    {dia:2, desde:'13:30', hasta:'16:45'},
    {dia:4, desde:'13:30', hasta:'20:00'}
  ],
  'Rehabilitarte Instituto Médico': [
    {dia:5, desde:'16:00', hasta:'18:00'}
  ],
  'Consultorios Privados del Norte': [
    {dia:5, desde:'18:00', hasta:'20:30'}
  ]
};

function hojaTurnosAgenda_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Turnos_Agenda');
  if(!sheet){
    sheet = ss.insertSheet('Turnos_Agenda');
    sheet.appendRow(['ID_Turno','Timestamp_Reserva','Consultorio','Fecha','Hora','Duracion_Min',
      'DNI','Nombre','Telefono','Email','Origen','Estado','Notas']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function minutosDesdeHHMM_(hhmm){
  const partes = String(hhmm||'0:0').split(':');
  return (parseInt(partes[0],10)||0)*60 + (parseInt(partes[1],10)||0);
}

function fechaYYYYMMDD_(d){
  const y = d.getFullYear();
  const m = String(d.getMonth()+1).padStart(2,'0');
  const dia = String(d.getDate()).padStart(2,'0');
  return y + '-' + m + '-' + dia;
}

// Calcula los horarios libres de uno o todos los consultorios para las
// próximas SEMANAS_AGENDA_ semanas. Excluye los horarios ya ocupados
// (cualquier fila con Estado distinto de 'cancelado') y, si es el día de
// hoy, los horarios que ya pasaron.
function getTurnosDisponibles(params){
  params = params || {};
  const consultorioFiltro = params.consultorio;
  const nombres = consultorioFiltro ? [consultorioFiltro] : Object.keys(CONSULTORIOS_HORARIOS_);

  const sheet = hojaTurnosAgenda_();
  const filas = sheet.getDataRange().getValues();
  const ocupados = {}; // clave: consultorio+'|'+fecha+'|'+hora
  for(let i=1;i<filas.length;i++){
    const estado = String(filas[i][11]||'').toLowerCase();
    if(estado === 'cancelado') continue;
    const clave = filas[i][2] + '|' + filas[i][3] + '|' + filas[i][4];
    ocupados[clave] = true;
  }

  const ahora = new Date();
  const minutosAhora = ahora.getHours()*60 + ahora.getMinutes();
  const resultado = {};

  nombres.forEach(nombre => {
    const horariosSemana = CONSULTORIOS_HORARIOS_[nombre];
    if(!horariosSemana) return;
    const duracion = duracionTurno_(nombre);
    const dias = [];

    for(let offset=0; offset < SEMANAS_AGENDA_*7; offset++){
      const fecha = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate()+offset);
      const bloques = horariosSemana.filter(b => b.dia === fecha.getDay());
      if(!bloques.length) continue;

      const fechaStr = fechaYYYYMMDD_(fecha);
      const horarios = [];
      bloques.forEach(bloque => {
        const desdeMin = minutosDesdeHHMM_(bloque.desde);
        const hastaMin = minutosDesdeHHMM_(bloque.hasta);
        for(let m = desdeMin; m + duracion <= hastaMin; m += duracion){
          if(offset === 0 && m <= minutosAhora) continue; // hoy: no ofrecer horarios ya pasados
          const hh = String(Math.floor(m/60)).padStart(2,'0');
          const mm = String(m%60).padStart(2,'0');
          const horaStr = hh + ':' + mm;
          const clave = nombre + '|' + fechaStr + '|' + horaStr;
          if(ocupados[clave]) continue;
          horarios.push(horaStr);
        }
      });

      if(horarios.length) dias.push({fecha: fechaStr, horarios});
    }

    resultado[nombre] = dias;
  });

  return { success:true, consultorios: resultado };
}

function reservarTurno(d){
  if(!d || !String(d.consultorio||'').trim()) return {success:false, error:'Falta el consultorio'};
  if(!CONSULTORIOS_HORARIOS_[d.consultorio]) return {success:false, error:'Consultorio no reconocido'};
  if(!String(d.fecha||'').trim() || !String(d.hora||'').trim()) return {success:false, error:'Falta la fecha o el horario'};
  if(!String(d.nombre||'').trim()) return {success:false, error:'Falta el nombre del paciente'};
  if(!String(d.telefono||'').trim() && !String(d.email||'').trim()){
    return {success:false, error:'Dejá al menos un teléfono o un mail de contacto del paciente'};
  }

  const sheet = hojaTurnosAgenda_();
  const filas = sheet.getDataRange().getValues();
  for(let i=1;i<filas.length;i++){
    const estado = String(filas[i][11]||'').toLowerCase();
    if(estado === 'cancelado') continue;
    if(filas[i][2]===d.consultorio && filas[i][3]===d.fecha && filas[i][4]===d.hora){
      return {success:false, error:'Ese horario ya se acaba de reservar — elegí otro.'};
    }
  }

  const id = Utilities.getUuid();
  sheet.appendRow([
    id, new Date().toISOString(), d.consultorio, d.fecha, d.hora, duracionTurno_(d.consultorio),
    d.dni||'', d.nombre||'', d.telefono||'', d.email||'', d.origen||'Sin especificar',
    'confirmado', d.notas||''
  ]);

  if(String(d.email||'').trim()){
    try{
      MailApp.sendEmail(d.email,
        'Turno confirmado — Dr. Muratori',
        'Hola ' + (d.nombre||'') + ',\n\n' +
        'Tu turno quedó confirmado:\n\n' +
        'Lugar: ' + d.consultorio + '\n' +
        'Fecha: ' + d.fecha + '\n' +
        'Hora: ' + d.hora + ' hs\n\n' +
        (d.notas ? ('Notas: ' + d.notas + '\n\n') : '') +
        'Dr. Luis Alberto Muratori — M.N. 100.540 · M.P. 9943',
        { name: 'Dr. Luis Alberto Muratori' });
    }catch(e){
      // El turno ya quedó guardado — que falle el mail de confirmación no
      // debe hacer perder la reserva.
    }
  }

  return { success:true, id: id };
}

// Libera un turno (por si el paciente cancela o hay que reprogramar) — no
// borra la fila, la marca 'cancelado' para que el horario vuelva a
// aparecer como disponible sin perder el registro de lo que pasó.
function cancelarTurno(d){
  if(!d || !String(d.idTurno||'').trim()) return {success:false, error:'Falta el ID del turno'};
  const sheet = hojaTurnosAgenda_();
  const filas = sheet.getDataRange().getValues();
  for(let i=1;i<filas.length;i++){
    if(filas[i][0] === d.idTurno){
      sheet.getRange(i+1, 12).setValue('cancelado'); // columna 12 = Estado
      return {success:true};
    }
  }
  return {success:false, error:'No se encontró ese turno'};
}
