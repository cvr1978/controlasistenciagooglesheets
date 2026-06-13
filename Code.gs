// ============================================================
// SISTEMA DE CONTROL DE ASISTENCIA CON QR - Google Apps Script
// ============================================================

// Zona horaria fija para Guatemala (UTC-6, sin horario de verano)
const TZ = 'America/Guatemala';

// Nombres de las hojas del spreadsheet
const SHEET_ESTUDIANTES  = 'Estudiantes';
const SHEET_ASISTENCIA   = 'Asistencia';
const SHEET_CONFIG       = 'Configuracion';
const SHEET_RESUMEN      = 'Resumen';
const SHEET_ESTADISTICAS = 'Estadísticas';

// Columnas hoja Estudiantes: ID | Nombre | Carrera | Email | Estado | Grado | Sección | Sexo
const COL_EST_ID       = 1;
const COL_EST_NOMBRE   = 2;
const COL_EST_PROGRAMA = 3;
const COL_EST_EMAIL    = 4;
const COL_EST_ESTADO   = 5;
const COL_EST_GRADO    = 6;
const COL_EST_SECCION  = 7;
const COL_EST_SEXO     = 8;

// Columnas hoja Asistencia: Fecha | Hora | ID | Nombre | Carrera | Grado | Sección | Materia | Observación | Sexo
const COL_ASIST_FECHA    = 1;
const COL_ASIST_HORA     = 2;
const COL_ASIST_ID       = 3;
const COL_ASIST_NOMBRE   = 4;
const COL_ASIST_PROGRAMA = 5;
const COL_ASIST_GRADO    = 6;
const COL_ASIST_SECCION  = 7;
const COL_ASIST_MATERIA  = 8;
const COL_ASIST_OBS      = 9;
const COL_ASIST_SEXO     = 10;
const NUM_COLS_ASIST     = 10;

// ------------------------------------------------------------
// PUNTO DE ENTRADA WEB APP
// ------------------------------------------------------------

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'scanner';

  if (page === 'dashboard') {
    return HtmlService.createTemplateFromFile('Dashboard')
      .evaluate()
      .setTitle('Dashboard Asistencia')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === 'estudiantes') {
    return HtmlService.createTemplateFromFile('Estudiantes')
      .evaluate()
      .setTitle('Gestión de Estudiantes')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === 'lista') {
    return HtmlService.createTemplateFromFile('ListaAsistencia')
      .evaluate()
      .setTitle('Lista de Asistencia')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Control de Asistencia QR')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// ------------------------------------------------------------
// INICIALIZACIÓN DEL SPREADSHEET
// ------------------------------------------------------------

function inicializarSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ---- Hoja Configuracion ----
  let sheetConfig = ss.getSheetByName(SHEET_CONFIG);
  if (!sheetConfig) {
    sheetConfig = ss.insertSheet(SHEET_CONFIG);
    sheetConfig.getRange('A1:B1').setValues([['Clave', 'Valor']]);
    sheetConfig.getRange('A2:B6').setValues([
      ['Institucion', 'Mi Institución Educativa'],
      ['Materia',     'Materia General'],
      ['Docente',     'Nombre del Docente'],
      ['Tolerancia',  '15'],
      ['ColorHeader', '#1a73e8']
    ]);
    sheetConfig.getRange('A1:B1').setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    sheetConfig.setFrozenRows(1);
  }

  // ---- Hoja Estudiantes ----
  let sheetEst = ss.getSheetByName(SHEET_ESTUDIANTES);
  if (!sheetEst) {
    sheetEst = ss.insertSheet(SHEET_ESTUDIANTES);
  }
  // Siempre actualizar encabezados (permite migración)
  const headersEst = ['ID / Código', 'Nombre Completo', 'Carrera / Programa', 'Email', 'Estado', 'Grado', 'Sección', 'Sexo'];
  sheetEst.getRange(1, 1, 1, headersEst.length).setValues([headersEst]);
  sheetEst.getRange(1, 1, 1, headersEst.length)
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sheetEst.setFrozenRows(1);
  [120, 220, 200, 200, 90, 90, 100, 80].forEach((w, i) => sheetEst.setColumnWidth(i + 1, w));

  // ---- Hoja Asistencia ----
  let sheetAsist = ss.getSheetByName(SHEET_ASISTENCIA);
  if (!sheetAsist) {
    sheetAsist = ss.insertSheet(SHEET_ASISTENCIA);
  }
  const headersAsist = ['Fecha', 'Hora', 'ID / Código', 'Nombre Completo', 'Carrera', 'Grado', 'Sección', 'Materia', 'Observación', 'Sexo'];
  sheetAsist.getRange(1, 1, 1, headersAsist.length).setValues([headersAsist]);
  sheetAsist.getRange(1, 1, 1, headersAsist.length)
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sheetAsist.setFrozenRows(1);
  [110, 80, 120, 220, 190, 80, 90, 160, 110, 80].forEach((w, i) => sheetAsist.setColumnWidth(i + 1, w));

  // ---- Hoja Resumen ----
  _obtenerOCrearHojaResumen(ss);

  // ---- Hoja Estadísticas ----
  _setupHojaEstadisticas(ss);

  SpreadsheetApp.getUi().alert('✅ Spreadsheet inicializado correctamente.');
}

// ------------------------------------------------------------
// REGISTRO DE ASISTENCIA POR QR
// ------------------------------------------------------------

function registrarAsistencia(codigoQr) {
  try {
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const sheetE = ss.getSheetByName(SHEET_ESTUDIANTES);
    const sheetA = ss.getSheetByName(SHEET_ASISTENCIA);
    const config = obtenerConfiguracion();

    if (!sheetE || !sheetA) {
      return { ok: false, mensaje: 'El spreadsheet no está inicializado.' };
    }

    const codigoLimpio = codigoQr.toString().trim();
    const est = _buscarEstudiante(sheetE, codigoLimpio);

    if (!est) {
      return { ok: false, mensaje: `Código "${codigoLimpio}" no encontrado.` };
    }
    if (est.estado.toString().toLowerCase() === 'inactivo') {
      return { ok: false, mensaje: `${est.nombre} está marcado como Inactivo.` };
    }

    const ahora  = new Date();
    const hoyStr = Utilities.formatDate(ahora, TZ, 'yyyy-MM-dd');

    if (_yaRegistroHoy(sheetA, codigoLimpio, hoyStr)) {
      return { ok: false, mensaje: `${est.nombre} ya tiene asistencia registrada hoy.`, estudiante: est };
    }

    const tolerancia = parseInt(config['Tolerancia'] || 15);
    const horaClase  = parseInt(config['HoraClaseH'] || 7) * 60 + parseInt(config['HoraClaseM'] || 0);
    const horaActual = ahora.getHours() * 60 + ahora.getMinutes();
    const obs        = horaActual <= horaClase + tolerancia ? 'Puntual' : 'Tarde';

    const fecha   = Utilities.formatDate(ahora, TZ, 'yyyy-MM-dd');
    const hora    = Utilities.formatDate(ahora, TZ, 'HH:mm:ss');
    const materia = config['Materia'] || 'General';

    _appendAsistencia(sheetA, fecha, hora, est, materia, obs);

    return {
      ok: true,
      mensaje: `✅ Asistencia registrada para ${est.nombre} — ${obs}`,
      estudiante: est,
      hora: hora,
      observacion: obs
    };

  } catch (err) {
    return { ok: false, mensaje: 'Error interno: ' + err.message };
  }
}

// ------------------------------------------------------------
// LISTA DE ASISTENCIA MASIVA (por grupo)
// ------------------------------------------------------------

/**
 * Devuelve los valores únicos de Carrera, Grado y Sección para poblar los filtros.
 */
function obtenerFiltrosDisponibles() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ESTUDIANTES);
  if (!sheet || sheet.getLastRow() < 2) return { carreras: [], grados: [], secciones: [] };

  const datos = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  const carreras  = new Set();
  const grados    = new Set();
  const secciones = new Set();

  datos.forEach(r => {
    if (r[0] === '') return;
    if (r[COL_EST_PROGRAMA - 1]) carreras.add(r[COL_EST_PROGRAMA - 1].toString().trim());
    if (r[COL_EST_GRADO   - 1]) grados.add(r[COL_EST_GRADO   - 1].toString().trim());
    if (r[COL_EST_SECCION - 1]) secciones.add(r[COL_EST_SECCION - 1].toString().trim());
  });

  return {
    carreras:  [...carreras].sort(),
    grados:    [...grados].sort(),
    secciones: [...secciones].sort()
  };
}

/**
 * Retorna los estudiantes que coincidan con los filtros,
 * junto con su estado de asistencia para la fecha indicada.
 * Filtros vacíos o "Todos" = sin filtrar.
 */
function obtenerEstudiantesPorGrupo(carrera, grado, seccion, fecha) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheetE = ss.getSheetByName(SHEET_ESTUDIANTES);
  const sheetA = ss.getSheetByName(SHEET_ASISTENCIA);

  if (!sheetE || sheetE.getLastRow() < 2) return [];

  const fechaBuscar = fecha || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');

  // Índice de asistencia del día: id -> observacion
  const asistHoy = {};
  if (sheetA && sheetA.getLastRow() > 1) {
    const datosA = sheetA.getRange(2, 1, sheetA.getLastRow() - 1, NUM_COLS_ASIST).getValues();
    datosA.forEach(r => {
      const f = r[COL_ASIST_FECHA - 1];
      if (!f) return;
      const fStr = _toDateStr(f);
      if (fStr === fechaBuscar) {
        asistHoy[r[COL_ASIST_ID - 1].toString().trim()] = r[COL_ASIST_OBS - 1].toString();
      }
    });
  }

  const datosE = sheetE.getRange(2, 1, sheetE.getLastRow() - 1, 8).getValues();
  const resultado = [];

  datosE.forEach(r => {
    if (r[0] === '') return;
    if (r[COL_EST_ESTADO - 1].toString().toLowerCase() === 'inactivo') return;

    const eCarrera  = r[COL_EST_PROGRAMA - 1].toString().trim();
    const eGrado    = r[COL_EST_GRADO    - 1].toString().trim();
    const eSeccion  = r[COL_EST_SECCION  - 1].toString().trim();

    if (carrera  && carrera  !== 'Todos' && eCarrera  !== carrera)  return;
    if (grado    && grado    !== 'Todos' && eGrado    !== grado)    return;
    if (seccion  && seccion  !== 'Todos' && eSeccion  !== seccion)  return;

    const id = r[COL_EST_ID - 1].toString().trim();
    resultado.push({
      id:        id,
      nombre:    r[COL_EST_NOMBRE   - 1],
      carrera:   eCarrera,
      grado:     eGrado,
      seccion:   eSeccion,
      email:     r[COL_EST_EMAIL    - 1],
      sexo:      r[COL_EST_SEXO     - 1].toString().trim(),
      estadoHoy: asistHoy[id] || null
    });
  });

  resultado.sort((a, b) => a.nombre.localeCompare(b.nombre));
  return resultado;
}

/**
 * Registra o ACTUALIZA la asistencia masiva para un grupo.
 * registros: [{ id, observacion }]  observacion: 'Puntual' | 'Tarde' | 'Ausente'
 * Si el estudiante ya tiene registro en esa fecha, actualiza su observación.
 */
function registrarAsistenciaMasiva(registros, materia, fecha) {
  try {
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const sheetE = ss.getSheetByName(SHEET_ESTUDIANTES);
    const sheetA = ss.getSheetByName(SHEET_ASISTENCIA);

    if (!sheetE || !sheetA) {
      return { ok: false, mensaje: 'El spreadsheet no está inicializado.' };
    }

    const tz         = TZ;
    const fechaUso   = fecha || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const horaUso    = Utilities.formatDate(new Date(), tz, 'HH:mm:ss');
    const materiaUso = materia || obtenerConfiguracion()['Materia'] || 'General';

    // Construir mapa: id -> número de fila en sheetA (1-based) para esa fecha
    const filaExistente = {};
    if (sheetA.getLastRow() > 1) {
      const datosA = sheetA.getRange(2, 1, sheetA.getLastRow() - 1, NUM_COLS_ASIST).getValues();
      datosA.forEach((r, idx) => {
        const f = r[COL_ASIST_FECHA - 1];
        if (!f) return;
        const fStr = _toDateStr(f);
        if (fStr === fechaUso) {
          filaExistente[r[COL_ASIST_ID - 1].toString().trim()] = idx + 2; // +2: cabecera + índice base-0
        }
      });
    }

    // Índice de estudiantes (8 columnas para incluir Sexo)
    const mapaEst = {};
    if (sheetE.getLastRow() > 1) {
      sheetE.getRange(2, 1, sheetE.getLastRow() - 1, 8).getValues().forEach(r => {
        if (r[0] !== '') mapaEst[r[0].toString().trim()] = r;
      });
    }

    let nuevos      = 0;
    let actualizados = 0;
    // Para el resumen global por sexo
    const resumenData = { totalH:0, totalM:0, presentesH:0, presentesM:0, ausentesH:0, ausentesM:0, puntuales:0, tarde:0 };
    // Para el resumen agrupado por grado/sección (se escribe en la hoja Resumen)
    const grupoResumen = {};

    registros.forEach(reg => {
      const id  = reg.id.toString().trim();
      const obs = reg.observacion || 'Ausente';
      const color = obs === 'Puntual' ? '#d9ead3' : obs === 'Tarde' ? '#fff2cc' : '#fce8e6';
      const filaE = mapaEst[id];
      const sexo    = filaE ? filaE[COL_EST_SEXO     - 1].toString().trim().toUpperCase() : '';
      const grado   = filaE ? filaE[COL_EST_GRADO    - 1].toString().trim() : '';
      const seccion = filaE ? filaE[COL_EST_SECCION  - 1].toString().trim() : '';
      const carrera = filaE ? filaE[COL_EST_PROGRAMA - 1].toString().trim() : '';
      const esH   = sexo === 'M';
      const esM   = sexo === 'F';

      // Acumular resumen global
      if (esH) resumenData.totalH++;
      else if (esM) resumenData.totalM++;
      if (obs !== 'Ausente') {
        if (esH) resumenData.presentesH++; else if (esM) resumenData.presentesM++;
        if (obs === 'Puntual') resumenData.puntuales++; else resumenData.tarde++;
      } else {
        if (esH) resumenData.ausentesH++; else if (esM) resumenData.ausentesM++;
      }

      // Acumular resumen por grado
      if (filaE) {
        const gKey = `${carrera}|${grado}|${seccion}`;
        if (!grupoResumen[gKey]) {
          grupoResumen[gKey] = { carrera, grado, seccion,
            totalH:0, totalM:0, presentesH:0, presentesM:0,
            ausentesH:0, ausentesM:0, puntuales:0, tarde:0 };
        }
        const gr = grupoResumen[gKey];
        if (esH) gr.totalH++; else if (esM) gr.totalM++;
        if (obs !== 'Ausente') {
          if (esH) gr.presentesH++; else if (esM) gr.presentesM++;
          if (obs === 'Puntual') gr.puntuales++; else gr.tarde++;
        } else {
          if (esH) gr.ausentesH++; else if (esM) gr.ausentesM++;
        }
      }

      if (filaExistente[id]) {
        const fila = filaExistente[id];
        sheetA.getRange(fila, COL_ASIST_MATERIA).setValue(materiaUso);
        sheetA.getRange(fila, COL_ASIST_OBS).setValue(obs);
        sheetA.getRange(fila, 1, 1, NUM_COLS_ASIST).setBackground(color);
        actualizados++;
        return;
      }

      if (!filaE) return;
      const est = {
        id:       filaE[COL_EST_ID       - 1],
        nombre:   filaE[COL_EST_NOMBRE   - 1],
        programa: filaE[COL_EST_PROGRAMA - 1],
        grado:    filaE[COL_EST_GRADO    - 1],
        seccion:  filaE[COL_EST_SECCION  - 1],
        sexo:     sexo
      };
      _appendAsistencia(sheetA, fechaUso, horaUso, est, materiaUso, obs);
      nuevos++;
    });

    const total     = resumenData.totalH + resumenData.totalM;
    const presentes = resumenData.presentesH + resumenData.presentesM;
    const ausentes  = resumenData.ausentesH  + resumenData.ausentesM;

    // Escribir resumen por grado en la hoja Resumen
    _escribirResumenHoja(ss, fechaUso, horaUso, materiaUso, grupoResumen);

    return {
      ok: true,
      mensaje: `✅ Asistencia guardada: ${nuevos} nuevos, ${actualizados} actualizados.`,
      resumen: {
        fecha:      fechaUso,
        materia:    materiaUso,
        totalH:     resumenData.totalH,
        totalM:     resumenData.totalM,
        total:      total,
        presentesH: resumenData.presentesH,
        presentesM: resumenData.presentesM,
        presentes:  presentes,
        ausentesH:  resumenData.ausentesH,
        ausentesM:  resumenData.ausentesM,
        ausentes:   ausentes,
        puntuales:  resumenData.puntuales,
        tarde:      resumenData.tarde
      }
    };

  } catch (err) {
    return { ok: false, mensaje: 'Error: ' + err.message };
  }
}

// ------------------------------------------------------------
// GESTIÓN DE ESTUDIANTES
// ------------------------------------------------------------

function obtenerEstudiantes() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ESTUDIANTES);
  if (!sheet || sheet.getLastRow() < 2) return [];

  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues()
    .filter(r => r[0] !== '')
    .map(r => ({
      id:       r[COL_EST_ID       - 1],
      nombre:   r[COL_EST_NOMBRE   - 1],
      programa: r[COL_EST_PROGRAMA - 1],
      email:    r[COL_EST_EMAIL    - 1],
      estado:   r[COL_EST_ESTADO   - 1],
      grado:    r[COL_EST_GRADO    - 1],
      seccion:  r[COL_EST_SECCION  - 1],
      sexo:     r[COL_EST_SEXO     - 1]
    }));
}

function agregarEstudiante(datos) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ESTUDIANTES);

    const existentes = sheet.getDataRange().getValues();
    for (let i = 1; i < existentes.length; i++) {
      if (existentes[i][0].toString().trim() === datos.id.toString().trim()) {
        return { ok: false, mensaje: 'Ya existe un estudiante con ese ID/Código.' };
      }
    }

    sheet.appendRow([
      datos.id, datos.nombre, datos.programa,
      datos.email || '', datos.estado || 'Activo',
      datos.grado || '', datos.seccion || '', datos.sexo || ''
    ]);
    return { ok: true, mensaje: `Estudiante ${datos.nombre} agregado correctamente.` };
  } catch (err) {
    return { ok: false, mensaje: 'Error: ' + err.message };
  }
}

function actualizarEstudiante(idOriginal, datos) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ESTUDIANTES);
    const filas = sheet.getDataRange().getValues();

    for (let i = 1; i < filas.length; i++) {
      if (filas[i][0].toString().trim() === idOriginal.toString().trim()) {
        sheet.getRange(i + 1, 1, 1, 8).setValues([[
          datos.id, datos.nombre, datos.programa,
          datos.email || '', datos.estado || 'Activo',
          datos.grado || '', datos.seccion || '', datos.sexo || ''
        ]]);
        return { ok: true, mensaje: 'Estudiante actualizado.' };
      }
    }
    return { ok: false, mensaje: 'Estudiante no encontrado.' };
  } catch (err) {
    return { ok: false, mensaje: 'Error: ' + err.message };
  }
}

function eliminarEstudiante(id) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ESTUDIANTES);
    const filas = sheet.getDataRange().getValues();

    for (let i = filas.length - 1; i >= 1; i--) {
      if (filas[i][0].toString().trim() === id.toString().trim()) {
        sheet.deleteRow(i + 1);
        return { ok: true, mensaje: 'Estudiante eliminado.' };
      }
    }
    return { ok: false, mensaje: 'Estudiante no encontrado.' };
  } catch (err) {
    return { ok: false, mensaje: 'Error: ' + err.message };
  }
}

// ------------------------------------------------------------
// DASHBOARD / REPORTES
// ------------------------------------------------------------

function obtenerResumenHoy() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { ok: false, errmsg: 'Spreadsheet no disponible' };

    const sheetA = ss.getSheetByName(SHEET_ASISTENCIA);
    const sheetE = ss.getSheetByName(SHEET_ESTUDIANTES);
    const hoyStr = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');

    // Índice de asistencia de hoy — hora se convierte a string para evitar
    // fallos de serialización en google.script.run con objetos Date anidados
    const asistHoy = {};
    if (sheetA && sheetA.getLastRow() > 1) {
      sheetA.getRange(2, 1, sheetA.getLastRow() - 1, NUM_COLS_ASIST).getValues().forEach(row => {
        const f = row[COL_ASIST_FECHA - 1];
        if (!f) return;
        if (_toDateStr(f) !== hoyStr) return;
        const h = row[COL_ASIST_HORA - 1];
        asistHoy[row[COL_ASIST_ID - 1].toString().trim()] = {
          hora:    h instanceof Date ? Utilities.formatDate(h, TZ, 'HH:mm') : String(h || ''),
          materia: String(row[COL_ASIST_MATERIA - 1] || ''),
          obs:     String(row[COL_ASIST_OBS     - 1] || '')
        };
      });
    }

    // Combinar todos los estudiantes activos con su asistencia de hoy
    const registros = [];
    if (sheetE && sheetE.getLastRow() > 1) {
      sheetE.getRange(2, 1, sheetE.getLastRow() - 1, 8).getValues().forEach(row => {
        if (!row[0]) return;
        if ((row[COL_EST_ESTADO - 1] || '').toString().toLowerCase() === 'inactivo') return;
        const id  = row[COL_EST_ID - 1].toString().trim();
        const att = asistHoy[id];
        registros.push({
          hora:     att ? att.hora    : '',
          id:       id,
          nombre:   String(row[COL_EST_NOMBRE   - 1] || ''),
          programa: String(row[COL_EST_PROGRAMA - 1] || ''),
          grado:    String(row[COL_EST_GRADO    - 1] || ''),
          seccion:  String(row[COL_EST_SECCION  - 1] || ''),
          materia:  att ? att.materia : '',
          obs:      att ? att.obs     : 'Ausente'
        });
      });
    }

    // Presentes primero (Puntual, Tarde), luego Ausentes
    registros.sort((a, b) => {
      const ord = { 'Puntual': 0, 'Tarde': 1, 'Ausente': 2 };
      return (ord[a.obs] !== undefined ? ord[a.obs] : 2) - (ord[b.obs] !== undefined ? ord[b.obs] : 2);
    });

    const presentes = registros.filter(r => r.obs !== 'Ausente').length;
    const puntuales = registros.filter(r => r.obs === 'Puntual').length;
    const tarde     = registros.filter(r => r.obs === 'Tarde').length;
    const ausentes  = registros.filter(r => r.obs === 'Ausente').length;

    return {
      ok: true,
      fecha: hoyStr,
      presentes, puntuales, tarde, ausentes,
      total: registros.length,
      registros
    };
  } catch (err) {
    return { ok: false, errmsg: err.message };
  }
}

/**
 * Devuelve estadísticas completas de asistencia para un estudiante específico.
 */
function obtenerEstadisticasEstudiante(id) {
  try {
    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const sheetE = ss.getSheetByName(SHEET_ESTUDIANTES);
    const sheetA = ss.getSheetByName(SHEET_ASISTENCIA);

    const est = _buscarEstudiante(sheetE, id.toString().trim());
    if (!est) return null;

    const registros = [];
    const porMes    = {};

    if (sheetA && sheetA.getLastRow() > 1) {
      sheetA.getRange(2, 1, sheetA.getLastRow() - 1, NUM_COLS_ASIST).getValues().forEach(r => {
        if (r[COL_ASIST_ID - 1].toString().trim() !== id.toString().trim()) return;
        const f = r[COL_ASIST_FECHA - 1];
        if (!f) return;
        const fStr = _toDateStr(f);
        const mes  = fStr.substring(0, 7);
        const obs  = r[COL_ASIST_OBS - 1].toString();
        const hora = r[COL_ASIST_HORA - 1];

        registros.push({
          fecha:   fStr,
          hora:    hora instanceof Date ? Utilities.formatDate(hora, TZ, 'HH:mm') : String(hora || '').substring(0, 5),
          materia: r[COL_ASIST_MATERIA - 1].toString(),
          grado:   r[COL_ASIST_GRADO   - 1].toString(),
          seccion: r[COL_ASIST_SECCION - 1].toString(),
          obs:     obs
        });

        if (!porMes[mes]) porMes[mes] = { total: 0, presentes: 0, puntuales: 0, tarde: 0, ausentes: 0 };
        porMes[mes].total++;
        if      (obs === 'Puntual') { porMes[mes].presentes++; porMes[mes].puntuales++; }
        else if (obs === 'Tarde')   { porMes[mes].presentes++; porMes[mes].tarde++;     }
        else                        { porMes[mes].ausentes++;                           }
      });
    }

    registros.sort((a, b) => b.fecha.localeCompare(a.fecha));

    const total     = registros.length;
    const presentes = registros.filter(r => r.obs !== 'Ausente').length;
    const puntuales = registros.filter(r => r.obs === 'Puntual').length;
    const tarde     = registros.filter(r => r.obs === 'Tarde').length;
    const ausentes  = registros.filter(r => r.obs === 'Ausente').length;
    const pct       = total > 0 ? Math.round(presentes / total * 100) : 0;

    return {
      estudiante: est,
      total, presentes, puntuales, tarde, ausentes, pct,
      porMes: Object.entries(porMes)
        .sort(([a], [b]) => a.localeCompare(b))
        .map(([mes, v]) => ({ mes, ...v })),
      registros
    };
  } catch (err) {
    return null;
  }
}

function obtenerAsistenciaPorFecha(fechaInicio, fechaFin) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheetA = ss.getSheetByName(SHEET_ASISTENCIA);
  if (!sheetA || sheetA.getLastRow() < 2) return [];

  const resultado = [];

  sheetA.getRange(2, 1, sheetA.getLastRow() - 1, NUM_COLS_ASIST).getValues().forEach(row => {
    const f = row[COL_ASIST_FECHA - 1];
    if (!f) return;
    const fStr = _toDateStr(f);
    if (fStr >= fechaInicio && fStr <= fechaFin) {
      const h = row[COL_ASIST_HORA - 1];
      resultado.push({
        fecha:    fStr,
        hora:     h instanceof Date ? Utilities.formatDate(h, TZ, 'HH:mm') : String(h || ''),
        id:       String(row[COL_ASIST_ID       - 1] || ''),
        nombre:   String(row[COL_ASIST_NOMBRE   - 1] || ''),
        programa: String(row[COL_ASIST_PROGRAMA - 1] || ''),
        grado:    String(row[COL_ASIST_GRADO    - 1] || ''),
        seccion:  String(row[COL_ASIST_SECCION  - 1] || ''),
        materia:  String(row[COL_ASIST_MATERIA  - 1] || ''),
        obs:      String(row[COL_ASIST_OBS      - 1] || '')
      });
    }
  });

  return resultado;
}

// ------------------------------------------------------------
// CONFIGURACIÓN
// ------------------------------------------------------------

function obtenerConfiguracion() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet || sheet.getLastRow() < 2) return {};

  const config = {};
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(r => {
    if (r[0]) config[r[0]] = r[1];
  });
  return config;
}

function guardarConfiguracion(datos) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_CONFIG);
    const filas = sheet.getDataRange().getValues();
    const claves = {};
    for (let i = 1; i < filas.length; i++) claves[filas[i][0]] = i + 1;

    Object.keys(datos).forEach(clave => {
      if (claves[clave]) sheet.getRange(claves[clave], 2).setValue(datos[clave]);
      else sheet.appendRow([clave, datos[clave]]);
    });
    return { ok: true, mensaje: 'Configuración guardada.' };
  } catch (err) {
    return { ok: false, mensaje: 'Error: ' + err.message };
  }
}

// ------------------------------------------------------------
// HELPERS INTERNOS
// ------------------------------------------------------------

/**
 * Convierte un valor de celda de fecha (Date object o string) a 'yyyy-MM-dd'
 * usando el timezone del script.
 * Evita el bug de new Date("2026-03-25") que parsea como UTC y da el día anterior.
 */
function _toDateStr(f) {
  if (!f) return '';
  if (f instanceof Date) {
    return Utilities.formatDate(f, TZ, 'yyyy-MM-dd');
  }
  // Es string — tomar solo los primeros 10 caracteres ("2026-03-25")
  return f.toString().substring(0, 10);
}

function _buscarEstudiante(sheet, id) {
  if (!sheet || sheet.getLastRow() < 2) return null;
  const datos = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  for (let i = 0; i < datos.length; i++) {
    if (datos[i][0].toString().trim() === id) {
      return {
        id:       datos[i][COL_EST_ID       - 1],
        nombre:   datos[i][COL_EST_NOMBRE   - 1],
        programa: datos[i][COL_EST_PROGRAMA - 1],
        email:    datos[i][COL_EST_EMAIL    - 1],
        estado:   datos[i][COL_EST_ESTADO   - 1],
        grado:    datos[i][COL_EST_GRADO    - 1],
        seccion:  datos[i][COL_EST_SECCION  - 1],
        sexo:     datos[i][COL_EST_SEXO     - 1].toString().trim()
      };
    }
  }
  return null;
}

function _yaRegistroHoy(sheet, id, hoyStr) {
  if (!sheet || sheet.getLastRow() < 2) return false;
  const datos = sheet.getRange(2, 1, sheet.getLastRow() - 1, NUM_COLS_ASIST).getValues();
  for (let i = 0; i < datos.length; i++) {
    const f = datos[i][COL_ASIST_FECHA - 1];
    if (!f) continue;
    if (datos[i][COL_ASIST_ID - 1].toString().trim() === id &&
        _toDateStr(f) === hoyStr) return true;
  }
  return false;
}

/**
 * Devuelve la hoja Resumen, creándola con encabezados si no existe.
 */
function _obtenerOCrearHojaResumen(ss) {
  let sheet = ss.getSheetByName(SHEET_RESUMEN);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_RESUMEN);
    const headers = [
      'Fecha', 'Hora', 'Carrera / Programa', 'Grado', 'Sección', 'Materia',
      'Total', 'Total H', 'Total M',
      'Presentes', 'Pres. H', 'Pres. M',
      'Ausentes', 'Aus. H', 'Aus. M',
      'Puntuales', 'Tardíos'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    [110, 80, 200, 80, 90, 160, 60, 60, 60, 80, 70, 70, 80, 70, 70, 85, 75]
      .forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }
  return sheet;
}

/**
 * Escribe una fila por cada grupo (grado/sección/carrera) en la hoja Resumen.
 */
function _escribirResumenHoja(ss, fecha, hora, materia, grupoResumen) {
  try {
    const sheet = _obtenerOCrearHojaResumen(ss);
    Object.values(grupoResumen).forEach(g => {
      const total     = g.totalH    + g.totalM;
      const presentes = g.presentesH + g.presentesM;
      const ausentes  = g.ausentesH  + g.ausentesM;
      sheet.appendRow([
        fecha, hora, g.carrera, g.grado, g.seccion, materia,
        total,    g.totalH,    g.totalM,
        presentes, g.presentesH, g.presentesM,
        ausentes,  g.ausentesH,  g.ausentesM,
        g.puntuales, g.tarde
      ]);
      const pct = total > 0 ? presentes / total : 0;
      const bg  = pct >= 0.8 ? '#d9ead3' : pct >= 0.5 ? '#fff2cc' : '#fce8e6';
      sheet.getRange(sheet.getLastRow(), 1, 1, 17).setBackground(bg);
    });
  } catch (err) {
    // No interrumpir el flujo principal si falla el resumen
    console.error('_escribirResumenHoja error: ' + err.message);
  }
}

function _appendAsistencia(sheet, fecha, hora, est, materia, obs) {
  sheet.appendRow([
    fecha, hora,
    est.id, est.nombre, est.programa || est.carrera || '',
    est.grado || '', est.seccion || '',
    materia, obs, est.sexo || ''
  ]);
  const col = obs === 'Puntual' ? '#d9ead3' : obs === 'Tarde' ? '#fff2cc' : '#fce8e6';
  sheet.getRange(sheet.getLastRow(), 1, 1, NUM_COLS_ASIST).setBackground(col);
}

// ------------------------------------------------------------
// HOJA ESTADÍSTICAS INDIVIDUAL
// ------------------------------------------------------------

/**
 * Trigger simple: se ejecuta cuando el usuario edita el desplegable de
 * la hoja Estadísticas (celda B2) para seleccionar un estudiante.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_ESTADISTICAS) return;
  if (e.range.getA1Notation() !== 'B2') return;
  const valor = (e.value || '').toString().trim();
  if (!valor || valor.startsWith('—')) {
    _limpiarAreaEstadisticas(sheet);
    return;
  }
  // El desplegable tiene formato "ID — Nombre"
  const id = valor.split(' — ')[0].trim();
  actualizarEstadisticasHoja(sheet, id);
}

/**
 * Crea (o reinicia) la hoja Estadísticas con el desplegable de estudiantes.
 * Llamar desde el menú: 🎓 Asistencia QR > 📊 Estadísticas por Estudiante.
 */
function crearHojaEstadisticas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _setupHojaEstadisticas(ss);
  SpreadsheetApp.getUi().alert(
    '✅ Hoja "Estadísticas" lista.\n\n' +
    'Usa el desplegable en la celda B2 para seleccionar un estudiante\n' +
    'y ver automáticamente sus estadísticas de asistencia.'
  );
}

/** Lógica interna de creación (sin alert, usable desde inicializarSpreadsheet). */
function _setupHojaEstadisticas(ss) {
  let sheet = ss.getSheetByName(SHEET_ESTADISTICAS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ESTADISTICAS);
  } else {
    sheet.clear();
  }

  // Anchos de columna
  [210, 230, 120, 110, 170, 110, 130].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.setRowHeight(1, 44);
  sheet.setRowHeight(2, 36);

  // ── Fila 1: Título ──
  sheet.getRange('A1:G1').merge()
    .setValue('📊  ESTADÍSTICAS DE ASISTENCIA INDIVIDUAL')
    .setBackground('#1a73e8').setFontColor('#ffffff')
    .setFontSize(15).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // ── Fila 2: Selector ──
  sheet.getRange('A2')
    .setValue('👤  Seleccionar Estudiante:')
    .setFontWeight('bold').setFontSize(11)
    .setBackground('#e8f0fe').setFontColor('#1a73e8')
    .setVerticalAlignment('middle');

  // Construir lista para el desplegable
  const sheetE = ss.getSheetByName(SHEET_ESTUDIANTES);
  const opciones = ['— Selecciona un estudiante —'];
  if (sheetE && sheetE.getLastRow() > 1) {
    sheetE.getRange(2, 1, sheetE.getLastRow() - 1, 2).getValues().forEach(r => {
      if (r[0]) opciones.push(r[0].toString().trim() + ' — ' + r[1].toString().trim());
    });
  }

  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(opciones, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('B2')
    .setDataValidation(regla)
    .setValue(opciones[0])
    .setBackground('#ffffff').setFontSize(11)
    .setBorder(true, true, true, true, false, false,
      '#1a73e8', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange('C2:G2').merge()
    .setValue('⬅  Elige un estudiante del desplegable para ver sus estadísticas')
    .setFontColor('#999999').setFontStyle('italic').setFontSize(10)
    .setVerticalAlignment('middle');

  sheet.setFrozenRows(2);
  return sheet;
}

/**
 * Rellena la hoja Estadísticas con los datos del estudiante indicado.
 */
function actualizarEstadisticasHoja(sheet, studentId) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const shE = ss.getSheetByName(SHEET_ESTUDIANTES);
  const shA = ss.getSheetByName(SHEET_ASISTENCIA);

  const est = _buscarEstudiante(shE, studentId);
  if (!est) return;

  // ── Calcular estadísticas ──
  const registros = [];
  const porMes    = {};

  if (shA && shA.getLastRow() > 1) {
    shA.getRange(2, 1, shA.getLastRow() - 1, NUM_COLS_ASIST).getValues().forEach(r => {
      if (r[COL_ASIST_ID - 1].toString().trim() !== studentId.toString().trim()) return;
      const f = r[COL_ASIST_FECHA - 1];
      if (!f) return;
      const fStr = _toDateStr(f);
      const mes  = fStr.substring(0, 7);
      const obs  = r[COL_ASIST_OBS     - 1].toString();
      const hora = r[COL_ASIST_HORA    - 1];
      registros.push({
        fecha:   fStr,
        hora:    hora ? hora.toString().substring(0, 5) : '',
        materia: r[COL_ASIST_MATERIA  - 1].toString(),
        grado:   r[COL_ASIST_GRADO    - 1].toString(),
        seccion: r[COL_ASIST_SECCION  - 1].toString(),
        obs:     obs
      });
      if (!porMes[mes]) porMes[mes] = { total:0, presentes:0, puntuales:0, tarde:0, ausentes:0 };
      porMes[mes].total++;
      if      (obs === 'Puntual') { porMes[mes].presentes++; porMes[mes].puntuales++; }
      else if (obs === 'Tarde')   { porMes[mes].presentes++; porMes[mes].tarde++;     }
      else                        { porMes[mes].ausentes++;                           }
    });
  }
  registros.sort((a, b) => b.fecha.localeCompare(a.fecha));

  const total     = registros.length;
  const presentes = registros.filter(r => r.obs !== 'Ausente').length;
  const puntuales = registros.filter(r => r.obs === 'Puntual').length;
  const tarde     = registros.filter(r => r.obs === 'Tarde').length;
  const ausentes  = registros.filter(r => r.obs === 'Ausente').length;
  const pct       = total > 0 ? Math.round(presentes / total * 100) : 0;
  const pctColor  = pct >= 80 ? '#2e7d32' : pct >= 60 ? '#e65100' : '#c62828';
  const meses     = Object.entries(porMes)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([mes, v]) => ({ mes, ...v }));

  // ── Limpiar área de contenido (fila 3 en adelante) ──
  _limpiarAreaEstadisticas(sheet);
  let fila = 3;

  // ── PERFIL ──
  sheet.getRange(fila, 1, 1, 7).merge()
    .setValue('👤  INFORMACIÓN DEL ESTUDIANTE')
    .setBackground('#1565c0').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(fila, 28); fila++;

  const perfilDatos = [
    ['Nombre',   est.nombre,                  'ID / Código', est.id],
    ['Carrera',  est.programa || '—',         'Grado',       est.grado   || '—'],
    ['Sección',  est.seccion  || '—',         'Sexo',        est.sexo === 'M' ? 'Masculino' : est.sexo === 'F' ? 'Femenino' : '—'],
    ['Estado',   est.estado   || '—',         'Email',       est.email   || '—']
  ];
  perfilDatos.forEach(([l1, v1, l2, v2]) => {
    sheet.getRange(fila, 1).setValue(l1).setFontWeight('bold').setBackground('#e8f0fe').setFontColor('#1a73e8').setVerticalAlignment('middle');
    sheet.getRange(fila, 2, 1, 2).merge().setValue(v1).setBackground('#ffffff').setVerticalAlignment('middle');
    sheet.getRange(fila, 4).setValue(l2).setFontWeight('bold').setBackground('#e8f0fe').setFontColor('#1a73e8').setVerticalAlignment('middle');
    sheet.getRange(fila, 5, 1, 3).merge().setValue(v2).setBackground('#ffffff').setVerticalAlignment('middle');
    sheet.setRowHeight(fila, 22); fila++;
  });
  fila++;

  // ── RESUMEN ──
  sheet.getRange(fila, 1, 1, 7).merge()
    .setValue('📊  RESUMEN GENERAL DE ASISTENCIA')
    .setBackground('#1565c0').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(fila, 28); fila++;

  sheet.getRange(fila, 1, 1, 6).setValues([['Total Clases','Presentes','Ausentes','Puntuales','Tardíos','% Asistencia']])
    .setBackground('#bbdefb').setFontColor('#0d47a1').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(fila, 22); fila++;

  sheet.getRange(fila, 1).setValue(total)    .setHorizontalAlignment('center').setFontSize(16).setFontWeight('bold').setBackground('#ffffff');
  sheet.getRange(fila, 2).setValue(presentes).setHorizontalAlignment('center').setFontSize(16).setFontWeight('bold').setFontColor('#2e7d32').setBackground('#d9ead3');
  sheet.getRange(fila, 3).setValue(ausentes) .setHorizontalAlignment('center').setFontSize(16).setFontWeight('bold').setFontColor('#c62828').setBackground('#fce8e6');
  sheet.getRange(fila, 4).setValue(puntuales).setHorizontalAlignment('center').setFontSize(16).setFontWeight('bold').setBackground('#ffffff');
  sheet.getRange(fila, 5).setValue(tarde)    .setHorizontalAlignment('center').setFontSize(16).setFontWeight('bold').setFontColor('#e65100').setBackground('#fff3e0');
  sheet.getRange(fila, 6).setValue(pct + '%').setHorizontalAlignment('center').setFontSize(18).setFontWeight('bold').setFontColor(pctColor).setBackground('#ffffff');
  sheet.setRowHeight(fila, 36); fila += 2;

  // ── TENDENCIA MENSUAL ──
  if (meses.length > 0) {
    sheet.getRange(fila, 1, 1, 7).merge()
      .setValue('📅  TENDENCIA MENSUAL')
      .setBackground('#1565c0').setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(fila, 28); fila++;

    sheet.getRange(fila, 1, 1, 7).setValues([['Mes','Total','Presentes','Ausentes','Puntuales','Tardíos','% Asist.']])
      .setBackground('#bbdefb').setFontColor('#0d47a1').setFontWeight('bold')
      .setHorizontalAlignment('center');
    sheet.setRowHeight(fila, 22); fila++;

    const MESES_ES = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
    meses.forEach(m => {
      const p   = m.total > 0 ? Math.round(m.presentes / m.total * 100) : 0;
      const pc  = p >= 80 ? '#2e7d32' : p >= 60 ? '#e65100' : '#c62828';
      const bg  = p >= 80 ? '#d9ead3' : p >= 60 ? '#fff3e0' : '#fce8e6';
      const [y, mo] = m.mes.split('-');
      const label   = MESES_ES[parseInt(mo) - 1] + ' ' + y;
      sheet.getRange(fila, 1, 1, 7).setBackground(bg).setHorizontalAlignment('center');
      sheet.getRange(fila, 1).setValue(label)    .setFontWeight('bold').setHorizontalAlignment('left');
      sheet.getRange(fila, 2).setValue(m.total);
      sheet.getRange(fila, 3).setValue(m.presentes).setFontColor('#2e7d32');
      sheet.getRange(fila, 4).setValue(m.ausentes) .setFontColor('#c62828');
      sheet.getRange(fila, 5).setValue(m.puntuales);
      sheet.getRange(fila, 6).setValue(m.tarde)    .setFontColor('#e65100');
      sheet.getRange(fila, 7).setValue(p + '%').setFontWeight('bold').setFontColor(pc);
      sheet.setRowHeight(fila, 20); fila++;
    });
    fila++;
  }

  // ── HISTORIAL COMPLETO ──
  sheet.getRange(fila, 1, 1, 7).merge()
    .setValue('🗒  HISTORIAL COMPLETO  (' + registros.length + ' registros)')
    .setBackground('#1565c0').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('left')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(fila, 28); fila++;

  sheet.getRange(fila, 1, 1, 6).setValues([['Fecha','Hora','Materia','Grado','Sección','Estado']])
    .setBackground('#bbdefb').setFontColor('#0d47a1').setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(fila, 22); fila++;

  if (registros.length === 0) {
    sheet.getRange(fila, 1, 1, 6).merge()
      .setValue('Sin registros de asistencia')
      .setFontColor('#aaaaaa').setHorizontalAlignment('center').setFontStyle('italic');
    sheet.setRowHeight(fila, 24);
  } else {
    // Batch write para rendimiento
    const filasDatos = registros.map(r => {
      const icon = r.obs === 'Puntual' ? '✅' : r.obs === 'Tarde' ? '⚠️' : '❌';
      return [r.fecha, r.hora, r.materia || '—', r.grado || '—', r.seccion || '—', icon + ' ' + r.obs];
    });
    sheet.getRange(fila, 1, filasDatos.length, 6).setValues(filasDatos).setHorizontalAlignment('center');
    sheet.getRange(fila, 1, filasDatos.length, 1).setHorizontalAlignment('left');

    registros.forEach((r, idx) => {
      const bg = r.obs === 'Puntual' ? '#d9ead3' : r.obs === 'Tarde' ? '#fff3e0' : '#fce8e6';
      const fc = r.obs === 'Puntual' ? '#2e7d32' : r.obs === 'Tarde' ? '#e65100' : '#c62828';
      sheet.getRange(fila + idx, 1, 1, 6).setBackground(bg);
      sheet.getRange(fila + idx, 6).setFontWeight('bold').setFontColor(fc);
      sheet.setRowHeight(fila + idx, 20);
    });
  }
}

function _limpiarAreaEstadisticas(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow >= 3) {
    sheet.getRange(3, 1, lastRow - 2, 7).clear();
  }
}

// ------------------------------------------------------------
// MENÚ EN GOOGLE SHEETS
// ------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎓 Asistencia QR')
    .addItem('▶ Abrir Escáner QR',          'abrirEscaner')
    .addItem('📋 Lista de Asistencia',       'abrirLista')
    .addItem('📊 Ver Dashboard',             'abrirDashboard')
    .addItem('👥 Gestionar Estudiantes',     'abrirEstudiantes')
    .addItem('📊 Estadísticas por Estudiante', 'crearHojaEstadisticas')
    .addSeparator()
    .addItem('⚙ Inicializar Spreadsheet',   'inicializarSpreadsheet')
    .addItem('🔗 Obtener URL de la App',     'mostrarUrl')
    .addToUi();
}

function _abrirPagina(sufijo, titulo) {
  const url  = ScriptApp.getService().getUrl() + sufijo;
  const html = HtmlService.createHtmlOutput(
    `<script>window.open('${url}','_blank');google.script.host.close();</script>`
  ).setWidth(10).setHeight(10);
  SpreadsheetApp.getUi().showModalDialog(html, titulo);
}

function abrirEscaner()     { _abrirPagina('',                  'Abriendo escáner...'); }
function abrirLista()       { _abrirPagina('?page=lista',       'Abriendo lista...'); }
function abrirDashboard()   { _abrirPagina('?page=dashboard',   'Abriendo dashboard...'); }
function abrirEstudiantes() { _abrirPagina('?page=estudiantes', 'Abriendo estudiantes...'); }

function mostrarUrl() {
  SpreadsheetApp.getUi().alert('URL de la Web App:\n\n' + ScriptApp.getService().getUrl());
}
