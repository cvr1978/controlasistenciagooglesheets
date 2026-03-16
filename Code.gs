// ============================================================
// SISTEMA DE CONTROL DE ASISTENCIA CON QR - Google Apps Script
// ============================================================

// Nombres de las hojas del spreadsheet
const SHEET_ESTUDIANTES = 'Estudiantes';
const SHEET_ASISTENCIA  = 'Asistencia';
const SHEET_CONFIG      = 'Configuracion';

// Columnas hoja Estudiantes: ID | Nombre | Programa | Email | Estado
const COL_EST_ID       = 1;
const COL_EST_NOMBRE   = 2;
const COL_EST_PROGRAMA = 3;
const COL_EST_EMAIL    = 4;
const COL_EST_ESTADO   = 5;  // Activo / Inactivo

// Columnas hoja Asistencia: Fecha | Hora | ID_Estudiante | Nombre | Programa | Materia | Observacion
const COL_ASIST_FECHA      = 1;
const COL_ASIST_HORA       = 2;
const COL_ASIST_ID         = 3;
const COL_ASIST_NOMBRE     = 4;
const COL_ASIST_PROGRAMA   = 5;
const COL_ASIST_MATERIA    = 6;
const COL_ASIST_OBS        = 7;

// ------------------------------------------------------------
// PUNTO DE ENTRADA WEB APP
// ------------------------------------------------------------

function doGet(e) {
  const page = e.parameter.page || 'scanner';

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

  // Página principal: escáner QR
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Control de Asistencia QR')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Helper para incluir archivos HTML parciales
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Retorna la URL de la web app
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
      ['Tolerancia',  '15'],   // minutos de tolerancia para llegar tarde
      ['ColorHeader', '#1a73e8']
    ]);
    sheetConfig.getRange('A1:B1').setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
    sheetConfig.setFrozenRows(1);
  }

  // ---- Hoja Estudiantes ----
  let sheetEst = ss.getSheetByName(SHEET_ESTUDIANTES);
  if (!sheetEst) {
    sheetEst = ss.insertSheet(SHEET_ESTUDIANTES);
    const headers = ['ID / Código', 'Nombre Completo', 'Programa / Curso', 'Email', 'Estado'];
    sheetEst.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheetEst.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1a73e8')
      .setFontColor('#ffffff');
    sheetEst.setFrozenRows(1);
    sheetEst.setColumnWidth(1, 120);
    sheetEst.setColumnWidth(2, 220);
    sheetEst.setColumnWidth(3, 200);
    sheetEst.setColumnWidth(4, 220);
    sheetEst.setColumnWidth(5, 100);
  }

  // ---- Hoja Asistencia ----
  let sheetAsist = ss.getSheetByName(SHEET_ASISTENCIA);
  if (!sheetAsist) {
    sheetAsist = ss.insertSheet(SHEET_ASISTENCIA);
    const headers = ['Fecha', 'Hora', 'ID / Código', 'Nombre Completo', 'Programa / Curso', 'Materia', 'Observación'];
    sheetAsist.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheetAsist.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1a73e8')
      .setFontColor('#ffffff');
    sheetAsist.setFrozenRows(1);
    sheetAsist.setColumnWidth(1, 110);
    sheetAsist.setColumnWidth(2, 90);
    sheetAsist.setColumnWidth(3, 120);
    sheetAsist.setColumnWidth(4, 220);
    sheetAsist.setColumnWidth(5, 200);
    sheetAsist.setColumnWidth(6, 160);
    sheetAsist.setColumnWidth(7, 160);
  }

  SpreadsheetApp.getUi().alert('✅ Spreadsheet inicializado correctamente.');
}

// ------------------------------------------------------------
// REGISTRO DE ASISTENCIA (llamado desde la Web App)
// ------------------------------------------------------------

/**
 * Registra la asistencia de un estudiante a partir del contenido del QR.
 * El QR del carnet debe contener el ID/código del estudiante.
 * Retorna un objeto { ok, mensaje, estudiante }
 */
function registrarAsistencia(codigoQr) {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const sheetE  = ss.getSheetByName(SHEET_ESTUDIANTES);
    const sheetA  = ss.getSheetByName(SHEET_ASISTENCIA);
    const config  = obtenerConfiguracion();

    if (!sheetE || !sheetA) {
      return { ok: false, mensaje: 'El spreadsheet no está inicializado. Ejecuta inicializarSpreadsheet().' };
    }

    // Limpiar código QR (trim y uppercase para comparación robusta)
    const codigoLimpio = codigoQr.toString().trim();

    // Buscar estudiante
    const datosEst = sheetE.getDataRange().getValues();
    let estudianteEncontrado = null;

    for (let i = 1; i < datosEst.length; i++) {
      const idHoja = datosEst[i][COL_EST_ID - 1].toString().trim();
      if (idHoja === codigoLimpio) {
        estudianteEncontrado = {
          id:       datosEst[i][COL_EST_ID - 1],
          nombre:   datosEst[i][COL_EST_NOMBRE - 1],
          programa: datosEst[i][COL_EST_PROGRAMA - 1],
          email:    datosEst[i][COL_EST_EMAIL - 1],
          estado:   datosEst[i][COL_EST_ESTADO - 1]
        };
        break;
      }
    }

    if (!estudianteEncontrado) {
      return {
        ok: false,
        mensaje: `Código "${codigoLimpio}" no encontrado. Verifica que el estudiante esté registrado.`
      };
    }

    if (estudianteEncontrado.estado.toString().toLowerCase() === 'inactivo') {
      return {
        ok: false,
        mensaje: `El estudiante ${estudianteEncontrado.nombre} está marcado como Inactivo.`
      };
    }

    // Verificar si ya registró asistencia hoy
    const ahora    = new Date();
    const hoyStr   = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const datosA   = sheetA.getDataRange().getValues();
    let yaRegistro = false;

    for (let i = 1; i < datosA.length; i++) {
      const fechaFila = datosA[i][COL_ASIST_FECHA - 1];
      const idFila    = datosA[i][COL_ASIST_ID - 1].toString().trim();
      if (idFila === codigoLimpio && fechaFila) {
        const fechaFilaStr = Utilities.formatDate(new Date(fechaFila), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        if (fechaFilaStr === hoyStr) {
          yaRegistro = true;
          break;
        }
      }
    }

    if (yaRegistro) {
      return {
        ok: false,
        mensaje: `${estudianteEncontrado.nombre} ya tiene asistencia registrada hoy.`,
        estudiante: estudianteEncontrado
      };
    }

    // Determinar observación (puntual / tarde)
    const horaActual    = ahora.getHours() * 60 + ahora.getMinutes();
    const toleranciaMin = parseInt(config['Tolerancia'] || 15);
    // Clase empieza a las 7:00 AM por defecto — ajustable en Config
    const horaClase     = parseInt(config['HoraClaseH'] || 7) * 60 + parseInt(config['HoraClaseM'] || 0);
    const observacion   = horaActual <= horaClase + toleranciaMin ? 'Puntual' : 'Tarde';

    // Registrar fila
    const fecha     = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const hora      = Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'HH:mm:ss');
    const materia   = config['Materia'] || 'General';

    sheetA.appendRow([
      fecha,
      hora,
      estudianteEncontrado.id,
      estudianteEncontrado.nombre,
      estudianteEncontrado.programa,
      materia,
      observacion
    ]);

    // Colorear fila según observación
    const ultimaFila = sheetA.getLastRow();
    const color = observacion === 'Puntual' ? '#d9ead3' : '#fff2cc';
    sheetA.getRange(ultimaFila, 1, 1, 7).setBackground(color);

    return {
      ok: true,
      mensaje: `✅ Asistencia registrada para ${estudianteEncontrado.nombre} — ${observacion}`,
      estudiante: estudianteEncontrado,
      hora: hora,
      observacion: observacion
    };

  } catch (err) {
    return { ok: false, mensaje: 'Error interno: ' + err.message };
  }
}

// ------------------------------------------------------------
// GESTIÓN DE ESTUDIANTES
// ------------------------------------------------------------

function obtenerEstudiantes() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ESTUDIANTES);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const datos = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  return datos
    .filter(r => r[0] !== '')
    .map(r => ({
      id:       r[COL_EST_ID - 1],
      nombre:   r[COL_EST_NOMBRE - 1],
      programa: r[COL_EST_PROGRAMA - 1],
      email:    r[COL_EST_EMAIL - 1],
      estado:   r[COL_EST_ESTADO - 1]
    }));
}

function agregarEstudiante(datos) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ESTUDIANTES);

    // Validar que el ID no exista
    const existentes = sheet.getDataRange().getValues();
    for (let i = 1; i < existentes.length; i++) {
      if (existentes[i][0].toString().trim() === datos.id.toString().trim()) {
        return { ok: false, mensaje: 'Ya existe un estudiante con ese ID/Código.' };
      }
    }

    sheet.appendRow([
      datos.id,
      datos.nombre,
      datos.programa,
      datos.email || '',
      datos.estado || 'Activo'
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
        sheet.getRange(i + 1, 1, 1, 5).setValues([[
          datos.id,
          datos.nombre,
          datos.programa,
          datos.email || '',
          datos.estado || 'Activo'
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
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheetA = ss.getSheetByName(SHEET_ASISTENCIA);
  const sheetE = ss.getSheetByName(SHEET_ESTUDIANTES);

  const hoyStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const totalEstudiantes = sheetE ? Math.max(0, sheetE.getLastRow() - 1) : 0;

  if (!sheetA || sheetA.getLastRow() < 2) {
    return { fecha: hoyStr, presentes: 0, puntuales: 0, tarde: 0, total: totalEstudiantes, registros: [] };
  }

  const datos = sheetA.getDataRange().getValues();
  const registrosHoy = [];

  for (let i = 1; i < datos.length; i++) {
    const fechaFila = datos[i][COL_ASIST_FECHA - 1];
    if (!fechaFila) continue;
    const fechaFilaStr = Utilities.formatDate(new Date(fechaFila), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (fechaFilaStr === hoyStr) {
      registrosHoy.push({
        hora:      datos[i][COL_ASIST_HORA - 1],
        id:        datos[i][COL_ASIST_ID - 1],
        nombre:    datos[i][COL_ASIST_NOMBRE - 1],
        programa:  datos[i][COL_ASIST_PROGRAMA - 1],
        materia:   datos[i][COL_ASIST_MATERIA - 1],
        obs:       datos[i][COL_ASIST_OBS - 1]
      });
    }
  }

  const puntuales = registrosHoy.filter(r => r.obs === 'Puntual').length;
  const tarde     = registrosHoy.filter(r => r.obs === 'Tarde').length;

  return {
    fecha:     hoyStr,
    presentes: registrosHoy.length,
    puntuales: puntuales,
    tarde:     tarde,
    total:     totalEstudiantes,
    registros: registrosHoy
  };
}

function obtenerAsistenciaPorFecha(fechaInicio, fechaFin) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const sheetA = ss.getSheetByName(SHEET_ASISTENCIA);

  if (!sheetA || sheetA.getLastRow() < 2) return [];

  const datos     = sheetA.getDataRange().getValues();
  const tz        = Session.getScriptTimeZone();
  const resultado = [];

  for (let i = 1; i < datos.length; i++) {
    const fechaFila = datos[i][COL_ASIST_FECHA - 1];
    if (!fechaFila) continue;
    const fechaFilaStr = Utilities.formatDate(new Date(fechaFila), tz, 'yyyy-MM-dd');
    if (fechaFilaStr >= fechaInicio && fechaFilaStr <= fechaFin) {
      resultado.push({
        fecha:    fechaFilaStr,
        hora:     datos[i][COL_ASIST_HORA - 1],
        id:       datos[i][COL_ASIST_ID - 1],
        nombre:   datos[i][COL_ASIST_NOMBRE - 1],
        programa: datos[i][COL_ASIST_PROGRAMA - 1],
        materia:  datos[i][COL_ASIST_MATERIA - 1],
        obs:      datos[i][COL_ASIST_OBS - 1]
      });
    }
  }

  return resultado;
}

// ------------------------------------------------------------
// CONFIGURACIÓN
// ------------------------------------------------------------

function obtenerConfiguracion() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet || sheet.getLastRow() < 2) return {};

  const datos  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const config = {};
  datos.forEach(r => {
    if (r[0]) config[r[0]] = r[1];
  });
  return config;
}

function guardarConfiguracion(datos) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_CONFIG);

    // Leer config actual y actualizar/agregar claves
    const filas = sheet.getDataRange().getValues();
    const claves = {};
    for (let i = 1; i < filas.length; i++) {
      claves[filas[i][0]] = i + 1; // fila en el sheet (1-based)
    }

    Object.keys(datos).forEach(clave => {
      if (claves[clave]) {
        sheet.getRange(claves[clave], 2).setValue(datos[clave]);
      } else {
        sheet.appendRow([clave, datos[clave]]);
      }
    });

    return { ok: true, mensaje: 'Configuración guardada.' };
  } catch (err) {
    return { ok: false, mensaje: 'Error: ' + err.message };
  }
}

// ------------------------------------------------------------
// MENÚ EN GOOGLE SHEETS
// ------------------------------------------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎓 Asistencia QR')
    .addItem('▶ Abrir Escáner QR',         'abrirEscaner')
    .addItem('📊 Ver Dashboard',            'abrirDashboard')
    .addItem('👥 Gestionar Estudiantes',    'abrirEstudiantes')
    .addSeparator()
    .addItem('⚙ Inicializar Spreadsheet',  'inicializarSpreadsheet')
    .addItem('🔗 Obtener URL de la App',    'mostrarUrl')
    .addToUi();
}

function abrirEscaner() {
  const url  = ScriptApp.getService().getUrl();
  const html = HtmlService.createHtmlOutput(
    `<script>window.open('${url}','_blank');google.script.host.close();</script>`
  ).setWidth(10).setHeight(10);
  SpreadsheetApp.getUi().showModalDialog(html, 'Abriendo...');
}

function abrirDashboard() {
  const url  = ScriptApp.getService().getUrl() + '?page=dashboard';
  const html = HtmlService.createHtmlOutput(
    `<script>window.open('${url}','_blank');google.script.host.close();</script>`
  ).setWidth(10).setHeight(10);
  SpreadsheetApp.getUi().showModalDialog(html, 'Abriendo...');
}

function abrirEstudiantes() {
  const url  = ScriptApp.getService().getUrl() + '?page=estudiantes';
  const html = HtmlService.createHtmlOutput(
    `<script>window.open('${url}','_blank');google.script.host.close();</script>`
  ).setWidth(10).setHeight(10);
  SpreadsheetApp.getUi().showModalDialog(html, 'Abriendo...');
}

function mostrarUrl() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert('URL de la Web App:\n\n' + url);
}
