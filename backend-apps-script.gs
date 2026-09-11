/**
 * EL EJE — Postulación Secretarías 2026
 * Backend: recibe las postulaciones del formulario web y las guarda en la hoja "Postulaciones".
 * Implementar como App web → Ejecutar como: Yo → Quién tiene acceso: Cualquier usuario.
 */
const SHEET_NAME = 'Postulaciones';

// [clave que envía el formulario, encabezado de la columna]
const COLUMNS = [
  ['recibido_el',      'Recibido'],
  ['nombre',           'Nombre y apellido'],
  ['dni',              'DNI'],
  ['edad',             'Edad'],
  ['telefono',         'Teléfono'],
  ['email',            'Email'],
  ['distrito',         'Distrito'],
  ['barrio',           'Barrio / Localidad'],
  ['ocupacion',        'Situación actual'],
  ['instagram',        'Instagram'],
  ['q_porque',         '¿Por qué El Eje?'],
  ['q_instrucciones',  'Tarea sin instrucciones'],
  ['q_equipo',         'Compañero que no cumplió'],
  ['q_situacion',      'Situación difícil resuelta'],
  ['q_desacuerdo',     'Desacuerdo con coordinador'],
  ['q_critica',        'Crítica recibida'],
  ['q_autonomia',      'Autonomía (1-10)'],
  ['q_lealtad',        'Lealtad vs. objetivos'],
  ['area',             'Secretaría'],
  ['subarea',          'Subárea'],
  ['area_porque',      '¿Por qué esta área?'],
  ['area_experiencia', 'Experiencia previa'],
  ['area_idea',        'Idea concreta'],
  ['presencial',       'Presencialidad BA'],
  ['extra',            'Comentario adicional'],
  ['acepto',           'Acepta contacto'],
  ['enviado_el',       'Hora del postulante']
];

// Ejecutar una vez desde el editor: crea la pestaña con encabezados y pide los permisos.
function setup() {
  getSheet_();
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!data.nombre || !data.dni) return json_({ ok: false, error: 'Datos incompletos' });

    const sheet = getSheet_();
    data.recibido_el = new Date();
    const row = COLUMNS.map(([key]) => clean_(data[key]));
    row[0] = data.recibido_el; // la fecha va como fecha real, no como texto
    sheet.appendRow(row);
    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

// Permite comprobar que la App web está activa abriendo la URL en el navegador.
function doGet() {
  return json_({ ok: true, servicio: 'El Eje — Postulaciones activo' });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS.map(c => c[1]));
    sheet.getRange(1, 1, 1, COLUMNS.length)
      .setFontWeight('bold').setBackground('#1A1A2E').setFontColor('#C8A951');
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('dd/mm/yyyy hh:mm');
    sheet.getRange('C:E').setNumberFormat('@'); // DNI, edad y teléfono como texto
  }
  return sheet;
}

// Fuerza texto plano: evita fórmulas inyectadas y que '+54 11...' se lea como fórmula (#ERROR!).
function clean_(v) {
  if (v === undefined || v === null) return '';
  let s = String(v).trim().slice(0, 5000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
