/****************************************************
 * SISTEMA CENTRAL DE TICKETS DE DISEÑO
 * Backend: Google Sheets + Google Apps Script
 *
 * CONFIGURACIÓN OBLIGATORIA ANTES DE USAR:
 * 1. Reemplaza SHEET_ID con el ID de tu Google Sheet
 *    (lo que va en la URL entre /d/ y /edit)
 * 2. Reemplaza DRIVE_FOLDER_ID con el ID de una carpeta
 *    de Google Drive donde se guardarán las capturas
 * 3. Implementar > Nueva implementación > Aplicación web
 *    - Ejecutar como: Yo
 *    - Quién tiene acceso: Cualquier usuario
 * 4. Copia la URL /exec resultante al CONFIG del index.html
 ****************************************************/

const SHEET_ID = 'PON_AQUI_EL_ID_DE_TU_GOOGLE_SHEET';
const DRIVE_FOLDER_ID = 'PON_AQUI_EL_ID_DE_TU_CARPETA_DE_DRIVE';

// Estas claves definen el orden de las columnas en cada hoja mensual.
// Si agregas un campo nuevo en el futuro, agrégalo aquí también (al final).
const HEADERS_KEYS = ['folio','fechaHora','marca','sucursal','solicitante','telefono',
  'esAjuste','ticketPrevio','ronda','complejidad','descripcion','referencia','urgente',
  'fechaEntregaEstimada','motivoAjusteFecha','estado','disenador','capturaURL',
  'fechaCaptura','comentarioSolicitante','fechaAprobacion','creadorId'];

const HEADERS_LABELS = ['Folio','Fecha y hora','Marca','Sucursal','Solicitante','Teléfono',
  'Es ajuste','Ticket previo','Ronda','Complejidad','Descripción','Referencia','Urgente',
  'Fecha entrega estimada','Motivo ajuste de fecha','Estado','Diseñador','URL captura',
  'Fecha captura','Comentario solicitante','Fecha aprobación','ID creador'];

// === PUNTOS DE ENTRADA ===

function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const hojas = obtenerHojasRecientes(ss, 6);
    let tickets = [];
    hojas.forEach(sheet => {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (!data[i][0]) continue;
        tickets.push(filaAObjeto(data[i]));
      }
    });
    return respuesta({ ok: true, tickets: tickets });
  } catch (err) {
    return respuesta({ ok: false, error: err.message });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const body = JSON.parse(e.postData.contents);
    let resultado;
    if (body.accion === 'guardar') resultado = guardarTicket(body);
    else if (body.accion === 'actualizar') resultado = actualizarTicket(body);
    else if (body.accion === 'subirCaptura') resultado = subirCaptura(body);
    else resultado = { ok: false, error: 'Acción no reconocida: ' + body.accion };
    return respuesta(resultado);
  } catch (err) {
    return respuesta({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

// === ACCIONES ===

function guardarTicket(body) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const folio = siguienteFolio();
  const fechaHora = new Date().toISOString();
  const sheet = hojaDelMes(ss, fechaHora);

  const registro = {
    folio: folio,
    fechaHora: fechaHora,
    marca: body.marca || '',
    sucursal: body.sucursal || '',
    solicitante: body.solicitante || '',
    telefono: body.telefono || '',
    esAjuste: body.esAjuste ? 'SI' : 'NO',
    ticketPrevio: body.ticketPrevio || '',
    ronda: body.ronda || 0,
    complejidad: body.complejidad || '',
    descripcion: body.descripcion || '',
    referencia: body.referencia || '',
    urgente: body.urgente ? 'SI' : 'NO',
    fechaEntregaEstimada: body.fechaEntregaEstimada || '',
    motivoAjusteFecha: body.motivoAjusteFecha || '',
    estado: 'En Fila de Espera',
    disenador: 'Sin Asignar',
    capturaURL: '',
    fechaCaptura: '',
    comentarioSolicitante: '',
    fechaAprobacion: '',
    creadorId: body.creadorId || ''
  };

  const fila = HEADERS_KEYS.map(k => registro[k]);
  sheet.appendRow(fila);
  return { ok: true, ticket: registro };
}

function actualizarTicket(body) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const ubicacion = buscarFilaPorFolio(ss, body.folio);
  if (!ubicacion) return { ok: false, error: 'Folio no encontrado: ' + body.folio };

  const campos = body.campos || {};
  Object.keys(campos).forEach(clave => {
    const col = HEADERS_KEYS.indexOf(clave);
    if (col === -1) return;
    ubicacion.sheet.getRange(ubicacion.fila, col + 1).setValue(campos[clave]);
  });
  return { ok: true };
}

function subirCaptura(body) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const ubicacion = buscarFilaPorFolio(ss, body.folio);
  if (!ubicacion) return { ok: false, error: 'Folio no encontrado: ' + body.folio };

  const carpeta = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const bytes = Utilities.base64Decode(body.imagenBase64);
  const blob = Utilities.newBlob(bytes, body.mimeType || 'image/jpeg', body.folio + '.jpg');
  const archivo = carpeta.createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = 'https://drive.google.com/uc?export=view&id=' + archivo.getId();

  const colCaptura = HEADERS_KEYS.indexOf('capturaURL') + 1;
  const colFechaCaptura = HEADERS_KEYS.indexOf('fechaCaptura') + 1;
  const colEstado = HEADERS_KEYS.indexOf('estado') + 1;
  ubicacion.sheet.getRange(ubicacion.fila, colCaptura).setValue(url);
  ubicacion.sheet.getRange(ubicacion.fila, colFechaCaptura).setValue(new Date().toISOString());
  ubicacion.sheet.getRange(ubicacion.fila, colEstado).setValue('En Revisión');

  return { ok: true, url: url };
}

// === UTILIDADES ===

function siguienteFolio() {
  const props = PropertiesService.getScriptProperties();
  const anio = new Date().getFullYear();
  const clave = 'folio_' + anio;
  const n = parseInt(props.getProperty(clave) || '0', 10) + 1;
  props.setProperty(clave, String(n));
  return 'TK-' + anio + '-' + String(n).padStart(4, '0');
}

function hojaDelMes(ss, fechaISO) {
  const nombre = fechaISO.substring(0, 7); // yyyy-MM
  let sheet = ss.getSheetByName(nombre);
  if (!sheet) {
    sheet = ss.insertSheet(nombre);
    sheet.appendRow(HEADERS_LABELS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS_LABELS.length).setFontWeight('bold').setBackground('#fed303');
  }
  return sheet;
}

function obtenerHojasRecientes(ss, n) {
  const todas = ss.getSheets().filter(s => /^\d{4}-\d{2}$/.test(s.getName()));
  todas.sort((a, b) => b.getName().localeCompare(a.getName()));
  return todas.slice(0, n);
}

function buscarFilaPorFolio(ss, folio) {
  const hojas = obtenerHojasRecientes(ss, 6);
  for (const sheet of hojas) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === folio) return { sheet: sheet, fila: i + 1 };
    }
  }
  return null;
}

function filaAObjeto(fila) {
  const obj = {};
  HEADERS_KEYS.forEach((k, i) => { obj[k] = fila[i]; });
  return obj;
}

function respuesta(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
