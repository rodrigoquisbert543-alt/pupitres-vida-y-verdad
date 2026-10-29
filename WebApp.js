/**
 * WEB APP - Punto de entrada
 * El QR de cada pupitre es la URL de la app con ?p=ID, así la cámara nativa del
 * celular abre el pupitre directamente (tras iniciar sesión si hace falta).
 */

function doGet(e) {
  try {
    return HtmlService.createHtmlOutputFromFile('App')
      .setTitle('Pupitres · Vida y Verdad')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:30px;color:#b00">Error: ' +
      String(error.message).replace(/</g, '&lt;') + '</p>');
  }
}

const API_PERMITIDAS_ = ['login', 'logout', 'sesionInfo', 'cambiarContrasena', 'listarUsuarios', 'guardarUsuario',
  'cargarInicio', 'obtenerPupitre', 'moverPupitre', 'cambiarEstado', 'crearPupitres', 'obtenerConfig', 'guardarConfig'];

/** API JSON para la versión alojada fuera de Apps Script (cámara en vivo). */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    if (API_PERMITIDAS_.indexOf(req.fn) < 0) throw new Error('Operación no permitida');
    out = globalThis[req.fn].apply(null, req.args || []);
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out === undefined ? { ok: true } : out))
    .setMimeType(ContentService.MimeType.JSON);
}

function iso_(d) {
  return (d instanceof Date && !isNaN(d)) ? d.toISOString() : '';
}

function appUrl_() {
  return String(cfg_('URL_App', '') || ScriptApp.getService().getUrl() || '');
}
