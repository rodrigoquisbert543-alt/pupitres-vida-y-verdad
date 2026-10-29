/**
 * ALERTAS Y NOTIFICACIONES
 * - Correo inmediato cuando se reporta un pupitre dañado / en mantenimiento / carpintería
 * - Resumen diario (trigger 7:00) con todo lo pendiente
 * - Alertas en el panel (campana)
 */

function dias_(desde, hasta) {
  return desde instanceof Date ? Math.floor((hasta - desde) / 86400000) : 0;
}

function calcularAlertas_(lista) {
  const hoy = new Date();
  const maxMant = Number(cfg_('Dias_Max_Mantenimiento'));
  const maxCarp = Number(cfg_('Dias_Max_Carpinteria'));
  const aviso = Number(cfg_('Dias_Aviso_Mantenimiento'));
  const out = [];
  lista.forEach(p => {
    const d = dias_(p.desde, hoy);
    const base = { id: p.id, curso: p.curso, dias: d, foto: p.foto || '' };
    if (p.estado === 'Dañado') out.push(Object.assign(base, { tipo: 'DANADO', sev: 1, msg: 'Dañado: necesita reparación' }));
    else if (p.estado === 'En Mantenimiento' && d > maxMant) out.push(Object.assign(base, { tipo: 'MANT_ATRASADO', sev: 2, msg: 'En mantenimiento hace ' + d + ' días' }));
    else if (p.estado === 'En Carpinteria' && d > maxCarp) out.push(Object.assign(base, { tipo: 'CARP_ATRASADO', sev: 2, msg: 'En carpintería hace ' + d + ' días' }));
    else if (p.estado === 'Bueno' && p.proxMant instanceof Date) {
      const falta = Math.ceil((p.proxMant - hoy) / 86400000);
      if (falta < 0) out.push(Object.assign(base, { tipo: 'MANT_VENCIDO', sev: 2, msg: 'Mantenimiento vencido hace ' + (-falta) + ' días' }));
      else if (falta <= aviso) out.push(Object.assign(base, { tipo: 'MANT_PROXIMO', sev: 3, msg: 'Mantenimiento programado en ' + falta + ' días' }));
    }
  });
  return out.sort((a, b) => a.sev - b.sev || b.dias - a.dias);
}

function destinatarios_() {
  const c = String(cfg_('Emails_Notificacion', '')).split(/[,;\s]+/).filter(x => /@/.test(x));
  if (c.length) return c.join(',');
  try { return Session.getEffectiveUser().getEmail(); } catch (e) { return ''; }
}

function esc_(t) {
  return String(t == null ? '' : t).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function notificarReporte_(p, estado, nota, s, fotoUrl) {
  const to = destinatarios_();
  if (!to) return;
  const foto = fotoUrl || p.foto;
  const html = '<div style="font-family:Arial,sans-serif;max-width:520px">' +
    '<h2 style="color:#0b3a8f">Pupitre ' + esc_(p.id) + ' → ' + esc_(estado) + '</h2>' +
    '<p><b>Curso:</b> ' + esc_(p.curso) + ' &nbsp; <b>Ubicación:</b> ' + esc_(p.ubicacion) + '<br>' +
    '<b>Reportado por:</b> ' + esc_(s.nombre || s.usuario) + ' (' + esc_(s.rol) + ')<br>' +
    '<b>Estado anterior:</b> ' + esc_(p.estado) + '</p>' +
    (nota ? '<p><b>Detalle:</b> ' + esc_(nota) + '</p>' : '') +
    (foto ? '<p><a href="' + esc_(foto) + '"><img src="' + esc_(foto) + '" style="max-width:100%;border-radius:8px"></a></p>' : '') +
    '<p><a href="' + esc_(appUrl_() + '?p=' + p.id) + '">Abrir en el sistema</a></p></div>';
  MailApp.sendEmail({ to: to, subject: '🪑 ' + p.id + ': ' + estado + ' · Vida y Verdad', htmlBody: html });
}

/** Trigger diario: envía un resumen sólo si hay pendientes importantes. */
function resumenDiario() {
  const alertas = calcularAlertas_(leerTodos_()).filter(a => a.sev <= 3);
  const to = destinatarios_();
  if (!alertas.length || !to) return;
  const filas = alertas.map(a => '<tr><td style="padding:4px 10px"><b>' + esc_(a.id) + '</b></td><td style="padding:4px 10px">' +
    esc_(a.curso) + '</td><td style="padding:4px 10px">' + esc_(a.msg) + '</td></tr>').join('');
  MailApp.sendEmail({
    to: to,
    subject: '🪑 Resumen de pupitres: ' + alertas.length + ' pendiente(s)',
    htmlBody: '<div style="font-family:Arial,sans-serif"><h2 style="color:#0b3a8f">Pupitres que requieren atención</h2>' +
      '<table style="border-collapse:collapse;font-size:14px">' + filas + '</table>' +
      '<p><a href="' + esc_(appUrl_()) + '">Abrir el sistema</a></p></div>'
  });
}

/** Ejecutar una vez (inicializarSistema ya lo hace). */
function instalarTriggers() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'resumenDiario').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('resumenDiario').timeBased().everyDays(1).atHour(7).create();
}
