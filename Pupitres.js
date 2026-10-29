/**
 * LÓGICA DE PUPITRES - inventario, estados, fotos, carpintería, lotes
 */

const ESTADOS = ['Bueno', 'Dañado', 'En Mantenimiento', 'En Carpinteria', 'Dado de Baja', 'Perdido'];
const LEGACY = { 'Roto': 'Dañado', 'En Reparacion': 'En Mantenimiento' };
const HEAD_PUP = ['ID_Pupitre', 'Curso', 'Aula', 'Estado', 'Situacion', 'Ubicacion_Actual', 'Ultimo_Usuario',
  'Ultima_Actualizacion', 'Fecha_Salida_Carpintero', 'Fecha_Retorno_Carpintero', 'Observaciones', 'QR_URL',
  'Foto_URL', 'Ultimo_Mantenimiento', 'Proximo_Mantenimiento', 'Estado_Desde'];
const CFG_DEFECTO = {
  Emails_Notificacion: '', Meses_Mantenimiento: 6, Dias_Max_Mantenimiento: 7,
  Dias_Max_Carpinteria: 15, Dias_Aviso_Mantenimiento: 15, URL_App: '', Carpeta_Fotos_ID: ''
};

/* ───────── Configuración ───────── */

let _cfg = null;
function cfgMapa_() {
  if (_cfg) return _cfg;
  _cfg = {};
  hoja_('Config').getDataRange().getValues().forEach((r, i) => { if (i > 0 && r[0]) _cfg[r[0]] = r[1]; });
  return _cfg;
}
function cfg_(k, def) {
  const v = cfgMapa_()[k];
  return (v === '' || v === undefined) ? (def !== undefined ? def : CFG_DEFECTO[k]) : v;
}
function setCfg_(k, valor) {
  const h = hoja_('Config');
  const v = h.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (v[i][0] === k) { h.getRange(i + 1, 2).setValue(valor); _cfg = null; return; }
  }
  h.appendRow([k, valor]);
  _cfg = null;
}

function obtenerConfig(token) {
  return protegido_(token, ['Director'], () => {
    const c = {};
    Object.keys(CFG_DEFECTO).filter(k => k !== 'Carpeta_Fotos_ID').forEach(k => { c[k] = cfg_(k); });
    c.URL_Detectada = ScriptApp.getService().getUrl();
    return { ok: true, config: c };
  });
}

function guardarConfig(token, c) {
  return protegido_(token, ['Director'], s => {
    ['Emails_Notificacion', 'Meses_Mantenimiento', 'Dias_Max_Mantenimiento', 'Dias_Max_Carpinteria',
      'Dias_Aviso_Mantenimiento', 'URL_App'].forEach(k => {
      if (c[k] === undefined) return;
      let v = c[k];
      if (k.indexOf('Dias') === 0 || k.indexOf('Meses') === 0) {
        v = parseInt(v, 10);
        if (!(v >= 1 && v <= 120)) throw new Error('Valor inválido en ' + k);
      } else v = String(v).trim();
      setCfg_(k, v);
    });
    registrarMov_('', s.usuario, s.rol, 'CONFIG_GUARDADA', '', '', '', '');
    return { ok: true };
  });
}

/* ───────── Lectura ───────── */

function normEstado_(e) {
  e = String(e || '').trim();
  return LEGACY[e] || (ESTADOS.indexOf(e) >= 0 ? e : 'Bueno');
}

function leerTodos_() {
  const h = hoja_('Pupitres');
  if (h.getMaxColumns() < HEAD_PUP.length) h.insertColumnsAfter(h.getMaxColumns(), HEAD_PUP.length - h.getMaxColumns());
  const n = h.getLastRow() - 1;
  if (n < 1) return [];
  const out = [];
  h.getRange(2, 1, n, HEAD_PUP.length).getValues().forEach((r, i) => {
    if (!r[0]) return;
    out.push({
      fila: i + 2, id: String(r[0]).trim(), curso: String(r[1]), aula: String(r[2]),
      estado: normEstado_(r[3]), situacion: r[4], ubicacion: String(r[5] || r[2] || ''),
      usuario: r[6], actualizado: r[7], obs: r[10], foto: r[12],
      ultMant: r[13], proxMant: r[14], desde: r[15] || r[7]
    });
  });
  return out;
}

function plano_(p) {
  return {
    id: p.id, curso: p.curso, aula: p.aula, estado: p.estado, situacion: p.situacion,
    ubicacion: p.ubicacion, usuario: p.usuario, obs: p.obs, foto: p.foto,
    actualizado: iso_(p.actualizado), ultMant: iso_(p.ultMant), proxMant: iso_(p.proxMant), desde: iso_(p.desde)
  };
}

function buscar_(id) {
  const k = String(id || '').trim().toUpperCase();
  return leerTodos_().filter(p => p.id.toUpperCase() === k)[0] || null;
}

/** Una sola llamada para pintar todo el panel. */
function cargarInicio(token) {
  return protegido_(token, null, () => {
    const lista = leerTodos_();
    return {
      ok: true, pupitres: lista.map(plano_), alertas: calcularAlertas_(lista),
      appUrl: appUrl_(), estados: ESTADOS
    };
  });
}

function obtenerPupitre(token, id) {
  return protegido_(token, null, () => {
    const p = buscar_(id);
    if (!p) return { ok: false, error: 'Pupitre no encontrado: ' + id };
    return { ok: true, pupitre: plano_(p), historial: historial_(p.id) };
  });
}

function historial_(id) {
  const h = hoja_('Movimientos');
  const ultima = h.getLastRow();
  if (ultima < 2) return [];
  const desde = Math.max(2, ultima - 2000);
  const v = h.getRange(desde, 1, ultima - desde + 1, 9).getValues();
  const out = [];
  for (let i = v.length - 1; i >= 0 && out.length < 15; i--) {
    if (String(v[i][1]).toUpperCase() !== String(id).toUpperCase()) continue;
    out.push({ fecha: iso_(v[i][0]), usuario: v[i][2], accion: v[i][4], anterior: v[i][5], nuevo: v[i][6], nota: v[i][7], foto: v[i][8] });
  }
  return out;
}

/* ───────── Registro ───────── */

function registrarMov_(id, usuario, rol, accion, anterior, nuevo, notas, foto) {
  try {
    hoja_('Movimientos').appendRow([new Date(), id || '', usuario, rol, accion, anterior || '', nuevo || '', notas || '', foto || '']);
  } catch (e) { Logger.log('Movimiento no registrado: ' + e.message); }
}

function situacionDe_(estado, actual) {
  if (estado === 'En Carpinteria') return 'Con Carpintero';
  if (estado === 'Perdido') return 'Extraviado';
  if (estado === 'Dado de Baja') return 'Dado de Baja';
  return (actual === 'En Aula' || actual === 'En Deposito') ? actual : 'En Aula';
}

/* ───────── Fotos ───────── */

function carpetaFotos_() {
  const id = cfg_('Carpeta_Fotos_ID', '');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* se recrea */ } }
  const f = DriveApp.createFolder('Pupitres Vida y Verdad - Fotos');
  setCfg_('Carpeta_Fotos_ID', f.getId());
  return f;
}

function guardarFoto_(dataUrl, id) {
  const m = /^data:image\/(jpeg|png|webp);base64,([A-Za-z0-9+/=]+)$/.exec(String(dataUrl));
  if (!m) throw new Error('Formato de foto no válido.');
  if (m[2].length > 3500000) throw new Error('La foto es demasiado grande.');
  const ext = m[1] === 'jpeg' ? 'jpg' : m[1];
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), 'image/' + m[1],
    id + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.' + ext);
  const file = carpetaFotos_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
}

/* ───────── Cambios de estado ───────── */

function cambiarEstado(token, d) {
  return protegido_(token, null, s => {
    const estado = String(d.estado || '');
    if (ESTADOS.indexOf(estado) < 0) return { ok: false, error: 'Estado no válido.' };
    if (s.rol === 'Conserje' && estado !== 'Dañado') return { ok: false, error: 'El Conserje solo puede reportar pupitres dañados.' };
    if (estado === 'Dado de Baja' && s.rol !== 'Director') return { ok: false, error: 'Solo el Director puede dar de baja.' };

    return conLock_(() => {
      const p = buscar_(d.id);
      if (!p) return { ok: false, error: 'Pupitre no encontrado.' };
      const h = hoja_('Pupitres');
      const ahora = new Date();
      const nota = String(d.nota || '').slice(0, 500);
      const previo = p.estado;
      const cambio = previo !== estado;
      const fotoUrl = d.foto ? guardarFoto_(d.foto, p.id) : '';
      const f = p.fila;

      h.getRange(f, 4, 1, 5).setValues([[estado, situacionDe_(estado, p.situacion), p.ubicacion, s.usuario, ahora]]);
      if (nota) h.getRange(f, 11).setValue(nota);
      if (fotoUrl) h.getRange(f, 13).setValue(fotoUrl);
      if (cambio) h.getRange(f, 16).setValue(ahora);

      if (estado === 'En Carpinteria' && cambio) {
        h.getRange(f, 9).setValue(ahora);
        const esperado = new Date(ahora.getTime() + Number(cfg_('Dias_Max_Carpinteria')) * 86400000);
        hoja_('Carpintero').appendRow([ahora, p.id, p.curso, nota || 'Enviado desde el sistema', s.usuario, esperado, '', '', '']);
      }
      if (previo === 'En Carpinteria' && cambio) {
        h.getRange(f, 10).setValue(ahora);
        const hc = hoja_('Carpintero');
        const vc = hc.getDataRange().getValues();
        for (let i = vc.length - 1; i >= 1; i--) {
          if (String(vc[i][1]) === p.id && !vc[i][6]) {
            hc.getRange(i + 1, 7, 1, 3).setValues([[ahora, estado, nota]]);
            break;
          }
        }
      }
      if (estado === 'Bueno' && (cambio || d.mant)) {
        h.getRange(f, 14, 1, 2).setValues([[ahora, sumarMeses_(ahora, Number(cfg_('Meses_Mantenimiento')))]]);
      }

      if (estado !== 'Bueno' && (cambio || nota || fotoUrl)) {
        hoja_('Incidencias').appendRow([ahora, p.id, p.curso, p.ubicacion, s.usuario, estado, false, nota, fotoUrl]);
      } else if (estado === 'Bueno') {
        resolverIncidencias_(p.id);
      }

      registrarMov_(p.id, s.usuario, s.rol, d.mant && !cambio ? 'MANTENIMIENTO_REALIZADO' : 'CAMBIO_ESTADO', previo, estado, nota, fotoUrl);
      if (estado !== 'Bueno' && (cambio || fotoUrl)) {
        try { notificarReporte_(p, estado, nota, s, fotoUrl); } catch (e) { Logger.log('Notificación falló: ' + e.message); }
      }
      return { ok: true, mensaje: p.id + ' → ' + estado, pupitre: plano_(buscar_(p.id)) };
    });
  });
}

function resolverIncidencias_(id) {
  const h = hoja_('Incidencias');
  const v = h.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][1]) === id && v[i][6] !== true) h.getRange(i + 1, 7).setValue(true);
  }
}

function sumarMeses_(fecha, meses) {
  const d = new Date(fecha.getTime());
  d.setMonth(d.getMonth() + meses);
  return d;
}

/** Cambia la ubicación actual (p. ej. un pupitre encontrado en otra aula). */
function moverPupitre(token, d) {
  return protegido_(token, null, s => conLock_(() => {
    const p = buscar_(d.id);
    const ubic = String(d.ubicacion || '').trim().slice(0, 60);
    if (!p) return { ok: false, error: 'Pupitre no encontrado.' };
    if (!ubic) return { ok: false, error: 'Indique la ubicación.' };
    const h = hoja_('Pupitres');
    h.getRange(p.fila, 6).setValue(ubic);
    h.getRange(p.fila, 7, 1, 2).setValues([[s.usuario, new Date()]]);
    registrarMov_(p.id, s.usuario, s.rol, 'MOVIDO', p.ubicacion, ubic, '', '');
    return { ok: true, mensaje: p.id + ' ahora está en ' + ubic, pupitre: plano_(buscar_(p.id)) };
  }));
}

/* ───────── Alta de pupitres por lote ───────── */

function crearPupitres(token, d) {
  return protegido_(token, ['Director'], s => conLock_(() => {
    const curso = String(d.curso || '').trim().slice(0, 40);
    const aula = String(d.aula || '').trim().slice(0, 40) || curso;
    const cantidad = parseInt(d.cantidad, 10);
    const estado = ESTADOS.indexOf(d.estado) >= 0 ? d.estado : 'Bueno';
    if (!curso) return { ok: false, error: 'Indique el curso.' };
    if (!(cantidad >= 1 && cantidad <= 500)) return { ok: false, error: 'La cantidad debe estar entre 1 y 500.' };

    const todos = leerTodos_();
    const existente = todos.filter(p => p.aula.toLowerCase() === aula.toLowerCase())[0];
    let prefijo = existente ? existente.id.replace(/-\d+$/, '') : 'P-' + (aula.toUpperCase().replace(/[^A-Z0-9]/g, '') || 'X');
    const re = new RegExp('^' + prefijo.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '-(\\d+)$', 'i');
    let max = 0;
    todos.forEach(p => { const m = re.exec(p.id); if (m) max = Math.max(max, parseInt(m[1], 10)); });

    const ahora = new Date();
    const url = appUrl_();
    const prox = estado === 'Bueno' ? sumarMeses_(ahora, Number(cfg_('Meses_Mantenimiento'))) : '';
    const ids = [], filas = [];
    for (let i = 1; i <= cantidad; i++) {
      const id = prefijo + '-' + String(max + i).padStart(3, '0');
      ids.push(id);
      filas.push([id, curso, aula, estado, situacionDe_(estado, 'En Aula'), aula, s.usuario, ahora, '', '', '',
        url + '?p=' + id, '', '', prox, ahora]);
    }
    const h = hoja_('Pupitres');
    h.getRange(h.getLastRow() + 1, 1, filas.length, HEAD_PUP.length).setValues(filas);
    registrarMov_('', s.usuario, s.rol, 'LOTE_CREADO', '', cantidad + ' × ' + aula, estado, '');
    return { ok: true, ids: ids, mensaje: cantidad + ' pupitres creados en ' + aula };
  }));
}
