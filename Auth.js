/**
 * AUTENTICACIÓN Y SESIONES - Sistema Pupitres Vida y Verdad
 * Las sesiones viven en el servidor (CacheService). El rol NUNCA se acepta del cliente.
 */

const SPREADSHEET_ID = '1nobCpwsgm9Q7gdXh7Nscwx3c2P2zx_oQquJdmW74G3k';
const SESION_SEG = 21600; // 6 h (máximo de CacheService); se renueva con cada uso
const MAX_INTENTOS = 5;

function ss_() { return SpreadsheetApp.openById(SPREADSHEET_ID); }

function hoja_(nombre) {
  const h = ss_().getSheetByName(nombre);
  if (!h) throw new Error('Falta la hoja "' + nombre + '". Ejecute inicializarSistema().');
  return h;
}

function hashPassword(password) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password));
  return raw.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function sesion_(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const raw = cache.get('S_' + token);
  if (!raw) return null;
  cache.put('S_' + token, raw, SESION_SEG);
  return JSON.parse(raw);
}

/** Ejecuta fn(sesion) sólo si el token es válido y el rol está permitido. */
function protegido_(token, roles, fn, permitirForzado) {
  try {
    const s = sesion_(token);
    if (!s) return { ok: false, sesion: true, error: 'Sesión expirada. Ingrese nuevamente.' };
    if (s.forzar && !permitirForzado) return { ok: false, forzar: true, error: 'Debe crear su contraseña personal primero.' };
    if (roles && roles.indexOf(s.rol) < 0) return { ok: false, error: 'No tiene permiso para esta acción.' };
    return fn(s);
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function conLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function login(usuario, contrasena) {
  try {
    const u = String(usuario || '').trim().toLowerCase();
    if (!u || !contrasena) return { ok: false, error: 'Ingrese usuario y contraseña.' };

    const cache = CacheService.getScriptCache();
    const claveIntentos = 'I_' + u;
    const intentos = parseInt(cache.get(claveIntentos) || '0', 10);
    if (intentos >= MAX_INTENTOS) return { ok: false, error: 'Demasiados intentos. Espere 10 minutos.' };

    const hoja = hoja_('Usuarios');
    const datos = hoja.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (!datos[i][0] || String(datos[i][0]).toLowerCase() !== u) continue;
      const nombreUsr = datos[i][0], nombre = datos[i][1], rol = datos[i][2], hash = datos[i][3], activo = datos[i][4];
      if (activo !== true && String(activo).toUpperCase() !== 'TRUE') return { ok: false, error: 'Usuario inactivo. Contacte al Director.' };
      if (hashPassword(contrasena) !== hash) {
        cache.put(claveIntentos, String(intentos + 1), 600);
        return { ok: false, error: 'Contraseña incorrecta.' };
      }
      cache.remove(claveIntentos);
      hoja.getRange(i + 1, 6).setValue(new Date());
      const token = Utilities.getUuid() + Utilities.getUuid();
      const forzar = datos[i][6] === true || String(datos[i][6]).toUpperCase() === 'TRUE';
      cache.put('S_' + token, JSON.stringify({ usuario: nombreUsr, nombre: nombre, rol: rol, forzar: forzar }), SESION_SEG);
      registrarMov_('', nombreUsr, rol, 'LOGIN', '', '', 'Inicio de sesión', '');
      return { ok: true, token: token, usuario: nombreUsr, nombre: nombre, rol: rol, cambiarClave: forzar };
    }
    cache.put(claveIntentos, String(intentos + 1), 600);
    return { ok: false, error: 'Usuario no encontrado.' };
  } catch (e) {
    return { ok: false, error: 'Error del sistema: ' + e.message };
  }
}

function sesionInfo(token) {
  const s = sesion_(token);
  return s ? { ok: true, usuario: s.usuario, nombre: s.nombre, rol: s.rol, cambiarClave: !!s.forzar } : { ok: false, sesion: true };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('S_' + token);
  return { ok: true };
}

function cambiarContrasena(token, actual, nueva) {
  return protegido_(token, null, s => {
    if (!nueva || String(nueva).length < 6) return { ok: false, error: 'La nueva contraseña debe tener al menos 6 caracteres.' };
    const hoja = hoja_('Usuarios');
    const datos = hoja.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).toLowerCase() === String(s.usuario).toLowerCase()) {
        if (hashPassword(actual) !== datos[i][3]) return { ok: false, error: 'La contraseña actual es incorrecta.' };
        if (actual === nueva) return { ok: false, error: 'La nueva contraseña debe ser distinta de la temporal.' };
        hoja.getRange(i + 1, 4).setValue(hashPassword(nueva));
        hoja.getRange(i + 1, 7).setValue(false);
        s.forzar = false;
        CacheService.getScriptCache().put('S_' + token, JSON.stringify(s), SESION_SEG);
        registrarMov_('', s.usuario, s.rol, 'CAMBIO_CONTRASENA', '', '', '', '');
        return { ok: true };
      }
    }
    return { ok: false, error: 'Usuario no encontrado.' };
  }, true);
}

function listarUsuarios(token) {
  return protegido_(token, ['Director'], () => {
    const datos = hoja_('Usuarios').getDataRange().getValues();
    const usuarios = [];
    for (let i = 1; i < datos.length; i++) {
      if (!datos[i][0] || !datos[i][2]) continue;
      usuarios.push({
        usuario: datos[i][0], nombre: datos[i][1], rol: datos[i][2],
        activo: datos[i][4] === true || String(datos[i][4]).toUpperCase() === 'TRUE',
        ultimoAcceso: iso_(datos[i][5])
      });
    }
    return { ok: true, usuarios: usuarios };
  });
}

/** Crea o actualiza un usuario. En actualización, la clave es opcional. */
function guardarUsuario(token, d) {
  return protegido_(token, ['Director'], s => conLock_(() => {
    const u = String(d.usuario || '').trim().toLowerCase();
    if (!/^[a-z0-9._-]{3,30}$/.test(u)) return { ok: false, error: 'Usuario inválido (3-30 letras/números).' };
    if (['Director', 'Regente', 'Conserje'].indexOf(d.rol) < 0) return { ok: false, error: 'Rol inválido.' };
    const hoja = hoja_('Usuarios');
    const datos = hoja.getDataRange().getValues();
    let fila = -1;
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).toLowerCase() === u) { fila = i + 1; break; }
    }
    if (fila < 0) {
      if (!d.clave || String(d.clave).length < 6) return { ok: false, error: 'La contraseña debe tener al menos 6 caracteres.' };
      hoja.appendRow([u, d.nombre || u, d.rol, hashPassword(d.clave), d.activo !== false, '', true]);
    } else {
      if (u === String(s.usuario).toLowerCase() && (d.activo === false || d.rol !== 'Director')) {
        return { ok: false, error: 'No puede quitarse el acceso de Director a sí mismo.' };
      }
      hoja.getRange(fila, 2, 1, 3).setValues([[d.nombre || u, d.rol, d.clave ? hashPassword(d.clave) : datos[fila - 1][3]]]);
      hoja.getRange(fila, 5).setValue(d.activo !== false);
      if (d.clave) hoja.getRange(fila, 7).setValue(true);
    }
    registrarMov_('', s.usuario, s.rol, 'USUARIO_GUARDADO', '', u, d.rol, '');
    return { ok: true };
  }));
}
