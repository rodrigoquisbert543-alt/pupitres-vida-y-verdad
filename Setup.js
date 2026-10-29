/**
 * INICIALIZACIÓN / MIGRACIÓN (no destructiva). Ejecutar una vez desde el editor:
 * inicializarSistema(). Es seguro volver a ejecutarla: no borra datos.
 */

function inicializarSistema() {
  const ss = ss_();
  const defs = {
    Pupitres: [HEAD_PUP, '#0b3a8f'],
    Usuarios: [['Usuario', 'Nombre_Completo', 'Rol', 'Hash_Contrasena', 'Activo', 'Ultimo_Acceso', 'Debe_Cambiar'], '#0f9d58'],
    Movimientos: [['Timestamp', 'ID_Pupitre', 'Usuario', 'Rol', 'Accion', 'Valor_Anterior', 'Valor_Nuevo', 'Notas', 'Foto_URL'], '#f4b400'],
    Incidencias: [['Timestamp', 'ID_Pupitre', 'Curso_Asignado', 'Ubicacion', 'Usuario', 'Tipo', 'Resuelta', 'Notas', 'Foto_URL'], '#db4437'],
    Carpintero: [['Fecha_Envio', 'ID_Pupitre', 'Curso', 'Descripcion_Falla', 'Enviado_Por', 'Fecha_Retorno_Esperada', 'Fecha_Retorno_Real', 'Estado_Retorno', 'Notas'], '#9c27b0'],
    Config: [['Parámetro', 'Valor'], '#607d8b']
  };
  Object.keys(defs).forEach(nombre => {
    const h = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
    const cab = defs[nombre][0];
    if (h.getMaxColumns() < cab.length) h.insertColumnsAfter(h.getMaxColumns(), cab.length - h.getMaxColumns());
    h.getRange(1, 1, 1, cab.length).setValues([cab]).setBackground(defs[nombre][1])
      .setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    h.setFrozenRows(1);
  });

  // Usuarios por defecto sólo si la hoja está vacía
  const hu = ss.getSheetByName('Usuarios');
  if (hu.getLastRow() < 2) {
    hu.getRange(2, 1, 3, 7).setValues([
      ['director', 'Director Administrativo', 'Director', hashPassword('director2024'), true, '', true],
      ['regente', 'Regente', 'Regente', hashPassword('regente2024'), true, '', true],
      ['conserje', 'Conserje', 'Conserje', hashPassword('conserje2024'), true, '', true]
    ]);
  } else {
    // Usuarios existentes que aún tienen la clave por defecto deben crear la suya
    const v = hu.getRange(2, 1, hu.getLastRow() - 1, 7).getValues();
    v.forEach((r, i) => {
      if (r[0] && r[6] === '' && r[3] === hashPassword(String(r[0]).toLowerCase() + '2024')) hu.getRange(i + 2, 7).setValue(true);
    });
  }
  if (hu.getMaxRows() < 201) hu.insertRowsAfter(hu.getMaxRows(), 201 - hu.getMaxRows());
  hu.getRange(2, 3, 200, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Director', 'Regente', 'Conserje'], true).build());

  // Config: agrega parámetros faltantes
  _cfg = null;
  const hc = ss.getSheetByName('Config');
  const existentes = cfgMapa_();
  if (!existentes.Nombre_Colegio) hc.appendRow(['Nombre_Colegio', 'Colegio Vida y Verdad']);
  Object.keys(CFG_DEFECTO).forEach(k => { if (!(k in existentes)) hc.appendRow([k, CFG_DEFECTO[k]]); });
  _cfg = null;

  migrarPupitres_(ss.getSheetByName('Pupitres'));
  instalarTriggers();
  Logger.log('✅ Sistema listo. URL: ' + ScriptApp.getService().getUrl());
}

function migrarPupitres_(h) {
  if (h.getMaxRows() < 5001) h.insertRowsAfter(h.getMaxRows(), 5001 - h.getMaxRows());
  h.getRange(2, 4, 5000, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).build());
  h.getRange(2, 5, 5000, 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['En Aula', 'En Deposito', 'Con Carpintero', 'Repuesto', 'Extraviado', 'Dado de Baja'], true).build());
  const n = h.getLastRow() - 1;
  if (n < 1) return;
  const rango = h.getRange(2, 1, n, HEAD_PUP.length);
  const v = rango.getValues();
  const url = appUrl_();
  const ahora = new Date();
  const meses = Number(cfg_('Meses_Mantenimiento'));
  v.forEach(r => {
    if (!r[0]) return;
    r[3] = normEstado_(r[3]);
    if (url) r[11] = url + '?p=' + r[0];
    if (!r[15]) r[15] = r[7] || ahora;
    if (!r[14] && r[3] === 'Bueno') r[14] = sumarMeses_(ahora, meses);
  });
  rango.setValues(v);
}
