/**
 * Recibe el progreso de las pizarras (robot, detectives de problemas…).
 * - Una pestaña por grupo (PEP 8, PEP 9…): una fila por alumna, siempre la más reciente.
 * - Una pestaña "Historial" con todos los envíos, por si quieres ver la evolución.
 */
function norm(s) {
  return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function limpia(v) {
  if (typeof v === 'string' && /^\d+\/\d+$/.test(v.trim())) {
    return v.trim().replace('/', ' de ');   // evita que "5/10" se convierta en fecha
  }
  return v;
}

function getHoja(ss, nombre) {
  var sh = ss.getSheetByName(nombre);
  if (!sh) sh = ss.insertSheet(nombre);
  return sh;
}

function cabeceras(sh, base) {
  var lc = sh.getLastColumn();
  var h = lc > 0 ? sh.getRange(1, 1, 1, lc).getValues()[0] : [];
  if (h.length === 0) {
    h = base.slice();
    sh.getRange(1, 1, 1, h.length).setValues([h]);
    sh.setFrozenRows(1);
  }
  return h;
}

function aseguraColumnas(sh, h, claves) {
  var nuevas = false;
  claves.forEach(function (k) {
    if (h.indexOf(k) === -1) { h.push(k); nuevas = true; }
  });
  if (nuevas) sh.getRange(1, 1, 1, h.length).setValues([h]);
  return h;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);   // si llegan muchas a la vez, las pone en cola
  try {
    var datos = JSON.parse(e.postData.contents);
    if (datos.tipo === 'nota') return guardaNota(datos);
    if (datos.tipo === 'ocultar') return ocultaNota(datos);
    var nombre = String(datos.nombre || '').trim() || 'Sin nombre';
    var puntos = datos.puntos || {};
    var grupo = String(puntos.Grupo || '').trim() || 'Sin grupo';
    var sinGrupo = {};
    Object.keys(puntos).forEach(function (k) { if (k !== 'Grupo') sinGrupo[k] = limpia(puntos[k]); });
    var claves = Object.keys(sinGrupo);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ahora = new Date();

    // 1) Pestaña del grupo: una fila por alumna
    var nomPestana = grupo.replace(/[\[\]\*\?\/\\:]/g, '-').substring(0, 90);
    var sh = getHoja(ss, nomPestana);
    var h = cabeceras(sh, ['Fecha', 'Nombre']);
    h = aseguraColumnas(sh, h, claves);
    var ultima = sh.getLastRow(), fila = -1;
    if (ultima > 1) {
      var nombres = sh.getRange(2, 2, ultima - 1, 1).getValues();
      var buscado = norm(nombre);
      for (var i = 0; i < nombres.length; i++) {
        if (norm(nombres[i][0]) === buscado) { fila = i + 2; break; }
      }
    }
    if (fila === -1) {
      var nueva = h.map(function (c) {
        if (c === 'Fecha') return ahora;
        if (c === 'Nombre') return nombre;
        return sinGrupo.hasOwnProperty(c) ? sinGrupo[c] : '';
      });
      sh.appendRow(nueva);
      if (sh.getLastRow() > 2) sh.getRange(2, 1, sh.getLastRow() - 1, h.length).sort(2);
    } else {
      var actual = sh.getRange(fila, 1, 1, h.length).getValues()[0];
      var actualizada = h.map(function (c, i) {
        if (c === 'Fecha') return ahora;
        if (c === 'Nombre') return actual[i] || nombre;
        if (sinGrupo.hasOwnProperty(c)) return sinGrupo[c];
        return actual[i] === undefined ? '' : actual[i];
      });
      sh.getRange(fila, 1, 1, h.length).setValues([actualizada]);
    }

    // 2) Historial: todos los envíos
    var hs = getHoja(ss, 'Historial');
    var hh = cabeceras(hs, ['Fecha', 'Grupo', 'Nombre']);
    hh = aseguraColumnas(hs, hh, claves);
    hs.appendRow(hh.map(function (c) {
      if (c === 'Fecha') return ahora;
      if (c === 'Grupo') return grupo;
      if (c === 'Nombre') return nombre;
      return sinGrupo.hasOwnProperty(c) ? sinGrupo[c] : '';
    }));

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/* ================= PREGUNTAS PARA PENSAR (indagacion.html) ================= */
/* Cambia este código por uno tuyo. Solo con él se pueden VER y OCULTAR las notas. */
var CODIGO_DOCENTE = 'CAMBIA-ESTE-CODIGO';

function jsonOut(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function hojaNotas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Indagación');
  if (!sh) {
    sh = ss.insertSheet('Indagación');
    sh.getRange(1, 1, 1, 8).setValues([['Fecha', 'Id', 'Grupo', 'Tarjeta', 'Nombre', 'Nota', 'Color', 'Oculta']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function guardaNota(d) {
  var texto = String(d.texto || '').trim().substring(0, 400);
  if (!texto) return jsonOut({ ok: false, error: 'vacia' });
  if (/^[=+\-@]/.test(texto)) texto = ' ' + texto;   // evita fórmulas
  var sh = hojaNotas(), id = String(d.id || '');
  var n = sh.getLastRow();
  if (id && n > 1) {   // si llega repetida (reintento), no se duplica
    var ids = sh.getRange(2, 2, n - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return jsonOut({ ok: true, repetida: true });
  }
  sh.appendRow([new Date(), id, String(d.grupo || '(sin grupo)').substring(0, 40), String(d.tarjeta || ''),
    String(d.nombre || 'Sin nombre').substring(0, 60), texto, String(d.color || '').substring(0, 10), '']);
  return jsonOut({ ok: true });
}

function ocultaNota(d) {
  if (String(d.codigo) !== CODIGO_DOCENTE) return jsonOut({ ok: false, error: 'codigo' });
  var sh = hojaNotas(), n = sh.getLastRow();
  if (n > 1) {
    var ids = sh.getRange(2, 2, n - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(d.id)) { sh.getRange(i + 2, 8).setValue('1'); break; }
  }
  return jsonOut({ ok: true });
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.accion === 'notas') {
    if (String(p.codigo) !== CODIGO_DOCENTE) return jsonOut({ ok: false, error: 'codigo' });
    var sh = hojaNotas(), n = sh.getLastRow(), out = [];
    if (n > 1) {
      var v = sh.getRange(2, 1, n - 1, 8).getValues();
      v.forEach(function (r) {
        if (String(r[3]) === String(p.tarjeta) && String(r[7]) !== '1') {
          out.push({ f: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), id: String(r[1]), grupo: r[2], nombre: r[4], texto: r[5], color: r[6] });
        }
      });
    }
    return jsonOut({ ok: true, notas: out });
  }
  return jsonOut({ ok: true, msg: 'activo' });
}

/** Para probar desde el editor: Ejecutar ▶ con esta función seleccionada. */
function probar() {
  doPost({ postData: { contents: JSON.stringify({
    nombre: 'Prueba Maribel',
    puntos: { Grupo: 'PEP 8', ROBOT_TotalRetos: 3, ROBOT_AciertoPrimera_pct: 67, DETECTIVES_UltimoSimulacro: '5/10' }
  }) } });
}
