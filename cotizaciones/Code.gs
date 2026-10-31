/**
 * FÁRA · Cotizaciones: puente entre la planilla y faraimagen.com.ar/cotizacion/<codigo>.
 *
 * doGet() ......... recibe un código y devuelve solo esa cotización.
 * setup() ......... arma la planilla entera. Se ejecuta una sola vez.
 * Menú FÁRA ....... "Nueva cotización" genera el código y deja la fila lista.
 * onEdit() ........ completa datos del Catálogo y la última modificación.
 *
 * Instrucciones en cotizaciones/README.md.
 */

const SITIO = 'https://faraimagen.com.ar/cotizacion/';
const LETRAS = 'abcdefghjkmnpqrstuvwxyz23456789'; // sin 0, o, 1, l, i
const LARGO_CODIGO = 8;
const FILAS = 500;
const CACHE_SEGUNDOS = 21600; // 6 horas, el máximo de Google

const H_USO = 'Cómo usar';
const H_COT = 'Cotizaciones';
const H_ITEMS = 'Ítems';
const H_CAT = 'Catálogo';
const H_CONFIG = 'Config';

const COL = {
  codigo: 0, clienteCodigo: 1, cliente: 2, contacto: 3, titulo: 4, intro: 5, emision: 6, vencimiento: 7,
  estado: 8, moneda: 9, descuentoTipo: 10, descuentoValor: 11, descuentoMotivo: 12, formaPago: 13,
  anticipoPct: 14, plazoEntrega: 15, condiciones: 16, notas: 17, disclaimer: 18, logoCliente: 19,
  total: 20, actualizada: 21, link: 22
};
const COL_ITEM = {
  codigo: 0, orden: 1, servicio: 2, descripcion: 3, entregables: 4, cantidad: 5, unidad: 6,
  precio: 7, descuentoPct: 8, mostrar: 9, importe: 10
};

/* ---------- web ---------- */

function doGet(e) {
  const pedido = String((e && e.parameter && e.parameter.codigo) || '');
  if (!/^[A-Z]{3}-[a-z2-9]{6,8}$/.test(pedido)) return responder('{"ok":false}');
  try {
    const guardada = CacheService.getScriptCache().get(claveCache(pedido));
    if (guardada) return responder(guardada);
    return responder(guardarEnCache(pedido));
  } catch (err) {
    return responder('{"ok":false}');
  }
}

function responder(json) {
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- copia rápida ----------
   Abrir la planilla es lo que más tarda, así que cada cotización queda lista en la
   caché de Google. Al editarla en la planilla, onEdit() la vuelve a armar en el momento. */

function claveCache(codigo) {
  return 'cot:' + codigo;
}

function guardarEnCache(codigo) {
  const datos = buscarCotizacion(codigo);
  const json = JSON.stringify(datos);
  if (datos.ok) CacheService.getScriptCache().put(claveCache(codigo), json, CACHE_SEGUNDOS);
  else CacheService.getScriptCache().remove(claveCache(codigo));
  return json;
}

// mantiene listas las cotizaciones enviadas: lo ejecuta el activador de "Activar carga rápida"
function refrescarCache() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H_COT).getRange(2, 1, FILAS, COL.estado + 1).getValues()
    .filter(function (r) { return texto(r[COL.codigo]) !== '' && r[COL.estado] === 'Enviada'; })
    .forEach(function (r) { guardarEnCache(texto(r[COL.codigo])); });
}

function activarCargaRapida() {
  const existe = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'refrescarCache'; });
  if (!existe) ScriptApp.newTrigger('refrescarCache').timeBased().everyHours(4).create();
  refrescarCache();
  SpreadsheetApp.getUi().alert('Listo: las cotizaciones enviadas quedan siempre preparadas.');
}

function buscarCotizacion(codigo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const zona = ss.getSpreadsheetTimeZone();
  const hoja = ss.getSheetByName(H_COT);
  const celda = buscarCodigo(hoja, codigo).findNext();
  if (!celda) return { ok: false };
  const f = hoja.getRange(celda.getRow(), 1, 1, COL.link + 1).getValues()[0];

  const estado = String(f[COL.estado]).trim().toLowerCase();
  if (['enviada', 'aceptada', 'vencida', 'pausada', 'anulada'].indexOf(estado) < 0) return { ok: false };
  // pausada y anulada no entregan ningún dato de la cotización
  if (estado === 'pausada' || estado === 'anulada') return { ok: true, estado: estado, codigo: codigo };

  const cfg = leerConfig(ss);
  const logo = texto(f[COL.logoCliente]);
  return {
    ok: true,
    estado: estado,
    codigo: codigo,
    cliente: texto(f[COL.cliente]),
    contacto: texto(f[COL.contacto]),
    titulo: texto(f[COL.titulo]),
    intro: texto(f[COL.intro]),
    emision: dia(f[COL.emision], zona),
    vencimiento: dia(f[COL.vencimiento], zona),
    moneda: f[COL.moneda] === 'USD' ? 'USD' : 'ARS',
    descuentoTipo: f[COL.descuentoTipo] === '%' ? '%' : 'monto',
    descuentoValor: f[COL.descuentoTipo] === '%' ? porcentaje(f[COL.descuentoValor]) : numero(f[COL.descuentoValor]),
    descuentoMotivo: texto(f[COL.descuentoMotivo]),
    formaPago: texto(f[COL.formaPago]),
    anticipoPct: porcentaje(f[COL.anticipoPct]),
    plazoEntrega: texto(f[COL.plazoEntrega]),
    condiciones: texto(f[COL.condiciones]) || texto(cfg.condiciones),
    notas: texto(f[COL.notas]),
    disclaimer: texto(f[COL.disclaimer]) || texto(cfg.disclaimer),
    logoCliente: /^https:\/\//.test(logo) ? logo : '',
    actualizada: dia(f[COL.actualizada], zona),
    items: leerItems(ss, codigo),
    // solo estos datos de Config salen a la web
    config: {
      whatsapp: texto(cfg.whatsapp),
      mail: texto(cfg.mail),
      instagram: texto(cfg.instagram),
      mediosPago: texto(cfg.mediosPago),
      impuestos: texto(cfg.impuestos)
    }
  };
}

function buscarCodigo(hoja, codigo) {
  return hoja.getRange(2, 1, FILAS, 1).createTextFinder(codigo).matchCase(true).matchEntireCell(true);
}

function leerItems(ss, codigo) {
  const hoja = ss.getSheetByName(H_ITEMS);
  const filas = buscarCodigo(hoja, codigo).findAll().map(function (c) { return c.getRow(); });
  if (!filas.length) return [];
  // solo el tramo donde están sus renglones, no la pestaña entera
  const desde = Math.min.apply(null, filas);
  const hasta = Math.max.apply(null, filas);
  return hoja.getRange(desde, 1, hasta - desde + 1, COL_ITEM.mostrar + 1).getValues()
    .filter(function (r) {
      return String(r[COL_ITEM.codigo]) === codigo && r[COL_ITEM.mostrar] === true && texto(r[COL_ITEM.servicio]) !== '';
    })
    .sort(function (a, b) { return numero(a[COL_ITEM.orden]) - numero(b[COL_ITEM.orden]); })
    .map(function (r) {
      return {
        servicio: texto(r[COL_ITEM.servicio]),
        descripcion: texto(r[COL_ITEM.descripcion]),
        entregables: texto(r[COL_ITEM.entregables]),
        cantidad: r[COL_ITEM.cantidad] === '' ? '' : numero(r[COL_ITEM.cantidad]),
        unidad: texto(r[COL_ITEM.unidad]),
        precio: r[COL_ITEM.precio] === '' ? '' : numero(r[COL_ITEM.precio]),
        descuentoPct: porcentaje(r[COL_ITEM.descuentoPct])
      };
    });
}

function leerConfig(ss) {
  const cfg = {};
  ss.getSheetByName(H_CONFIG).getDataRange().getValues().slice(1).forEach(function (r) {
    if (r[2]) cfg[r[2]] = r[1];
  });
  return cfg;
}

const texto = function (v) { return String(v === null || v === undefined ? '' : v).trim(); };
const numero = function (v) { const n = Number(v); return isFinite(n) ? n : 0; };
// acepta 10 y también 10% (que la planilla guarda como 0,1)
const porcentaje = function (v) { const n = numero(v); return n > 0 && n < 1 ? n * 100 : n; };
const dia = function (v, zona) {
  return Object.prototype.toString.call(v) === '[object Date]' ? Utilities.formatDate(v, zona, 'yyyy-MM-dd') : '';
};

/* ---------- menú ---------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('FÁRA')
    .addItem('Nueva cotización', 'nuevaCotizacion')
    .addSeparator()
    .addItem('Ver ítems de esta cotización', 'verItems')
    .addItem('Agregar ítem', 'agregarItem')
    .addItem('Ver todos los ítems', 'verTodosLosItems')
    .addSeparator()
    .addItem('Activar carga rápida', 'activarCargaRapida')
    .addToUi();
}

function nuevaCotizacion() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Nueva cotización', 'Código de cliente: 3 letras, por ejemplo CEN', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const prefijo = r.getResponseText().trim().toUpperCase();
  if (!/^[A-Z]{3}$/.test(prefijo)) {
    ui.alert('Tienen que ser 3 letras, sin acentos ni números.');
    return;
  }
  const creada = crearCotizacion(SpreadsheetApp.getActiveSpreadsheet(), prefijo);
  ui.alert('Cotización creada en Borrador', 'Completala y pasala a Enviada para activarla:\n\n' + SITIO + creada.codigo, ui.ButtonSet.OK);
}

function verItems() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const codigo = codigoActivo(ss);
  if (!codigo) {
    SpreadsheetApp.getUi().alert('Pará el cursor en la fila de una cotización y volvé a tocar Ver ítems.');
    return;
  }
  filtrarItems(ss.getSheetByName(H_ITEMS), codigo).activate();
}

function agregarItem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const codigo = codigoActivo(ss);
  if (!codigo) {
    SpreadsheetApp.getUi().alert('Pará el cursor en la fila de una cotización o abrí sus ítems con Ver ítems.');
    return;
  }
  const items = ss.getSheetByName(H_ITEMS);
  const fila = sumarRenglon(items, codigo);
  filtrarItems(items, codigo).activate();
  items.getRange(fila, COL_ITEM.servicio + 1).activate();
}

function verTodosLosItems() {
  const items = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H_ITEMS);
  quitarFiltro(items);
  items.activate();
}

// la cotización de la fila activa, o la que se está viendo en Ítems
function codigoActivo(ss) {
  const hoja = ss.getActiveSheet();
  const fila = ss.getActiveRange().getRow();
  if (hoja.getName() === H_COT) return fila > 1 ? texto(hoja.getRange(fila, COL.codigo + 1).getValue()) : '';
  if (hoja.getName() !== H_ITEMS) return '';
  const filtro = hoja.getFilter();
  const criterio = filtro && filtro.getColumnFilterCriteria(COL_ITEM.codigo + 1);
  if (criterio) return texto(criterio.getCriteriaValues()[0]);
  return fila > 1 ? texto(hoja.getRange(fila, COL_ITEM.codigo + 1).getValue()) : '';
}

function filtrarItems(items, codigo) {
  const filtro = items.getFilter() || items.getRange(1, 1, FILAS + 1, COL_ITEM.importe + 1).createFilter();
  filtro.setColumnFilterCriteria(COL_ITEM.codigo + 1, SpreadsheetApp.newFilterCriteria().whenTextEqualTo(codigo).build());
  return items;
}

function quitarFiltro(items) {
  const filtro = items.getFilter();
  if (filtro) filtro.remove();
}

function sumarRenglon(items, codigo) {
  const orden = items.getRange(2, 1, FILAS, 2).getValues()
    .filter(function (r) { return r[0] === codigo; })
    .reduce(function (max, r) { return Math.max(max, numero(r[1])); }, 0);
  const fila = primeraLibre(items);
  items.getRange(fila, COL_ITEM.codigo + 1, 1, 2).setValues([[codigo, orden + 1]]);
  items.getRange(fila, COL_ITEM.mostrar + 1).setValue(true);
  return fila;
}

function crearCotizacion(ss, prefijo) {
  const hoja = ss.getSheetByName(H_COT);
  const items = ss.getSheetByName(H_ITEMS);
  const cfg = leerConfig(ss);
  quitarFiltro(items);
  const codigo = codigoNuevo(hoja, prefijo);
  const hoy = new Date();
  const vence = new Date(hoy.getTime());
  vence.setDate(vence.getDate() + (numero(cfg.validezDias) || 15));

  const fila = primeraLibre(hoja);
  hoja.getRange(fila, COL.codigo + 1, 1, 2).setValues([[codigo, prefijo]]);
  hoja.getRange(fila, COL.emision + 1, 1, 4).setValues([[hoy, vence, 'Borrador', 'ARS']]);
  hoja.getRange(fila, COL.actualizada + 1).setValue(hoy);

  // tres renglones de ítems ya unidos a la cotización: los vacíos no se muestran
  [1, 2, 3].forEach(function () { sumarRenglon(items, codigo); });
  hoja.getRange(fila, COL.cliente + 1).activate();
  return { codigo: codigo, fila: fila };
}

function codigoNuevo(hoja, prefijo) {
  const usados = hoja.getRange(2, COL.codigo + 1, FILAS, 1).getValues().map(function (r) { return r[0]; });
  let codigo;
  do {
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid());
    let azar = '';
    for (let i = 0; i < LARGO_CODIGO; i++) azar += LETRAS[(bytes[i] & 255) % LETRAS.length];
    codigo = prefijo + '-' + azar;
  } while (usados.indexOf(codigo) >= 0);
  return codigo;
}

// getLastRow no sirve: las fórmulas automáticas llenan la columna hasta abajo
function primeraLibre(hoja) {
  const col = hoja.getRange(2, 1, FILAS, 1).getValues();
  for (let i = 0; i < col.length; i++) if (col[i][0] === '') return i + 2;
  throw new Error('La pestaña ' + hoja.getName() + ' está llena.');
}

/* ---------- automático al editar ---------- */

function onEdit(e) {
  if (!e || !e.range || e.range.getRow() < 2) return;
  const hoja = e.range.getSheet();
  const fila = e.range.getRow();

  if (hoja.getName() === H_CONFIG) {
    refrescarCache();
    return;
  }
  if (hoja.getName() === H_COT) {
    if (e.range.getColumn() !== COL.actualizada + 1) marcarCambio(hoja, fila);
    refrescarFilas(hoja, e.range);
    return;
  }
  if (hoja.getName() !== H_ITEMS) return;

  if (e.range.getNumRows() === 1 && e.range.getColumn() === COL_ITEM.servicio + 1) completarDesdeCatalogo(hoja, fila, e.value);

  const codigo = hoja.getRange(fila, COL_ITEM.codigo + 1).getValue();
  const cot = e.source.getSheetByName(H_COT);
  const n = cot.getRange(2, COL.codigo + 1, FILAS, 1).getValues().findIndex(function (r) { return r[0] === codigo; });
  if (codigo && n >= 0) marcarCambio(cot, n + 2);
  refrescarFilas(hoja, e.range);
}

// el código siempre está en la columna A, en Cotizaciones y en Ítems
function refrescarFilas(hoja, rango) {
  const codigos = hoja.getRange(rango.getRow(), 1, rango.getNumRows(), 1).getValues()
    .map(function (r) { return texto(r[0]); })
    .filter(function (c, i, todos) { return c !== '' && todos.indexOf(c) === i; });
  codigos.forEach(guardarEnCache);
}

function marcarCambio(hoja, fila) {
  hoja.getRange(fila, COL.actualizada + 1).setValue(new Date());
}

function completarDesdeCatalogo(hoja, fila, servicio) {
  const base = hoja.getParent().getSheetByName(H_CAT).getDataRange().getValues()
    .find(function (r) { return r[0] === servicio; });
  if (!base) return;
  const desc = hoja.getRange(fila, COL_ITEM.descripcion + 1);
  const precio = hoja.getRange(fila, COL_ITEM.precio + 1);
  if (desc.getValue() === '') desc.setValue(base[1]);
  if (precio.getValue() === '' && base[2] !== '') precio.setValue(base[2]);
}

/* ---------- armado de la planilla ---------- */

const C_CREMA = '#F2EBDD';
const C_CREMA2 = '#E9DFCB';
const C_ESPRESSO = '#1A1410';
const C_VERDE = '#7A8B6F';
const C_ROJO = '#A8473A';
const C_BANDA = '#F7F2E8';

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(H_COT)) throw new Error('La planilla ya está armada. No hace falta volver a ejecutar setup.');

  ss.setSpreadsheetLocale('es_AR');
  ss.setSpreadsheetTimeZone('America/Argentina/Buenos_Aires');
  const inicial = ss.getSheets()[0];

  armarUso(ss);
  armarCotizaciones(ss);
  armarItems(ss);
  armarCatalogo(ss);
  armarConfig(ss);
  ss.deleteSheet(inicial);

  cargarEjemplo(ss);
  ss.getSheetByName(H_USO).activate();
  ss.toast('Planilla lista. La cotización de ejemplo está en la pestaña Cotizaciones.', 'FÁRA', 10);
}

// encabezado: verde se puede tocar, rojo es automático
function encabezado(hoja, titulos, automaticas) {
  hoja.getRange(1, 1, 1, titulos.length).setValues([titulos])
    .setFontWeight('bold').setFontColor(C_CREMA).setBackground(C_VERDE)
    .setVerticalAlignment('middle').setWrap(true);
  automaticas.forEach(function (col) { hoja.getRange(1, col + 1).setBackground(C_ROJO); });
  hoja.setFrozenRows(1);
  hoja.setRowHeight(1, 44);
  hoja.getDataRange().setFontFamily('Archivo');
}

function lista(hoja, fila, col, filas, opciones) {
  const regla = SpreadsheetApp.newDataValidation().requireValueInList(opciones, true).setAllowInvalid(false).build();
  hoja.getRange(fila, col + 1, filas, 1).setDataValidation(regla);
}

// filas apenas alternadas para leer mejor
function bandas(hoja, columnas, formula) {
  hoja.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula).setBackground(C_BANDA)
    .setRanges([hoja.getRange(2, 1, FILAS, columnas)]).build()]);
}

function automatica(hoja, col) {
  hoja.getRange(2, col + 1, FILAS, 1).setBackground(C_CREMA2).protect()
    .setDescription('Columna automática').setWarningOnly(true);
}

function armarCotizaciones(ss) {
  const h = ss.insertSheet(H_COT);
  encabezado(h, ['Código', 'Código de cliente', 'Cliente', 'Contacto', 'Título de la propuesta', 'Mensaje de introducción',
    'Fecha de emisión', 'Vencimiento', 'Estado', 'Moneda', 'Tipo de descuento', 'Descuento', 'Motivo del descuento',
    'Forma de pago', 'Anticipo %', 'Plazo de entrega', 'Condiciones', 'Notas', 'Disclaimer', 'Logo del cliente (link)',
    'Total', 'Última modificación', 'Link'], [COL.codigo, COL.total, COL.actualizada, COL.link]);

  lista(h, 2, COL.estado, FILAS, ['Borrador', 'Enviada', 'Aceptada', 'Vencida', 'Anulada', 'Pausada']);
  lista(h, 2, COL.moneda, FILAS, ['ARS', 'USD']);
  lista(h, 2, COL.descuentoTipo, FILAS, ['%', 'Monto']);
  h.getRange(2, COL.emision + 1, FILAS, 2).setNumberFormat('dd/mm/yyyy');
  h.getRange(2, COL.actualizada + 1, FILAS, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  h.getRange(2, COL.descuentoValor + 1, FILAS, 1).setNumberFormat('#,##0.##');
  h.getRange(2, COL.anticipoPct + 1, FILAS, 1).setNumberFormat('0"%"');
  h.getRange(2, COL.total + 1, FILAS, 1).setNumberFormat('#,##0');
  h.getRange(2, COL.codigo + 1, FILAS, 1).setNumberFormat('@');

  // con la planilla en es_AR las fórmulas separan con punto y coma
  const neto = "SUMIF('" + H_ITEMS + "'!A2:A;A2:A;'" + H_ITEMS + "'!K2:K)";
  h.getRange(2, COL.total + 1).setFormula(
    '=ARRAYFORMULA(IF(A2:A="";"";IF(K2:K="%";' + neto + '*(1-IF(L2:L<1;L2:L;L2:L/100));' + neto + '-L2:L)))');
  h.getRange(2, COL.link + 1).setFormula('=ARRAYFORMULA(IF(A2:A="";"";"' + SITIO + '"&A2:A))');

  [COL.codigo, COL.total, COL.actualizada, COL.link].forEach(function (c) { automatica(h, c); });
  bandas(h, COL.link + 1, '=AND($A2<>"";ISEVEN(ROW()))');
  h.setColumnWidths(1, 23, 130);
  h.setColumnWidth(COL.intro + 1, 280);
  h.setColumnWidth(COL.titulo + 1, 240);
  h.setColumnWidth(COL.link + 1, 360);
  h.setFrozenColumns(3);
}

function armarItems(ss) {
  const h = ss.insertSheet(H_ITEMS);
  encabezado(h, ['Código', 'Orden', 'Servicio', 'Descripción', 'Entregables (uno por renglón)', 'Cantidad', 'Unidad',
    'Precio', 'Descuento %', 'Mostrar', 'Importe'], [COL_ITEM.importe]);

  const hojaCatalogo = ss.getSheetByName(H_CAT) || ss.insertSheet(H_CAT);
  const catalogo = SpreadsheetApp.newDataValidation()
    .requireValueInRange(hojaCatalogo.getRange('A2:A' + FILAS), true)
    .setAllowInvalid(true).build();
  h.getRange(2, COL_ITEM.servicio + 1, FILAS, 1).setDataValidation(catalogo);
  h.getRange(2, COL_ITEM.mostrar + 1, FILAS, 1).insertCheckboxes();
  h.getRange(2, COL_ITEM.precio + 1, FILAS, 1).setNumberFormat('#,##0.##');
  h.getRange(2, COL_ITEM.descuentoPct + 1, FILAS, 1).setNumberFormat('0"%"');
  h.getRange(2, COL_ITEM.importe + 1, FILAS, 1).setNumberFormat('#,##0');
  h.getRange(2, COL_ITEM.codigo + 1, FILAS, 1).setNumberFormat('@');
  h.getRange(2, COL_ITEM.importe + 1).setFormula(
    '=ARRAYFORMULA(IF(A2:A="";"";IF(J2:J=TRUE;H2:H*(1-IF(I2:I<1;I2:I;I2:I/100));0)))');

  automatica(h, COL_ITEM.importe);
  // cada cotización toma su color según su fila en Cotizaciones
  bandas(h, COL_ITEM.importe + 1, '=AND($A2<>"";ISEVEN(MATCH($A2;INDIRECT("' + H_COT + '!$A:$A");0)))');
  h.setColumnWidths(1, 11, 110);
  h.setColumnWidth(COL_ITEM.servicio + 1, 200);
  h.setColumnWidth(COL_ITEM.descripcion + 1, 300);
  h.setColumnWidth(COL_ITEM.entregables + 1, 240);
}

function armarCatalogo(ss) {
  const h = ss.getSheetByName(H_CAT) || ss.insertSheet(H_CAT);
  encabezado(h, ['Servicio', 'Descripción', 'Precio de referencia'], []);
  const base = ['Fotos de producto', 'Fotos de estudio', 'Landing', 'Menú digital', 'Impresión 3D', 'Asesoría de marketing'];
  h.getRange(2, 1, base.length, 3).setValues(base.map(function (s) { return [s, '', '']; }));
  h.getRange(2, 3, FILAS, 1).setNumberFormat('#,##0.##');
  h.setColumnWidth(1, 220);
  h.setColumnWidth(2, 380);
  h.setColumnWidth(3, 150);
}

function armarConfig(ss) {
  const h = ss.insertSheet(H_CONFIG);
  encabezado(h, ['Dato', 'Valor', 'Clave'], [2]);
  const filas = [
    ['WhatsApp', '5493885735415', 'whatsapp'],
    ['Mail', 'hola@faraimagen.com.ar', 'mail'],
    ['Instagram', '@fara.imagen', 'instagram'],
    ['Validez por defecto (días)', 15, 'validezDias'],
    ['Medios de pago (alias, bancos, cuotas)', '', 'mediosPago'],
    ['Disclaimer por defecto', '', 'disclaimer'],
    ['Condiciones por defecto', '', 'condiciones'],
    ['Impuestos (por ejemplo: Los precios no incluyen IVA.)', '', 'impuestos']
  ];
  h.getRange(2, 2, filas.length, 1).setNumberFormat('@');
  h.getRange(2, 1, filas.length, 3).setValues(filas);
  h.getRange(2, 3, filas.length, 1).setBackground(C_CREMA2).protect().setDescription('No cambiar').setWarningOnly(true);
  h.setColumnWidth(1, 340);
  h.setColumnWidth(2, 420);
  h.setColumnWidth(3, 120);
}

function armarUso(ss) {
  const h = ss.insertSheet(H_USO);
  const lineas = [
    ['FÁRA · Cotizaciones'],
    ['Esta planilla es privada. El cliente solo ve su cotización en la web, con su link.'],
    [''],
    ['Crear una cotización'],
    ['1. Menú FÁRA → Nueva cotización. Escribí las 3 letras del cliente.'],
    ['2. En Cotizaciones, completá la fila nueva.'],
    ['3. Parado en esa fila, tocá FÁRA → Ver ítems y completá los renglones. Para sumar otro: FÁRA → Agregar ítem. En Precio va el total del renglón: la cantidad no multiplica.'],
    ['4. Cambiá el Estado a Enviada y mandale el link de la última columna.'],
    [''],
    ['Cambiar una cotización'],
    ['Editá la celda. El cliente ve el cambio al recargar la página. Para volver a la lista completa: FÁRA → Ver todos los ítems.'],
    [''],
    ['Estados'],
    ['Borrador: el link dice que no está disponible.'],
    ['Enviada: se ve completa. Si pasa la fecha de vencimiento, se muestra vencida sola.'],
    ['Aceptada: se ve completa con el aviso de aceptada.'],
    ['Pausada y Anulada: un mensaje, sin datos ni precios.'],
    ['Vencida: se ve completa con el botón para pedir actualización.'],
    [''],
    ['Qué se puede tocar'],
    ['Encabezado verde: se puede editar.'],
    ['Encabezado rojo: automático, no lo toques.'],
    ['No borres cotizaciones: quedan como historial. Para sacarlas de la web, usá Anulada.']
  ];
  h.getRange(1, 1, lineas.length, 1).setValues(lineas).setFontFamily('Archivo').setBackground(C_CREMA).setFontColor(C_ESPRESSO).setWrap(true);
  h.setColumnWidth(1, 720);
  h.getRange(1, 1).setFontSize(20).setFontWeight('bold');
  [4, 10, 13, 20].forEach(function (f) { h.getRange(f, 1).setFontWeight('bold').setFontSize(13); });
  h.getRange(21, 1).setBackground(C_VERDE).setFontColor(C_CREMA);
  h.getRange(22, 1).setBackground(C_ROJO).setFontColor(C_CREMA);
  h.setHiddenGridlines(true);
}

// cliente inventado para probar: se puede anular y dejar como referencia
function cargarEjemplo(ss) {
  const creada = crearCotizacion(ss, 'EJE');
  const hoja = ss.getSheetByName(H_COT);
  const f = creada.fila;
  hoja.getRange(f, COL.cliente + 1, 1, 4).setValues([['Café Ejemplo', 'Laura Pérez',
    'Fotos de producto y landing para tu marca',
    'Hola Laura, gracias por la charla de ayer.\nAcá está la propuesta con lo que hablamos para la carta nueva y la web.']]);
  hoja.getRange(f, COL.estado + 1).setValue('Enviada');
  hoja.getRange(f, COL.descuentoTipo + 1, 1, 6).setValues([['%', 10, 'Descuento por contratar los tres servicios juntos.',
    'Transferencia bancaria.', 50, '15 días hábiles desde la sesión de fotos.']]);

  const items = ss.getSheetByName(H_ITEMS);
  items.getRange(2, COL_ITEM.servicio + 1, 3, 7).setValues([
    ['Fotos de producto', 'Sesión en estudio con fondo y luz definidos para tu marca.', '12 fotos editadas\nFormato web y redes', 12, 'fotos', 300000, 10],
    ['Landing', 'Una página con tu oferta y botón de WhatsApp.', '', 1, '', 350000, 0],
    ['Menú digital', 'Carta online que editás desde una planilla.', '', 1, '', 150000, 0]
  ]);
}
