/**
 * FÁRA · Cotizaciones y entregas: puente entre la planilla y
 * faraimagen.com.ar/cotizacion/<codigo> y faraimagen.com.ar/entrega/<codigo>.
 *
 * doGet() ......... recibe un código y devuelve solo esa cotización o entrega.
 * setup() ......... arma la planilla entera. Se ejecuta una sola vez.
 * agregarEntregas() suma las pestañas de entregas a una planilla ya armada.
 * Menús FÁRA ...... "Nueva cotización" y "Nueva entrega" generan el código y dejan la fila lista.
 * onEdit() ........ completa datos del Catálogo, la última modificación y la fecha límite,
 *                   y crea la entrega de una cotización al aceptarla o marcar un pago.
 * avisosEntregas() manda los mails de vencimiento y marca las vencidas. Corre solo una vez por día.
 *
 * Instrucciones en cotizaciones/README.md y entregas/README.md.
 */

const SITIO = 'https://faraimagen.com.ar/cotizacion/';
const SITIO_ENTREGA = 'https://faraimagen.com.ar/entrega/';
const LETRAS = 'abcdefghjkmnpqrstuvwxyz23456789'; // sin 0, o, 1, l, i
const LARGO_CODIGO = 8;
const FILAS = 500;
const CACHE_SEGUNDOS = 21600; // 6 horas, el máximo de Google
const RE_CODIGO = /^[A-Z]{3}-[a-z2-9]{6,8}$/;
const PLAZO_ENTREGA = 30; // días, si Config no dice otra cosa
const AVISO_ENTREGA = 5;
const MAIL_FARA = 'hola@faraimagen.com.ar';

const H_USO = 'Cómo usar';
const H_COT = 'Cotizaciones';
const H_ITEMS = 'Ítems';
const H_CAT = 'Catálogo';
const H_CONFIG = 'Config';
const H_ENT = 'Entregas';
const H_ARC = 'Archivos';

const COL = {
  codigo: 0, clienteCodigo: 1, cliente: 2, contacto: 3, titulo: 4, intro: 5, emision: 6, vencimiento: 7,
  estado: 8, moneda: 9, descuentoTipo: 10, descuentoValor: 11, descuentoMotivo: 12, formaPago: 13,
  anticipoPct: 14, pago: 15, plazoEntrega: 16, condiciones: 17, notas: 18, disclaimer: 19,
  logoCliente: 20, total: 21, actualizada: 22, link: 23
};
const PAGOS = { 'Pendiente': 'pendiente', 'Anticipo pagado': 'anticipo', 'Pagado total': 'total' };
const COL_ITEM = {
  codigo: 0, orden: 1, servicio: 2, descripcion: 3, entregables: 4, cantidad: 5, unidad: 6,
  precio: 7, descuentoPct: 8, mostrar: 9, importe: 10
};
const COL_ENT = {
  codigo: 0, clienteCodigo: 1, empresa: 2, saludo: 3, mail: 4, titulo: 5, mensaje: 6, logoCliente: 7, limite: 8,
  estado: 9, avisoPrevio: 10, avisoUltimo: 11, link: 12, cotizacion: 13
};
const COL_ARC = { codigo: 0, orden: 1, nombre: 2, tipo: 3, link: 4, detalle: 5, fecha: 6, visible: 7 };
const ESTADOS_ENT = {
  'Pendiente de pago': 'pendiente', 'En proceso': 'proceso', 'Disponible': 'disponible',
  'Vencida': 'borrada', 'Archivos borrados': 'borrada', 'Pausada': 'pausada', 'Anulada': 'anulada'
};
const TIPOS_ARC = { 'Descarga': 'descarga', 'Web': 'web', 'Entregado': 'entregado', 'En proceso': 'proceso' };
const VISIBLES_ARC = { 'Al pagar': 'pago', 'Entrega parcial': 'parcial', 'Oculto': 'oculto' };

/* ---------- web ---------- */

function doGet(e) {
  if (e && e.parameter && e.parameter.entrega) return responder(pedirEntrega(String(e.parameter.entrega)));
  const pedido = String((e && e.parameter && e.parameter.codigo) || '');
  if (!RE_CODIGO.test(pedido)) return responder('{"ok":false}');
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName(H_COT).getRange(2, 1, FILAS, COL.estado + 1).getValues()
    .filter(function (r) { return texto(r[COL.codigo]) !== '' && r[COL.estado] === 'Enviada'; })
    .forEach(function (r) { guardarEnCache(texto(r[COL.codigo])); });
  const entregas = ss.getSheetByName(H_ENT);
  if (!entregas) return;
  entregas.getRange(2, 1, FILAS, COL_ENT.estado + 1).getValues()
    .filter(function (r) { return texto(r[COL_ENT.codigo]) !== '' && ['pendiente', 'proceso', 'disponible'].indexOf(ESTADOS_ENT[r[COL_ENT.estado]]) >= 0; })
    .forEach(function (r) { guardarEntrega(texto(r[COL_ENT.codigo])); });
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
    pago: PAGOS[f[COL.pago]] || 'pendiente',
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

// solo el tramo donde están sus renglones, no la pestaña entera
function renglonesDe(hoja, codigo, columnas) {
  const filas = buscarCodigo(hoja, codigo).findAll().map(function (c) { return c.getRow(); });
  if (!filas.length) return [];
  const desde = Math.min.apply(null, filas);
  const hasta = Math.max.apply(null, filas);
  return hoja.getRange(desde, 1, hasta - desde + 1, columnas).getValues();
}

function leerItems(ss, codigo) {
  return renglonesDe(ss.getSheetByName(H_ITEMS), codigo, COL_ITEM.mostrar + 1)
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

/* ---------- entregas ----------
   La ficha guardada tiene todos los renglones, links incluidos, y nunca sale a la web.
   En cada pedido vistaEntrega() arma lo que corresponde mostrar ese día: así un link
   vence a tiempo aunque la ficha siga en la caché, y los que el cliente todavía no
   puede ver nunca llegan a su navegador. */

function claveEntrega(codigo) {
  return 'ent:' + codigo;
}

function pedirEntrega(codigo) {
  if (!RE_CODIGO.test(codigo)) return '{"ok":false}';
  try {
    const guardada = CacheService.getScriptCache().get(claveEntrega(codigo));
    return JSON.stringify(vistaEntrega(JSON.parse(guardada || guardarEntrega(codigo))));
  } catch (err) {
    return '{"ok":false}';
  }
}

function guardarEntrega(codigo) {
  const ficha = leerEntrega(codigo);
  const json = JSON.stringify(ficha);
  if (ficha.ok) CacheService.getScriptCache().put(claveEntrega(codigo), json, CACHE_SEGUNDOS);
  else CacheService.getScriptCache().remove(claveEntrega(codigo));
  return json;
}

function leerEntrega(codigo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const zona = ss.getSpreadsheetTimeZone();
  const hoja = ss.getSheetByName(H_ENT);
  const celda = buscarCodigo(hoja, codigo).findNext();
  if (!celda) return { ok: false };
  const f = hoja.getRange(celda.getRow(), 1, 1, COL_ENT.link + 1).getValues()[0];

  const estado = ESTADOS_ENT[f[COL_ENT.estado]];
  if (!estado) return { ok: false };
  if (estado === 'pausada' || estado === 'anulada') return { ok: true, estado: estado, codigo: codigo };

  const cfg = leerConfig(ss);
  const logo = texto(f[COL_ENT.logoCliente]);
  return {
    ok: true,
    estado: estado,
    codigo: codigo,
    zona: zona,
    cliente: texto(f[COL_ENT.empresa]),
    titulo: texto(f[COL_ENT.titulo]),
    mensaje: texto(f[COL_ENT.mensaje]),
    logoCliente: /^https:\/\//.test(logo) ? logo : '',
    limite: dia(f[COL_ENT.limite], zona),
    archivos: leerArchivos(ss, codigo, zona),
    config: { whatsapp: texto(cfg.whatsapp), mail: texto(cfg.mail), instagram: texto(cfg.instagram) }
  };
}

function leerArchivos(ss, codigo, zona) {
  return renglonesDe(ss.getSheetByName(H_ARC), codigo, COL_ARC.visible + 1)
    .filter(function (r) {
      return String(r[COL_ARC.codigo]) === codigo && texto(r[COL_ARC.nombre]) !== '' && r[COL_ARC.visible] !== 'Oculto';
    })
    .sort(function (a, b) { return numero(a[COL_ARC.orden]) - numero(b[COL_ARC.orden]); })
    .map(function (r) {
      const link = texto(r[COL_ARC.link]);
      return {
        nombre: texto(r[COL_ARC.nombre]),
        tipo: TIPOS_ARC[r[COL_ARC.tipo]] || 'descarga',
        link: /^https:\/\//.test(link) ? link : '',
        detalle: texto(r[COL_ARC.detalle]),
        fecha: dia(r[COL_ARC.fecha], zona),
        // sin elegir, espera al pago: es lo más seguro
        visible: VISIBLES_ARC[r[COL_ARC.visible]] || 'pago'
      };
    });
}

function vistaEntrega(ficha, hoy) {
  if (!ficha.ok) return { ok: false };
  if (ficha.estado === 'pausada' || ficha.estado === 'anulada') return ficha;

  const fecha = hoy || Utilities.formatDate(new Date(), ficha.zona, 'yyyy-MM-dd');
  const dias = ficha.limite ? diasEntre(fecha, ficha.limite) : null;
  const pendiente = ficha.estado === 'pendiente';
  const sinPlazo = pendiente || ficha.estado === 'proceso';
  const vencida = ficha.estado === 'borrada' || (ficha.estado === 'disponible' && dias !== null && dias < 0);

  return {
    ok: true,
    estado: sinPlazo ? ficha.estado : vencida ? 'vencida' : 'disponible',
    codigo: ficha.codigo,
    cliente: ficha.cliente,
    titulo: ficha.titulo,
    mensaje: ficha.mensaje,
    logoCliente: ficha.logoCliente,
    // el plazo corre recién desde que la entrega está disponible
    limite: sinPlazo ? '' : ficha.limite,
    dias: sinPlazo ? null : dias,
    archivos: ficha.archivos
      .filter(function (a) { return !pendiente || a.visible === 'parcial'; })
      .map(function (a) {
        const vencido = vencida && a.tipo === 'descarga';
        // en proceso, lo que espera el pago se lista sin su link
        const bloqueado = ficha.estado === 'proceso' && a.visible === 'pago' && a.tipo !== 'proceso' && a.link !== '';
        return {
          nombre: a.nombre,
          tipo: a.tipo,
          // lo que está en proceso todavía no se abre, aunque el renglón ya tenga link
          link: vencido || bloqueado || a.tipo === 'proceso' ? '' : a.link,
          detalle: a.detalle,
          fecha: a.fecha,
          parcial: a.visible === 'parcial',
          vencido: vencido,
          bloqueado: bloqueado
        };
      }),
    config: ficha.config
  };
}

const diasEntre = function (desde, hasta) { return Math.round((Date.parse(hasta) - Date.parse(desde)) / 86400000); };

/* ---------- avisos por mail ---------- */

// lo ejecuta el activador de "Activar avisos por mail", una vez por día
function avisosEntregas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(H_ENT);
  const zona = ss.getSpreadsheetTimeZone();
  const cfg = leerConfig(ss);
  const anticipo = numero(cfg.avisoEntregaDias) || AVISO_ENTREGA;
  const hoy = Utilities.formatDate(new Date(), zona, 'yyyy-MM-dd');
  const remitente = opcionesMail(cfg);

  hoja.getRange(2, 1, FILAS, COL_ENT.link + 1).getValues().forEach(function (f, i) {
    const limite = dia(f[COL_ENT.limite], zona);
    if (f[COL_ENT.estado] !== 'Disponible' || !limite) return;
    const dias = diasEntre(hoy, limite);
    if (dias < 0) {
      // queda en rojo hasta que borres la carpeta y la pases a Archivos borrados
      hoja.getRange(i + 2, COL_ENT.estado + 1).setValue('Vencida');
      guardarEntrega(texto(f[COL_ENT.codigo]));
      return;
    }
    const mail = texto(f[COL_ENT.mail]);
    const col = dias === 0 ? COL_ENT.avisoUltimo : COL_ENT.avisoPrevio;
    if (!mail || dias > anticipo || f[col] !== '') return;
    try {
      const aviso = textoAviso(f, dias, limite);
      GmailApp.sendEmail(mail, aviso.asunto, aviso.cuerpo, remitente);
      hoja.getRange(i + 2, col + 1).setValue(new Date());
    } catch (err) {
      // un mail mal escrito no frena a los demás; mañana se vuelve a intentar
      console.error('No se pudo avisar a ' + f[COL_ENT.codigo] + ': ' + err.message);
    }
  });
}

// Gmail solo deja usar un remitente ya agregado en "Enviar correo como"
function opcionesMail(cfg) {
  const mail = texto(cfg.mail) || MAIL_FARA;
  const opciones = { name: 'FÁRA', replyTo: mail };
  if (GmailApp.getAliases().indexOf(mail) >= 0) opciones.from = mail;
  return opciones;
}

function textoAviso(f, dias, limite) {
  const saludo = texto(f[COL_ENT.saludo]);
  const titulo = texto(f[COL_ENT.titulo]);
  const fecha = limite.split('-').reverse().join('/');
  const asunto = dias === 0 ? 'Hoy es el último día para descargar tu entrega de FÁRA'
    : dias === 1 ? 'Mañana vence tu entrega de FÁRA'
    : 'Tu entrega de FÁRA vence en ' + dias + ' días';
  return {
    asunto: asunto,
    cuerpo: (saludo ? 'Hola ' + saludo + ',' : 'Hola,') + '\n\n' +
      'Tu entrega' + (titulo ? ' "' + titulo + '"' : '') + ' está disponible hasta el ' + fecha + '.\n' +
      'Descargá todo antes de esa fecha: después los archivos se borran y recuperarlos tiene un costo extra.\n\n' +
      SITIO_ENTREGA + f[COL_ENT.codigo] + '\n\n' +
      'Cualquier duda, respondé este mail.\n\nFÁRA'
  };
}

function activarAvisos() {
  const existe = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'avisosEntregas'; });
  if (!existe) ScriptApp.newTrigger('avisosEntregas').timeBased().everyDays(1).atHour(9).inTimezone('America/Argentina/Buenos_Aires').create();
  SpreadsheetApp.getUi().alert('Listo: los avisos salen solos todos los días a la mañana.');
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
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('FÁRA Cotizaciones')
    .addItem('Nueva cotización', 'nuevaCotizacion')
    .addSeparator()
    .addItem('Ver ítems de esta cotización', 'verItems')
    .addItem('Agregar ítem', 'agregarItem')
    .addItem('Ver todos los ítems', 'verTodosLosItems')
    .addSeparator()
    .addItem('Activar carga rápida', 'activarCargaRapida')
    .addToUi();
  ui.createMenu('FÁRA Entregas')
    .addItem('Nueva entrega', 'nuevaEntrega')
    .addItem('Ver archivos de esta entrega', 'verArchivos')
    .addItem('Agregar archivo', 'agregarArchivo')
    .addItem('Ver todos los archivos', 'verTodosLosArchivos')
    .addSeparator()
    .addItem('Activar avisos por mail', 'activarAvisos')
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
  filtrarPorCodigo(ss.getSheetByName(H_ITEMS), codigo).activate();
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
  filtrarPorCodigo(items, codigo).activate();
  items.getRange(fila, COL_ITEM.servicio + 1).activate();
}

function verTodosLosItems() {
  const items = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H_ITEMS);
  quitarFiltro(items);
  items.activate();
}

function nuevaEntrega() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const r = ui.prompt('Nueva entrega', 'Código de cliente: 3 letras, por ejemplo CEN', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const prefijo = r.getResponseText().trim().toUpperCase();
  if (!/^[A-Z]{3}$/.test(prefijo)) {
    ui.alert('Tienen que ser 3 letras, sin acentos ni números.');
    return;
  }
  const creada = crearEntrega(ss, prefijo);
  ss.getSheetByName(H_ENT).getRange(creada.fila, COL_ENT.empresa + 1).activate();
  ui.alert('Entrega creada en Borrador', 'Completala, cargá sus archivos y pasala a Pendiente de pago o Disponible:\n\n' + SITIO_ENTREGA + creada.codigo, ui.ButtonSet.OK);
}

function verArchivos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const codigo = codigoActivo(ss, H_ENT, H_ARC);
  if (!codigo) {
    SpreadsheetApp.getUi().alert('Pará el cursor en la fila de una entrega y volvé a tocar Ver archivos.');
    return;
  }
  filtrarPorCodigo(ss.getSheetByName(H_ARC), codigo).activate();
}

function agregarArchivo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const codigo = codigoActivo(ss, H_ENT, H_ARC);
  if (!codigo) {
    SpreadsheetApp.getUi().alert('Pará el cursor en la fila de una entrega o abrí sus archivos con Ver archivos.');
    return;
  }
  const archivos = ss.getSheetByName(H_ARC);
  const fila = sumarArchivo(archivos, codigo);
  filtrarPorCodigo(archivos, codigo).activate();
  archivos.getRange(fila, COL_ARC.nombre + 1).activate();
}

function verTodosLosArchivos() {
  const archivos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H_ARC);
  quitarFiltro(archivos);
  archivos.activate();
}

// el código de la fila activa, o el que se está viendo en la pestaña de renglones.
// En todas estas pestañas el código está en la columna A.
function codigoActivo(ss, padre, renglones) {
  const nombrePadre = padre || H_COT;
  const nombreRenglones = renglones || H_ITEMS;
  const hoja = ss.getActiveSheet();
  const fila = ss.getActiveRange().getRow();
  if (hoja.getName() === nombrePadre) return fila > 1 ? texto(hoja.getRange(fila, 1).getValue()) : '';
  if (hoja.getName() !== nombreRenglones) return '';
  const filtro = hoja.getFilter();
  const criterio = filtro && filtro.getColumnFilterCriteria(1);
  if (criterio) return texto(criterio.getCriteriaValues()[0]);
  return fila > 1 ? texto(hoja.getRange(fila, 1).getValue()) : '';
}

function filtrarPorCodigo(hoja, codigo) {
  const filtro = hoja.getFilter() || hoja.getRange(1, 1, FILAS + 1, hoja.getLastColumn()).createFilter();
  filtro.setColumnFilterCriteria(1, SpreadsheetApp.newFilterCriteria().whenTextEqualTo(codigo).build());
  return hoja;
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

function sumarArchivo(archivos, codigo) {
  const orden = archivos.getRange(2, 1, FILAS, 2).getValues()
    .filter(function (r) { return r[0] === codigo; })
    .reduce(function (max, r) { return Math.max(max, numero(r[1])); }, 0);
  const fila = primeraLibre(archivos);
  archivos.getRange(fila, COL_ARC.codigo + 1, 1, 2).setValues([[codigo, orden + 1]]);
  archivos.getRange(fila, COL_ARC.tipo + 1).setValue('Descarga');
  archivos.getRange(fila, COL_ARC.visible + 1).setValue('Al pagar');
  return fila;
}

// items: los renglones ya armados. Sin items quedan tres vacíos, que no se muestran
function crearEntrega(ss, prefijo, datos, items) {
  const hoja = ss.getSheetByName(H_ENT);
  const archivos = ss.getSheetByName(H_ARC);
  quitarFiltro(archivos);
  const codigo = codigoNuevo(hoja, prefijo);
  const fila = primeraLibre(hoja);
  hoja.getRange(fila, COL_ENT.codigo + 1, 1, 2).setValues([[codigo, prefijo]]);
  hoja.getRange(fila, COL_ENT.estado + 1).setValue(datos ? 'En proceso' : 'Borrador');
  if (datos) {
    hoja.getRange(fila, COL_ENT.empresa + 1, 1, 2).setValues([[datos.empresa, datos.saludo]]);
    hoja.getRange(fila, COL_ENT.titulo + 1).setValue(datos.titulo);
    hoja.getRange(fila, COL_ENT.logoCliente + 1).setValue(datos.logoCliente);
    hoja.getRange(fila, COL_ENT.cotizacion + 1).setValue(datos.cotizacion);
  }
  (items || [{}, {}, {}]).forEach(function (it) {
    const renglon = sumarArchivo(archivos, codigo);
    if (!it.nombre) return;
    archivos.getRange(renglon, COL_ARC.nombre + 1, 1, 2).setValues([[it.nombre, 'En proceso']]);
    archivos.getRange(renglon, COL_ARC.detalle + 1).setValue(it.detalle);
    archivos.getRange(renglon, COL_ARC.visible + 1).setValue(it.visible);
  });
  return { codigo: codigo, fila: fila };
}

// al aceptar o marcar un pago en la planilla: nace su entrega, una sola vez, con los mismos ítems
function entregaDeCotizacion(ss, hoja, fila) {
  const f = hoja.getRange(fila, 1, 1, COL.link + 1).getValues()[0];
  const cotizacion = texto(f[COL.codigo]);
  const pagada = f[COL.pago] === 'Pagado total';
  const conPago = pagada || f[COL.pago] === 'Anticipo pagado';
  if (!cotizacion || (f[COL.estado] !== 'Aceptada' && !conPago)) return;

  const entregas = ss.getSheetByName(H_ENT);
  const ya = entregas.getRange(2, COL_ENT.cotizacion + 1, FILAS, 1).createTextFinder(cotizacion)
    .matchCase(true).matchEntireCell(true).findNext();
  if (ya) {
    if (pagada) habilitarArchivos(ss, texto(entregas.getRange(ya.getRow(), COL_ENT.codigo + 1).getValue()));
    return;
  }
  const visible = pagada ? 'Entrega parcial' : 'Al pagar';
  const items = leerItems(ss, cotizacion).map(function (it) {
    return { nombre: it.servicio, detalle: it.entregables || it.descripcion, visible: visible };
  });
  const creada = crearEntrega(ss, cotizacion.slice(0, 3), {
    empresa: texto(f[COL.cliente]),
    saludo: texto(f[COL.contacto]),
    titulo: texto(f[COL.titulo]),
    logoCliente: texto(f[COL.logoCliente]),
    cotizacion: cotizacion
  }, items);
  guardarEntrega(creada.codigo);
  ss.toast('Se creó su entrega, en la pestaña Entregas: ' + SITIO_ENTREGA + creada.codigo, 'FÁRA', 10);
}

// con todo pagado ya no hay nada que espere: lo cargado se puede ver a medida que esté
function habilitarArchivos(ss, codigo) {
  const archivos = ss.getSheetByName(H_ARC);
  archivos.getRange(2, 1, FILAS, COL_ARC.visible + 1).getValues().forEach(function (r, i) {
    if (r[COL_ARC.codigo] === codigo && r[COL_ARC.visible] === 'Al pagar') {
      archivos.getRange(i + 2, COL_ARC.visible + 1).setValue('Entrega parcial');
    }
  });
  guardarEntrega(codigo);
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
  hoja.getRange(fila, COL.pago + 1).setValue('Pendiente');
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
    const toca = function (col) { return e.range.getColumn() <= col + 1 && e.range.getLastColumn() >= col + 1; };
    if (toca(COL.estado) || toca(COL.pago)) {
      for (let f = e.range.getRow(); f <= e.range.getLastRow(); f++) entregaDeCotizacion(e.source, hoja, f);
    }
    return;
  }
  if (hoja.getName() === H_ENT) {
    ordenarPlazo(hoja, e.range);
    refrescarFilas(hoja, e.range, guardarEntrega);
    return;
  }
  if (hoja.getName() === H_ARC) {
    refrescarFilas(hoja, e.range, guardarEntrega);
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

// el código siempre está en la columna A, en todas las pestañas
function refrescarFilas(hoja, rango, guardar) {
  const codigos = hoja.getRange(rango.getRow(), 1, rango.getNumRows(), 1).getValues()
    .map(function (r) { return texto(r[0]); })
    .filter(function (c, i, todos) { return c !== '' && todos.indexOf(c) === i; });
  codigos.forEach(function (c) { (guardar || guardarEnCache)(c); });
}

// el plazo arranca al pasar a Disponible, y una fecha nueva vuelve a habilitar los avisos
function ordenarPlazo(hoja, rango) {
  const toca = function (col) { return rango.getColumn() <= col + 1 && rango.getLastColumn() >= col + 1; };
  if (!toca(COL_ENT.estado) && !toca(COL_ENT.limite)) return;
  const dias = numero(leerConfig(hoja.getParent()).plazoEntregaDias) || PLAZO_ENTREGA;
  for (let fila = rango.getRow(); fila <= rango.getLastRow(); fila++) {
    if (toca(COL_ENT.limite)) hoja.getRange(fila, COL_ENT.avisoPrevio + 1, 1, 2).clearContent();
    const estado = hoja.getRange(fila, COL_ENT.estado + 1).getValue();
    const limite = hoja.getRange(fila, COL_ENT.limite + 1);
    if (toca(COL_ENT.estado) && estado === 'Disponible' && limite.getValue() === '') {
      const vence = new Date();
      vence.setDate(vence.getDate() + dias);
      limite.setValue(vence);
    }
  }
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
const C_AMBAR = '#C8873A'; // pestañas y encabezados de cotizaciones
const C_PETROLEO = '#2F5D62'; // pestañas y encabezados de entregas
const C_GRIS = '#8C8378'; // lo general
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
  sumarPestanasEntregas(ss);
  ss.deleteSheet(inicial);
  ordenarPestanas(ss);

  cargarEjemplo(ss);
  ss.getSheetByName(H_USO).activate();
  ss.toast('Planilla lista. La cotización de ejemplo está en la pestaña Cotizaciones.', 'FÁRA', 10);
}

// cada sección con su color, de izquierda a derecha
function ordenarPestanas(ss) {
  [[H_USO, C_GRIS], [H_COT, C_AMBAR], [H_ITEMS, C_AMBAR], [H_CAT, C_AMBAR], [H_ENT, C_PETROLEO], [H_ARC, C_PETROLEO], [H_CONFIG, C_GRIS]]
    .forEach(function (p, i) {
      const hoja = ss.getSheetByName(p[0]).setTabColor(p[1]);
      ss.setActiveSheet(hoja);
      ss.moveActiveSheet(i + 1);
    });
}

// encabezado: el color de la sección se puede tocar, rojo es automático
function encabezado(hoja, titulos, automaticas, color) {
  hoja.getRange(1, 1, 1, titulos.length).setValues([titulos])
    .setFontWeight('bold').setFontColor(C_CREMA).setBackground(color || C_AMBAR)
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
function bandas(hoja, columnas, formula, reglasPrevias) {
  hoja.setConditionalFormatRules((reglasPrevias || []).concat([SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula).setBackground(C_BANDA)
    .setRanges([hoja.getRange(2, 1, FILAS, columnas)]).build()]));
}

function automatica(hoja, col) {
  hoja.getRange(2, col + 1, FILAS, 1).setBackground(C_CREMA2).protect()
    .setDescription('Columna automática').setWarningOnly(true);
}

function armarCotizaciones(ss) {
  const h = ss.insertSheet(H_COT);
  encabezado(h, ['Código', 'Código de cliente', 'Empresa', 'Nombre para el mail', 'Título de la propuesta', 'Mensaje de introducción',
    'Fecha de emisión', 'Vencimiento', 'Estado', 'Moneda', 'Tipo de descuento', 'Descuento', 'Motivo del descuento',
    'Forma de pago', 'Anticipo %', 'Pago', 'Plazo de entrega', 'Condiciones', 'Notas', 'Disclaimer', 'Logo del cliente (link)',
    'Total', 'Última modificación', 'Link'], [COL.codigo, COL.total, COL.actualizada, COL.link]);

  lista(h, 2, COL.estado, FILAS, ['Borrador', 'Enviada', 'Aceptada', 'Vencida', 'Anulada', 'Pausada']);
  lista(h, 2, COL.moneda, FILAS, ['ARS', 'USD']);
  lista(h, 2, COL.descuentoTipo, FILAS, ['%', 'Monto']);
  lista(h, 2, COL.pago, FILAS, Object.keys(PAGOS));
  h.getRange(2, COL.emision + 1, FILAS, 2).setNumberFormat('dd/mm/yyyy');
  h.getRange(2, COL.actualizada + 1, FILAS, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  h.getRange(2, COL.descuentoValor + 1, FILAS, 1).setNumberFormat('#,##0.##');
  h.getRange(2, COL.anticipoPct + 1, FILAS, 1).setNumberFormat('0"%"');
  h.getRange(2, COL.total + 1, FILAS, 1).setNumberFormat('#,##0');
  h.getRange(2, COL.codigo + 1, FILAS, 1).setNumberFormat('@');

  // con la planilla en es_AR las fórmulas separan con punto y coma
  const neto = "SUMIF('" + H_ITEMS + "'!A2:A;A2:A;'" + H_ITEMS + "'!K2:K)";
  // las fórmulas van en el encabezado: así sobreviven si se borra la primera fila de datos
  h.getRange(1, COL.total + 1).setFormula(
    '={"Total";ARRAYFORMULA(IF(A2:A="";"";IF(K2:K="%";' + neto + '*(1-IF(L2:L<1;L2:L;L2:L/100));' + neto + '-L2:L)))}');
  h.getRange(1, COL.link + 1).setFormula('={"Link";ARRAYFORMULA(IF(A2:A="";"";"' + SITIO + '"&A2:A))}');

  [COL.codigo, COL.total, COL.actualizada, COL.link].forEach(function (c) { automatica(h, c); });
  bandas(h, COL.link + 1, '=AND($A2<>"";ISEVEN(ROW()))');
  h.setColumnWidths(1, COL.link + 1, 130);
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
  h.getRange(1, COL_ITEM.importe + 1).setFormula(
    '={"Importe";ARRAYFORMULA(IF(A2:A="";"";IF(J2:J=TRUE;H2:H*(1-IF(I2:I<1;I2:I;I2:I/100));0)))}');

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

// suma las entregas a una planilla ya armada, sin tocar lo que hay
function agregarEntregas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(H_ENT)) {
    SpreadsheetApp.getUi().alert('Las pestañas de entregas ya están.');
    return;
  }
  sumarPestanasEntregas(ss);
  ss.getSheetByName(H_ENT).activate();
  SpreadsheetApp.getUi().alert('Listo: ya están las pestañas Entregas y Archivos.');
}

function sumarPestanasEntregas(ss) {
  armarEntregas(ss);
  armarArchivos(ss);
  sumarConfigEntregas(ss);
  sumarUsoEntregas(ss);
}

function armarEntregas(ss) {
  const h = ss.insertSheet(H_ENT);
  const automaticas = [COL_ENT.codigo, COL_ENT.avisoPrevio, COL_ENT.avisoUltimo, COL_ENT.link, COL_ENT.cotizacion];
  encabezado(h, ['Código', 'Código de cliente', 'Empresa', 'Nombre para el mail', 'Mail del cliente', 'Título de la entrega', 'Mensaje',
    'Logo del cliente (link)', 'Fecha límite', 'Estado', 'Aviso previo enviado', 'Aviso último día enviado', 'Link', 'Cotización'],
    automaticas, C_PETROLEO);

  lista(h, 2, COL_ENT.estado, FILAS, ['Borrador'].concat(Object.keys(ESTADOS_ENT)));
  h.getRange(2, COL_ENT.mail + 1, FILAS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireTextIsEmail().setAllowInvalid(false).build());
  h.getRange(2, COL_ENT.limite + 1, FILAS, 1).setNumberFormat('dd/mm/yyyy').setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  h.getRange(2, COL_ENT.avisoPrevio + 1, FILAS, 2).setNumberFormat('dd/mm/yyyy');
  h.getRange(2, COL_ENT.codigo + 1, FILAS, 1).setNumberFormat('@');
  h.getRange(1, COL_ENT.link + 1).setFormula('={"Link";ARRAYFORMULA(IF(A2:A="";"";"' + SITIO_ENTREGA + '"&A2:A))}');

  automaticas.forEach(function (c) { automatica(h, c); });
  // en rojo: venció y falta borrar o cerrar la carpeta. Al pasar a Archivos borrados se apaga
  const vencida = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($A2<>"";OR($J2="Vencida";AND($J2="Disponible";$I2<>"";$I2<TODAY())))')
    .setBackground(C_ROJO).setFontColor(C_CREMA)
    .setRanges([h.getRange(2, 1, FILAS, COL_ENT.link + 1)]).build();
  bandas(h, COL_ENT.link + 1, '=AND($A2<>"";ISEVEN(ROW()))', [vencida]);
  h.setColumnWidths(1, COL_ENT.link + 1, 130);
  h.setColumnWidth(COL_ENT.saludo + 1, 170);
  h.setColumnWidth(COL_ENT.mail + 1, 220);
  h.setColumnWidth(COL_ENT.titulo + 1, 240);
  h.setColumnWidth(COL_ENT.mensaje + 1, 280);
  h.setColumnWidth(COL_ENT.link + 1, 360);
  h.setFrozenColumns(3);
}

function armarArchivos(ss) {
  const h = ss.insertSheet(H_ARC);
  encabezado(h, ['Código', 'Orden', 'Qué entregás', 'Tipo', 'Link', 'Detalle', 'Fecha', 'Cuándo se ve'], [], C_PETROLEO);

  lista(h, 2, COL_ARC.tipo, FILAS, Object.keys(TIPOS_ARC));
  lista(h, 2, COL_ARC.visible, FILAS, Object.keys(VISIBLES_ARC));
  h.getRange(2, COL_ARC.fecha + 1, FILAS, 1).setNumberFormat('dd/mm/yyyy');
  h.getRange(2, COL_ARC.codigo + 1, FILAS, 1).setNumberFormat('@');

  // cada entrega toma su color según su fila en Entregas
  bandas(h, COL_ARC.visible + 1, '=AND($A2<>"";ISEVEN(MATCH($A2;INDIRECT("' + H_ENT + '!$A:$A");0)))');
  h.setColumnWidths(1, COL_ARC.visible + 1, 110);
  h.setColumnWidth(COL_ARC.nombre + 1, 240);
  h.setColumnWidth(COL_ARC.link + 1, 320);
  h.setColumnWidth(COL_ARC.detalle + 1, 280);
}

function sumarConfigEntregas(ss) {
  const h = ss.getSheetByName(H_CONFIG);
  const filas = [
    ['Entregas: días para descargar', PLAZO_ENTREGA, 'plazoEntregaDias'],
    ['Entregas: días antes para avisar por mail', AVISO_ENTREGA, 'avisoEntregaDias']
  ];
  const desde = h.getLastRow() + 1;
  h.getRange(desde, 1, filas.length, 3).setValues(filas).setFontFamily('Archivo');
  h.getRange(desde, 3, filas.length, 1).setBackground(C_CREMA2).protect().setDescription('No cambiar').setWarningOnly(true);
}

function sumarUsoEntregas(ss) {
  const h = ss.getSheetByName(H_USO);
  const lineas = [
    [''],
    ['Crear una entrega'],
    ['Se crea sola al aceptar una cotización. Para una entrega suelta:'],
    ['1. Menú FÁRA Entregas → Nueva entrega. Escribí las 3 letras del cliente.'],
    ['2. En Entregas, completá la fila. Empresa sale en la página; Nombre para el mail es el saludo: Hola Leandro. Con mail, le llegan los avisos de vencimiento.'],
    ['3. Parado en esa fila, tocá FÁRA Entregas → Ver archivos y cargá un renglón por cosa que entregás.'],
    ['Tipo: Descarga vence. Web no vence. Entregado es sin link, como una impresión o una charla, y queda como registro. En proceso es lo que todavía falta: cuando esté, cambiá el Tipo y pegá el link.'],
    ['Cuándo se ve: Al pagar, Entrega parcial (se ve antes del pago) u Oculto.'],
    ['4. Elegí el Estado y mandale el link de la última columna.'],
    [''],
    ['Estados de una entrega'],
    ['Borrador: el link dice que no está disponible.'],
    ['Pendiente de pago: avisa que falta el pago y muestra solo las entregas parciales. El plazo no corre.'],
    ['En proceso: el cliente ve la lista mientras la vas terminando. Lo marcado Al pagar sale sin link hasta que lo pague. El plazo no corre y no salen mails.'],
    ['Disponible: se ve todo. Si la Fecha límite está vacía, se completa sola con los días de Config.'],
    ['Vencida: pasa sola el día después de la Fecha límite y la fila se pone en rojo. Borrá o cerrá la carpeta y pasala a Archivos borrados.'],
    ['Pausada y Anulada: un mensaje, sin datos ni links.']
  ];
  const desde = h.getLastRow() + 1;
  h.getRange(desde, 1, lineas.length, 1).setValues(lineas).setFontFamily('Archivo').setBackground(C_CREMA).setFontColor(C_ESPRESSO).setWrap(true);
  [1, 10].forEach(function (i) { h.getRange(desde + i, 1).setFontWeight('bold').setFontSize(13); });
}

function armarUso(ss) {
  const h = ss.insertSheet(H_USO);
  const lineas = [
    ['FÁRA · Cotizaciones'],
    ['Esta planilla es privada. El cliente solo ve su cotización en la web, con su link.'],
    [''],
    ['Crear una cotización'],
    ['1. Menú FÁRA Cotizaciones → Nueva cotización. Escribí las 3 letras del cliente.'],
    ['2. En Cotizaciones, completá la fila nueva.'],
    ['3. Parado en esa fila, tocá FÁRA Cotizaciones → Ver ítems y completá los renglones. Para sumar otro: Agregar ítem. En Precio va el total del renglón: la cantidad no multiplica.'],
    ['4. Cambiá el Estado a Enviada y mandale el link de la última columna.'],
    [''],
    ['Cambiar una cotización'],
    ['Editá la celda. El cliente ve el cambio al recargar la página. Para volver a la lista completa: FÁRA Cotizaciones → Ver todos los ítems.'],
    [''],
    ['Estados'],
    ['Borrador: el link dice que no está disponible.'],
    ['Enviada: se ve completa. Si pasa la fecha de vencimiento, se muestra vencida sola.'],
    ['Aceptada: se ve completa con el aviso de aceptada. Al ponerla vos en Aceptada, o marcar un pago, se crea sola su entrega con los mismos ítems.'],
    ['Pago: elegí Anticipo pagado o Pagado total a medida que cobrás. Si te paga otro anticipo, cambiá primero Anticipo %. El link sigue activo como registro.'],
    ['Pausada y Anulada: un mensaje, sin datos ni precios.'],
    ['Vencida: se ve completa con el botón para pedir actualización.'],
    [''],
    ['Qué se puede tocar'],
    ['Encabezado de color: se puede editar.'],
    ['Encabezado rojo: automático, no lo toques.'],
    ['No borres cotizaciones: quedan como historial. Para sacarlas de la web, usá Anulada.']
  ];
  h.getRange(1, 1, lineas.length, 1).setValues(lineas).setFontFamily('Archivo').setBackground(C_CREMA).setFontColor(C_ESPRESSO).setWrap(true);
  h.setColumnWidth(1, 720);
  h.getRange(1, 1).setFontSize(20).setFontWeight('bold');
  [4, 10, 13, 21].forEach(function (f) { h.getRange(f, 1).setFontWeight('bold').setFontSize(13); });
  h.getRange(22, 1).setBackground(C_AMBAR).setFontColor(C_CREMA);
  h.getRange(23, 1).setBackground(C_ROJO).setFontColor(C_CREMA);
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
  hoja.getRange(f, COL.descuentoTipo + 1, 1, 5).setValues([['%', 10, 'Descuento por contratar los tres servicios juntos.',
    'Transferencia bancaria.', 50]]);
  hoja.getRange(f, COL.plazoEntrega + 1).setValue('15 días hábiles desde la sesión de fotos.');

  const items = ss.getSheetByName(H_ITEMS);
  items.getRange(2, COL_ITEM.servicio + 1, 3, 7).setValues([
    ['Fotos de producto', 'Sesión en estudio con fondo y luz definidos para tu marca.', '12 fotos editadas\nFormato web y redes', 12, 'fotos', 300000, 10],
    ['Landing', 'Una página con tu oferta y botón de WhatsApp.', '', 1, '', 350000, 0],
    ['Menú digital', 'Carta online que editás desde una planilla.', '', 1, '', 150000, 0]
  ]);
}
