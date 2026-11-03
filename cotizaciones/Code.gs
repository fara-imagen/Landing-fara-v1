/**
 * FÁRA · Cotizaciones y entregas: puente entre la planilla y
 * faraimagen.com.ar/cotizacion/<codigo> y faraimagen.com.ar/entrega/<codigo>.
 *
 * doGet() ......... recibe un código y devuelve solo esa cotización o entrega.
 * setup() ......... arma la planilla entera. Se ejecuta una sola vez.
 * Menús FÁRA ...... "Nueva cotización" y "Nueva entrega" generan el código y dejan la fila lista.
 * onEdit() ........ completa datos del Catálogo, la última modificación y la fecha límite,
 *                   y crea la entrega de una cotización al aceptarla o marcar un pago.
 * avisosEntregas() manda los mails de vencimiento y marca las vencidas. Corre solo una vez por día.
 * doPost() ........ atiende a faraimagen.com.ar/panel-cotizaciones: sin la clave no lee ni cambia nada.
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
const LOGO_MAIL = 'https://res.cloudinary.com/eekglpse/image/upload/f_png,h_160/c_crop,g_center,w_0.96,h_0.9/e_make_transparent:35/e_trim:10/e_colorize:100,co_rgb:1A1410/v1792547783/fara_marca_agua_logotipo_crema.svg';
const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

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
  logoCliente: 20, total: 21, actualizada: 22, link: 23, mail: 24
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
  const f = hoja.getRange(celda.getRow(), 1, 1, COL_ENT.cotizacion + 1).getValues()[0];

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
    cotizacion: texto(f[COL_ENT.cotizacion]),
    archivos: leerArchivos(ss, codigo, zona),
    config: { whatsapp: texto(cfg.whatsapp), mail: texto(cfg.mail), instagram: texto(cfg.instagram), avisoDias: diasDeAviso(cfg) }
  };
}

function leerArchivos(ss, codigo, zona) {
  return renglonesDe(ss.getSheetByName(H_ARC), codigo, COL_ARC.visible + 1)
    .filter(function (r) {
      return String(r[COL_ARC.codigo]) === codigo && texto(r[COL_ARC.nombre]) !== '' && r[COL_ARC.visible] !== 'Oculto';
    })
    .sort(function (a, b) { return numero(a[COL_ARC.orden]) - numero(b[COL_ARC.orden]); })
    .map(function (r) {
      return {
        nombre: texto(r[COL_ARC.nombre]),
        tipo: TIPOS_ARC[r[COL_ARC.tipo]] || 'descarga',
        link: linkValido(r[COL_ARC.link]),
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
  // sin nada para descargar no hay plazo: una web o un trabajo entregado no vencen
  const sinPlazo = pendiente || ficha.estado === 'proceso' || (ficha.estado === 'disponible' && !ficha.archivos.some(vence));
  const vencida = ficha.estado === 'borrada' || (ficha.estado === 'disponible' && !sinPlazo && dias !== null && dias < 0);

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
      .filter(function (a) { return (!pendiente || a.visible === 'parcial') && !(vencida && a.tipo === 'proceso'); })
      .map(function (a) {
        const vencido = vencida && a.tipo === 'descarga';
        // en proceso, lo que espera el pago de su cotización se lista sin su link
        const esperaPago = ficha.estado === 'proceso' && a.visible === 'pago' && !!ficha.cotizacion;
        const bloqueado = esperaPago && a.tipo !== 'proceso' && a.link !== '';
        return {
          nombre: a.nombre,
          tipo: a.tipo,
          // lo que está en proceso todavía no se abre, aunque el renglón ya tenga link
          link: vencido || bloqueado || a.tipo === 'proceso' ? '' : a.link,
          detalle: a.detalle,
          fecha: a.fecha,
          parcial: a.visible === 'parcial',
          vencido: vencido,
          esperaPago: esperaPago,
          bloqueado: bloqueado
        };
      }),
    config: ficha.config
  };
}

const diasEntre = function (desde, hasta) { return Math.round((Date.parse(hasta) - Date.parse(desde)) / 86400000); };
const vence = function (a) { return a.tipo === 'proceso' || (a.tipo === 'descarga' && a.link !== ''); };
// 0 en Config apaga los avisos; vacío usa el valor de siempre
const diasDeAviso = function (cfg) { return texto(cfg.avisoEntregaDias) === '' ? AVISO_ENTREGA : numero(cfg.avisoEntregaDias); };
// acepta el link como se copie del navegador, con o sin https://. Nada que no sea una dirección web
const linkValido = function (v) {
  const link = texto(v);
  if (/^https?:\/\//i.test(link)) return link;
  return /^[a-z0-9-]+(\.[a-z0-9-]+)+([/?#]|$)/i.test(link) ? 'https://' + link : '';
};

/* ---------- avisos por mail ---------- */

// lo ejecuta el activador de "Activar avisos por mail", una vez por día
function avisosEntregas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(H_ENT);
  const zona = ss.getSpreadsheetTimeZone();
  const cfg = leerConfig(ss);
  const anticipo = diasDeAviso(cfg);
  const hoy = Utilities.formatDate(new Date(), zona, 'yyyy-MM-dd');
  const remitente = opcionesMail(cfg);

  hoja.getRange(2, 1, FILAS, COL_ENT.link + 1).getValues().forEach(function (f, i) {
    const limite = dia(f[COL_ENT.limite], zona);
    const codigo = texto(f[COL_ENT.codigo]);
    if (f[COL_ENT.estado] !== 'Disponible' || !limite || !leerArchivos(ss, codigo, zona).some(vence)) return;
    const dias = diasEntre(hoy, limite);
    if (dias < 0) {
      // queda en rojo hasta que borres la carpeta y la pases a Archivos borrados
      hoja.getRange(i + 2, COL_ENT.estado + 1).setValue('Vencida');
      guardarEntrega(codigo);
      return;
    }
    const mail = texto(f[COL_ENT.mail]);
    const col = dias === 0 ? COL_ENT.avisoUltimo : COL_ENT.avisoPrevio;
    if (!mail || !anticipo || dias > anticipo || f[col] !== '') return;
    try {
      const aviso = textoAviso(f, dias, limite);
      GmailApp.sendEmail(mail, aviso.asunto, aviso.cuerpo, Object.assign({ htmlBody: aviso.html }, remitente));
      hoja.getRange(i + 2, col + 1).setValue(new Date());
    } catch (err) {
      // un mail mal escrito no frena a los demás; mañana se vuelve a intentar
      console.error('No se pudo avisar a ' + codigo + ': ' + err.message);
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
  const fecha = fechaEnLetras(limite);
  const link = SITIO_ENTREGA + f[COL_ENT.codigo];
  const asunto = dias === 0 ? 'Hoy es el último día para descargar tu entrega de FÁRA'
    : dias === 1 ? 'Mañana vence tu entrega de FÁRA'
    : 'Tu entrega de FÁRA vence en ' + dias + ' días';
  const encabezado = dias === 0 ? 'Hoy es el último día para descargar'
    : dias === 1 ? 'Tu entrega vence mañana'
    : 'Tu entrega vence en ' + dias + ' días';
  const hola = saludo ? 'Hola ' + saludo + ',' : 'Hola,';
  const disponible = 'Tu entrega' + (titulo ? ' "' + titulo + '"' : '') + ' está disponible hasta el ' + fecha + '.';
  const borrado = 'Descargá todo a tiempo: después de esa fecha los archivos se borran.';
  const costo = 'Si necesitás recuperarlos después de esa fecha, tiene un costo extra.';
  return {
    asunto: asunto,
    cuerpo: hola + '\n\n' + disponible + '\n' + borrado + ' ' + costo + '\n\n' + link + '\n\nCualquier duda, respondé este mail.\n\nFÁRA',
    html: mailHtml(encabezado, hola, disponible, borrado, costo, link)
  };
}

// los mails solo aceptan estilos escritos en cada elemento y tablas para ordenar
function mailHtml(encabezado, hola, disponible, borrado, costo, link) {
  const letra = 'font-family:Arial,Helvetica,sans-serif;';
  const parrafo = 'margin:0 0 14px;' + letra + 'font-size:16px;line-height:1.5;color:' + C_ESPRESSO + ';';
  return '<body style="margin:0;padding:0;background:' + C_CREMA + ';">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + C_CREMA + ';"><tr><td align="center" style="padding:24px 12px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:480px;background:#FFFFFF;border-radius:16px;"><tr><td style="padding:28px 28px 24px;">' +
    '<img src="' + LOGO_MAIL + '" alt="FÁRA" height="24" style="display:block;height:24px;width:auto;border:0;margin:0 0 24px;">' +
    '<h1 style="margin:0 0 16px;' + letra + 'font-size:24px;line-height:1.2;font-weight:bold;color:' + C_ESPRESSO + ';">' + escapar(encabezado) + '</h1>' +
    '<p style="' + parrafo + '">' + escapar(hola) + '</p>' +
    '<p style="' + parrafo + '">' + escapar(disponible) + ' ' + escapar(borrado) + '</p>' +
    '<table role="presentation" cellpadding="0" cellspacing="0" style="margin:22px 0;"><tr><td style="background:' + C_AMBAR + ';border-radius:10px;">' +
    '<a href="' + link + '" style="display:inline-block;padding:14px 28px;' + letra + 'font-size:16px;font-weight:bold;color:' + C_ESPRESSO + ';text-decoration:none;">Ver mi entrega</a>' +
    '</td></tr></table>' +
    '<p style="margin:0;' + letra + 'font-size:13px;line-height:1.5;font-style:italic;color:' + C_MARRON + ';">' + escapar(costo) + '</p>' +
    '</td></tr><tr><td style="padding:16px 28px 22px;border-top:1px solid ' + C_CREMA2 + ';' + letra + 'font-size:13px;line-height:1.5;color:' + C_MARRON + ';">' +
    'Cualquier duda, respondé este mail.' +
    '</td></tr></table></td></tr></table></body>';
}

const escapar = function (v) {
  return String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
};
const fechaEnLetras = function (txt) {
  const p = txt.split('-');
  return Number(p[2]) + ' de ' + MESES[Number(p[1]) - 1] + ' de ' + p[0];
};

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
    .addItem('Crear clave del panel', 'crearClavePanel')
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
  const propios = archivos.getRange(2, 1, FILAS, COL_ARC.visible + 1).getValues()
    .filter(function (r) { return r[0] === codigo; });
  const orden = propios.reduce(function (max, r) { return Math.max(max, numero(r[COL_ARC.orden])); }, 0);
  // si la entrega ya está abierta, lo nuevo también
  const anterior = propios.length ? propios[propios.length - 1][COL_ARC.visible] : '';
  const fila = primeraLibre(archivos);
  archivos.getRange(fila, COL_ARC.codigo + 1, 1, 2).setValues([[codigo, orden + 1]]);
  archivos.getRange(fila, COL_ARC.tipo + 1).setValue('Descarga');
  archivos.getRange(fila, COL_ARC.visible + 1).setValue(anterior === 'Entrega parcial' ? anterior : 'Al pagar');
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
    hoja.getRange(fila, COL_ENT.empresa + 1, 1, 4).setValues([[datos.empresa, datos.saludo, datos.mail, datos.titulo]]);
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
  const f = hoja.getRange(fila, 1, 1, COL.mail + 1).getValues()[0];
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
    mail: texto(f[COL.mail]),
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
  try {
    hoja.getRange(fila, COL.cliente + 1).activate();
  } catch (err) {
    // desde el panel no hay planilla abierta que mover
  }
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

/* ---------- panel para el celular ----------
   faraimagen.com.ar/panel-cotizaciones le habla a doPost(), y cada pedido trae la clave.
   Acá solo queda su huella: la clave no se puede leer ni desde el script.
   Escribir por script no dispara onEdit, así que cada acción hace a mano lo mismo que él. */

const PROP_CLAVE = 'panelClave';
const PROP_CARPETA = 'panelCarpetaFotos';
const CARPETA_FOTOS = 'FÁRA · Fotos de entregas';
const MAX_FOTO = 10000000; // caracteres de la foto ya comprimida: unos 7 MB
const INTENTOS_PANEL = 10;
const BLOQUEO_PANEL = 900; // segundos sin atender después de tantas claves mal
const ESTADOS_COT = ['Borrador', 'Enviada', 'Aceptada', 'Vencida', 'Anulada', 'Pausada'];
const MAX_TEXTO = 500;
const MAX_ITEMS = 30;

function doPost(e) {
  let pedido;
  try {
    pedido = JSON.parse(e.postData.contents);
  } catch (err) {
    return responder('{"ok":false}');
  }
  const cache = CacheService.getScriptCache();
  const fallos = Number(cache.get('panel:fallos')) || 0;
  if (fallos >= INTENTOS_PANEL) return responder('{"ok":false,"bloqueado":true}');
  if (!claveCorrecta(pedido.clave)) {
    cache.put('panel:fallos', String(fallos + 1), BLOQUEO_PANEL);
    return responder('{"ok":false,"clave":true}');
  }
  if (fallos) cache.remove('panel:fallos');
  if (!Object.prototype.hasOwnProperty.call(ACCIONES_PANEL, pedido.accion)) return responder('{"ok":false}');
  try {
    const args = Array.isArray(pedido.args) ? pedido.args : [];
    return responder(JSON.stringify({ ok: true, datos: ACCIONES_PANEL[pedido.accion].apply(null, args) }));
  } catch (err) {
    return responder(JSON.stringify({ ok: false, error: err.message }));
  }
}

function huella(clave) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, clave, Utilities.Charset.UTF_8));
}

function claveCorrecta(clave) {
  const guardada = PropertiesService.getScriptProperties().getProperty(PROP_CLAVE);
  return !!guardada && typeof clave === 'string' && huella(clave) === guardada;
}

function crearClavePanel() {
  const ui = SpreadsheetApp.getUi();
  const propiedades = PropertiesService.getScriptProperties();
  if (propiedades.getProperty(PROP_CLAVE)) {
    const seguir = ui.alert('Crear clave del panel', 'Ya hay una clave. Si creás otra, la anterior deja de servir y hay que volver a entrar en cada celular. ¿Seguís?', ui.ButtonSet.YES_NO);
    if (seguir !== ui.Button.YES) return;
  }
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid());
  let clave = '';
  for (let i = 0; i < 16; i++) clave += (i && i % 4 === 0 ? '-' : '') + LETRAS[(bytes[i] & 255) % LETRAS.length];
  propiedades.setProperty(PROP_CLAVE, huella(clave));
  ui.alert('Clave del panel', 'Copiala ahora: no se vuelve a mostrar.\n\n' + clave + '\n\nEntrá a faraimagen.com.ar/panel-cotizaciones desde tu celular y pegala una sola vez.', ui.ButtonSet.OK);
}

// lo que se escribe en la planilla: una celda que empieza con = se tomaría como fórmula
const dato = function (v, max) {
  const t = texto(v).slice(0, max || MAX_TEXTO);
  return /^[=+\-@]/.test(t) ? ' ' + t : t;
};
const elegir = function (v, opciones) {
  if (opciones.indexOf(v) < 0) throw new Error('Ese valor no es válido.');
  return v;
};
const importe = function (v) {
  const n = Number(v);
  if (v === '' || v === null || v === undefined || !isFinite(n) || n < 0) throw new Error('Un número no es válido.');
  return n;
};
const mailDelPanel = function (v) {
  const mail = texto(v).slice(0, 200);
  if (mail && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(mail)) throw new Error('Ese mail no parece válido.');
  return mail;
};
const pedirCodigo = function (v) {
  const codigo = texto(v);
  if (!RE_CODIGO.test(codigo)) throw new Error('Código no válido.');
  return codigo;
};

function filaDe(hoja, codigo) {
  const celda = buscarCodigo(hoja, codigo).findNext();
  if (!celda) throw new Error('No encontré ' + codigo + '.');
  return celda.getRow();
}

function leerCatalogo(ss) {
  return ss.getSheetByName(H_CAT).getRange(2, 1, FILAS, 3).getValues()
    .filter(function (r) { return texto(r[0]) !== ''; })
    .map(function (r) { return { servicio: texto(r[0]), descripcion: texto(r[1]), precio: r[2] === '' ? '' : numero(r[2]) }; });
}

function panelDatos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const zona = ss.getSpreadsheetTimeZone();
  const cotizaciones = ss.getSheetByName(H_COT).getRange(2, 1, FILAS, COL.mail + 1).getValues()
    .filter(function (r) { return texto(r[COL.codigo]) !== ''; })
    .map(function (r) {
      return {
        codigo: texto(r[COL.codigo]), empresa: texto(r[COL.cliente]), contacto: texto(r[COL.contacto]), mail: texto(r[COL.mail]),
        titulo: texto(r[COL.titulo]), estado: texto(r[COL.estado]), pago: texto(r[COL.pago]),
        moneda: r[COL.moneda] === 'USD' ? 'USD' : 'ARS', total: numero(r[COL.total])
      };
    }).reverse();
  const entregas = ss.getSheetByName(H_ENT).getRange(2, 1, FILAS, COL_ENT.cotizacion + 1).getValues()
    .filter(function (r) { return texto(r[COL_ENT.codigo]) !== ''; })
    .map(function (r) {
      return {
        codigo: texto(r[COL_ENT.codigo]), empresa: texto(r[COL_ENT.empresa]), contacto: texto(r[COL_ENT.saludo]), mail: texto(r[COL_ENT.mail]),
        titulo: texto(r[COL_ENT.titulo]), estado: texto(r[COL_ENT.estado]), limite: dia(r[COL_ENT.limite], zona),
        cotizacion: texto(r[COL_ENT.cotizacion])
      };
    }).reverse();
  return {
    cotizaciones: cotizaciones,
    entregas: entregas,
    catalogo: leerCatalogo(ss),
    sitio: SITIO,
    sitioEntrega: SITIO_ENTREGA,
    opciones: {
      estadosCot: ESTADOS_COT,
      pagos: Object.keys(PAGOS),
      estadosEnt: ['Borrador'].concat(Object.keys(ESTADOS_ENT)),
      tipos: Object.keys(TIPOS_ARC),
      visibles: Object.keys(VISIBLES_ARC)
    }
  };
}

// lo que manda el formulario del panel, ya validado: sirve para crear y para editar
function leerFormulario(ss, d) {
  d = d || {};
  const catalogo = leerCatalogo(ss);
  const elegidos = (Array.isArray(d.items) ? d.items : []).slice(0, MAX_ITEMS).map(function (it) {
    const servicio = dato(it && it.servicio, 120);
    if (!servicio) throw new Error('A un ítem le falta el nombre.');
    const base = catalogo.filter(function (c) { return c.servicio === servicio; })[0];
    const conPrecio = texto(it.precio) !== '';
    return {
      servicio: servicio,
      descripcion: dato(it.descripcion, 1000) || (base ? base.descripcion : ''),
      entregables: dato(it.entregables, 1000),
      cantidad: texto(it.cantidad) === '' ? '' : importe(it.cantidad),
      unidad: dato(it.unidad, 60),
      precio: conPrecio ? importe(it.precio) : (base ? base.precio : ''),
      descuentoPct: texto(it.descuentoPct) === '' ? '' : importe(it.descuentoPct),
      mostrar: it.mostrar !== false,
      alCatalogo: !base && it.alCatalogo === true
    };
  });
  const extra = d.extra || {};
  const conDescuento = texto(extra.descuentoValor) !== '';
  const anticipo = texto(extra.anticipoPct) === '' ? '' : importe(extra.anticipoPct);
  if (anticipo !== '' && anticipo > 100) throw new Error('El anticipo no puede pasar de 100.');
  const vencimiento = texto(extra.vencimiento);
  if (vencimiento && !/^\d{4}-\d{2}-\d{2}$/.test(vencimiento)) throw new Error('La fecha no es válida.');
  const empresa = dato(d.empresa, 120);
  if (!empresa) throw new Error('Falta la empresa.');
  return {
    empresa: empresa, contacto: dato(d.contacto, 120), titulo: dato(d.titulo, 200), intro: dato(extra.intro, 1000),
    mail: mailDelPanel(d.mail),
    moneda: extra.moneda ? elegir(extra.moneda, ['ARS', 'USD']) : 'ARS',
    descuentoTipo: conDescuento ? elegir(extra.descuentoTipo, ['%', 'Monto']) : '',
    descuentoValor: conDescuento ? importe(extra.descuentoValor) : '',
    descuentoMotivo: conDescuento ? dato(extra.descuentoMotivo, 200) : '',
    formaPago: dato(extra.formaPago, 300), anticipo: anticipo, plazoEntrega: dato(extra.plazoEntrega, 200),
    condiciones: dato(extra.condiciones, 1000), notas: dato(extra.notas, 1000),
    vencimiento: vencimiento ? new Date(vencimiento + 'T12:00:00') : null,
    items: elegidos
  };
}

function escribirFormulario(ss, hoja, fila, codigo, v) {
  hoja.getRange(fila, COL.cliente + 1, 1, 4).setValues([[v.empresa, v.contacto, v.titulo, v.intro]]);
  if (v.vencimiento) hoja.getRange(fila, COL.vencimiento + 1).setValue(v.vencimiento);
  hoja.getRange(fila, COL.moneda + 1, 1, 6).setValues([[v.moneda, v.descuentoTipo, v.descuentoValor, v.descuentoMotivo, v.formaPago, v.anticipo]]);
  hoja.getRange(fila, COL.plazoEntrega + 1, 1, 3).setValues([[v.plazoEntrega, v.condiciones, v.notas]]);
  hoja.getRange(fila, COL.mail + 1).setValue(v.mail);

  // se reusan los renglones que ya tiene; los que sobran quedan libres
  const items = ss.getSheetByName(H_ITEMS);
  const filas = buscarCodigo(items, codigo).findAll().map(function (c) { return c.getRow(); })
    .sort(function (a, b) { return a - b; });
  v.items.forEach(function (it, i) {
    const renglon = filas[i] || sumarRenglon(items, codigo);
    items.getRange(renglon, COL_ITEM.orden + 1).setValue(i + 1);
    items.getRange(renglon, COL_ITEM.servicio + 1, 1, 5).setValues([[it.servicio, it.descripcion, it.entregables, it.cantidad, it.unidad]]);
    items.getRange(renglon, COL_ITEM.precio + 1, 1, 3).setValues([[it.precio, it.descuentoPct, it.mostrar]]);
    if (it.alCatalogo) {
      const hojaCat = ss.getSheetByName(H_CAT);
      hojaCat.getRange(primeraLibre(hojaCat), 1, 1, 3).setValues([[it.servicio, it.descripcion, it.precio]]);
    }
  });
  filas.slice(v.items.length).forEach(function (f) { items.getRange(f, 1, 1, COL_ITEM.mostrar + 1).clearContent(); });
  marcarCambio(hoja, fila);
  guardarEnCache(codigo);
}

function panelNuevaCotizacion(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prefijo = texto(d && d.prefijo).toUpperCase();
  if (!/^[A-Z]{3}$/.test(prefijo)) throw new Error('Son 3 letras, sin acentos ni números.');
  const v = leerFormulario(ss, d);
  const creada = crearCotizacion(ss, prefijo);
  escribirFormulario(ss, ss.getSheetByName(H_COT), creada.fila, creada.codigo, v);
  return panelDatos();
}

function panelCotizacion(codigo) {
  codigo = pedirCodigo(codigo);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(H_COT);
  const f = hoja.getRange(filaDe(hoja, codigo), 1, 1, COL.mail + 1).getValues()[0];
  return {
    codigo: codigo, empresa: texto(f[COL.cliente]), contacto: texto(f[COL.contacto]), mail: texto(f[COL.mail]),
    titulo: texto(f[COL.titulo]), intro: texto(f[COL.intro]), moneda: f[COL.moneda] === 'USD' ? 'USD' : 'ARS',
    vencimiento: dia(f[COL.vencimiento], ss.getSpreadsheetTimeZone()), descuentoTipo: texto(f[COL.descuentoTipo]) || '%',
    descuentoValor: texto(f[COL.descuentoValor]), descuentoMotivo: texto(f[COL.descuentoMotivo]), formaPago: texto(f[COL.formaPago]),
    anticipoPct: texto(f[COL.anticipoPct]), plazoEntrega: texto(f[COL.plazoEntrega]), condiciones: texto(f[COL.condiciones]), notas: texto(f[COL.notas]),
    items: renglonesDe(ss.getSheetByName(H_ITEMS), codigo, COL_ITEM.mostrar + 1)
      .filter(function (r) { return String(r[COL_ITEM.codigo]) === codigo && texto(r[COL_ITEM.servicio]) !== ''; })
      .sort(function (a, b) { return numero(a[COL_ITEM.orden]) - numero(b[COL_ITEM.orden]); })
      .map(function (r) {
        return {
          servicio: texto(r[COL_ITEM.servicio]), descripcion: texto(r[COL_ITEM.descripcion]), entregables: texto(r[COL_ITEM.entregables]),
          cantidad: texto(r[COL_ITEM.cantidad]), unidad: texto(r[COL_ITEM.unidad]), precio: texto(r[COL_ITEM.precio]),
          descuentoPct: texto(r[COL_ITEM.descuentoPct]), mostrar: r[COL_ITEM.mostrar] !== false
        };
      })
  };
}

function panelGuardarCotizacion(codigo, d) {
  codigo = pedirCodigo(codigo);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(H_COT);
  const fila = filaDe(hoja, codigo);
  escribirFormulario(ss, hoja, fila, codigo, leerFormulario(ss, d));
  return panelDatos();
}

function panelCambiarCotizacion(codigo, campo, valor) {
  codigo = pedirCodigo(codigo);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(H_COT);
  if (campo === 'mail') {
    const filaMail = filaDe(hoja, codigo);
    hoja.getRange(filaMail, COL.mail + 1).setValue(mailDelPanel(valor));
    marcarCambio(hoja, filaMail);
    return panelDatos();
  }
  const opciones = { estado: ESTADOS_COT, pago: Object.keys(PAGOS) }[campo];
  if (!opciones) throw new Error('Ese dato no se cambia desde el panel.');
  elegir(valor, opciones);
  const fila = filaDe(hoja, codigo);
  hoja.getRange(fila, (campo === 'estado' ? COL.estado : COL.pago) + 1).setValue(valor);
  marcarCambio(hoja, fila);
  guardarEnCache(codigo);
  entregaDeCotizacion(ss, hoja, fila);
  return panelDatos();
}

function panelCambiarEntrega(codigo, campo, valor) {
  codigo = pedirCodigo(codigo);
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H_ENT);
  if (campo === 'mail') {
    hoja.getRange(filaDe(hoja, codigo), COL_ENT.mail + 1).setValue(mailDelPanel(valor));
    return panelDatos();
  }
  if (campo !== 'estado') throw new Error('Ese dato no se cambia desde el panel.');
  elegir(valor, ['Borrador'].concat(Object.keys(ESTADOS_ENT)));
  const celda = hoja.getRange(filaDe(hoja, codigo), COL_ENT.estado + 1);
  celda.setValue(valor);
  ordenarPlazo(hoja, celda);
  guardarEntrega(codigo);
  return panelDatos();
}

function panelRenglones(codigo) {
  codigo = pedirCodigo(codigo);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const zona = ss.getSpreadsheetTimeZone();
  filaDe(ss.getSheetByName(H_ENT), codigo);
  return renglonesDe(ss.getSheetByName(H_ARC), codigo, COL_ARC.visible + 1)
    .filter(function (r) { return String(r[COL_ARC.codigo]) === codigo; })
    .sort(function (a, b) { return numero(a[COL_ARC.orden]) - numero(b[COL_ARC.orden]); })
    .map(function (r) {
      return {
        orden: numero(r[COL_ARC.orden]), nombre: texto(r[COL_ARC.nombre]), tipo: texto(r[COL_ARC.tipo]) || 'Descarga',
        link: texto(r[COL_ARC.link]), detalle: texto(r[COL_ARC.detalle]), fecha: dia(r[COL_ARC.fecha], zona),
        visible: texto(r[COL_ARC.visible]) || 'Al pagar'
      };
    });
}

// vacío borra el link; si no, tiene que ser una dirección web que la página pueda abrir
const linkDelPanel = function (v) {
  if (texto(v) === '') return '';
  const link = linkValido(v);
  if (!link) throw new Error('Ese link no sirve. Pegalo como lo copiaste del navegador.');
  return link.slice(0, 1000);
};

function panelSumarArchivo(codigo, d) {
  codigo = pedirCodigo(codigo);
  d = d || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  filaDe(ss.getSheetByName(H_ENT), codigo);
  const nombre = dato(d.nombre, 200);
  if (!nombre) throw new Error('Falta el nombre de lo que entregás.');
  const tipo = elegir(d.tipo || 'Descarga', Object.keys(TIPOS_ARC));
  const visible = d.visible ? elegir(d.visible, Object.keys(VISIBLES_ARC)) : '';
  const link = linkDelPanel(d.link);
  const archivos = ss.getSheetByName(H_ARC);
  const fila = sumarArchivo(archivos, codigo);
  archivos.getRange(fila, COL_ARC.nombre + 1, 1, 2).setValues([[nombre, tipo]]);
  archivos.getRange(fila, COL_ARC.link + 1, 1, 2).setValues([[link, dato(d.detalle, 300)]]);
  if (visible) archivos.getRange(fila, COL_ARC.visible + 1).setValue(visible);
  guardarEntrega(codigo);
  return panelRenglones(codigo);
}

function filaArchivo(archivos, codigo, orden) {
  const i = archivos.getRange(2, 1, FILAS, COL_ARC.visible + 1).getValues()
    .findIndex(function (r) { return String(r[COL_ARC.codigo]) === codigo && numero(r[COL_ARC.orden]) === numero(orden); });
  if (i < 0) throw new Error('No encontré ese renglón.');
  return i + 2;
}

// se busca por id: la carpeta se puede renombrar o mover, y si se borra se crea otra
function carpetaFotos() {
  const propiedades = PropertiesService.getScriptProperties();
  const id = propiedades.getProperty(PROP_CARPETA);
  if (id) {
    try {
      const carpeta = DriveApp.getFolderById(id);
      if (!carpeta.isTrashed()) return carpeta;
    } catch (err) {
      // ya no existe: se crea otra
    }
  }
  const nueva = DriveApp.createFolder(CARPETA_FOTOS);
  propiedades.setProperty(PROP_CARPETA, nueva.getId());
  return nueva;
}

function panelSubirFoto(codigo, orden, d) {
  codigo = pedirCodigo(codigo);
  d = d || {};
  const tipo = elegir(d.tipo, ['image/webp', 'image/jpeg']);
  const base64 = String(d.base64 || '');
  if (!base64 || base64.length > MAX_FOTO) throw new Error('La foto es muy pesada.');
  const hojaArchivos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H_ARC);
  const fila = filaArchivo(hojaArchivos, codigo, orden);
  const nombre = codigo + '-' + numero(orden) + '-' + new Date().getTime() + (tipo === 'image/webp' ? '.webp' : '.jpg');
  const archivo = carpetaFotos().createFile(Utilities.newBlob(Utilities.base64Decode(base64), tipo, nombre));
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  hojaArchivos.getRange(fila, COL_ARC.tipo + 1, 1, 2).setValues([['Entregado', archivo.getUrl()]]);
  guardarEntrega(codigo);
  return panelRenglones(codigo);
}

function panelEditarArchivo(codigo, orden, campo, valor) {
  codigo = pedirCodigo(codigo);
  const columna = { tipo: COL_ARC.tipo, visible: COL_ARC.visible, link: COL_ARC.link }[campo];
  if (columna === undefined) throw new Error('Ese dato no se cambia desde el panel.');
  const nuevo = campo === 'tipo' ? elegir(valor, Object.keys(TIPOS_ARC))
    : campo === 'visible' ? elegir(valor, Object.keys(VISIBLES_ARC)) : linkDelPanel(valor);
  const archivos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(H_ARC);
  archivos.getRange(filaArchivo(archivos, codigo, orden), columna + 1).setValue(nuevo);
  guardarEntrega(codigo);
  return panelRenglones(codigo);
}

const ACCIONES_PANEL = {
  datos: panelDatos,
  nuevaCotizacion: panelNuevaCotizacion,
  cotizacion: panelCotizacion,
  guardarCotizacion: panelGuardarCotizacion,
  cambiarCotizacion: panelCambiarCotizacion,
  cambiarEntrega: panelCambiarEntrega,
  renglones: panelRenglones,
  sumarArchivo: panelSumarArchivo,
  subirFoto: panelSubirFoto,
  editarArchivo: panelEditarArchivo
};

/* ---------- armado de la planilla ---------- */

const C_CREMA = '#F2EBDD';
const C_CREMA2 = '#E9DFCB';
const C_ESPRESSO = '#1A1410';
const C_MARRON = '#5C5047';
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
  armarEntregas(ss);
  armarArchivos(ss);
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
    'Total', 'Última modificación', 'Link', 'Mail del cliente'], [COL.codigo, COL.total, COL.actualizada, COL.link]);

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
    ['Impuestos (por ejemplo: Los precios no incluyen IVA.)', '', 'impuestos'],
    ['Entregas: días para descargar', PLAZO_ENTREGA, 'plazoEntregaDias'],
    ['Entregas: días antes para avisar por mail', AVISO_ENTREGA, 'avisoEntregaDias']
  ];
  h.getRange(2, 2, filas.length, 1).setNumberFormat('@');
  h.getRange(2, 1, filas.length, 3).setValues(filas);
  h.getRange(2, 3, filas.length, 1).setBackground(C_CREMA2).protect().setDescription('No cambiar').setWarningOnly(true);
  h.setColumnWidth(1, 340);
  h.setColumnWidth(2, 420);
  h.setColumnWidth(3, 120);
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
    .whenFormulaSatisfied('=AND($A2<>"";$J2="Vencida")')
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

  // un link que no se va a abrir queda a la vista
  const linkMalo = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($E2<>"";NOT(REGEXMATCH($E2;"(?i)^(https?://|[a-z0-9-]+(\\.[a-z0-9-]+)+([/?#]|$))")))')
    .setBackground(C_ROJO).setFontColor(C_CREMA)
    .setRanges([h.getRange(2, COL_ARC.link + 1, FILAS, 1)]).build();
  // cada entrega toma su color según su fila en Entregas
  bandas(h, COL_ARC.visible + 1, '=AND($A2<>"";ISEVEN(MATCH($A2;INDIRECT("' + H_ENT + '!$A:$A");0)))', [linkMalo]);
  h.setColumnWidths(1, COL_ARC.visible + 1, 110);
  h.setColumnWidth(COL_ARC.nombre + 1, 240);
  h.setColumnWidth(COL_ARC.link + 1, 320);
  h.setColumnWidth(COL_ARC.detalle + 1, 280);
}

function armarUso(ss) {
  const h = ss.insertSheet(H_USO);
  const lineas = [
    'FÁRA · Cotizaciones y entregas',
    'Esta planilla es privada. El cliente solo ve lo suyo en la web, con su link.',
    '',
    'COTIZACIONES · pestañas ámbar · menú FÁRA Cotizaciones',
    'Crear una cotización',
    '1. Menú FÁRA Cotizaciones → Nueva cotización. Escribí las 3 letras del cliente.',
    '2. En Cotizaciones, completá la fila nueva. Empresa es la marca. Nombre para el mail es a quién le escribís.',
    '3. Parado en esa fila, tocá FÁRA Cotizaciones → Ver ítems y completá los renglones. Para sumar otro: Agregar ítem. En Precio va el total del renglón: la cantidad no multiplica.',
    '4. Cambiá el Estado a Enviada y mandale el link de la última columna.',
    '',
    'Cambiar una cotización',
    'Editá la celda. El cliente ve el cambio al recargar la página. Para volver a la lista completa: FÁRA Cotizaciones → Ver todos los ítems.',
    '',
    'Estados',
    'Borrador: el link dice que no está disponible.',
    'Enviada: se ve completa. Si pasa la fecha de vencimiento, se muestra vencida sola.',
    'Aceptada: se ve completa con el aviso de aceptada. Al ponerla vos en Aceptada, o marcar un pago, se crea sola su entrega con los mismos ítems.',
    'Pago: elegí Anticipo pagado o Pagado total a medida que cobrás. Si te paga otro anticipo, cambiá primero Anticipo %. El link sigue activo como registro.',
    'Vencida: se ve completa con el botón para pedir actualización.',
    'Pausada y Anulada: un mensaje, sin datos ni precios.',
    '',
    'ENTREGAS · pestañas petróleo · menú FÁRA Entregas',
    'Cómo nace una entrega',
    'Sola: al poner una cotización en Aceptada, o marcar un pago, aparece en Entregas con sus ítems en Archivos. Completá el Mail del cliente.',
    'Suelta: menú FÁRA Entregas → Nueva entrega. Escribí las 3 letras del cliente y completá la fila.',
    'Empresa sale en la página. Nombre para el mail es el saludo: Hola Leandro.',
    'Parado en la fila, tocá FÁRA Entregas → Ver archivos. Podés agregar, cambiar u ocultar renglones. Mandale el link de la columna Link.',
    '',
    'Archivos',
    'Tipo: Descarga vence. Web no vence. Entregado es sin link, como una impresión o una charla. En proceso es lo que todavía falta: cuando esté, cambiá el Tipo y pegá el link.',
    'Cuándo se ve: Al pagar espera el pago. Entrega parcial se abre desde ya. Oculto no aparece.',
    'Un link marcado en rojo no se va a abrir: revisalo.',
    '',
    'Estados',
    'Borrador: el link dice que no está disponible.',
    'Pendiente de pago: avisa que falta el pago y muestra solo las entregas parciales. El plazo no corre.',
    'En proceso: el cliente ve la lista mientras la vas terminando. Lo que espera el pago sale sin link. El plazo no corre y no salen mails.',
    'Disponible: se ve todo. Si la Fecha límite está vacía, se completa sola con los días de Config.',
    'Vencida: pasa sola el día después de la Fecha límite y la fila se pone en rojo. Borrá o cerrá la carpeta y pasala a Archivos borrados.',
    'Pausada y Anulada: un mensaje, sin datos ni links.',
    '',
    'GENERAL · pestañas grises',
    'Qué se puede tocar',
    'Encabezado de color: se puede editar.',
    'Encabezado rojo: automático, no lo toques.',
    'No borres filas: quedan como historial. Para sacar algo de la web, usá Anulada u Oculto.'
  ];
  h.getRange(1, 1, lineas.length, 1).setValues(lineas.map(function (l) { return [l]; }))
    .setFontFamily('Archivo').setBackground(C_CREMA).setFontColor(C_ESPRESSO).setWrap(true);
  h.setColumnWidth(1, 720);
  h.getRange(1, 1).setFontSize(20).setFontWeight('bold');
  [[4, C_AMBAR], [22, C_PETROLEO], [41, C_GRIS]].forEach(function (b) {
    h.getRange(b[0], 1).setBackground(b[1]).setFontColor(C_CREMA).setFontWeight('bold').setFontSize(13);
  });
  [5, 11, 14, 23, 29, 34, 42].forEach(function (f) { h.getRange(f, 1).setFontWeight('bold').setFontSize(12); });
  h.getRange(43, 1).setBackground(C_PETROLEO).setFontColor(C_CREMA);
  h.getRange(44, 1).setBackground(C_ROJO).setFontColor(C_CREMA);
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
