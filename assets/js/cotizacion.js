/* ==========================================================
   FÁRA · assets/js/cotizacion.js
   Página de cotización: lee el código de la dirección, le pide esa
   cotización al script de la planilla "FÁRA · Cotizaciones" y la muestra.
   Todo el contenido se escribe como texto (nunca como HTML) y una celda
   vacía simplemente no muestra su bloque.
   ========================================================== */
(function (w, d) {
  'use strict';

  /* API: la dirección /exec del script publicado (ver cotizaciones/README.md).
     Vacía = la página muestra "no disponible". */
  const API = 'https://script.google.com/macros/s/AKfycbxkYwQ-kbpUrgIhF8RF58sCLI0ZEj-TtX3VcdI8tccmWyJNjqBTI6rKr6DZVTwiBxC47A/exec';
  const WHATSAPP = '5493885735415';
  const MAIL = 'hola@faraimagen.com.ar';
  const INSTAGRAM = 'fara.imagen';
  const ESPERA_MAX = 20000;

  const raiz = d.getElementById('cot');

  /* ---------- utilidades ---------- */

  /* crea un elemento: h('p',{class:'x'},'texto',otroElemento) */
  function h(tag, attrs) {
    const el = d.createElement(tag);
    Object.keys(attrs || {}).forEach(function (k) {
      if (attrs[k] !== undefined && attrs[k] !== null) el.setAttribute(k, attrs[k]);
    });
    Array.prototype.slice.call(arguments, 2).forEach(function (c) {
      if (c === null || c === undefined || c === false) return;
      el.append(c);
    });
    return el;
  }

  const lleno = function (v) { return v !== null && v !== undefined && String(v).trim() !== ''; };
  const num = function (v) { const n = Number(v); return isFinite(n) ? n : 0; };
  const redondo = function (n) { return Math.round(n * 100) / 100; };

  function plata(n, moneda) {
    const valor = redondo(n);
    return new Intl.NumberFormat('es-AR', {
      style: 'currency',
      currency: moneda === 'USD' ? 'USD' : 'ARS',
      minimumFractionDigits: Number.isInteger(valor) ? 0 : 2,
      maximumFractionDigits: 2
    }).format(valor);
  }

  /* las fechas viajan como AAAA-MM-DD: se leen en hora local para no correrse un día */
  function fecha(txt) {
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(txt || ''));
    return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
  }
  function fechaLarga(txt) {
    const f = fecha(txt);
    return f ? f.toLocaleDateString('es-AR', { day: 'numeric', month: 'long', year: 'numeric' }) : '';
  }
  function diasHasta(txt) {
    const f = fecha(txt);
    if (!f) return null;
    const hoy = new Date();
    return Math.round((f - new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate())) / 86400000);
  }

  function linkWhatsApp(numero, texto) {
    return 'https://wa.me/' + (String(numero || '').replace(/\D/g, '') || WHATSAPP) + '?text=' + encodeURIComponent(texto);
  }

  function icono(ruta) {
    const svg = d.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('viewBox', '0 0 24 24');
    svg.setAttribute('aria-hidden', 'true');
    const p = d.createElementNS('http://www.w3.org/2000/svg', 'path');
    p.setAttribute('d', ruta);
    svg.appendChild(p);
    return svg;
  }

  /* ---------- cuentas ---------- */

  function calcular(c) {
    const items = (c.items || []).map(function (it) {
      /* el precio es el total del renglón: la cantidad es solo informativa */
      const bruto = num(it.precio);
      const descuento = bruto * Math.min(num(it.descuentoPct), 100) / 100;
      return { bruto: bruto, descuento: descuento, neto: bruto - descuento };
    });
    const subtotal = items.reduce(function (s, i) { return s + i.bruto; }, 0);
    const descItems = items.reduce(function (s, i) { return s + i.descuento; }, 0);
    const neto = subtotal - descItems;
    const general = c.descuentoTipo === '%'
      ? neto * Math.min(num(c.descuentoValor), 100) / 100
      : Math.min(num(c.descuentoValor), neto);
    const total = neto - general;
    const anticipo = total * Math.min(num(c.anticipoPct), 100) / 100;
    return {
      items: items, subtotal: subtotal, descuentos: descItems + general, general: general,
      total: total, anticipo: anticipo, saldo: total - anticipo,
      /* sin anticipo no hay nada que marcar como "anticipo pagado" */
      pago: c.pago === 'total' || (c.pago === 'anticipo' && anticipo > 0) ? c.pago : 'pendiente'
    };
  }

  /* ---------- bloques ---------- */

  function fila(nombre, valor, clase) {
    return h('div', { class: 'cot-fila' + (clase ? ' ' + clase : '') }, h('dt', {}, nombre), h('dd', {}, valor));
  }

  function textoPlazo(dias) {
    if (dias === 0) return 'Vence hoy';
    if (dias === 1) return 'Queda 1 día';
    return 'Quedan ' + dias + ' días';
  }

  function cabecera(c, vencida, aceptada) {
    const dias = diasHasta(c.vencimiento);
    const datos = h('dl', { class: 'cot-datos' },
      lleno(c.emision) && h('div', {}, h('dt', {}, 'Fecha'), h('dd', {}, fechaLarga(c.emision))),
      lleno(c.vencimiento) && h('div', {}, h('dt', {}, vencida ? 'Venció' : 'Válida hasta'), h('dd', {}, fechaLarga(c.vencimiento)))
    );
    return h('div', { class: 'post-hero' },
      h('span', { class: 'eyebrow' }, 'Propuesta', h('span', { class: 'cot-codigo' }, c.codigo)),
      h('h1', {}, lleno(c.titulo) ? c.titulo : 'Tu propuesta'),
      lleno(c.cliente) && h('p', { class: 'cot-para' }, 'Para ' + c.cliente + (lleno(c.contacto) ? ' · ' + c.contacto : '')),
      datos.childElementCount ? datos : null,
      !vencida && !aceptada && dias !== null && h('p', { class: 'cot-plazo' }, h('mark', {}, textoPlazo(dias)))
    );
  }

  function bloqueItem(it, cuenta, c) {
    const entregables = lleno(it.entregables)
      ? String(it.entregables).split('\n').map(function (t) { return t.trim(); }).filter(Boolean) : [];
    const cantidad = lleno(it.cantidad) ? num(it.cantidad) : 1;
    const conPrecio = cuenta.bruto > 0;
    const detalle = (cantidad === 1 && !lleno(it.unidad)) ? '' : cantidad + (lleno(it.unidad) ? ' ' + it.unidad : '');
    return h('article', { class: 'post-body cot-item' },
      h('h3', {}, it.servicio || ''),
      lleno(it.descripcion) && h('p', { class: 'cot-texto' }, it.descripcion),
      entregables.length ? h.apply(null, ['ul', {}].concat(entregables.map(function (t) { return h('li', {}, t); }))) : null,
      conPrecio && h('div', { class: 'cot-item-pie' },
        h('div', { class: 'cot-item-cuenta' },
          detalle && h('span', { class: 'meta' }, detalle),
          h('span', { class: 'cot-precio' },
            cuenta.descuento > 0 && h('s', {}, plata(cuenta.bruto, c.moneda)),
            plata(cuenta.neto, c.moneda))
        ),
        cuenta.descuento > 0 && h('p', { class: 'cot-ahorro' }, h('mark', {}, 'Ahorrás ' + plata(cuenta.descuento, c.moneda)))
      )
    );
  }

  function filaPagado() {
    return h('div', { class: 'cot-fila cot-fila-pagado' },
      h('dt', {}, icono('M5 12.5l4.5 4.5L19 7.5'), 'Pagado'),
      h('dd', {}));
  }

  function resumen(c, t, acciones) {
    const m = c.moneda;
    const hayAnticipo = t.anticipo > 0;
    const anticipoPagado = hayAnticipo && t.pago !== 'pendiente';
    const saldoPagado = hayAnticipo && t.pago === 'total';
    const totalPagado = !hayAnticipo && t.pago === 'total';
    const filas = h('dl', { class: 'cot-filas' },
      t.descuentos > 0 && fila('Subtotal', plata(t.subtotal, m)),
      t.descuentos > 0 && fila('Descuentos', '− ' + plata(t.descuentos, m)),
      fila('Total', plata(t.total, m), 'cot-fila-total' + (totalPagado ? ' cot-fila-tachada' : '')),
      t.descuentos > 0 && h('div', { class: 'cot-fila cot-fila-ahorro' },
        h('dt', { class: 'cot-oculto' }, 'Ahorro'),
        h('dd', {}, h('mark', {}, 'Ahorrás ' + plata(t.descuentos, m)))),
      totalPagado && filaPagado(),
      hayAnticipo && fila('Anticipo (' + num(c.anticipoPct) + '%)', plata(t.anticipo, m), anticipoPagado ? 'cot-fila-tachada' : ''),
      anticipoPagado && filaPagado(),
      hayAnticipo && (saldoPagado
        ? fila('Saldo', plata(t.saldo, m), 'cot-fila-tachada')
        : fila(anticipoPagado ? 'Saldo pendiente' : 'Saldo', plata(t.saldo, m), anticipoPagado ? 'cot-fila-saldo' : '')),
      saldoPagado && filaPagado()
    );
    return h('section', { class: 'post-cta cot-resumen', 'aria-label': 'Resumen de la propuesta' },
      h('h2', {}, 'Resumen'),
      filas,
      t.general > 0 && lleno(c.descuentoMotivo) && h('p', { class: 'cot-motivo' }, c.descuentoMotivo),
      acciones
    );
  }

  function acordeon(titulo, partes) {
    const contenido = partes.filter(Boolean);
    if (!contenido.length) return null;
    const flecha = icono('M6 9l6 6 6-6');
    return h('details', { class: 'cot-acc' },
      h('summary', {}, titulo, flecha),
      h.apply(null, ['div', { class: 'cot-acc-cuerpo' }].concat(contenido))
    );
  }

  function parrafo(txt, etiqueta) {
    if (!lleno(txt)) return null;
    return h('p', { class: 'cot-texto' }, etiqueta ? h('strong', {}, etiqueta + ': ') : null, String(txt));
  }

  function condiciones(c) {
    const cfg = c.config || {};
    const bloques = [
      acordeon('Forma de pago', [parrafo(c.formaPago), parrafo(cfg.mediosPago), parrafo(cfg.impuestos)]),
      acordeon('Plazo de entrega', [parrafo(c.plazoEntrega)]),
      acordeon('Condiciones', [parrafo(c.condiciones)]),
      acordeon('Notas', [parrafo(c.notas)])
    ].filter(Boolean);
    if (!bloques.length) return null;
    return h.apply(null, ['section', { class: 'cot-seccion', 'aria-label': 'Pago y condiciones' }, h('h2', {}, 'Pago y condiciones')].concat(bloques));
  }

  function aviso(titulo, texto, boton) {
    return h('div', { class: 'cot-aviso' }, h('h2', {}, titulo), h('p', {}, texto), boton);
  }

  function boton(clase, texto, href) {
    return h('a', { class: 'btn ' + clase, href: href, target: '_blank', rel: 'noopener' }, texto);
  }

  function pie(c) {
    const cfg = c.config || {};
    const contacto = d.getElementById('cotContacto');
    const wa = cfg.whatsapp || WHATSAPP;
    const mail = cfg.mail || MAIL;
    const ig = String(cfg.instagram || INSTAGRAM).replace(/^@/, '');
    contacto.replaceChildren(
      h('a', { href: 'https://wa.me/' + String(wa).replace(/\D/g, ''), target: '_blank', rel: 'noopener' }, 'WhatsApp'),
      h('a', { href: 'mailto:' + mail }, mail),
      h('a', { href: 'https://instagram.com/' + ig, target: '_blank', rel: 'noopener' }, '@' + ig)
    );
    const actualizada = d.getElementById('cotActualizada');
    if (fecha(c.actualizada)) {
      actualizada.textContent = 'Última actualización: ' + fechaLarga(c.actualizada);
      actualizada.hidden = false;
    }
    const disclaimer = d.getElementById('cotDisclaimer');
    if (lleno(c.disclaimer)) {
      disclaimer.textContent = c.disclaimer;
      disclaimer.hidden = false;
    }
    d.getElementById('cotLegal').hidden = actualizada.hidden && disclaimer.hidden;
    d.getElementById('cotPie').hidden = false;
  }

  /* ---------- pantallas ---------- */

  function mostrar(contenido) {
    raiz.replaceChildren(contenido);
    raiz.removeAttribute('aria-busy');
  }

  /* pantalla de aviso: usa .nf de style.css */
  function pantalla(titulo, texto, codigo, contenidoBoton) {
    const mensaje = 'Hola Matías, te escribo por la propuesta' + (codigo ? ' ' + codigo : '') + '.';
    mostrar(h('section', { class: 'nf' },
      h('div', { class: 'wrap nf-box' },
        h('h1', {}, titulo),
        h('p', {}, texto),
        contenidoBoton || boton('btn-accent', 'Escribime por WhatsApp', linkWhatsApp(WHATSAPP, mensaje))
      )
    ));
  }

  function noDisponible() {
    pantalla('Esta cotización no está disponible', 'Si te llegó este link, escribime y lo revisamos.');
  }

  function propuesta(c) {
    const t = calcular(c);
    const dias = diasHasta(c.vencimiento);
    /* con un pago registrado la propuesta ya está en marcha: no vence ni se vuelve a aceptar */
    const conPago = t.pago !== 'pendiente';
    const vencida = !conPago && (c.estado === 'vencida' || (c.estado !== 'aceptada' && dias !== null && dias < 0));
    const aceptada = c.estado === 'aceptada' || conPago;
    const wa = (c.config || {}).whatsapp;

    if (lleno(c.logoCliente)) {
      const logo = d.getElementById('cotLogoCliente');
      logo.onerror = function () { logo.hidden = true; };
      logo.src = c.logoCliente;
      logo.alt = lleno(c.cliente) ? c.cliente : '';
      logo.hidden = false;
    }

    const consulta = boton('btn-sec', 'Tengo una consulta',
      linkWhatsApp(wa, 'Hola Matías, tengo una consulta sobre la propuesta ' + c.codigo + '.'));
    const acciones = (vencida || aceptada)
      ? h('div', { class: 'cot-acciones' }, consulta)
      : h('div', { class: 'cot-acciones' },
          boton('btn-accent', 'Aceptar propuesta',
            linkWhatsApp(wa, 'Hola Matías, acepto la propuesta ' + c.codigo + ' por un total de ' + plata(t.total, c.moneda) + '.')),
          consulta);

    const itemsVisibles = (c.items || []).length
      ? h.apply(null, ['section', { class: 'cot-seccion', 'aria-label': 'Qué incluye' }, h('h2', {}, 'Qué incluye')]
          .concat(c.items.map(function (it, i) { return bloqueItem(it, t.items[i], c); })))
      : null;

    mostrar(h('div', { class: 'wrap cot-wrap' },
      cabecera(c, vencida, aceptada),
      vencida && aviso('Esta propuesta venció', 'Los valores pueden haber cambiado. Escribime y te la actualizo.',
        boton('btn-accent', 'Pedir actualización',
          linkWhatsApp(wa, 'Hola Matías, la propuesta ' + c.codigo + ' venció. ¿Me la actualizás?'))),
      t.pago === 'total' && aviso('Propuesta pagada', 'Gracias. Este link queda disponible para que lo consultes cuando quieras.'),
      t.pago !== 'total' && c.estado === 'aceptada' && aviso('Propuesta aceptada', 'Gracias. Ya estamos trabajando con estos valores. Cualquier duda, escribime.'),
      h('div', { class: 'cot-cuerpo' },
        h('div', { class: 'cot-detalle' },
          lleno(c.intro) && h('div', { class: 'post-body cot-texto' }, c.intro),
          itemsVisibles),
        resumen(c, t, acciones),
        condiciones(c))
    ));
    pie(c);
  }

  function pintar(c) {
    if (!c || !c.ok) return noDisponible();
    if (c.estado === 'anulada') {
      return pantalla('Esta propuesta fue anulada', 'Si querés retomarla, escribime y armamos una nueva.', c.codigo);
    }
    if (c.estado === 'pausada') {
      return pantalla('Esta propuesta está en pausa', 'La estoy revisando. Si necesitás avanzar, escribime.', c.codigo);
    }
    propuesta(c);
  }

  /* ---------- arranque ---------- */

  function codigoDeLaDireccion() {
    const m = /^\/cotizacion\/([^/]+)\/?$/.exec(w.location.pathname);
    const codigo = m ? decodeURIComponent(m[1]) : '';
    return /^[A-Za-z]{3}-[a-z2-9]{6,8}$/.test(codigo) ? codigo : '';
  }

  function cargar() {
    const codigo = codigoDeLaDireccion();
    if (!codigo || !API) return noDisponible();
    const corte = new AbortController();
    const reloj = w.setTimeout(function () { corte.abort(); }, ESPERA_MAX);
    w.fetch(API + (API.indexOf('?') < 0 ? '?' : '&') + 'codigo=' + encodeURIComponent(codigo), { signal: corte.signal })
      .then(function (r) { if (!r.ok) throw new Error(String(r.status)); return r.json(); })
      .then(pintar)
      .catch(function () {
        pantalla('No pudimos cargar la propuesta', 'Volvé a intentar en un momento.', codigo,
          h('a', { class: 'btn btn-accent', href: w.location.href }, 'Reintentar'));
      })
      .finally(function () { w.clearTimeout(reloj); });
  }

  cargar();
})(window, document);
