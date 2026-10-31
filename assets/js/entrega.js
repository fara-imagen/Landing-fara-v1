/* ==========================================================
   FÁRA · assets/js/entrega.js
   Página de entrega: lee el código de la dirección, le pide esa entrega
   al script de la planilla "FÁRA · Cotizaciones" y la muestra.
   El script decide qué links salen: acá nunca llega uno vencido o sin pagar.
   Todo el contenido se escribe como texto (nunca como HTML) y una celda
   vacía simplemente no muestra su bloque.
   ========================================================== */
(function (w, d) {
  'use strict';

  /* API: la misma dirección /exec de assets/js/cotizacion.js */
  const API = 'https://script.google.com/macros/s/AKfycbxkYwQ-kbpUrgIhF8RF58sCLI0ZEj-TtX3VcdI8tccmWyJNjqBTI6rKr6DZVTwiBxC47A/exec';
  const WHATSAPP = '5493885735415';
  const MAIL = 'hola@faraimagen.com.ar';
  const INSTAGRAM = 'fara.imagen';
  const ESPERA_MAX = 20000;
  const DIAS_URGENTE = 5;
  const BOTONES = { descarga: 'Descargar', web: 'Abrir', entregado: 'Ver' };

  const raiz = d.getElementById('ent');

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

  /* las fechas viajan como AAAA-MM-DD: se leen en hora local para no correrse un día */
  function fechaLarga(txt) {
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(txt || ''));
    if (!m) return '';
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]))
      .toLocaleDateString('es-AR', { day: 'numeric', month: 'long', year: 'numeric' });
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

  function boton(clase, texto, href, etiqueta) {
    return h('a', { class: 'btn ' + clase, href: href, target: '_blank', rel: 'noopener noreferrer', 'aria-label': etiqueta }, texto);
  }

  /* ---------- bloques ---------- */

  function textoPlazo(dias) {
    if (dias === 0) return 'Último día';
    if (dias === 1) return 'Queda 1 día';
    return 'Quedan ' + dias + ' días';
  }

  function cabecera(e) {
    const datos = h('dl', { class: 'cot-datos' },
      lleno(e.limite) && h('div', {},
        h('dt', {}, e.estado === 'vencida' ? 'Venció' : 'Disponible hasta'),
        h('dd', {}, fechaLarga(e.limite))));
    return h('div', { class: 'post-hero' },
      h('span', { class: 'eyebrow' }, 'Entrega', h('span', { class: 'cot-codigo' }, e.codigo)),
      h('h1', {}, lleno(e.titulo) ? e.titulo : 'Tu entrega'),
      lleno(e.cliente) && h('p', { class: 'cot-para' }, 'Para ' + e.cliente),
      lleno(e.mensaje) && h('p', { class: 'cot-para cot-texto' }, e.mensaje),
      datos.childElementCount ? datos : null,
      e.estado === 'disponible' && e.dias !== null && h('p', { class: 'cot-plazo' }, h('mark', {}, textoPlazo(e.dias)))
    );
  }

  function aviso(titulo, texto, botonAviso, clase) {
    return h('div', { class: 'cot-aviso' + (clase ? ' ' + clase : '') }, h('h2', {}, titulo), h('p', {}, texto), botonAviso);
  }

  function avisoEstado(e) {
    const wa = (e.config || {}).whatsapp;
    if (e.estado === 'pendiente') {
      return aviso('Tu trabajo está listo',
        'La entrega completa se habilita cuando se complete el pago. Apenas esté, la vas a ver en este mismo link.',
        boton('btn-accent', 'Escribime por WhatsApp',
          linkWhatsApp(wa, 'Hola Matías, te escribo por el pago de la entrega ' + e.codigo + '.')));
    }
    if (e.estado === 'proceso') {
      return aviso('Tu trabajo está en marcha',
        'Vas a ver cada parte a medida que la termine.');
    }
    if (e.estado === 'vencida') {
      return aviso('El plazo de descarga terminó',
        'Los archivos ya no están disponibles. Escribime y vemos cómo te los reenvío.',
        boton('btn-accent', 'Pedir los archivos',
          linkWhatsApp(wa, 'Hola Matías, venció la entrega ' + e.codigo + '. ¿Me reenviás los archivos?')));
    }
    if (e.dias === null || !e.archivos.some(function (a) { return a.tipo === 'descarga' && lleno(a.link); })) return null;
    const urgente = e.dias <= DIAS_URGENTE;
    return aviso(e.dias === 0 ? 'Hoy es el último día para descargar' : 'Descargá antes del ' + fechaLarga(e.limite),
      'Guardá todo en tu compu o en tu nube. ' + (e.dias === 0 ? 'Mañana' : 'Después de esa fecha') + ' los archivos dejan de estar disponibles.',
      null, urgente ? 'ent-aviso-urgente' : '');
  }

  function renglon(a) {
    const enProceso = a.tipo === 'proceso';
    const visita = enProceso ? h('mark', {}, 'En proceso')
      : a.bloqueado ? h('mark', {}, 'Se habilita con el pago')
      : lleno(a.link) ? boton('btn-accent', BOTONES[a.tipo] || 'Abrir', a.link, (BOTONES[a.tipo] || 'Abrir') + ': ' + a.nombre) : null;
    const pie = a.vencido ? 'Venció el plazo de descarga'
      : enProceso && lleno(a.fecha) ? 'Fecha estimada: ' + fechaLarga(a.fecha)
      : a.tipo === 'entregado' && lleno(a.fecha) ? 'Entregado el ' + fechaLarga(a.fecha)
      : a.tipo === 'entregado' ? 'Entregado'
      : lleno(a.fecha) ? fechaLarga(a.fecha) : '';
    return h('article', { class: 'post-body ent-item' + (a.vencido ? ' ent-item-vencido' : '') },
      h('div', { class: 'ent-item-texto' },
        h('h3', {}, a.nombre),
        lleno(a.detalle) && h('p', { class: 'cot-texto' }, a.detalle),
        pie && h('p', { class: 'meta' }, a.tipo === 'entregado' ? icono('M5 12.5l4.5 4.5L19 7.5') : null, pie)),
      visita
    );
  }

  function seccion(titulo, archivos) {
    if (!archivos.length) return null;
    return h('section', { class: 'cot-seccion', 'aria-label': titulo },
      h('h2', {}, titulo),
      h.apply(null, ['div', { class: 'ent-lista' }].concat(archivos.map(renglon))));
  }

  function pie(e) {
    const cfg = e.config || {};
    const wa = cfg.whatsapp || WHATSAPP;
    const mail = cfg.mail || MAIL;
    const ig = String(cfg.instagram || INSTAGRAM).replace(/^@/, '');
    d.getElementById('entContacto').replaceChildren(
      h('a', { href: 'https://wa.me/' + String(wa).replace(/\D/g, ''), target: '_blank', rel: 'noopener' }, 'WhatsApp'),
      h('a', { href: 'mailto:' + mail }, mail),
      h('a', { href: 'https://instagram.com/' + ig, target: '_blank', rel: 'noopener' }, '@' + ig)
    );
    d.getElementById('entPie').hidden = false;
  }

  /* ---------- pantallas ---------- */

  function mostrar(contenido) {
    raiz.replaceChildren(contenido);
    raiz.removeAttribute('aria-busy');
  }

  /* pantalla de aviso: usa .nf de style.css */
  function pantalla(titulo, texto, codigo, contenidoBoton) {
    const mensaje = 'Hola Matías, te escribo por la entrega' + (codigo ? ' ' + codigo : '') + '.';
    mostrar(h('section', { class: 'nf' },
      h('div', { class: 'wrap nf-box' },
        h('h1', {}, titulo),
        h('p', {}, texto),
        contenidoBoton || boton('btn-accent', 'Escribime por WhatsApp', linkWhatsApp(WHATSAPP, mensaje))
      )
    ));
  }

  function noDisponible() {
    pantalla('Esta entrega no está disponible', 'Si te llegó este link, escribime y lo revisamos.');
  }

  function entrega(e) {
    if (lleno(e.logoCliente)) {
      const logo = d.getElementById('entLogoCliente');
      logo.onerror = function () { logo.hidden = true; };
      logo.src = e.logoCliente;
      logo.alt = lleno(e.cliente) ? e.cliente : '';
      logo.hidden = false;
    }
    e.archivos = e.archivos || [];
    const pendiente = e.estado === 'pendiente';

    mostrar(h('div', { class: 'wrap cot-wrap' },
      cabecera(e),
      avisoEstado(e),
      seccion(pendiente ? 'Entregas parciales' : 'Qué incluye', e.archivos)
    ));
    pie(e);
  }

  function pintar(e) {
    if (!e || !e.ok) return noDisponible();
    if (e.estado === 'anulada') {
      return pantalla('Esta entrega fue anulada', 'Si tenés alguna duda, escribime.', e.codigo);
    }
    if (e.estado === 'pausada') {
      return pantalla('Esta entrega está en pausa', 'La estoy revisando. Si necesitás algo, escribime.', e.codigo);
    }
    entrega(e);
  }

  /* ---------- arranque ---------- */

  function codigoDeLaDireccion() {
    const m = /^\/entrega\/([^/]+)\/?$/.exec(w.location.pathname);
    const codigo = m ? decodeURIComponent(m[1]) : '';
    return /^[A-Za-z]{3}-[a-z2-9]{6,8}$/.test(codigo) ? codigo : '';
  }

  function cargar() {
    const codigo = codigoDeLaDireccion();
    if (!codigo || !API) return noDisponible();
    const corte = new AbortController();
    const reloj = w.setTimeout(function () { corte.abort(); }, ESPERA_MAX);
    w.fetch(API + '?entrega=' + encodeURIComponent(codigo), { signal: corte.signal })
      .then(function (r) { if (!r.ok) throw new Error(String(r.status)); return r.json(); })
      .then(pintar)
      .catch(function () {
        pantalla('No pudimos cargar la entrega', 'Volvé a intentar en un momento.', codigo,
          h('a', { class: 'btn btn-accent', href: w.location.href }, 'Reintentar'));
      })
      .finally(function () { w.clearTimeout(reloj); });
  }

  cargar();
})(window, document);
