/* ==========================================================
   FÁRA · assets/js/tracking.js
   Medición con Meta Pixel + API de Conversiones (servidor).

   Cada evento sale dos veces con el MISMO event_id:
     1) por el Pixel del navegador (fbq)
     2) por el servidor (/.netlify/functions/capi → Meta)
   Meta ve el mismo event_id y cuenta el evento una sola vez.

   Requisitos: el snippet del Pixel en el <head> define window.MK_PIXEL_ID
   y window.MK_PV_ID (id del PageView). Si MK_PIXEL_ID sigue siendo el
   marcador TU_PIXEL_ID, este archivo no envía nada (modo apagado).

   Eventos automáticos (por atributos HTML, sin tocar este archivo):
     - Cualquier link a wa.me ............ Contact
     - [data-track="agenda"] ............. Schedule
     - form[data-track-form] (submit) .... Lead
     - [data-track-view="Nombre"] ........ ViewContent al llegar a esa sección
   ========================================================== */
(function (w, d) {
  'use strict';

  var ENDPOINT = '/.netlify/functions/capi';
  var PIXEL = String(w.MK_PIXEL_ID || '');
  var ENABLED = PIXEL !== '' && PIXEL.indexOf('TU_') !== 0;
  var LEAD_KEY = 'mk_lead_sent';

  /* ---------- utilidades ---------- */

  /* id único por evento: es lo que Meta usa para no contarlo dos veces */
  function generateEventId() {
    try {
      if (w.crypto && w.crypto.randomUUID) return w.crypto.randomUUID();
      var b = new Uint8Array(16);
      w.crypto.getRandomValues(b);
      return Array.prototype.map.call(b, function (x) { return ('0' + x.toString(16)).slice(-2); }).join('');
    } catch (e) {
      return Date.now().toString(16) + Math.random().toString(16).slice(2);
    }
  }

  function readCookie(name) {
    var m = d.cookie.match(new RegExp('(?:^|;\\s*)' + name + '=([^;]*)'));
    return m ? decodeURIComponent(m[1]) : '';
  }

  /* cookie _fbp: la crea el Pixel e identifica el navegador */
  function getFbp() {
    return readCookie('_fbp');
  }

  /* cookie _fbc: identifica el clic en un anuncio. Si todavía no existe
     pero la URL trae fbclid, se arma con el formato de Meta */
  function getFbc() {
    var c = readCookie('_fbc');
    if (c) return c;
    try {
      var id = new URLSearchParams(w.location.search).get('fbclid');
      if (id) return 'fb.1.' + Date.now() + '.' + id;
    } catch (e) { /* sin fbclid */ }
    return '';
  }

  function store(key, val) {
    try { w.sessionStorage.setItem(key, val); } catch (e) { /* modo privado */ }
  }
  function stored(key) {
    try { return w.sessionStorage.getItem(key); } catch (e) { return null; }
  }

  /* ---------- envío al servidor ---------- */
  function sendToServer(name, id, data) {
    var body = JSON.stringify({
      event_name: name,
      event_id: id,
      event_source_url: w.location.href,
      fbp: getFbp(),
      fbc: getFbc(),
      custom_data: data || {}
    });
    try {
      w.fetch(ENDPOINT, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: body,
        keepalive: true /* sigue viajando aunque la página se cierre o abra WhatsApp */
      }).catch(function () { /* la medición nunca debe romper la web */ });
    } catch (e) { /* idem */ }
  }

  /* ---------- API pública ---------- */

  /* Dispara un evento por el Pixel y por el servidor con el mismo event_id.
     opts.eventId permite reutilizar un id ya creado. Devuelve el event_id. */
  function trackEvent(name, data, opts) {
    var id = (opts && opts.eventId) || generateEventId();
    data = data || {};
    if (!ENABLED) {
      if (w.console && console.debug) console.debug('[MK apagado]', name, id, data);
      return id;
    }
    try {
      if (typeof w.fbq === 'function') w.fbq('track', name, data, { eventID: id });
    } catch (e) { /* idem */ }
    sendToServer(name, id, data);
    return id;
  }

  /* Dispara un evento cada vez que se hace clic en el elemento (o elementos del selector).
     data puede ser un objeto o una función (elemento) => objeto. */
  function trackClick(target, name, data, opts) {
    var els = typeof target === 'string' ? d.querySelectorAll(target) : [target];
    Array.prototype.forEach.call(els, function (el) {
      if (!el || el.__mkClick) return;
      el.__mkClick = true;
      el.addEventListener('click', function () {
        trackEvent(name, typeof data === 'function' ? data(el) : data, opts);
      });
    });
  }

  /* Dispara ViewContent una sola vez cuando la sección entra en pantalla */
  function trackView(target, name, data) {
    var el = typeof target === 'string' ? d.querySelector(target) : target;
    if (!el) return;
    var fire = function () { trackEvent('ViewContent', merge({ content_name: name }, data)); };
    if (!('IntersectionObserver' in w)) { fire(); return; }
    var io = new IntersectionObserver(function (es) {
      es.forEach(function (e) {
        if (e.isIntersecting) { io.disconnect(); fire(); }
      });
    }, { threshold: 0.35 });
    io.observe(el);
  }

  /* Lead al enviar un formulario. Marca la sesión para que /gracias/ no lo repita */
  function trackForm(form, data) {
    if (!form || form.__mkForm) return;
    form.__mkForm = true;
    form.addEventListener('submit', function () {
      var id = trackEvent('Lead', merge({ content_name: 'Formulario' }, data));
      store(LEAD_KEY, id);
    });
  }

  /* Dispara un evento una sola vez por sesión del navegador (para /gracias/) */
  function trackOnce(name, data, key) {
    key = key || ('mk_once_' + name);
    if (stored(key)) return stored(key);
    var id = trackEvent(name, data);
    store(key, id);
    return id;
  }

  /* ViewContent para páginas de contenido (por ejemplo, un post del blog) */
  function viewContent(name, data) {
    return trackEvent('ViewContent', merge({ content_name: name }, data));
  }

  function merge(a, b) {
    for (var k in b) if (Object.prototype.hasOwnProperty.call(b, k)) a[k] = b[k];
    return a;
  }

  /* dónde de la página estaba el botón (para saber cuál convierte más) */
  function ubicacion(el) {
    if (el.closest('#mant')) return 'mantenimiento';
    if (el.closest('header')) return 'menu';
    if (el.closest('.hero')) return 'portada';
    if (el.closest('#contacto')) return 'cierre';
    if (el.closest('footer')) return 'pie';
    return 'pagina';
  }

  /* ---------- automatismos ---------- */
  function init() {
    /* PageView: el navegador ya lo mandó con el snippet del head; acá va el del servidor con el mismo id */
    if (ENABLED && w.MK_PV_ID) sendToServer('PageView', w.MK_PV_ID, {});

    /* clics: un solo listener para todos los botones (también los que se crean después) */
    d.addEventListener('click', function (e) {
      var t = e.target && e.target.closest ? e.target : null;
      if (!t) return;
      var marca = t.closest('[data-track]');
      var tipo = marca ? marca.getAttribute('data-track') : '';
      if (tipo === 'lead') return; /* el Lead se dispara al enviar el formulario */
      if (tipo === 'agenda') {
        trackEvent('Schedule', { content_name: 'Agenda', button_location: ubicacion(marca) });
        return;
      }
      var wa = t.closest('a[href^="https://wa.me/"], a[href^="https://api.whatsapp.com/"]');
      if (wa) trackEvent('Contact', { content_name: 'WhatsApp', button_location: ubicacion(wa) });
    });

    /* secciones con data-track-view */
    Array.prototype.forEach.call(d.querySelectorAll('[data-track-view]'), function (el) {
      trackView(el, el.getAttribute('data-track-view'));
    });

    /* formularios con data-track-form */
    Array.prototype.forEach.call(d.querySelectorAll('form[data-track-form]'), function (f) {
      trackForm(f);
    });
  }

  if (d.readyState === 'loading') d.addEventListener('DOMContentLoaded', init);
  else init();

  w.MK = {
    enabled: ENABLED,
    trackEvent: trackEvent,
    trackClick: trackClick,
    generateEventId: generateEventId,
    getFbp: getFbp,
    getFbc: getFbc,
    trackView: trackView,
    trackForm: trackForm,
    trackOnce: trackOnce,
    viewContent: viewContent
  };
})(window, document);
