/* ==========================================================
   FÁRA · netlify/functions/capi.js
   Recibe un evento desde assets/js/tracking.js y lo manda a la
   API de Conversiones de Meta, con el mismo event_id que el Pixel.

   Variables de entorno (Netlify → Site configuration → Environment variables):
     META_PIXEL_ID ........ ID del Pixel (obligatoria)
     META_ACCESS_TOKEN .... token de la API de Conversiones (obligatoria)
     META_TEST_EVENT_CODE . código de "Probar eventos" (opcional, solo para pruebas)
     META_API_VERSION ..... versión de la API (opcional, por defecto v24.0)
     ALLOWED_HOSTS ........ dominios extra permitidos, separados por coma (opcional)

   Ningún secreto está escrito en este archivo.
   ========================================================== */
'use strict';

const VERSION_POR_DEFECTO = 'v24.0';
const EVENTOS_PERMITIDOS = new Set(['PageView', 'ViewContent', 'Contact', 'Schedule', 'Lead']);
const MAX_BODY = 8 * 1024;

function responder(status, cuerpo) {
  return {
    statusCode: status,
    headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' },
    body: JSON.stringify(cuerpo)
  };
}

function hostDe(url) {
  try { return new URL(url).hostname.toLowerCase(); } catch (e) { return ''; }
}

/* solo se aceptan eventos que vienen de tu dominio o de las vistas previas de Netlify */
function hostPermitido(host) {
  if (!host) return false;
  const extra = (process.env.ALLOWED_HOSTS || '').split(',').map((h) => h.trim().toLowerCase()).filter(Boolean);
  return host === 'faraimagen.com.ar' ||
    host.endsWith('.faraimagen.com.ar') ||
    host.endsWith('.netlify.app') ||
    extra.includes(host);
}

/* custom_data: solo valores simples y cortos */
function limpiarCustom(datos) {
  const salida = {};
  if (!datos || typeof datos !== 'object') return salida;
  Object.keys(datos).slice(0, 20).forEach((k) => {
    const v = datos[k];
    if (typeof v === 'string') salida[k.slice(0, 40)] = v.slice(0, 200);
    else if (typeof v === 'number' || typeof v === 'boolean') salida[k.slice(0, 40)] = v;
  });
  return salida;
}

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return responder(405, { ok: false, mensaje: 'Método no permitido' });
  }

  /* sin variables de entorno: responde con éxito y avisa que falta configurar */
  const pixelId = process.env.META_PIXEL_ID;
  const token = process.env.META_ACCESS_TOKEN;
  if (!pixelId || !token) {
    return responder(200, { ok: true, configurado: false, mensaje: 'sin configurar' });
  }

  const h = event.headers || {};
  const origen = hostDe(h.origin || h.referer || '');
  if (!hostPermitido(origen)) {
    return responder(403, { ok: false, mensaje: 'Origen no permitido' });
  }

  let crudo = event.body || '';
  if (event.isBase64Encoded) crudo = Buffer.from(crudo, 'base64').toString('utf8');
  if (crudo.length > MAX_BODY) return responder(413, { ok: false, mensaje: 'Evento demasiado grande' });

  let e;
  try { e = JSON.parse(crudo); } catch (err) { return responder(400, { ok: false, mensaje: 'JSON inválido' }); }

  if (!e || !EVENTOS_PERMITIDOS.has(e.event_name)) {
    return responder(400, { ok: false, mensaje: 'Evento no permitido' });
  }
  if (typeof e.event_id !== 'string' || !/^[A-Za-z0-9_-]{8,64}$/.test(e.event_id)) {
    return responder(400, { ok: false, mensaje: 'event_id inválido' });
  }

  const ip = h['x-nf-client-connection-ip'] || String(h['x-forwarded-for'] || '').split(',')[0].trim();
  const userData = { client_user_agent: h['user-agent'] || '' };
  if (ip) userData.client_ip_address = ip;
  if (typeof e.fbp === 'string' && /^fb\.\d\.\d+\.[\w-]+$/.test(e.fbp)) userData.fbp = e.fbp;
  if (typeof e.fbc === 'string' && /^fb\.\d\.\d+\.[\w-]+$/.test(e.fbc)) userData.fbc = e.fbc;

  const evento = {
    event_name: e.event_name,
    event_time: Math.floor(Date.now() / 1000),
    event_id: e.event_id,
    event_source_url: String(e.event_source_url || '').slice(0, 500),
    action_source: 'website',
    user_data: userData,
    custom_data: limpiarCustom(e.custom_data)
  };

  const paquete = { data: [evento], access_token: token };
  if (process.env.META_TEST_EVENT_CODE) paquete.test_event_code = process.env.META_TEST_EVENT_CODE;

  const version = process.env.META_API_VERSION || VERSION_POR_DEFECTO;
  const url = 'https://graph.facebook.com/' + encodeURIComponent(version) + '/' + encodeURIComponent(pixelId) + '/events';

  const control = new AbortController();
  const timer = setTimeout(() => control.abort(), 8000);
  try {
    const r = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(paquete),
      signal: control.signal
    });
    const resp = await r.json().catch(() => ({}));
    if (!r.ok) {
      /* se registra el motivo, nunca el token */
      console.error('CAPI rechazó el evento', r.status, resp && resp.error && resp.error.message);
      return responder(502, { ok: false, mensaje: 'Meta rechazó el evento' });
    }
    return responder(200, { ok: true, configurado: true, eventos_recibidos: resp.events_received || 0 });
  } catch (err) {
    console.error('CAPI sin respuesta', err && err.name);
    return responder(502, { ok: false, mensaje: 'No se pudo contactar a Meta' });
  } finally {
    clearTimeout(timer);
  }
};
