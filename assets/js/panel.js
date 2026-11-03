/* ==========================================================
   FÁRA · assets/js/panel.js
   Panel privado: habla con el script de la planilla "FÁRA · Cotizaciones".
   La clave se pide una vez y queda solo en este navegador. Sin ella el
   script no entrega nada: esta página, sola, no tiene ningún dato.
   Todo el contenido se escribe como texto (nunca como HTML).
   ========================================================== */
(function (w, d) {
  'use strict';

  /* API: la misma dirección /exec de assets/js/cotizacion.js */
  const API = 'https://script.google.com/macros/s/AKfycbxkYwQ-kbpUrgIhF8RF58sCLI0ZEj-TtX3VcdI8tccmWyJNjqBTI6rKr6DZVTwiBxC47A/exec';
  const GUARDADA = 'fara-panel-clave';
  const ESPERA_MAX = 30000;
  const ESPERA_FOTO = 90000;
  const FOTO_LADO = 1600; /* píxeles del lado más largo */
  const FOTO_CALIDAD = 0.8;
  const RE_CODIGO = /^[A-Z]{3}-[a-z2-9]{6,8}$/;

  const raiz = d.getElementById('pan');
  const capaHoja = d.getElementById('panHoja');
  const pestanas = d.getElementById('panPestanas');
  const salir = d.getElementById('panSalir');
  const tabCot = d.getElementById('panTabCot');
  const tabEnt = d.getElementById('panTabEnt');
  const aviso = d.getElementById('panAviso');
  const velo = d.getElementById('panVelo');

  let clave = '';
  let D = null;
  let pestana = 'cot';
  let abierta = '';
  const renglones = {};
  const borradores = {}; // links escritos sin guardar: sobreviven a cada repintado

  /* ---------- utilidades ---------- */

  /* crea un elemento: h('p',{class:'x',onclick:fn},'texto',otroElemento,[lista]) */
  function h(tag, attrs) {
    const el = d.createElement(tag);
    Object.keys(attrs || {}).forEach(function (k) {
      const v = attrs[k];
      if (k.indexOf('on') === 0) el.addEventListener(k.slice(2), v);
      else if (k === 'value') el.value = v;
      else if (v !== false && v !== null && v !== undefined) el.setAttribute(k, v);
    });
    Array.prototype.slice.call(arguments, 2).forEach(function (hijos) {
      [].concat(hijos).forEach(function (c) {
        if (c === null || c === undefined || c === false) return;
        el.append(c);
      });
    });
    return el;
  }

  const plata = function (n, moneda) { return (moneda === 'USD' ? 'US$ ' : '$ ') + Number(n || 0).toLocaleString('es-AR'); };
  const decimal = function (v) { return String(v).replace(',', '.').trim(); };

  function fechaCorta(iso) {
    const p = String(iso || '').split('-');
    return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : '';
  }

  /* el almacenamiento puede estar bloqueado: sin él, la clave dura hasta cerrar la pestaña */
  function leerClave() {
    try { return w.localStorage.getItem(GUARDADA) || ''; } catch (e) { return ''; }
  }
  function guardarClave(v) {
    try {
      if (v) w.localStorage.setItem(GUARDADA, v);
      else w.localStorage.removeItem(GUARDADA);
    } catch (e) { /* sigue sin recordarla */ }
  }

  /* ---------- servidor ---------- */

  function srv(accion) {
    const args = Array.prototype.slice.call(arguments, 1);
    const corte = new AbortController();
    const reloj = setTimeout(function () { corte.abort(); }, accion === 'subirFoto' ? ESPERA_FOTO : ESPERA_MAX);
    /* text/plain evita la consulta previa del navegador, que Apps Script no responde */
    return fetch(API, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ clave: clave, accion: accion, args: args }),
      credentials: 'omit',
      signal: corte.signal
    }).then(function (r) { return r.json(); }, function () {
      throw new Error('No pude conectar. Probá de nuevo.');
    }).then(function (r) {
      clearTimeout(reloj);
      if (r.ok) return r.datos;
      if (r.clave || r.bloqueado) {
        const motivo = new Error(r.bloqueado ? 'Demasiados intentos. Esperá 15 minutos.' : 'Esa clave no es correcta.');
        motivo.entrada = true;
        throw motivo;
      }
      throw new Error(r.error || 'No se pudo hacer.');
    }).finally(function () { clearTimeout(reloj); });
  }

  let reloj;
  function avisar(texto, mal) {
    aviso.textContent = texto;
    aviso.className = 'pan-aviso' + (mal ? ' pan-aviso-mal' : '');
    aviso.hidden = false;
    clearTimeout(reloj);
    reloj = setTimeout(function () { aviso.hidden = true; }, mal ? 6000 : 2500);
  }

  function hacer(promesa, listo) {
    velo.hidden = false;
    return promesa.then(function (r) {
      velo.hidden = true;
      if (listo) listo(r);
    }, function (e) {
      velo.hidden = true;
      if (e.entrada) pedirClave(e.message);
      else avisar(e.message, true);
    });
  }

  function copiar(texto) {
    function aMano() {
      const t = h('textarea', { value: texto });
      d.body.append(t);
      t.select();
      try { d.execCommand('copy'); avisar('Link copiado'); } catch (e) { avisar('No pude copiar', true); }
      t.remove();
    }
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(texto).then(function () { avisar('Link copiado'); }, aMano);
    } else aMano();
  }

  /* ---------- entrada ---------- */

  function pedirClave(error) {
    guardarClave('');
    clave = '';
    D = null;
    capaHoja.textContent = '';
    pestanas.hidden = true;
    salir.hidden = true;
    d.body.className = '';
    raiz.removeAttribute('aria-busy');
    const campo = h('input', {
      id: 'panClave', type: 'password', autocomplete: 'current-password', autocapitalize: 'none',
      spellcheck: 'false', required: true
    });
    const mensaje = h('p', { class: 'pan-error', role: 'alert' }, error || '');
    const entrar = function (ev) {
      ev.preventDefault();
      const escrita = campo.value.trim().toLowerCase();
      if (!escrita) return;
      clave = escrita;
      velo.hidden = false;
      srv('datos').then(function (r) {
        velo.hidden = true;
        guardarClave(clave);
        D = r;
        arrancar();
      }, function (e) {
        velo.hidden = true;
        clave = '';
        mensaje.textContent = e.message;
        campo.select();
      });
    };
    raiz.textContent = '';
    raiz.append(h('section', { class: 'post-hero pan-entrada' },
      h('span', { class: 'eyebrow' }, 'Panel privado'),
      h('h1', {}, 'Entrá con tu clave'),
      h('p', { class: 'meta' }, 'La pegás una sola vez: este celular la recuerda.'),
      h('form', { onsubmit: entrar },
        h('label', { for: 'panClave' }, 'Clave'),
        campo,
        mensaje,
        h('button', { class: 'btn btn-accent', type: 'submit' }, 'Entrar'))));
    campo.focus();
  }

  function sinConexion(texto) {
    raiz.removeAttribute('aria-busy');
    raiz.textContent = '';
    raiz.append(h('section', { class: 'post-hero pan-entrada' },
      h('h1', {}, 'No pude abrir el panel'),
      h('p', { class: 'meta' }, texto),
      h('form', { onsubmit: function (ev) { ev.preventDefault(); w.location.reload(); } },
        h('button', { class: 'btn btn-accent', type: 'submit' }, 'Probar de nuevo'))));
  }

  /* ---------- pantalla ---------- */

  function fichas(opciones, actual, elegir) {
    return h('div', { class: 'pan-fichas' }, opciones.map(function (o) {
      return h('button', {
        class: 'pan-ficha', type: 'button', 'aria-pressed': String(o === actual),
        onclick: function () { if (o !== actual) elegir(o); }
      }, o);
    }));
  }

  function marca(texto, tono) {
    return h('span', { class: 'pan-marca' + (tono ? ' pan-marca-' + tono : '') }, texto);
  }

  function cabecera(eyebrow, titulo, boton) {
    return h('section', { class: 'post-hero pan-cabecera' },
      h('span', { class: 'eyebrow' }, eyebrow),
      h('h1', {}, titulo),
      boton);
  }

  function cuantas(n, una, varias) { return n + ' ' + (n === 1 ? una : varias); }

  function botonesLink(codigo, contacto, esEntrega) {
    if (!RE_CODIGO.test(codigo)) return null;
    const link = (esEntrega ? D.sitioEntrega : D.sitio) + codigo;
    const texto = 'Hola' + (contacto ? ' ' + contacto : '') + ', ' + (esEntrega ? 'acá está tu entrega: ' : 'te paso tu cotización: ') + link;
    return h('div', { class: 'pan-acciones' },
      h('button', { class: 'btn pan-btn', type: 'button', onclick: function () { copiar(link); } }, 'Copiar link'),
      h('a', { class: 'btn pan-btn-linea', target: '_blank', rel: 'noopener', href: 'https://wa.me/?text=' + encodeURIComponent(texto) }, 'Mandar por WhatsApp'));
  }

  function tarjeta(codigo, titulo, detalle, datos, cuerpo, alAbrir) {
    const abre = abierta === codigo;
    return h('article', { class: 'pan-tarjeta' },
      h('button', {
        class: 'pan-cab', type: 'button', 'aria-expanded': String(abre),
        onclick: function () {
          abierta = abre ? '' : codigo;
          if (!abre && alAbrir) alAbrir();
          pintar();
        }
      },
        h('h2', {}, titulo || 'Sin empresa'),
        h('small', {}, detalle || 'Sin título'),
        h('span', { class: 'pan-datos' }, datos)),
      abre ? cuerpo() : null);
  }

  /* cotizaciones */
  function pintarCot() {
    raiz.append(cabecera('Cotizaciones', cuantas(D.cotizaciones.length, 'cotización', 'cotizaciones'),
      h('button', { class: 'btn btn-accent', type: 'button', onclick: function () { abrirFormulario(); } }, 'Nueva cotización')));
    if (!D.cotizaciones.length) raiz.append(h('p', { class: 'pan-vacio' }, 'Todavía no hay cotizaciones.'));
    raiz.append(h('div', { class: 'pan-lista' }, D.cotizaciones.map(function (c) {
      return tarjeta(c.codigo, c.empresa, c.titulo, [
        marca(c.estado || 'Sin estado', c.estado === 'Enviada' || c.estado === 'Aceptada' ? 'ok' : ''),
        marca(c.pago || 'Pendiente'),
        h('span', { class: 'pan-total' }, plata(c.total, c.moneda))
      ], function () {
        return h('div', { class: 'pan-cuerpo' },
          h('h3', {}, 'Estado'),
          fichas(D.opciones.estadosCot, c.estado, function (v) { cambiarCot(c.codigo, 'estado', v); }),
          h('h3', {}, 'Pago'),
          fichas(D.opciones.pagos, c.pago, function (v) { cambiarCot(c.codigo, 'pago', v); }),
          botonesLink(c.codigo, c.contacto, false),
          h('div', { class: 'pan-acciones' },
            h('button', {
              class: 'btn pan-btn', type: 'button',
              onclick: function () { hacer(srv('cotizacion', c.codigo), abrirFormulario); }
            }, 'Editar cotización')),
          campoMail(c.codigo, c.mail, function (v, listo) {
            hacer(srv('cambiarCotizacion', c.codigo, 'mail', v), function (r) {
              listo();
              D = r;
              pintar();
              avisar('Mail guardado');
            });
          }));
      });
    })));
  }

  function cambiarCot(codigo, campo, valor) {
    hacer(srv('cambiarCotizacion', codigo, campo, valor), function (r) {
      D = r;
      pintar();
      avisar('Listo');
    });
  }

  /* el borrador sobrevive a cada repintado, como los links */
  function campoMail(codigo, actual, guardar) {
    const id = 'mail' + codigo;
    const campo = h('input', {
      type: 'email', inputmode: 'email', autocomplete: 'off', placeholder: 'nombre@empresa.com', 'aria-label': 'Mail del cliente',
      value: id in borradores ? borradores[id] : actual,
      oninput: function (ev) { borradores[id] = ev.target.value; }
    });
    return [
      h('h3', {}, 'Mail del cliente'),
      h('div', { class: 'pan-fichas' }, campo),
      h('div', { class: 'pan-acciones' },
        h('button', {
          class: 'btn pan-btn-linea', type: 'button',
          onclick: function () { guardar(campo.value, function () { delete borradores[id]; }); }
        }, 'Guardar mail'))
    ];
  }

  /* sin datos crea una cotización nueva; con datos edita esa */
  function abrirFormulario(actual) {
    aviso.hidden = true;
    const f = {};
    const elegidos = []; /* del Catálogo y nuevos: cada uno sabe leerse y se quita solo */
    const total = h('p', {});
    function campo(id, etiqueta, attrs, area) {
      const domId = 'panN' + id;
      f[id] = h(area ? 'textarea' : 'input', Object.assign({ id: domId, autocomplete: 'off' }, attrs, { oninput: sumar }));
      return h('div', { class: 'pan-campo' }, h('label', { for: domId }, etiqueta), f[id]);
    }
    function sumar() {
      const activos = elegidos.filter(function (e) { return e.activo(); });
      let suma = activos.reduce(function (t, e) { return t + (Number(decimal(e.precio.value)) || 0); }, 0);
      const descuento = Number(decimal(f.descuentoValor ? f.descuentoValor.value : '')) || 0;
      if (descuento && f.descuentoTipo.value === '%') suma = suma * (1 - descuento / 100);
      else if (descuento) suma = suma - descuento;
      total.textContent = cuantas(activos.length, 'ítem', 'ítems') + ' · ' + plata(suma, f.moneda ? f.moneda.value : 'ARS');
    }
    /* descripción, entregables, cantidad y unidad: lo mismo que la pestaña Ítems */
    function detalleItem(id, descripcion) {
      const d2 = {};
      const c2 = function (k, etiqueta, attrs, area) {
        d2[k] = h(area ? 'textarea' : 'input', Object.assign({ id: id + k, autocomplete: 'off' }, attrs));
        return h('div', { class: 'pan-campo' }, h('label', { for: id + k }, etiqueta), d2[k]);
      };
      return {
        campos: d2,
        nodos: [
          c2('descripcion', 'Descripción', { value: descripcion || '' }, true),
          c2('entregables', 'Entregables (uno por renglón)', {}, true),
          h('div', { class: 'pan-doble' },
            c2('cantidad', 'Cantidad', { inputmode: 'decimal' }),
            c2('unidad', 'Unidad', { maxlength: '60', placeholder: 'fotos' }))
        ]
      };
    }
    function leerItem(e, servicio) {
      return {
        servicio: servicio, precio: decimal(e.precio.value), descripcion: e.campos.descripcion.value, entregables: e.campos.entregables.value,
        cantidad: decimal(e.campos.cantidad.value), unidad: e.campos.unidad.value
      };
    }
    const usados = actual ? actual.items.map(function (it) { return it.servicio; }) : [];
    const lista = D.catalogo.filter(function (c) { return usados.indexOf(c.servicio) < 0; }).map(function (c, i) {
      const id = 'panS' + i;
      const detalle = detalleItem(id, c.descripcion);
      const caja = h('div', { class: 'pan-item-detalle', hidden: true }, detalle.nodos);
      const tilde = h('input', {
        id: id, type: 'checkbox',
        onchange: function () { caja.hidden = !tilde.checked; sumar(); }
      });
      const precio = h('input', {
        type: 'text', inputmode: 'decimal', placeholder: 'Precio', value: c.precio === '' ? '' : String(c.precio),
        'aria-label': 'Precio de ' + c.servicio, oninput: sumar
      });
      const item = { precio: precio, campos: detalle.campos, activo: function () { return tilde.checked; }, leer: function () { return leerItem(item, c.servicio); } };
      elegidos.push(item);
      return h('div', { class: 'pan-item' }, h('div', { class: 'pan-servicio' }, tilde, h('label', { for: id }, c.servicio), precio), caja);
    });
    const previos = h('div', {}); /* lo que ya tiene arriba; lo nuevo, debajo del Catálogo y cerca del botón */
    const nuevos = h('div', {});
    let cuenta = 0;
    function sumarNuevo(previo) {
      const id = 'panI' + (cuenta++);
      const detalle = detalleItem(id, previo ? previo.descripcion : '');
      const nombre = h('input', { id: id + 'n', maxlength: '120', autocomplete: 'off', placeholder: 'Ej: Video para redes', value: previo ? previo.servicio : '' });
      const precio = h('input', { id: id + 'p', type: 'text', inputmode: 'decimal', autocomplete: 'off', oninput: sumar, value: previo ? previo.precio : '' });
      const alCatalogo = h('input', { id: id + 'c', type: 'checkbox' });
      if (previo) {
        detalle.campos.entregables.value = previo.entregables;
        detalle.campos.cantidad.value = previo.cantidad;
        detalle.campos.unidad.value = previo.unidad;
      }
      const item = {
        previo: !!previo, precio: precio, campos: detalle.campos, activo: function () { return true; },
        leer: function () {
          const leido = Object.assign(leerItem(item, nombre.value), { alCatalogo: alCatalogo.checked });
          if (previo) { leido.descuentoPct = previo.descuentoPct; leido.mostrar = previo.mostrar; }
          return leido;
        }
      };
      const caja = h('div', { class: 'pan-item pan-item-nuevo' },
        h('div', { class: 'pan-campo' }, h('label', { for: id + 'n' }, 'Nombre del ítem'), nombre),
        h('div', { class: 'pan-campo' }, h('label', { for: id + 'p' }, 'Precio'), precio),
        detalle.nodos,
        previo ? null : h('label', { class: 'pan-tilde', for: id + 'c' }, alCatalogo, 'Guardar también en el Catálogo'),
        h('button', {
          class: 'pan-quitar', type: 'button',
          onclick: function () { elegidos.splice(elegidos.indexOf(item), 1); caja.remove(); sumar(); }
        }, 'Quitar ítem'));
      elegidos.push(item);
      (previo ? previos : nuevos).append(caja);
      sumar();
      if (!previo) nombre.focus();
    }
    f.moneda = h('select', { id: 'panNmoneda', onchange: sumar }, h('option', { value: 'ARS' }, 'ARS'), h('option', { value: 'USD' }, 'USD'));
    f.descuentoTipo = h('select', { id: 'panNdescuentoTipo', onchange: sumar }, h('option', { value: '%' }, '%'), h('option', { value: 'Monto' }, 'Monto'));
    const cerrar = function () { capaHoja.textContent = ''; };
    const crear = function () {
      const datos = {
        prefijo: f.prefijo.value, empresa: f.empresa.value, contacto: f.contacto.value, mail: f.mail.value, titulo: f.titulo.value,
        items: elegidos.filter(function (e) { return e.activo(); })
          .sort(function (a, b) { return Number(!!b.previo) - Number(!!a.previo); }) /* lo que ya tenía conserva su orden, arriba */
          .map(function (e) { return e.leer(); }),
        extra: {
          intro: f.intro.value, moneda: f.moneda.value, descuentoTipo: f.descuentoTipo.value, descuentoValor: decimal(f.descuentoValor.value),
          descuentoMotivo: f.descuentoMotivo.value, formaPago: f.formaPago.value, anticipoPct: decimal(f.anticipoPct.value),
          plazoEntrega: f.plazoEntrega.value, condiciones: f.condiciones.value, notas: f.notas.value,
          vencimiento: f.vencimiento ? f.vencimiento.value : ''
        }
      };
      hacer(actual ? srv('guardarCotizacion', actual.codigo, datos) : srv('nuevaCotizacion', datos), function (r) {
        D = r;
        pestana = 'cot';
        abierta = actual ? actual.codigo : (D.cotizaciones[0] ? D.cotizaciones[0].codigo : '');
        cerrar();
        pintar();
        avisar(actual ? 'Cotización guardada' : 'Cotización creada en Borrador');
      });
    };
    capaHoja.textContent = '';
    capaHoja.append(h('div', { class: 'pan-hoja', role: 'dialog', 'aria-modal': 'true', 'aria-label': actual ? 'Editar cotización' : 'Nueva cotización' },
      h('div', { class: 'pan-barra' },
        h('h1', {}, actual ? 'Editar cotización' : 'Nueva cotización'),
        h('button', { type: 'button', onclick: cerrar }, 'Cerrar')),
      h('div', { class: 'pan-interior' },
        actual ? h('p', { class: 'pan-nota' }, 'El cliente ve los cambios en su mismo link apenas guardás. Si ya tiene entrega, esa no cambia.') : null,
        h('section', { class: 'pan-bloque' },
          h('h2', {}, 'Cliente'),
          h('div', { class: 'pan-doble pan-doble-letras' },
            campo('prefijo', '3 letras', { maxlength: '3', autocapitalize: 'characters', placeholder: 'CEN' }),
            campo('empresa', 'Empresa', { maxlength: '120' })),
          campo('contacto', 'Nombre para el mail', { maxlength: '120' }),
          campo('mail', 'Mail del cliente', { type: 'email', inputmode: 'email', maxlength: '200', placeholder: 'nombre@empresa.com' })),
        h('section', { class: 'pan-bloque' },
          h('h2', {}, 'Propuesta'),
          campo('titulo', 'Título', { maxlength: '200' }),
          campo('intro', 'Mensaje de introducción', {}, true)),
        h('section', { class: 'pan-bloque' },
          h('h2', {}, 'Qué cotizás'),
          h('p', { class: 'pan-ayuda' }, 'Los precios y textos que cambies acá valen solo para esta cotización.'),
          previos,
          lista.length || actual ? lista : h('p', { class: 'pan-vacio' }, 'El Catálogo está vacío.'),
          nuevos,
          h('button', { class: 'btn pan-btn-linea pan-sumar', type: 'button', onclick: function () { sumarNuevo(); } }, 'Agregar ítem nuevo')),
        h('section', { class: 'pan-bloque' },
          h('h2', {}, 'Condiciones'),
          campo('formaPago', 'Forma de pago', { maxlength: '300' }),
          h('div', { class: 'pan-doble' },
            campo('anticipoPct', 'Anticipo %', { inputmode: 'decimal' }),
            campo('plazoEntrega', 'Plazo de entrega', { maxlength: '200' })),
          h('div', { class: 'pan-doble' },
            h('div', { class: 'pan-campo' }, h('label', { for: 'panNdescuentoTipo' }, 'Tipo de descuento'), f.descuentoTipo),
            campo('descuentoValor', 'Descuento', { inputmode: 'decimal' })),
          campo('descuentoMotivo', 'Motivo del descuento', { maxlength: '200' }),
          h('div', { class: 'pan-campo' }, h('label', { for: 'panNmoneda' }, 'Moneda'), f.moneda),
          actual ? campo('vencimiento', 'Válida hasta', { type: 'date' }) : null,
          campo('condiciones', 'Condiciones', {}, true),
          campo('notas', 'Notas', {}, true))),
      h('div', { class: 'pan-pie' }, total, h('button', { class: 'btn btn-accent', type: 'button', onclick: crear }, actual ? 'Guardar' : 'Crear'))));
    if (actual) {
      f.prefijo.value = actual.codigo.slice(0, 3);
      f.prefijo.disabled = true;
      ['empresa', 'contacto', 'mail', 'titulo', 'intro', 'formaPago', 'anticipoPct', 'plazoEntrega', 'descuentoValor', 'descuentoMotivo', 'condiciones', 'notas', 'vencimiento', 'moneda', 'descuentoTipo']
        .forEach(function (k) { f[k].value = actual[k]; });
      actual.items.forEach(function (it) { sumarNuevo(it); });
    }
    sumar();
    if (!actual) f.prefijo.focus();
  }

  /* entregas */
  function pintarEnt() {
    raiz.append(cabecera('Entregas', cuantas(D.entregas.length, 'entrega', 'entregas')));
    if (!D.entregas.length) raiz.append(h('p', { class: 'pan-vacio' }, 'Todavía no hay entregas.'));
    raiz.append(h('div', { class: 'pan-lista' }, D.entregas.map(function (e) {
      return tarjeta(e.codigo, e.empresa, e.titulo, [
        marca(e.estado || 'Sin estado', e.estado === 'Disponible' ? 'ok' : e.estado === 'Vencida' ? 'mal' : ''),
        e.limite ? marca('Hasta ' + fechaCorta(e.limite)) : null
      ], function () { return cuerpoEntrega(e); }, function () { cargarRenglones(e.codigo); });
    })));
  }

  function cargarRenglones(codigo) {
    srv('renglones', codigo).then(function (r) {
      renglones[codigo] = r;
      pintar();
    }, function (e) {
      if (e.entrada) pedirClave(e.message);
      else avisar(e.message, true);
    });
  }

  function cuerpoEntrega(e) {
    const filas = renglones[e.codigo];
    const cuerpo = h('div', { class: 'pan-cuerpo' },
      h('h3', {}, 'Estado'),
      fichas(D.opciones.estadosEnt, e.estado, function (v) {
        hacer(srv('cambiarEntrega', e.codigo, 'estado', v), function (r) {
          D = r;
          pintar();
          avisar('Listo');
        });
      }),
      botonesLink(e.codigo, e.contacto, true),
      campoMail(e.codigo, e.mail, function (v, listo) {
        hacer(srv('cambiarEntrega', e.codigo, 'mail', v), function (r) {
          listo();
          D = r;
          pintar();
          avisar('Mail guardado');
        });
      }),
      h('h3', {}, 'Renglones'));
    if (!filas) {
      cuerpo.append(h('p', { class: 'pan-vacio' }, 'Cargando…'));
      return cuerpo;
    }
    if (!filas.length) cuerpo.append(h('p', { class: 'pan-vacio' }, 'Sin renglones todavía.'));
    filas.forEach(function (r) { cuerpo.append(renglon(e.codigo, r)); });
    cuerpo.append(formRenglon(e.codigo));
    return cuerpo;
  }

  function editarArchivo(codigo, orden, campo, valor) {
    hacer(srv('editarArchivo', codigo, orden, campo, valor), function (r) {
      if (campo === 'link') delete borradores[codigo + orden];
      renglones[codigo] = r;
      pintar();
      avisar('Listo');
    });
  }

  /* la foto se achica y se pasa a WebP acá, antes de subir, para que pese poco */
  function comprimir(archivo) {
    return createImageBitmap(archivo).then(function (foto) {
      const escala = Math.min(1, FOTO_LADO / Math.max(foto.width, foto.height));
      const lienzo = d.createElement('canvas');
      lienzo.width = Math.round(foto.width * escala);
      lienzo.height = Math.round(foto.height * escala);
      lienzo.getContext('2d').drawImage(foto, 0, 0, lienzo.width, lienzo.height);
      const aBlob = function (tipo) {
        return new Promise(function (ok) { lienzo.toBlob(ok, tipo, FOTO_CALIDAD); });
      };
      /* algunos navegadores no escriben WebP: ahí sale JPG */
      return aBlob('image/webp').then(function (b) { return b && b.type === 'image/webp' ? b : aBlob('image/jpeg'); });
    }).then(function (blob) {
      return new Promise(function (ok, mal) {
        const lector = new FileReader();
        lector.onload = function () { ok({ tipo: blob.type, base64: String(lector.result).split(',')[1] }); };
        lector.onerror = mal;
        lector.readAsDataURL(blob);
      });
    }).catch(function () {
      throw new Error('No pude leer esa foto.');
    });
  }

  let fotoPara = null;
  const entradaFoto = h('input', {
    type: 'file', accept: 'image/*', hidden: true,
    onchange: function () {
      const archivo = entradaFoto.files[0];
      const destino = fotoPara;
      entradaFoto.value = '';
      if (!archivo || !destino) return;
      hacer(comprimir(archivo).then(function (foto) {
        return srv('subirFoto', destino.codigo, destino.orden, foto);
      }), function (r) {
        delete borradores[destino.codigo + destino.orden];
        renglones[destino.codigo] = r;
        pintar();
        avisar('Foto guardada');
      });
    }
  });
  d.body.append(entradaFoto);

  function renglon(codigo, r) {
    const id = codigo + r.orden;
    const link = h('input', {
      type: 'url', inputmode: 'url', placeholder: 'Pegá el link', 'aria-label': 'Link de ' + (r.nombre || 'este renglón'),
      value: id in borradores ? borradores[id] : r.link,
      oninput: function (ev) { borradores[id] = ev.target.value; }
    });
    return h('div', { class: 'pan-renglon' },
      h('strong', {}, r.nombre || 'Sin nombre'),
      r.detalle ? h('small', {}, r.detalle) : null,
      h('h3', {}, 'Tipo'),
      fichas(D.opciones.tipos, r.tipo, function (v) { editarArchivo(codigo, r.orden, 'tipo', v); }),
      h('h3', {}, 'Cuándo se ve'),
      fichas(D.opciones.visibles, r.visible, function (v) { editarArchivo(codigo, r.orden, 'visible', v); }),
      h('h3', {}, 'Link'),
      h('div', { class: 'pan-fichas' }, link),
      h('div', { class: 'pan-acciones' },
        h('button', { class: 'btn pan-btn-linea', type: 'button', onclick: function () { editarArchivo(codigo, r.orden, 'link', link.value); } }, 'Guardar link'),
        h('button', {
          class: 'btn pan-btn', type: 'button',
          onclick: function () {
            fotoPara = { codigo: codigo, orden: r.orden };
            entradaFoto.click();
          }
        }, 'Agregar foto')));
  }

  function formRenglon(codigo) {
    const base = 'panR' + codigo;
    const nombre = h('input', { id: base + 'n', maxlength: '200' });
    const tipo = h('select', { id: base + 't' }, D.opciones.tipos.map(function (o) { return h('option', { value: o }, o); }));
    const link = h('input', { id: base + 'l', type: 'url', inputmode: 'url', placeholder: 'Link (opcional)' });
    const visible = h('select', { id: base + 'v' }, D.opciones.visibles.map(function (o) { return h('option', { value: o }, o); }));
    const detalle = h('input', { id: base + 'd', maxlength: '300' });
    return h('details', {},
      h('summary', {}, 'Agregar renglón'),
      h('label', { for: base + 'n' }, 'Qué entregás'), nombre,
      h('label', { for: base + 't' }, 'Tipo'), tipo,
      h('label', { for: base + 'l' }, 'Link'), link,
      h('label', { for: base + 'v' }, 'Cuándo se ve'), visible,
      h('label', { for: base + 'd' }, 'Detalle'), detalle,
      h('div', { class: 'pan-acciones' },
        h('button', {
          class: 'btn pan-btn', type: 'button',
          onclick: function () {
            hacer(srv('sumarArchivo', codigo, { nombre: nombre.value, tipo: tipo.value, visible: visible.value, link: link.value, detalle: detalle.value }), function (r) {
              renglones[codigo] = r;
              pintar();
              avisar('Renglón agregado');
            });
          }
        }, 'Agregar')));
  }

  /* ---------- general ---------- */

  function pintar() {
    if (!D) return;
    d.body.className = pestana === 'ent' ? 'pan-ent' : '';
    tabCot.setAttribute('aria-pressed', String(pestana === 'cot'));
    tabEnt.setAttribute('aria-pressed', String(pestana === 'ent'));
    /* el repintado no mueve la pantalla: un cambio no te manda arriba de todo */
    const x = w.scrollX;
    const y = w.scrollY;
    raiz.textContent = '';
    if (pestana === 'cot') pintarCot();
    else pintarEnt();
    w.scrollTo(x, y);
  }

  function cambiarPestana(p) {
    pestana = p;
    abierta = '';
    pintar();
    w.scrollTo(0, 0);
  }

  function arrancar() {
    raiz.removeAttribute('aria-busy');
    pestanas.hidden = false;
    salir.hidden = false;
    pintar();
  }

  tabCot.addEventListener('click', function () { cambiarPestana('cot'); });
  tabEnt.addEventListener('click', function () { cambiarPestana('ent'); });
  salir.addEventListener('click', function () {
    if (w.confirm('¿Salir? Para volver vas a necesitar la clave.')) pedirClave();
  });

  clave = leerClave();
  if (!clave) {
    pedirClave();
    return;
  }
  srv('datos').then(function (r) {
    D = r;
    arrancar();
  }, function (e) {
    if (e.entrada) pedirClave(e.message);
    else sinConexion(e.message);
  });
})(window, document);
