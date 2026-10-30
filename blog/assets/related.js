/* ==========================================================
   FÁRA · blog/assets/related.js
   Arma el bloque "Quizás te interese" al pie de cada post.
   Uso: en el post, <div id="relacionados" data-post="SLUG"></div>
   y cargar posts.js y después este archivo.
   Criterio: primero los que comparten más tags o categoría, después los más nuevos.
   Si no hay otros posts, el bloque no aparece.
   ========================================================== */
(function () {
  'use strict';
  var posts = window.MK_POSTS || [];
  var box = document.getElementById('relacionados');
  if (!box || !posts.length) return;

  var actual = box.getAttribute('data-post');
  var yo = posts.filter(function (p) { return p.slug === actual; })[0] || { tags: [], category: '' };

  function puntaje(p) {
    var s = 0;
    (p.tags || []).forEach(function (t) { if ((yo.tags || []).indexOf(t) !== -1) s += 2; });
    if (p.category && p.category === yo.category) s += 1;
    return s;
  }

  var otros = posts
    .filter(function (p) { return p.slug !== actual; })
    .map(function (p) { return { p: p, s: puntaje(p) }; })
    .sort(function (a, b) { return b.s - a.s || (a.p.date < b.p.date ? 1 : -1); })
    .slice(0, 3);
  if (!otros.length) return;

  function el(tag, clase, texto) {
    var n = document.createElement(tag);
    if (clase) n.className = clase;
    if (texto) n.textContent = texto;
    return n;
  }

  var bloque = el('aside', 'related');
  bloque.setAttribute('aria-labelledby', 'rel-titulo');
  var h = el('h2', '', 'Quizás te interese');
  h.id = 'rel-titulo';
  bloque.appendChild(h);

  var lista = el('div', 'blog-list');
  otros.forEach(function (o) {
    var a = el('a', 'post-card');
    a.href = '/blog/' + o.p.slug + '/';
    a.appendChild(el('span', 'eyebrow', o.p.category || 'Blog'));
    a.appendChild(el('h3', '', o.p.title));
    a.appendChild(el('p', '', o.p.description));
    var pie = el('span', 'foot');
    pie.appendChild(el('span', 'meta', o.p.category || 'Blog'));
    var mas = el('span', 'more', 'Leer ');
    mas.insertAdjacentHTML('beforeend', '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 12h14M13 6l6 6-6 6"/></svg>');
    pie.appendChild(mas);
    a.appendChild(pie);
    lista.appendChild(a);
  });
  bloque.appendChild(lista);
  box.appendChild(bloque);
})();
