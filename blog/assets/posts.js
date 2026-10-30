/* ==========================================================
   FÁRA · blog/assets/posts.js
   LISTA ÚNICA DE POSTS. El índice del blog y el bloque
   "Quizás te interese" leen todo desde acá.

   Para sumar un post nuevo:
   1) Copiá la carpeta blog/por-que-te-piden-descuento/ con otro nombre
      (el nombre de la carpeta es el slug) y editá su index.html.
   2) Agregá un objeto nuevo a esta lista, con el mismo slug.
   3) Sumá la dirección nueva a sitemap.xml.
   Campos: slug, title, description, date (AAAA-MM-DD), category, tags (lista).
   Los tags en común entre posts definen cuáles se recomiendan juntos.
   ========================================================== */
window.MK_POSTS = [
  {
    slug: 'por-que-te-piden-descuento',
    title: 'Por qué te piden descuento aunque tu producto sea excelente',
    description: 'Si tus fotos, tu catálogo y tus redes se ven más chicos que tu producto, el cliente duda y pide rebaja. Cómo detectarlo y qué mejorar primero.',
    date: '2026-10-29',
    category: 'Marca y precio',
    tags: ['precio', 'imagen de marca', 'catálogo', 'fotografía de producto']
  }
];
