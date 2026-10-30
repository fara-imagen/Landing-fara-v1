FÁRA · faraimagen.com.ar
Sitio en HTML, CSS y JS puros. Sin frameworks, sin pasos de compilación.
Se publica solo en Netlify cuando se hace el merge en GitHub.

==========================================================
1) QUÉ MARCADOR REEMPLAZAR Y DÓNDE
==========================================================

El único marcador es TU_PIXEL_ID (el ID de tu Pixel de Meta, son 15 o 16 números).
Mientras diga TU_PIXEL_ID la medición queda apagada y no se envía nada a Meta.

Aparece 2 veces en el <head> de cada una de estas 6 páginas:

  index.html
  404.html
  privacidad/index.html
  gracias/index.html
  blog/index.html
  blog/por-que-te-piden-descuento/index.html
  (y el <head> de cada post nuevo que crees copiando el de ejemplo)

Las 2 apariciones por página son:
  - window.MK_PIXEL_ID='TU_PIXEL_ID';
  - la dirección dentro de <noscript> (...tr?id=TU_PIXEL_ID&ev=PageView...)

Cómo reemplazarlo rápido: en un editor como VS Code, abrí la carpeta del sitio,
usá "Buscar y reemplazar en archivos" (Ctrl+Mayús+H), buscá TU_PIXEL_ID,
poné tu ID y tocá "Reemplazar todo".

==========================================================
2) VARIABLES SECRETAS (van en Netlify, nunca en el código)
==========================================================

Netlify → Site configuration → Environment variables:

  META_PIXEL_ID ........ el mismo ID del Pixel
  META_ACCESS_TOKEN .... token de la API de Conversiones
                         (Events Manager → tu Pixel → Configuración →
                          API de Conversiones → Generar token de acceso)
  META_TEST_EVENT_CODE . OPCIONAL. Código de "Probar eventos". Usalo solo
                         para probar y borralo después.
  META_API_VERSION ..... OPCIONAL. Por defecto v24.0.
  ALLOWED_HOSTS ........ OPCIONAL. Dominios extra permitidos, separados por coma.

Después de cargar o cambiar variables, hacé un nuevo deploy (Deploys → Trigger deploy).
Si faltan META_PIXEL_ID o META_ACCESS_TOKEN, la función responde "sin configurar"
y la web funciona igual.

==========================================================
3) CÓMO SABER QUE LA MEDICIÓN FUNCIONA
==========================================================

1. Events Manager → tu Pixel → "Probar eventos".
2. Cargá META_TEST_EVENT_CODE en Netlify y hacé deploy.
3. Abrí la web, tocá un botón de WhatsApp y llegá a "Cómo trabajo".
4. Tienen que aparecer PageView, Contact y ViewContent, cada uno con
   origen "Navegador" y "Servidor" y marcado como desduplicado.

Eventos del sitio:
  PageView ...... al cargar cualquier página
  ViewContent ... al llegar a "Cómo trabajo" y al abrir un post del blog
  Contact ....... en cualquier botón de WhatsApp (menú, portada, cierre)
  Schedule ...... listo para un botón de agenda (agregale data-track="agenda")
  Lead .......... listo para un formulario (agregale data-track-form) y en /gracias/
                  (en /gracias/ solo cuenta si la persona llegó desde una página del sitio;
                   quien entra directo por el link no suma un Lead falso)

==========================================================
4) CÓMO SUMAR UN POST AL BLOG
==========================================================

1. Copiá la carpeta blog/por-que-te-piden-descuento/ con otro nombre
   (ese nombre es la dirección del post, en minúsculas y con guiones).
2. Editá el index.html de la carpeta nueva: título, descripción, canonical,
   og:url, datos del JSON-LD, texto y el data-post="..." del bloque relacionados.
3. Agregá un objeto nuevo en blog/assets/posts.js con el mismo slug.
   El bloque "Quizás te interese" y el índice del blog se actualizan solos.
4. Sumá la dirección nueva en sitemap.xml.
5. Si el post lleva imágenes, todas necesitan texto alt.

==========================================================
5) QUÉ HAY EN CADA CARPETA
==========================================================

  index.html ...................... la landing (lee textos y fotos de tu planilla de Google)
  404.html ........................ página de error
  gracias/ ........................ confirmación de formulario (noindex, dispara Lead una vez)
  privacidad/ ..................... política de privacidad
  blog/ ........................... índice, lista de posts (assets/posts.js),
                                    bloque relacionados (assets/related.js) y los posts
  assets/style.css ................ sistema de diseño de FÁRA (variables en :root)
  assets/js/tracking.js ........... objeto MK: medición con Pixel + servidor
  netlify/functions/capi.js ....... envía los eventos a la API de Conversiones
  netlify.toml .................... encabezados de seguridad y carpeta de funciones
  sitemap.xml, robots.txt ......... para Google
  llms.txt, ai.txt ................ presentación para asistentes de IA

==========================================================
6) DATOS QUE SE REPITEN EN VARIAS PÁGINAS
==========================================================

Si cambian, hay que actualizarlos en estos archivos:
  - WhatsApp 5493885735415 ........ index.html (respaldo; manda el de la planilla),
                                    gracias/, blog/por-que-te-piden-descuento/, llms.txt, ai.txt
  - Mail hola@faraimagen.com.ar ... privacidad/, llms.txt, ai.txt
  - Dominio faraimagen.com.ar ..... canonical de cada página, sitemap.xml, robots.txt,
                                    llms.txt, ai.txt, capi.js (lista de dominios permitidos)
  - Instagram ..................... llms.txt, ai.txt y la planilla (general.instagram)
  - Precio "desde ARS 150.000" .... llms.txt, ai.txt

==========================================================
7) LO QUE NO SE INCLUYÓ
==========================================================

La calculadora de precios (precios/index.html) no se creó porque hoy no cobrás
por cantidad. Si más adelante la querés, se agrega sin tocar lo demás.

Reglas del sitio: todas las imágenes llevan texto alt y no hay ningún secreto
escrito en el código.
