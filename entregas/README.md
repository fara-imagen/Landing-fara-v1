# Entregas de FÁRA

Cada entrega es un link: `faraimagen.com.ar/entrega/<codigo>`.
Los datos viven en la misma planilla privada **FÁRA · Cotizaciones**, en las pestañas **Entregas** y **Archivos**.
La página no guarda archivos: muestra links a Drive, Pixieset, una web o lo que uses.
Para crear o cambiar una entrega no hace falta deploy: editás la planilla y el cliente ve el cambio al recargar.

## Sola, al aceptar una cotización
Cuando ponés una cotización en `Aceptada`, o marcás un pago, se crea sola su entrega, una sola vez:
- Toma empresa, nombre, título y logo de la cotización.
- Cada ítem de la cotización aparece en Archivos con Tipo `En proceso`. Podés agregar, cambiar u ocultar lo que quieras.
- Queda en Estado `En proceso`. El cliente no ve nada hasta que le mandes el link, que está en la pestaña Entregas.
- Falta solo el **Mail del cliente**, porque la cotización no lo tiene. Sin mail no salen los avisos.

El botón Aceptar de la página solo te manda un WhatsApp: no crea nada.

Una vez creada es independiente: si cambiás la cotización, la entrega no cambia, y si anulás la cotización, la entrega hay que anularla aparte.

## Crear una entrega suelta
1. En la planilla: menú **FÁRA Entregas → Nueva entrega**. Escribí las 3 letras del cliente.
2. En **Entregas**, completá la fila. **Empresa** es lo que sale en la página. **Nombre para el mail** es el saludo: con `Leandro` el mail dice "Hola Leandro,". También podés poner `equipo de Ceniza`. **Mensaje** es opcional: sale arriba, debajo del título, y si queda vacío no se ve nada. Con **Mail del cliente** le llegan los avisos de vencimiento.
3. Parado en esa fila, tocá **FÁRA Entregas → Ver archivos** y cargá un renglón por cada cosa que entregás. Para sumar otro: **FÁRA Entregas → Agregar archivo**.
4. Elegí el **Estado** y mandale al cliente el link de la última columna.

Un renglón con link muestra un botón. Sin link, se ve solo el texto. Podés pegar el link con o sin `https://`. Si la dirección no sirve, la celda se pone en rojo.

## Qué cargar en Archivos

| Tipo | Para qué | ¿Vence? |
|---|---|---|
| Descarga | Drive, Pixieset o cualquier link | Sí |
| Web | Su web o landing | No |
| Entregado | Lo que no se descarga: impresiones 2D/3D, una charla por Zoom. Queda como registro y puede llevar un link, por ejemplo a una foto | No |
| En proceso | Lo que todavía falta. Sale con la etiqueta **En proceso**, sin link. En **Fecha** podés poner una fecha estimada | No |

| Cuándo se ve | Qué pasa |
|---|---|
| Al pagar | Con link, se abre recién en Disponible. Si la entrega viene de una cotización, en `En proceso` se lista con la etiqueta **Se habilita con el pago**. En una entrega suelta, `En proceso` no bloquea: para hacerlo, usá `Pendiente de pago` |
| Entrega parcial | Se abre desde ya, también en Pendiente de pago, bajo el título **Entregas parciales** |
| Oculto | No aparece nunca |

## Entrega parcial antes del pago
1. Dejá la entrega en `Pendiente de pago`.
2. Cargá un renglón con lo que ya le das, por ejemplo una carpeta con 10 fotos o el link de Zoom, y marcalo `Entrega parcial`.
3. Cuando te pague, pasá la entrega a `Disponible`.

Los links de `Al pagar` no llegan al navegador del cliente hasta ese momento. Ni mirando el código de la página los puede ver.

## Trabajo en marcha
Te sirve de lista para vos y para el cliente: ve todo lo acordado y cómo avanza.
1. Mandale el link cuando quieras. La entrega ya está en `En proceso`.
2. Cuando terminás algo, cambiá ese renglón a `Descarga`, `Web` o `Entregado` y pegá el link.
3. Cuando esté todo, pasá la entrega a `Disponible`. Ahí arranca el plazo y salen los avisos.

Mientras la entrega esté en `En proceso` no vence nada y no sale ningún mail.

**Si ya pagó todo:** los renglones vienen en `Entrega parcial` y el link se abre apenas lo pegás. Si marcás `Pagado total` después, cambian solos.

**Si pagó la mitad o nada:** los renglones vienen en `Al pagar`. Aunque pegues el link, el cliente no lo ve hasta el pago. Para darle algo antes, marcá ese renglón `Entrega parcial`.

Con **FÁRA Entregas → Agregar archivo**, el renglón nuevo copia el `Cuándo se ve` del último renglón de esa entrega.

## Estados

| Estado | Qué ve el cliente |
|---|---|
| Borrador | "Esta entrega no está disponible" |
| Pendiente de pago | "Tu trabajo está listo", el aviso de pago y solo las entregas parciales. El plazo no corre |
| En proceso | "Tu trabajo está en marcha" y la lista, salvo lo Oculto. Lo que espera el pago sale sin link. El plazo no corre |
| Disponible | Todo, con la fecha límite y el aviso para descargar |
| Vencida | Se pone sola el día después de la Fecha límite. El mensaje de plazo terminado, con botón de WhatsApp |
| Archivos borrados | Lo mismo que Vencida. Lo ponés vos cuando ya borraste la carpeta |
| Pausada / Anulada | Solo un mensaje |

Si en la entrega no hay nada para descargar, por ejemplo solo una web, no vence, no sale la fecha y no salen avisos.

Al pasar el límite, una entrega `Disponible` vence sola en la web, y a la mañana siguiente su Estado cambia a `Vencida` en la planilla: se ocultan los links de tipo Descarga y aparece el botón de WhatsApp para pedir los archivos. Lo de tipo Web y Entregado sigue a la vista.

La página avisa antes y después de la fecha que recuperar los archivos tiene un costo extra. Ese costo no sale en la página: lo acordás con el cliente.

## Plazo
Al pasar a `Disponible`, si **Fecha límite** está vacía se completa sola con 30 días.
- Para cambiar los 30 días de todas las entregas nuevas: **Config → Entregas: días para descargar**.
- Para cambiar una sola: escribí otra fecha en su fila.
- Para reabrir una vencida: poné una fecha nueva y pasala a `Disponible`. Los avisos se vuelven a activar.

El cliente puede descargar hasta el final del día de la fecha límite, hora de Argentina.

## Avisos por mail
El cliente recibe un mail 5 días antes y otro el último día. Salen solos a la mañana, solo con la entrega en `Disponible`.
- Para cambiar los 5 días: **Config → Entregas: días antes para avisar por mail**. Con `0` no sale ningún aviso.
- Si la dejás con menos de 5 días, el primer aviso sale al día siguiente de cargarla y el último, el último día.
- Las columnas **Aviso previo enviado** y **Aviso último día enviado** muestran cuándo salió cada uno.
- Sin mail en la fila, no sale nada. Si un mail falla, se vuelve a intentar al día siguiente.
- Gmail deja mandar hasta 100 mails por día con esta cuenta.

### Enviar desde hola@faraimagen.com.ar
Gmail solo lo permite si esa dirección ya está en la cuenta fara.imagen@gmail.com, en **Configuración → Cuentas e importación → Enviar correo como**.
Si está, el script la usa sola. Si no está, el mail sale desde fara.imagen@gmail.com y las respuestas igual llegan a hola@faraimagen.com.ar.

## Cuando vence: borrar de verdad
**Ocultar un link en la página no quita el permiso en Drive ni en Pixieset.** Quien lo haya guardado puede seguir entrando.

Por eso la fila `Vencida` queda en **rojo** en la planilla. Cuando la veas:
1. Borrá la carpeta o cerrale el permiso en Drive o Pixieset.
2. Pasá el Estado a `Archivos borrados`. El rojo se apaga.

## Sacarla de la web
Cambiá el Estado a `Anulada` o `Pausada`. Para sacar un solo renglón, ponelo en `Oculto`.
No borres filas: si borrás una fila entera, la copia rápida puede seguir mostrando ese link hasta 6 horas.

## Qué no tocar
- Columnas con encabezado rojo: son automáticas.
- El código de una entrega ya enviada: si cambia, el link deja de funcionar.

## La planilla
Los colores de las pestañas separan cada cosa:
- **Ámbar:** cotizaciones. Cotizaciones, Ítems y Catálogo.
- **Petróleo:** entregas. Entregas y Archivos.
- **Gris:** general. Cómo usar y Config.

Cada sección tiene su menú: **FÁRA Cotizaciones** y **FÁRA Entregas**.

## Dónde está cada cosa
- `entrega/index.html`: la plantilla.
- `assets/entrega.css`: estilos propios de la entrega. El resto sale de `style.css` y `cotizacion.css`.
- `assets/js/entrega.js`: lee la entrega y la muestra.
- `cotizaciones/Code.gs`: el script de la planilla, compartido con las cotizaciones.
- `netlify.toml`: la regla que manda `/entrega/*` a la plantilla.

Este repo es público: acá nunca van datos reales de clientes.
