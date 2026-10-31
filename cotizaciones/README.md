# Cotizaciones de FÁRA

Cada cotización es un link: `faraimagen.com.ar/cotizacion/<codigo>`.
Los datos viven en la planilla privada **FÁRA · Cotizaciones** (cuenta fara.imagen@gmail.com).
Para crear o cambiar una cotización no hace falta deploy: editás la planilla y el cliente ve el cambio al recargar.

## Crear una cotización
1. En la planilla: menú **FÁRA → Nueva cotización**. Escribí las 3 letras del cliente.
2. En **Cotizaciones**, completá la fila nueva.
3. Parado en esa fila, tocá **FÁRA → Ver ítems** y completá los renglones. Si elegís un servicio del Catálogo, se cargan solos la descripción y el precio. Para sumar otro renglón: **FÁRA → Agregar ítem**.
4. Cambiá el **Estado** a `Enviada` y mandale al cliente el link de la última columna.

Para el mismo cliente podés crear todas las cotizaciones que quieras: cada una tiene su propio código y su propio link.

En **Precio** va el total de ese renglón. La cantidad es informativa: no multiplica.

## Anticipo y pagos
Dejá **Anticipo %** vacío si no cobrás anticipo. Si te paga otro porcentaje, cambiá **Anticipo %** antes de marcar el pago.

En la columna **Pago** elegí:
- `Pendiente`: se ve normal.
- `Anticipo pagado`: el anticipo sale tachado con el tick de pagado y queda el saldo pendiente.
- `Pagado total`: todo sale pagado y aparece el aviso de propuesta pagada.

Con cualquier pago marcado, la página saca el botón de aceptar y deja de vencer. El link sigue activo como registro hasta que la pases a `Anulada`.

## Cambiar una cotización
Editá la celda en la planilla. Listo. Para volver a ver todos los ítems: **FÁRA → Ver todos los ítems**.

## Sacarla de la web
Cambiá el Estado a `Anulada` o `Pausada`: el cliente ve un mensaje, sin datos ni precios.
No borres filas: quedan como historial.

| Estado | Qué ve el cliente |
|---|---|
| Borrador | "Esta cotización no está disponible" |
| Enviada | La propuesta completa. Si pasa el vencimiento, aparece vencida sola |
| Aceptada | La propuesta completa con el aviso de aceptada |
| Vencida | La propuesta completa y el botón para pedir actualización |
| Pausada / Anulada | Solo un mensaje |

## Qué no tocar
- Columnas con encabezado rojo: son automáticas.
- La columna Clave de la pestaña Config.
- El código de una cotización ya enviada: si cambia, el link deja de funcionar.

## Instalación (una sola vez)
1. Abrí la planilla con la cuenta fara.imagen@gmail.com.
2. **Extensiones → Apps Script**.
3. Borrá lo que aparece, pegá todo `Code.gs` y tocá **Guardar**.
4. Arriba, elegí la función `setup` y tocá **Ejecutar**.
5. Google pide permiso: **Revisar permisos**, elegí la cuenta, **Configuración avanzada**, **Ir a… (no seguro)**, **Permitir**. Ese aviso aparece porque el script es tuyo y no está verificado por Google.
6. **Implementar → Nueva implementación**, engranaje → **Aplicación web**.
7. Ejecutar como: **Yo**. Quién tiene acceso: **Cualquier persona**. Tocá **Implementar**.
8. Copiá la URL que termina en `/exec` y pegala en `API`, en `assets/js/cotizacion.js`.
9. En la planilla, tocá **FÁRA → Activar carga rápida** y aceptá el permiso. Mantiene listas las cotizaciones enviadas para que abran enseguida.

Cada cotización queda guardada en una copia rápida de Google, que se actualiza sola al editar la planilla. Excepción: si borrás una fila entera en vez de usar `Anulada`, la copia puede seguir viéndose hasta 6 horas.

Si más adelante cambia `Code.gs`: **Implementar → Gestionar implementaciones → lápiz → Versión: Nueva versión → Implementar**. La URL no cambia.

## Dónde está cada cosa
- `cotizacion/index.html`: la plantilla.
- `assets/cotizacion.css`: estilos propios de la cotización.
- `assets/js/cotizacion.js`: lee la cotización y la muestra.
- `cotizaciones/Code.gs`: el script de la planilla.
- `netlify.toml`: la regla que manda `/cotizacion/*` a la plantilla.

Este repo es público: acá nunca van datos reales de clientes.
