# Cotizaciones de FÁRA

Cada cotización es un link: `faraimagen.com.ar/cotizacion/<codigo>`.
Los datos viven en la planilla privada **FÁRA · Cotizaciones** (cuenta fara.imagen@gmail.com).
Para crear o cambiar una cotización no hace falta deploy: editás la planilla y el cliente ve el cambio al recargar.

## Crear una cotización
1. En la planilla: menú **FÁRA → Nueva cotización**. Escribí las 3 letras del cliente.
2. En **Cotizaciones**, completá la fila nueva.
3. En **Ítems**, completá los renglones con ese código. Si elegís un servicio del Catálogo, se cargan solos la descripción y el precio. Para sumar otro renglón, copiá el código en una fila libre.
4. Cambiá el **Estado** a `Enviada` y mandale al cliente el link de la última columna.

## Cambiar una cotización
Editá la celda en la planilla. Listo.

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

Si más adelante cambia `Code.gs`: **Implementar → Gestionar implementaciones → lápiz → Versión: Nueva versión → Implementar**. La URL no cambia.

## Dónde está cada cosa
- `cotizacion/index.html`: la plantilla.
- `assets/cotizacion.css`: estilos propios de la cotización.
- `assets/js/cotizacion.js`: lee la cotización y la muestra.
- `cotizaciones/Code.gs`: el script de la planilla.
- `netlify.toml`: la regla que manda `/cotizacion/*` a la plantilla.

Este repo es público: acá nunca van datos reales de clientes.
