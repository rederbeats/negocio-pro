# NEGOCIO PRO — arquitectura de datos (v0.1)

## Alcance y grano

La primera versión contiene Dashboard, Ventas, Gastos, Productos y Configuración. Cada fila de Ventas representa **una línea de venta**; varias líneas de una factura pueden compartir referencia documental, pero cada línea tiene un ID único. Cada fila de Gastos representa un gasto. Los ID de producto son únicos y estables. Se reservan ID de cliente y proveedor para las futuras hojas, sin fingir que hoy exista integridad referencial con ellas.

## Relaciones y dependencias

```text
Configuración ──> listas de categorías y tipos IVA (validación)
Productos ──────> Ventas (ID de producto; validación de existencia)
Ventas ─────────> Dashboard (ventas netas, coste estimado y evolución)
Gastos ─────────> Dashboard (gastos netos y evolución)
```

Ventas conserva precio unitario, coste unitario y tipo de IVA **como datos de la operación**. El catálogo no altera ventas históricas al cambiar un precio o coste. La validación avisa de un ID de producto inexistente; no se calcula automáticamente un precio de catálogo para evitar cambios retroactivos.

## Definiciones financieras

- Precio unitario, coste unitario e importe de gasto se introducen **sin IVA**. Descuento es un porcentaje de 0 a 100 %.
- Base venta = cantidad × precio × (1 − descuento). IVA repercutido = base × tipo. Total factura = base + IVA.
- Coste de producto estimado = cantidad × coste unitario de la línea; beneficio operativo estimado = ventas netas − coste de producto estimado − gastos netos. No es beneficio contable ni fiscal.
- IVA soportado de gastos se muestra aparte. Ningún IVA se incluye en ingresos, costes o beneficio.
- Dashboard y resumen mensual usan fecha de operación, no fecha de cobro o pago. La futura Tesorería mostrará flujos de caja sin alimentar el resultado operativo.
- No se modelan todavía devoluciones, abonos, impagos, existencias, amortizaciones ni conciliación de IVA. Los importes negativos están bloqueados en esta versión; las correcciones se hacen anulando/reemplazando una línea.

## Contrato de futuras hojas

Clientes y Proveedores tendrán ID únicos; sus ID opcionales en operaciones se validarán cuando existan esas hojas. Inventario tendrá stock inicial más movimientos de entrada/salida; una venta no generará automáticamente otra salida si también se registra el movimiento. Tesorería referenciará el ID de operación y un ID de movimiento de caja, sin sumarse al resultado. Objetivos e Informes leerán agregados de Ventas y Gastos.

## Capacidad y compatibilidad

Las plantillas preparan 500 filas de ventas, 500 de gastos y 200 productos. Para ampliar, el generador acepta `--ventas`, `--gastos` y `--productos`; se regenera el libro y se trasladan los datos de entrada. Recalcular en Excel antes de usar los indicadores. Google Sheets puede importar el archivo, pero pueden cambiar validaciones, protección, gráficos y formato; no se promete equivalencia.
