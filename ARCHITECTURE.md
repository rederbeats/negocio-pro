# NEGOCIO PRO — arquitectura de datos (v0.2)

## Alcance y grano

El libro contiene Dashboard, Ventas, Gastos, Productos, Clientes, Proveedores, Inventario y Configuración. Cada fila de Ventas representa **una línea de venta**; varias líneas de una factura pueden compartir referencia documental, pero cada línea tiene un ID único. Cada fila de Gastos representa un gasto. Los ID de producto, cliente, proveedor y movimiento son únicos en sus hojas.

## Relaciones y dependencias

```text
Configuración ──> listas de categorías y tipos IVA (validación)
Productos ──────> Ventas (ID de producto; validación de existencia)
Clientes ───────> Ventas (ID opcional; si existe, debe ser válido)
Proveedores ────> Gastos (ID opcional; si existe, debe ser válido)
Productos ──────> Inventario (solo productos físicos)
Inventario ─────> Productos (stock disponible y alerta)
Configuración ──> Ventas y Gastos (ejercicio y rangos con nombre)
Ventas ─────────> Dashboard (ventas netas, coste estimado y evolución)
Gastos ─────────> Dashboard (gastos netos y evolución)
Clientes, Proveedores y Productos ──> Dashboard (indicadores)
```

Ventas conserva precio unitario, coste unitario y tipo de IVA **como datos de la operación**. El catálogo no altera ventas históricas al cambiar un precio o coste. La validación avisa de un ID de producto inexistente; no se calcula automáticamente un precio de catálogo para evitar cambios retroactivos. Un ID de cliente/proveedor en blanco es válido; uno informado debe apuntar a una fila válida y única. Los campos de contacto son opcionales; los emails informados reciben una comprobación básica de forma, no de entregabilidad. Las fichas con situación Inactivo siguen siendo referencias válidas para conservar operaciones históricas.

## Inventario

El tipo de producto vacío se interpreta como `Sin control` para conservar productos creados con la versión anterior. Los productos físicos requieren stock inicial y mínimo; los demás no admiten movimientos.

Productos distingue `Físico`, `Servicio` y `Sin control`. Solo `Físico` admite stock inicial, mínimo y movimientos. Inventario registra movimientos con ID único, fecha, ID de producto, tipo, cantidad positiva y referencia. Los tipos `Entrada` y `Ajuste +` suman; `Salida` y `Ajuste -` restan. El stock disponible es stock inicial más impactos de movimientos `OK`. Un movimiento inválido queda en `REVISAR` y no altera stock. Una `Salida` exige referencia única entre las salidas para evitar duplicados aunque cambie el ID de movimiento. Una referencia puede ser el ID de línea de venta; no se impone que toda salida corresponda a una venta, porque también hay mermas y usos internos. Para salidas parciales se usan referencias únicas distintas. El stock negativo se muestra como alerta, no se oculta ni se bloquea.

**Las ventas no descuentan stock automáticamente.** Por cada venta de producto físico que deba reducir existencias, el usuario registra una sola `Salida` con referencia única. No debe registrar una segunda salida por esa misma operación. Los movimientos no alimentan Ventas, Gastos ni el beneficio; una compra de existencias puede estar en Gastos según el criterio de gestión del usuario, pero la entrada de stock no vuelve a sumar el gasto.

Cada operación tiene una fórmula de Estado. `OK` requiere fecha válida del ejercicio, ID único, referencias válidas y entradas numéricas dentro de los límites. `REVISAR` identifica filas incompletas o inconsistentes; `FUERA DE EJERCICIO` identifica operaciones válidas de otro año. Solo `OK` alimenta importes calculados y agregados del Dashboard. Los contadores del Dashboard muestran las filas excluidas, evitando indicadores aparentemente completos cuando hay errores. Las validaciones de Excel ayudan a prevenir entradas, mientras que Estado también detecta datos pegados, cambios posteriores en catálogos y duplicaciones.

## Definiciones financieras

- Precio unitario, coste unitario e importe de gasto se introducen **sin IVA**. Descuento es un porcentaje de 0 a 100 %.
- Base venta = cantidad × precio × (1 − descuento). IVA repercutido = base × tipo. Total factura = base + IVA.
- Coste de producto estimado = cantidad × coste unitario de la línea; beneficio operativo estimado = ventas netas − coste de producto estimado − gastos netos. No es beneficio contable ni fiscal.
- IVA soportado de gastos se muestra aparte. Ningún IVA se incluye en ingresos, costes o beneficio.
- Dashboard y resumen mensual usan fecha de operación, no fecha de cobro o pago. La futura Tesorería mostrará flujos de caja sin alimentar el resultado operativo.
- No se modelan todavía devoluciones, abonos, impagos, existencias, amortizaciones ni conciliación de IVA. Los importes negativos están bloqueados en esta versión; las correcciones se hacen anulando/reemplazando una línea.

## Contrato de futuras hojas

Tesorería referenciará el ID de operación y un ID de movimiento de caja, sin sumarse al resultado. Objetivos e Informes leerán agregados de Ventas y Gastos.

## Capacidad y compatibilidad

Las plantillas preparan 500 filas de ventas, 500 de gastos, 200 productos, 200 clientes, 200 proveedores y 500 movimientos. Para ampliar, el generador acepta `--ventas`, `--gastos`, `--productos`, `--clientes`, `--proveedores` y `--movimientos`; se regenera el libro y se trasladan los datos de entrada. Recalcular en Excel antes de usar los indicadores. Google Sheets puede importar el archivo, pero pueden cambiar validaciones, protección, gráficos y formato; no se promete equivalencia.

La prueba automatizada no sustituye el recálculo ni las pruebas interactivas en Excel. La protección de hoja permite selección y filtros; la ordenación de rangos protegidos puede no funcionar en todas las versiones. Las listas usan rangos con nombre para evitar referencias directas a otra hoja en validación de datos.
