# NEGOCIO PRO · VALENCIA LABS

Generador reproducible de una plantilla Excel para gestión de pequeños negocios españoles. Incluye **12 hojas**: Instrucciones, Dashboard, Ventas, Gastos, Productos, Clientes, Proveedores, Inventario, Tesorería, Objetivos, Informes y Configuración. Instrucciones se abre primero y permite empezar sin conocer fórmulas de Excel.

**VALENCIA LABS** es la marca comercial; **NEGOCIO PRO** es el nombre del producto. El cambio de marca no altera el esquema de datos ni los cálculos.

Los archivos para el comprador están en [`release/`](release/). `dist/` contiene la salida directa y reproducible de `openpyxl`; los libros de `release/` se recalcularon con LibreOffice y pasaron la comparación estructural indicada en [`release/VERIFICACIONES.md`](release/VERIFICACIONES.md).

## Generación

Requiere Python 3.11+.

```powershell
python -m pip install -r requirements.txt
python negocio_pro.py
python -m pytest -q
```

El generador crea `dist/negocio_pro_limpio.xlsx` y `dist/negocio_pro_demo.xlsx`. La demo lleva datos **ficticios** con emails `example.invalid`. Para aumentar capacidad: `python negocio_pro.py --ventas 1000 --gastos 1000 --productos 400 --clientes 400 --proveedores 400 --movimientos 1000 --caja 1000 --objetivos 240`.

## Uso

1. Abre el libro en Excel y permite el recálculo. Empieza en la primera hoja, Instrucciones. Sus pasos y enlaces te llevan a las hojas necesarias. En Configuración, edita empresa y ejercicio.
2. Crea productos con ID único y selecciona `Físico`, `Servicio` o `Sin control`. En los físicos indica stock inicial y mínimo. Registra cada línea de venta con ID único, ID de producto existente, precio y coste unitarios de esa operación, cantidad, descuento y tipo IVA. El precio y el coste se guardan en la venta para mantener su historia.
3. Si conoces el cliente, créalo en Clientes y usa su ID en Ventas. Puedes dejar el ID vacío. Haz lo mismo con Proveedores y Gastos. Registra gastos con fecha, ID único, categoría, base neta y tipo IVA.
4. Para una venta física que reduzca existencias, registra **una sola Salida** en Inventario con cantidad y referencia única (por ejemplo, el ID de venta). Las ventas no descuentan stock por sí mismas. Entradas y ajustes también se registran allí. Un movimiento de stock no crea un gasto ni modifica el resultado financiero.
5. Consulta el Dashboard por ejercicio. Verde claro indica entrada; blanco/gris indica cálculo. La columna **Estado** muestra `OK`, `REVISAR` o `FUERA DE EJERCICIO`. Las filas que no estén en `OK` se excluyen de los indicadores del ejercicio; las operaciones válidas de otro año conservan sus importes para conciliar cobros y pagos. Corrige cualquier alerta antes de utilizar el resultado. Los filtros, paneles inmovilizados y enlaces de navegación facilitan el uso. Clientes muestra facturación y número de compras por cliente; Proveedores muestra gasto por proveedor; el Dashboard resume gasto por categoría y alertas de stock.
6. En Configuración, define cuentas como Banco y Efectivo, su **saldo inicial al comenzar el ejercicio** y el **ejercicio al que corresponde ese saldo**. Si cambias el ejercicio, revisa los saldos iniciales y actualiza ese año; mientras no coincidan, el saldo final se oculta y el Dashboard muestra `REVISAR SALDO`. En Tesorería registra cada cobro o pago con su fecha efectiva, cuenta e importe **con IVA**. Si procede de una operación, usa el ID de venta para un Cobro o el ID de gasto para un Pago; puedes registrar varios movimientos parciales con IDs de caja distintos. Ventas y Gastos muestran lo cobrado/pagado y pendiente incluso al cambiar el ejercicio. Un movimiento sin referencia es independiente. No registres en Tesorería una venta o gasto como si fuera una segunda operación de resultado.
7. En Objetivos, introduce ejercicio, mes y las metas que quieras: ventas, beneficio o límite de gastos, **sin IVA**. Cada ejercicio y mes admite una fila. `SIN META`, `SIN ACTIVIDAD` y `n.a.` distinguen metas ausentes, meses sin operaciones y porcentajes con meta cero. Informes muestra los mismos resultados mensuales y anuales del Dashboard, más desgloses por categoría y producto.

Los desplegables usan rangos con nombre. Las fechas deben ser fechas reales de Excel, la cantidad debe ser positiva y el descuento debe estar entre 0 % y 100 %. Se acepta precio cero y descuento del 100 %. Cada ID debe ser único en su hoja. Un producto inexistente o duplicado deja la venta en `REVISAR`; un cliente o proveedor vacío es válido, pero un ID informado debe existir y ser válido. El tipo de producto vacío se interpreta como `Sin control` para conservar datos de versiones anteriores. La protección permite seleccionar celdas, editar entradas y usar filtros; la ordenación de una hoja protegida puede depender de Excel y del rango seleccionado, por lo que se recomienda filtrar. Una operación fuera del ejercicio sigue visible, con estado propio, pero no suma al Dashboard del ejercicio elegido.

**Ventas, Gastos, Productos y Objetivos usan importes sin IVA; Tesorería usa importes efectivamente cobrados o pagados, con IVA.** El IVA repercutido y soportado se calcula en columnas separadas. Facturación, gastos y beneficio estimado excluyen IVA. El beneficio resta el coste estimado de los productos vendidos y los gastos netos; no equivale a caja ni a beneficio contable o fiscal. El saldo inicial de caja es manual y debe revisarse al cambiar de ejercicio. El archivo no es software contable certificado ni una herramienta de cumplimiento fiscal.

## Validación de resultados

Para la demo, la conciliación esperada con aritmética independiente es: venta V-001, total con IVA 871,20 €, cobrado 500 €, pendiente 371,20 €; gasto G-001, total con IVA 145,20 €, pagado 100 €, pendiente 45,20 €. Caja: saldo inicial 1.100 €, cobros 790,40 €, pagos 142,35 € y saldo final 1.748,05 €. La meta anual de ventas introducida es 1.000 €, por lo que el cumplimiento esperado es 96 %. Comprueba estos valores tras recalcular en Excel.

`openpyxl` escribe fórmulas y no calcula sus resultados. Esta versión se recalculó con **LibreOffice** en copias de ambos libros. La demo produjo ventas netas **960,00 €**, coste de producto **285,00 €**, gastos netos **155,00 €**, beneficio estimado **520,00 €**, margen **54,17 %** y caja final **1.748,05 €**. No se hallaron errores de fórmula calculados. La comparación posterior confirmó que se conservaron las validaciones, protecciones, filtros, gráficos y nombres definidos; LibreOffice normalizó algunos códigos de formato numérico. Estas comprobaciones **no demuestran equivalencia visual ni funcionamiento interactivo en Microsoft Excel**. Antes de vender, abre ambos libros en Excel, fuerza recálculo completo (`Ctrl+Alt+F9`), guarda y confirma la demo: enero, ventas 720 €, coste 180 €, gastos 120 €, beneficio 420 €; febrero, ventas 240 €, coste 105 €, gastos 35 €, beneficio 100 €; los otros diez meses, cero. Verifica también **2 clientes con compras**, **2 proveedores válidos**, stock **9** de P-002, desplegables, filtros, alertas, enlaces y protección.

Google Sheets puede importar el libro, pero sus validaciones, protección, gráficos, filtros o formato pueden comportarse de otra manera. No se garantiza equivalencia total.

## Alcance pendiente

La validación final en Excel y los ajustes que surjan de pruebas con usuarios quedan pendientes antes de comercializar. No se incluyen devoluciones, abonos, impagos ni cálculos fiscales.
