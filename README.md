# NEGOCIO PRO

Generador reproducible de una plantilla Excel para gestión de pequeños negocios españoles. **Versión actual:** Dashboard, Ventas, Gastos, Productos, Clientes, Proveedores, Inventario y Configuración. Tesorería, Objetivos, Informes e Instrucciones dentro del libro quedan para fases posteriores.

## Generación

Requiere Python 3.11+.

```powershell
python -m pip install -r requirements.txt
python negocio_pro.py
python -m pytest -q
```

El generador crea `dist/negocio_pro_limpio.xlsx` y `dist/negocio_pro_demo.xlsx`. La demo lleva datos **ficticios** con emails `example.invalid`. Para aumentar capacidad: `python negocio_pro.py --ventas 1000 --gastos 1000 --productos 400 --clientes 400 --proveedores 400 --movimientos 1000`.

## Uso

1. Abre el libro en Excel y permite el recálculo. Edita empresa y ejercicio en Configuración.
2. Crea productos con ID único y selecciona `Físico`, `Servicio` o `Sin control`. En los físicos indica stock inicial y mínimo. Registra cada línea de venta con ID único, ID de producto existente, precio y coste unitarios de esa operación, cantidad, descuento y tipo IVA. El precio y el coste se guardan en la venta para mantener su historia.
3. Si conoces el cliente, créalo en Clientes y usa su ID en Ventas. Puedes dejar el ID vacío. Haz lo mismo con Proveedores y Gastos. Registra gastos con fecha, ID único, categoría, base neta y tipo IVA.
4. Para una venta física que reduzca existencias, registra **una sola Salida** en Inventario con cantidad y referencia única (por ejemplo, el ID de venta). Las ventas no descuentan stock por sí mismas. Entradas y ajustes también se registran allí. Un movimiento de stock no crea un gasto ni modifica el resultado financiero.
5. Consulta el Dashboard por ejercicio. Amarillo claro indica entrada; blanco/gris indica cálculo. La columna **Estado** muestra `OK`, `REVISAR` o `FUERA DE EJERCICIO`. Las filas que no estén en `OK` se excluyen de los indicadores; el Dashboard cuenta las filas excluidas. Corrige cualquier alerta antes de utilizar el resultado. Los filtros y paneles inmovilizados facilitan la navegación. Clientes muestra facturación y número de compras por cliente; Proveedores muestra gasto por proveedor; el Dashboard resume gasto por categoría y alertas de stock.

Los desplegables usan rangos con nombre. Las fechas deben ser fechas reales de Excel, la cantidad debe ser positiva y el descuento debe estar entre 0 % y 100 %. Se acepta precio cero y descuento del 100 %. Cada ID debe ser único en su hoja. Un producto inexistente o duplicado deja la venta en `REVISAR`; un cliente o proveedor vacío es válido, pero un ID informado debe existir y ser válido. El tipo de producto vacío se interpreta como `Sin control` para conservar datos de versiones anteriores. La protección permite seleccionar celdas, editar entradas y usar filtros; la ordenación de una hoja protegida puede depender de Excel y del rango seleccionado, por lo que se recomienda filtrar. Una operación fuera del ejercicio sigue visible, con estado propio, pero no suma al Dashboard del ejercicio elegido.

**Todas las entradas monetarias son sin IVA.** El IVA repercutido y soportado se calcula en columnas separadas. Facturación, gastos y beneficio estimado excluyen IVA. El beneficio resta el coste estimado de los productos vendidos y los gastos netos; no equivale a caja ni a beneficio contable o fiscal. El archivo no es software contable certificado ni una herramienta de cumplimiento fiscal.

## Validación de resultados

`openpyxl` escribe fórmulas y no calcula sus resultados. Las pruebas comprueban estructura, referencias, fórmulas, protecciones y, con aritmética Python independiente, los importes esperados de los datos demo. **No confirman que Excel haya evaluado las fórmulas ni el funcionamiento interactivo de desplegables y protección.** No había Excel ni LibreOffice disponibles en el entorno de generación. Antes de distribuir, abre cada libro en Excel moderno, fuerza recálculo completo (`Ctrl+Alt+F9`), guarda y verifica en la demo: ventas netas **960,00 €**, coste de producto **285,00 €**, gastos netos **155,00 €**, beneficio estimado **520,00 €** y margen **54,17 %**. Enero: ventas 720 €, coste 180 €, gastos 120 €, beneficio 420 €. Febrero: ventas 240 €, coste 105 €, gastos 35 €, beneficio 100 €. Los otros diez meses deben ser cero. Verifica además **2 clientes con compras**, **2 proveedores válidos** y stock disponible **9** del producto P-002 (10 inicial − 3 salida + 2 entrada). Comprueba entradas, desplegables, filtros, alertas y protección en la interfaz. Un motor compatible como LibreOffice puede automatizar un recálculo, pero hay que contrastar los resultados y el diseño en Excel antes de distribuir.

Google Sheets puede importar el libro, pero sus validaciones, protección, gráficos, filtros o formato pueden comportarse de otra manera. No se garantiza equivalencia total.

## Alcance pendiente

Tesorería, Objetivos, Informes e Instrucciones dentro del libro se añadirán en fases posteriores. Por ahora estas instrucciones están en el README. No se incluyen devoluciones, abonos, impagos ni cálculos fiscales.
