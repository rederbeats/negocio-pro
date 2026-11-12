# NEGOCIO PRO

Generador reproducible de una plantilla Excel para gestión de pequeños negocios españoles. **Primera versión funcional:** Dashboard, Ventas, Gastos, Productos y Configuración. Las otras siete hojas del producto final están planificadas en [ARCHITECTURE.md](ARCHITECTURE.md), pero aún no se incluyen.

## Generación

Requiere Python 3.11+.

```powershell
python -m pip install -r requirements.txt
python negocio_pro.py
python -m pytest -q
```

El generador crea `dist/negocio_pro_limpio.xlsx` y `dist/negocio_pro_demo.xlsx`. La demo lleva datos **ficticios**. Para aumentar capacidad: `python negocio_pro.py --ventas 1000 --gastos 1000 --productos 400`.

## Uso

1. Abre el libro en Excel y permite el recálculo. Edita empresa y ejercicio en Configuración.
2. Crea productos con ID único. Registra cada línea de venta con un ID de venta único, ID de producto existente, precio y coste unitarios de esa operación, cantidad, descuento y tipo IVA. El precio y el coste se guardan en la venta para mantener su historia.
3. Registra gastos con fecha, ID único, categoría, base neta y tipo IVA.
4. Consulta el Dashboard por ejercicio. Amarillo claro indica entrada; blanco/gris indica cálculo. Las alertas rojas señalan duplicados, referencias desconocidas o filas incompletas. Los filtros y paneles inmovilizados facilitan la navegación.

**Todas las entradas monetarias son sin IVA.** El IVA repercutido y soportado se calcula en columnas separadas. Facturación, gastos y beneficio estimado excluyen IVA. El beneficio resta el coste estimado de los productos vendidos y los gastos netos; no equivale a caja ni a beneficio contable o fiscal. El archivo no es software contable certificado ni una herramienta de cumplimiento fiscal.

## Validación de resultados

`openpyxl` escribe fórmulas y no calcula sus resultados. Las pruebas comprueban estructura, referencias, fórmulas y protecciones; no confirman valores cacheados. Tras generar, abre cada libro en Excel moderno, fuerza recálculo completo (`Ctrl+Alt+F9`), guarda y verifica en la demo: ventas netas **960,00 €**, coste de producto **285,00 €**, gastos netos **155,00 €**, beneficio estimado **520,00 €** y margen **54,17 %**. Comprueba también enero y febrero en la tabla mensual. Un motor compatible como LibreOffice puede servir para automatizar un recálculo, pero hay que contrastar sus resultados y el diseño en Excel antes de distribuir.

Google Sheets puede importar el libro, pero sus validaciones, protección, gráficos, filtros o formato pueden comportarse de otra manera. No se garantiza equivalencia total.

## Alcance pendiente

Clientes, Proveedores, Inventario, Tesorería, Objetivos, Informes e Instrucciones dentro del libro se añadirán en fases posteriores. Por ahora estas instrucciones están en el README. No se incluyen devoluciones, abonos, impagos ni cálculos fiscales.
