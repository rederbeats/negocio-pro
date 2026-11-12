# Verificaciones de la edición VALENCIA LABS · NEGOCIO PRO

Fecha de auditoría: 12/11/2026.

## Realizadas

- Suite de pytest completa: **24 pruebas superadas** (incluye escenarios recalculados con LibreOffice; actualizar el número si se repite la auditoría).
- Revisión de marca: textos, propiedades y XML interno de los cuatro XLSX (`dist` y `release`) usan VALENCIA LABS; no quedan referencias a la marca anterior. NEGOCIO PRO conserva su nombre.
- Las 13.069 fórmulas de cada libro generado son idénticas a las de la versión anterior al cambio de marca; se mantuvieron las 12 hojas y los nombres de archivo.
- Se generaron las versiones limpia y demo desde el mismo código de `main` y se recalcularon copias con LibreOffice **26.2.6.3** instalado localmente. Los XLSX de esta carpeta son copias exactas de esas versiones recalculadas.
- La demo recalculada mostró 960 € de ventas netas, 285 € de coste de producto, 155 € de gastos netos, 520 € de beneficio estimado, margen 54,17 %, caja final 1.748,05 € y cumplimiento de ventas 96 %.
- Una prueba con cobro excesivo en 2027 marcó solo el movimiento nuevo como `REVISAR`; el cobro previo siguió válido y el pendiente permaneció en 371,20 €.
- Una prueba con el cobro válido restante en 2027 dejó pendiente cero sin cambiar el beneficio de gestión de ese año. Un saldo inicial de otro ejercicio provocó `REVISAR SALDO`.
- La versión limpia no contiene operaciones, clientes, proveedores, correos ficticios ni texto de demostración en el libro.
- En ambos archivos completos recalculados: 12 hojas, 13.069 fórmulas por libro, sin celdas de error calculadas ni fórmulas con `#REF!`, `#DIV/0!` o `#VALUE!`. LibreOffice no mostró errores compatibles con referencias circulares en la demo, la plantilla limpia o los escenarios probados; esto no demuestra que todos los datos futuros estén libres de errores.
- La comparación entre salida original y copia recalculada conservó número de validaciones, formatos condicionales, gráficos, nombres definidos, protección, filtros, paneles inmovilizados y hoja de apertura. Se comprobaron celdas de entrada desbloqueadas y fórmulas bloqueadas. LibreOffice normalizó la sintaxis de algunos formatos numéricos sin cambiar su significado aparente.
- Se inspeccionaron anclas de gráficos y posiciones de tablas: empiezan debajo de los indicadores y no ocupan las tablas laterales. Esta es una comprobación estructural, no una revisión visual en Excel.

## Pendientes antes de vender en Shopify

- Abrir **ambos archivos de esta carpeta** en Microsoft Excel, forzar el recálculo completo y confirmar los valores de la demo y la ausencia de errores.
- Probar en la interfaz de Excel la edición de celdas verdes, desplegables, filtros, enlaces internos y bloqueo de fórmulas.
- Revisar visualmente las 12 hojas en Excel, especialmente anchos, texto, gráficos, colores, zoom e impresión en pantallas habituales.
- Hacer una prueba de uso sin ayuda con una persona del público objetivo y corregir los puntos de confusión.
- Confirmar la política de soporte y entrega de archivos en la ficha de Shopify. No anunciar equivalencia total con Google Sheets ni cumplimiento contable o fiscal.

**Estado comercial:** pendiente de estas comprobaciones en Microsoft Excel. El recálculo con LibreOffice y las pruebas automatizadas reducen el riesgo, pero no sustituyen la validación del producto en la aplicación de destino.
