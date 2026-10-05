# Gestión de Inventario CECOSF Martín Henríquez — v1.0

Herramienta web local basada en la estructura de Gestión de Inventario ABG, adaptada para CECOSF Martín Henríquez.

## Datos iniciales
- Maestro generado desde `CONSUMOS MH.xlsx`.
- 204 productos con consumos históricos enero–agosto 2026.
- 135 productos heredaron metadatos mediante equivalencia exacta/segura con el maestro ABG.
- 1 producto(s) identificado(s) como controlado(s).

## Regla especial MH para controlados
Los controlados pueden figurar en el sistema para registro/indicación, pero se marcan como **solo registro / no dispensables en MH**. Por lo tanto no generan alertas de stock crítico, quiebre, sobrestock ni solicitud de reposición. La acción operacional es derivar a ABG según el flujo local.

## Stock
La herramienta funciona sin stock inicial: permite consultar histórico, CPM y curvas de consumo. Las alertas de cobertura/stock se activan cuando se cargue un archivo de Stock Actual Rayen de MH.

## Pedidos
El módulo visible de pedidos queda desactivado hasta contar con una plantilla propia de MH. No se incluye la plantilla ABG para evitar pedidos erróneos.

## Persistencia
Los datos se guardan en localStorage con prefijo exclusivo `mh_v1_`, separado del programa ABG aunque ambos se usen en el mismo navegador.

- v1.1: se elimina únicamente el reporte de trazadores. IAAPS/FOFAR se mantienen para priorizar el inventario rotatorio y facilitar redistribuciones rápidas de trazadores.

- v1.2 filtra el maestro para mostrar únicamente medicamentos. Se retiraron aerocámaras, material dental, jeringas, lancetas, dispositivos, preservativos y otros insumos.
- IAAPS/FOFAR se mantienen para priorización del inventario rotatorio, pero sin reporte de trazadores.

- v1.3 restaura insumos y dispositivos clínicos que deben mantenerse en la herramienta MH, incluyendo aerocámaras, bajadas de suero, cepillos dentales, cintas de glicemia, equipos de monitoreo, jeringas, lancetas, preservativos, seda dental, T de cobre, test de embarazo y similares.
- Se revierte el filtro amplio de la v1.2 para evitar eliminar artículos que sí deben gestionarse en CECOSF MH.

- v1.4: rediseño visual completo con una estética más rosada, suave y amable.
- Se ajustó la interfaz con tonos más cálidos, tarjetas redondeadas, paneles más delicados y una presentación más tierna.
- Se corrigieron textos para identificar el maestro como MH.

- v1.5: refuerzo del diseño rosado y cálido.
- Mientras no exista una carga de stock Rayen, la ausencia de stock no cuenta como excepción para todos los productos.
- Se ajusta el lenguaje a "productos" en inventario rotativo porque MH gestiona medicamentos, insumos y dispositivos.

- v1.6: CSS con nombre nuevo para evitar caché de GitHub Pages y aplicar el rediseño rosado de forma visible.
- Se corrige el encabezado de Maestro ABG a Maestro MH y se usa “productos” en lugar de “fármacos” para el total del maestro.

- v2.0: conexión con Supabase para datos compartidos entre dispositivos.
- Stock, consumos mensuales e inventarios rotatorios/generales se sincronizan en línea.
- Se mantiene almacenamiento local como respaldo operativo si no hay conexión.
- Primer ingreso: usar el botón “Conectar nube” y el código de acceso MH.
- En la primera conexión, si la base está vacía, se migran automáticamente consumos, stock e inventarios que existan en ese navegador.

- v2.1: conexión automática a Supabase.
- Ya no se solicita código de acceso al abrir la herramienta.
- Al abrir MH, la app se conecta y descarga los datos compartidos automáticamente.
- Cada carga de stock, consumo e inventario se sigue sincronizando en línea.
- El indicador superior muestra “Sincronizado” cuando la conexión está activa.
- Al pulsar el indicador de nube se fuerza una resincronización, sin pedir credenciales.

- v2.2: reparación de migración del último stock local.
- Al conectarse, la app compara la fecha del último stock guardado en el navegador con Supabase.
- Si Supabase no tiene stock, o el stock local es más reciente, lo sube automáticamente conservando su fecha.
- Si existe una cabecera de stock en Supabase pero quedó sin detalle, la app la repara automáticamente.

- v2.3: corrige ficha de producto que podía quedar mostrando consumos locales antiguos mientras la nube terminaba de sincronizar.
- Si una ficha está abierta cuando termina la descarga desde Supabase, se refresca automáticamente.
- La ficha muestra explícitamente cuál es el último mes de consumo disponible.
