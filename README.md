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
