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
