# Dashboard de Ventas — Philagro (agroinsumos)

Genera un dashboard HTML **autocontenido** (se abre con doble clic, sin
internet, sin CDN) a partir del Excel de facturación. Proyecto
**independiente** del dashboard de cobranzas: no comparte ni pisa sus
archivos — todo vive bajo esta carpeta (`ventas/`).

## Instalación (una sola vez)

```bash
pip install pandas openpyxl
cd ventas && npm install        # trae Chart.js, que se inyecta inline en el HTML
```

Chart.js **no** se commitea al repo: se instala con `npm install` y el
generador lo **inyecta inline** en cada dashboard. El HTML resultante queda
100 % autocontenido (sin CDN, sin fetch, se abre con doble clic offline).

## Uso

```bash
python3 ventas_dashboard.py <FACTURACION.xlsx>
```

Genera siempre un archivo **nuevo**: `dashboard_ventas_<nombre-del-excel>.html`.
Si ya existe, agrega `_v2`, `_v3`, … Nunca sobrescribe.

Después, verificar que cierre sin errores:

```bash
node verificar_ventas.mjs dashboard_ventas_<archivo>.html
```
(Si `node` no encuentra el módulo `playwright`, correr con
`NODE_PATH=/opt/node22/lib/node_modules node verificar_ventas.mjs ...`)

## Actualización incremental (histórico)

`ventas_dashboard.py` mantiene un histórico acumulado en
`data/historico_ventas.json` (se crea solo, no tocar a mano). Cada corrida:

1. Detecta qué meses trae el Excel nuevo.
2. Para esos meses, **reemplaza** la versión del histórico por la del Excel
   nuevo (se asume más completa: p. ej. "julio a la fecha").
3. Los meses que el Excel nuevo no menciona **no se tocan**.
4. Los meses que no existían todavía se agregan.
5. Avisa por consola si algún mes reemplazado trajo **menos** comprobantes
   que el histórico anterior (posible dato incompleto).

El dashboard resultante siempre se arma sobre el histórico completo
acumulado, no solo sobre el Excel de la corrida — es un dashboard
**histórico** multi-año.

## Columnas esperadas del Excel

`Fechacomprobante`, `Transacconsubtiponombre`, `Comprobante`, `Cliente`,
`Costo`, `Condicionpago`, `Moneda`, `Vendedor`, `Producto`, `Cantidad`,
`Precio`, `Importemonsecundaria`, `Principioactivo` (pueden venir otras
columnas de más, se ignoran).

## Criterios congelados

No cambiar sin pedido explícito — están implementados en
`ventas_dashboard.py`:

1. **Moneda**: siempre USD, tomada de `Importemonsecundaria` tal cual viene
   (firmada; las notas de crédito son negativas y se netean solas al sumar).
   No se convierte nada, ni aunque haya filas marcadas en pesos.
2. **Precio unitario por fila**: si `Transacconsubtiponombre == "Nota Liquido
   producto"` → se usa `Costo`; en cualquier otro caso (incluye
   `Asesoramiento-Comisiones` con `Costo == 0`, que es intencional) → se usa
   `Precio`.
3. **Plazo (días)** desde `Condicionpago`: `CONTADO` = 0; rangos
   ("30-60 DIAS", "45-75-105") toman el **máximo**; condiciones no numéricas
   (CANJE, PLATAFORMA, TARJETA, COMPENSACION, GRANOS, …) quedan sin plazo —
   se excluyen del cuadro de "Plazos por producto" y quedan al final del
   ordenamiento por condición de pago.
4. **Cantidad de operaciones** = comprobantes distintos (columna
   `Comprobante`), no cantidad de líneas.

## Vistas del dashboard

- KPIs: ventas totales (en miles de USD), comprobantes, clientes distintos,
  vendedor principal.
- Gráfico de ventas por mes + Top 15 de la dimensión activa.
- Cuadro dinámico "Ventas por [dimensión] y mes" con selector
  (Vendedor / Producto / Cliente / Principio activo / Condición de pago —
  esta última ordenada por días de plazo).
- Ventas por mes y año (filas = meses, columnas = años).
- Cantidades por producto (campo `Cantidad`) por mes.
- Precios por producto: precio unitario promedio **ponderado por USD** —
  Σ(precio × importe) / Σ(importe) — por mes.
- Plazos por producto: plazo máximo (días), clientes a ese plazo, cantidad
  y precio unitario en esa venta.
- Comparativo por año: barras agrupadas, eje X = mes, una serie por año.

Filtros: Año, Mes, Vendedor, Cliente, Producto (todos multi-selección con
chips + buscador + "marcar visibles"), más Desde/Hasta. Todos combinan con
AND; vacío = todos. Es histórico: soporta múltiples años (se arma solo
según las fechas de los datos).

## Estructura

```
ventas/
  ventas_dashboard.py           generador (Python + pandas)
  verificar_ventas.mjs          verificador (Node + Playwright headless)
  package.json                  dependencia Chart.js (se inyecta inline)
  assets/plantilla_ventas.html  plantilla con el motor de filtros/tablas/
                                gráficos y dos placeholders:
                                  /*__DATOS_JS__*/  → datos JSON embebidos
                                  /*__CHARTJS__*/   → Chart.js inline
  data/historico_ventas.json    histórico acumulado (se genera solo)
  dashboard_ventas_*.html       salidas generadas (se generan solas)
```
