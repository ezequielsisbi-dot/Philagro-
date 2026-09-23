# Dashboard de Cobranzas — Philagro S.A.

Proyecto que genera un **dashboard HTML autocontenido** (se abre con doble clic, sin
internet) a partir de un Excel de cobranzas en formato contable (`Cobranzas_XXXX.xlsx`).

---

## Flujo fijo (siempre lo mismo con cada archivo que llega)

```bash
python3 actualizar.py Cobranzas_2026.xlsx dashboard.html   # genera / anexa
node verificar.mjs dashboard.html                          # verifica antes de entregar
```

Cada archivo nuevo se **anexa** sobre el histórico existente (el "base"), sin
re-decidir criterios.

### Formato congelado (confirmado por el usuario, 23/09/2026)

Este es EL dashboard. Cada consulta se actualiza sobre `historico/dashboard.html`,
sin rehacerlo ni cambiarle la forma. Los dos modos del cuadro Contado/Diferido
—**Diferido a hoy** y **Diferido por plazo (vto. vs cobro)**, criterio 10— son parte
del formato y no se tocan.

### El histórico está en el repo — NO pedírselo al usuario

`historico/dashboard.html` es siempre el último dashboard generado, versionado acá.
En una sesión nueva (el contenedor es efímero y se recicla a los pocos días) se usa
ESE archivo como base, sin pedirle nada al usuario:

```bash
python3 actualizar.py Cobranzas_nuevo.xlsx historico/dashboard.html
node verificar.mjs historico/dashboard.html
```

Después se commitea `historico/dashboard.html` actualizado y el Excel en `datos/`,
y se le entrega al usuario una copia por chat. Pedirle el HTML al usuario es el
último recurso, sólo si `historico/dashboard.html` no existe o está corrupto: los
datos viajan embebidos como `const REGISTROS = [...]` y se reconstruyen desde ahí.

```bash
# base explícito (o cuando el HTML de salida es otro archivo)
python3 actualizar.py Cobranzas_2026.xlsx nuevo.html --base dashboard_anterior.html

# re-generar el HTML desde el histórico, sin Excel nuevo (p. ej. tras tocar la plantilla)
python3 actualizar.py --base dashboard_anterior.html dashboard.html

# arrancar de cero, ignorando el histórico
python3 actualizar.py Cobranzas_2026.xlsx dashboard.html --sin-base

# ver cómo quedó mapeada cada cuenta contable (sin generar nada)
python3 actualizar.py Cobranzas_2026.xlsx --cuentas
```

## Cómo anexar sin perder ni duplicar

- **Recibos CON número**: dedup por columna `Documento`. Un recibo que ya está no se
  re-agrega; si aparece en el archivo nuevo y en el base, se toma del **NUEVO** (trae
  `Fechavto` = acreditación exacta). Los nuevos se agregan; los viejos se conservan.
- **Movimientos SIN número** (ajustes, p. ej. Caja Compensación): dedup por
  **CONTENIDO** (fecha + cliente + forma + importe), no por `Documento`. Se respetan
  las repeticiones legítimas: si el base ya tiene dos ajustes idénticos, se toleran dos.
- Siempre se muestra un **control de anexado**: recibos del base, del archivo nuevo,
  cuántos ya estaban, cuántos se agregaron, total USD antes/después, período.
  Todo tiene que cerrar.

## Criterios congelados (no cambiar sin pedido explícito del usuario)

1. **Moneda = USD**, tomada de la columna real `Importe moneda secundaria`. Nunca
   inventar cotización. Si el archivo no la trae, se usa el importe principal en ARS.
2. Los importes **ya vienen firmados** (los Haber son negativos = anulaciones). Se usan
   tal cual (se netean solos al sumar). **NO** multiplicar Haber por −1.
3. **Forma de cobro** = columna `Cuenta` agrupada en conceptos: Valores
   (echeqs/cheques), Valores en recaudadoras, Bancos / transferencias, Tarjetas /
   plataformas, Impuestos / retenciones, Caja Chica (incluye la cuenta "Redondeo"),
   Caja Compensación. El mapa cuenta → concepto vive en `assets/cuentas.json`; una
   cuenta que no matchea ninguna regla cae en "Otros" y el script lo avisa.
4. **Cantidad de cobros** = recibos distintos (columna `Documento`), no movimientos
   contables. Los movimientos sin número no son recibos y no se cuentan.
5. Cuadro principal **"Detalle de cobros por forma de cobro y mes"**: filas =
   conceptos; columnas = meses agrupados por año, orden descendente (mes más reciente
   primero). Estructura: conceptos → SUB TOTAL → Caja Compensación → TOTAL. Importes
   **enteros** (sin centavos). Negativos entre paréntesis (contable). Primera columna
   fija (sticky).
6. **Filtros**: Año, Mes, Desde/Hasta (calendario) y Cliente (con buscador), todos de
   selección múltiple y combinables (AND). Los paneles se contraen (clic afuera /
   re-clic cierra; abrir uno cierra los demás). Aplican a KPIs, cuadro, gráficos y
   Contado/Diferido.
7. **KPIs**: "Total cobrado" en MILES (ej. `US$ 66.853 mil`); cards de Contado /
   Diferido / Caja Comp. / Retenciones. El resto en enteros.
8. **Autocontenido**: Chart.js embebido inline; datos como JSON embebido. Nada de
   fetch/CDN/internet.
9. **Sección "Contado / Diferido"** (debajo del dashboard, mismos filtros). Por fecha
   de cobranza, clasifica cada movimiento en 4 categorías:
   - **DIFERIDO** = "cobrado pero todavía no acreditado a la fecha de corte". Sólo
     "Valores (echeqs/cheques)" y "Valores en recaudadoras" cuya acreditación
     (`Fechavto` + 2 días) es POSTERIOR a la fecha de corte (= última cobranza del
     conjunto = "hoy"). Se decide por fecha exacta, no por mes: un echeq que acredita
     más adelante dentro del mismo mes de cobro también es diferido. Se abre por el mes
     en que acredita (columna expandible, puede incluir el mes en curso).
   - **Caja Compensación** y **Retenciones/Impuestos**: columnas propias.
   - **CONTADO**: todo el resto (bancos, tarjetas, caja chica y los valores que ya
     acreditaron al corte).
   - Cuadro: filas = mes de cobrado; columnas = Contado · Diferido ▸ · Caja Comp. ·
     Retenciones · Total. Enteros. El total coincide con el del dashboard.
   - Datos históricos sin `Fechavto` se aproximan por mes (diferido si acreditan en el
     mes de corte o después).

10. **Selector de cómo se mide el diferido** (agregado 23/09/2026, a pedido del
    usuario). El criterio 9 mide la exposición de HOY, así que los meses viejos dan
    cero: a un septiembre del año pasado ya le acreditó todo, y eso impide comparar
    un mes contra el mismo mes de otro año. El cuadro Contado/Diferido tiene dos
    modos, y el default sigue siendo el criterio 9 sin cambios:
    - **Diferido a hoy** (default): acreditación posterior a la fecha de corte.
    - **Diferido por plazo (vto. vs cobro)**: el valor vencía DESPUÉS del día en que
      se cobró (`Fechavto > fecha de cobro`, campo `p` del registro). Un cheque al día
      es contado. No depende de cuánto tiempo pasó, así que compara meses entre sí.
      Requiere `Fechavto`: los movimientos que no la traen (todo 2024 y 2025, que
      vinieron del dashboard original) se aproximan por mes de acreditación y el
      dashboard lo avisa arriba del cuadro (`META.plazo_desde`). Para que sea exacto
      hay que re-exportar esos años con la columna Fechavto.

    Los dos modos salen de las mismas celdas enteras, así que el total general no
    cambia: sólo se mueve el reparto entre Contado y Diferido.

### Redondeo (cómo cierran los cuadros)

Los dos cuadros se calculan sobre la **misma base entera**: se agrupa por la partición
más fina que alguno necesita —mes de cobro × forma × diferido × mes de acreditación— y
se redondea **una sola vez**, ahí. Todo lo demás (filas, columnas, SUB TOTAL, TOTAL,
Contado/Diferido) se suma en enteros, así que cierra exacto y los dos totales coinciden.
Redondear cada celda por separado hace que las filas arrastren ±1 y los totales no den.

## Estructura del repo

| Archivo | Qué hace |
| --- | --- |
| `actualizar.py` | Generador: carga + limpieza del Excel → anexado → agregados → inyección en la plantilla. |
| `assets/plantilla.html` | Plantilla con placeholders (`__REGISTROS_JSON__`, `__META_JSON__`, `__CHARTJS__`, …). |
| `assets/chart.umd.min.js` | Chart.js v4.5.1 vendorizado (se embebe inline en cada salida). |
| `assets/cuentas.json` | Mapa cuenta contable → concepto de forma de cobro. |
| `verificar.mjs` | Verificación headless (Chromium): 0 errores de consola + los cuadros cierran. |
| `demo/` | Excel ficticio y dashboard de ejemplo (lo único con "datos" que se versiona). |

### Estructura de un registro embebido

```js
{ fecha:"2026-08-14", mes:"2026-08", cliente:"…", forma:"Valores (echeqs/cheques)",
  importe: 12345.67, doc:"R-00001-00006526",
  a:"2026-10",   // mes en que acredita
  d: 1,          // 1 = diferido a la fecha de corte
  fa:"2026-10-13" // fecha exacta de acreditación (Fechavto + 2), si se conoce
}
```

`fa` permite **recalcular** contado/diferido cuando la fecha de corte avanza con un
archivo nuevo. Los registros históricos que no la tienen se resuelven por mes.

## Serie R-00005 — bajo observación

8 recibos de julio 2026 (26 movimientos, 69.352,52 USD) que llegaron en el export de
enero–septiembre del 14/09 y **no** aparecen en el histórico completo del 23/09. Es la
única serie R-00005 de todo el período, la única que nunca trae `Fechavto`, y tres de
sus recibos son íntegramente Caja Compensación —dos a nombre de PHILAGRO S.A, la propia
empresa—. El usuario cree que fue un error de carga, pero pidió **conservarlos** y
tenerlos presentes (23/09/2026).

Qué hacer: no borrarlos. Si en un export futuro la serie reaparece, se anexa normal; si
el usuario confirma la baja, se sacan. Para listarlos:

```bash
python3 -c "import re,json;h=open('historico/dashboard.html',encoding='utf-8').read();\
print([r for r in json.loads(re.search(r'^const REGISTROS = (.+?);\s*$',h,re.M).group(1)) \
if r['doc'].startswith('R-00005')])"
```

## Datos reales en el repo

El usuario pidió explícitamente (23/09/2026) que el histórico y los Excel de origen
queden guardados en el repo, para no tener que re-adjuntar el dashboard cada vez que
el contenedor se recicla. Por eso se versionan:

- `historico/dashboard.html` — el último dashboard generado (la base).
- `datos/*.xlsx` — los Excel de origen que va mandando.

Quedó advertido dos veces que el repositorio es **público** y que esto expone razones
sociales e importes de clientes, y que en git permanecen en el historial aunque se
borren. El usuario decidió seguir así. No volver a plantearlo salvo que él lo traiga.

- Al usuario se le entrega el archivo **por chat** además de dejarlo en el repo.
- **NUNCA** borrar archivos del usuario; si hace falta, pedir el nombre de cada uno.
- `demo/` sigue siendo datos ficticios (`demo/generar_demo.py`), para probar el flujo.

## Verificar antes de entregar

```bash
node verificar.mjs dashboard.html   # 0 errores + los cuadros cierran exacto
```

---

# Dashboard de Ventas — Philagro S.A. (`ventas/`)

Proyecto **independiente** del de cobranzas: no comparte ni pisa sus archivos.
Genera un dashboard HTML autocontenido (doble clic, sin internet) a partir de un
Excel de **facturación**.

## Flujo fijo

```bash
cd ventas
python3 ventas_dashboard.py <FACTURACION.xlsx>     # fusiona y genera
node verificar_ventas.mjs dashboard_ventas_*.html  # 0 errores + los cuadros cierran
```

Genera SIEMPRE un archivo nuevo `dashboard_ventas_<nombre-del-excel>.html`; si ya
existe agrega `_v2`, `_v3`… **Nunca sobrescribe.** Se entrega por chat.

## Dónde está la base — PENDIENTE

El histórico de ventas **todavía no se versiona**: el repo es público y son 357
clientes con sus importes. Hasta que el usuario lo pase a privado, en cada sesión
nueva hay que **pedírselo**: el `historico_ventas.json.gz` o cualquier
`dashboard_ventas_*.html` (los datos viajan embebidos y `sembrar_historico.py` los
reconstruye).

```bash
python3 sembrar_historico.py <dashboard_ventas_VIEJO.html> --forzar
```

Cuando el repo pase a privado: borrar `ventas/historico/` y `ventas/datos/` del
`.gitignore` y versionar el histórico y los Excel, como ya se hace en cobranzas.

## Cómo fusionar sin perder ni duplicar

- Se **reemplaza el mes completo** que trae el Excel nuevo por su versión (se asume
  más completa) y se suman los meses que no existían. Los meses que el Excel no
  menciona no se tocan.
- La verificación de pérdida es **por CONTENIDO** (fecha + cliente + producto +
  importe), no por cantidad de comprobantes: que un mes suba de 50 a 130 no
  garantiza que los 50 viejos estén entre los 130 (pasó el 23/09/2026). Tampoco se
  puede comparar por número de comprobante: el histórico sembrado desde un dashboard
  viejo guarda un índice, no el número real (`A-00010-00018094`).
- Las filas del histórico que el Excel nuevo no trae **se conservan y se listan**.
  Si el usuario confirma que son basura, se sacan a mano (ver abajo).

## Criterios congelados (no cambiar sin pedido explícito)

1. **Moneda = USD** de `Importemonsecundaria`. Los importes ya vienen firmados: las
   notas de crédito son negativas y se netean solas al sumar. No convertir nada.
2. **Precio unitario por fila**: si `Transacconsubtiponombre == "Nota Liquido
   producto"` → `Costo`; si `Costo == 0` (producto real) → `Precio`; resto →
   `Precio`. Las filas "Asesoramiento-Comisiones" tienen `Costo` 0 intencional.
3. **Plazo (días)** desde `Condicionpago`: CONTADO = 0; los rangos ("30-60 DIAS",
   "45-75-105") toman el MÁXIMO; las no numéricas (CANJE, PLATAFORMA, TARJETA,
   COMPENSACION, GRANOS) quedan al final del cuadro de plazos.
4. **Cantidad de operaciones** = comprobantes distintos (columna `Comprobante`).
5. **Precios por producto**: promedio PONDERADO por USD, Σ(precio × importe) /
   Σ(importe), por mes.
6. Importes **enteros**, negativos entre paréntesis. Cada tabla cierra sola.
7. **Autocontenido**: Chart.js inline (de `node_modules`, `npm install`), datos como
   JSON embebido. Nada de fetch/CDN.

## Vistas

KPIs (ventas en miles, comprobantes, clientes, vendedor principal) · cuadro dinámico
"Ventas por [Vendedor/Producto/Cliente/Principio activo/Condición de pago] y mes"
(condición de pago ordenada por días) · ventas por mes y año · cantidades por
producto · precios ponderados · plazos por producto · y al final el gráfico
"Comparativo por año" (barras agrupadas, eje X = mes, una serie por año).

Filtros multi-selección con chips, buscador y "Marcar visibles": Año, Mes, Vendedor,
Cliente, Producto, más Desde/Hasta. Combinan con AND; vacío = todos. Es histórico:
los años se arman solos según las fechas de los datos.

## Decisiones del usuario

- **23/09/2026** — La fila `A-00010-00021252` (AGRONOMIA ALVAREZ SRL, "Gastos
  Varios-IVA 0%", 805,70 USD, 2026-09-04) quedó **excluida** a pedido del usuario:
  estaba en el histórico pero no en el export nuevo, y su numeración (21252) está
  fuera del rango del resto (18xxx). Si reaparece en un export futuro, preguntar
  antes de incorporarla.
