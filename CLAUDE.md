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

Si no tenés el base a mano (sesión nueva), pedile al usuario que adjunte el último
dashboard HTML: los datos vienen embebidos como `const REGISTROS = [...]` y se
reconstruyen desde ahí.

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

## PRIVACIDAD

- Los datos reales **NUNCA** se commitean a git. El dashboard versionado (`demo/`) es
  un demo con datos de ejemplo generados por `demo/generar_demo.py`.
- Al usuario se le entrega el archivo **por chat**, no por el repo.
- **NUNCA** borrar archivos del usuario; si hace falta, pedir el nombre de cada uno.

## Verificar antes de entregar

```bash
node verificar.mjs dashboard.html   # 0 errores + los cuadros cierran exacto
```
