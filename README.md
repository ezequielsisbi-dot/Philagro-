# Dashboard de Cobranzas — Philagro S.A.

Convierte el Excel de cobranzas (formato contable) en un **dashboard HTML que se abre
con doble clic**, sin internet y sin instalar nada del lado del que lo mira.

![conceptos](https://img.shields.io/badge/moneda-USD-blue) ![autocontenido](https://img.shields.io/badge/HTML-autocontenido-green)

## Qué muestra

- **KPIs**: total cobrado (en miles), cantidad de cobros, clientes distintos, cliente principal.
- **Cuadro** de cobros por forma de cobro y mes (conceptos → SUB TOTAL → Caja Compensación → TOTAL).
- **Gráficos**: total por mes, ranking de clientes (top 15), evolución por forma de cobro.
- **Contado / Diferido**: qué parte de lo cobrado ya acreditó y qué parte son valores
  que acreditan más adelante, abierto por mes de acreditación.
- **Filtros** combinables de Año, Mes, Desde/Hasta y Cliente (todos de selección múltiple).

## Uso

```bash
pip install openpyxl

# genera el dashboard; si dashboard.html ya existe, lo usa como histórico y le anexa
python3 actualizar.py Cobranzas_2026.xlsx dashboard.html

# verificación antes de entregar (Chromium headless)
node verificar.mjs dashboard.html
```

El script imprime un **control de anexado** (recibos del base, del archivo nuevo,
cuántos ya estaban, cuántos se agregaron, total antes/después y período) para que se
pueda chequear que todo cierra.

### Otras formas de invocarlo

```bash
python3 actualizar.py Cobranzas_2026.xlsx nuevo.html --base anterior.html  # base explícito
python3 actualizar.py --base anterior.html dashboard.html                  # re-generar sin Excel
python3 actualizar.py Cobranzas_2026.xlsx dashboard.html --sin-base        # arrancar de cero
python3 actualizar.py Cobranzas_2026.xlsx --cuentas                        # ver mapeo de cuentas
```

## Si el Excel trae una cuenta contable nueva

El script avisa (`ATENCIÓN: N cuenta(s) sin mapear`) y las manda a "Otros". Se arregla
agregando el patrón en [`assets/cuentas.json`](assets/cuentas.json) — no hace falta
tocar el código.

## Demo

`demo/` tiene un Excel ficticio y su dashboard, para probar el flujo sin datos reales:

```bash
python3 demo/generar_demo.py demo/Cobranzas_demo.xlsx
python3 actualizar.py demo/Cobranzas_demo.xlsx demo/dashboard_demo.html --sin-base
node verificar.mjs demo/dashboard_demo.html
```

## Privacidad

Los datos reales no se versionan: `.gitignore` bloquea los `.xlsx` y los dashboards
generados, salvo los de `demo/`. El dashboard con datos reales se entrega por chat.

Los criterios de cálculo (moneda, signos, conceptos, contado/diferido, redondeo) están
congelados y documentados en [CLAUDE.md](CLAUDE.md).
