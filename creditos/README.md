# Asignación de crédito por cliente — Philagro S.A.

Propone un **límite de crédito por cliente** a partir de las ventas y la condición
de pago. Proyecto **independiente** de los dashboards de cobranzas y ventas: todo
vive bajo `creditos/`, no pisa ni comparte archivos con ellos.

> Alcance: esta propuesta mide **sólo comportamiento comercial** (qué compró el
> cliente, a qué plazo, con qué continuidad). Es el piso del análisis, no la
> decisión. Antes de aprobar hay que contrastarla con balances, índices, antigüedad
> en el rubro, deuda tomada con terceros y garantías.

## Instalación

```bash
pip install pandas openpyxl numpy
```

## Uso

```bash
python3 asignar_creditos.py <FACTURACION.xlsx>          # Excel de facturación
python3 asignar_creditos.py <dashboard_ventas_XX.html>  # dashboard de ventas ya generado
python3 asignar_creditos.py <historico_ventas.json>     # histórico del proyecto ventas
```

Genera `creditos_sugeridos_<fuente>.xlsx` (hoja **Límites sugeridos** + hoja
**Criterios** con los parámetros de la corrida) y un resumen por consola.

Verificar el motor de cálculo:

```bash
python3 verificar_creditos.py     # 20 chequeos contra cálculo manual
```

## La idea central: el crédito se ocupa y se libera

Una factura **ocupa** cupo desde que se emite hasta que vence (fecha + plazo de su
condición de pago). Ahí se **libera** y el cupo vuelve a estar disponible.

De eso se desprenden las dos cosas que pediste:

- **Cuándo se renueva**: el ciclo es el plazo de la condición de venta. A 90 días el
  cupo rota 4 veces al año; a 180 días, 2 veces.
- **Las facturas dentro del plazo acumulan**: si a un cliente a 60 días le facturás
  el 10/01 y otra vez el 10/02, el 10/02 tiene las dos abiertas. El límite tiene que
  cubrir esa **suma**, no la factura más grande.

El script calcula la **exposición diaria real**: para cada día del historial, cuánto
tenía abierto ese cliente. Cada factura vence con **su propia** condición, así que un
cliente con mezcla de 30 y 180 días se modela exacto. Las notas de crédito entran
negativas y liberan cupo solas.

## Cómo se llega al límite

**Paso 1 — Plazo ponderado.** Promedio de días ponderado por USD facturado (sólo
ventas a crédito con importe positivo; las NC se netean en los volúmenes pero no
distorsionan el promedio). De ahí salen los **ciclos por año** = 365 / plazo.

**Paso 2 — Dos medidas de necesidad**, porque el agro es estacional:

| Medida | Fórmula | Para qué sirve |
| --- | --- | --- |
| Saldo medio (rotación) | `ventas 12m × plazo / 365` | El clásico. Correcto si el cliente compra parejo todo el año. |
| Exposición observada | pico y percentil 95 de la serie diaria | Lo que realmente ocupó. En campaña gruesa concentrada, el pico puede ser **3 o 4 veces** el saldo medio. |

La base es `máx(P95, saldo medio)`. Se usa el **P95 y no el pico** para que un mes
excepcional no fije el límite de todo el año; el pico queda en el Excel como
referencia. Con `--base max` se usa el pico, con `--base media` sólo la rotación.

**Paso 3 — Factor de antigüedad (γ).** Un cliente sin historia no puede recibir la
misma línea que uno de cinco campañas:

| Relación | Factor |
| --- | --- |
| ≥ 24 meses y ≥ 2 campañas | 1,00 |
| 12 a 24 meses | 0,90 |
| 6 a 12 meses | 0,75 |
| < 6 meses | 0,60 |

**Paso 4 — Holgura de crecimiento.** ×1,10 por defecto (`--crecimiento`), para que la
línea no frene una campaña que crece 10 %.

**Límite = base × γ × holgura**, redondeado al escalón comercial de arriba
(500 / 1.000 / 5.000 / 10.000 USD según el tamaño).

**Paso 5 — Topes de cartera.**

- `--cap-concentracion 0.10` (default): ningún cliente por encima del 10 % de la
  cartera de crédito. Si `% × nº de clientes < 1` el tope es imposible de cumplir y
  se ignora con aviso, en lugar de aplastar a todos al mismo número.
- `--capacidad <USD>`: techo global de financiación de Philagro. Si la suma de
  líneas lo supera, se prorratea todo proporcionalmente.

## Qué NO entra en el cálculo automático

- **CONTADO** (plazo 0): no ocupa crédito, no genera línea.
- **Condiciones sin plazo numérico** (CANJE, GRANOS, PLATAFORMA, TARJETA,
  COMPENSACION…): son exposición real pero se cancelan de otra forma (grano a
  cosecha, un tercero que asume el riesgo). No se les puede aplicar el modelo de
  días, así que se informan aparte, en su propia columna, con alerta. **Se deciden a
  mano.**

## Columnas del Excel

`Nro cliente` · `Cliente` · `Vendedor` · `Ventas 12m USD` (abierto en a crédito /
contado / sin plazo) · `Plazo pond. (días)` · `Ciclos/año` · `Exposición pico` ·
`Exposición P95` · `Saldo al corte` · `Saldo medio (rotación)` · `Base de cálculo` ·
`Antigüedad (meses)` · `Factor antigüedad` · **`LÍMITE SUGERIDO USD`** ·
`Ventas anuales que soporta` · `% de la cartera` · `Tope aplicado` · `Alertas`

**`Ventas anuales que soporta`** = límite × ciclos/año. Es la lectura comercial del
número: con esa línea, hasta cuánto puede comprar el cliente en el año sin pedir
excepción.

### Alertas que marca

- ventas sin plazo definido (canje/granos/plataforma): decidir aparte
- opera sólo contado: no requiere línea
- cliente nuevo (< 12 meses de relación)
- muy estacional: el pico triplica al saldo medio
- neto negativo en 12m (NC > facturas): revisar

## Sobre el "Nro de cliente"

El Excel de facturación trae `Cliente` como **texto**. Si el nombre viene con el
código adelante (`1001 - AGRO DEL SUR SRL`), el script lo separa en `Nro cliente` +
`Cliente`. Si no, la columna queda vacía: para tener el número de cuenta real hay que
exportar el reporte con esa columna incluida.

## Parámetros

| Flag | Default | Qué hace |
| --- | --- | --- |
| `--corte AAAA-MM-DD` | última factura | Fecha de corte del análisis. |
| `--ventana N` | 24 | Meses de historia para medir exposición (2 campañas). |
| `--base p95\|max\|media` | p95 | Qué medida usar como base. |
| `--crecimiento F` | 1.10 | Holgura para crecimiento. |
| `--cap-concentracion F` | 0.10 | Tope por cliente sobre la cartera. 0 = sin tope. |
| `--capacidad USD` | 0 | Techo global de financiación. 0 = sin techo. |
| `--min-ventas USD` | 0 | Ignora clientes por debajo de ese volumen. |
| `--salida ruta.xlsx` | auto | Ruta del Excel de salida. |

## PRIVACIDAD

Igual que el resto del repo: los datos reales **nunca** se commitean. El Excel de
salida y el de facturación quedan excluidos por `.gitignore` y se entregan al usuario
**por chat**. `generar_demo_ventas.py` crea datos ficticios para probar el motor.
