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

### El percentil va sobre los días CON saldo

No sobre el calendario. Es la diferencia entre *"¿cuánto debe un día cualquiera del
año?"* y *"cuando nos debe, ¿cuánto nos debe?"*. Para sizear un límite, la pregunta
correcta es la segunda.

Sobre el calendario, un cliente que compra dos veces al año a 15 días tiene saldo 15
días de 365: el percentil 95 cae en un día de saldo cero y el límite se va a cero. Caso
real de la cartera: un cliente facturado por US$ 9.240 en una sola operación salía con
una línea de US$ 500 — no se le puede vender ni lo que ya se le vendió. Sobre los días
con saldo sale US$ 7.000.

Para los clientes activos casi todo el año el cambio es marginal (en la cartera real,
US$ 111.343 → US$ 113.176 en el mayor). Sólo corrige a los compradores esporádicos, que
es donde el percentil sobre calendario no significaba nada.

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

## Trabajar por campaña

Philagro cierra campaña **1 de abril a 31 de marzo** (`campaña 2025/26` = abr-2025 a
mar-2026), así que hay que pasar `--inicio-campania 4`. El default del script es 7
(julio-junio, el año agrícola estándar de INDEC).

```bash
python3 asignar_creditos.py VENTAS.xlsx --campania 2025/26 --inicio-campania 4
```

La campaña es el período correcto para medir crédito: el año calendario parte la
gruesa al medio y deja la venta de sep-dic en un año y su cobranza (a 180 días,
contra cosecha) en el siguiente.

El script imprime siempre la **facturación por mes calendario** y sugiere el corte
según los datos: la campaña arranca el mes siguiente al de menor facturación.

### El arrastre entre campañas

Se mide sobre la campaña elegida, pero **arrastrando las facturas anteriores**. La
distinción importa: una factura de la campaña previa que sigue abierta el 1 de abril
ocupa crédito y tiene que contar. Si se filtran los datos a la campaña y listo, el
arranque muestra menos saldo del real — justo cuando la gruesa vendida a 180 días
todavía no se cobró.

Por eso la serie de exposición se arma con **todo** el historial del archivo y recién
después se recorta al período que se informa:

- `Ventas del período` → sólo lo facturado dentro de la campaña.
- `Exposición pico / P95` → incluye lo que venía abierto de antes.
- `Saldo al abrir el período` → cuánto arrastraba el cliente el primer día. Si es > 0
  queda marcado en Alertas.

El script informa **cuántos días de arrastre** tuvo disponibles. Si ese arrastre es
menor al plazo más largo de la cartera, avisa cuántos días del arranque quedan
subestimados igual. Con dos campañas cargadas y analizando la segunda, el arrastre es
de 365 días y el problema desaparece.

### Cuando hay una sola campaña

La antigüedad **no se puede medir**: con menos de 18 meses no hay forma de distinguir
un cliente de cinco campañas de uno que entró este año, así que γ se **neutraliza en
1,00 para todos** (con aviso) en vez de castigar a toda la cartera por una limitación
del archivo. Se fuerza con `--antiguedad on|off`.

La cuenta de campañas usa el mes de inicio configurado, no el año calendario: con
`--inicio-campania 4`, una compra de feb-2026 pertenece a la campaña 2025.

## Condiciones en cuotas

`30 - 60 - 90 DIAS` significa tres pagos, no un pago a 90. Tomar el máximo —como
hace el dashboard de ventas— **sobreestima la exposición**: a los 60 días ya cobraste
dos tercios. Acá cada cuota ocupa cupo por su propio plazo, en partes iguales, que es
la convención de plaza. El `plazo ponderado` que se informa es el promedio de las
cuotas.

Es una divergencia **deliberada** con el criterio congelado de ventas, porque el
propósito es otro: allá se informa el plazo más largo otorgado, acá se mide crédito
ocupado.

## Actividad propia vs. arrastre

Se calculan dos series de exposición, porque responden preguntas distintas:

| Serie | Qué incluye | Para qué |
| --- | --- | --- |
| **Con arrastre** | facturas previas todavía abiertas + las del período | Qué riesgo se corrió de verdad. Columna `Pico con arrastre`. |
| **Actividad propia** | sólo lo facturado dentro del período | Qué exposición genera el cliente hoy. **Es la que fija el límite.** |

El límite lo fija la actividad propia porque **mira para adelante**. Un cliente que
compró fuerte la campaña pasada y este año casi no compró arrastra un pico alto que ya
se está liquidando: darle línea por ese pico es financiar una retirada. Caso real de
la cartera: un cliente con US$ 696.000 comprados a 270-360 días en 2024/25 y sólo
US$ 70.000 en 2025/26 mostraba un pico con arrastre de US$ 608.000. Por actividad
propia le corresponden US$ 65.000.

Cuando las dos medidas divergen más de 50 %, se marca en Alertas (`viene bajando` /
`viene creciendo fuerte`).

## Días sobre el límite

`Días/año sobre el límite` cuenta, sobre la **actividad propia**, cuántos días del
período el saldo habría superado el límite propuesto. Cada uno de esos días es un
pedido de excepción: es la medida operativa de cuánta fricción genera la línea.

Se usa la actividad propia y no el saldo con arrastre porque el límite se aplica a la
campaña siguiente, cuando el arrastre ya se liquidó. Para el arrastre está la alerta
`ARRANCA EXCEDIDO`, que marca a los clientes cuyo saldo al cierre ya supera el límite
propuesto.

## Qué NO entra en el cálculo automático## Columnas del Excel

`Nro cliente` · `Cliente` · `Vendedor` · `Ventas 12m USD` (abierto en a crédito /
contado / sin plazo) · `Plazo pond. (días)` · `Ciclos/año` · `Exposición pico` ·
`Exposición P95` · `Saldo al corte` · `Saldo medio (rotación)` · `Base de cálculo` ·
`Saldo al abrir el período` · `Ventas campaña anterior` · `Var. vs campaña anterior` ·
`Antigüedad (meses)` · `Campañas` · `Factor antigüedad` · **`LÍMITE SUGERIDO USD`** ·
`Ventas anuales que soporta` · `Días/año sobre el límite` · `% de la cartera` · `Tope aplicado` · `Alertas`

**`Ventas anuales que soporta`** = límite × ciclos/año. Es la lectura comercial del
número: con esa línea, hasta cuánto puede comprar el cliente en el año sin pedir
excepción.

### Alertas que marca

- ventas sin plazo definido (canje/granos/plataforma): decidir aparte
- opera sólo contado: no requiere línea
- cliente nuevo (< 12 meses de relación)
- muy estacional: el pico triplica al saldo medio
- neto negativo en el período (NC > facturas): revisar
- abrió el período con US$ X de la campaña anterior
- compró en una sola campaña
- cayó X% / creció X% vs la campaña anterior
- ARRANCA EXCEDIDO: al cierre debía más que el límite propuesto

## Sobre el "Nro de cliente"

El Excel de facturación trae `Cliente` como **texto**. Si el nombre viene con el
código adelante (`1001 - AGRO DEL SUR SRL`), el script lo separa en `Nro cliente` +
`Cliente`. Si no, la columna queda vacía: para tener el número de cuenta real hay que
exportar el reporte con esa columna incluida.

## Parámetros

| Flag | Default | Qué hace |
| --- | --- | --- |
| `--corte AAAA-MM-DD` | última factura | Fecha de corte del análisis. |
| `--campania 2025/26` | — | Analiza una campaña en vez de toda la historia. |
| `--inicio-campania MM` | 7 (julio) | Mes en que arranca la campaña. **Philagro: 4.** |
| `--antiguedad auto\|on\|off` | auto | Factor γ. `auto` = sólo si hay ≥ 18 meses de datos. |
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
