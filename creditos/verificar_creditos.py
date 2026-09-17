#!/usr/bin/env python3
"""Verifica el motor de exposicion contra casos de calculo manual."""
import sys
from datetime import datetime
from pathlib import Path
import numpy as np
sys.path.insert(0, str(Path(__file__).resolve().parent))
import pandas as pd
from asignar_creditos import (serie_exposicion, parsear_dias, escalonar,
                               analizar, etiqueta_campania, parsear_cuotas)

def d(s): return datetime.strptime(s, "%Y-%m-%d")
fallas = []
def chk(nombre, obtenido, esperado):
    ok = abs(obtenido - esperado) < 1e-6
    print(f"  [{'OK ' if ok else 'MAL'}] {nombre}: {obtenido:,.2f} (esperado {esperado:,.2f})")
    if not ok: fallas.append(nombre)

print("1. Una factura de 1.000 a 60 dias")
s = serie_exposicion([(d("2026-01-10"), 60, 1000)], d("2026-01-01"), d("2026-06-30"))
chk("pico", s.max(), 1000)
chk("ocupa exactamente 60 dias", (s > 0).sum(), 60)
chk("libre antes de facturar", s[8], 0)
chk("libre al vencer (dia 61)", s[(d("2026-03-11") - d("2026-01-01")).days], 0)

print("\n2. Dos facturas DENTRO del plazo se acumulan")
s = serie_exposicion([(d("2026-01-10"), 60, 1000), (d("2026-02-10"), 60, 800)],
                     d("2026-01-01"), d("2026-06-30"))
chk("pico = 1.000 + 800", s.max(), 1800)

print("\n3. Dos facturas FUERA del plazo NO se acumulan (se renovo el cupo)")
s = serie_exposicion([(d("2026-01-10"), 60, 1000), (d("2026-05-10"), 60, 800)],
                     d("2026-01-01"), d("2026-12-31"))
chk("pico = solo la mayor", s.max(), 1000)

print("\n4. Nota de credito libera cupo")
s = serie_exposicion([(d("2026-01-10"), 60, 1000), (d("2026-01-20"), 60, -300)],
                     d("2026-01-01"), d("2026-06-30"))
chk("pico inicial", s.max(), 1000)
chk("saldo tras la NC", s[(d("2026-01-25") - d("2026-01-01")).days], 700)

print("\n5. CONTADO no ocupa credito")
s = serie_exposicion([(d("2026-01-10"), 0, 5000)], d("2026-01-01"), d("2026-06-30"))
chk("pico", s.max(), 0)

print("\n6. Plazos mixtos: cada factura vence con SU condicion")
s = serie_exposicion([(d("2026-01-10"), 30, 500), (d("2026-01-10"), 180, 500)],
                     d("2026-01-01"), d("2026-12-31"))
chk("pico ambas juntas", s.max(), 1000)
chk("tras vencer la de 30 queda la de 180", s[(d("2026-03-01") - d("2026-01-01")).days], 500)

print("\n7. Formula de rotacion: compra 1.000/mes a 90 dias -> saldo ~3.000")
fac = [(d(f"2026-{m:02d}-01"), 90, 1000) for m in range(1, 13)]
s = serie_exposicion(fac, d("2026-01-01"), d("2026-12-31"))
chk("saldo estabilizado", s[(d("2026-06-15") - d("2026-01-01")).days], 3000)
chk("rotacion: 12.000 * 90/365", round(12000 * 90 / 365), 2959)

print("\n8. parsear_dias (criterio congelado de ventas)")
for cond, esp in [("CONTADO", 0), ("30-60 DIAS", 60), ("45-75-105", 105),
                  ("180 DIAS", 180), ("CANJE", None), ("", None)]:
    obt = parsear_dias(cond)
    ok = obt == esp
    print(f"  [{'OK ' if ok else 'MAL'}] {cond or '(vacio)':<12} -> {obt} (esperado {esp})")
    if not ok: fallas.append(f"parsear_dias({cond})")

print("\n9. Escalones de redondeo")
for monto, esp in [(0, 0), (1_234, 1_500), (12_100, 13_000), (61_200, 65_000), (340_000, 340_000)]:
    chk(f"escalonar({monto:,})", escalonar(monto), esp)

print("\n10. Arrastre: la factura de la campania anterior sigue ocupando cupo")
# Campania abr-mar. Factura del 01/02/2025 a 180 dias -> vence 31/07/2025, o sea
# que sigue abierta cuando arranca la campania 2025/26 el 01/04/2025.
df = pd.DataFrame([
    {"fecha": "2025-02-01", "comprobante": 1, "cliente": "X", "vendedor": "V",
     "condicionpago": "180 DIAS", "importe": 1000},
    {"fecha": "2025-05-01", "comprobante": 2, "cliente": "X", "vendedor": "V",
     "condicionpago": "180 DIAS", "importe": 600},
])
df["fecha"] = pd.to_datetime(df["fecha"])
res, warmup, _ = analizar(df, pd.Timestamp("2025-04-01"), pd.Timestamp("2026-03-31"),
                       "p95", 1.0, usar_gamma=False, mes_campania=4)
r = res.iloc[0]
chk("dias de arrastre", warmup, 59)
chk("saldo al abrir la campania", r["exp_apertura"], 1000)
chk("pico CON arrastre (riesgo real corrido)", r["exp_pico_total"], 1600)
chk("pico de la campania actual (sin arrastre)", r["exp_pico"], 600)
chk("ventas DEL periodo (no incluye la de febrero)", r["ventas_total_p"], 600)

print("\n11. El limite NO lo fija el arrastre: un cliente que se retira")
# Compro fuerte antes del periodo y casi nada dentro: el pico historico es alto,
# pero la linea tiene que seguir a la actividad actual.
df2 = pd.DataFrame([
    {"fecha": "2025-02-01", "comprobante": 1, "cliente": "Y", "vendedor": "V",
     "condicionpago": "270 DIAS", "importe": 500000},
    {"fecha": "2025-09-01", "comprobante": 2, "cliente": "Y", "vendedor": "V",
     "condicionpago": "270 DIAS", "importe": 40000},
])
df2["fecha"] = pd.to_datetime(df2["fecha"])
res2, _, _s2 = analizar(df2, pd.Timestamp("2025-04-01"), pd.Timestamp("2026-03-31"),
                   "max", 1.0, usar_gamma=False, mes_campania=4)
r2 = res2.iloc[0]
chk("riesgo real corrido", r2["exp_pico_total"], 540000)
chk("base del limite = solo la campania actual", r2["exp_pico"], 40000)
chk("el limite sigue a la actividad, no al arrastre", r2["limite_calc"], 40000)
ok = "viene bajando" in r2["alertas"]
print(f"  [{'OK ' if ok else 'MAL'}] avisa que viene bajando: {r2['alertas'][:70]}")
if not ok: fallas.append("alerta viene bajando")

print("\n12. Etiqueta de campania (abr-mar): feb-2026 cae en la campania 2025")
et = etiqueta_campania(pd.to_datetime(pd.Series(
    ["2025-02-15", "2025-04-01", "2026-02-15", "2026-04-01"])), 4)
for i, esp in enumerate([2024, 2025, 2025, 2026]):
    ok = int(et.iloc[i]) == esp
    print(f"  [{'OK ' if ok else 'MAL'}] {et.index[i]}: campania {int(et.iloc[i])} (esperado {esp})")
    if not ok: fallas.append(f"campania[{i}]")

print("\n13. Cuotas: 30-60-90 ocupa menos cupo que 90 DIAS de una")
df = pd.DataFrame([{"fecha": "2025-06-01", "comprobante": 1, "cliente": "X",
                    "vendedor": "V", "condicionpago": "30 - 60 - 90 DIAS", "importe": 900}])
df["fecha"] = pd.to_datetime(df["fecha"])
res, _, _s = analizar(df, pd.Timestamp("2025-04-01"), pd.Timestamp("2026-03-31"),
                  "max", 1.0, usar_gamma=False, mes_campania=4)
chk("pico = las 3 cuotas juntas al inicio", res.iloc[0]["exp_pico"], 900)
chk("plazo representativo = promedio (30+60+90)/3", res.iloc[0]["plazo_pond"], 60)
s60 = serie_exposicion([(datetime(2025, 6, 1), 30, 300), (datetime(2025, 6, 1), 60, 300),
                        (datetime(2025, 6, 1), 90, 300)], d("2025-06-01"), d("2025-12-31"))
chk("a los 45 dias quedan 2 cuotas", s60[45], 600)
chk("a los 75 dias queda 1 cuota", s60[75], 300)
uno = serie_exposicion([(datetime(2025, 6, 1), 90, 900)], d("2025-06-01"), d("2025-12-31"))
chk("a 90 dias de una, a los 75 sigue entero", uno[75], 900)

print("\n14. parsear_cuotas")
for cond, esp in [("CONTADO", [0]), ("60 DIAS", [60]), ("30 - 60 - 90 DIAS", [30, 60, 90]),
                  ("45-75-105", [45, 75, 105]), ("180 DIAS PESIFICADO", [180]),
                  ("CANJE/COMPENSACION", None), ("PLATAFORMA NERA", None)]:
    obt = parsear_cuotas(cond)
    ok = obt == esp
    print(f"  [{'OK ' if ok else 'MAL'}] {cond:<22} -> {obt} (esperado {esp})")
    if not ok: fallas.append(f"parsear_cuotas({cond})")

print("\n" + ("TODO OK" if not fallas else f"FALLARON {len(fallas)}: {fallas}"))
sys.exit(1 if fallas else 0)
