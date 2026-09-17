#!/usr/bin/env python3
"""Verifica el motor de exposicion contra casos de calculo manual."""
import sys
from datetime import datetime
from pathlib import Path
import numpy as np
sys.path.insert(0, str(Path(__file__).resolve().parent))
import pandas as pd
from asignar_creditos import (serie_exposicion, parsear_dias, escalonar,
                               analizar, etiqueta_campania)

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
res, warmup = analizar(df, pd.Timestamp("2025-04-01"), pd.Timestamp("2026-03-31"),
                       "p95", 1.0, usar_gamma=False, mes_campania=4)
r = res.iloc[0]
chk("dias de arrastre", warmup, 59)
chk("saldo al abrir la campania", r["exp_apertura"], 1000)
chk("pico = arrastre + factura nueva", r["exp_pico"], 1600)
chk("ventas DEL periodo (no incluye la de febrero)", r["ventas_total_p"], 600)

print("\n11. Sin arrastre el pico se subestima (misma data, periodo completo)")
res2, w2 = analizar(df, pd.Timestamp("2025-05-01"), pd.Timestamp("2026-03-31"),
                    "p95", 1.0, usar_gamma=False, mes_campania=4)
chk("pico con la de febrero ya arrastrada", res2.iloc[0]["exp_pico"], 1600)

print("\n12. Etiqueta de campania (abr-mar): feb-2026 cae en la campania 2025")
et = etiqueta_campania(pd.to_datetime(pd.Series(
    ["2025-02-15", "2025-04-01", "2026-02-15", "2026-04-01"])), 4)
for i, esp in enumerate([2024, 2025, 2025, 2026]):
    ok = int(et.iloc[i]) == esp
    print(f"  [{'OK ' if ok else 'MAL'}] {et.index[i]}: campania {int(et.iloc[i])} (esperado {esp})")
    if not ok: fallas.append(f"campania[{i}]")

print("\n" + ("TODO OK" if not fallas else f"FALLARON {len(fallas)}: {fallas}"))
sys.exit(1 if fallas else 0)
