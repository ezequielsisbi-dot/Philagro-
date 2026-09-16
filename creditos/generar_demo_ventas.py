#!/usr/bin/env python3
"""Genera un Excel de facturacion FICTICIO con la estructura real, para probar
asignar_creditos.py sin usar datos del cliente. No son datos de Philagro."""

import random
from pathlib import Path
import pandas as pd

random.seed(7)
BASE = Path(__file__).resolve().parent

CLIENTES = [
    ("1001 - AGRO DEL SUR SRL", ["30-60 DIAS", "90 DIAS"], 8),
    ("1002 - ESTANCIA LA MARIA SA", ["120 DIAS", "90-120 DIAS"], 5),
    ("1003 - COOP. AGRICOLA CENTRO", ["CONTADO"], 12),
    ("1004 - DON PEDRO AGROPECUARIA", ["45-75-105"], 4),
    ("1005 - SEMILLERO NORTE SRL", ["30 DIAS", "CANJE"], 7),
    ("1006 - CAMPO VERDE SA", ["180 DIAS"], 3),
    ("1007 - HNOS GOMEZ SH", ["60 DIAS"], 6),
]
PRODUCTOS = ["GLIFOSATO 62%", "ATRAZINA 90", "CIPERMETRINA 25", "FERTILIZANTE NPK"]
VENDEDORES = ["J. PEREZ", "M. LOPEZ", "R. SOSA"]
# Campania gruesa concentrada: set-dic mucho mas cargado que el resto.
PESO_MES = {1: .4, 2: .3, 3: .6, 4: .8, 5: .7, 6: .4, 7: .3,
            8: .6, 9: 1.8, 10: 2.5, 11: 2.2, 12: 1.2}

filas = []
nro = 5000
for anio in (2024, 2025, 2026):
    for mes in range(1, 13):
        if anio == 2026 and mes > 8:
            continue
        for cliente, condiciones, ops in CLIENTES:
            n = max(0, round(ops * PESO_MES[mes] * random.uniform(.5, 1.4) / 3))
            for _ in range(n):
                nro += 1
                dia = random.randint(1, 28)
                cond = random.choice(condiciones)
                for _ in range(random.randint(1, 3)):
                    cant = random.randint(10, 400)
                    precio = round(random.uniform(4, 28), 2)
                    filas.append({
                        "Fechacomprobante": f"{anio}-{mes:02d}-{dia:02d}",
                        "Transacconsubtiponombre": "Factura",
                        "Comprobante": nro,
                        "Cliente": cliente,
                        "Costo": 0,
                        "Condicionpago": cond,
                        "Moneda": "USD",
                        "Vendedor": random.choice(VENDEDORES),
                        "Producto": random.choice(PRODUCTOS),
                        "Cantidad": cant,
                        "Precio": precio,
                        "Importemonsecundaria": round(cant * precio, 2),
                        "Principioactivo": "DEMO",
                    })

# Un par de notas de credito (negativas) para verificar que netean solas.
for _ in range(6):
    f = random.choice(filas).copy()
    f["Comprobante"] = f["Comprobante"] + 900000
    f["Transacconsubtiponombre"] = "Nota de Credito"
    f["Importemonsecundaria"] = -abs(f["Importemonsecundaria"]) * 0.3
    filas.append(f)

destino = BASE / "demo_facturacion.xlsx"
pd.DataFrame(filas).to_excel(destino, index=False)
print(f"{destino}  ({len(filas)} filas)")
