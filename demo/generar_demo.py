#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera un Excel de cobranzas FICTICIO con la misma estructura que el real.

    python3 demo/generar_demo.py demo/Cobranzas_demo.xlsx

Sirve para probar el flujo y para versionar un dashboard de ejemplo en el repo
sin subir datos de clientes (ver PRIVACIDAD en CLAUDE.md).
"""
import datetime as dt
import random
import sys

from openpyxl import Workbook

CLIENTES = [
    "AGRO DEMO S.A.", "CAMPOS DEL SUR S.R.L.", "SEMILLAS EJEMPLO S.A.",
    "COOPERATIVA MODELO LTDA", "INSUMOS PRUEBA S.A.S.", "ACOPIO FICTICIO S.A.",
    "DISTRIBUIDORA TESTIGO S.R.L.", "ESTANCIA MUESTRA S.A.",
]
# (cuenta contable, peso relativo, ¿es un valor con vencimiento?)
CUENTAS = [
    ("Valores a depositar - echeq", 34, True),
    ("Valores en recaudadoras", 6, True),
    ("Banco Nación cta cte", 22, False),
    ("Banco Galicia transferencias", 14, False),
    ("Tarjeta / Mercado Pago", 5, False),
    ("Retenciones IIBB", 8, False),
    ("Percepciones IVA", 4, False),
    ("Caja Chica", 2, False),
    ("Redondeo", 1, False),
    ("Caja de Compensación", 4, False),
]


def main() -> None:
    salida = sys.argv[1] if len(sys.argv) > 1 else "demo/Cobranzas_demo.xlsx"
    rnd = random.Random(20260814)  # semilla fija: el demo es reproducible

    wb = Workbook()
    hoja = wb.active
    hoja.title = "Cobranzas"
    hoja.append(["Listado de cobranzas — DATOS DE EJEMPLO"])
    hoja.append([])
    hoja.append(["Fecha", "Documento", "Cliente", "Cuenta", "Importe",
                 "Importe moneda secundaria", "Fechavto"])

    cuentas = [c for c in CUENTAS for _ in range(c[1])]
    inicio, fin = dt.date(2025, 9, 1), dt.date(2026, 8, 14)
    dias = (fin - inicio).days
    nro = 0
    for _ in range(320):
        fecha = inicio + dt.timedelta(days=rnd.randint(0, dias))
        if fecha.weekday() >= 5:
            fecha -= dt.timedelta(days=2)
        nro += 1
        doc = f"R-00001-{nro:08d}"
        cliente = rnd.choice(CLIENTES)
        for _ in range(rnd.randint(1, 3)):
            cuenta, _peso, es_valor = rnd.choice(cuentas)
            usd = round(rnd.uniform(400, 45000), 2)
            if "Retenciones" in cuenta or "Percepciones" in cuenta:
                usd = round(usd * 0.02, 2)
            if rnd.random() < 0.03:      # anulaciones: ya vienen negativas
                usd = -usd
            fvto = fecha + dt.timedelta(days=rnd.choice([15, 30, 45, 60, 90, 120])) if es_valor else None
            hoja.append([fecha, doc, cliente, cuenta, round(usd * 1350, 2), usd, fvto])

    # Ajustes sin número de documento (se deduplican por contenido).
    for mes in (11, 12, 1, 2, 3):
        anio = 2025 if mes >= 9 else 2026
        hoja.append([dt.date(anio, mes, 28), None, rnd.choice(CLIENTES),
                     "Caja de Compensación", 0, round(rnd.uniform(-9000, 9000), 2), None])

    for col, ancho in zip("ABCDEFG", (12, 20, 34, 30, 16, 24, 12)):
        hoja.column_dimensions[col].width = ancho
    wb.save(salida)
    print(f"Escrito {salida} · {hoja.max_row - 3} movimientos · {nro} recibos")


if __name__ == "__main__":
    main()
