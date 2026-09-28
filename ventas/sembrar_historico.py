#!/usr/bin/env python3
"""
sembrar_historico.py — Siembra el histórico a partir de un dashboard ya generado.

Uso:
    python3 sembrar_historico.py <dashboard_ventas_VIEJO.html>

Sirve para arrancar el histórico con todo lo que ya estaba en un dashboard
anterior, sin tener que volver a procesar todos los Excel viejos. Después de
esto, cada Excel nuevo ("julio a la fecha") se fusiona incrementalmente con
`ventas_dashboard.py` sin duplicar nada.

Lee los datos embebidos del HTML (REGISTROS + catálogos) y los reexpande al
formato de `data/historico_ventas.json`.

Se niega a pisar un histórico existente salvo que se pase --forzar.
"""

import json
import re
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
HISTORICO_PATH = DATA_DIR / "historico_ventas.json"


def extraer_const(html: str, nombre: str):
    """Extrae `const NOMBRE = <json>;` del HTML embebido."""
    patron = re.compile(r"const\s+" + nombre + r"\s*=\s*", re.MULTILINE)
    m = patron.search(html)
    if not m:
        raise SystemExit(f"ERROR: no se encontró 'const {nombre}' en el HTML.")
    inicio = m.end()
    # El valor termina en el ';' que cierra la sentencia, al final de la línea.
    fin = html.index(";\n", inicio)
    return json.loads(html[inicio:fin])


def main():
    args = [a for a in sys.argv[1:] if a != "--forzar"]
    forzar = "--forzar" in sys.argv[1:]
    if len(args) != 1:
        raise SystemExit("Uso: python3 sembrar_historico.py <dashboard_ventas_VIEJO.html> [--forzar]")

    html_path = Path(args[0]).resolve()
    if not html_path.exists():
        raise SystemExit(f"ERROR: no existe el archivo {html_path}")

    if HISTORICO_PATH.exists() and not forzar:
        raise SystemExit(
            f"ERROR: ya existe {HISTORICO_PATH}.\n"
            "  Sembrar lo pisaría entero. Si es lo que querés, volvé a correr con --forzar."
        )

    html = html_path.read_text(encoding="utf-8")

    registros = extraer_const(html, "REGISTROS")
    vendedores = extraer_const(html, "VENDEDORES")
    productos = extraer_const(html, "PRODUCTOS")
    clientes = extraer_const(html, "CLIENTES")
    principios = extraer_const(html, "PRINCIPIOS")
    condiciones = extraer_const(html, "CONDICIONES")

    filas = []
    for r in registros:
        filas.append({
            "fecha": r["f"],
            "mes": r["m"],
            "comprobante": r["o"],
            "vendedor": vendedores[r["v"]],
            "producto": productos[r["p"]],
            "cliente": clientes[r["c"]],
            "principioactivo": principios[r["pa"]],
            "condicionpago": condiciones[r["cp"]]["n"],
            "cantidad": r["q"],
            "precio_unitario": r["pu"],
            "importe": r["imp"],
        })

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with open(HISTORICO_PATH, "w", encoding="utf-8") as f:
        json.dump(filas, f, ensure_ascii=False, separators=(",", ":"))

    meses = sorted({f["mes"] for f in filas})
    comprobantes = {f["comprobante"] for f in filas}
    print(f"Histórico sembrado desde {html_path.name}")
    print(f"  Archivo:      {HISTORICO_PATH}")
    print(f"  Filas:        {len(filas)}")
    print(f"  Comprobantes: {len(comprobantes)}")
    print(f"  Meses:        {meses[0]} a {meses[-1]}  ({len(meses)} meses)")
    print("\nYa podés correr: python3 ventas_dashboard.py <FACTURACION_NUEVA.xlsx>")


if __name__ == "__main__":
    main()
