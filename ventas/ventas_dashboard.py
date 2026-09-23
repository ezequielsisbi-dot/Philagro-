#!/usr/bin/env python3
"""
ventas_dashboard.py — Genera el Dashboard de Ventas de Philagro (HTML autocontenido)
a partir de un Excel de facturación.

Uso:
    python3 ventas_dashboard.py <FACTURACION.xlsx>

Proyecto INDEPENDIENTE del dashboard de cobranzas: no comparte ni pisa sus archivos.
Todo lo que produce vive bajo ventas/ (este directorio).

Salida:
    dashboard_ventas_<nombre-del-excel>.html   (nunca pisa uno existente: _v2, _v3, …)

Histórico incremental:
    ventas/data/historico_ventas.json — acumula TODAS las filas ya vistas.
    Cada corrida detecta qué meses trae el Excel nuevo, descarta la versión vieja
    de esos meses del histórico y la reemplaza por la del Excel nuevo (se asume
    más completa), sumando además los meses que no existían todavía.
"""

import json
from collections import Counter
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "assets" / "plantilla_ventas.html"
DATA_DIR = BASE_DIR / "data"
HISTORICO_PATH = DATA_DIR / "historico_ventas.json"

# Chart.js se inyecta INLINE en el HTML final (el dashboard queda autocontenido,
# sin CDN ni fetch). Se busca primero una copia vendorizada en assets/vendor/ y,
# si no está, la instalada por npm.
CHARTJS_PATHS = [
    BASE_DIR / "assets" / "vendor" / "chart.umd.min.js",
    BASE_DIR / "node_modules" / "chart.js" / "dist" / "chart.umd.min.js",
    BASE_DIR / "node_modules" / "chart.js" / "dist" / "chart.umd.js",
]

REQUIRED_COLS = [
    "Fechacomprobante", "Transacconsubtiponombre", "Comprobante", "Cliente",
    "Costo", "Condicionpago", "Moneda", "Vendedor", "Producto", "Cantidad",
    "Precio", "Importemonsecundaria", "Principioactivo",
]

SIN_PRINCIPIO = "(sin principio activo)"

MESES_ABBR = ["ene.", "feb.", "mar.", "abr.", "may.", "jun.",
              "jul.", "ago.", "sep.", "oct.", "nov.", "dic."]

NO_NUMERICO_CONTADO = {"CONTADO"}


# ---------------------------------------------------------------------------
# Lectura y limpieza del Excel
# ---------------------------------------------------------------------------

def leer_excel(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, engine="openpyxl")
    faltantes = [c for c in REQUIRED_COLS if c not in df.columns]
    if faltantes:
        raise SystemExit(
            f"ERROR: al Excel le faltan columnas requeridas: {', '.join(faltantes)}"
        )
    return df


def parsear_dias(condicion: str):
    """CONTADO -> 0. Rango de dias ("30-60 DIAS", "45-75-105") -> el MAXIMO.
    No numerica (CANJE, PLATAFORMA, TARJETA, COMPENSACION, GRANOS, ...) -> None."""
    if condicion is None:
        return None
    texto = str(condicion).strip().upper()
    if not texto:
        return None
    numeros = [int(n) for n in re.findall(r"\d+", texto)]
    if numeros:
        return max(numeros)
    if texto in NO_NUMERICO_CONTADO:
        return 0
    return None


def precio_unitario(row) -> float:
    """Criterio congelado:
    - Transacconsubtiponombre == "Nota Liquido producto" -> usar Costo.
    - En cualquier otro caso (incluye Costo == 0, p.ej. Asesoramiento-Comisiones) -> usar Precio.
    """
    tipo = str(row["Transacconsubtiponombre"]).strip() if pd.notna(row["Transacconsubtiponombre"]) else ""
    if tipo == "Nota Liquido producto":
        return float(row["Costo"]) if pd.notna(row["Costo"]) else 0.0
    return float(row["Precio"]) if pd.notna(row["Precio"]) else 0.0


def normalizar_comprobante(v):
    try:
        if float(v).is_integer():
            return int(v)
    except (TypeError, ValueError):
        pass
    return str(v).strip()


def limpiar(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df = df.dropna(subset=["Fechacomprobante", "Comprobante", "Vendedor", "Producto", "Cliente"])

    out = pd.DataFrame()
    out["fecha"] = pd.to_datetime(df["Fechacomprobante"]).dt.strftime("%Y-%m-%d")
    out["mes"] = out["fecha"].str.slice(0, 7)
    out["comprobante"] = df["Comprobante"].map(normalizar_comprobante)
    out["vendedor"] = df["Vendedor"].fillna("").astype(str).str.strip()
    out["producto"] = df["Producto"].fillna("").astype(str).str.strip()
    out["cliente"] = df["Cliente"].fillna("").astype(str).str.strip()
    principio = df["Principioactivo"].fillna("").astype(str).str.strip()
    principio = principio.where(~principio.isin(["", "nan", "None"]), SIN_PRINCIPIO)
    out["principioactivo"] = principio
    out["condicionpago"] = df["Condicionpago"].fillna("").astype(str).str.strip()
    out["cantidad"] = pd.to_numeric(df["Cantidad"], errors="coerce").fillna(0.0)
    out["precio_unitario"] = df.apply(precio_unitario, axis=1)
    # Moneda = USD desde Importemonsecundaria, firmada tal cual viene (NC negativas se
    # netean solas al sumar). Si hay filas en PESOS se usa igual Importemonsecundaria.
    out["importe"] = pd.to_numeric(df["Importemonsecundaria"], errors="coerce").fillna(0.0)
    return out


# ---------------------------------------------------------------------------
# Histórico incremental (fusión sin duplicar)
# ---------------------------------------------------------------------------

def cargar_historico() -> list:
    if HISTORICO_PATH.exists():
        with open(HISTORICO_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def guardar_historico(filas: list):
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with open(HISTORICO_PATH, "w", encoding="utf-8") as f:
        json.dump(filas, f, ensure_ascii=False, separators=(",", ":"))


def _clave_contenido(fila: dict) -> tuple:
    """Identidad de una fila que sobrevive a cambios de numeracion.

    El historico sembrado desde un dashboard viejo no conserva el numero de
    comprobante real (guarda un indice), asi que comparar por `comprobante`
    entre historico y Excel nuevo no siempre es posible. El contenido si.
    """
    return (
        fila["fecha"],
        str(fila["cliente"]).strip().upper(),
        str(fila["producto"]).strip().upper(),
        round(float(fila["importe"]), 2),
    )


def fusionar(historico: list, nuevas: pd.DataFrame) -> list:
    """Reemplaza en el historico los meses que trae el Excel nuevo por su version
    (se asume mas completa), y agrega los meses que no existian. No toca los
    meses que el Excel nuevo no menciona.

    Verifica 0 filas perdidas POR CONTENIDO, no por cantidad: que un mes pase de
    50 a 130 comprobantes no garantiza que los 50 viejos esten entre los 130. Las
    filas del historico que el Excel nuevo no trae se CONSERVAN y se informan.
    """
    meses_nuevos = set(nuevas["mes"].unique())

    viejos_por_mes = {}
    for fila in historico:
        viejos_por_mes.setdefault(fila["mes"], set()).add(fila["comprobante"])

    conservados = [f for f in historico if f["mes"] not in meses_nuevos]
    nuevas_filas = nuevas.to_dict(orient="records")

    # Filas del historico, en los meses reemplazados, que el Excel nuevo no trae.
    pisados = [f for f in historico if f["mes"] in meses_nuevos]
    nuevas_por_clave = Counter(_clave_contenido(f) for f in nuevas_filas)
    vistas = Counter()
    rescatadas = []
    for fila in pisados:
        clave = _clave_contenido(fila)
        vistas[clave] += 1
        if vistas[clave] > nuevas_por_clave.get(clave, 0):
            rescatadas.append(fila)

    fusionado = conservados + nuevas_filas + rescatadas

    nuevos_por_mes = {}
    for fila in nuevas_filas:
        nuevos_por_mes.setdefault(fila["mes"], set()).add(fila["comprobante"])

    print("\n--- Fusion con historico ---")
    if not historico:
        print(f"Histórico vacío: se inicializa con {len(nuevas_filas)} filas / {len(meses_nuevos)} meses.")
    for mes in sorted(meses_nuevos):
        antes = len(viejos_por_mes.get(mes, set()))
        despues = len(nuevos_por_mes.get(mes, set()))
        if mes in viejos_por_mes:
            print(f"  {mes}: reemplazado — comprobantes {antes} -> {despues}")
        else:
            print(f"  {mes}: mes nuevo — {despues} comprobantes agregados.")

    if rescatadas:
        print(f"\n  ATENCION: {len(rescatadas)} fila(s) del histórico que el Excel nuevo NO trae.")
        print("  Se CONSERVAN (no se pierde nada). Revisar si fueron anuladas en el sistema:")
        for fila in sorted(rescatadas, key=lambda f: (f["fecha"], str(f["cliente"]))):
            print(f"    {fila['fecha']} | {str(fila['cliente'])[:26]:26} | "
                  f"{str(fila['producto'])[:26]:26} | {fila['importe']:>10.2f} | "
                  f"comp {fila['comprobante']}")
    else:
        print("  Verificación de pérdida: 0 filas del histórico quedaron fuera.")

    meses_conservados = sorted({f["mes"] for f in conservados})
    if meses_conservados:
        print(f"  Meses históricos sin tocar: {', '.join(meses_conservados)}")
    print(f"Total filas fusionadas: {len(fusionado)} (antes: {len(historico)})")
    return fusionado


# ---------------------------------------------------------------------------
# Construcción del payload para la plantilla
# ---------------------------------------------------------------------------

def construir_payload(filas: list) -> dict:
    df = pd.DataFrame(filas)

    vendedores = sorted(df["vendedor"].unique())
    productos = sorted(df["producto"].unique())
    clientes = sorted(df["cliente"].unique())
    principios = [SIN_PRINCIPIO] + sorted(
        p for p in df["principioactivo"].unique() if p != SIN_PRINCIPIO
    )

    condiciones_nombres = sorted(df["condicionpago"].unique())
    condiciones = [{"n": n, "d": parsear_dias(n)} for n in condiciones_nombres]

    idx_v = {n: i for i, n in enumerate(vendedores)}
    idx_p = {n: i for i, n in enumerate(productos)}
    idx_c = {n: i for i, n in enumerate(clientes)}
    idx_pa = {n: i for i, n in enumerate(principios)}
    idx_cp = {n: i for i, n in enumerate(condiciones_nombres)}

    registros = []
    for f in filas:
        registros.append({
            "f": f["fecha"],
            "m": f["mes"],
            "o": f["comprobante"],
            "v": idx_v[f["vendedor"]],
            "p": idx_p[f["producto"]],
            "c": idx_c[f["cliente"]],
            "pa": idx_pa[f["principioactivo"]],
            "cp": idx_cp[f["condicionpago"]],
            "q": f["cantidad"],
            "pu": round(f["precio_unitario"], 4),
            "imp": round(f["importe"], 2),
        })
    # Orden estable por fecha, como en el original.
    registros.sort(key=lambda r: (r["f"], r["o"] if isinstance(r["o"], int) else 0))

    meses_keys = sorted(df["mes"].unique())
    meses = []
    for k in meses_keys:
        anio, mes_num = k.split("-")
        meses.append({"key": k, "label": f"{MESES_ABBR[int(mes_num) - 1]} {anio}"})

    fechas = sorted(df["fecha"].unique())
    total_ops = df["comprobante"].nunique()

    meta = {
        "generado": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "n_registros": len(registros),
        "total_ops": int(total_ops),
        "rango_min": fechas[0] if fechas else "",
        "rango_max": fechas[-1] if fechas else "",
    }

    return {
        "REGISTROS": registros,
        "MESES": meses,
        "VENDEDORES": vendedores,
        "PRODUCTOS": productos,
        "CLIENTES": clientes,
        "PRINCIPIOS": principios,
        "CONDICIONES": condiciones,
        "META": meta,
    }


# ---------------------------------------------------------------------------
# Render de la plantilla
# ---------------------------------------------------------------------------

def cargar_chartjs() -> str:
    for p in CHARTJS_PATHS:
        if p.exists():
            return p.read_text(encoding="utf-8")
    raise SystemExit(
        "ERROR: no se encontró Chart.js para vendorizar inline.\n"
        f"  Buscado en:\n    " + "\n    ".join(str(p) for p in CHARTJS_PATHS) + "\n"
        "  Solución: correr `npm install` dentro de ventas/ (una sola vez)."
    )


def render_html(payload: dict) -> str:
    if not TEMPLATE_PATH.exists():
        raise SystemExit(f"ERROR: no se encuentra la plantilla en {TEMPLATE_PATH}")
    template = TEMPLATE_PATH.read_text(encoding="utf-8")

    def j(v):
        return json.dumps(v, ensure_ascii=False, separators=(",", ":"))

    bloque = (
        f'const REGISTROS  = {j(payload["REGISTROS"])};\n'
        f'const MESES       = {j(payload["MESES"])};\n'
        f'const VENDEDORES  = {j(payload["VENDEDORES"])};\n'
        f'const PRODUCTOS   = {j(payload["PRODUCTOS"])};\n'
        f'const CLIENTES    = {j(payload["CLIENTES"])};\n'
        f'const PRINCIPIOS  = {j(payload["PRINCIPIOS"])};\n'
        f'const CONDICIONES = {j(payload["CONDICIONES"])};\n'
        f'const META        = {j(payload["META"])};'
    )

    if "/*__DATOS_JS__*/" not in template:
        raise SystemExit("ERROR: la plantilla no tiene el placeholder /*__DATOS_JS__*/")
    if "/*__CHARTJS__*/" not in template:
        raise SystemExit("ERROR: la plantilla no tiene el placeholder /*__CHARTJS__*/")

    html = template.replace("/*__DATOS_JS__*/", bloque, 1)
    html = html.replace("/*__CHARTJS__*/", cargar_chartjs(), 1)
    return html


def siguiente_nombre_libre(base: Path) -> Path:
    if not base.exists():
        return base
    stem, suf = base.stem, base.suffix
    i = 2
    while True:
        candidato = base.with_name(f"{stem}_v{i}{suf}")
        if not candidato.exists():
            return candidato
        i += 1


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) != 2:
        raise SystemExit("Uso: python3 ventas_dashboard.py <FACTURACION.xlsx>")

    xlsx_path = Path(sys.argv[1]).resolve()
    if not xlsx_path.exists():
        raise SystemExit(f"ERROR: no existe el archivo {xlsx_path}")

    print(f"Leyendo {xlsx_path.name} ...")
    df_raw = leer_excel(xlsx_path)
    print(f"  {len(df_raw)} filas leídas.")

    df_limpio = limpiar(df_raw)
    print(f"  {len(df_limpio)} filas válidas tras la limpieza.")

    historico = cargar_historico()
    fusionado = fusionar(historico, df_limpio)
    guardar_historico(fusionado)

    payload = construir_payload(fusionado)
    html = render_html(payload)

    salida_base = BASE_DIR / f"dashboard_ventas_{xlsx_path.stem}.html"
    salida = siguiente_nombre_libre(salida_base)
    salida.write_text(html, encoding="utf-8")

    print("\n--- Dashboard generado ---")
    print(f"  Archivo:        {salida}")
    print(f"  Registros:      {payload['META']['n_registros']}")
    print(f"  Comprobantes:   {payload['META']['total_ops']}")
    print(f"  Rango de fechas:{payload['META']['rango_min']} a {payload['META']['rango_max']}")
    print(f"  Vendedores:     {len(payload['VENDEDORES'])}")
    print(f"  Productos:      {len(payload['PRODUCTOS'])}")
    print(f"  Clientes:       {len(payload['CLIENTES'])}")
    print(f"\nVerificar con: node verificar_ventas.mjs {salida.name}")


if __name__ == "__main__":
    main()
