#!/usr/bin/env python3
"""
asignar_creditos.py -- Propuesta de limite de credito por cliente segun ventas y plazo.

Uso:
    python3 asignar_creditos.py <FACTURACION.xlsx>
    python3 asignar_creditos.py <dashboard_ventas_XXX.html>
    python3 asignar_creditos.py <historico_ventas.json>

Salida: Excel `creditos_sugeridos_<fuente>.xlsx` + resumen por consola.

La logica esta documentada en creditos/README.md. Resumen:
cada factura ocupa credito desde su fecha hasta su vencimiento (fecha + plazo de
su condicion de pago). El limite tiene que cubrir el pico de esa ocupacion.
"""

import argparse
import json
import re
import sys
import unicodedata
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent

REQUIRED_COLS = [
    "Fechacomprobante", "Transacconsubtiponombre", "Comprobante", "Cliente",
    "Costo", "Condicionpago", "Moneda", "Vendedor", "Producto", "Cantidad",
    "Precio", "Importemonsecundaria", "Principioactivo",
]

NO_NUMERICO_CONTADO = {"CONTADO"}

MESES_ABBR = ["ene.", "feb.", "mar.", "abr.", "may.", "jun.",
              "jul.", "ago.", "sep.", "oct.", "nov.", "dic."]

# Escalones comerciales de redondeo del limite (USD).
ESCALONES = [(5_000, 500), (50_000, 1_000), (200_000, 5_000), (float("inf"), 10_000)]


# ---------------------------------------------------------------------------
# Criterios heredados del dashboard de ventas (congelados, no re-decidir)
# ---------------------------------------------------------------------------

def parsear_dias(condicion):
    """CONTADO -> 0. Rango ("30-60 DIAS", "45-75-105") -> el MAXIMO.
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


def precio_unitario(row):
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


# ---------------------------------------------------------------------------
# Lectura: acepta Excel de facturacion, dashboard de ventas o historico JSON
# ---------------------------------------------------------------------------

def leer_excel(path):
    df = pd.read_excel(path, engine="openpyxl")
    faltantes = [c for c in REQUIRED_COLS if c not in df.columns]
    if faltantes:
        raise SystemExit(f"ERROR: al Excel le faltan columnas requeridas: {', '.join(faltantes)}")
    df = df.dropna(subset=["Fechacomprobante", "Comprobante", "Cliente"])
    out = pd.DataFrame()
    out["fecha"] = pd.to_datetime(df["Fechacomprobante"])
    out["comprobante"] = df["Comprobante"].map(normalizar_comprobante)
    out["cliente"] = df["Cliente"].fillna("").astype(str).str.strip()
    out["vendedor"] = df["Vendedor"].fillna("").astype(str).str.strip()
    out["condicionpago"] = df["Condicionpago"].fillna("").astype(str).str.strip()
    out["importe"] = pd.to_numeric(df["Importemonsecundaria"], errors="coerce").fillna(0.0)
    return out


def extraer_const(html, nombre):
    patron = re.compile(r"const\s+" + nombre + r"\s*=\s*", re.MULTILINE)
    m = patron.search(html)
    if not m:
        raise SystemExit(f"ERROR: no se encontro 'const {nombre}' en el HTML.")
    inicio = m.end()
    fin = html.index(";\n", inicio)
    return json.loads(html[inicio:fin])


def leer_dashboard_html(path):
    html = Path(path).read_text(encoding="utf-8")
    registros = extraer_const(html, "REGISTROS")
    clientes = extraer_const(html, "CLIENTES")
    vendedores = extraer_const(html, "VENDEDORES")
    condiciones = extraer_const(html, "CONDICIONES")
    filas = [{
        "fecha": r["f"],
        "comprobante": r["o"],
        "cliente": clientes[r["c"]],
        "vendedor": vendedores[r["v"]],
        "condicionpago": condiciones[r["cp"]]["n"],
        "importe": r["imp"],
    } for r in registros]
    df = pd.DataFrame(filas)
    df["fecha"] = pd.to_datetime(df["fecha"])
    return df


def leer_historico_json(path):
    filas = json.loads(Path(path).read_text(encoding="utf-8"))
    df = pd.DataFrame([{
        "fecha": f["fecha"],
        "comprobante": f["comprobante"],
        "cliente": f["cliente"],
        "vendedor": f.get("vendedor", ""),
        "condicionpago": f["condicionpago"],
        "importe": f["importe"],
    } for f in filas])
    df["fecha"] = pd.to_datetime(df["fecha"])
    return df


def cargar(path):
    p = Path(path)
    if not p.exists():
        raise SystemExit(f"ERROR: no existe el archivo {p}")
    suf = p.suffix.lower()
    if suf in (".xlsx", ".xlsm", ".xls"):
        return leer_excel(p), "excel"
    if suf in (".html", ".htm"):
        return leer_dashboard_html(p), "dashboard"
    if suf == ".json":
        return leer_historico_json(p), "historico"
    raise SystemExit(f"ERROR: extension no soportada ({suf}). Use .xlsx, .html o .json")


# ---------------------------------------------------------------------------
# Motor de exposicion: cuanto credito ocupa cada cliente, dia por dia
# ---------------------------------------------------------------------------

def serie_exposicion(facturas, dia_ini, dia_fin):
    """Serie diaria de credito ocupado.

    Cada factura ocupa credito desde su fecha (inclusive) hasta su vencimiento
    (fecha + plazo, exclusive). Se arma con un array de diferencias y se acumula:
    O(facturas + dias) en vez de O(facturas x dias).

    Las notas de credito entran con importe negativo y liberan cupo solas.
    Las facturas de CONTADO (plazo 0) no ocupan credito: no entran.
    """
    n = (dia_fin - dia_ini).days + 1
    if n <= 0:
        return np.zeros(0)
    delta = np.zeros(n + 1)
    for fecha, dias, importe in facturas:
        if dias is None or pd.isna(dias) or dias <= 0:
            continue
        i = (fecha - dia_ini).days
        j = i + int(dias)
        if j <= 0 or i >= n:
            continue
        delta[max(i, 0)] += importe
        if j < n:
            delta[j] -= importe
    return np.cumsum(delta[:n])


def escalonar(monto):
    """Redondea el limite al escalon comercial de arriba."""
    if monto <= 0:
        return 0
    for tope, paso in ESCALONES:
        if monto < tope:
            return int(np.ceil(monto / paso) * paso)
    return int(monto)


def factor_antiguedad(meses, campanias):
    """Cliente nuevo = menos limite hasta que haya historia que lo respalde."""
    if meses >= 24 and campanias >= 2:
        return 1.00
    if meses >= 12:
        return 0.90
    if meses >= 6:
        return 0.75
    return 0.60


def rango_campania(etiqueta, mes_inicio):
    """'2025/26' (o '2025') -> (1-jul-2025, 30-jun-2026) con mes_inicio=7."""
    m = re.match(r"^\s*(\d{4})\s*(?:[/-]\s*(\d{2,4}))?\s*$", str(etiqueta))
    if not m:
        raise SystemExit("ERROR: --campania va como 2025/26 o 2025")
    anio = int(m.group(1))
    ini = pd.Timestamp(year=anio, month=mes_inicio, day=1)
    fin = ini + pd.DateOffset(years=1) - pd.Timedelta(days=1)
    return ini, fin


def detectar_campania(df, mes_inicio):
    """Mes de menor facturacion: el corte natural de campania segun los datos."""
    por_mes = df.assign(m=df["fecha"].dt.month).groupby("m")["importe"].sum()
    por_mes = por_mes.reindex(range(1, 13), fill_value=0.0)
    valle = int(por_mes.idxmin())
    sugerido = valle % 12 + 1          # la campania arranca despues del valle
    return por_mes, valle, sugerido


def separar_codigo(nombre):
    """Si el nombre del cliente trae el numero de cuenta adelante, lo separa."""
    m = re.match(r"^\s*(\d{3,})\s*[-–]\s*(.+)$", nombre)
    if m:
        return m.group(1), m.group(2).strip()
    return None, nombre


def norm_texto(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Z0-9]", "", s.upper())


# ---------------------------------------------------------------------------
# Analisis por cliente
# ---------------------------------------------------------------------------

def analizar(df, corte, meses_ventana, base_modo, crecimiento,
             usar_gamma=True, desde=None):
    df = df.copy()
    df["dias"] = df["condicionpago"].map(parsear_dias)

    # La ventana nunca puede empezar antes del primer dato: rellenar con ceros
    # dias sin informacion deprimiria el percentil 95 de toda la cartera.
    ini_ventana = corte - pd.DateOffset(months=meses_ventana)
    primer_dato = df["fecha"].min()
    if desde is not None:
        ini_ventana = max(ini_ventana, pd.Timestamp(desde))
    ini_ventana = max(ini_ventana, primer_dato - pd.Timedelta(days=1))
    ini_12m = max(corte - pd.DateOffset(months=12), primer_dato - pd.Timedelta(days=1))

    filas = []
    for cliente, g in df.groupby("cliente", sort=False):
        g12 = g[(g["fecha"] > ini_12m) & (g["fecha"] <= corte)]
        gv = g[(g["fecha"] > ini_ventana) & (g["fecha"] <= corte)]

        credito12 = g12.loc[g12["dias"].notna() & (g12["dias"] > 0), "importe"].sum()
        contado12 = g12.loc[g12["dias"] == 0, "importe"].sum()
        sinplazo12 = g12.loc[g12["dias"].isna(), "importe"].sum()
        total12 = g12["importe"].sum()

        # Plazo ponderado por USD, solo sobre ventas a credito y con importe positivo
        # (las NC negativas distorsionarian el promedio; se netean en los volumenes).
        cred = g[g["dias"].notna() & (g["dias"] > 0) & (g["importe"] > 0)]
        cred_v = cred[(cred["fecha"] > ini_ventana) & (cred["fecha"] <= corte)]
        base_plazo = cred_v if cred_v["importe"].sum() > 0 else cred
        if base_plazo["importe"].sum() > 0:
            plazo_pond = float(np.average(base_plazo["dias"], weights=base_plazo["importe"]))
        else:
            plazo_pond = 0.0

        facturas = list(zip(gv["fecha"].dt.to_pydatetime(), gv["dias"], gv["importe"]))
        dia_ini = ini_ventana.to_pydatetime()
        serie = serie_exposicion(facturas, dia_ini, corte.to_pydatetime())
        if serie.size:
            exp_max = float(serie.max())
            exp_p95 = float(np.percentile(serie, 95))
            exp_hoy = float(serie[-1])
            dias_con_saldo = int((serie > 0).sum())
        else:
            exp_max = exp_p95 = exp_hoy = 0.0
            dias_con_saldo = 0

        # Formula de rotacion: el saldo medio que sostiene ese volumen anual.
        nec_media = credito12 * plazo_pond / 365.0 if plazo_pond > 0 else 0.0

        if base_modo == "max":
            base = max(exp_max, nec_media)
        elif base_modo == "media":
            base = nec_media
        else:
            base = max(exp_p95, nec_media)

        primera = g["fecha"].min()
        meses_rel = (corte.year - primera.year) * 12 + (corte.month - primera.month)
        campanias = g["fecha"].dt.year.nunique()
        gamma = factor_antiguedad(meses_rel, campanias) if usar_gamma else 1.00

        limite = base * gamma * crecimiento
        codigo, nombre = separar_codigo(cliente)

        alertas = []
        if sinplazo12 > 0:
            alertas.append("ventas sin plazo definido (canje/granos/plataforma): decidir aparte")
        if plazo_pond == 0 and credito12 == 0 and total12 != 0:
            alertas.append("opera solo contado: no requiere linea")
        if usar_gamma and meses_rel < 12:
            alertas.append(f"cliente nuevo ({meses_rel} meses de relacion)")
        elif not usar_gamma and primera > ini_ventana + pd.Timedelta(days=45):
            alertas.append(f"primera compra del periodo recien en {primera:%m/%Y}: "
                           "sin historia previa en el archivo")
        if exp_max > 0 and nec_media > 0 and exp_max / nec_media > 3:
            alertas.append("muy estacional: el pico triplica al saldo medio")
        if total12 < 0:
            alertas.append("neto negativo en 12m (NC > facturas): revisar")

        filas.append({
            "nro": codigo,
            "cliente": nombre,
            "cliente_raw": cliente,
            "vendedor": g["vendedor"].mode().iloc[0] if not g["vendedor"].mode().empty else "",
            "ventas_credito_12m": credito12,
            "ventas_contado_12m": contado12,
            "ventas_sin_plazo_12m": sinplazo12,
            "ventas_total_12m": total12,
            "plazo_pond": plazo_pond,
            "ciclos_anio": 365.0 / plazo_pond if plazo_pond > 0 else np.nan,
            "exp_pico": exp_max,
            "exp_p95": exp_p95,
            "exp_al_corte": exp_hoy,
            "nec_media": nec_media,
            "base": base,
            "meses_relacion": meses_rel,
            "campanias": campanias,
            "gamma": gamma,
            "limite_calc": limite,
            "dias_con_saldo": dias_con_saldo,
            "alertas": "; ".join(alertas),
        })

    return pd.DataFrame(filas)


def aplicar_topes(res, cap_pct, capacidad):
    """Concentracion por cliente y, si se informa, techo global de financiacion."""
    res = res.copy()
    res["limite_sugerido"] = res["limite_calc"].map(escalonar)
    res["tope_aplicado"] = ""

    avisos = []
    if cap_pct and cap_pct > 0:
        # Un tope del x% es matematicamente imposible si x% * n_clientes < 1:
        # aplicarlo igual aplastaria a toda la cartera al mismo numero.
        if cap_pct * len(res) < 1:
            avisos.append(
                f"tope de concentracion {cap_pct:.0%} ignorado: con {len(res)} clientes "
                f"el minimo viable es {1/len(res):.0%}")
        else:
            total = res["limite_sugerido"].sum()
            tope = total * cap_pct
            excede = res["limite_sugerido"] > tope
            res.loc[excede, "tope_aplicado"] = f"concentracion {cap_pct:.0%}"
            res.loc[excede, "limite_sugerido"] = escalonar(tope)

    if capacidad and capacidad > 0:
        total = res["limite_sugerido"].sum()
        if total > capacidad:
            factor = capacidad / total
            res["limite_sugerido"] = (res["limite_sugerido"] * factor).map(escalonar)
            res["tope_aplicado"] = res["tope_aplicado"].str.cat(
                pd.Series([f"prorrateo x{factor:.2f}"] * len(res), index=res.index),
                sep="; ").str.strip("; ")

    total = res["limite_sugerido"].sum()
    res["pct_cartera"] = res["limite_sugerido"] / total if total else 0.0
    res["ventas_soportadas"] = np.where(
        res["plazo_pond"] > 0,
        res["limite_sugerido"] * 365.0 / res["plazo_pond"].replace(0, np.nan),
        np.nan)
    return res, avisos


# ---------------------------------------------------------------------------
# Salida
# ---------------------------------------------------------------------------

COLS_SALIDA = [
    ("nro", "Nro cliente"),
    ("cliente", "Cliente"),
    ("vendedor", "Vendedor"),
    ("ventas_total_12m", "Ventas 12m USD"),
    ("ventas_credito_12m", "  a credito"),
    ("ventas_contado_12m", "  contado"),
    ("ventas_sin_plazo_12m", "  sin plazo"),
    ("plazo_pond", "Plazo pond. (dias)"),
    ("ciclos_anio", "Ciclos/anio"),
    ("exp_pico", "Exposicion pico USD"),
    ("exp_p95", "Exposicion P95 USD"),
    ("exp_al_corte", "Saldo al corte USD"),
    ("nec_media", "Saldo medio (rotacion)"),
    ("base", "Base de calculo"),
    ("meses_relacion", "Antiguedad (meses)"),
    ("gamma", "Factor antiguedad"),
    ("limite_sugerido", "LIMITE SUGERIDO USD"),
    ("ventas_soportadas", "Ventas anuales que soporta"),
    ("pct_cartera", "% de la cartera"),
    ("tope_aplicado", "Tope aplicado"),
    ("alertas", "Alertas"),
]


def exportar(res, destino, meta):
    salida = res[[c for c, _ in COLS_SALIDA]].rename(columns=dict(COLS_SALIDA))
    with pd.ExcelWriter(destino, engine="openpyxl") as xl:
        salida.to_excel(xl, sheet_name="Limites sugeridos", index=False)
        pd.DataFrame(meta.items(), columns=["Parametro", "Valor"]).to_excel(
            xl, sheet_name="Criterios", index=False)
        ws = xl.sheets["Limites sugeridos"]
        ws.freeze_panes = "C2"
        for col in ws.columns:
            largo = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max(largo + 2, 10), 42)
    return destino


def siguiente_nombre_libre(base):
    if not base.exists():
        return base
    i = 2
    while True:
        cand = base.with_name(f"{base.stem}_v{i}{base.suffix}")
        if not cand.exists():
            return cand
        i += 1


def main():
    ap = argparse.ArgumentParser(description="Propuesta de limite de credito por cliente.")
    ap.add_argument("fuente", help="Excel de facturacion, dashboard de ventas HTML o historico JSON")
    ap.add_argument("--corte", help="Fecha de corte AAAA-MM-DD (default: ultima factura)")
    ap.add_argument("--ventana", type=int, default=24,
                    help="Meses de historia para medir la exposicion (default 24)")
    ap.add_argument("--campania", help="Analizar una campania: 2025/26 (o 2025)")
    ap.add_argument("--inicio-campania", type=int, default=7, metavar="MM",
                    help="Mes en que arranca la campania (default 7 = julio)")
    ap.add_argument("--antiguedad", choices=["auto", "on", "off"], default="auto",
                    help="Factor de antiguedad. auto = solo si hay >=18 meses de datos")
    ap.add_argument("--base", choices=["p95", "max", "media"], default="p95",
                    help="Que medida de exposicion usar como base (default p95)")
    ap.add_argument("--crecimiento", type=float, default=1.10,
                    help="Holgura para crecimiento, ej 1.10 = +10%% (default 1.10)")
    ap.add_argument("--cap-concentracion", type=float, default=0.10,
                    help="Tope por cliente como fraccion de la cartera (default 0.10; 0 = sin tope)")
    ap.add_argument("--capacidad", type=float, default=0,
                    help="Techo global de financiacion en USD (0 = sin techo)")
    ap.add_argument("--min-ventas", type=float, default=0,
                    help="Ignorar clientes con ventas 12m por debajo de este monto")
    ap.add_argument("--salida", help="Ruta del Excel de salida")
    args = ap.parse_args()

    df, origen = cargar(args.fuente)
    if df.empty:
        raise SystemExit("ERROR: la fuente no tiene filas utilizables.")

    por_mes, valle, mes_sugerido = detectar_campania(df, args.inicio_campania)

    desde = None
    periodo = "historia completa"
    if args.campania:
        desde, fin = rango_campania(args.campania, args.inicio_campania)
        df = df[(df["fecha"] >= desde) & (df["fecha"] <= fin)]
        if df.empty:
            raise SystemExit(f"ERROR: no hay ventas entre {desde:%d/%m/%Y} y {fin:%d/%m/%Y}")
        periodo = f"campania {args.campania} ({desde:%m/%Y} a {fin:%m/%Y})"

    corte = pd.Timestamp(args.corte) if args.corte else df["fecha"].max()
    primer_dato = df["fecha"].min()
    span_meses = (corte.year - primer_dato.year) * 12 + (corte.month - primer_dato.month)

    if args.antiguedad == "on":
        usar_gamma = True
    elif args.antiguedad == "off":
        usar_gamma = False
    else:
        usar_gamma = span_meses >= 18

    desconocidas = sorted({c for c in df["condicionpago"].unique() if parsear_dias(c) is None})
    plazo_max = max([d for d in df["condicionpago"].map(parsear_dias) if d] or [0])

    res = analizar(df, corte, args.ventana, args.base, args.crecimiento,
                   usar_gamma=usar_gamma, desde=desde)
    if args.min_ventas:
        res = res[res["ventas_total_12m"] >= args.min_ventas]
    res = res[res["limite_calc"] > 0].copy()
    res, avisos = aplicar_topes(res, args.cap_concentracion, args.capacidad)
    res = res.sort_values("limite_sugerido", ascending=False).reset_index(drop=True)

    meta = {
        "Fuente": Path(args.fuente).name,
        "Tipo de fuente": origen,
        "Periodo analizado": periodo,
        "Fecha de corte": corte.strftime("%Y-%m-%d"),
        "Historia disponible (meses)": span_meses,
        "Factor de antiguedad": "aplicado" if usar_gamma else "neutralizado (historia corta)",
        "Ventana de exposicion (meses)": args.ventana,
        "Base de calculo": args.base,
        "Holgura de crecimiento": args.crecimiento,
        "Tope de concentracion": args.cap_concentracion,
        "Capacidad global USD": args.capacidad or "sin techo",
        "Generado": datetime.now().strftime("%d/%m/%Y %H:%M"),
    }

    destino = Path(args.salida) if args.salida else siguiente_nombre_libre(
        BASE_DIR / f"creditos_sugeridos_{Path(args.fuente).stem}.xlsx")
    exportar(res, destino, meta)

    # ---- Resumen por consola ----
    print(f"\nFuente: {Path(args.fuente).name}  ({origen})")
    print(f"Periodo: {periodo}")
    print(f"Datos: {primer_dato:%d/%m/%Y} a {corte:%d/%m/%Y}  ({span_meses} meses)   Base: {args.base}")
    print(f"Clientes con linea propuesta: {len(res)}")
    print(f"Suma de limites: US$ {res['limite_sugerido'].sum():,.0f}")
    print(f"Ventas 12m (total): US$ {res['ventas_total_12m'].sum():,.0f}")
    if not usar_gamma:
        avisos.append(
            f"con {span_meses} meses de datos NO se puede medir antiguedad: el factor "
            "quedo neutro (1,00) para todos. Con mas anios, o con las fechas de alta, "
            "se castiga al cliente nuevo")
    if plazo_max and span_meses <= 14:
        avisos.append(
            f"los primeros ~{plazo_max} dias del periodo subestiman la exposicion: las "
            "facturas abiertas que vienen de la campania anterior no estan en el archivo")
    for a in avisos:
        print(f"AVISO: {a}")

    print("\nFacturacion por mes calendario (para ubicar el corte de campania):")
    tot = por_mes.sum()
    for m in range(1, 13):
        v = por_mes[m]
        barra = "#" * int(round(40 * v / por_mes.max())) if por_mes.max() else ""
        marca = "  <- mes mas flojo" if m == valle else ""
        print(f"  {MESES_ABBR[m-1]:<5} {v:>12,.0f}  {barra}{marca}")
    print(f"  Corte de campania sugerido por los datos: arranca en "
          f"{MESES_ABBR[mes_sugerido-1]} (usando --inicio-campania {mes_sugerido})")
    if desconocidas:
        print(f"\nCondiciones de pago SIN plazo numerico ({len(desconocidas)}): "
              f"{', '.join(desconocidas[:12])}")
        print("  -> esas ventas no generan linea automatica; se marcan como alerta.")

    print("\nTop 20 por limite sugerido:")
    cab = f"{'Nro':>8}  {'Cliente':<34} {'Vts 12m':>12} {'Plazo':>6} {'Pico':>12} {'LIMITE':>12}"
    print(cab)
    print("-" * len(cab))
    for _, r in res.head(20).iterrows():
        print(f"{(r['nro'] or '-'): >8}  {r['cliente'][:34]:<34} "
              f"{r['ventas_total_12m']:>12,.0f} {r['plazo_pond']:>6.0f} "
              f"{r['exp_pico']:>12,.0f} {r['limite_sugerido']:>12,.0f}")

    print(f"\nExcel generado: {destino}")
    print("Recordatorio: esta propuesta es SOLO comportamiento comercial. Antes de")
    print("aprobar hay que contrastarla con balances, indices, antiguedad en el rubro,")
    print("deuda tomada con terceros y garantias.")


if __name__ == "__main__":
    main()
