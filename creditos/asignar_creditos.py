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


def parsear_cuotas(condicion):
    """Vencimientos de la condicion, en dias. "30 - 60 - 90 DIAS" -> [30, 60, 90].

    Diferencia deliberada con el dashboard de ventas, que toma el MAXIMO: para
    exposicion de credito el maximo sobreestima, porque a los 60 dias ya cobraste
    un tercio. Cada cuota ocupa cupo por su propio plazo y se asume partes iguales,
    que es la convencion de plaza.
    """
    if condicion is None:
        return None
    texto = str(condicion).strip().upper()
    if not texto:
        return None
    numeros = [int(n) for n in re.findall(r"\d+", texto)]
    if numeros:
        return sorted(set(numeros))
    if texto in NO_NUMERICO_CONTADO:
        return [0]
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

def etiqueta_campania(fechas, mes_inicio):
    """Campania a la que pertenece cada fecha. Con mes_inicio=4, el 15/02/2026
    cae en la campania 2025 (abr-2025 a mar-2026)."""
    return fechas.dt.year.where(fechas.dt.month >= mes_inicio, fechas.dt.year - 1)


def analizar(df, desde, corte, base_modo, crecimiento,
             usar_gamma=True, mes_campania=7, plazo_otras=None):
    """Mide sobre [desde, corte], pero arrastra las facturas anteriores.

    La distincion importa: una factura de la campania previa que sigue abierta el
    primer dia del periodo OCUPA credito y tiene que contar. Por eso la serie de
    exposicion se arma con TODO el historial y recien despues se recorta al periodo
    que se informa. Sin eso, el arranque del periodo muestra menos saldo del real.
    """
    df = df.copy()
    # cuotas_orig conserva la lectura literal de la condicion, para poder seguir
    # informando cuanto se vendio por canje / plataforma / tarjeta. cuotas es lo
    # que usa el motor: si se fijo un plazo para esas condiciones, va ahi.
    df["cuotas_orig"] = df["condicionpago"].map(parsear_cuotas)
    df["cuotas"] = df["cuotas_orig"]
    if plazo_otras is not None:
        df["cuotas"] = [c if isinstance(c, list) else [plazo_otras]
                        for c in df["cuotas_orig"]]
    # Plazo representativo de la condicion = promedio de sus cuotas (una sola cuota
    # devuelve ese mismo plazo). Es el que corresponde para medir credito.
    df["dias"] = df["cuotas"].map(
        lambda c: float(np.mean(c)) if isinstance(c, list) and c else None)
    df = df[df["fecha"] <= corte]

    primer_dato = df["fecha"].min()
    desde = max(pd.Timestamp(desde), primer_dato - pd.Timedelta(days=1))
    # Dias de arrastre disponibles antes del periodo informado.
    warmup = max((desde - primer_dato).days, 0)

    # Campania anterior: mismo largo, inmediatamente antes del periodo.
    ini_previo = desde - (corte - desde) - pd.Timedelta(days=1)

    filas = []
    series = {}
    for cliente, g in df.groupby("cliente", sort=False):
        gp = g[(g["fecha"] >= desde) & (g["fecha"] <= corte)]
        gprev = g[(g["fecha"] >= ini_previo) & (g["fecha"] < desde)]
        ventas_prev = gprev["importe"].sum()

        credito_p = gp.loc[gp["dias"].notna() & (gp["dias"] > 0), "importe"].sum()
        contado_p = gp.loc[gp["dias"] == 0, "importe"].sum()
        sinplazo_p = gp.loc[gp["cuotas_orig"].isna(), "importe"].sum()
        total_p = gp["importe"].sum()

        # Plazo ponderado por USD, solo sobre ventas a credito con importe positivo
        # (las NC negativas distorsionarian el promedio; se netean en los volumenes).
        cred = g[g["dias"].notna() & (g["dias"] > 0) & (g["importe"] > 0)]
        cred_p = cred[(cred["fecha"] >= desde) & (cred["fecha"] <= corte)]
        base_plazo = cred_p if cred_p["importe"].sum() > 0 else cred
        if base_plazo["importe"].sum() > 0:
            plazo_pond = float(np.average(base_plazo["dias"], weights=base_plazo["importe"]))
        else:
            plazo_pond = 0.0

        # Serie con TODAS las facturas del cliente; el recorte viene despues.
        # Cada factura se abre en sus cuotas: cada una ocupa cupo por su plazo.
        facturas = []
        for fecha, cuotas, importe in zip(g["fecha"].dt.to_pydatetime(),
                                          g["cuotas"], g["importe"]):
            if not isinstance(cuotas, list) or not cuotas:
                continue
            parte = importe / len(cuotas)
            for d in cuotas:
                facturas.append((fecha, d, parte))
        # Dos series distintas, porque responden dos preguntas distintas:
        #  - TOTAL   (arrastre + periodo): que riesgo se corrio de verdad.
        #  - PROPIA  (solo lo facturado en el periodo): que exposicion genera la
        #    actividad actual del cliente. Es la que fija el limite, porque el
        #    limite mira para adelante. Un cliente que compro fuerte la campania
        #    pasada y este ano casi no compro arrastra un pico alto que ya se esta
        #    liquidando: darle linea por ese pico es financiar una retirada.
        ini_py = primer_dato.to_pydatetime()
        corte_py = corte.to_pydatetime()
        i_desde = max((desde - primer_dato).days, 0)

        serie_total = serie_exposicion(facturas, ini_py, corte_py)[i_desde:]
        propias = [f for f in facturas if f[0] >= desde.to_pydatetime()]
        serie_propia = serie_exposicion(propias, ini_py, corte_py)[i_desde:]

        if serie_total.size:
            exp_max_total = float(serie_total.max())
            exp_hoy = float(serie_total[-1])
            exp_apertura = float(serie_total[0])
            dias_con_saldo = int((serie_total > 0).sum())
        else:
            exp_max_total = exp_hoy = exp_apertura = 0.0
            dias_con_saldo = 0

        # El percentil se toma sobre los dias en que el cliente DEBE algo, no sobre
        # el calendario. Un cliente que compra dos veces al ano a 15 dias tiene
        # saldo 15 dias de 365: sobre el calendario el P95 cae en un dia de saldo
        # cero y el limite se va a cero, cuando en realidad necesita cubrir la
        # operacion que hace. Sobre los dias con saldo la pregunta es la correcta:
        # cuando nos debe, cuanto nos debe.
        con_saldo = serie_propia[serie_propia > 0]
        if serie_propia.size:
            exp_max = float(serie_propia.max())
            exp_p95 = float(np.percentile(con_saldo, 95)) if con_saldo.size else 0.0
        else:
            exp_max = exp_p95 = 0.0

        # Formula de rotacion: el saldo medio que sostiene ese volumen anual.
        dias_periodo = max((corte - desde).days, 1)
        credito_anualizado = credito_p * 365.0 / dias_periodo
        nec_media = credito_anualizado * plazo_pond / 365.0 if plazo_pond > 0 else 0.0

        if base_modo == "max":
            base = max(exp_max, nec_media)
        elif base_modo == "media":
            base = nec_media
        else:
            base = max(exp_p95, nec_media)

        primera = g["fecha"].min()
        meses_rel = (corte.year - primera.year) * 12 + (corte.month - primera.month)
        campanias = int(etiqueta_campania(g["fecha"], mes_campania).nunique())
        gamma = factor_antiguedad(meses_rel, campanias) if usar_gamma else 1.00

        limite = base * gamma * crecimiento
        codigo, nombre = separar_codigo(cliente)
        # Para medir friccion futura se usa la serie de la actividad propia: el
        # limite se va a aplicar a la campania que viene, cuando el arrastre de la
        # anterior ya se liquido. El arrastre se informa aparte (saldo al corte).
        series[cliente] = serie_propia

        alertas = []
        if sinplazo_p > 0:
            alertas.append(
                f"US$ {sinplazo_p:,.0f} por canje/plataforma/tarjeta"
                + (f", computados a {plazo_otras} dias" if plazo_otras is not None
                   else ": sin plazo, decidir aparte"))
        if plazo_pond == 0 and credito_p == 0 and total_p != 0:
            alertas.append("opera solo contado: no requiere linea")
        if usar_gamma and meses_rel < 12:
            alertas.append(f"cliente nuevo ({meses_rel} meses de relacion)")
        elif not usar_gamma and primera > desde + pd.Timedelta(days=45):
            alertas.append(f"primera compra del periodo recien en {primera:%m/%Y}: "
                           "sin historia previa en el archivo")
        if usar_gamma and campanias < 2:
            alertas.append("compro en una sola campania")
        if ventas_prev > 0 and total_p > 0 and total_p / ventas_prev < 0.5:
            alertas.append(f"cayo {1 - total_p / ventas_prev:.0%} vs la campania anterior")
        elif ventas_prev > 0 and total_p / ventas_prev > 2:
            alertas.append(f"crecio {total_p / ventas_prev - 1:.0%} vs la campania anterior")
        if exp_max > 0 and nec_media > 0 and exp_max / nec_media > 3:
            alertas.append("muy estacional: el pico triplica al saldo medio")
        if total_p < 0:
            alertas.append("neto negativo en el periodo (NC > facturas): revisar")
        if warmup > 0 and exp_apertura > 0:
            alertas.append(f"abrio el periodo con US$ {exp_apertura:,.0f} de la campania anterior")
        if exp_max > 0 and exp_max_total > exp_max * 1.5:
            alertas.append(f"el riesgo real llego a US$ {exp_max_total:,.0f} por arrastre, "
                           "muy por encima de lo que genera la campania actual: viene bajando")
        elif exp_max_total > 0 and exp_max > exp_max_total * 1.5:
            alertas.append("la campania actual genera mas exposicion que el pico historico: "
                           "viene creciendo fuerte")

        filas.append({
            "nro": codigo,
            "cliente": nombre,
            "cliente_raw": cliente,
            "vendedor": g["vendedor"].mode().iloc[0] if not g["vendedor"].mode().empty else "",
            "ventas_credito_p": credito_p,
            "ventas_contado_p": contado_p,
            "ventas_sin_plazo_p": sinplazo_p,
            "ventas_total_p": total_p,
            "ventas_previa": ventas_prev,
            "var_campania": (total_p / ventas_prev - 1) if ventas_prev > 0 else np.nan,
            "plazo_pond": plazo_pond,
            "ciclos_anio": 365.0 / plazo_pond if plazo_pond > 0 else np.nan,
            "exp_pico": exp_max,
            "exp_p95": exp_p95,
            "exp_pico_total": exp_max_total,
            "exp_apertura": exp_apertura,
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

    return pd.DataFrame(filas), warmup, series


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
    ("ventas_total_p", "Ventas del periodo USD"),
    ("ventas_credito_p", "  a credito"),
    ("ventas_contado_p", "  contado"),
    ("ventas_sin_plazo_p", "  sin plazo"),
    ("ventas_previa", "Ventas campania anterior"),
    ("var_campania", "Var. vs campania anterior"),
    ("plazo_pond", "Plazo pond. (dias)"),
    ("ciclos_anio", "Ciclos/anio"),
    ("exp_pico", "Pico campania actual USD"),
    ("exp_pico_total", "Pico con arrastre USD"),
    ("exp_p95", "Exposicion P95 USD"),
    ("exp_apertura", "Saldo al abrir el periodo"),
    ("exp_al_corte", "Saldo al corte USD"),
    ("nec_media", "Saldo medio (rotacion)"),
    ("base", "Base de calculo"),
    ("meses_relacion", "Antiguedad (meses)"),
    ("campanias", "Campanias"),
    ("gamma", "Factor antiguedad"),
    ("limite_sugerido", "LIMITE SUGERIDO USD"),
    ("ventas_soportadas", "Ventas anuales que soporta"),
    ("dias_sobre_limite", "Dias/ano sobre el limite"),
    ("exceso_max", "Exceso maximo USD"),
    ("exceso_max_pct", "Exceso maximo %"),
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
    ap.add_argument("--campania", help="Analizar una campania: 2025/26 (o 2025)")
    ap.add_argument("--inicio-campania", type=int, default=7, metavar="MM",
                    help="Mes en que arranca la campania (default 7 = julio)")
    ap.add_argument("--plazo-otras", type=int, metavar="N",
                    help="Plazo en dias para condiciones sin numero (canje, plataforma, "
                         "tarjeta). Sin esto quedan fuera del calculo")
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
                    help="Ignorar clientes con ventas del periodo por debajo de este monto")
    ap.add_argument("--salida", help="Ruta del Excel de salida")
    args = ap.parse_args()

    df, origen = cargar(args.fuente)
    if df.empty:
        raise SystemExit("ERROR: la fuente no tiene filas utilizables.")

    por_mes, valle, mes_sugerido = detectar_campania(df, args.inicio_campania)
    primer_dato_total = df["fecha"].min()
    ultimo_dato = df["fecha"].max()

    if args.campania:
        desde, fin = rango_campania(args.campania, args.inicio_campania)
        if not ((df["fecha"] >= desde) & (df["fecha"] <= fin)).any():
            raise SystemExit(f"ERROR: no hay ventas entre {desde:%d/%m/%Y} y {fin:%d/%m/%Y}.\n"
                             f"  El archivo va de {primer_dato_total:%d/%m/%Y} a "
                             f"{ultimo_dato:%d/%m/%Y}.")
        corte = min(pd.Timestamp(args.corte) if args.corte else fin, fin)
        periodo = f"campania {args.campania} ({desde:%m/%Y} a {fin:%m/%Y})"
    else:
        desde = primer_dato_total
        corte = pd.Timestamp(args.corte) if args.corte else ultimo_dato
        periodo = "historia completa"

    # La antiguedad se mide contra TODO el archivo, no contra el periodo informado.
    span_meses = ((corte.year - primer_dato_total.year) * 12
                  + (corte.month - primer_dato_total.month))

    if args.antiguedad == "on":
        usar_gamma = True
    elif args.antiguedad == "off":
        usar_gamma = False
    else:
        usar_gamma = span_meses >= 18

    desconocidas = sorted({c for c in df["condicionpago"].unique() if parsear_dias(c) is None})
    plazo_max = max([d for d in df["condicionpago"].map(parsear_dias) if d] or [0])

    res, warmup, series = analizar(df, desde, corte, args.base, args.crecimiento,
                                   usar_gamma=usar_gamma, mes_campania=args.inicio_campania,
                                   plazo_otras=args.plazo_otras)
    if args.min_ventas:
        res = res[res["ventas_total_p"] >= args.min_ventas]
    res = res[res["limite_calc"] > 0].copy()
    res, avisos = aplicar_topes(res, args.cap_concentracion, args.capacidad)
    res = res.sort_values("limite_sugerido", ascending=False).reset_index(drop=True)

    # Cuantos dias del periodo el saldo real habria superado el limite propuesto.
    # Es la medida operativa: cada uno de esos dias es un pedido de excepcion.
    # Ademas de cuantos dias, CUANTO se pasa. Sin la magnitud, pasarse US$ 16 en
    # una linea de 3.000 se lee igual que pasarse US$ 80.000: el redondeo del
    # limite solo ya alcanza para generar cientos de dias de exceso trivial.
    res["dias_sobre_limite"] = [
        int((series[c] > lim).sum()) if c in series else 0
        for c, lim in zip(res["cliente_raw"], res["limite_sugerido"])]
    res["exceso_max"] = [
        float(max((series[c] - lim).max(), 0.0)) if c in series and series[c].size else 0.0
        for c, lim in zip(res["cliente_raw"], res["limite_sugerido"])]
    res["exceso_max_pct"] = np.where(res["limite_sugerido"] > 0,
                                     res["exceso_max"] / res["limite_sugerido"], 0.0)

    # Arranca la campania que viene ya pasado de linea por deuda de la anterior.
    excedido = res["exp_al_corte"] > res["limite_sugerido"]
    res.loc[excedido, "alertas"] = (
        res.loc[excedido, "alertas"].str.cat(
            "ARRANCA EXCEDIDO: al 31/03 debia US$ "
            + res.loc[excedido, "exp_al_corte"].map(lambda v: f"{v:,.0f}")
            + ", por encima del limite propuesto", sep="; ").str.strip("; "))

    # El archivo puede no traer codigo de cliente. Se numera alfabeticamente para
    # que el numero sea estable entre corridas y entre campanias.
    sin_codigo = res["nro"].isna().all()
    if sin_codigo:
        orden = {n: i + 1 for i, n in enumerate(sorted(res["cliente_raw"]))}
        res["nro"] = res["cliente_raw"].map(orden).map(lambda i: f"P{i:03d}")

    meta = {
        "Fuente": Path(args.fuente).name,
        "Tipo de fuente": origen,
        "Periodo analizado": periodo,
        "Fecha de corte": corte.strftime("%Y-%m-%d"),
        "Historia disponible (meses)": span_meses,
        "Arrastre previo al periodo (dias)": warmup,
        "Factor de antiguedad": "aplicado" if usar_gamma else "neutralizado (historia corta)",
        "Plazo mas largo de la cartera (dias)": plazo_max,
        "Base de calculo": args.base,
        "Holgura de crecimiento": args.crecimiento,
        "Plazo para canje/plataforma/tarjeta": (f"{args.plazo_otras} dias"
                                                if args.plazo_otras is not None
                                                else "excluidas"),
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
    print(f"Datos: {primer_dato_total:%d/%m/%Y} a {ultimo_dato:%d/%m/%Y}  ({span_meses} meses)")
    print(f"Mide {desde:%d/%m/%Y} a {corte:%d/%m/%Y}, arrastrando {warmup} dias previos.   "
          f"Base: {args.base}")
    print(f"Clientes con linea propuesta: {len(res)}")
    print(f"Suma de limites: US$ {res['limite_sugerido'].sum():,.0f}")
    print(f"Ventas del periodo (total): US$ {res['ventas_total_p'].sum():,.0f}")
    if not usar_gamma:
        avisos.append(
            f"con {span_meses} meses de datos NO se puede medir antiguedad: el factor "
            "quedo neutro (1,00) para todos. Con mas anios, o con las fechas de alta, "
            "se castiga al cliente nuevo")
    if plazo_max and warmup < plazo_max:
        avisos.append(
            f"el arrastre disponible ({warmup} dias) es menor al plazo mas largo de la "
            f"cartera ({plazo_max} dias): los primeros ~{plazo_max - warmup} dias del "
            "periodo subestiman la exposicion")
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
        if args.plazo_otras is not None:
            print(f"  -> computadas a {args.plazo_otras} dias.")
        else:
            print("  -> esas ventas no generan linea automatica; se marcan como alerta.")

    print("\nTop 20 por limite sugerido:")
    cab = f"{'Nro':>8}  {'Cliente':<34} {'Vts periodo':>13} {'Plazo':>6} {'Pico':>12} {'LIMITE':>12}"
    print(cab)
    print("-" * len(cab))
    for _, r in res.head(20).iterrows():
        print(f"{(r['nro'] or '-'): >8}  {r['cliente'][:34]:<34} "
              f"{r['ventas_total_p']:>13,.0f} {r['plazo_pond']:>6.0f} "
              f"{r['exp_pico']:>12,.0f} {r['limite_sugerido']:>12,.0f}")

    if sin_codigo:
        print("\nAVISO: el archivo no trae codigo de cliente. Se numero alfabeticamente")
        print("  (P001, P002, ...). Para el numero de cuenta real hay que exportar el")
        print("  reporte con esa columna.")
    print(f"\nExcel generado: {destino}")
    print("Recordatorio: esta propuesta es SOLO comportamiento comercial. Antes de")
    print("aprobar hay que contrastarla con balances, indices, antiguedad en el rubro,")
    print("deuda tomada con terceros y garantias.")


if __name__ == "__main__":
    main()
