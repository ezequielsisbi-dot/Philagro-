#!/usr/bin/env python3
"""
resumen_lc.py -- Vista ejecutiva del limite de credito: Cliente / LC / Comentarios.

Uso:
    python3 resumen_lc.py <creditos_sugeridos_XXX.xlsx> [--salida-html X.html]

Toma la salida detallada de asignar_creditos.py y arma la version corta para
comite: tres columnas, mas un HTML autocontenido (se abre con doble clic, sin
internet) con buscador y orden por columna.
"""

import argparse
import html as html_mod
import json
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent


def fmt(v):
    return f"{v:,.0f}".replace(",", ".")


def normalizar_miles(t):
    """Los avisos vienen con separador de miles ingles; se pasan a punto."""
    return re.sub(r"\d{1,3}(?:,\d{3})+", lambda m: m.group(0).replace(",", "."), t)


def construir_comentario(r):
    """Devuelve [(tipo, frase)]. El tipo lo decide quien arma la frase, no un
    regex sobre el texto final: los importes llevan punto de miles y cualquier
    intento de cortar 'hasta el primer punto' parte el numero al medio."""
    partes = []

    dias = int(r.get("Dias/ano sobre el limite") or 0)
    if dias == 0:
        partes.append(("ok", "No quedaria excedido ningun dia de la campania."))
    elif dias <= 15:
        partes.append(("", f"Quedaria excedido {dias} dias de la campania (tolerable)."))
    elif dias <= 45:
        partes.append(("aviso", f"Quedaria excedido {dias} dias de la campania: "
                                "va a pedir excepcion varias veces."))
    else:
        partes.append(("aviso", f"Quedaria excedido {dias} dias de la campania: "
                                "el limite le queda corto para como compra."))

    plazo = r.get("Plazo pond. (dias)")
    if pd.notna(plazo) and plazo > 0:
        txt = (f"Opera a {plazo:.0f} dias, el cupo le rota {365.0 / plazo:.1f} "
               "veces al ano")
        soporta = r.get("Ventas anuales que soporta")
        if pd.notna(soporta):
            txt += f"; con este LC puede comprar hasta US$ {fmt(soporta)} al ano"
        partes.append(("", txt + "."))

    alertas = str(r.get("Alertas") or "")
    if "ARRANCA EXCEDIDO" in alertas:
        partes.append(("alerta",
                       f"ATENCION: al cierre de campania debia US$ "
                       f"{fmt(r.get('Saldo al corte USD'))}, arranca por encima del LC. "
                       "Antes de habilitar, plan de cobranza."))
    if "viene bajando" in alertas:
        partes.append(("", f"Llego a deber US$ {fmt(r.get('Pico con arrastre USD'))} "
                           "arrastrando saldo viejo; el LC sigue a lo que compra hoy, "
                           "no a ese pico."))
    if "viene creciendo fuerte" in alertas:
        partes.append(("aviso", "Creciendo fuerte: revisar el LC a mitad de campania."))
    for a in alertas.split("; "):
        if a.startswith(("cayo ", "crecio ")) or "canje/plataforma/tarjeta" in a:
            a = normalizar_miles(a)
            partes.append(("", a[0].upper() + a[1:] + "."))
    if "cliente nuevo" in alertas or "una sola campania" in alertas:
        camp = r.get("Campanias")
        partes.append(("aviso", f"Poca historia "
                                f"({int(camp) if pd.notna(camp) else '?'} campania/s): "
                                "LC de arranque, revisar con documentacion."))
    if "solo contado" in alertas:
        partes.append(("", "Opera solo contado: no necesita linea."))
    if "muy estacional" in alertas:
        partes.append(("", "Muy estacional: concentra la compra en pocas semanas."))
    tope = r.get("Tope aplicado")
    if pd.notna(tope) and str(tope).strip() not in ("", "nan"):
        partes.append(("aviso", f"Limitado por {tope}."))

    return partes


def texto_plano(partes):
    return " ".join(t for _, t in partes)


PLANTILLA = """<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITULO__</title>
<style>
  :root {
    --bg:#f6f7f5; --panel:#fff; --tinta:#1c1c1a; --suave:#6b6b66; --linea:#e2e2dd;
    --acento:#4a6b3d; --alerta:#a8442a; --ok:#3d6b4a; --tibio:#8a6d24;
  }
  @media (prefers-color-scheme: dark) { :root:not([data-theme="light"]) {
    --bg:#16181a; --panel:#1e2124; --tinta:#e8e8e4; --suave:#9a9a94; --linea:#2f3336;
    --acento:#8fb87a; --alerta:#e08a6d; --ok:#7fb894; --tibio:#d4b463;
  } }
  :root[data-theme="dark"] {
    --bg:#16181a; --panel:#1e2124; --tinta:#e8e8e4; --suave:#9a9a94; --linea:#2f3336;
    --acento:#8fb87a; --alerta:#e08a6d; --ok:#7fb894; --tibio:#d4b463;
  }
  * { box-sizing:border-box; }
  body { margin:0; background:var(--bg); color:var(--tinta); font:15px/1.5
    -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif; }
  .wrap { max-width:1180px; margin:0 auto; padding:28px 16px 64px; }
  h1 { font-size:22px; margin:0 0 4px; letter-spacing:-.01em; }
  .sub { color:var(--suave); font-size:13px; margin-bottom:22px; }
  .kpis { display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr));
    gap:10px; margin-bottom:20px; }
  .kpi { background:var(--panel); border:1px solid var(--linea); border-radius:9px;
    padding:12px 14px; }
  .kpi .v { font-size:20px; font-weight:600; letter-spacing:-.02em; }
  .kpi .l { font-size:11px; color:var(--suave); text-transform:uppercase;
    letter-spacing:.05em; margin-top:2px; }
  .barra { display:flex; gap:8px; flex-wrap:wrap; align-items:center; margin-bottom:14px; }
  input[type=search] { flex:1; min-width:200px; padding:9px 12px; font-size:14px;
    border:1px solid var(--linea); border-radius:8px; background:var(--panel);
    color:var(--tinta); }
  .chip { padding:7px 12px; font-size:13px; border:1px solid var(--linea);
    border-radius:999px; background:var(--panel); color:var(--suave); cursor:pointer; }
  .chip[aria-pressed=true] { background:var(--acento); border-color:var(--acento);
    color:#fff; }
  .cont { background:var(--panel); border:1px solid var(--linea); border-radius:10px;
    overflow:auto; }
  table { border-collapse:collapse; width:100%; font-size:14px; }
  th { text-align:left; font-size:11px; text-transform:uppercase; letter-spacing:.05em;
    color:var(--suave); padding:11px 14px; border-bottom:1px solid var(--linea);
    position:sticky; top:0; background:var(--panel); cursor:pointer;
    white-space:nowrap; }
  th:hover { color:var(--tinta); }
  td { padding:11px 14px; border-bottom:1px solid var(--linea); vertical-align:top; }
  tr:last-child td { border-bottom:none; }
  .cli { font-weight:600; min-width:200px; }
  .cli .nro { display:block; font-weight:400; font-size:11px; color:var(--suave); }
  .lc { text-align:right; font-variant-numeric:tabular-nums; white-space:nowrap;
    font-weight:600; font-size:15px; }
  .com { color:var(--suave); font-size:13px; line-height:1.55; }
  .com b { color:var(--alerta); font-weight:600; }
  .com i { color:var(--ok); font-style:normal; }
  .com u { color:var(--tibio); text-decoration:none; }
  .pie { color:var(--suave); font-size:12px; margin-top:18px; line-height:1.6; }
  .vacio { padding:34px; text-align:center; color:var(--suave); }
  @media print { .barra,.kpis { display:none; } .cont { border:none; } body { background:#fff; } }
</style></head><body><div class="wrap">
<h1>__TITULO__</h1>
<div class="sub">__SUB__</div>
<div class="kpis">__KPIS__</div>
<div class="barra">
  <input type="search" id="q" placeholder="Buscar cliente...">
  <button class="chip" id="fExc" aria-pressed="false">Arrancan excedidos</button>
  <button class="chip" id="fFric" aria-pressed="false">Se exceden +45 dias</button>
  <button class="chip" id="fCae" aria-pressed="false">Cayeron fuerte</button>
</div>
<div class="cont"><table>
<thead><tr><th data-k="cliente">Cliente</th><th data-k="lc" style="text-align:right">LC asignado (US$)</th><th data-k="com">Comentarios</th></tr></thead>
<tbody id="tb"></tbody></table><div class="vacio" id="vacio" hidden>Sin resultados.</div></div>
<div class="pie">__PIE__</div>
</div>
<script>
const D = __DATOS__;
const tb = document.getElementById('tb'), vacio = document.getElementById('vacio');
let orden = {k:'lc', desc:true};
const esc = s => String(s).replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
const TAG = {alerta:'b', ok:'i', aviso:'u'};
const realza = partes => partes.map(([tipo, t]) =>
  TAG[tipo] ? `<${TAG[tipo]}>${esc(t)}</${TAG[tipo]}>` : esc(t)).join(' ');
function visibles() {
  const q = document.getElementById('q').value.trim().toLowerCase();
  const fe = document.getElementById('fExc').getAttribute('aria-pressed') === 'true';
  const ff = document.getElementById('fFric').getAttribute('aria-pressed') === 'true';
  const fc = document.getElementById('fCae').getAttribute('aria-pressed') === 'true';
  let r = D.filter(d => !q || d.cliente.toLowerCase().includes(q) || d.nro.toLowerCase().includes(q));
  if (fe) r = r.filter(d => d.exc);
  if (ff) r = r.filter(d => d.dias > 45);
  if (fc) r = r.filter(d => d.cae);
  const s = orden.desc ? -1 : 1;
  return r.sort((a, b) => {
    const x = a[orden.k], y = b[orden.k];
    return (typeof x === 'number' ? x - y : String(x).localeCompare(String(y), 'es')) * s;
  });
}
function pintar() {
  const r = visibles();
  vacio.hidden = r.length > 0;
  tb.innerHTML = r.map(d =>
    `<tr><td class="cli">${esc(d.cliente)}<span class="nro">${esc(d.nro)}</span></td>` +
    `<td class="lc">${d.lcf}</td><td class="com">${realza(d.com)}</td></tr>`).join('');
}
document.getElementById('q').addEventListener('input', pintar);
for (const id of ['fExc', 'fFric', 'fCae']) {
  document.getElementById(id).addEventListener('click', e => {
    const b = e.currentTarget;
    b.setAttribute('aria-pressed', b.getAttribute('aria-pressed') === 'true' ? 'false' : 'true');
    pintar();
  });
}
document.querySelectorAll('th[data-k]').forEach(th => th.addEventListener('click', () => {
  const k = th.dataset.k;
  orden = {k, desc: orden.k === k ? !orden.desc : k === 'lc'};
  pintar();
}));
pintar();
</script></body></html>
"""


def render_html(res, titulo, sub, kpis, pie):
    datos = []
    for _, r in res.iterrows():
        alertas = str(r.get("Alertas") or "")
        datos.append({
            "nro": str(r.get("Nro cliente") or ""),
            "cliente": str(r["Cliente"]),
            "lc": float(r["LIMITE SUGERIDO USD"]),
            "lcf": fmt(r["LIMITE SUGERIDO USD"]),
            "com": r["_partes"],
            "dias": int(r.get("Dias/ano sobre el limite") or 0),
            "exc": "ARRANCA EXCEDIDO" in alertas,
            "cae": "cayo " in alertas,
        })
    kpis_html = "".join(
        f'<div class="kpi"><div class="v">{v}</div><div class="l">{l}</div></div>'
        for l, v in kpis)
    return (PLANTILLA
            .replace("__TITULO__", html_mod.escape(titulo))
            .replace("__SUB__", html_mod.escape(sub))
            .replace("__KPIS__", kpis_html)
            .replace("__PIE__", pie)
            .replace("__DATOS__", json.dumps(datos, ensure_ascii=False)))


def main():
    ap = argparse.ArgumentParser(description="Vista ejecutiva del LC por cliente.")
    ap.add_argument("detalle", help="Excel generado por asignar_creditos.py")
    ap.add_argument("--salida-xlsx")
    ap.add_argument("--salida-html")
    args = ap.parse_args()

    det = pd.read_excel(args.detalle, sheet_name="Limites sugeridos")
    crit = pd.read_excel(args.detalle, sheet_name="Criterios")
    criterios = dict(zip(crit["Parametro"], crit["Valor"]))

    det["_partes"] = det.apply(construir_comentario, axis=1)
    det["Comentarios"] = det["_partes"].map(texto_plano)
    det = det.sort_values("LIMITE SUGERIDO USD", ascending=False).reset_index(drop=True)

    resumen = pd.DataFrame({
        "Cliente": det["Nro cliente"].fillna("").astype(str) + "  " + det["Cliente"],
        "LC asignado (US$)": det["LIMITE SUGERIDO USD"],
        "Comentarios": det["Comentarios"],
    })

    destino_x = Path(args.salida_xlsx or (BASE_DIR / "LC_resumen.xlsx"))
    with pd.ExcelWriter(destino_x, engine="openpyxl") as xl:
        resumen.to_excel(xl, sheet_name="LC por cliente", index=False)
        det.drop(columns=["_partes"]).to_excel(xl, sheet_name="Detalle", index=False)
        crit.to_excel(xl, sheet_name="Criterios", index=False)
        ws = xl.sheets["LC por cliente"]
        ws.freeze_panes = "A2"
        ws.column_dimensions["A"].width = 46
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 120
        for fila in ws.iter_rows(min_row=2, min_col=2, max_col=2):
            for c in fila:
                c.number_format = "#,##0"
        for fila in ws.iter_rows(min_row=1, max_row=1):
            for c in fila:
                c.font = c.font.copy(bold=True)
        for fila in ws.iter_rows(min_row=2, min_col=3, max_col=3):
            for c in fila:
                c.alignment = c.alignment.copy(wrap_text=True, vertical="top")

    total = det["LIMITE SUGERIDO USD"].sum()
    n_exc = det["Alertas"].astype(str).str.contains("ARRANCA EXCEDIDO").sum()
    n_fric = (det["Dias/ano sobre el limite"] > 45).sum()
    n_ok = (det["Dias/ano sobre el limite"] == 0).sum()
    kpis = [("Clientes", f"{len(det):,}".replace(",", ".")),
            ("Suma de LC", f"US$ {fmt(total)}"),
            ("LC promedio", f"US$ {fmt(total / len(det))}"),
            ("Sin excederse nunca", f"{n_ok}"),
            ("Se exceden +45 dias", f"{n_fric}"),
            ("Arrancan excedidos", f"{n_exc}")]

    titulo = "Limite de credito por cliente - Philagro S.A."
    sub = (f"{criterios.get('Periodo analizado', '')} &middot; corte "
           f"{criterios.get('Fecha de corte', '')} &middot; base "
           f"{criterios.get('Base de calculo', '')} &middot; canje/plataforma/tarjeta: "
           f"{criterios.get('Plazo para canje/plataforma/tarjeta', '')}")
    pie = ("Propuesta basada <b>solo en comportamiento comercial</b> (que compro cada "
           "cliente, a que plazo y con que continuidad). Antes de aprobar hay que "
           "contrastarla con balances, indices, antiguedad en el rubro, deuda tomada con "
           "terceros y garantias.<br>El LC cubre el pico de saldo que genera la "
           "actividad actual del cliente; el cupo se libera al vencer cada factura. "
           f"Generado el {datetime.now():%d/%m/%Y %H:%M}.")

    destino_h = Path(args.salida_html or (BASE_DIR / "LC_resumen.html"))
    destino_h.write_text(
        render_html(det, titulo, html_mod.unescape(sub.replace("&middot;", "·")),
                    kpis, pie), encoding="utf-8")

    print(f"Excel: {destino_x}")
    print(f"HTML : {destino_h}")
    print(f"Clientes: {len(det)}   Suma de LC: US$ {fmt(total)}")


if __name__ == "__main__":
    main()
