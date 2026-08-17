#!/usr/bin/env node
/**
 * Verificación headless del dashboard generado.
 *
 *     node verificar.mjs [dashboard.html]
 *
 * Abre el HTML en Chromium y chequea:
 *   1. Cero errores de consola / excepciones de página.
 *   2. Que el cuadro principal CIERRE: filas, columnas, SUB TOTAL y TOTAL.
 *   3. Que el cuadro Contado/Diferido cierre y coincida con el total del cuadro.
 *   4. Que KPIs, filtros y gráficos hayan renderizado.
 *   5. Que el archivo sea autocontenido (sin fetch/CDN/internet).
 *
 * Sale con código 1 si algo falla.
 */
import { createRequire } from "node:module";
import { readFileSync, existsSync } from "node:fs";
import { resolve } from "node:path";
import { pathToFileURL } from "node:url";

// Playwright está instalado global en este entorno; si el proyecto lo tiene
// local, ese gana.
const require_ = createRequire(import.meta.url);
let chromium;
try {
  ({ chromium } = require_("playwright"));
} catch {
  ({ chromium } = require_("/opt/node22/lib/node_modules/playwright"));
}

const archivo = resolve(process.argv[2] || "dashboard.html");
if (!existsSync(archivo)) {
  console.error(`No existe ${archivo}`);
  process.exit(1);
}

const fallas = [];
const oks = [];
const check = (cond, msg) => (cond ? oks.push(msg) : fallas.push(msg));

// ---- 1) Autocontenido: nada de red ----------------------------------------
const html = readFileSync(archivo, "utf8");
check(!/<script[^>]+\ssrc=/i.test(html), "Sin <script src=…> externo");
check(!/<link[^>]+stylesheet/i.test(html), "Sin hoja de estilos externa");
check(!/\bfetch\s*\(|XMLHttpRequest|cdn\.jsdelivr|unpkg\.com|cdnjs/i.test(html),
  "Sin fetch/XHR/CDN");

// ---- 2) Render en Chromium -------------------------------------------------
const browser = await chromium.launch();
const page = await browser.newPage();
const errores = [];
page.on("console", (m) => { if (m.type() === "error") errores.push(m.text()); });
page.on("pageerror", (e) => errores.push(String(e)));
page.on("request", (r) => {
  if (!r.url().startsWith("file:") && !r.url().startsWith("data:")) {
    errores.push(`pedido de red: ${r.url()}`);
  }
});

await page.goto(pathToFileURL(archivo).href, { waitUntil: "load" });
await page.waitForTimeout(600);

const r = await page.evaluate(() => {
  // "1.234" -> 1234 ; "(1.234)" -> -1234 ; "–" -> 0
  const num = (t) => {
    const s = (t || "").trim();
    if (!s || s === "–" || s === "-") return 0;
    const neg = s.startsWith("(");
    const n = parseInt(s.replace(/[^0-9]/g, ""), 10) || 0;
    return neg ? -n : n;
  };
  const celdas = (tr) => Array.from(tr.querySelectorAll("td")).slice(1).map((td) => num(td.textContent));

  const out = { errores: [] };

  // --- Cuadro principal (forma de cobro × mes) ---
  const tabla = document.querySelector("table.pivot:not(.cd)");
  const filas = Array.from(tabla.querySelectorAll("tbody tr"));
  const pie = tabla.querySelector("tfoot tr");
  const nCols = celdas(pie).length - 1; // sin la columna Total

  const conceptos = filas.filter((f) => !f.classList.contains("subtotal"));
  const subtotal = filas.find((f) => f.classList.contains("subtotal"));

  // a) cada fila: suma de meses = su columna Total
  for (const f of filas.concat([pie])) {
    const v = celdas(f);
    const meses = v.slice(0, nCols);
    const tot = v[nCols];
    const suma = meses.reduce((a, b) => a + b, 0);
    if (suma !== tot) {
      out.errores.push(`fila "${f.cells[0].textContent.trim()}": suma ${suma} ≠ total ${tot}`);
    }
  }

  // b) cada columna: suma de conceptos + Caja Comp. = fila TOTAL
  const totPie = celdas(pie);
  for (let j = 0; j <= nCols; j++) {
    const suma = conceptos.reduce((a, f) => a + celdas(f)[j], 0);
    if (suma !== totPie[j]) out.errores.push(`columna ${j}: suma ${suma} ≠ TOTAL ${totPie[j]}`);
  }

  // c) SUB TOTAL = conceptos sin Caja Compensación
  if (subtotal) {
    const cc = conceptos.find((f) => /Caja Compensación/.test(f.cells[0].textContent));
    const sinCC = conceptos.filter((f) => f !== cc);
    const sub = celdas(subtotal);
    for (let j = 0; j <= nCols; j++) {
      const suma = sinCC.reduce((a, f) => a + celdas(f)[j], 0);
      if (suma !== sub[j]) out.errores.push(`SUB TOTAL col ${j}: ${suma} ≠ ${sub[j]}`);
    }
  }
  out.totalCuadro = totPie[nCols];

  // --- Cuadro Contado / Diferido ---
  const cd = document.querySelector("table.pivot.cd");
  const cdPie = cd.querySelector("tfoot tr");
  const cdFilas = Array.from(cd.querySelectorAll("tbody tr"));
  const cdCols = celdas(cdPie).length; // Contado, Diferido, Caja, Reten, Total
  for (const f of cdFilas) {
    const v = celdas(f);
    const suma = v[0] + v[1] + v[cdCols - 3] + v[cdCols - 2];
    if (suma !== v[cdCols - 1]) {
      out.errores.push(`C/D fila "${f.cells[0].textContent.trim()}": ${suma} ≠ ${v[cdCols - 1]}`);
    }
  }
  for (let j = 0; j < cdCols; j++) {
    const suma = cdFilas.reduce((a, f) => a + celdas(f)[j], 0);
    if (suma !== celdas(cdPie)[j]) out.errores.push(`C/D columna ${j}: ${suma} ≠ ${celdas(cdPie)[j]}`);
  }
  out.totalCD = celdas(cdPie)[cdCols - 1];

  // --- UI mínima ---
  out.kpis = document.querySelectorAll("#kpis .kpi").length;
  out.filtros = document.querySelectorAll(".ms .ms-btn").length;
  out.canvas = Array.from(document.querySelectorAll("canvas"))
    .filter((c) => c.width > 0 && c.height > 0).length;
  out.contador = (document.getElementById("contador").textContent || "").trim();
  out.sub = (document.getElementById("subheader").textContent || "").trim();
  return out;
});

// ---- 3) Filtros: que el panel abra y cierre --------------------------------
await page.click("#ms-anio .ms-btn");
const abierto = await page.isVisible("#ms-anio .ms-panel");
await page.click("#ms-anio .ms-btn");
const cerrado = !(await page.isVisible("#ms-anio .ms-panel"));

// ---- 4) Contado/Diferido: desplegar la columna de acreditación -------------
await page.click("#cd-tg");
await page.waitForTimeout(200);
const acreditacion = await page.locator("table.cd th.acre").count();

await browser.close();

check(errores.length === 0, `Cero errores de consola${errores.length ? ` (${errores.length})` : ""}`);
check(r.errores.length === 0, `Los cuadros cierran exacto${r.errores.length ? ` (${r.errores.length} desvíos)` : ""}`);
check(r.totalCuadro === r.totalCD, `Total cuadro (${r.totalCuadro}) = total Contado/Diferido (${r.totalCD})`);
check(r.kpis === 4, `4 KPIs renderizados (${r.kpis})`);
check(r.filtros === 3, `3 filtros multi-selección (${r.filtros})`);
check(r.canvas >= 4, `Gráficos dibujados (${r.canvas} canvas)`);
check(abierto && cerrado, "El panel de filtro abre y cierra");
check(acreditacion > 0, `Diferido se despliega por mes de acreditación (${acreditacion} columnas)`);
const [vistos, totales] = (r.contador.match(/[\d.]+/g) || []).map((n) => parseInt(n.replace(/\./g, ""), 10));
check(/cobros/.test(r.contador) && vistos === totales,
  `Contador de cobros sin filtros: "${r.contador}"`);

for (const e of errores.slice(0, 10)) console.error("  consola:", e);
for (const e of r.errores.slice(0, 15)) console.error("  cuadro:", e);

console.log(`\nVerificación de ${archivo}`);
console.log(r.sub);
for (const o of oks) console.log("  ok   ", o);
for (const f of fallas) console.log("  FALLA", f);
console.log(fallas.length ? `\n${fallas.length} verificación(es) fallida(s).` : "\nTodo OK.");
process.exit(fallas.length ? 1 : 0);
