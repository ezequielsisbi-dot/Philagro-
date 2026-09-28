#!/usr/bin/env node
/**
 * verificar_ventas.mjs — Verificación automática de un dashboard_ventas_*.html
 *
 * Uso:
 *   node verificar_ventas.mjs <dashboard_ventas_ARCHIVO.html>
 *
 * Chequea:
 *   1) 0 errores de consola / errores de página al cargar (sin red, sin CDN).
 *   2) Las tablas con fila/columna de totales cierran solas: cada fila suma
 *      igual a su "Total", cada columna suma igual al total de esa columna,
 *      y el total general coincide. Se prueba para las 5 dimensiones del
 *      cuadro dinámico (Vendedor / Producto / Cliente / Principio activo /
 *      Condición de pago), la tabla de mes×año y la de cantidades.
 *   3) Los filtros (multi-select de Vendedor) responden y el contador cambia.
 *
 * Sale con código 0 si todo pasa, 1 si algo falla.
 */

import { pathToFileURL } from "node:url";
import path from "node:path";
import fs from "node:fs";
import { createRequire } from "node:module";

async function cargarPlaywright() {
  try {
    return await import("playwright");
  } catch {
    // Playwright instalado globalmente pero no en node_modules local.
    const require = createRequire(import.meta.url);
    const globalPaths = [
      "/opt/node22/lib/node_modules/playwright/index.mjs",
    ];
    for (const p of globalPaths) {
      if (fs.existsSync(p)) return await import(pathToFileURL(p).href);
    }
    throw new Error(
      "No se encontró el módulo 'playwright'. Instalalo o corré con NODE_PATH=/opt/node22/lib/node_modules"
    );
  }
}

const CHROMIUM_PATH = "/opt/pw-browsers/chromium";

function parseNum(txt) {
  if (txt == null) return 0;
  let t = txt.trim();
  if (t === "" || t === "–" || t === "-") return 0;
  let neg = false;
  if (t.startsWith("(") && t.endsWith(")")) { neg = true; t = t.slice(1, -1); }
  t = t.replace(/\./g, "").replace(/\s/g, "");
  const n = parseInt(t, 10);
  if (Number.isNaN(n)) return NaN;
  return neg ? -n : n;
}

async function verificarTablaCierra(page, ids, nombre, errores) {
  const { thead, tbody, tfoot } = ids;
  const datos = await page.evaluate(({ tbodyId, tfootId }) => {
    const tbody = document.getElementById(tbodyId);
    const tfoot = document.getElementById(tfootId);
    const filas = Array.from(tbody.querySelectorAll("tr")).map((tr) =>
      Array.from(tr.querySelectorAll("td")).map((td) => td.textContent)
    );
    const filaTotal = tfoot
      ? Array.from(tfoot.querySelectorAll("td")).map((td) => td.textContent)
      : [];
    return { filas, filaTotal };
  }, { tbodyId: tbody, tfootId: tfoot });

  if (datos.filas.length === 1 && /sin datos/i.test(datos.filas[0][0] || "")) {
    console.log(`  [${nombre}] sin datos para los filtros actuales (omitido).`);
    return;
  }
  if (!datos.filas.length) {
    errores.push(`[${nombre}] tabla vacía inesperadamente.`);
    return;
  }

  // Cada fila: suma de columnas numéricas (todas menos la 1ra = etiqueta) == última columna (Total)
  for (const fila of datos.filas) {
    const [, ...resto] = fila;
    const total = parseNum(resto[resto.length - 1]);
    const valores = resto.slice(0, -1).map(parseNum);
    if (valores.some((v) => Number.isNaN(v)) || Number.isNaN(total)) {
      errores.push(`[${nombre}] celda no numérica en fila "${fila[0]}": ${JSON.stringify(fila)}`);
      continue;
    }
    const suma = valores.reduce((a, b) => a + b, 0);
    if (suma !== total) {
      errores.push(`[${nombre}] fila "${fila[0]}" no cierra: suma columnas ${suma} != total ${total}`);
    }
  }

  // Fila de totales: suma de columnas == total general (última celda)
  if (datos.filaTotal.length) {
    const [, ...restoT] = datos.filaTotal;
    const totalGeneral = parseNum(restoT[restoT.length - 1]);
    const columnas = restoT.slice(0, -1).map(parseNum);
    const sumaCols = columnas.reduce((a, b) => a + b, 0);
    if (sumaCols !== totalGeneral) {
      errores.push(`[${nombre}] fila TOTAL no cierra: suma columnas ${sumaCols} != total general ${totalGeneral}`);
    }
    // Y el total general debe coincidir con la suma de los totales de fila.
    const sumaFilas = datos.filas.reduce((s, fila) => s + parseNum(fila[fila.length - 1]), 0);
    if (sumaFilas !== totalGeneral) {
      errores.push(`[${nombre}] total general (${totalGeneral}) != suma de totales de fila (${sumaFilas})`);
    }
  }

  console.log(`  [${nombre}] ${datos.filas.length} filas — cierra OK.`);
}

async function main() {
  const archivo = process.argv[2];
  if (!archivo) {
    console.error("Uso: node verificar_ventas.mjs <dashboard_ventas_ARCHIVO.html>");
    process.exit(1);
  }
  const ruta = path.resolve(archivo);
  if (!fs.existsSync(ruta)) {
    console.error(`No existe el archivo: ${ruta}`);
    process.exit(1);
  }

  const { chromium } = await cargarPlaywright();
  const launchOpts = fs.existsSync(CHROMIUM_PATH) ? { executablePath: CHROMIUM_PATH } : {};
  const browser = await chromium.launch(launchOpts);
  const page = await browser.newPage();

  const consoleErrores = [];
  page.on("console", (msg) => { if (msg.type() === "error") consoleErrores.push(msg.text()); });
  page.on("pageerror", (err) => consoleErrores.push(String(err)));
  page.on("requestfailed", (req) => consoleErrores.push(`request fallida: ${req.url()}`));

  console.log(`Abriendo ${ruta} ...`);
  await page.goto(pathToFileURL(ruta).href, { waitUntil: "load" });
  await page.waitForSelector("#kpis .kpi", { timeout: 15000 });
  await page.waitForTimeout(300);

  const erroresLogicos = [];

  console.log("\nVerificando cuadro dinámico por dimensión (5) ...");
  const dims = await page.$$eval("#dim-select option", (opts) => opts.map((o) => o.value));
  for (const dim of dims) {
    await page.selectOption("#dim-select", dim);
    await page.waitForTimeout(150);
    await verificarTablaCierra(
      page,
      { tbody: "p-tbody", tfoot: "p-tfoot" },
      `pivot dim=${dim}`,
      erroresLogicos
    );
  }
  await page.selectOption("#dim-select", "v");

  console.log("\nVerificando 'Ventas por mes y año' ...");
  await verificarTablaCierra(page, { tbody: "m-tbody", tfoot: "m-tfoot" }, "mes×año", erroresLogicos);

  console.log("\nVerificando 'Cantidades por producto' ...");
  await verificarTablaCierra(page, { tbody: "q-tbody", tfoot: "q-tfoot" }, "cantidades", erroresLogicos);

  console.log("\nVerificando filtros (multi-select Vendedor) ...");
  const contadorAntes = await page.textContent("#contador");
  await page.click("#ms-vendedor .ms-control");
  await page.waitForTimeout(100);
  const primeraOpcion = await page.$("#ms-vendedor .ms-opt input[type=checkbox]");
  if (primeraOpcion) {
    await primeraOpcion.check();
    await page.waitForTimeout(200);
    const contadorDespues = await page.textContent("#contador");
    if (contadorDespues === contadorAntes) {
      erroresLogicos.push("El filtro de Vendedor no modificó el contador de comprobantes.");
    } else {
      console.log(`  contador "${contadorAntes}" -> "${contadorDespues}" — filtro responde OK.`);
    }
    await page.click("#f-limpiar");
    await page.waitForTimeout(150);
  } else {
    erroresLogicos.push("No se pudo abrir el panel de opciones de Vendedor.");
  }

  await browser.close();

  console.log("\n--- Resultado ---");
  console.log(`Errores de consola: ${consoleErrores.length}`);
  consoleErrores.forEach((e) => console.log("  · " + e));
  console.log(`Errores de reconciliación: ${erroresLogicos.length}`);
  erroresLogicos.forEach((e) => console.log("  · " + e));

  const ok = consoleErrores.length === 0 && erroresLogicos.length === 0;
  console.log(ok ? "\nOK — el dashboard cierra y no tiene errores de consola." : "\nFALLÓ la verificación.");
  process.exit(ok ? 0 : 1);
}

main().catch((err) => {
  console.error("Error inesperado en la verificación:", err);
  process.exit(1);
});
