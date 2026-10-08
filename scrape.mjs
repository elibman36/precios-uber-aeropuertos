// Lee los precios que publica Uber en sus páginas de rutas y los agrega a data/precios.csv.
// Uso: node scrape.mjs
import { chromium } from 'playwright';
import fs from 'node:fs';

const RUTAS = JSON.parse(fs.readFileSync('rutas.json', 'utf8'));
const CSV = 'data/precios.csv';
const DEBUG_DIR = 'debug';

// Nombres de productos de Uber en Argentina (para asociar cada precio a su producto)
const PRODUCTOS = /^(Uber\s?X|UberX|Comfort|Uber\s?Comfort|Black|Uber\s?Black|XL|Uber\s?XL|Flash|Uber\s?Flash|Moto|Uber\s?Moto|Planet|Uber\s?Planet|Pet|Uber\s?Pet|Taxi|Uber\s?Taxi|Premier|Green|Uber\s?Green|Van|Priority|Prioridad|Ahorro|Espera y ahorra|Wait and save)\b/i;

// "ARS 25.000", "$25.000 - $32.000", "ARS 25.000–32.000", "ARS25,000"
const PRECIO = /(ARS|\$)\s?([\d.,]+)(?:\s?[-–]\s?(?:ARS|\$)?\s?([\d.,]+))?/;

// Formato argentino: "25.000" → 25000, "25.000,50" → 25000.5; también tolera "25,000"
function numero(txt) {
  if (!txt) return '';
  const limpio = txt.replace(/[.,](?=\d{3}(\D|$))/g, '').replace(',', '.');
  const n = parseFloat(limpio);
  return Number.isFinite(n) ? Math.round(n) : '';
}

function csvCampo(v) {
  const s = String(v ?? '');
  return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

// Recorre el texto visible línea por línea: cada precio se asocia al último producto visto.
function extraerPrecios(texto) {
  const lineas = texto.split('\n').map(l => l.trim()).filter(Boolean);
  const filas = [];
  let producto = '';
  lineas.forEach((linea, i) => {
    if (PRODUCTOS.test(linea) && linea.length < 40) producto = linea;
    const m = linea.match(PRECIO);
    if (!m) return;
    const min = numero(m[2]);
    if (min === '' || min < 1000) return; // descarta números sueltos que no son tarifas
    filas.push({
      producto: producto || '(sin producto)',
      precio_min: min,
      precio_max: numero(m[3]) || min,
      moneda: 'ARS',
      texto: linea.slice(0, 120),
      contexto: lineas.slice(Math.max(0, i - 3), i + 2).join(' | '),
    });
  });
  // Uber a veces repite el mismo bloque (mobile/desktop): deduplicar
  const vistas = new Set();
  return filas.filter(f => {
    const k = `${f.producto}|${f.precio_min}|${f.precio_max}`;
    if (vistas.has(k)) return false;
    vistas.add(k);
    return true;
  });
}

const ahora = new Date();
const fecha = ahora.toLocaleDateString('en-CA', { timeZone: 'America/Argentina/Buenos_Aires' });
const hora = ahora.toLocaleTimeString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires', hour: '2-digit', minute: '2-digit', hour12: false });

fs.mkdirSync(DEBUG_DIR, { recursive: true });
const browser = await chromium.launch();
const context = await browser.newContext({
  locale: 'es-AR',
  timezoneId: 'America/Argentina/Buenos_Aires',
  userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36',
  viewport: { width: 1366, height: 900 },
});

const nuevas = [];
const linksRutas = new Set();
let fallas = 0;

for (const ruta of RUTAS) {
  if (!ruta.url) {
    console.log(`\n== ${ruta.id}: sin URL configurada, se busca en los links de las otras páginas`);
    continue;
  }
  console.log(`\n== ${ruta.id}: ${ruta.url}`);
  const page = await context.newPage();
  try {
    const resp = await page.goto(ruta.url, { waitUntil: 'domcontentloaded', timeout: 60000 });
    console.log(`HTTP ${resp?.status()} · título: ${await page.title()}`);
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    // Los precios pueden cargarse más abajo en la página
    for (let i = 0; i < 6; i++) { await page.mouse.wheel(0, 1200); await page.waitForTimeout(700); }

    const texto = await page.evaluate(() => document.body.innerText);
    fs.writeFileSync(`${DEBUG_DIR}/${ruta.id}.txt`, texto);
    fs.writeFileSync(`${DEBUG_DIR}/${ruta.id}.html`, await page.content());
    await page.screenshot({ path: `${DEBUG_DIR}/${ruta.id}.png`, fullPage: true }).catch(() => {});

    for (const href of await page.$$eval('a[href*="/r/routes/"]', as => as.map(a => a.href))) linksRutas.add(href);

    const filas = extraerPrecios(texto);
    if (filas.length === 0) {
      fallas++;
      console.log('!! No se encontraron precios. Líneas del texto con números grandes o "$"/"ARS":');
      texto.split('\n').filter(l => /ARS|\$|\d{2}[.,]\d{3}/.test(l)).slice(0, 40).forEach(l => console.log('   ', l.trim().slice(0, 160)));
      console.log('!! Primeras 80 líneas de la página:');
      texto.split('\n').map(l => l.trim()).filter(Boolean).slice(0, 80).forEach(l => console.log('   ', l.slice(0, 160)));
    }
    for (const f of filas) {
      console.log(`  ${f.producto}: ${f.precio_min}–${f.precio_max}   [${f.contexto.slice(0, 150)}]`);
      nuevas.push([fecha, hora, ruta.id, f.producto, f.precio_min, f.precio_max, f.moneda, f.texto]);
    }
  } catch (e) {
    fallas++;
    console.log(`!! Error: ${e.message}`);
  } finally {
    await page.close();
  }
}
await browser.close();

if (RUTAS.some(r => !r.url)) {
  console.log('\n== Links a otras rutas encontrados (para completar rutas.json):');
  [...linksRutas].filter(h => /aeroparque|newbery|aep|ezeiza/i.test(h)).forEach(h => console.log('   ', h));
  console.log(`   (${linksRutas.size} links de rutas en total)`);
  [...linksRutas].slice(0, 60).forEach(h => console.log('    ·', h));
}

if (nuevas.length) {
  fs.appendFileSync(CSV, nuevas.map(f => f.map(csvCampo).join(',')).join('\n') + '\n');
  console.log(`\nSe agregaron ${nuevas.length} filas a ${CSV}`);
}
process.exit(fallas ? 1 : 0);
