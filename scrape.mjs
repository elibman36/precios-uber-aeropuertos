// Lee los precios que publica Uber en sus páginas de rutas y los agrega a data/precios.csv.
// Uso: node scrape.mjs
import { chromium } from 'playwright';
import fs from 'node:fs';

const RUTAS = JSON.parse(fs.readFileSync('rutas.json', 'utf8'));
const CSV = 'data/precios.csv';
const DEBUG_DIR = 'debug';

// "$34172*", "$ 25.000", "ARS 25.000-32.000"
const PRECIO = /(ARS|\$)\s?([\d.,]+)(?:\s?[-–]\s?(?:ARS|\$)?\s?([\d.,]+))?/;
// Líneas que no son nombre de producto: asientos ("4"), rating ("4.9"), rótulos sueltos
const NO_NOMBRE = /^(\d+([.,]\d+)?|Precio promedio)$/i;

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

// La página de ruta de Uber muestra cada producto como:
//   UberX | 4 (asientos) | 4.9 (rating) | $34172* | descripción
// y arriba un resumen: "Tarifa promedia de la ruta | $32972", "Tiempo promedio de viaje | 45 minutos".
function extraerPrecios(texto) {
  const lineas = texto.split('\n').map(l => l.trim()).filter(Boolean);
  const valorDe = re => {
    const i = lineas.findIndex(l => re.test(l));
    return i >= 0 && lineas[i + 1] ? numero((lineas[i + 1].match(/[\d.,]+/) || [])[0]) : '';
  };
  const resumen = {
    tiempo_min: valorDe(/^Tiempo promedio/i),
    distancia_km: valorDe(/^Distancia promedio/i),
  };
  const filas = [];
  lineas.forEach((linea, i) => {
    const m = linea.match(PRECIO);
    if (!m || m.index > 2) return; // el precio va al principio de la línea
    const precio = numero(m[2]);
    if (precio === '' || precio < 1000) return;
    let j = i - 1;
    while (j >= 0 && (NO_NOMBRE.test(lineas[j]) || PRECIO.test(lineas[j]))) j--;
    let producto = j >= 0 ? lineas[j] : '(sin nombre)';
    if (/^Tarifa promedi/i.test(producto)) producto = 'Promedio de la ruta';
    filas.push({ producto: producto.slice(0, 40), precio, precio_max: numero(m[3]) || precio,
                 contexto: lineas.slice(Math.max(0, i - 3), i + 2).join(' | ') });
  });
  // Uber a veces repite el mismo bloque (mobile/desktop): deduplicar
  const vistas = new Set();
  return {
    resumen,
    filas: filas.filter(f => {
      const k = `${f.producto}|${f.precio}`;
      if (vistas.has(k)) return false;
      vistas.add(k);
      return true;
    }),
  };
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
  const urls = [].concat(ruta.url || []);
  if (!urls.length) {
    console.log(`\n== ${ruta.id}: sin URL configurada, se busca en los links de las otras páginas`);
    continue;
  }
  let page;
  try {
    // "url" puede ser una lista: se usa la primera que responde con una página de ruta
    let ok = false;
    for (const url of urls) {
      console.log(`\n== ${ruta.id}: ${url}`);
      await page?.close();
      page = await context.newPage(); // página nueva por intento: un error no contamina el siguiente
      try {
        const resp = await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
        console.log(`HTTP ${resp?.status()} · título: ${await page.title()} · final: ${page.url()}`);
        if ((!resp || resp.ok()) && /\/r\/routes\//.test(page.url())) { ok = true; break; }
      } catch (e) {
        console.log(`   no responde: ${e.message.split('\n')[0]}`);
      }
    }
    if (!ok) throw new Error('ninguna URL respondió con una página de ruta');
    await page.waitForLoadState('networkidle', { timeout: 30000 }).catch(() => {});
    // Los precios pueden cargarse más abajo en la página
    for (let i = 0; i < 6; i++) { await page.mouse.wheel(0, 1200); await page.waitForTimeout(700); }

    const texto = await page.evaluate(() => document.body.innerText);
    fs.writeFileSync(`${DEBUG_DIR}/${ruta.id}.txt`, texto);
    fs.writeFileSync(`${DEBUG_DIR}/${ruta.id}.html`, await page.content());
    await page.screenshot({ path: `${DEBUG_DIR}/${ruta.id}.png`, fullPage: true }).catch(() => {});

    for (const href of await page.$$eval('a[href*="/r/routes/"]', as => as.map(a => a.href))) linksRutas.add(href);

    const { resumen, filas } = extraerPrecios(texto);
    console.log(`  tiempo promedio: ${resumen.tiempo_min} min · distancia promedio: ${resumen.distancia_km} km`);
    if (filas.length === 0) {
      fallas++;
      console.log('!! No se encontraron precios. Líneas del texto con números grandes o "$"/"ARS":');
      texto.split('\n').filter(l => /ARS|\$|\d{2}[.,]\d{3}/.test(l)).slice(0, 40).forEach(l => console.log('   ', l.trim().slice(0, 160)));
      console.log('!! Primeras 80 líneas de la página:');
      texto.split('\n').map(l => l.trim()).filter(Boolean).slice(0, 80).forEach(l => console.log('   ', l.slice(0, 160)));
    }
    for (const f of filas) {
      console.log(`  ${f.producto}: ${f.precio}   [${f.contexto.slice(0, 150)}]`);
      nuevas.push([fecha, hora, ruta.id, f.producto, f.precio, 'ARS', resumen.tiempo_min, resumen.distancia_km]);
    }
  } catch (e) {
    fallas++;
    console.log(`!! Error: ${e.message}`);
  } finally {
    await page?.close();
  }
}
await browser.close();

if (RUTAS.some(r => !r.url)) {
  console.log('\n== Links a otras rutas encontrados (para completar rutas.json):');
  [...linksRutas].filter(h => /aeroparque|newbery|aep/i.test(h)).forEach(h => console.log('   ', h));
  console.log(`   (${linksRutas.size} links de rutas en total)`);
  [...linksRutas].slice(0, 60).forEach(h => console.log('    ·', h));
}

if (nuevas.length) {
  fs.appendFileSync(CSV, nuevas.map(f => f.map(csvCampo).join(',')).join('\n') + '\n');
  console.log(`\nSe agregaron ${nuevas.length} filas a ${CSV}`);
}
process.exit(fallas ? 1 : 0);
