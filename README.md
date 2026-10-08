# precios-uber-aeropuertos

Registro diario de los precios que publica Uber en sus páginas de rutas para:

- **CABA → Ezeiza (EZE)**
- **CABA → Aeroparque (AEP)**

Todos los días a las 09:07 (hora argentina) un GitHub Action abre esas páginas, lee los precios por producto (UberX, Comfort, etc.) y los agrega a [`data/precios.csv`](data/precios.csv).

## Datos

| columna | descripción |
|---|---|
| `fecha` | fecha de la lectura (hora argentina) |
| `hora_ar` | hora de la lectura |
| `ruta` | `CABA-EZE` o `CABA-AEP` |
| `producto` | producto de Uber según la página |
| `precio_min`, `precio_max` | rango publicado, en pesos |
| `moneda` | `ARS` |
| `texto` | texto original de donde se sacó el precio |

Son los **valores de referencia que Uber muestra en la página pública**, no la cotización en vivo de la app (que varía por hora y demanda).

## Uso

- Correrlo a mano: pestaña **Actions → Precios diarios → Run workflow**.
- Agregar rutas: sumar una entrada a `rutas.json` con la URL de la página de Uber.
- Si una corrida falla, en el run quedan las capturas de la página (artefacto `debug`) para ver qué cambió.

Local: `npm install && npx playwright install chromium && node scrape.mjs`
