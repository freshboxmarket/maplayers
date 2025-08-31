(async function () {
  const params = new URLSearchParams(location.search);
  const cfgUrl = params.get('cfg') || './config/app.config.json';
  const cfg = await (await fetch(cfgUrl + cb())).json();

  const map = L.map('map', { preferCanvas: true }).setView([43.55, -80.25], 8);
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', { attribution:'&copy; OpenStreetMap' }).addTo(map);

  // Legend
  renderLegend(cfg);

  // Load selection (CSV → Set of route keys, uppercase/trim)
  async function loadSelection() {
    const selUrl = params.get('sel') || cfg.selection.url;
    const txt = await (await fetch(selUrl + cb())).text();
    const rows = parseCsv(txt);  // array of objects keyed by header
    const keyCol = cfg.selection.schema.keys;
    const dayCol = cfg.selection.schema.day;
    const delim  = cfg.selection.schema.delimiter || ',';

    const keys = new Set();
    for (const r of rows) {
      const rawKeys = (r[keyCol] || '').split(delim).map(s => s.trim()).filter(Boolean);
      if (!cfg.selection.mergeDays) {
        const day = (r[dayCol] || '').toString().trim();
        rawKeys.forEach(k => keys.add(`${day}||${k.toUpperCase()}`));
      } else {
        rawKeys.forEach(k => keys.add(k.toUpperCase()));
      }
    }
    return keys;
  }

  // Load day layers
  const layers = [];
  for (const Lcfg of cfg.layers) {
    const gj = await (await fetch(Lcfg.url + cb())).json();
    const color = (cfg.style.perDay[Lcfg.day] && cfg.style.perDay[Lcfg.day].color) || '#666';

    const layer = L.geoJSON(gj, {
      style: () => ({
        color,
        weight: cfg.style.dimmed.weight,
        opacity: cfg.style.dimmed.opacity,
        fillOpacity: cfg.style.dimmed.opacity
      }),
      onEachFeature: (feat, lyr) => {
        const props = feat.properties || {};
        const key = props[cfg.fields.key];
        const muni = props[cfg.fields.muni];
        lyr.bindTooltip(`${Lcfg.day}: ${muni || 'area'} (${key || '-'})`, { sticky: true });
        // stash for later re-styling
        lyr._routeKey = key ? key.toString().toUpperCase() : '';
        lyr._day = props[cfg.fields.day] || Lcfg.day;
        lyr._baseColor = color;
      }
    }).addTo(map);
    layers.push(layer);
  }

  // Apply selection styling
  async function applySelection() {
    const selected = await loadSelection();
    const bounds = [];
    layers.forEach(layer => {
      layer.eachLayer(lyr => {
        const keyHit = cfg.selection.mergeDays
          ? selected.has(lyr._routeKey)
          : selected.has(`${lyr._day}||${lyr._routeKey}`);

        if (keyHit) {
          lyr.setStyle({
            color: lyr._baseColor,
            weight: cfg.style.selected.weight,
            opacity: cfg.style.selected.opacity,
            fillOpacity: cfg.style.selected.opacity
          });
          if (lyr.getBounds) bounds.push(lyr.getBounds());
          lyr.addTo(layer);
        } else {
          if (cfg.style.unselectedMode === 'hide') {
            layer.removeLayer(lyr);
          } else {
            lyr.setStyle({
              color: lyr._baseColor,
              weight: cfg.style.dimmed.weight,
              opacity: cfg.style.dimmed.opacity,
              fillOpacity: cfg.style.dimmed.opacity
            });
            lyr.addTo(layer);
          }
        }
      });
    });

    if (cfg.behavior.autoZoom && bounds.length) {
      const all = bounds.reduce((acc, b) => acc.extend(b), L.latLngBounds(bounds[0]));
      map.fitBounds(all.pad(0.1));
    }
  }

  await applySelection();

  // Optional auto-refresh of selection CSV
  if (cfg.behavior.refreshSeconds > 0) {
    setInterval(applySelection, cfg.behavior.refreshSeconds * 1000);
  }

  // Helpers
  function parseCsv(text) {
    const rows = text.trim().split(/\r?\n/);
    const headers = rows.shift().split(',').map(h => h.trim().replace(/^\ufeff/, '')); // strip BOM
    return rows.map(line => {
      const cols = splitCsvLine(line);
      const obj = {};
      headers.forEach((h, i) => obj[h] = (cols[i] || '').trim());
      return obj;
    });
  }

  function splitCsvLine(line) {
    // simple CSV splitter (no nested quotes needed for your sheet)
    const out = []; let cur = ''; let inQ = false;
    for (let i=0; i<line.length; i++) {
      const ch = line[i];
      if (ch === '"' ) { inQ = !inQ; continue; }
      if (ch === ',' && !inQ) { out.push(cur); cur=''; continue; }
      cur += ch;
    }
    out.push(cur);
    return out;
  }

  function cb() {
    const n = Date.now();
    return (typeof URL === 'function') ? `?cb=${n}` : '';
  }

  function renderLegend(cfg) {
    const el = document.getElementById('legend');
    const rows = Object.entries(cfg.style.perDay).map(([day, st]) =>
      `<div class="row"><span class="swatch" style="background:${st.color}"></span>${day}</div>`).join('');
    el.innerHTML = `<h4>Zones</h4>${rows}<small>Selected = bright • Others = ${cfg.style.unselectedMode}</small>`;
  }
})();
