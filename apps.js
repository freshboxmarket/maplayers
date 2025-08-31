(async function () {
  const qs = new URLSearchParams(location.search);
  const cfgUrl = qs.get('cfg') || './config/app.config.json';

  try {
    // Load config
    const cfg = await fetchJson(cfgUrl);

    // Map
    const map = L.map('map', { preferCanvas: true }).setView([43.55, -80.25], 8);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution:'&copy; OpenStreetMap'
    }).addTo(map);

    renderLegend(cfg);

    // Load all day layers
    const dayLayers = [];
    for (const Lcfg of cfg.layers) {
      const gj = await fetchJson(Lcfg.url);
      const color = (cfg.style.perDay[Lcfg.day] && cfg.style.perDay[Lcfg.day].color) || '#666';

      const layer = L.geoJSON(gj, {
        style: () => ({
          color,
          weight: cfg.style.dimmed.weight,
          opacity: cfg.style.dimmed.opacity,
          fillOpacity: cfg.style.dimmed.opacity
        }),
        onEachFeature: (feat, lyr) => {
          const p = feat.properties || {};
          const key = (p[cfg.fields.key] ?? '').toString().toUpperCase().trim();
          const muni = p[cfg.fields.muni] ?? '';
          const day  = (p[cfg.fields.day] ?? Lcfg.day).toString().trim();
          lyr._routeKey  = key;      // used for selection matching
          lyr._day       = day;      // used if mergeDays=false
          lyr._baseColor = color;    // per-day color
          lyr.bindTooltip(`${day}: ${muni || 'area'}${key ? ` (${key})` : ''}`, { sticky: true });
        }
      }).addTo(map);

      dayLayers.push(layer);
    }

    // Selection logic
    async function applySelection() {
      const selUrl = qs.get('sel') || cfg.selection.url;
      const csvText = await fetchText(selUrl);
      const rows = parseCsv(csvText);

      const keyCol = cfg.selection.schema.keys;
      const dayCol = cfg.selection.schema.day;
      const delim  = cfg.selection.schema.delimiter || ',';

      // Build selected key set
      const selected = new Set();
      for (const r of rows) {
        const rawKeys = (r[keyCol] || '').split(delim).map(s => s.trim()).filter(Boolean);
        if (cfg.selection.mergeDays) {
          rawKeys.forEach(k => selected.add(k.toUpperCase()));
        } else {
          const day = (r[dayCol] || '').toString().trim();
          rawKeys.forEach(k => selected.add(`${day}||${k.toUpperCase()}`));
        }
      }

      // Style features
      const bounds = [];
      let selectedCount = 0;
      dayLayers.forEach(layer => {
        layer.eachLayer(lyr => {
          const hit = cfg.selection.mergeDays
            ? selected.has(lyr._routeKey)
            : selected.has(`${lyr._day}||${lyr._routeKey}`);

          if (hit) {
            lyr.setStyle({
              color: lyr._baseColor,
              weight: cfg.style.selected.weight,
              opacity: cfg.style.selected.opacity,
              fillOpacity: cfg.style.selected.opacity
            });
            if (lyr.getBounds) bounds.push(lyr.getBounds());
            selectedCount++;
            layer.addLayer(lyr);
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
              layer.addLayer(lyr);
            }
          }
        });
      });

      // Auto-zoom
      if (cfg.behavior.autoZoom && bounds.length) {
        const all = bounds.reduce((acc, b) => acc.extend(b), L.latLngBounds(bounds[0]));
        map.fitBounds(all.pad(0.1));
      }

      setStatus(`Selected features: ${selectedCount} • Keys: ${Array.from(selected).join(', ') || '—'}`);
    }

    await applySelection();

    // Optional auto-refresh of selection CSV
    const refresh = Number(cfg.behavior.refreshSeconds || 0);
    if (refresh > 0) {
      setInterval(applySelection, refresh * 1000);
    }

  } catch (e) {
    showError(e);
    console.error(e);
  }

  // ---------- Helpers ----------
  function cb(u) {
    const sep = u.includes('?') ? '&' : '?';
    return `${sep}cb=${Date.now()}`;
  }

  async function fetchJson(url) {
    const res = await fetch(url + cb(url));
    if (!res.ok) throw new Error(`Fetch failed (${res.status}) for ${url}`);
    return await res.json();
  }

  async function fetchText(url) {
    const res = await fetch(url + cb(url));
    if (!res.ok) throw new Error(`Fetch failed (${res.status}) for ${url}`);
    return await res.text();
  }

  function parseCsv(text) {
    // RFC4180-ish parser (handles quoted commas and double-quotes)
    const rows = [];
    let i = 0, field = '', row = [], inQ = false;
    const pushField = () => { row.push(field); field=''; };
    const pushRow = () => { rows.push(row); row=[]; };

    // Normalize newlines
    text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    const N = text.length;

    while (i < N) {
      const ch = text[i];
      if (inQ) {
        if (ch === '"') {
          if (i + 1 < N && text[i+1] === '"') { field += '"'; i += 2; continue; }
          inQ = false; i++; continue;
        }
        field += ch; i++; continue;
      } else {
        if (ch === '"') { inQ = true; i++; continue; }
        if (ch === ',') { pushField(); i++; continue; }
        if (ch === '\n') { pushField(); pushRow(); i++; continue; }
        field += ch; i++; continue;
      }
    }
    pushField(); pushRow();

    if (rows.length === 0) return [];
    const headers = rows.shift().map(h => h.replace(/^\uFEFF/, '').trim()); // strip BOM
    const out = rows
      .filter(r => r.some(v => (v || '').trim() !== ''))
      .map(r => {
        const o = {};
        headers.forEach((h, idx) => { o[h] = (r[idx] || '').trim(); });
        return o;
      });
    return out;
  }

  function renderLegend(cfg) {
    const el = document.getElementById('legend');
    const rows = Object.entries(cfg.style.perDay).map(([day, st]) =>
      `<div class="row"><span class="swatch" style="background:${st.color}"></span>${day}</div>`).join('');
    el.innerHTML = `<h4>Layers</h4>${rows}<div style="margin-top:6px;opacity:.7">Unselected: ${cfg.style.unselectedMode}</div>`;
  }

  function setStatus(msg) {
    const el = document.getElementById('status');
    el.textContent = msg;
  }

  function showError(e) {
    const el = document.getElementById('error');
    el.style.display = 'block';
    el.innerHTML = `<strong>Load error</strong><br>${e && e.message ? e.message : e}`;
  }
})();
