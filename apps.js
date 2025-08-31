(async function () {
  const qs = new URLSearchParams(location.search);
  const cfgUrl = qs.get('cfg') || './config/app.config.json';

  try {
    // Load config
    const cfg = await fetchJson(cfgUrl);

    // Map
    const map = L.map('map', { preferCanvas: true }).setView([43.55, -80.25], 8);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution:'&copy; OpenStreetMap contributors'
    }).addTo(map);

    renderLegend(cfg);

    // Load day layers
    const dayLayers = [];
    for (const Lcfg of cfg.layers) {
      const gj = await fetchJson(Lcfg.url);
      const perDay = cfg.style.perDay[Lcfg.day] || {};
      const color = perDay.stroke || '#666';
      const fillColor = perDay.fill || '#ccc';

      const layer = L.geoJSON(gj, {
        style: () => ({
          color,
          weight: cfg.style.dimmed.weightPx,
          opacity: cfg.style.dimmed.strokeOpacity,
          fillColor,
          fillOpacity: cfg.style.dimmed.fillOpacity
        }),
        onEachFeature: (feat, lyr) => {
          const p = feat.properties || {};
          const key = (p[cfg.fields.key] ?? '').toString().toUpperCase().trim();
          const muni = p[cfg.fields.muni] ?? '';
          const day  = (p[cfg.fields.day] ?? Lcfg.day).toString().trim();

          // Stash for later
          lyr._routeKey  = key;
          lyr._day       = day;
          lyr._perDay    = perDay;

          lyr.bindTooltip(`${day}: ${muni || 'area'}${key ? ` (${key})` : ''}`, { sticky: true });
        }
      }).addTo(map);

      dayLayers.push(layer);
    }

    // Apply selection
    async function applySelection() {
      const selUrl = qs.get('sel') || cfg.selection.url;
      const csvText = await fetchText(selUrl);
      const rows = parseCsv(csvText);

      const keyCol = cfg.selection.schema.keys;
      const dayCol = cfg.selection.schema.day;
      const delim  = cfg.selection.schema.delimiter || ',';

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

      // Style each feature
      const bounds = [];
      let selectedCount = 0;

      dayLayers.forEach(layer => {
        layer.eachLayer(lyr => {
          const hit = cfg.selection.mergeDays
            ? selected.has(lyr._routeKey)
            : selected.has(`${lyr._day}||${lyr._routeKey}`);

          const perDay = lyr._perDay || {};
          if (hit) {
            lyr.setStyle({
              color: perDay.stroke || '#666',
              weight: cfg.style.selected.weightPx,
              opacity: cfg.style.selected.strokeOpacity,
              fillColor: perDay.fill || '#ccc',
              fillOpacity: perDay.fillOpacity != null ? perDay.fillOpacity : 0.8
            });
            if (lyr.getBounds) bounds.push(lyr.getBounds());
            selectedCount++;
            layer.addLayer(lyr);
          } else {
            if (cfg.style.unselectedMode === 'hide') {
              layer.removeLayer(lyr);
            } else {
              lyr.setStyle({
                color: perDay.stroke || '#666',
                weight: cfg.style.dimmed.weightPx,
                opacity: cfg.style.dimmed.strokeOpacity,
                fillColor: perDay.fill || '#ccc',
                fillOpacity: cfg.style.dimmed.fillOpacity
              });
              layer.addLayer(lyr);
            }
          }
        });
      });

      // Zoom to selection
      if (cfg.behavior.autoZoom && bounds.length) {
        const all = bounds.reduce((acc, b) => acc.extend(b), L.latLngBounds(bounds[0]));
        map.fitBounds(all.pad(0.1));
      }

      setStatus(`Selected features: ${selectedCount}  •  Keys: ${Array.from(selected).join(', ') || '—'}`);
    }

    await applySelection();

    // Optional auto-refresh
    const refresh = Number(cfg.behavior.refreshSeconds || 0);
    if (refresh > 0) setInterval(applySelection, refresh * 1000);

  } catch (e) {
    showError(e);
    console.error(e);
  }

  // ---------- helpers ----------
  function cb(u) { return (u.includes('?') ? '&' : '?') + 'cb=' + Date.now(); }

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
    // Robust CSV (quotes + commas)
    const out = [];
    let i=0, f='', r=[], q=false;
    text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    const N = text.length;
    const pushF = ()=>{ r.push(f); f=''; };
    const pushR = ()=>{ out.push(r); r=[]; };
    while (i < N) {
      const c = text[i];
      if (q) {
        if (c === '"') {
          if (i+1<N && text[i+1] === '"') { f += '"'; i+=2; continue; }
          q = false; i++; continue;
        }
        f += c; i++; continue;
      } else {
        if (c === '"') { q = true; i++; continue; }
        if (c === ',') { pushF(); i++; continue; }
        if (c === '\n') { pushF(); pushR(); i++; continue; }
        f += c; i++; continue;
      }
    }
    pushF(); pushR();
    if (!out.length) return [];
    const headers = out.shift().map(h => h.replace(/^\uFEFF/,'').trim());
    return out
      .filter(row => row.some(v => (v||'').trim() !== ''))
      .map(row => {
        const obj = {};
        headers.forEach((h, idx) => obj[h] = (row[idx] || '').trim());
        return obj;
      });
  }

  function renderLegend(cfg) {
    const el = document.getElementById('legend');
    const rows = Object.entries(cfg.style.perDay).map(([day, st]) =>
      `<div class="row">
         <span class="swatch" style="background:${st.fill}; border-color:${st.stroke}"></span>${day}
       </div>`
    ).join('');
    el.innerHTML = `<h4>Layers</h4>${rows}<div style="margin-top:6px;opacity:.7">Unselected: ${cfg.style.unselectedMode}</div>`;
  }

  function setStatus(msg) {
    document.getElementById('status').textContent = msg;
  }
  function showError(e) {
    const el = document.getElementById('error');
    el.style.display = 'block';
    el.innerHTML = `<strong>Load error</strong><br>${e && e.message ? e.message : e}`;
  }
})();
