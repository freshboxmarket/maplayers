(async function () {
  const qs = new URLSearchParams(location.search);
  const cfgUrl = qs.get('cfg') || './config/app.config.json';

  try {
    const cfg = await fetchJson(cfgUrl);

    // Base map
    const map = L.map('map', { preferCanvas: true }).setView([43.55, -80.25], 8);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution:'&copy; OpenStreetMap contributors'
    }).addTo(map);

    renderLegend(cfg);

    // Load all day layers; keep a union of all bounds so we can always fit something
    const dayLayers = [];
    let allBounds = null;
    let totalFeatures = 0;

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
          fillOpacity: Math.max(0.08, (perDay.fillOpacity ?? 0.8) * (cfg.style.dimmed.fillFactor ?? 0.3))
        }),
        onEachFeature: (feat, lyr) => {
          totalFeatures++;
          const p = feat.properties || {};
          const key = (p[cfg.fields.key] ?? '').toString().toUpperCase().trim();
          const muni = p[cfg.fields.muni] ?? '';
          const day  = (p[cfg.fields.day] ?? Lcfg.day).toString().trim();

          lyr._routeKey  = key;
          lyr._day       = day;
          lyr._perDay    = perDay;

          lyr.bindTooltip(`${day}: ${muni || 'area'}${key ? ` (${key})` : ''}`, { sticky: true });

          // grow union bounds while we iterate
          if (lyr.getBounds) {
            const b = lyr.getBounds();
            allBounds = allBounds ? allBounds.extend(b) : L.latLngBounds(b);
          }
        }
      }).addTo(map);

      dayLayers.push(layer);
    }

    // Selection application
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

      // If no keys parsed, don't hide everything — just keep everything dim and fit to all
      const hasSelection = selected.size > 0;

      let selectedCount = 0;
      const selectedBounds = [];

      dayLayers.forEach(layer => {
        layer.eachLayer(lyr => {
          const hit = cfg.selection.mergeDays
            ? selected.has(lyr._routeKey)
            : selected.has(`${lyr._day}||${lyr._routeKey}`);

          const perDay = lyr._perDay || {};
          if (hasSelection && hit) {
            lyr.setStyle({
              color: perDay.stroke || '#666',
              weight: cfg.style.selected.weightPx,
              opacity: cfg.style.selected.strokeOpacity,
              fillColor: perDay.fill || '#ccc',
              fillOpacity: perDay.fillOpacity != null ? perDay.fillOpacity : 0.8
            });
            if (lyr.getBounds) selectedBounds.push(lyr.getBounds());
            selectedCount++;
            layer.addLayer(lyr);
          } else {
            if (cfg.style.unselectedMode === 'hide' && hasSelection) {
              layer.removeLayer(lyr);
            } else {
              lyr.setStyle({
                color: perDay.stroke || '#666',
                weight: cfg.style.dimmed.weightPx,
                opacity: cfg.style.dimmed.strokeOpacity,
                fillColor: perDay.fill || '#ccc',
                fillOpacity: Math.max(0.08, (perDay.fillOpacity ?? 0.8) * (cfg.style.dimmed.fillFactor ?? 0.3))
              });
              layer.addLayer(lyr);
            }
          }
        });
      });

      if (cfg.behavior.autoZoom) {
        if (hasSelection && selectedBounds.length) {
          const allSel = selectedBounds.reduce((acc, b) => acc.extend(b), L.latLngBounds(selectedBounds[0]));
          map.fitBounds(allSel.pad(0.1));
        } else if (allBounds) {
          map.fitBounds(allBounds.pad(0.1));
        }
      }

      setStatus(`Features on map: ${totalFeatures} • Selected: ${selectedCount} • Keys: ${hasSelection ? Array.from(selected).join(', ') : '—'}`);
    }

    await applySelection();

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
    // robust CSV with quotes
    const rows = []; let i=0, f='', r=[], q=false;
    text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    const N = text.length;
    const pushF=()=>{ r.push(f); f=''; }, pushR=()=>{ rows.push(r); r=[]; };
    while (i<N) {
      const c = text[i];
      if (q) { if (c === '"') { if (i+1<N && text[i+1] === '"') { f+='"'; i+=2; } else { q=false; i++; } }
               else { f+=c; i++; } }
      else { if (c === '"') { q=true; i++; }
             else if (c === ',') { pushF(); i++; }
             else if (c === '\n') { pushF(); pushR(); i++; }
             else { f+=c; i++; } }
    }
    pushF(); pushR();
    if (!rows.length) return [];
    const headers = rows.shift().map(h => h.replace(/^\uFEFF/, '').trim());
    return rows
      .filter(row => row.some(v => (v||'').trim() !== ''))
      .map(row => {
        const o = {};
        headers.forEach((h, idx) => o[h] = (row[idx] || '').trim());
        return o;
      });
  }
  function renderLegend(cfg) {
    const el = document.getElementById('legend');
    const rows = Object.entries(cfg.style.perDay).map(([day, st]) =>
      `<div class="row">
         <span class="swatch" style="background:${st.fill}; border-color:${st.stroke}"></span>${day}
       </div>`).join('');
    el.innerHTML = `<h4>Layers</h4>${rows}<div style="margin-top:6px;opacity:.7">Unselected: ${cfg.style.unselectedMode}</div>`;
  }
  function setStatus(msg) { document.getElementById('status').textContent = msg; }
  function showError(e) {
    const el = document.getElementById('error');
    el.style.display = 'block';
    el.innerHTML = `<strong>Load error</strong><br>${e && e.message ? e.message : e}`;
  }
})();
