(async function () {
  const qs = new URLSearchParams(location.search);
  const cfgUrl = qs.get('cfg') || './config/app.config.json';

  try {
    const cfg = await fetchJson(cfgUrl);

    // Map
    const map = L.map('map', { preferCanvas: true }).setView([43.55, -80.25], 8);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution:'&copy; OpenStreetMap contributors'
    }).addTo(map);

    renderLegend(cfg);

    // Load day layers (keep overall bounds so we always show something)
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
          fillOpacity: dimFill(perDay, cfg)
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

          if (lyr.getBounds) {
            const b = lyr.getBounds();
            allBounds = allBounds ? allBounds.extend(b) : L.latLngBounds(b);
          }
        }
      }).addTo(map);

      dayLayers.push(layer);
    }

    // Selection
    await applySelection();
    const refresh = Number(cfg.behavior.refreshSeconds || 0);
    if (refresh > 0) setInterval(applySelection, refresh * 1000);

    async function applySelection() {
      const selUrl = qs.get('sel') || cfg.selection.url;
      const csvText = await fetchText(selUrl);

      // Parse CSV into rows (array of arrays), then locate the header row dynamically.
      const rows = parseCsvRows(csvText);
      const hdr = findHeader(rows, cfg.selection.schema);

      if (!hdr) {
        setStatus('No header detected (looked for "day" and "zone keys"). Showing all dimmed.');
        if (cfg.behavior.autoZoom && allBounds) map.fitBounds(allBounds.pad(0.1));
        return;
      }

      const { headerIndex, colMap } = hdr;
      const keysIdx = colMap.keys;
      const dayIdx  = colMap.day;

      // Build the selected key set
      const selected = new Set();
      for (let i = headerIndex + 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r || r.length === 0) continue;
        const rawKeys = ((r[keysIdx] || '') + '')
          .split(cfg.selection.schema.delimiter || ',')
          .map(s => s.trim())
          .filter(Boolean);
        if (cfg.selection.mergeDays) {
          rawKeys.forEach(k => selected.add(k.toUpperCase()));
        } else {
          const dayVal = ((r[dayIdx] || '') + '').trim();
          rawKeys.forEach(k => selected.add(`${dayVal}||${k.toUpperCase()}`));
        }
      }

      const hasSelection = selected.size > 0;

      let selectedCount = 0;
      const selBounds = [];

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
            if (lyr.getBounds) selBounds.push(lyr.getBounds());
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
                fillOpacity: dimFill(perDay, cfg)
              });
              layer.addLayer(lyr);
            }
          }
        });
      });

      if (cfg.behavior.autoZoom) {
        if (hasSelection && selBounds.length) {
          const allSel = selBounds.reduce((acc, b) => acc.extend(b), L.latLngBounds(selBounds[0]));
          map.fitBounds(allSel.pad(0.1));
        } else if (allBounds) {
          map.fitBounds(allBounds.pad(0.1));
        }
      }

      setStatus(`Features on map: ${totalFeatures} • Selected: ${selectedCount} • Keys: ${hasSelection ? Array.from(selected).join(', ') : '—'}`);
    }

  } catch (e) {
    showError(e);
    console.error(e);
  }

  // ------- helpers -------
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

  // CSV → array of arrays
  function parseCsvRows(text) {
    const out = []; let i=0, f='', r=[], q=false;
    text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    const N = text.length;
    const pushF=()=>{ r.push(f); f=''; }, pushR=()=>{ out.push(r); r=[]; };
    while (i<N) {
      const c = text[i];
      if (q) {
        if (c === '"') {
          if (i+1<N && text[i+1] === '"') { f+='"'; i+=2; } else { q=false; i++; }
        } else { f+=c; i++; }
      } else {
        if (c === '"') { q=true; i++; }
        else if (c === ',') { pushF(); i++; }
        else if (c === '\n') { pushF(); pushR(); i++; }
        else { f+=c; i++; }
      }
    }
    pushF(); pushR();
    // trim BOM/whitespace
    return out.map(row => row.map(v => (v||'').replace(/^\uFEFF/,'').trim()));
  }

  // Find a header row that contains both a day col and a keys col (tolerant of wording)
  function findHeader(rows, schema) {
    const wantDay  = (schema.day  || 'day').toLowerCase();
    const wantKeys = (schema.keys || 'zone keys').toLowerCase();

    const isDayLike  = s => {
      s = (s||'').toLowerCase();
      return s === wantDay || ['day','weekday'].includes(s);
    };
    const isKeysLike = s => {
      s = (s||'').toLowerCase();
      return s === wantKeys || ['zone keys','zone key','keys','route keys','selected keys'].includes(s);
    };

    for (let i=0; i<rows.length; i++) {
      const row = rows[i] || [];
      let dayIdx = -1, keysIdx = -1;
      row.forEach((h, idx) => {
        if (isDayLike(h)  && dayIdx  === -1) dayIdx  = idx;
        if (isKeysLike(h) && keysIdx === -1) keysIdx = idx;
      });
      if (keysIdx !== -1) {
        // day column optional when mergeDays=true; still record if present
        return { headerIndex: i, colMap: { day: dayIdx, keys: keysIdx } };
      }
    }
    return null;
  }

  function dimFill(perDay, cfg) {
    const base = perDay.fillOpacity != null ? perDay.fillOpacity : 0.8;
    const factor = cfg.style.dimmed.fillFactor != null ? cfg.style.dimmed.fillFactor : 0.3;
    return Math.max(0.08, base * factor);
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
