(async function () {
  const qs = new URLSearchParams(location.search);
  const cfgUrl = qs.get('cfg') || './config/app.config.json';

  try {
    const cfg = await fetchJson(cfgUrl);

    // --- map & tiles ---
    const map = L.map('map', { preferCanvas: true }).setView([43.55, -80.25], 8);

    // grayscale base tiles
    const base = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution:'&copy; OpenStreetMap contributors'
    }).addTo(map);
    // make base grayscale
    const baseEl = base.getContainer?.();
    if (baseEl) { baseEl.style.filter = 'grayscale(1)'; baseEl.style.webkitFilter = 'grayscale(1)'; }

    // containers for layers & later boundary
    const dayLayers = [];
    const boundaryFeatures = []; // to clip the color tiles
    let allBounds = null;
    let totalFeatures = 0;

    // legend scaffold (we'll fill counts later)
    renderLegend(cfg, {});

    // --- load 4 day layers ---
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
          // collect boundary features for later tile clipping
          boundaryFeatures.push({ type: 'Feature', geometry: feat.geometry });

          const p = feat.properties || {};
          const key  = (p[cfg.fields.key]  ?? '').toString().toUpperCase().trim();
          const muni = (p[cfg.fields.muni] ?? '').toString().trim();
          const day  = (p[cfg.fields.day]  ?? Lcfg.day).toString().trim();

          lyr._routeKey = key;
          lyr._day      = day;
          lyr._perDay   = perDay;

          // pretty, permanent label: "s13 • PERTH EAST"
          const label = [key, muni].filter(Boolean).join(' • ');
          lyr.bindTooltip(label, { permanent: true, direction: 'center', className: 'lbl' });

          if (lyr.getBounds) {
            const b = lyr.getBounds();
            allBounds = allBounds ? allBounds.extend(b) : L.latLngBounds(b);
          }
        }
      }).addTo(map);

      dayLayers.push({ day: Lcfg.day, layer, perDay });
    }

    // --- color tiles clipped to union of delivery polygons ---
    try {
      if (L.TileLayer && L.TileLayer.boundaryCanvas && boundaryFeatures.length) {
        const boundaryFC = { type: 'FeatureCollection', features: boundaryFeatures };
        const colorTiles = L.TileLayer.boundaryCanvas('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
          boundary: boundaryFC,
          attribution: ''
        });
        colorTiles.addTo(map).setZIndex(2);
        base.setZIndex(1);
      }
    } catch(e) { /* fallback to grayscale-only base if plugin not present */ }

    // initial fit (in case selection is empty)
    if (cfg.behavior.autoZoom && allBounds) map.fitBounds(allBounds.pad(0.1));

    // --- selection & styling ---
    await applySelection();
    const refresh = Number(cfg.behavior.refreshSeconds || 0);
    if (refresh > 0) setInterval(applySelection, refresh * 1000);

    async function applySelection() {
      const selUrl = qs.get('sel') || cfg.selection.url;
      const csvText = await fetchText(selUrl);
      const rowsAA = parseCsvRows(csvText);
      const hdr = findHeader(rowsAA, cfg.selection.schema);

      const selected = new Set();
      if (hdr) {
        const { headerIndex, colMap } = hdr;
        const keysIdx = colMap.keys;
        const dayIdx  = colMap.day;

        for (let i = headerIndex + 1; i < rowsAA.length; i++) {
          const r = rowsAA[i];
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
      }

      // counts per day
      const counts = {}; // { day: { selected, total } }
      let selectedCount = 0;
      const selBounds = [];
      const hasSelection = selected.size > 0;

      dayLayers.forEach(({ day, layer, perDay }) => {
        let dayTotal = 0, daySel = 0;
        layer.eachLayer(lyr => {
          dayTotal++;
          const hit = cfg.selection.mergeDays
            ? selected.has(lyr._routeKey)
            : selected.has(`${lyr._day}||${lyr._routeKey}`);

          if (hasSelection && hit) {
            daySel++;
            lyr.setStyle({
              color: perDay.stroke || '#666',
              weight: cfg.style.selected.weightPx,
              opacity: cfg.style.selected.strokeOpacity,
              fillColor: perDay.fill || '#ccc',
              fillOpacity: perDay.fillOpacity != null ? perDay.fillOpacity : 0.8
            });
            // brighten label
            const tip = lyr.getTooltip();
            if (tip?.getElement()) tip.getElement().classList.remove('dim');

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
              // dim label
              const tip = lyr.getTooltip();
              if (tip?.getElement()) tip.getElement().classList.add('dim');

              layer.addLayer(lyr);
            }
          }
        });
        counts[day] = { selected: daySel, total: dayTotal };
      });

      if (cfg.behavior.autoZoom) {
        if (hasSelection && selBounds.length) {
          const allSel = selBounds.reduce((acc, b) => acc.extend(b), L.latLngBounds(selBounds[0]));
          map.fitBounds(allSel.pad(0.1));
        } else if (allBounds) {
          map.fitBounds(allBounds.pad(0.1));
        }
      }

      renderLegend(cfg, counts);
      setStatus(`Features on map: ${totalFeatures} • Selected: ${selectedCount} • Keys: ${hasSelection ? Array.from(selected).join(', ') : '—'}`);
    }

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
    return out.map(row => row.map(v => (v||'').replace(/^\uFEFF/,'').trim()));
  }

  // find header row that includes a keys column (and day if present)
  function findHeader(rows, schema) {
    const wantDay  = (schema.day  || 'day').toLowerCase();
    const wantKeys = (schema.keys || 'zone keys').toLowerCase();
    const isDayLike  = s => ['day','weekday',wantDay].includes((s||'').toLowerCase());
    const isKeysLike = s => ['zone keys','zone key','keys','route keys','selected keys',wantKeys].includes((s||'').toLowerCase());
    for (let i=0;i<rows.length;i++){
      const row = rows[i] || [];
      let dayIdx=-1, keysIdx=-1;
      row.forEach((h, idx) => {
        if (isDayLike(h)  && dayIdx  === -1) dayIdx  = idx;
        if (isKeysLike(h) && keysIdx === -1) keysIdx = idx;
      });
      if (keysIdx !== -1) return { headerIndex: i, colMap: { day: dayIdx, keys: keysIdx } };
    }
    return null;
  }

  function dimFill(perDay, cfg) {
    const base = perDay.fillOpacity != null ? perDay.fillOpacity : 0.8;
    const factor = cfg.style.dimmed.fillFactor != null ? cfg.style.dimmed.fillFactor : 0.3;
    return Math.max(0.08, base * factor);
  }

  function renderLegend(cfg, counts) {
    const el = document.getElementById('legend');
    const rowsHtml = cfg.layers.map(Lcfg => {
      const st = cfg.style.perDay[Lcfg.day] || {};
      const c  = (counts && counts[Lcfg.day]) ? counts[Lcfg.day] : { selected: 0, total: 0 };
      return `<div class="row">
        <span class="swatch" style="background:${st.fill}; border-color:${st.stroke}"></span>
        <div>${Lcfg.day}</div>
        <div class="counts">${c.selected}/${c.total}</div>
      </div>`;
    }).join('');
    el.innerHTML = `<h4>Layers</h4>${rowsHtml}<div style="margin-top:6px;opacity:.7">Unselected: ${cfg.style.unselectedMode}</div>`;
  }

  function setStatus(msg) { document.getElementById('status').textContent = msg; }
  function showError(e) {
    const el = document.getElementById('error');
    el.style.display = 'block';
    el.innerHTML = `<strong>Load error</strong><br>${e && e.message ? e.message : e}`;
  }
})();
