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
    const baseEl = base.getContainer?.();
    if (baseEl) { baseEl.style.filter = 'grayscale(1)'; baseEl.style.webkitFilter = 'grayscale(1)'; }

    // state
    const dayLayers = [];            // [{day, layer, perDay}]
    const boundaryFeatures = [];     // for clipping color tiles
    let allBounds = null;            // union of all features
    let selectionBounds = null;      // union of selected features (last apply)
    let totalFeatures = 0;
    let hasSelection = false;

    // focus state (one feature at a time) + its heavy outline overlay
    let currentFocus = null;
    let focusOutline = null;

    renderLegend(cfg, {});

    // --- load layers ---
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
          boundaryFeatures.push({ type: 'Feature', geometry: feat.geometry });

          const p   = feat.properties || {};
          const key = (p[cfg.fields.key]  ?? '').toString().toUpperCase().trim();
          const muniRaw = (p[cfg.fields.muni] ?? '').toString().trim();
          const muni = smartTitleCase(muniRaw);
          const day  = (p[cfg.fields.day]  ?? Lcfg.day).toString().trim();

          lyr._routeKey = key;
          lyr._day      = day;
          lyr._perDay   = perDay;
          lyr._labelTxt = muni;
          lyr._isSelected = false; // set in applySelection

          // click-to-focus: highlight border + show label; only one focused at a time
          lyr.on('click', (e) => {
            L.DomEvent.stopPropagation(e);
            focusFeature(lyr);
          });

          if (lyr.getBounds) {
            const b = lyr.getBounds();
            allBounds = allBounds ? allBounds.extend(b) : L.latLngBounds(b);
          }
        }
      }).addTo(map);

      dayLayers.push({ day: Lcfg.day, layer, perDay });
    }

    // color tiles clipped to union of delivery polygons
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
    } catch(e) { /* ok */ }

    // background click: clear focus + refit to selected extent
    map.on('click', () => {
      clearFocus();
      if (hasSelection && selectionBounds) map.fitBounds(selectionBounds.pad(0.1));
    });

    // any map movement: clear transient focus (so only selected labels remain)
    map.on('movestart', () => { clearFocus(); });

    await applySelection(); // centers on selection at launch

    const refresh = Number(cfg.behavior.refreshSeconds || 0);
    if (refresh > 0) setInterval(applySelection, refresh * 1000);

    // ---------- selection & styling ----------
    async function applySelection() {
      // changing selection cancels transient focus
      clearFocus();

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

      hasSelection = selected.size > 0;
      selectionBounds = null;

      // counts + apply styles/labels
      const counts = {};
      let selectedCount = 0;
      const selBounds = [];

      dayLayers.forEach(({ day, layer, perDay }) => {
        let dayTotal = 0, daySel = 0;

        layer.eachLayer(lyr => {
          dayTotal++;

          const hit = hasSelection && (
            cfg.selection.mergeDays
              ? selected.has(lyr._routeKey)
              : selected.has(`${lyr._day}||${lyr._routeKey}`)
          );

          lyr._isSelected = !!hit;

          if (hit) {
            applySelectedStyle(lyr, perDay, cfg);
            showLabel(lyr, lyr._labelTxt);
            if (lyr.getBounds) selBounds.push(lyr.getBounds());
            selectedCount++; daySel++;
          } else {
            applyDimmedStyle(lyr, perDay, cfg);
            // only keep label if this (now) unselected feature is the current focus
            if (currentFocus === lyr) showLabel(lyr, lyr._labelTxt);
            else hideLabel(lyr);
          }

          layer.addLayer(lyr);
        });

        counts[day] = { selected: daySel, total: dayTotal };
      });

      // center on selected primarily
      if (cfg.behavior.autoZoom) {
        if (hasSelection && selBounds.length) {
          selectionBounds = selBounds.reduce((acc, b) => acc.extend(b), L.latLngBounds(selBounds[0]));
          map.fitBounds(selectionBounds.pad(0.1));
        } else if (allBounds) {
          map.fitBounds(allBounds.pad(0.1));
        }
      }

      renderLegend(cfg, counts);
      setStatus(`Features on map: ${totalFeatures} • Selected: ${selectedCount}${hasSelection ? '' : ' • (no keys parsed; showing all dimmed)'}`);
    }

    // ---------- focus/highlight helpers ----------
    function focusFeature(lyr) {
      // clear previous focus (removes outline & hides its label if not selected)
      clearFocus();

      currentFocus = lyr;

      // heavy dashed outline behind the feature
      try {
        const geom = lyr.toGeoJSON();
        focusOutline = L.geoJSON(geom, {
          style: { color: '#000', weight: 6, opacity: 0.9, fillOpacity: 0, dashArray: '6 3' }
        }).addTo(map);
      } catch (e) { /* if toGeoJSON not available, skip outline */ }

      // bump main stroke a bit
      lyr.setStyle({ weight: (cfg.style.selected.weightPx || 2) + 2, opacity: 1 });
      showLabel(lyr, lyr._labelTxt);
      lyr.bringToFront?.();

      const b = lyr.getBounds?.();
      if (b) map.fitBounds(b.pad(0.2));
    }

    function clearFocus() {
      // remove outline if present
      if (focusOutline) { map.removeLayer(focusOutline); focusOutline = null; }

      if (currentFocus) {
        // restore style/label of previously focused feature
        const lyr = currentFocus;
        const perDay = lyr._perDay || {};
        if (lyr._isSelected) {
          applySelectedStyle(lyr, perDay, cfg);
          showLabel(lyr, lyr._labelTxt);  // selected labels stay
        } else {
          applyDimmedStyle(lyr, perDay, cfg);
          hideLabel(lyr);                  // unselected labels vanish
        }
        currentFocus = null;
      }
    }

  } catch (e) {
    showError(e);
    console.error(e);
  }

  // ---------- shared helpers ----------
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
      if (q) { if (c === '"') { if (i+1<N && text[i+1] === '"') { f+='"'; i+=2; } else { q=false; i++; } }
               else { f+=c; i++; } }
      else { if (c === '"') { q=true; i++; }
             else if (c === ',') { pushF(); i++; }
             else if (c === '\n') { pushF(); pushR(); i++; }
             else { f+=c; i++; } }
    }
    pushF(); pushR();
    return out.map(row => row.map(v => (v||'').replace(/^\uFEFF/, '').trim()));
  }

  // header row detection
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

  function applySelectedStyle(lyr, perDay, cfg) {
    lyr.setStyle({
      color: perDay.stroke || '#666',
      weight: cfg.style.selected.weightPx,
      opacity: cfg.style.selected.strokeOpacity,
      fillColor: perDay.fill || '#ccc',
      fillOpacity: perDay.fillOpacity != null ? perDay.fillOpacity : 0.8
    });
  }

  function applyDimmedStyle(lyr, perDay, cfg) {
    lyr.setStyle({
      color: perDay.stroke || '#666',
      weight: cfg.style.dimmed.weightPx,
      opacity: cfg.style.dimmed.strokeOpacity,
      fillColor: perDay.fill || '#ccc',
      fillOpacity: dimFill(perDay, cfg)
    });
  }

  function showLabel(lyr, text) {
    if (lyr.getTooltip()) lyr.unbindTooltip();
    lyr.bindTooltip(text, { permanent: true, direction: 'center', className: 'lbl' });
  }
  function hideLabel(lyr) {
    const t = lyr.getTooltip();
    if (t) lyr.unbindTooltip();
  }

  function smartTitleCase(s) {
    if (!s) return '';
    s = s.toLowerCase();
    const parts = s.split(/([\/-])/g);
    const small = new Set(['and','or','of','the','a','an','in','on','at','by','for','to','de','la','le','el','du','von','van','di','da','del']);
    const fixWord = (w, isFirst) => {
      if (!w) return w;
      if (!isFirst && small.has(w)) return w;
      if (w === 'st' || w === 'st.' ) return 'St.';
      if (w === 'mt' || w === 'mt.' ) return 'Mt.';
      return w.charAt(0).toUpperCase() + w.slice(1);
    };
    let tokenIdx = 0;
    for (let i=0;i<parts.length;i++){
      if (parts[i] === '/' || parts[i] === '-') continue;
      const tokens = parts[i].split(/\s+/).map((tok) => fixWord(tok, tokenIdx++ === 0));
      parts[i] = tokens.join(' ');
    }
    return parts.map(p => (p === '/' || p === '-') ? p : p).join('').replace(/\s+/g,' ').trim();
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
