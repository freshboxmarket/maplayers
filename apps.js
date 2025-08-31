(async function () {
  const qs = new URLSearchParams(location.search);
  const cfgUrl = qs.get('cfg') || './config/app.config.json';

  try {
    const cfg = await fetchJson(cfgUrl);

    // --- map & tiles ---
    const map = L.map('map', { preferCanvas: true }).setView([43.55, -80.25], 8);

    // grayscale base
    const base = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution:'&copy; OpenStreetMap contributors'
    }).addTo(map);
    const baseEl = base.getContainer?.();
    if (baseEl) { baseEl.style.filter = 'grayscale(1)'; baseEl.style.webkitFilter = 'grayscale(1)'; }

    const dayLayers = [];        // [{day, layer, perDay}]
    const boundaryFeatures = []; // for clipping color tiles
    let allBounds = null;        // union of all features
    let selectionBounds = null;  // union of selected features (last apply)
    let hasSelection = false;
    let totalFeatures = 0;

    // customers
    const customerLayer = L.layerGroup().addTo(map);
    let customerCount = 0;

    // at most one focused feature at a time
    let currentFocus = null;

    renderLegend(cfg, {}, customerCount);

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

          const p    = feat.properties || {};
          const key  = (p[cfg.fields.key]  ?? '').toString().toUpperCase().trim();
          const muni = smartTitleCase((p[cfg.fields.muni] ?? '').toString().trim());
          const day  = (p[cfg.fields.day]  ?? Lcfg.day).toString().trim();

          lyr._routeKey = key;
          lyr._day      = day;
          lyr._perDay   = perDay;
          lyr._labelTxt = muni;
          lyr._isSelected = false; // set during applySelection

          // click-to-focus (toggle)
          lyr.on('click', (e) => {
            L.DomEvent.stopPropagation(e);
            if (currentFocus === lyr) { clearFocus(true); }
            else { focusFeature(lyr); }
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

    // click background → clear focus and recenter to selection
    map.on('click', () => clearFocus(true));
    // move the map (pan/zoom) → clear focus (do NOT recenter)
    map.on('movestart', () => clearFocus(false));

    // initial selection
    await applySelection();

    // customers overlay (optional)
    await loadCustomersIfAny();

    const refresh = Number(cfg.behavior.refreshSeconds || 0);
    if (refresh > 0) setInterval(async () => {
      await applySelection();
      await loadCustomersIfAny();
    }, refresh * 1000);

    async function applySelection() {
      // selection change cancels a manual focus
      clearFocus(false);

      const selUrl = qs.get('sel') || cfg.selection.url;
      const csvText = await fetchText(selUrl);
      const rowsAA = parseCsvRows(csvText);
      const hdr = findHeader(rowsAA, cfg.selection.schema);

      const selectedSet = new Set();
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
            rawKeys.forEach(k => selectedSet.add(k.toUpperCase()));
          } else {
            const dayVal = ((r[dayIdx] || '') + '').trim();
            rawKeys.forEach(k => selectedSet.add(`${dayVal}||${k.toUpperCase()}`));
          }
        }
      }

      hasSelection = selectedSet.size > 0;
      selectionBounds = null;

      // counts & styling
      const counts = {};
      let selectedCount = 0;
      const selBounds = [];

      dayLayers.forEach(({ day, layer, perDay }) => {
        let dayTotal = 0, daySel = 0;

        layer.eachLayer(lyr => {
          dayTotal++;
          const hit = cfg.selection.mergeDays
            ? selectedSet.has(lyr._routeKey)
            : selectedSet.has(`${lyr._day}||${lyr._routeKey}`);

          lyr._isSelected = !!hit;

          if (hit) {
            daySel++;
            applyStyleSelected(lyr, perDay, cfg);
            showLabel(lyr, lyr._labelTxt);      // labels ON for selected
            if (lyr.getBounds) selBounds.push(lyr.getBounds());
            selectedCount++;
            layer.addLayer(lyr);
          } else {
            if (cfg.style.unselectedMode === 'hide' && hasSelection) {
              hideLabel(lyr);
              layer.removeLayer(lyr);
            } else {
              applyStyleDim(lyr, perDay, cfg);
              hideLabel(lyr);                    // labels OFF for unselected
              layer.addLayer(lyr);
            }
          }
        });

        counts[day] = { selected: daySel, total: dayTotal };
      });

      // center primarily on selected
      if (cfg.behavior.autoZoom) {
        if (hasSelection && selBounds.length) {
          selectionBounds = selBounds.reduce((acc, b) => acc.extend(b), L.latLngBounds(selBounds[0]));
          map.fitBounds(selectionBounds.pad(0.1));
        } else if (allBounds) {
          map.fitBounds(allBounds.pad(0.1));
        }
      }

      renderLegend(cfg, counts, customerCount);
      setStatus(`Features on map: ${totalFeatures} • Selected: ${selectedCount} • Customers: ${customerCount}${hasSelection ? '' : ' • (no keys parsed; showing all dimmed)'}`);
    }

    // ---------- customers overlay ----------
    async function loadCustomersIfAny() {
      customerLayer.clearLayers();
      customerCount = 0;

      // allow ?cust=off to suppress customers
      const custToggle = qs.get('cust');
      if (custToggle && custToggle.toLowerCase() === 'off') {
        renderLegend(cfg, null, customerCount);
        setStatus(`Customers: ${customerCount}`);
        return;
      }

      if (!cfg.customers || cfg.customers.enabled === false) {
        renderLegend(cfg, null, customerCount);
        setStatus(`Customers: ${customerCount}`);
        return;
      }

      const custUrl = cfg.customers.url;
      if (!custUrl) { renderLegend(cfg, null, customerCount); return; }

      const text = await fetchText(custUrl);
      const rows = parseCsvRows(text);
      if (!rows.length) { renderLegend(cfg, null, customerCount); return; }

      // find header row
      const hdrIdx = findCustomerHeaderIndex(rows, cfg.customers.schema);
      if (hdrIdx === -1) { renderLegend(cfg, null, customerCount); return; }

      const mapIdx = headerIndexMap(rows[hdrIdx], cfg.customers.schema);

      // build markers
      for (let i = hdrIdx + 1; i < rows.length; i++) {
        const r = rows[i]; if (!r || r.length === 0) continue;

        const coord = (mapIdx.coords !== -1) ? r[mapIdx.coords] : '';
        const note  = (mapIdx.note   !== -1) ? r[mapIdx.note]   : '';

        const ll = parseLatLng(coord);
        if (!ll) continue;

        const s = cfg.style.customers || {};
        const m = L.circleMarker([ll.lat, ll.lng], {
          radius: s.radius || 6,
          color:  s.stroke || '#111',
          weight: s.weightPx || 2,
          opacity: s.opacity != null ? s.opacity : 0.9,
          fillColor: s.fill || '#fff',
          fillOpacity: s.fillOpacity != null ? s.fillOpacity : 0.85
        }).addTo(customerLayer);

        const popup = note ? `<div style="max-width:240px">${escapeHtml(note)}</div>`
                           : `<div>${ll.lat.toFixed(6)}, ${ll.lng.toFixed(6)}</div>`;
        m.bindPopup(popup);

        customerCount++;
      }

      renderLegend(cfg, null, customerCount);
      // leave map extent unchanged; this overlay is passive
      setStatus(`Customers plotted: ${customerCount}`);
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

  // selection header finder
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

  // customers header helpers
  function findCustomerHeaderIndex(rows, schema) {
    const wantCoords = (schema.coords || 'Verified Coordinates').toLowerCase();
    const wantNote   = (schema.note   || 'Order Note').toLowerCase();
    for (let i=0;i<rows.length;i++) {
      const hdr = rows[i] || [];
      const hasCoords = hdr.some(h => (h||'').toLowerCase() === wantCoords);
      const hasNote   = hdr.some(h => (h||'').toLowerCase() === wantNote);
      if (hasCoords || hasNote) return i;
    }
    return -1;
  }
  function headerIndexMap(hdrRow, schema) {
    const wantCoords = (schema.coords || 'Verified Coordinates').toLowerCase();
    const wantNote   = (schema.note   || 'Order Note').toLowerCase();
    const idx = (name) => hdrRow.findIndex(h => (h||'').toLowerCase() === name);
    return {
      coords: idx(wantCoords),
      note:   idx(wantNote)
    };
  }
  function parseLatLng(s) {
    const m = String(s||'').trim().match(/(-?\d+(?:\.\d+)?)\s*[, ]\s*(-?\d+(?:\.\d+)?)/);
    if (!m) return null;
    const lat = parseFloat(m[1]), lng = parseFloat(m[2]);
    if (!isFinite(lat) || !isFinite(lng)) return null;
    if (lat < -90 || lat > 90 || lng < -180 || lng > 180) return null;
    return { lat, lng };
  }

  function dimFill(perDay, cfg) {
    const base = perDay.fillOpacity != null ? perDay.fillOpacity : 0.8;
    const factor = cfg.style.dimmed.fillFactor != null ? cfg.style.dimmed.fillFactor : 0.3;
    return Math.max(0.08, base * factor);
  }

  // label helpers: permanent tooltip on/off
  function showLabel(lyr, text) {
    if (lyr.getTooltip()) lyr.unbindTooltip();
    lyr.bindTooltip(text, { permanent: true, direction: 'center', className: 'lbl' });
  }
  function hideLabel(lyr) {
    const t = lyr.getTooltip();
    if (t) lyr.unbindTooltip();
  }

  // Municipality title case (handles slashes/hyphens and common particles)
  function smartTitleCase(s) {
    if (!s) return '';
    s = s.toLowerCase();
    const parts = s.split(/([\/-])/g);
    const small = new Set(['and','or','of','the','a','an','in','on','at','by','for','to','de','la','le','el','du','von','van','di','da','del']);
    const fixWord = (w, isFirst) => {
      if (!w) return w;
      if (!isFirst && small.has(w)) return w;
      if (w === 'st' || w === 'st.') return 'St.';
      if (w === 'mt' || w === 'mt.') return 'Mt.';
      return w.charAt(0).toUpperCase() + w.slice(1);
    };
    let tokenIdx = 0;
    for (let i=0;i<parts.length;i++){
      if (parts[i] === '/' || parts[i] === '-') continue;
      const tokens = parts[i].split(/\s+/).map(tok => fixWord(tok, tokenIdx++ === 0));
      parts[i] = tokens.join(' ');
    }
    return parts.map(p => p === '/' ? '/' : (p === '-' ? '-' : p)).join('').replace(/\s+/g,' ').trim();
  }

  function renderLegend(cfg, counts, customers) {
    const el = document.getElementById('legend');
    const rowsHtml = (cfg.layers || []).map(Lcfg => {
      const st = (cfg.style && cfg.style.perDay && cfg.style.perDay[Lcfg.day]) || {};
      const c  = (counts && counts[Lcfg.day]) ? counts[Lcfg.day] : { selected: 0, total: 0 };
      return `<div class="row">
        <span class="swatch" style="background:${st.fill}; border-color:${st.stroke}"></span>
        <div>${Lcfg.day}</div>
        <div class="counts">${c.selected}/${c.total}</div>
      </div>`;
    }).join('');
    const cust = (typeof customers === 'number')
      ? `<div class="row" style="margin-top:6px;border-top:1px solid #eee;padding-top:6px">
           <span class="swatch" style="background:#fff;border-color:#111"></span>
           <div>Customers</div>
           <div class="counts">${customers}</div>
         </div>`
      : '';
    el.innerHTML = `<h4>Layers</h4>${rowsHtml}${cust}<div style="margin-top:6px;opacity:.7">Unselected: ${cfg.style.unselectedMode}</div>`;
  }

  function setStatus(msg) { document.getElementById('status').textContent = msg; }
  function showError(e) {
    const el = document.getElementById('error');
    el.style.display = 'block';
    el.innerHTML = `<strong>Load error</strong><br>${e && e.message ? e.message : e}`;
  }
})();
