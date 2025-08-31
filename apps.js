(async function () {
  const qs = new URLSearchParams(location.search);
  const cfgUrl = qs.get('cfg') || './config/app.config.json';

  try {
    const cfg = await fetchJson(cfgUrl);

    // --- map & tiles ---
    const map = L.map('map', { preferCanvas: true }).setView([43.55, -80.25], 8);
    const base = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution:'&copy; OpenStreetMap contributors'
    }).addTo(map);
    const baseEl = base.getContainer?.();
    if (baseEl) { baseEl.style.filter = 'grayscale(1)'; baseEl.style.webkitFilter = 'grayscale(1)'; }

    const dayLayers = [];              // [{day, layer, perDay}]
    const boundaryFeatures = [];       // for clipping color tiles
    let allBounds = null;
    let selectionBounds = null;
    let hasSelection = false;
    let totalFeatures = 0;

    // Customers
    const customerLayer = L.layerGroup().addTo(map);
    const customerMarkers = [];        // [{ marker, lat, lng }]
    let customerCount = 0;
    let custWithinSel = 0;
    let custOutsideSel = 0;

    // one focused polygon at a time
    let currentFocus = null;

    // coverage sets for Turf
    const coveragePolysAll = [];       // all polygons (for recolor)
    let coveragePolysSelected = [];    // selected-only (for counts)

    renderLegend(cfg, {}, custWithinSel, custOutsideSel);

    // --- polygons ---
    for (const Lcfg of (cfg.layers || [])) {
      const gj = await fetchJson(Lcfg.url);
      const perDay = (cfg.style && cfg.style.perDay && cfg.style.perDay[Lcfg.day]) || {};
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
          lyr._turfFeat = { type: 'Feature', properties: { day }, geometry: feat.geometry };

          // keep a record for "color-by-any-polygon"
          coveragePolysAll.push({ feat: lyr._turfFeat, perDay });

          lyr.on('click', (e) => {
            L.DomEvent.stopPropagation(e);
            if (currentFocus === lyr) clearFocus(true);
            else focusFeature(lyr);
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
          boundary: boundaryFC, attribution: ''
        });
        colorTiles.addTo(map).setZIndex(2);
        base.setZIndex(1);
      }
    } catch(e) { /* ok */ }

    // clear focus interactions
    map.on('click', () => { clearFocus(true); });
    map.on('movestart', () => { clearFocus(false); map.closePopup(); });
    map.on('zoomstart', () => { map.closePopup(); });

    // initial selection + customers
    await applySelection();
    await loadCustomersIfAny();

    const refresh = Number(cfg.behavior.refreshSeconds || 0);
    if (refresh > 0) setInterval(async () => {
      await applySelection();
      await loadCustomersIfAny();
    }, refresh * 1000);

    // ---------- selection over polygons ----------
    async function applySelection() {
      clearFocus(false);

      const selUrl = qs.get('sel') || cfg.selection.url;
      const csvText = await fetchText(selUrl);
      const rowsAA = parseCsvRows(csvText);
      const hdr = findHeader(rowsAA, cfg.selection.schema);

      const selectedSet = new Set();
      if (hdr) {
        const { headerIndex, colMap } = hdr;
        const keysIdx = colMap.keys, dayIdx  = colMap.day;
        for (let i = headerIndex + 1; i < rowsAA.length; i++) {
          const r = rowsAA[i]; if (!r || r.length === 0) continue;
          const rawKeys = ((r[keysIdx] || '') + '')
            .split(cfg.selection.schema.delimiter || ',')
            .map(s => s.trim()).filter(Boolean);
          if (cfg.selection.mergeDays) rawKeys.forEach(k => selectedSet.add(k.toUpperCase()));
          else {
            const dayVal = ((r[dayIdx] || '') + '').trim();
            rawKeys.forEach(k => selectedSet.add(`${dayVal}||${k.toUpperCase()}`));
          }
        }
      }

      hasSelection = selectedSet.size > 0;
      selectionBounds = null;
      coveragePolysSelected = [];

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
            showLabel(lyr, lyr._labelTxt);
            coveragePolysSelected.push({ feat: lyr._turfFeat, perDay });
            if (lyr.getBounds) selBounds.push(lyr.getBounds());
            selectedCount++;
            layer.addLayer(lyr);
          } else {
            if (cfg.style.unselectedMode === 'hide' && hasSelection) {
              hideLabel(lyr); layer.removeLayer(lyr);
            } else {
              applyStyleDim(lyr, perDay, cfg);
              hideLabel(lyr); layer.addLayer(lyr);
            }
          }
        });
        counts[day] = { selected: daySel, total: dayTotal };
      });

      if (cfg.behavior.autoZoom) {
        if (hasSelection && selBounds.length) {
          selectionBounds = selBounds.reduce((acc, b) => acc.extend(b), L.latLngBounds(selBounds[0]));
          map.fitBounds(selectionBounds.pad(0.1));
        } else if (allBounds) {
          map.fitBounds(allBounds.pad(0.1));
        }
      }

      // recolor + recount customers relative to the NEW selection
      recolorAndRecountCustomers();

      renderLegend(cfg, counts, custWithinSel, custOutsideSel);
      setStatus(`Features: ${totalFeatures} • Selected: ${selectedCount} • Customers (in/out): ${custWithinSel}/${custOutsideSel}`);
    }

    // ---------- customers overlay ----------
    async function loadCustomersIfAny() {
      customerLayer.clearLayers();
      customerMarkers.length = 0;
      customerCount = 0;

      const custToggle = qs.get('cust');
      if (custToggle && custToggle.toLowerCase() === 'off') {
        custWithinSel = 0; custOutsideSel = 0;
        renderLegend(cfg, null, custWithinSel, custOutsideSel);
        return;
      }
      if (!cfg.customers || cfg.customers.enabled === false || !cfg.customers.url) {
        custWithinSel = 0; custOutsideSel = 0;
        renderLegend(cfg, null, custWithinSel, custOutsideSel);
        return;
      }

      const text = await fetchText(cfg.customers.url);
      const rows = parseCsvRows(text);
      if (!rows.length) {
        custWithinSel = 0; custOutsideSel = 0;
        renderLegend(cfg, null, custWithinSel, custOutsideSel);
        return;
      }

      const hdrIdx = findCustomerHeaderIndex(rows, cfg.customers.schema);
      if (hdrIdx === -1) {
        custWithinSel = 0; custOutsideSel = 0;
        renderLegend(cfg, null, custWithinSel, custOutsideSel);
        return;
      }
      const mapIdx = headerIndexMap(rows[hdrIdx], cfg.customers.schema);

      // default style
      const s = cfg.style.customers || {};
      const baseStyle = {
        radius: s.radius || 9,
        color:  s.stroke || '#111',
        weight: s.weightPx || 2,
        opacity: s.opacity != null ? s.opacity : 0.95,
        fillColor: s.fill || '#ffffff',
        fillOpacity: s.fillOpacity != null ? s.fillOpacity : 0.95
      };

      // markers
      for (let i = hdrIdx + 1; i < rows.length; i++) {
        const r = rows[i]; if (!r || r.length === 0) continue;
        const coord = (mapIdx.coords !== -1) ? r[mapIdx.coords] : '';
        const note  = (mapIdx.note   !== -1) ? r[mapIdx.note]   : '';
        const ll = parseLatLng(coord);
        if (!ll) continue;

        const m = L.circleMarker([ll.lat, ll.lng], baseStyle).addTo(customerLayer);
        const popupHtml = note
          ? `<div style="max-width:260px">${escapeHtml(note)}</div>`
          : `<div>${ll.lat.toFixed(6)}, ${ll.lng.toFixed(6)}</div>`;
        m.bindPopup(popupHtml, { autoClose: true, closeOnClick: true });

        customerMarkers.push({ marker: m, lat: ll.lat, lng: ll.lng });
        customerCount++;
      }

      // color + count
      recolorAndRecountCustomers();

      // keep extent unchanged
      renderLegend(cfg, null, custWithinSel, custOutsideSel);
      setStatus(`Customers (in/out): ${custWithinSel}/${custOutsideSel}`);
    }

    // recolor (by ANY polygon) + recount (by SELECTED polygons)
    function recolorAndRecountCustomers() {
      let inSel = 0, outSel = 0;
      const turfOn = (typeof turf !== 'undefined');

      for (const rec of customerMarkers) {
        // base style (white)
        const s = cfg.style.customers || {};
        const style = {
          radius: s.radius || 9,
          color:  s.stroke || '#111',
          weight: s.weightPx || 2,
          opacity: s.opacity != null ? s.opacity : 0.95,
          fillColor: s.fill || '#ffffff',
          fillOpacity: s.fillOpacity != null ? s.fillOpacity : 0.95
        };

        if (turfOn) {
          const pt = turf.point([rec.lng, rec.lat]);

          // color by any polygon hit (first match wins)
          for (let j = 0; j < coveragePolysAll.length; j++) {
            if (turf.booleanPointInPolygon(pt, coveragePolysAll[j].feat)) {
              const pd = coveragePolysAll[j].perDay || {};
              style.color = pd.stroke || style.color;
              style.fillColor = pd.fill || style.fillColor;
              break;
            }
          }

          // count by selected polygons
          let insideSel = false;
          if (hasSelection && coveragePolysSelected.length) {
            for (let k = 0; k < coveragePolysSelected.length; k++) {
              if (turf.booleanPointInPolygon(pt, coveragePolysSelected[k].feat)) { insideSel = true; break; }
            }
          }
          if (insideSel) inSel++; else outSel++;
        } else {
          // no Turf → can't classify; count all as "outside"
          outSel++;
        }

        rec.marker.setStyle(style);
      }

      custWithinSel = inSel;
      custOutsideSel = outSel;
    }

    // ---------- polygon focus/highlight ----------
    function focusFeature(lyr) {
      if (currentFocus && currentFocus !== lyr) restoreFeature(currentFocus);
      currentFocus = lyr;

      const perDay = lyr._perDay || {};
      const baseWeight = lyr._isSelected ? cfg.style.selected.weightPx : cfg.style.dimmed.weightPx;
      const hiWeight = Math.max(3, Math.round(baseWeight * 3));

      lyr.setStyle({
        color: perDay.stroke || '#666',
        weight: hiWeight,
        opacity: 1.0,
        fillColor: perDay.fill || '#ccc',
        fillOpacity: lyr._isSelected
          ? (perDay.fillOpacity != null ? perDay.fillOpacity : 0.8)
          : dimFill(perDay, cfg)
      });

      showLabel(lyr, lyr._labelTxt);
      lyr.bringToFront?.();

      const b = lyr.getBounds?.();
      if (b) map.fitBounds(b.pad(0.2));
    }
    function restoreFeature(lyr) {
      const perDay = lyr._perDay || {};
      if (lyr._isSelected) {
        applyStyleSelected(lyr, perDay, cfg); showLabel(lyr, lyr._labelTxt);
      } else {
        applyStyleDim(lyr, perDay, cfg); hideLabel(lyr);
      }
    }
    function clearFocus(recenter) {
      if (!currentFocus) {
        if (recenter && hasSelection && selectionBounds) map.fitBounds(selectionBounds.pad(0.1));
        return;
      }
      restoreFeature(currentFocus);
      currentFocus = null;
      if (recenter) {
        if (hasSelection && selectionBounds) map.fitBounds(selectionBounds.pad(0.1));
        else if (allBounds) map.fitBounds(allBounds.pad(0.1));
      }
    }

    // ---------- polygon style helpers ----------
    function applyStyleSelected(lyr, perDay, cfg) {
      lyr.setStyle({
        color: perDay.stroke || '#666',
        weight: cfg.style.selected.weightPx,
        opacity: cfg.style.selected.strokeOpacity,
        fillColor: perDay.fill || '#ccc',
        fillOpacity: perDay.fillOpacity != null ? perDay.fillOpacity : 0.8
      });
    }
    function applyStyleDim(lyr, perDay, cfg) {
      lyr.setStyle({
        color: perDay.stroke || '#666',
        weight: cfg.style.dimmed.weightPx,
        opacity: cfg.style.dimmed.strokeOpacity,
        fillColor: perDay.fill || '#ccc',
        fillOpacity: dimFill(perDay, cfg)
      });
    }

  } catch (e) {
    showError(e);
    console.error(e);
  }

  // ---------- shared helpers ----------
  function cb(u) { return (u.includes('?') ? '&' : '?') + 'cb=' + Date.now(); }
  async function fetchJson(url) { const res = await fetch(url + cb(url)); if (!res.ok) throw new Error(`Fetch ${res.status} for ${url}`); return res.json(); }
  async function fetchText(url) { const res = await fetch(url + cb(url)); if (!res.ok) throw new Error(`Fetch ${res.status} for ${url}`); return res.text(); }

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
    return { coords: idx(wantCoords), note: idx(wantNote) };
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

  // labels
  function showLabel(lyr, text) { if (lyr.getTooltip()) lyr.unbindTooltip(); lyr.bindTooltip(text, { permanent: true, direction: 'center', className: 'lbl' }); }
  function hideLabel(lyr) { if (lyr.getTooltip()) lyr.unbindTooltip(); }

  // escapes for popup HTML
  function escapeHtml(s) {
    return String(s || '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  // Municipality title case
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

  // legend
  function renderLegend(cfg, counts, custIn, custOut) {
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
    const custBlock = `<div class="row" style="margin-top:6px;border-top:1px solid #eee;padding-top:6px">
        <div>Customers within selection:</div><div class="counts">${custIn ?? 0}</div>
      </div>
      <div class="row">
        <div>Customers outside selection:</div><div class="counts">${custOut ?? 0}</div>
      </div>`;
    el.innerHTML = `<h4>Layers</h4>${rowsHtml}${custBlock}`;
  }

  function setStatus(msg) { const n = document.getElementById('status'); if (n) n.textContent = msg || ''; }
  function showError(e) {
    const el = document.getElementById('error');
    el.style.display = 'block';
    el.innerHTML = `<strong>Load error</strong><br>${e && e.message ? e.message : e}`;
  }
})();
