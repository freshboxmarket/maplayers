(async function () {
  const qs = new URLSearchParams(location.search);
  const cfgUrl = qs.get('cfg') || './config/app.config.json';

  try {
    phase('Loading config…');
    const cfg = await fetchJson(cfgUrl);

    // --- map & tiles ---
    phase('Initializing map…');
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

    // per-day customers inside selection (legend fractions)
    const custByDayInSel = { Wednesday: 0, Thursday: 0, Friday: 0, Saturday: 0 };

    // one focused polygon at a time
    let currentFocus = null;

    // coverage sets for Turf
    const coveragePolysAll = [];       // all polygons (for _custAny & dimming)
    let coveragePolysSelected = [];    // only selected polygons

    // selected municipality list (status)
    let selectedMunicipalities = [];

    renderLegend(cfg, null, custWithinSel, custOutsideSel);

    // --- polygons ---
    phase('Loading polygon layers…');
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
          const rawKey  = (p[cfg.fields.key]  ?? '').toString().trim();
          const keyNorm = normalizeKey(rawKey);
          const muni = smartTitleCase((p[cfg.fields.muni] ?? '').toString().trim());
          const day  = Lcfg.day; // use layer day

          lyr._routeKey = keyNorm;       // normalized like W1/T6/F12/S2
          lyr._day      = day;
          lyr._perDay   = perDay;
          lyr._labelTxt = muni;
          lyr._isSelected = false;
          lyr._custAny = 0;
          lyr._custSel = 0;
          lyr._turfFeat = { type: 'Feature', properties: { day, muni, key: keyNorm }, geometry: feat.geometry };

          coveragePolysAll.push({ feat: lyr._turfFeat, perDay, layerRef: lyr });

          lyr.on('click', (e) => {
            L.DomEvent.stopPropagation(e);
            if (currentFocus === lyr) openPolygonPopup(lyr);
            else { focusFeature(lyr); openPolygonPopup(lyr); }
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
    } catch(e) { /* non-fatal */ }

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

    // ---------- selection ----------
    async function applySelection() {
      clearFocus(false);

      const selUrl = qs.get('sel') || cfg.selection.url;
      phase('Loading selection…');
      const csvText = await fetchText(selUrl);
      const rowsAA = parseCsvRows(csvText);

      const hdr = findHeaderFlexible(rowsAA, cfg.selection.schema);
      const selectedSet = new Set();

      if (hdr) {
        const { headerIndex, keysCol } = hdr;
        for (let i = headerIndex + 1; i < rowsAA.length; i++) {
          const r = rowsAA[i]; if (!r || r.length === 0) continue;
          const raw = ((r[keysCol] || '') + '');
          const keys = splitKeys(raw, cfg.selection.schema.delimiter).map(normalizeKey).filter(Boolean);
          keys.forEach(k => selectedSet.add(k));
        }
      }

      if (selectedSet.size === 0) {
        warn(`Selection parsed 0 keys. Check your published CSV has a "zone keys" column with values in B48:E51.`);
      } else {
        info(`Loaded ${selectedSet.size} selected key(s).`);
      }

      hasSelection = selectedSet.size > 0;
      selectionBounds = null;
      coveragePolysSelected = [];
      selectedMunicipalities = [];

      const selBounds = [];

      dayLayers.forEach(({ layer, perDay }) => {
        layer.eachLayer(lyr => {
          const hit = selectedSet.has(lyr._routeKey);
          lyr._isSelected = !!hit;

          if (hit) {
            applyStyleSelected(lyr, perDay, cfg);
            showLabel(lyr, lyr._labelTxt);
            coveragePolysSelected.push({ feat: lyr._turfFeat, perDay, layerRef: lyr });
            if (lyr.getBounds) selBounds.push(lyr.getBounds());
            layer.addLayer(lyr);
            if (lyr._labelTxt) selectedMunicipalities.push(lyr._labelTxt);
          } else {
            if (cfg.style.unselectedMode === 'hide' && hasSelection) {
              hideLabel(lyr); layer.removeLayer(lyr);
            } else {
              applyStyleDim(lyr, perDay, cfg);
              hideLabel(lyr); layer.addLayer(lyr);
            }
          }
        });
      });

      // de-dupe & sort municipalities
      selectedMunicipalities = Array.from(new Set(selectedMunicipalities))
        .sort((a,b) => a.toLowerCase().localeCompare(b.toLowerCase()));

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

      // per-day legend fractions: day-in-selection / total-in-selection
      const legendCounts = {
        Wednesday: { selected: custByDayInSel.Wednesday, total: custWithinSel },
        Thursday:  { selected: custByDayInSel.Thursday,  total: custWithinSel },
        Friday:    { selected: custByDayInSel.Friday,    total: custWithinSel },
        Saturday:  { selected: custByDayInSel.Saturday,  total: custWithinSel }
      };

      renderLegend(cfg, legendCounts, custWithinSel, custOutsideSel);
      setStatus(makeStatusLine(selectedMunicipalities, custWithinSel, custOutsideSel));
    }

    // ---------- customers ----------
    async function loadCustomersIfAny() {
      customerLayer.clearLayers();
      customerMarkers.length = 0;
      customerCount = 0;

      const custToggle = qs.get('cust');
      if (custToggle && custToggle.toLowerCase() === 'off') {
        custWithinSel = custOutsideSel = 0;
        resetDayCounts();
        renderLegend(cfg, null, custWithinSel, custOutsideSel);
        setStatus(makeStatusLine(selectedMunicipalities, custWithinSel, custOutsideSel));
        return;
      }
      if (!cfg.customers || cfg.customers.enabled === false || !cfg.customers.url) {
        custWithinSel = custOutsideSel = 0;
        resetDayCounts();
        renderLegend(cfg, null, custWithinSel, custOutsideSel);
        setStatus(makeStatusLine(selectedMunicipalities, custWithinSel, custOutsideSel));
        return;
      }

      phase('Loading customers…');
      const text = await fetchText(cfg.customers.url);
      const rows = parseCsvRows(text);
      if (!rows.length) {
        custWithinSel = custOutsideSel = 0;
        resetDayCounts();
        renderLegend(cfg, null, custWithinSel, custOutsideSel);
        setStatus(makeStatusLine(selectedMunicipalities, custWithinSel, custOutsideSel));
        return;
      }

      const hdrIdx = findCustomerHeaderIndex(rows, cfg.customers.schema);
      if (hdrIdx === -1) {
        warn('Customers CSV: no header with "Verified Coordinates" or "Order Note" found.');
        custWithinSel = custOutsideSel = 0;
        resetDayCounts();
        renderLegend(cfg, null, custWithinSel, custOutsideSel);
        setStatus(makeStatusLine(selectedMunicipalities, custWithinSel, custOutsideSel));
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
      let added = 0;
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
        added++;
      }
      customerCount = added;
      info(`Loaded ${customerCount} customers.`);

      // color + count
      recolorAndRecountCustomers();

      const legendCounts = {
        Wednesday: { selected: custByDayInSel.Wednesday, total: custWithinSel },
        Thursday:  { selected: custByDayInSel.Thursday,  total: custWithinSel },
        Friday:    { selected: custByDayInSel.Friday,    total: custWithinSel },
        Saturday:  { selected: custByDayInSel.Saturday,  total: custWithinSel }
      };

      renderLegend(cfg, legendCounts, custWithinSel, custOutsideSel);
      setStatus(makeStatusLine(selectedMunicipalities, custWithinSel, custOutsideSel));
    }

    // keep this INSIDE the try so it can see custByDayInSel
    function resetDayCounts(){
      custByDayInSel.Wednesday = 0;
      custByDayInSel.Thursday  = 0;
      custByDayInSel.Friday    = 0;
      custByDayInSel.Saturday  = 0;
    }

    // recolor + recount
    function recolorAndRecountCustomers() {
      let inSel = 0, outSel = 0;
      resetDayCounts();

      // reset per-polygon counts
      coveragePolysAll.forEach(rec => { if (rec.layerRef) { rec.layerRef._custAny = 0; rec.layerRef._custSel = 0; } });

      const turfOn = (typeof turf !== 'undefined');

      // customer outside style
      const cst = cfg.style.customers || {};
      const out = cst.outside || {};
      const outsideStroke = out.stroke || '#7a7a7a';
      const outsideFill   = out.fill   || '#c7c7c7';
      const outsideOpacity = out.opacity != null ? out.opacity : 0.8;
      const outsideFillOpacity = out.fillOpacity != null ? out.fillOpacity : 0.6;

      for (const rec of customerMarkers) {
        // default outside (grey)
        let style = {
          radius:  (cst.radius || 9),
          color:   outsideStroke,
          weight:  (cst.weightPx || 2),
          opacity: outsideOpacity,
          fillColor: outsideFill,
          fillOpacity: outsideFillOpacity
        };

        let insideSel = false;
        let selDay = null;

        if (turfOn) {
          const pt = turf.point([rec.lng, rec.lat]);

          // count for ANY polygon (total customers per municipality)
          for (let j = 0; j < coveragePolysAll.length; j++) {
            if (turf.booleanPointInPolygon(pt, coveragePolysAll[j].feat)) {
              const lyr = coveragePolysAll[j].layerRef;
              if (lyr) lyr._custAny += 1;
              break;
            }
          }

          // color & per-day count using SELECTED polygons only
          if (hasSelection && coveragePolysSelected.length) {
            for (let k = 0; k < coveragePolysSelected.length; k++) {
              if (turf.booleanPointInPolygon(pt, coveragePolysSelected[k].feat)) {
                insideSel = true;
                selDay = (coveragePolysSelected[k].feat.properties.day || '').trim();
                const pd = coveragePolysSelected[k].perDay || {};
                style = {
                  radius:  (cst.radius || 9),
                  color:   pd.stroke || (cst.stroke || '#111'),
                  weight:  (cst.weightPx || 2),
                  opacity: (cst.opacity != null ? cst.opacity : 0.95),
                  fillColor: pd.fill || (cst.fill || '#ffffff'),
                  fillOpacity: (cst.fillOpacity != null ? cst.fillOpacity : 0.95)
                };
                const lyr = coveragePolysSelected[k].layerRef;
                if (lyr) lyr._custSel += 1;
                break;
              }
            }
          }
        }

        if (insideSel) {
          inSel++;
          if (selDay && custByDayInSel[selDay] != null) custByDayInSel[selDay] += 1;
        } else {
          outSel++;
        }

        rec.marker.setStyle(style);
      }

      custWithinSel = inSel;
      custOutsideSel = outSel;
    }

    // polygon popup with counts
    function openPolygonPopup(lyr) {
      const muni = lyr._labelTxt || 'Municipality';
      const totalAny = Number(lyr._custAny || 0);
      const inSel    = Number(lyr._custSel || 0);
      const html = `<div><strong>${escapeHtml(muni)}</strong><br>Customers: ${totalAny}` +
                   (lyr._isSelected ? ` <span style="opacity:.8">(in selection: ${inSel})</span>` : '') +
                   `</div>`;
      const center = lyr.getBounds?.().getCenter?.();
      if (center) L.popup({autoClose:true, closeOnClick:true}).setLatLng(center).setContent(html).openOn(map);
    }

    // focus/highlight helpers
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

    // style helpers
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

  // ---------- shared helpers (safe outside the try) ----------
  function cb(u) { return (u.includes('?') ? '&' : '?') + 'cb=' + Date.now(); }
  async function fetchJson(url) { const res = await fetch(url + cb(url)); if (!res.ok) throw new Error(`Fetch ${res.status} for ${url}`); return res.json(); }
  async function fetchText(url) { const res = await fetch(url + cb(url)); if (!res.ok) throw new Error(`Fetch ${res.status} for ${url}`); return res.text(); }

  // CSV parser
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

  // selection header finder (tolerant)
  function findHeaderFlexible(rows, schema) {
    if (!rows || !rows.length) return null;
    const wantKeys = (schema.keys || 'zone keys').toLowerCase();
    const like = s => (s||'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
    for (let i=0;i<rows.length;i++){
      const row = rows[i] || [];
      let keysCol = -1;
      row.forEach((h, idx) => {
        const v = like(h);
        if (keysCol === -1 && (v === like(wantKeys) || v.startsWith('zone key') || v === 'keys' || (v.includes('selected') && v.includes('keys'))))
          keysCol = idx;
      });
      if (keysCol !== -1) return { headerIndex: i, keysCol };
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

  // key utils
  function splitKeys(s, delim) { return String(s||'').split(/[;,]/).map(x => x.trim()).filter(Boolean); }
  function normalizeKey(s) {
    s = String(s || '').trim().toUpperCase();
    const m = s.match(/^([WTFS])0*(\d+)$/);
    if (m) return m[1] + String(parseInt(m[2], 10));
    return s;
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

  // legend & status & error helpers
  function renderLegend(cfg, legendCounts, custIn, custOut) {
    const el = document.getElementById('legend');
    const rowsHtml = (cfg.layers || []).map(Lcfg => {
      const st = (cfg.style && cfg.style.perDay && cfg.style.perDay[Lcfg.day]) || {};
      const c  = (legendCounts && legendCounts[Lcfg.day]) ? legendCounts[Lcfg.day] : { selected: 0, total: 0 };
      const frac = (c.total > 0) ? `${c.selected}/${c.total}` : '0/0';
      return `<div class="row">
        <span class="swatch" style="background:${st.fill}; border-color:${st.stroke}"></span>
        <div>${Lcfg.day}</div>
        <div class="counts">${frac}</div>
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

  function makeStatusLine(selMunis, inCount, outCount) {
    const muniList = (Array.isArray(selMunis) && selMunis.length)
      ? selMunis.join(', ')
      : '—';
    return `Customers (in/out): ${inCount}/${outCount} • Municipalities: ${muniList}`;
  }

  function setStatus(msg) { const n = document.getElementById('status'); if (n) n.textContent = msg || ''; }
  function phase(msg){ setStatus(`⏳ ${msg}`); }
  function info(msg){ setStatus(`ℹ️ ${msg}`); }
  function warn(msg){ const e = document.getElementById('error'); e.style.display='block'; e.innerHTML = `<strong>Warning</strong><br>${escapeHtml(msg)}`; }
  function showError(e) {
    const el = document.getElementById('error');
    el.style.display = 'block';
    el.innerHTML = `<strong>Load error</strong><br>${e && e.message ? escapeHtml(e.message) : e}`;
  }
})();
