/* apps.js — viewer + manual-first route stepper + PNG snapshots
   --------------------------------------------------------------
   Defaults to **manual** user-driven controls when a batch is provided.
   Buttons: Prev · Next · Stats · Snap
   Flow:
     • Start: shows ALL selected routes at once (no customers)
     • Next/Prev: zoom to one route and show **only that route’s customers**
     • Stats toggle: show/hide 5-line dispatch card overlay (also drawn into PNG)
     • Snap: capture Leaflet map via leaflet-image, draw the stats card onto the canvas,
             POST to Apps Script doPost (cb=...) or download fallback

   Keeps all existing features (3-tier layers, CSV selection/assignments, customers, overlays).
   Snapshot images are map-only + drawn stats (no UI buttons).
*/
(async function () {
  // ---- error surfacing ----
  window.addEventListener('error', (e) => showTopError('Script error', (e && e.message) ? e.message : String(e)));
  window.addEventListener('unhandledrejection', (e) => showTopError('Promise error', (e?.reason?.message) || String(e?.reason || e)));

  // ---- URL / config ----
  const qs = new URLSearchParams(location.search);
  const cfgUrl = qs.get('cfg') || './config/app.config.json';
  const parseJSON = (s, fb=null) => { try { return JSON.parse(s); } catch { return fb; } };

  const driverMetaParam = parseJSON(qs.get('driverMeta') || '[]', []);
  const assignMapParam  = parseJSON(qs.get('assignMap')  || '{}', {});
  const cbUrl           = qs.get('cb') || '';
  const batchParam      = qs.get('batch'); // websafe b64 JSON

  // Manual is the default when batch exists; pass &auto=1 to re-enable auto-export loop
  const manualMode = !!batchParam && qs.get('auto') !== '1';

  // Runtime state for manual controller
  let batchItems = parseBatchItems(batchParam); // [{name, day, driver, keys[], stats, outName}]
  let currentIndex = -1;         // -1 = overview (all routes at once)
  let statsVisible = false;      // UI toggle
  let manualSelectedKeys = null; // overrides selection when set
  let runtimeCustEnabled = null; // null=default pipeline, true=force on (loaded), false=force off (none loaded)

  // Driver assignment + palette
  let activeAssignMap = { ...assignMapParam };
  let driverMeta      = Array.isArray(driverMetaParam) ? [...driverMetaParam] : [];

  // ---- map init ----
  phase('Loading config…');
  const cfg = await fetchJson(cfgUrl);

  if (typeof L === 'undefined' || !document.getElementById('map')) {
    renderError('Leaflet or #map element not available.');
    return;
  }

  phase('Initializing map…');
  const map = L.map('map', { preferCanvas: true }).setView([43.55, -80.25], 8);
  const osmTiles = 'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png';
  const base = L.tileLayer(osmTiles, { attribution: '&copy; OpenStreetMap contributors', crossOrigin: true }).addTo(map);
  try { const el = base.getContainer?.(); if (el) el.style.filter = 'grayscale(1)'; } catch {}

  // ---- collections / state ----
  const baseDayLayers = [], quadDayLayers = [], subqDayLayers = [], allDaySets=[baseDayLayers,quadDayLayers,subqDayLayers];
  let allBounds=null, selectionBounds=null, hasSelection=false;

  const boundaryFeatures = [];
  const customerLayer = L.layerGroup().addTo(map);
  const customerMarkers = [];   // [{ marker, lat, lng, visible:bool }]
  let customerCount=0, custWithinSel=0, custOutsideSel=0;
  const custByDayInSel = { Wednesday:0, Thursday:0, Friday:0, Saturday:0 };

  let currentFocus=null;
  let coveragePolysAll=[], coveragePolysSelected=[];
  let selectedMunicipalities=[];
  let driverSelectedCounts={}, driverOverlays={};

  let selectedOrderedKeys = [];
  let visibleSelectedKeysSet = new Set();

  // Render empty panels first
  renderLegend(cfg, null, custWithinSel, custOutsideSel, false);
  renderDriversPanel([], {}, false, {}, 0);

  // ---- load polygon layers (base → quads → subquads) ----
  phase('Loading polygon layers…');
  await loadLayerSet(cfg.layers, baseDayLayers, true);
  if (cfg.layersQuadrants?.length)    await loadLayerSet(cfg.layersQuadrants, quadDayLayers, true);
  if (cfg.layersSubquadrants?.length) await loadLayerSet(cfg.layersSubquadrants, subqDayLayers, true);

  // Optional boundary mask
  try {
    if (L.TileLayer?.boundaryCanvas && boundaryFeatures.length) {
      const boundaryFC = { type:'FeatureCollection', features: boundaryFeatures };
      const maskTiles = L.TileLayer.boundaryCanvas(osmTiles, { boundary: boundaryFC, attribution: '' });
      maskTiles.addTo(map).setZIndex(2); base.setZIndex(1);
    }
  } catch {}

  map.on('click', () => { clearFocus(true); });
  map.on('movestart', () => { clearFocus(false); map.closePopup(); });
  map.on('zoomstart',  () => { map.closePopup(); });

  // ---- initial selection + customers ----
  // Manual/overview start: union of all route keys, and customers OFF
  if (manualMode && batchItems.length) {
    manualSelectedKeys = unionAllKeys(batchItems);
    runtimeCustEnabled = false; // start with no customers
  }

  await applySelection();
  await loadCustomersIfAny();

  // ---- MANUAL CONTROLS (default when batch present) ----
  if (manualMode && batchItems.length) {
    injectManualUI();
    // Overview view (index -1)
    await zoomToOverview();
  } else if (!manualMode && batchItems.length) {
    // optional: auto-export loop remains available by passing auto=1
    await runAutoExport(batchItems);
  }

  // =================================================================
  // UI: Manual controller
  // =================================================================
  function injectManualUI(){
    // CSS
    const css = document.createElement('style');
    css.textContent = `
      .route-toolbar{position:absolute;top:10px;left:50%;transform:translateX(-50%);z-index:9000;display:flex;gap:8px}
      .route-toolbar button{background:#111;color:#fff;border:none;border-radius:6px;padding:8px 12px;font:600 14px system-ui;cursor:pointer;opacity:.9}
      .route-toolbar button:hover{opacity:1}
      .stats-card{position:absolute;top:12px;right:12px;z-index:8500;background:rgba(255,255,255,.92);backdrop-filter:saturate(120%) blur(2px);border-radius:10px;box-shadow:0 6px 18px rgba(0,0,0,.12);padding:10px 12px;max-width:min(52vw,640px);display:none}
      .stats-card.visible{display:block}
      .stats-card h4{margin:0 0 4px 0;font:700 16px system-ui}
      .stats-card .sub{font:500 13px system-ui;color:#333;margin:0 0 8px 0}
      .stats-card ul{margin:0;padding:0;list-style:none}
      .stats-card li{font:500 14px system-ui;margin:2px 0}
    `;
    document.head.appendChild(css);

    // Toolbar
    const bar = document.createElement('div');
    bar.className = 'route-toolbar';
    bar.innerHTML = `
      <button id="btnPrev" title="Previous route">◀ Prev</button>
      <button id="btnNext" title="Next route">Next ▶</button>
      <button id="btnStats" title="Toggle stats card">Stats</button>
      <button id="btnSnap"  title="Snapshot to Drive">📸 Snap</button>
    `;
    document.body.appendChild(bar);

    // Stats card (DOM preview; snapshot draws separately onto canvas)
    const card = document.createElement('div');
    card.id = 'statsCard';
    card.className = 'stats-card';
    card.innerHTML = `<h4></h4><div class="sub"></div><ul></ul>`;
    document.body.appendChild(card);

    // Wire buttons
    document.getElementById('btnPrev').addEventListener('click', async () => { await stepRoute(-1); });
    document.getElementById('btnNext').addEventListener('click', async () => { await stepRoute(+1); });
    document.getElementById('btnStats').addEventListener('click', () => toggleStats());
    document.getElementById('btnSnap').addEventListener('click', async () => { await snapshotCurrent(); });

    updateButtons();
  }

  function updateButtons(){
    const haveFocus = currentIndex >= 0 && currentIndex < batchItems.length;
    const prev = document.getElementById('btnPrev');
    const next = document.getElementById('btnNext');
    const stat = document.getElementById('btnStats');
    const snap = document.getElementById('btnSnap');
    if (prev) prev.disabled = batchItems.length === 0;
    if (next) next.disabled = batchItems.length === 0;
    if (stat) stat.disabled = !haveFocus;
    if (snap) snap.disabled = !haveFocus;
  }

  function toggleStats(){
    statsVisible = !statsVisible;
    const card = document.getElementById('statsCard');
    if (!card) return;
    if (statsVisible) card.classList.add('visible'); else card.classList.remove('visible');
  }

  function renderStatsCard(it){
    const card = document.getElementById('statsCard'); if (!card) return;
    if (!it) { card.classList.remove('visible'); return; }
    const s = it?.stats || {};
    card.querySelector('h4').textContent = `${it.day} — ${it.driver}`;
    card.querySelector('.sub').textContent = String(it.name || '');
    const lines = [
      `Deliveries: ${s.deliveries || 0}`,
      `Base boxes: ${s.baseBoxes || 0}`,
      `Customs: ${s.customs || 0}`,
      `Add-ons: ${s.addOns || 0}`,
      `Apts: ${s.apartments || 0}`
    ];
    card.querySelector('ul').innerHTML = lines.map(t => `<li>${escapeHtml(t)}</li>`).join('');
    if (statsVisible) card.classList.add('visible'); else card.classList.remove('visible');
  }

  async function stepRoute(delta){
    if (!batchItems.length) return;

    // From overview (-1), Next focuses index 0
    if (currentIndex === -1 && delta > 0) currentIndex = 0;
    else if (currentIndex === -1 && delta < 0) currentIndex = batchItems.length - 1; // wrap to last
    else currentIndex = (currentIndex + delta + batchItems.length) % batchItems.length;

    const it = batchItems[currentIndex];
    manualSelectedKeys = (it.keys || []).map(normalizeKey);
    runtimeCustEnabled = true; // on focus: load customers and show only this route's
    await applySelection();
    await loadCustomersIfAny();

    // Zoom to focused route
    if (selectionBounds) map.fitBounds(selectionBounds.pad(0.08));

    renderStatsCard(it);
    updateButtons();
  }

  async function zoomToOverview(){
    currentIndex = -1;
    manualSelectedKeys = unionAllKeys(batchItems);
    runtimeCustEnabled = false; // hide customers
    statsVisible = false; renderStatsCard(null);

    await applySelection();
    await loadCustomersIfAny();

    if (selectionBounds) map.fitBounds(selectionBounds.pad(0.08));
    updateButtons();
  }

  async function snapshotCurrent(){
    const it = (currentIndex >=0 && currentIndex < batchItems.length) ? batchItems[currentIndex] : null;
    if (!it) return; // no-op on overview

    // Render Leaflet canvas
    const canvas = await new Promise((resolve, reject) =>
      window.leafletImage ? leafletImage(map, (err, c) => (err ? reject(err) : resolve(c)))
                          : reject(new Error('leaflet-image not loaded')));

    // Draw stats card onto canvas if visible
    if (statsVisible) drawStatsOntoCanvas(canvas, it);

    // Post PNG to Apps Script receiver (no-cors to avoid preflight)
    const png = canvas.toDataURL('image/png');
    const name = it.outName || `${it.driver}_${it.day}.png`;
    if (cbUrl) {
      try {
        await fetch(cbUrl, {
          method: 'POST',
          mode: 'no-cors',
          body: JSON.stringify({ name, day: it.day, driver: it.driver, routeName: it.name, pngBase64: png })
        });
      } catch (e) {
        console.warn('POST failed, falling back to download:', e);
        downloadFallback(png, name);
      }
    } else {
      downloadFallback(png, name);
    }
  }

  function drawStatsOntoCanvas(canvas, it){
    const ctx = canvas.getContext('2d');
    const w = canvas.width, h = canvas.height;
    const pad = Math.round(Math.min(w, h) * 0.02);
    const lineH = Math.round(Math.min(w, h) * 0.03);
    const boxW = Math.min(Math.round(w * 0.5), 620);
    const boxH = lineH * 6 + pad * 2;

    ctx.fillStyle = 'rgba(255,255,255,0.92)';
    roundRect(ctx, w - boxW - pad, pad, boxW, boxH, 10).fill();

    ctx.fillStyle = '#111';
    ctx.font = 'bold ' + (lineH * 0.9) + 'px system-ui, -apple-system, Segoe UI, Roboto, sans-serif';
    const x = w - boxW - pad + pad;
    let y = pad + lineH;

    ctx.fillText(`${it.day} — ${it.driver}`, x, y); y += Math.round(lineH * 1.1);
    ctx.font = 'normal ' + (lineH * 0.85) + 'px system-ui, -apple-system, Segoe UI, Roboto, sans-serif';
    ctx.fillText(String(it.name || ''), x, y); y += Math.round(lineH * 1.05);

    const s = it.stats || {};
    [
      `Deliveries: ${s.deliveries || 0}`,
      `Base boxes: ${s.baseBoxes || 0}`,
      `Customs: ${s.customs || 0}`,
      `Add-ons: ${s.addOns || 0}`,
      `Apts: ${s.apartments || 0}`
    ].forEach(t => { ctx.fillText(t, x, y); y += Math.round(lineH * 1.0); });
  }

  function roundRect(ctx, x, y, w, h, r) {
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.arcTo(x + w, y,   x + w, y + h, r);
    ctx.arcTo(x + w, y + h, x,   y + h, r);
    ctx.arcTo(x,   y + h, x,   y,   r);
    ctx.arcTo(x,   y,   x + w, y,   r);
    ctx.closePath();
    return ctx;
  }

  function downloadFallback(uri, name){
    const a = document.createElement('a'); a.href = uri; a.download = name; document.body.appendChild(a); a.click(); a.remove();
  }

  function unionAllKeys(items){
    const set = new Set();
    (items||[]).forEach(it => (it.keys||[]).forEach(k => set.add(normalizeKey(k))));
    return Array.from(set);
  }

  // =================================================================
  // Optional: legacy auto-export (only if auto=1 in URL)
  // =================================================================
  async function runAutoExport(items){
    // hide side panels for clean snapshots
    const legend = document.getElementById('legend');
    const drivers= document.getElementById('drivers');
    const status = document.getElementById('status');
    const err    = document.getElementById('error');
    [legend,drivers,status,err].forEach(n => { if(n) n.style.display='none'; });

    for (let i=0;i<items.length;i++){
      const it = items[i];
      try{
        manualSelectedKeys = (it.keys||[]).map(normalizeKey);
        runtimeCustEnabled = true;
        await applySelection();
        await loadCustomersIfAny();
        if (selectionBounds) map.fitBounds(selectionBounds.pad(0.08));
        await sleep(350);

        const canvas = await new Promise((resolve, reject) =>
          window.leafletImage ? leafletImage(map, (err, c) => (err ? reject(err) : resolve(c)))
                              : reject(new Error('leaflet-image not loaded')));
        drawStatsOntoCanvas(canvas, it);
        const png = canvas.toDataURL('image/png');
        const name = it.outName || `${it.driver}_${it.day}.png`;
        if (cbUrl) {
          await fetch(cbUrl, { method:'POST', mode:'no-cors', body: JSON.stringify({ name, day: it.day, driver: it.driver, routeName: it.name, pngBase64: png }) });
        } else { downloadFallback(png, name); }
      }catch(e){ console.error('auto export item failed', e); }
    }
  }

  // =================================================================
  // LOADERS / SELECTION / CUSTOMERS — core viewer behavior
  // =================================================================
  async function loadLayerSet(arr, collector, addToMap = true) {
    for (const Lcfg of (arr || [])) {
      const gj = await fetchJson(Lcfg.url);
      const perDay = (cfg.style?.perDay?.[Lcfg.day]) || {};
      const color = perDay.stroke || '#666';
      const fillColor = perDay.fill || '#ccc';
      const features = [];

      const layer = L.geoJSON(gj, {
        style: () => ({
          color,
          weight: cfg.style?.dimmed?.weightPx ?? 1,
          opacity: cfg.style?.dimmed?.strokeOpacity ?? 0.35,
          fillColor,
          fillOpacity: dimFill(perDay, cfg)
        }),
        onEachFeature: (feat, lyr) => {
          const p = feat.properties || {};
          const rawKey  = (p[cfg.fields.key]  ?? '').toString().trim();
          const keyNorm = normalizeKey(rawKey);
          const muni    = smartTitleCase((p[cfg.fields.muni] ?? '').toString().trim());
          const day     = Lcfg.day;

          lyr._routeKey   = keyNorm;
          lyr._baseKey    = baseKeyFrom(keyNorm);
          lyr._day        = day;
          lyr._perDay     = perDay;
          lyr._labelTxt   = muni;
          lyr._isSelected = false;
          lyr._custAny    = 0;
          lyr._custSel    = 0;
          lyr._turfFeat   = { type:'Feature', properties:{ day, muni, key:keyNorm }, geometry: feat.geometry };

          boundaryFeatures.push({ type:'Feature', geometry: feat.geometry });
          features.push(lyr);

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
      });

      collector.push({ day: Lcfg.day, layer, perDay, features });
      if (addToMap) layer.addTo(map);
    }
  }

  async function applySelection() {
    clearFocus(false);

    const selectedSet = new Set();
    selectedOrderedKeys = [];
    let loadedSelectionCsv = false;

    if (Array.isArray(manualSelectedKeys) && manualSelectedKeys.length) {
      manualSelectedKeys.map(normalizeKey).forEach(k => { if (!selectedSet.has(k)) { selectedSet.add(k); selectedOrderedKeys.push(k); } });
    } else {
      const selUrl = qs.get('sel') || (cfg.selection && cfg.selection.url) || '';
      if (selUrl) {
        phase('Loading selection…');
        try {
          const csvText = await fetchText(selUrl);
          const rowsAA = parseCsvRows(csvText);
          loadedSelectionCsv = rowsAA?.length > 0;

          const hdr = findHeaderFlexible(rowsAA, cfg.selection?.schema || { keys: 'zone keys' });
          if (hdr) {
            const { headerIndex, keysCol } = hdr;
            for (let i=headerIndex+1;i<rowsAA.length;i++) {
              const r = rowsAA[i]; if (!r?.length) continue;
              const ks = splitKeys(r[keysCol]).map(normalizeKey);
              for (const k of ks) { if (!selectedSet.has(k)) { selectedSet.add(k); selectedOrderedKeys.push(k); } }
            }
          }

          const knownDriverNames = new Set((driverMeta||[]).map(d => String(d.name||'').toLowerCase()).filter(Boolean));
          const derivedAssign = extractAssignmentsFromCsv(rowsAA, knownDriverNames);
          activeAssignMap = { ...activeAssignMap, ...derivedAssign };
          driverMeta = ensureDriverMeta(driverMeta, Object.values(activeAssignMap));
        } catch (e) { warn(`Selection CSV load failed (${e?.message||e}). Showing all zones.`); }
      }
    }

    const noKeys = selectedSet.size === 0;
    hasSelection = !noKeys;
    selectionBounds = null;
    coveragePolysSelected = [];
    selectedMunicipalities = [];
    visibleSelectedKeysSet = new Set();

    const recordVisibleSelectedKey = (k) => { if (k) visibleSelectedKeysSet.add(k); };

    const quadBasesSelected = new Set();
    const subqBasesSelected = new Set();
    const subqQuadsSelected = new Set();

    if (!noKeys) {
      for (const k of selectedSet) {
        if (isSubquadrantKey(k)) { subqBasesSelected.add(baseKeyFrom(k)); const pq = basePlusQuad(k); if (pq) subqQuadsSelected.add(pq); }
        else if (isQuadrantKey(k)) quadBasesSelected.add(baseKeyFrom(k));
      }
    }

    const setFeatureVisible = (entry, lyr, visible, isSelected) => {
      const has = entry.layer.hasLayer(lyr);
      if (visible && !has) entry.layer.addLayer(lyr);
      if (!visible && has) entry.layer.removeLayer(lyr);
      lyr._isSelected = !!(visible && isSelected);

      if (visible) {
        if (isSelected) applyStyleSelected(lyr, entry.perDay, cfg);
        else            applyStyleDim(lyr, entry.perDay, cfg);

        if (isSelected) {
          recordVisibleSelectedKey(lyr._routeKey);
          if (lyr.getBounds) { const b = lyr.getBounds(); selectionBounds = selectionBounds ? selectionBounds.extend(b) : L.latLngBounds(b); }
          coveragePolysSelected.push({ feat: lyr._turfFeat, perDay: entry.perDay, layerRef: lyr });
          if (lyr._labelTxt) selectedMunicipalities.push(lyr._labelTxt);
          showLabel(lyr, lyr._labelTxt);
        } else {
          hideLabel(lyr);
        }
      } else {
        hideLabel(lyr);
      }
    };

    if (noKeys) {
      info(loadedSelectionCsv ? 'No selection keys found; showing all zones.' : 'Showing all zones.');
      for (const arr of allDaySets) for (const entry of arr) for (const lyr of entry.features) setFeatureVisible(entry, lyr, true, false);
      rebuildCoverageFromVisible();
      selectedMunicipalities = [];
      if (cfg.behavior?.autoZoom) { if (allBounds) map.fitBounds(allBounds.pad(0.1)); }
    } else {
      for (const entry of baseDayLayers)  for (const lyr of entry.features) { const k=lyr._routeKey; const visible = selectedSet.has(k) && !quadBasesSelected.has(k) && !subqBasesSelected.has(k); setFeatureVisible(entry, lyr, visible, visible); }
      for (const entry of quadDayLayers)  for (const lyr of entry.features) { const k=lyr._routeKey; const bq=basePlusQuad(k); const visible = selectedSet.has(k) && !subqQuadsSelected.has(bq); setFeatureVisible(entry, lyr, visible, visible); }
      for (const entry of subqDayLayers)  for (const lyr of entry.features) { const k=lyr._routeKey; const visible = selectedSet.has(k); setFeatureVisible(entry, lyr, visible, visible); }

      rebuildCoverageFromVisible();
      selectedMunicipalities = Array.from(new Set(selectedMunicipalities)).sort((a,b)=>a.toLowerCase().localeCompare(b.toLowerCase()));

      if (cfg.behavior?.autoZoom) {
        if (selectionBounds) map.fitBounds(selectionBounds.pad(0.1)); else if (allBounds) map.fitBounds(allBounds.pad(0.1));
      }
      info(`Loaded ${selectedSet.size} selected key(s).`);
    }

    recolorAndRecountCustomers();

    const legendCounts = {
      Wednesday: { selected: custByDayInSel.Wednesday, total: custWithinSel },
      Thursday:  { selected: custByDayInSel.Thursday,  total: custWithinSel },
      Friday:    { selected: custByDayInSel.Friday,    total: custWithinSel },
      Saturday:  { selected: custByDayInSel.Saturday,  total: custWithinSel }
    };
    const activeKeysOrderedLower = selectedOrderedKeys.filter(k => visibleSelectedKeysSet.has(k)).map(k => k.toLowerCase());

    renderLegend(cfg, legendCounts, custWithinSel, custOutsideSel, false);
    setStatus(makeStatusLine(selectedMunicipalities, custWithinSel, custOutsideSel, activeKeysOrderedLower));
    await rebuildDriverOverlays();
  }

  function rebuildCoverageFromVisible() {
    coveragePolysAll = [];
    for (const arr of allDaySets) for (const entry of arr) for (const lyr of entry.features) {
      if (entry.layer.hasLayer(lyr)) { coveragePolysAll.push({ feat: lyr._turfFeat, perDay: entry.perDay, layerRef: lyr }); lyr._custAny = 0; lyr._custSel = 0; }
    }
  }

  async function loadCustomersIfAny() {
    // runtimeCustEnabled === false → no customers at all (overview)
    // runtimeCustEnabled === true  → load customers, and we'll hide/show per-route in recolor step
    // runtimeCustEnabled === null  → follow normal cfg behavior
    customerLayer.clearLayers();
    customerMarkers.length = 0;
    customerCount = 0;

    const custCfg = cfg.customers || {};
    const forceOff = (runtimeCustEnabled === false);
    const forceOn  = (runtimeCustEnabled === true);

    if (forceOff || !custCfg.enabled || !custCfg.url) {
      custWithinSel = custOutsideSel = 0; resetDayCounts();
      renderLegend(cfg, null, custWithinSel, custOutsideSel, false);
      setStatus(makeStatusLine(selectedMunicipalities, custWithinSel, custOutsideSel, []));
      renderDriversPanel(driverMeta, driverOverlays, true, driverSelectedCounts, custWithinSel);
      return;
    }

    phase('Loading customers…');
    const text = await fetchText(custCfg.url);
    const rows = parseCsvRows(text);
    if (!rows.length) { custWithinSel = custOutsideSel = 0; resetDayCounts(); renderLegend(cfg, null, custWithinSel, custOutsideSel, false); setStatus(makeStatusLine(selectedMunicipalities, custWithinSel, custOutsideSel, [])); renderDriversPanel(driverMeta, driverOverlays, true, driverSelectedCounts, custWithinSel); return; }

    const hdrIdx = findCustomerHeaderIndex(rows, custCfg.schema || { coords: 'Verified Coordinates', note: 'Order Note' });
    if (hdrIdx === -1) { warn('Customers CSV: header not found.'); custWithinSel = custOutsideSel = 0; resetDayCounts(); renderLegend(cfg, null, custWithinSel, custOutsideSel, false); setStatus(makeStatusLine(selectedMunicipalities, custWithinSel, custOutsideSel, [])); renderDriversPanel(driverMeta, driverOverlays, true, driverSelectedCounts, custWithinSel); return; }

    const mapIdx = headerIndexMap(rows[hdrIdx], custCfg.schema || { coords: 'Verified Coordinates', note: 'Order Note' });
    const s = cfg.style?.customers || {};
    const baseStyle = { radius: s.radius || 9, color: s.stroke || '#111', weight: s.weightPx || 2, opacity: s.opacity ?? 0.95, fillColor: s.fill || '#ffffff', fillOpacity: s.fillOpacity ?? 0.95 };

    let added = 0;
    for (let i = hdrIdx + 1; i < rows.length; i++) {
      const r = rows[i]; if (!r?.length) continue;
      const coord = (mapIdx.coords !== -1) ? r[mapIdx.coords] : '';
      const note  = (mapIdx.note   !== -1) ? r[mapIdx.note]   : '';
      const ll = parseLatLng(coord); if (!ll) continue;
      const m = L.circleMarker([ll.lat, ll.lng], baseStyle).addTo(customerLayer);
      const popupHtml = note ? `<div style="max-width:260px">${escapeHtml(note)}</div>` : `<div>${ll.lat.toFixed(6)}, ${ll.lng.toFixed(6)}</div>`;
      m.bindPopup(popupHtml, { autoClose: true, closeOnClick: true });
      customerMarkers.push({ marker: m, lat: ll.lat, lng: ll.lng, visible: true });
      added++;
    }
    customerCount = added; info(`Loaded ${customerCount} customers.`);

    recolorAndRecountCustomers();

    const legendCounts = { Wednesday: { selected: custByDayInSel.Wednesday, total: custWithinSel }, Thursday: { selected: custByDayInSel.Thursday, total: custWithinSel }, Friday: { selected: custByDayInSel.Friday, total: custWithinSel }, Saturday: { selected: custByDayInSel.Saturday, total: custWithinSel } };
    renderLegend(cfg, legendCounts, custWithinSel, custOutsideSel, false);

    const activeKeysOrderedLower = selectedOrderedKeys.filter(k => visibleSelectedKeysSet.has(k)).map(k => k.toLowerCase());
    setStatus(makeStatusLine(selectedMunicipalities, custWithinSel, custOutsideSel, activeKeysOrderedLower));

    renderDriversPanel(driverMeta, driverOverlays, true, driverSelectedCounts, custWithinSel);
  }

  function resetDayCounts(){ custByDayInSel.Wednesday=0; custByDayInSel.Thursday=0; custByDayInSel.Friday=0; custByDayInSel.Saturday=0; }

  function recolorAndRecountCustomers() {
    let inSel = 0, outSel = 0;
    resetDayCounts();
    for (const rec of coveragePolysAll) { if (rec.layerRef) { rec.layerRef._custAny = 0; rec.layerRef._custSel = 0; } }

    const turfOn = (typeof turf !== 'undefined');
    const cst = cfg.style?.customers || {};
    const outStyle = {
      radius: cst.radius || 9,
      color:  '#7a7a7a',
      weight: cst.weightPx || 2,
      opacity: 0.8,
      fillColor: '#c7c7c7',
      fillOpacity: 0.6
    };

    // In manual focused mode: hide non-route customers entirely
    const onlySelectedCustomers = manualMode && (currentIndex >= 0);

    for (const rec of customerMarkers) {
      let show = true;
      let style = { ...outStyle };
      let insideSel = false;
      let selDay = null;

      if (turfOn) {
        const pt = turf.point([rec.lng, rec.lat]);

        // any visible polygon (for totalAny)
        for (let j = 0; j < coveragePolysAll.length; j++) {
          if (turf.booleanPointInPolygon(pt, coveragePolysAll[j].feat)) {
            const lyr = coveragePolysAll[j].layerRef;
            if (lyr) lyr._custAny += 1;
            break;
          }
        }

        // in current selection?
        for (let k = 0; k < coveragePolysSelected.length; k++) {
          if (turf.booleanPointInPolygon(pt, coveragePolysSelected[k].feat)) {
            insideSel = true;
            selDay = (coveragePolysSelected[k].feat.properties.day || '').trim();
            const pd = coveragePolysSelected[k].perDay || {};
            style = {
              radius:  (cst.radius || 9),
              color:   pd.stroke || (cst.stroke || '#111'),
              weight:  (cst.weightPx || 2),
              opacity: (cst.opacity ?? 0.95),
              fillColor: pd.fill || (cst.fill || '#ffffff'),
              fillOpacity: (cst.fillOpacity ?? 0.95)
            };
            const lyr = coveragePolysSelected[k].layerRef;
            if (lyr) lyr._custSel += 1;
            break;
          }
        }
      }

      // In focused (per-route) mode: hide customers outside selection entirely
      if (onlySelectedCustomers && !insideSel) {
        show = false;
      }

      if (insideSel) {
        inSel++;
        if (selDay && custByDayInSel[selDay] != null) custByDayInSel[selDay] += 1;
      } else {
        outSel++;
      }

      // add/remove marker from map based on "show"
      if (show && !rec.visible) { customerLayer.addLayer(rec.marker); rec.visible = true; }
      else if (!show && rec.visible) { customerLayer.removeLayer(rec.marker); rec.visible = false; }

      if (show) rec.marker.setStyle(style);
    }

    custWithinSel = inSel;
    custOutsideSel = outSel;

    driverSelectedCounts = computeDriverCounts();
    renderDriversPanel(driverMeta, driverOverlays, true, driverSelectedCounts, custWithinSel);
  }

  function computeDriverCounts() {
    const out = {};
    for (const rec of coveragePolysSelected) {
      const lyr = rec.layerRef; if (!lyr) continue;
      const drv = lookupDriverForKey(lyr._routeKey);
      if (!drv) continue;
      out[drv] = (out[drv] || 0) + (lyr._custSel || 0);
    }
    return out;
  }

  async function rebuildDriverOverlays() {
    // remove existing
    Object.values(driverOverlays).forEach(rec => { try { map.removeLayer(rec.group); } catch {} });
    for (const k of Object.keys(driverOverlays)) delete driverOverlays[k];

    // group selected features by driver
    const byDriver = new Map(); // name -> [geoJSON Feature]
    coveragePolysSelected.forEach(rec => {
      const key = rec.layerRef && rec.layerRef._routeKey;
      if (!key) return;
      const drv = lookupDriverForKey(key);
      if (!drv) return;
      if (!byDriver.has(drv)) byDriver.set(drv, []);
      byDriver.get(drv).push(rec.layerRef._turfFeat);
    });

    if (byDriver.size === 0) {
      renderDriversPanel(driverMeta, {}, false, driverSelectedCounts, custWithinSel);
      return;
    }

    byDriver.forEach((features, name) => {
      const meta = (driverMeta.find(d => (d.name || '').toLowerCase() === (name || '').toLowerCase())
                   || { name, color: colorFromName(name) });
      const color = meta.color || '#888';

      // halo under colored outline
      const halo = L.geoJSON({ type:'FeatureCollection', features }, {
        interactive:false,
        style:()=>({ color:'#ffffff', weight:(cfg.drivers?.outlineHaloWeightPx ?? 10), opacity:0.95, lineJoin:'round', lineCap:'round', fill:false })
      });

      const outline = L.geoJSON({ type:'FeatureCollection', features }, {
        interactive:false,
        style:()=>({ color, weight:(cfg.drivers?.strokeWeightPx ?? 4), opacity:1.0, lineJoin:'round', lineCap:'round', dashArray:(cfg.drivers?.dashArray ?? '6 4'), fill:false })
      });

      const group = L.featureGroup([halo, outline]).addTo(map);
      let labelMarker = null;
      try {
        const b = outline.getBounds();
        const center = b && b.getCenter ? b.getCenter() : null;
        if (center) {
          labelMarker = L.marker(center, { opacity: 0 })
            .bindTooltip(name, { permanent:true, direction:'center', className: (cfg.drivers?.labelClass || 'lbl dim') });
          group.addLayer(labelMarker);
        }
      } catch {}

      try { outline.bringToFront(); } catch {}
      driverOverlays[name] = { group, color, labelMarker };
    });

    renderDriversPanel(driverMeta, driverOverlays, true, driverSelectedCounts, custWithinSel);
  }

  // exact key, then base key
  function lookupDriverForKey(key) {
    if (!key) return null;
    if (activeAssignMap[key]) return activeAssignMap[key];
    const m = String(key).match(/^([WTFS]\d+)/);
    if (m && activeAssignMap[m[1]]) return activeAssignMap[m[1]];
    return null;
  }

  // =================================================================
  // POPUP / FOCUS / STYLES
  // =================================================================
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

  function focusFeature(lyr) {
    if (currentFocus && currentFocus !== lyr) restoreFeature(currentFocus);
    currentFocus = lyr;

    const perDay = lyr._perDay || {};
    const baseWeight = lyr._isSelected ? (cfg.style?.selected?.weightPx ?? 2) : (cfg.style?.dimmed?.weightPx ?? 1);
    const hiWeight = Math.max(3, Math.round(baseWeight * 3));

    lyr.setStyle({
      color: perDay.stroke || '#666',
      weight: hiWeight,
      opacity: 1.0,
      fillColor: perDay.fill || '#ccc',
      fillOpacity: lyr._isSelected ? (perDay.fillOpacity ?? 0.8) : dimFill(perDay, cfg)
    });

    showLabel(lyr, lyr._labelTxt);
    lyr.bringToFront?.();

    const b = lyr.getBounds?.();
    if (b) map.fitBounds(b.pad(0.2));
  }
  function restoreFeature(lyr) {
    const perDay = lyr._perDay || {};
    if (lyr._isSelected) { applyStyleSelected(lyr, perDay, cfg); showLabel(lyr, lyr._labelTxt); }
    else { applyStyleDim(lyr, perDay, cfg); hideLabel(lyr); }
  }
  function clearFocus(recenter) {
    if (!currentFocus) { if (recenter && hasSelection && selectionBounds) map.fitBounds(selectionBounds.pad(0.1)); return; }
    restoreFeature(currentFocus);
    currentFocus = null;
    if (recenter) {
      if (hasSelection && selectionBounds) map.fitBounds(selectionBounds.pad(0.1));
      else if (allBounds) map.fitBounds(allBounds.pad(0.1));
    }
  }
  function applyStyleSelected(lyr, perDay, cfg) {
    lyr.setStyle({
      color: perDay.stroke || '#666',
      weight: cfg.style?.selected?.weightPx ?? 2,
      opacity: cfg.style?.selected?.strokeOpacity ?? 1.0,
      fillColor: perDay.fill || '#ccc',
      fillOpacity: perDay.fillOpacity ?? 0.8
    });
  }
  function applyStyleDim(lyr, perDay, cfg) {
    lyr.setStyle({
      color: perDay.stroke || '#666',
      weight: cfg.style?.dimmed?.weightPx ?? 1,
      opacity: cfg.style?.dimmed?.strokeOpacity ?? 0.35,
      fillColor: perDay.fill || '#ccc',
      fillOpacity: dimFill(perDay, cfg)
    });
  }

  // =================================================================
  // UI PANELS (kept)
  // =================================================================
  function renderDriversPanel(metaList, overlays, defaultOn=false, countsMap={}, totalSelected=0) {
    const el = document.getElementById('drivers'); if (!el) return;

    const metaByName = new Map((metaList||[]).map(d => [String(d.name||'').toLowerCase(), d]));
    const overlayDrivers = Object.keys(overlays || {});
    const haveOverlays = overlayDrivers.length > 0;

    const allDriverNames = haveOverlays
      ? overlayDrivers
      : Array.from(new Set(Object.values(activeAssignMap))).sort((a,b)=>a.toLowerCase().localeCompare(b.toLowerCase()));

    const orderIndex = new Map((metaList||[]).map((d,i)=>[String(d.name||'').toLowerCase(), i]));
    allDriverNames.sort((a,b)=>{
      const ia = orderIndex.has(a.toLowerCase()) ? orderIndex.get(a.toLowerCase()) : 9999;
      const ib = orderIndex.has(b.toLowerCase()) ? orderIndex.get(b.toLowerCase()) : 9999;
      if (ia !== ib) return ia - ib;
      return a.toLowerCase().localeCompare(b.toLowerCase());
    });

    const rows = allDriverNames.map(name => {
      const meta = metaByName.get(name.toLowerCase()) || { name, color: colorFromName(name) };
      const col = meta.color || '#888';
      const safeName = escapeHtml(name || '');
      const isPresent = !!overlays?.[name];
      const onAttr = isPresent && defaultOn ? 'checked' : '';
      const count = typeof countsMap[name] === 'number' ? countsMap[name] : 0;
      const frac = `${count}/${totalSelected || 0}`;
      return `<div class="row">
        <input type="checkbox" data-driver="${safeName}" aria-label="Toggle ${safeName}" ${onAttr} ${isPresent ? '' : 'disabled'}>
        <span class="swatch" style="background:${col}; border-color:${col}"></span>
        <div>${safeName}</div>
        <div class="counts" style="margin-left:auto">${frac}</div>
      </div>`;
    }).join('');

    el.innerHTML = `<h4>Drivers</h4>${rows}`;

    el.querySelectorAll('input[type="checkbox"]').forEach(cb => {
      const name = cb.getAttribute('data-driver');
      if (!overlays?.[name]) return; // disabled
      toggleDriverOverlay(name, !!cb.checked);
      cb.addEventListener('change', (e) => {
        const name = e.target.getAttribute('data-driver');
        toggleDriverOverlay(name, e.target.checked);
      });
    });
  }

  function toggleDriverOverlay(name, on) {
    const rec = driverOverlays[name];
    if (!rec) return;
    if (on) {
      rec.group.addTo(map);
      try { if (rec.labelMarker) rec.labelMarker.openTooltip(); } catch{}
    } else {
      map.removeLayer(rec.group);
    }
  }

  function renderLegend(cfg, legendCounts, custIn, custOut, outsideToggle) {
    const el = document.getElementById('legend'); if (!el) return;

    const rowsHtml = (cfg.layers || []).map(Lcfg => {
      const st = (cfg.style?.perDay?.[Lcfg.day]) || {};
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

    const toggle = `<div class="row" style="margin-top:4px">
        <input type="checkbox" id="toggleOutside" ${outsideToggle ? 'checked' : ''} aria-label="Highlight outside customers">
        <div>Highlight outside customers</div>
      </div>`;

    el.innerHTML = `<h4>Layers</h4>${rowsHtml}${custBlock}${toggle}`;

    const cb = document.getElementById('toggleOutside');
    if (cb) {
      cb.addEventListener('change', () => {
        // left as-is; snapshot ignores this DOM-only option
        // (customers outside selection are hidden in focused mode anyway)
      });
    }
  }

  // =================================================================
  // HELPERS / PARSERS / UTILS
  // =================================================================
  function cb(u) { return (u.includes('?') ? '&' : '?') + 'cb=' + Date.now(); }
  async function fetchJson(url) { const res = await fetch(url + cb(url)); if (!res.ok) throw new Error(`Fetch ${res.status} for ${url}`); return res.json(); }
  async function fetchText(url) { const res = await fetch(url + cb(url)); if (!res.ok) throw new Error(`Fetch ${res.status} for ${url}`); return res.text(); }
  function sleep(ms){ return new Promise(r => setTimeout(r, ms)); }

  function parseBatchItems(b64){ if (!b64) return []; try { const json = atob(String(b64).replace(/-/g,'+').replace(/_/g,'/')); const arr = JSON.parse(json); return Array.isArray(arr) ? arr : []; } catch { return []; } }

  function parseCsvRows(text) {
    const out = []; let i=0, f='', r=[], q=false;
    text = String(text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    const N = text.length;
    const pushF=()=>{ r.push(f); f=''; }, pushR=()=>{ out.push(r); r=[]; };
    while (i<N) {
      const c = text[i];
      if (q) {
        if (c === '"') { if (i+1<N && text[i+1] === '"') { f+='"'; i+=2; } else { q=false; i++; } }
        else { f+=c; i++; }
      } else {
        if (c === '"') { q=true; i++; }
        else if (c === ',') { pushF(); i++; }
        else if (c === '\n') { pushF(); pushR(); i++; }
        else { f+=c; i++; }
      }
    }
    pushF(); pushR();
    return out.map(row => row.map(v => (v||'').replace(/^\uFEFF/, '').trim()));
  }

  function findHeaderFlexible(rows, schema) {
    if (!rows?.length) return null;
    const wantKeys = ((schema && schema.keys) || 'zone keys').toLowerCase();
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

  function findCustomerHeaderIndex(rows, schema) {
    const wantCoords = ((schema && schema.coords) || 'Verified Coordinates').toLowerCase();
    const wantNote   = ((schema && schema.note)   || 'Order Note').toLowerCase();
    for (let i=0;i<rows.length;i++) {
      const hdr = rows[i] || [];
      const hasCoords = hdr.some(h => (h||'').toLowerCase() === wantCoords);
      const hasNote   = hdr.some(h => (h||'').toLowerCase() === wantNote);
      if (hasCoords || hasNote) return i;
    }
    return -1;
  }
  function headerIndexMap(hdrRow, schema) {
    const wantCoords = ((schema && schema.coords) || 'Verified Coordinates').toLowerCase();
    const wantNote   = ((schema && schema.note)   || 'Order Note').toLowerCase();
    const idx = (name) => hdrRow.findIndex(h => (h||'').toLowerCase() === name);
    return { coords: idx(wantCoords), note: idx(wantNote) };
  }

  function splitKeys(s) { return String(s||'').split(/[;,/|]/).map(x => x.trim()).filter(Boolean); }
  function normalizeKey(s) {
    s = String(s || '').trim().toUpperCase();
    const m = s.match(/^([WTFS])0*(\d+)(_.+)?$/);
    if (m) return m[1] + String(parseInt(m[2], 10)) + (m[3] || '');
    return s;
  }
  function baseKeyFrom(key) { const m = String(key||'').toUpperCase().match(/^([WTFS]\d+)/); return m ? m[1] : String(key||'').toUpperCase(); }
  function quadParts(key) {
    const m = String(key || '').toUpperCase().match(/_(NE|NW|SE|SW)(?:_(TL|TR|LL|LR))?$/);
    return m ? { quad: m[1], sub: m[2] || null } : null;
  }
  function basePlusQuad(key) { const p = quadParts(key); return p ? (baseKeyFrom(key) + '_' + p.quad) : null; }
  function isSubquadrantKey(k) { const p = quadParts(k); return !!(p && p.sub); }
  function isQuadrantKey(k)    { const p = quadParts(k); return !!(p && !p.sub); }

  function parseLatLng(s) {
    const m = String(s||'').trim().match(/(-?\d+(?:\.\d+)?)\s*[, ]\s*(-?\d+(?:\.\d+)?)/);
    if (!m) return null;
    const lat = parseFloat(m[1]), lng = parseFloat(m[2]);
    if (!isFinite(lat) || !isFinite(lng)) return null;
    if (lat < -90 || lat > 90 || lng < -180 || lng > 180) return null;
    return { lat, lng };
  }

  function dimFill(perDay, cfg) {
    const base = perDay.fillOpacity ?? 0.8;
    const factor = cfg.style?.dimmed?.fillFactor ?? 0.3;
    return Math.max(0.08, base * factor);
  }

  function showLabel(lyr, text) { if (lyr.getTooltip()) lyr.unbindTooltip(); lyr.bindTooltip(text, { permanent: true, direction: 'center', className: 'lbl' }); }
  function hideLabel(lyr) { if (lyr.getTooltip()) lyr.unbindTooltip(); }

  function escapeHtml(s) {
    return String(s || '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

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

  function showTopError(title, msg) {
    const el = document.getElementById('error');
    if (!el) return;
    el.style.display = 'block';
    el.innerHTML = `<strong>${escapeHtml(title)}</strong><br>${escapeHtml(String(msg || ''))}`;
  }
  function renderError(msg) { showTopError('Load error', msg); }
  function phase(msg){ setStatus(`⏳ ${msg}`); }
  function info(msg){ setStatus(`ℹ️ ${msg}`); }
  function warn(msg){ showTopError('Warning', msg); }

  // Status line now includes "active zone keys"
  function setStatus(msg) { const n = document.getElementById('status'); if (n) n.textContent = msg || ''; }
  function makeStatusLine(selMunis, inCount, outCount, activeKeysLowerArray) {
    const muniList = (Array.isArray(selMunis) && selMunis.length) ? selMunis.join(', ') : '—';
    const keysTxt = (Array.isArray(activeKeysLowerArray) && activeKeysLowerArray.length)
      ? activeKeysLowerArray.join(', ')
      : '—';
    return `Customers (in/out): ${inCount}/${outCount} • Municipalities: ${muniList} • active zone keys: ${keysTxt}`;
  }

  function extractAssignmentsFromCsv(rows, knownNamesSet) {
    const like = (s) => String(s||'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
    const dayWords = new Set(['mon','monday','tue','tues','tuesday','wed','weds','wednesday','thu','thur','thurs','thursday','fri','friday','sat','saturday','sun','sunday']);
    const looksLikeName = (s) => {
      const v = like(s);
      if (!v) return false;
      if (dayWords.has(v)) return false;
      if (knownNamesSet && knownNamesSet.size) return knownNamesSet.has(v);
      return /^[a-z][a-z .'\-]{1,30}$/.test(v) && !/\d/.test(v);
    };
    const isKeysHeader = (s) => { const v = like(s); return v.includes('key') || v.includes('zone') || v.includes('route'); };
    const isDriverHeader = (s) => { const v = like(s); return v.includes('driver') || v === 'name' || v.includes('assigned'); };

    // Try header-based table
    let headerIndex=-1, driverCol=-1, keysCol=-1;
    for (let i=0; i<Math.min(rows.length, 50); i++) {
      const row = rows[i] || [];
      let d=-1, k=-1;
      for (let c=0;c<row.length;c++) {
        if (d === -1 && isDriverHeader(row[c])) d = c;
        if (k === -1 && isKeysHeader(row[c]))   k = c;
      }
      if (d !== -1 && k !== -1) { headerIndex = i; driverCol = d; keysCol = k; break; }
    }

    const out = {};
    if (headerIndex !== -1) {
      for (let r = headerIndex + 1; r < rows.length; r++) {
        const row = rows[r] || [];
        const dn = (row[driverCol] || '').trim();
        const ks = (row[keysCol]   || '').trim();
        if (!dn || !ks) continue;
        if (!looksLikeName(dn)) continue; // reject “Sat”
        splitKeys(ks).map(normalizeKey).forEach(k => { out[k] = dn; });
      }
      return out;
    }

    // Fallback: scan first 40 rows; only accept KNOWN names
    for (let r = 0; r < Math.min(rows.length, 40); r++) {
      const row = rows[r] || [];
      if (!row.length) continue;

      let dn = '';
      for (let c=0;c<row.length;c++) {
        const cell = (row[c]||'').trim();
        if (!cell) continue;
        if (looksLikeName(cell)) { dn = cell; break; }
      }
      if (!dn) continue;

      // find any cell(s) with keys and merge
      let allKeys = [];
      for (let c=0;c<row.length;c++) {
        const cell = (row[c]||'').trim();
        if (/[WTFS]\s*0*\d+(_(NE|NW|SE|SW)(_(TL|TR|LL|LR))?)?/i.test(cell)) {
          allKeys = allKeys.concat(splitKeys(cell).map(normalizeKey));
        }
      }
      allKeys = Array.from(new Set(allKeys)).filter(Boolean);
      allKeys.forEach(k => { out[k] = dn; });
    }
    return out;
  }

  function ensureDriverMeta(metaList, driverNames) {
    const out = Array.isArray(metaList) ? [...metaList] : [];
    const seen = new Set(out.map(d => String(d.name||'').toLowerCase()));
    for (const nm of (driverNames || [])) {
      const low = String(nm||'').toLowerCase(); if (!low) continue;
      if (!seen.has(low)) { out.push({ name: nm, color: colorFromName(nm) }); seen.add(low); }
    }
    return out;
  }

  function colorFromName(name) {
    let h=0; const s=65, l=50; const str = String(name||'');
    for (let i=0;i<str.length;i++) h = (h*31 + str.charCodeAt(i)) % 360;
    return `hsl(${h}, ${s}%, ${l}%)`;
  }

})().catch(e => {
  const el = document.getElementById('error');
  if (el) {
    el.style.display = 'block';
    el.innerHTML = `<strong>Load error</strong><br>${String(e && e.message ? e.message : e)}`;
  }
  console.error(e);
});
