/* apps.js — 3-tier (base → quadrants → subquadrants) splice-aware viewer + batch snapshots
   - Loads base, quadrant, and subquadrant layers (if present in config).
   - Visibility precedence (per base): Subquadrant > Quadrant > Base.
     * Base (e.g., W1) shows only if selected AND no quads/subquads of that base are selected.
     * A Quadrant (e.g., W1_NE) shows only if selected AND no subquads exist for that quadrant (W1_NE_*).
     * Subquadrants (e.g., W1_SW_TL) show if selected.
   - Exact "active zone keys" shown in status, using the selection order and only keys actually drawn.
   - All existing features kept: outside-customer toggle, never-blank fallback, driver overlays, labels, CSV parsing, counts.
   - NEW: Batch snapshots — pass ?batch=<websafe b64 JSON> with items [{day,driver,name,keys[],outName?,stats:{deliveries,baseBoxes,customs,addOns,apartments}}].
          Also pass ?cb=<Apps Script Web App URL> so the page can POST the PNG to Drive.
*/
(async function () {
  // ---- Surface uncaught errors to the page error box ----
  window.addEventListener('error', (e) => showTopError('Script error', (e && e.message) ? e.message : String(e)));
  window.addEventListener('unhandledrejection', (e) => showTopError('Promise error', (e?.reason?.message) || String(e?.reason || e)));

  // ---- URL / config ----
  const qs = new URLSearchParams(location.search);
  const cfgUrl = qs.get('cfg') || './config/app.config.json';

  const parseJSON = (s, fb=null) => { try { return JSON.parse(s); } catch { return fb; } };
  const driverMetaParam = parseJSON(qs.get('driverMeta') || '[]', []);   // [{name,color}, ...]
  const assignMapParam  = parseJSON(qs.get('assignMap')  || '{}', {});   // {"S9":"Devin","S9_NE_TL":"Devin",...}

  // --- Snapshot/batch globals (NEW) ---
  const cbUrl = qs.get('cb') || '';          // Apps Script receiver (doPost Web App)
  let manualSelectedKeys = null;             // override selection for batch
  const batchParam = qs.get('batch');        // websafe base64 JSON array from launcher

  let activeAssignMap = { ...assignMapParam };     // authoritative mapping after CSV merge
  let driverMeta      = Array.isArray(driverMetaParam) ? [...driverMetaParam] : [];
  let highlightOutside = false;

  // ---- map init ----
  phase('Loading config…');
  const cfg = await fetchJson(cfgUrl);

  if (typeof L === 'undefined' || !document.getElementById('map')) {
    renderError('Leaflet or #map element not available. Check <link/script> tags and container element.');
    return;
  }

  phase('Initializing map…');
  const map = L.map('map', { preferCanvas: true }).setView([43.55, -80.25], 8);
  const osmTiles = 'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png';
  const base = L.tileLayer(osmTiles, {
    attribution: '&copy; OpenStreetMap contributors',
    crossOrigin: true                // NEW: needed for leaflet-image canvas export
  }).addTo(map);
  try {
    const el = base.getContainer?.();
    if (el) { el.style.filter = 'grayscale(1)'; el.style.webkitFilter = 'grayscale(1)'; }
  } catch {}

  // ---- collections / state ----
  const baseDayLayers = [];   // [{day, layer, perDay, features: LeafletLayer[]}]  (W1, W2, F12…)
  const quadDayLayers = [];   // same for quadrants (…_NE, …_SW…)
  const subqDayLayers = [];   // same for subquadrants (…_NE_TL, …_SW_LR…)
  const allDaySets    = [baseDayLayers, quadDayLayers, subqDayLayers];

  let allBounds = null;
  let selectionBounds = null;
  let hasSelection = false;

  const boundaryFeatures = [];
  const customerLayer = L.layerGroup().addTo(map);
  const customerMarkers = [];      // [{marker, lat, lng}]
  let customerCount = 0;
  let custWithinSel = 0;
  let custOutsideSel = 0;
  const custByDayInSel = { Wednesday: 0, Thursday: 0, Friday: 0, Saturday: 0 };

  let currentFocus = null;
  let coveragePolysAll = [];       // visible polygons (all)
  let coveragePolysSelected = [];  // visible polygons in current selection
  let selectedMunicipalities = [];

  let driverSelectedCounts = {};   // { "Devin": n, ... } (selected only)
  const driverOverlays = {};       // name -> { group, color, labelMarker }

  // For status line: preserve selection ORDER + show only keys actually drawn
  let selectedOrderedKeys = [];     // normalized (upper) in order parsed from CSV or manual override
  let visibleSelectedKeysSet = new Set(); // keys that actually rendered due to precedence

  // Render empty panels first
  renderLegend(cfg, null, custWithinSel, custOutsideSel, highlightOutside);
  renderDriversPanel([], {}, false, {}, 0);

  // ---- load polygon layers (base → quads → subquads) ----
  phase('Loading polygon layers…');
  await loadLayerSet(cfg.layers, baseDayLayers, true); // required
  if (cfg.layersQuadrants && cfg.layersQuadrants.length) {
    await loadLayerSet(cfg.layersQuadrants, quadDayLayers, true);
  }
  if (cfg.layersSubquadrants && cfg.layersSubquadrants.length) {
    await loadLayerSet(cfg.layersSubquadrants, subqDayLayers, true);
  }

  // Optional boundary-canvas mask
  try {
    if (L.TileLayer && L.TileLayer.boundaryCanvas && boundaryFeatures.length) {
      const boundaryFC = { type: 'FeatureCollection', features: boundaryFeatures };
      const maskTiles = L.TileLayer.boundaryCanvas(osmTiles, { boundary: boundaryFC, attribution: '' });
      maskTiles.addTo(map).setZIndex(2);
      base.setZIndex(1);
    }
  } catch (e) { console.warn('boundaryCanvas unavailable (non-fatal):', e); }

  map.on('click', () => { clearFocus(true); });
  map.on('movestart', () => { clearFocus(false); map.closePopup(); });
  map.on('zoomstart', () => { map.closePopup(); });

  // ---- initial selection + customers ----
  await applySelection();       // never blank
  await loadCustomersIfAny();

  // optional auto-refresh
  const refresh = Number(cfg.behavior?.refreshSeconds || 0);
  if (refresh > 0) setInterval(async () => {
    await applySelection();
    await loadCustomersIfAny();
  }, refresh * 1000);

  // Kick off batch snapshots if the URL contained ?batch=...
  await runBatchIfRequested();

  // =================================================================
  // LOADERS
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

          lyr._routeKey   = keyNorm;                 // exact key: S9 / S9_NE / S9_NE_TL
          lyr._baseKey    = baseKeyFrom(keyNorm);    // base: S9
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

  // =================================================================
  // SELECTION (base + quadrants + subquadrants) + CSV-derived assignments
  // =================================================================
  async function applySelection() {
    clearFocus(false);

    // Build selection (CSV or manual override for batch)
    const selectedSet = new Set();
    selectedOrderedKeys = [];
    let loadedSelectionCsv = false;

    if (Array.isArray(manualSelectedKeys) && manualSelectedKeys.length) {
      // Batch mode: explicit keys
      manualSelectedKeys.map(normalizeKey).forEach(k => {
        if (!selectedSet.has(k)) { selectedSet.add(k); selectedOrderedKeys.push(k); }
      });
    } else {
      // Normal mode: CSV-based
      const selUrl = qs.get('sel') || (cfg.selection && cfg.selection.url) || '';
      if (!selUrl) {
        warn('No selection URL configured — showing all zones.');
      } else {
        phase('Loading selection…');
        try {
          const csvText = await fetchText(selUrl);
          const rowsAA = parseCsvRows(csvText);
          loadedSelectionCsv = rowsAA?.length > 0;

          // A) parse selection keys from “zone keys”
          const hdr = findHeaderFlexible(rowsAA, cfg.selection?.schema || { keys: 'zone keys' });
          if (hdr) {
            const { headerIndex, keysCol } = hdr;
            for (let i = headerIndex + 1; i < rowsAA.length; i++) {
              const r = rowsAA[i]; if (!r || r.length === 0) continue;
              const ks = splitKeys(r[keysCol]).map(normalizeKey);
              for (const k of ks) {
                if (!selectedSet.has(k)) { selectedSet.add(k); selectedOrderedKeys.push(k); }
              }
            }
          }

          // B) parse driver assignments (ONLY accept known driver names if provided)
          const knownDriverNames = new Set(
            (driverMeta || []).map(d => String(d.name || '').toLowerCase()).filter(Boolean)
          );
          const derivedAssign = extractAssignmentsFromCsv(rowsAA, knownDriverNames);
          activeAssignMap = { ...activeAssignMap, ...derivedAssign };  // CSV overwrites seed
          driverMeta = ensureDriverMeta(driverMeta, Object.values(activeAssignMap));
        } catch (e) {
          warn(`Selection CSV load failed (${e && e.message ? e.message : e}). Showing all zones.`);
        }
      }
    }

    const noKeys = selectedSet.size === 0;
    hasSelection = !noKeys;
    selectionBounds = null;
    coveragePolysSelected = [];
    selectedMunicipalities = [];
    visibleSelectedKeysSet = new Set();

    // Helper: record keys that truly rendered as selected
    const recordVisibleSelectedKey = (k) => { if (k) visibleSelectedKeysSet.add(k); };

    // Override tracking:
    //  - bases with any selected quadrants
    //  - bases with any selected subquadrants
    //  - specific quadrants (e.g., "W1_SW") that have selected subquadrants
    const quadBasesSelected = new Set();
    const subqBasesSelected = new Set();
    const subqQuadsSelected = new Set(); // strings like "W1_NE"

    if (!noKeys) {
      for (const k of selectedSet) {
        if (isSubquadrantKey(k)) {
          subqBasesSelected.add(baseKeyFrom(k));
          const pq = basePlusQuad(k);
          if (pq) subqQuadsSelected.add(pq);
        } else if (isQuadrantKey(k)) {
          quadBasesSelected.add(baseKeyFrom(k));
        }
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
          if (lyr.getBounds) {
            const b = lyr.getBounds();
            selectionBounds = selectionBounds ? selectionBounds.extend(b) : L.latLngBounds(b);
          }
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
      for (const arr of allDaySets) {
        for (const entry of arr) {
          for (const lyr of entry.features) {
            setFeatureVisible(entry, lyr, true, false); // visible, not "selected"
          }
        }
      }
      rebuildCoverageFromVisible();
      selectedMunicipalities = [];
      if (cfg.behavior?.autoZoom) {
        if (allBounds) map.fitBounds(allBounds.pad(0.1));
      }
    } else {
      // BASE: show only if base key selected AND no quad/subquad selected for that base
      for (const entry of baseDayLayers) {
        for (const lyr of entry.features) {
          const keyBase = lyr._routeKey; // base layer uses base key (e.g., W1)
          const visible = selectedSet.has(keyBase) && !quadBasesSelected.has(keyBase) && !subqBasesSelected.has(keyBase);
          setFeatureVisible(entry, lyr, visible, visible);
        }
      }
      // QUADRANTS: show only if exact quad selected AND no subquads selected for that specific quadrant
      for (const entry of quadDayLayers) {
        for (const lyr of entry.features) {
          const kQuad = lyr._routeKey;      // e.g., W1_NE
          const bq = basePlusQuad(kQuad);   // "W1_NE"
          const visible = selectedSet.has(kQuad) && !subqQuadsSelected.has(bq);
          setFeatureVisible(entry, lyr, visible, visible);
        }
      }
      // SUBQUADRANTS: show only if exact subquad selected
      for (const entry of subqDayLayers) {
        for (const lyr of entry.features) {
          const kSub = lyr._routeKey; // e.g., W1_NE_TL
          const visible = selectedSet.has(kSub);
          setFeatureVisible(entry, lyr, visible, visible);
        }
      }

      rebuildCoverageFromVisible();
      selectedMunicipalities = Array.from(new Set(selectedMunicipalities))
        .sort((a, b) => a.toLowerCase().localeCompare(b.toLowerCase()));

      if (cfg.behavior?.autoZoom) {
        if (selectionBounds) map.fitBounds(selectionBounds.pad(0.1));
        else if (allBounds) map.fitBounds(allBounds.pad(0.1));
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

    // Build "active zone keys" list in original selection order, but only keys that actually drew
    const activeKeysOrderedLower = selectedOrderedKeys
      .filter(k => visibleSelectedKeysSet.has(k))
      .map(k => k.toLowerCase());

    renderLegend(cfg, legendCounts, custWithinSel, custOutsideSel, highlightOutside);
    setStatus(makeStatusLine(selectedMunicipalities, custWithinSel, custOutsideSel, activeKeysOrderedLower));
    await rebuildDriverOverlays();
  }

  function rebuildCoverageFromVisible() {
    coveragePolysAll = [];
    for (const arr of allDaySets) {
      for (const entry of arr) {
        for (const lyr of entry.features) {
          if (entry.layer.hasLayer(lyr)) {
            coveragePolysAll.push({ feat: lyr._turfFeat, perDay: entry.perDay, layerRef: lyr });
            lyr._custAny = 0; lyr._custSel = 0;
          }
        }
      }
    }
  }

  // =================================================================
  // CUSTOMERS
  // =================================================================
  async function loadCustomersIfAny() {
    customerLayer.clearLayers();
    customerMarkers.length = 0;
    customerCount = 0;

    const custToggle = qs.get('cust');
    if (custToggle && custToggle.toLowerCase() === 'off') {
      custWithinSel = custOutsideSel = 0;
      resetDayCounts();
      renderLegend(cfg, null, custWithinSel, custOutsideSel, highlightOutside);
      setStatus(makeStatusLine(selectedMunicipalities, custWithinSel, custOutsideSel, []));
      renderDriversPanel(driverMeta, driverOverlays, true, driverSelectedCounts, custWithinSel);
      return;
    }
    if (!cfg.customers || cfg.customers.enabled === false || !cfg.customers.url) {
      custWithinSel = custOutsideSel = 0;
      resetDayCounts();
      renderLegend(cfg, null, custWithinSel, custOutsideSel, highlightOutside);
      setStatus(makeStatusLine(selectedMunicipalities, custWithinSel, custOutsideSel, []));
      renderDriversPanel(driverMeta, driverOverlays, true, driverSelectedCounts, custWithinSel);
      return;
    }

    phase('Loading customers…');
    const text = await fetchText(cfg.customers.url);
    const rows = parseCsvRows(text);
    if (!rows.length) {
      custWithinSel = custOutsideSel = 0;
      resetDayCounts();
      renderLegend(cfg, null, custWithinSel, custOutsideSel, highlightOutside);
      setStatus(makeStatusLine(selectedMunicipalities, custWithinSel, custOutsideSel, []));
      renderDriversPanel(driverMeta, driverOverlays, true, driverSelectedCounts, custWithinSel);
      return;
    }

    const hdrIdx = findCustomerHeaderIndex(rows, cfg.customers.schema || { coords: 'Verified Coordinates', note: 'Order Note' });
