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

    const dayLayers = [];            // [{day, layer, perDay}]
    const boundaryFeatures = [];     // for clipping color tiles
    let allBounds = null;            // union of all features
    let selectionBounds = null;      // union of selected features (last apply)
    let hasSelection = false;
    let totalFeatures = 0;
    let currentFocus = null;         // lyr we zoomed into via click (if any)

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
          const muni = smartTitleCase(muniRaw);            // << label text you asked for
          const day  = (p[cfg.fields.day]  ?? Lcfg.day).toString().trim();

          lyr._routeKey = key;
          lyr._day      = day;
          lyr._perDay   = perDay;
          lyr._labelTxt = muni;                            // store the pretty label

          // start with NO label bound; we attach/detach when (de)selected or focused

          // click-to-focus
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
    } catch(e) { /* ok to ignore */ }

    // click outside to clear focus and refit to selection
    map.on('click', () => {
      if (currentFocus) clearFocus();
      else if (hasSelection && selectionBounds) map.fitBounds(selectionBounds.pad(0.1));
    });

    await applySelection(); // center on selected at launch

    const refresh = Number(cfg.behavior.refreshSeconds || 0);
    if (refresh > 0) setInterval(applySelection, refresh * 1000);

    async function applySelection() {
      // any change to selection cancels a manual focus
      if (currentFocus) clearFocus();

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

      // counts & styling
      const counts = {};
      let selectedCount = 0;
      const selBounds = [];

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
            showLabel(lyr, lyr._labelTxt);
            if (lyr.getBounds) selBounds.push(lyr.getBounds());
            selectedCount++;
            layer.addLayer(lyr);
          } else {
            if (cfg.style.unselectedMode === 'hide' && hasSelection) {
              hideLabel(lyr);
              layer.removeLayer(lyr);
            } else {
              lyr.setStyle({
                color: perDay.stroke || '#666',
                weight: cfg.style.dimmed.weightPx,
                opacity: cfg.style.dimmed.strokeOpacity,
                fillColor: perDay.fill || '#ccc',
                fillOpacity: dimFill(perDay, cfg)
              });
              if (currentFocus === lyr) showLabel(lyr, lyr._labelTxt); // focused unselected stays labeled
              else hideLabel(lyr);                                     // otherwise no labels for unselected
              layer.addLayer(lyr);
            }
          }
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

    // ---- focus helpers ----
    function focusFeature(lyr) {
      currentFocus = lyr;
      showLabel(lyr, lyr._labelTxt);
      lyr.bringToFront?.();
      const b = lyr.getBounds?.();
      if (b) map.fitBounds(b.pad(0.2));
    }
    function clearFocus() {
      if (!currentFocus) return;
      // if that feature is not selected, hide its label again
      if (!hasSelection ||
          (hasSelection && !featureIsSelected(currentFocus))) {
        hideLabel(currentFocus);
      }
      currentFocus = null;
      if (hasSelection && selectionBounds) map.fitBounds(selectionBounds.pad(0.1));
      else if (allBounds) map.fitBounds(allBounds.pad(0.1));
    }
    function featureIsSelected(lyr) {
      // we infer from current style: selected has higher weight
      return lyr.options && lyr.options.weight === cfg.style.selected.weightPx;
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
  // find header row including keys (and day if present)
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

  // label helpers: permanent tooltip on/off
  function showLabel(lyr, text) {
    // rebind to ensure 'permanent' is applied
    if (lyr.getTooltip()) lyr.unbindTooltip();
    lyr.bindTooltip(text, { permanent: true, direction: 'center', className: 'lbl' });
  }
  function hideLabel(lyr) {
    const t = lyr.getTooltip();
    if (t) lyr.unbindTooltip();
  }

  // smart title case for MUNICIPALITY NAMES
  function smartTitleCase(s) {
    if (!s) return '';
    s = s.toLowerCase();
    // split by slashes or hyphens but keep the separators
    const parts = s.split(/([\/-])/g);
    const small = new Set(['and','or','of','the','a','an','in','on','at','by','for','to','de','la','le','el','du','von','van','di','da','del']);
    const fixWord = (w, isFirst) => {
      if (!w) return w;
      if (!isFirst && small.has(w)) return w;
      // St. / Mt. abbreviations
      if (w === 'st' || w === 'st.' ) return 'St.';
      if (w === 'mt' || w === 'mt.' ) return 'Mt.';
      // normal cap
      return w.charAt(0).toUpperCase() + w.slice(1);
    };
    // rebuild with title-casing around separators
    let tokenIdx = 0;
    for (let i=0;i<parts.length;i++){
      if (parts[i] === '/' || parts[i] === '-') continue;
      const tokens = parts[i].split(/\s+/).map((tok, j) => fixWord(tok, tokenIdx++ === 0));
      parts[i] = tokens.join(' ');
    }
    return parts.map(p => p === '/' ? '/' : (p === '-' ? '-' : p)).join('').replace(/\s+/g,' ').trim();
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
