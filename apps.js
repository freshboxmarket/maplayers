// apps.js — vNext READY (finished)
// - Works with ?cfg, ?assignMap, ?driverMeta, ?batch (websafe b64), ?cb (webhook),
//   optional ?manual=1, ?diag=1, and Drive hints (?driveDirect=1&gClientId=...&driveFolderId|driveFolderUrl|driveFolder).
//
// Notes
// - Prevents blank shots by rendering the map via leaflet-image whenever the frame touches the map.
// - Draws the banner stats as a single line (D/H/E/F/G texts) onto the canvas at its DOM position (no E–G headers).
// - Save dock shows storage readiness and returns a clickable Drive link.
// - Direct Drive uses Google Identity Services (no gapi client needed). Scope: drive.file.

(function () {

  // ---- Robust fetch helpers (top-level, single source of truth) ----
  async function fetchText(url) {
    const res = await fetch(url, { cache: 'no-store' });
    if (!res.ok) throw new Error(`HTTP ${res.status} for ${url}`);
    return await res.text();
  }
  async function fetchJson(url) {
    return JSON.parse(await fetchText(url));
  }

  function start() {
    (async function () {
      // ---------- base UI + error surfacing ----------
      ensureUiShell();
      injectBaseCss();
      // Hardened global error hooks so they can never crash the app
      window.addEventListener('error', (e) => { try { showTopError('Script error', e?.message || String(e)); } catch {} });
      window.addEventListener('unhandledrejection', (e) => { try { showTopError('Promise error', e?.reason?.message || String(e?.reason || e)); } catch {} });

      // ---------- URL/config ----------
      const qs = new URLSearchParams(location.search);
      const cfgUrl = qs.get('cfg') || './config/app.config.json';
      const parseJSON = (s, fb=null) => { try { return JSON.parse(s); } catch { return fb; } };

      const driverMetaParam = parseJSON(qs.get('driverMeta') || '[]', []);
      const assignMapParam  = parseJSON(qs.get('assignMap')  || '{}', {});
      const batchParam      = qs.get('batch');
      const manualMode      = (qs.get('manual') === '1') || (!!batchParam && qs.get('auto') !== '1');
      const diagMode        = qs.get('diag') === '1';
      const cbUrl           = qs.get('cb') || '';

      // Drive hints
      const DRIVE_FOLDER_DEFAULT_NAME = 'Screenshots';
      const driveDirect = qs.get('driveDirect') === '1';
      const gClientId   = qs.get('gClientId') || '';
      const driveFolderRaw = qs.get('driveFolder') || qs.get('driveFolderId') || qs.get('driveFolderUrl') || '';
      const driveFolderIdCandidate = extractDriveFolderId(driveFolderRaw);

      // Runtime state
      let batchItems = parseBatchItems(batchParam);
      let driverMeta = Array.isArray(driverMetaParam) ? [...driverMetaParam] : [];
      let activeAssignMap = { ...assignMapParam };
      let currentIndex = -1;
      let manualSelectedKeys = null;
      let runtimeCustEnabled = null;

      // Outside highlight: persisted
      const LS_KEYS = {
        frame: 'dispatchViewer.snapFrameRect',
        banner:'dispatchViewer.bannerPos',
        dock:  'dispatchViewer.snapDockPos',
        outside: 'dispatchViewer.highlightOutside',
        driveFolderCachePrefix: 'dispatchViewer.driveFolderId.'
      };
      let outsideHighlight = false;
      try { outsideHighlight = localStorage.getItem(LS_KEYS.outside) === '1'; } catch {}

      // ---------- fetch config ----------
      phase('Loading config…');
      const cfg = await fetchJson(cfgUrl);

      // ---------- map init ----------
      ensureMapRoot();
      const mapEl = document.getElementById('map');
      ensureMinHeight(mapEl);

      // FIX A: don't call undefined renderError; surface a safe error and exit
      if (typeof L === 'undefined' || !mapEl) { showTopError('Startup', 'Leaflet or #map missing.'); return; }

      phase('Initializing map…');
      const map = L.map('map', { preferCanvas: true }).setView([43.55, -80.25], 8);
      const osmTiles = 'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png';
      const base = L.tileLayer(osmTiles, {
        attribution: '&copy; OpenStreetMap contributors',
        crossOrigin: 'anonymous'        // important for taint prevention
      }).addTo(map);

      // Collections/state
      const baseDayLayers = [], quadDayLayers = [], subqDayLayers = [], allDaySets=[baseDayLayers,quadDayLayers,subqDayLayers];
      let allBounds=null, selectionBounds=null, hasSelection=false;
      const boundaryFeatures = [];

      // Customers
      const customerLayer = L.layerGroup().addTo(map);
      const customerMarkers = [];
      let customerCount=0, custWithinSel=0, custOutsideSel=0;
      const custByDayInSel = { Wednesday:0, Thursday:0, Friday:0, Saturday:0 };

      // Selection/focus
      let coveragePolysAll=[], coveragePolysSelected=[];
      let selectedMunicipalities=[];
      let selectedOrderedKeys = [];
      let visibleSelectedKeysSet = new Set();
      let driverSelectedCounts={}, driverOverlays={}, currentFocus=null;

      // Snap/dock/banner state
      let statsVisible = false;
      let snapArmed = false;
      let snapEls = null, dockEls = null;
      let lastPngDataUrl = null, lastSuggestedName = null;

      // ---------- scaffolding (now that cfg exists) ----------
      renderLegend(cfg, {}, 0, 0, outsideHighlight);
      renderDriversPanel([], {}, false, {}, 0);
      updateDiagnostics();

      // ---------- load layers ----------
      phase('Loading polygon layers…');
      await loadLayerSet(cfg.layers, baseDayLayers, true);
      if (cfg.layersQuadrants?.length)    await loadLayerSet(cfg.layersQuadrants,  quadDayLayers, true);
      if (cfg.layersSubquadrants?.length) await loadLayerSet(cfg.layersSubquadrants, subqDayLayers, true);

      // NEW: surface status once layers done
      if (totalFeatureCount() === 0) {
        warn('No polygon features loaded; check cfg URLs and CORS.');
        setStatus('No layers found — check cfg.');
      } else {
        setStatus('Layers loaded.');
      }

      // Optional boundary mask (plugin optional)
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
      setTimeout(()=>map.invalidateSize(), 50);

      // ---------- initial load ----------
      if (manualMode && batchItems.length) {
        manualSelectedKeys = unionAllKeys(batchItems);
        runtimeCustEnabled = false; // hide customers in overview
      }
      await applySelection();
      await loadCustomersIfAny();
      updateDiagnostics();

      // ---------- UI ----------
      if (manualMode && batchItems.length) {
        injectManualUI();
        ensureSnapshotUi();
        ensureDockUi();
        await zoomToOverview();
      } else if (!manualMode && batchItems.length) {
        await runAutoExport(batchItems);
      }

      // =================================================================
      // Controls / Toolbar / Banner
      // =================================================================
      function injectManualUI(){
        injectToolbarCss();

        // toolbar
        const bar = document.createElement('div');
        bar.className = 'route-toolbar';
        bar.innerHTML = `
          <button id="btnPrev" aria-label="Previous">◀ Prev</button>
          <button id="btnNext" aria-label="Next">Next ▶</button>
          <button id="btnStats" aria-label="Toggle stats">Stats</button>
          <button id="btnSnap" aria-label="Snapshot">📸 Snap</button>
        `;
        document.body.appendChild(bar);

        // banner (draggable) — two rows: meta + one-line stats
        const banner = document.createElement('div');
        banner.id = 'dispatchBanner';
        banner.className = 'dispatch-banner';
        banner.innerHTML = `<div class="row r1 meta"></div><div class="row r2 stats"></div>`;
        document.body.appendChild(banner);
        restoreBannerPos(banner);
        makeBannerDraggable(banner);

        document.getElementById('btnPrev').addEventListener('click', async () => { await stepRouteCycle(-1); });
        document.getElementById('btnNext').addEventListener('click', async () => { await stepRouteCycle(+1); });
        document.getElementById('btnStats').addEventListener('click', toggleStats);
        document.getElementById('btnSnap').addEventListener('click', onSnapClick);

        // keyboard
        window.addEventListener('keydown', async (e)=>{
          const tag = (e.target && e.target.tagName || '').toLowerCase();
          if (tag === 'input' || tag === 'textarea' || e.target?.isContentEditable) return;
          if (!batchItems.length) return;
          if (e.key === 'ArrowRight') { e.preventDefault(); await stepRouteCycle(+1); }
          else if (e.key === 'ArrowLeft') { e.preventDefault(); await stepRouteCycle(-1); }
          else if (e.key.toLowerCase() === 's') { e.preventDefault(); toggleStats(); }
          else if (e.key === 'Escape') { cancelFraming(); }
        });

        updateButtons();
      }

      function toggleStats(){
        statsVisible = !statsVisible;
        const banner = document.getElementById('dispatchBanner');
        if (!banner) return;
        banner.classList.toggle('visible', statsVisible);
        if (statsVisible && currentIndex >= 0) renderDispatchBanner(batchItems[currentIndex]);
      }

      // --- STAT TEXT FROM D/H/E/F/G as TEXT (no counters) ---
      function stringifyStat(val){
        if (val == null) return '';
        if (typeof val === 'number') return String(val);
        if (typeof val === 'string') return val.trim();
        if (Array.isArray(val)) return val.map(stringifyStat).filter(Boolean).join(' | ');
        if (typeof val === 'object') {
          const prefer = ['baseBoxesText','customsText','addOnsText','text','desc','description','value','label'];
          for (const k of prefer) if (val[k]) return stringifyStat(val[k]);
          const entries = Object.entries(val).map(([,v]) => stringifyStat(v)).filter(Boolean);
          return entries.join(' | ');
        }
        return String(val);
      }
      function trySumExpression(s){
        if (!s) return null;
        const t = String(s).trim();
        const strict = /^\s*\d+(?:\.\d+)?(?:\s*\+\s*\d+(?:\.\d+)?)*\s*$/;
        if (!strict.test(t)) return null;
        return t.split('+').map(x => Number(x.trim())).reduce((a,b)=>a+b,0) + '';
      }
      function normalizeQtyLike(s){ const str = stringifyStat(s); const summed = trySumExpression(str); return (summed || str || '—'); }
      function pickByKeysLike(obj, mustIncludes){
        const want = mustIncludes.map(s => s.toLowerCase());
        for (const [k,v] of Object.entries(obj||{})) {
          const lk = k.toLowerCase();
          if (want.every(w => lk.includes(w))) return v;
        }
        return undefined;
      }

      // REPLACED: getStatsTexts — treat E/F/G as TEXT; also accept D/E/F/G/H fallbacks
      function getStatsTexts(statsObj) {
        const s = statsObj || {};

        // Prefer “text” fields/arrays; accept letter fallbacks too
        const deliveriesRaw = s.deliveriesText ?? s.deliveries ?? s.D;
        const aptsRaw       = s.apartmentsText ?? s.apartments ?? s.H;

        const baseRaw = s.baseBoxesText ?? s.base ?? s.baseBoxes ?? s.E ?? pickByKeysLike(s, ['base','box']);
        const custRaw = s.customsText    ?? s.customs   ?? s.F   ?? pickByKeysLike(s, ['custom']);
        const addRaw  = s.addOnsText     ?? s.addOns    ?? s.add_ons ?? s.addons ?? s.G ?? pickByKeysLike(s, ['add']);

        const asText = (v) => stringifyStat(v) || '';

        return {
          deliveriesTxt: asText(deliveriesRaw),
          apartmentsTxt: asText(aptsRaw),
          baseTxt: asText(baseRaw),
          custTxt: asText(custRaw),
          addTxt:  asText(addRaw),
        };
      }
      
+      // Build one-line stats string in order D,H,E,F,G **with headers**.
+      // Preserves your “single-line” rule by collapsing newlines to spaces.
+      function buildStatsOneLiner(t) {
+        const seg = (label, val) => {
+          const v = String(val || '').trim();
+          if (!v || v === '—') return '';
+          const oneLine = v.replace(/\s*\n+\s*/g, ' ');
+          return `${label}: ${oneLine}`;
+        };
+        const parts = [
+          seg('# deliveries', t.deliveriesTxt),
+          seg('# apartments', t.apartmentsTxt),
+          seg('# base boxes', t.baseTxt),
+          seg('# customs',    t.custTxt),
+          seg('# add ons',    t.addTxt),
+        ].filter(Boolean);
+        return parts.join(' • ');
+      }


      function renderDispatchBanner(it){
        const banner = document.getElementById('dispatchBanner'); if (!banner) return;
        if (!it) { banner.classList.remove('visible'); return; }
        const s = it?.stats || {};
        const t = getStatsTexts(s);
        const row1 = `<strong>Day:</strong> ${escapeHtml(it.day||'')} &nbsp; - &nbsp; <strong>Driver:</strong> ${escapeHtml(it.driver||'')} &nbsp; - &nbsp; <strong>Route Name:</strong> ${escapeHtml(it.name||'')}`;
        const row2 = buildStatsOneLiner(t);
        const r1 = banner.querySelector('.r1');
        const r2 = banner.querySelector('.r2');
        if (r1) r1.innerHTML = row1;
        if (r2) r2.textContent = row2;
        if (statsVisible) banner.classList.add('visible');
      }

      // Draw the banner directly onto a snapshot canvas at its DOM position
      function drawBannerOntoCanvas(canvas, it, frameRect){
        if (!it || !statsVisible) return;
        const domBanner = document.getElementById('dispatchBanner');
        if (!domBanner || !domBanner.classList.contains('visible')) return;

        const s = it?.stats || {};
        const t = getStatsTexts(s);

        const row1 = `Day: ${it.day||''}  -  Driver: ${it.driver||''}  -  Route Name: ${it.name||''}`;
        const row2 = buildStatsOneLiner(t);

        const bRect = domBanner.getBoundingClientRect();
        const frameW = (typeof frameRect?.width === 'number' && frameRect.width > 0) ? Math.round(frameRect.width) : canvas.width;
        const dpr = canvas.width / Math.max(1, frameW);

        let x = Math.round((bRect.left - (frameRect.left || 0)) * dpr);
        let y = Math.round((bRect.top  - (frameRect.top  || 0)) * dpr);
        let boxW = Math.max(120 * dpr, Math.min(canvas.width, Math.round(bRect.width * dpr)));

        if (x > canvas.width || y > canvas.height || (x + 10) < 0 || (y + 10) < 0) return;
        x = Math.max(0, Math.min(x, canvas.width - 10));
        y = Math.max(0, Math.min(y, canvas.height - 10));

        const ctx = canvas.getContext('2d');
        const pad = Math.max(8 * dpr, Math.round(boxW * 0.04));
        const lineH = Math.max(16 * dpr, Math.round(Math.min(canvas.width, canvas.height) * 0.027));

        const fontFamily = 'system-ui, -apple-system, Segoe UI, Roboto, sans-serif';
        const fBold = (px)=>`700 ${px}px ${fontFamily}`;
        const fNorm = (px)=>`600 ${px}px ${fontFamily}`;

        const wrap = (text, fontPx) => {
          const font = fBold(fontPx);
          ctx.font = font;
          const words = String(text||'').split(/\s+/);
          const out = [];
          let cur='';
          for (const w of words){
            const test = cur ? cur + ' ' + w : w;
            if (ctx.measureText(test).width <= boxW - pad*2) cur = test;
            else { if (cur) out.push(cur); cur = w; }
          }
          if (cur) out.push(cur);
          return { lines: out, font };
        };

        const metaFontPx = Math.round(lineH*0.95);
        const metaWrapped = wrap(row1, metaFontPx);

        // Stats one-liner shrink-to-fit
        let statsFontPx = Math.round(lineH*0.90);
        ctx.font = fNorm(statsFontPx);
        let statsWidth = ctx.measureText(row2).width;
        const targetW = boxW - pad*2;
        while (statsWidth > targetW && statsFontPx > Math.max(10, Math.round(lineH*0.6))) {
          statsFontPx -= 1;
          ctx.font = fNorm(statsFontPx);
          statsWidth = ctx.measureText(row2).width;
        }

        const totalLines = metaWrapped.lines.length + 1; // meta lines + 1 stats line
        const boxH = pad*2 + totalLines * lineH + Math.round(lineH*0.2);

        ctx.fillStyle = 'rgba(255,255,255,0.92)'; roundRect(ctx, x, y, boxW, boxH, 12 * dpr).fill();

        let yy = y + pad + lineH; ctx.fillStyle='#111';
        ctx.font = metaWrapped.font; for (const ln of metaWrapped.lines) { ctx.fillText(ln, x+pad, yy); yy += lineH; }
        ctx.font = fNorm(statsFontPx); ctx.fillText(row2, x+pad, yy);
      }
      function roundRect(ctx, x, y, w, h, r) { ctx.beginPath(); ctx.moveTo(x + r, y); ctx.arcTo(x + w, y, x + w, y + h, r); ctx.arcTo(x + w, y + h, x, y + h, r); ctx.arcTo(x, y + h, x, y, r); ctx.arcTo(x, y, x + w, y, r); ctx.closePath(); return ctx; }

      // =================================================================
      // Route cycling
      // =================================================================
      async function stepRouteCycle(delta){
        const N = batchItems.length; if (!N) return;

        if (delta > 0) currentIndex = (currentIndex === -1) ? 0 : (currentIndex === N-1 ? -1 : currentIndex+1);
        else           currentIndex = (currentIndex === -1) ? N-1 : (currentIndex === 0 ? -1 : currentIndex-1);

        if (currentIndex === -1) { await zoomToOverview(); return; }

        const it = batchItems[currentIndex];
        manualSelectedKeys = (it.keys || []).map(normalizeKey);
        runtimeCustEnabled = true;
        await applySelection();
        await loadCustomersIfAny();
        if (selectionBounds) fitWithHints(selectionBounds, it?.view || null);
        renderDispatchBanner(it);
        updateButtons(); updateDiagnostics();
      }

      async function zoomToOverview(){
        currentIndex = -1;
        manualSelectedKeys = unionAllKeys(batchItems);
        runtimeCustEnabled = false;
        statsVisible = false; renderDispatchBanner(null);
        await applySelection();
        await loadCustomersIfAny();
        if (selectionBounds) fitWithHints(selectionBounds, null);
        updateButtons(); updateDiagnostics();
      }

      function updateButtons(){
        const haveFocus = currentIndex >= 0 && currentIndex < batchItems.length;
        const byId = id => document.getElementById(id);
        byId('btnPrev') && (byId('btnPrev').disabled = !batchItems.length);
        byId('btnNext') && (byId('btnNext').disabled = !batchItems.length);
        byId('btnStats') && (byId('btnStats').disabled = !haveFocus);
        byId('btnSnap')  && (byId('btnSnap').disabled  = !haveFocus);
      }

      // =================================================================
      // Snapshot: frame → wait tiles → capture → banner → save
      // =================================================================]
      function ensureSnapshotUi(){
        if (snapEls) return;
        const overlay = document.createElement('div');
        overlay.className = 'snap-overlay'; overlay.style.display='none';
        overlay.innerHTML = `
          <div class="helper">Drag to frame your capture. Click the <b>red Snap</b> to take it. &nbsp;•&nbsp; Press <b>Esc</b> to cancel.</div>
          <div class="frame">
            <div class="handle nw" data-dir="nw"></div><div class="handle n" data-dir="n"></div><div class="handle ne" data-dir="ne"></div>
            <div class="handle w" data-dir="w"></div> <div class="handle e" data-dir="e"></div>
            <div class="handle sw" data-dir="sw"></div><div class="handle s" data-dir="s"></div><div class="handle se" data-dir="se"></div>
          </div>
        `;
        document.body.appendChild(overlay);

        const flashCrop = document.createElement('div');
        flashCrop.className='snap-flash-crop';
        document.body.appendChild(flashCrop);

        snapEls = { overlay, helper: overlay.querySelector('.helper'), frame: overlay.querySelector('.frame'), flashCrop };
        bindFramingGestures(snapEls.frame);
      }

      function ensureDockUi(){
        if (dockEls) return;
        const dock = document.createElement('div');
        dock.className = 'snap-dock';
        dock.innerHTML = `
          <div class="head" id="snapDockHead">
            <div class="title">Snapshot</div>
            <div class="spacer"></div>
            <button class="x" id="snapDockClose" aria-label="Close">✕</button>
          </div>
          <div class="body">
            <img class="preview" id="snapDockImg" alt="Snapshot preview"/>
            <div class="row">
              <input type="text" name="snapName" id="snapDockName" placeholder="filename.png" />
            </div>
            <div class="row">
              <button class="btn" id="snapDockSave">Save</button>
              <button class="btn secondary" id="snapDockDownload">Download</button>
              <button class="btn secondary" id="snapDockExit">Exit</button>
            </div>
            <div class="link" id="snapDockLink"></div>
            <div class="note" id="snapDockNote"></div>
            <div class="note" id="snapDockTargets"></div>
          </div>
        `;
        document.body.appendChild(dock);

        const el = {
          dock,
          img:      dock.querySelector('#snapDockImg'),
          name:     dock.querySelector('#snapDockName'),
          saveBtn:  dock.querySelector('#snapDockSave'),
          dlBtn:    dock.querySelector('#snapDockDownload'),
          exitBtn:  dock.querySelector('#snapDockExit'),
          closeBtn: dock.querySelector('#snapDockClose'),
          title:    dock.querySelector('.title'),
          note:     dock.querySelector('#snapDockNote'),
          link:     dock.querySelector('#snapDockLink'),
          targets:  dock.querySelector('#snapDockTargets')
        };
        makeDockDraggable(dock, dock.querySelector('#snapDockHead')); restoreDockPos(dock);

        el.saveBtn.addEventListener('click', onSaveClick);
        el.dlBtn.addEventListener('click', onDownloadClick);
        const closeDock = ()=>{ dock.style.display='none'; lastPngDataUrl=null; lastSuggestedName=null; el.note.textContent=''; el.link.textContent=''; saveDockPos(dock); };
        el.exitBtn.addEventListener('click', closeDock);
        el.closeBtn.addEventListener('click', closeDock);

        dockEls = el;
      }

      function showDock(pngDataUrl, suggestedName){
        ensureDockUi();
        const { dock, img, name, title, note, link, targets } = dockEls;
        title.textContent = 'Snapshot'; note.textContent = ''; link.textContent='';
        img.src = pngDataUrl || '';
        name.value = suggestedName || 'snapshot.png';
        dock.style.display = 'block';
        targets.innerHTML = storageSummaryHtml();
        saveDockPos(dock);
      }

      function storageSummaryHtml(){
        const parts = [];
        parts.push(`<div><strong>Webhook:</strong> ${cbUrl ? escapeHtml(new URL(cbUrl, location.href).origin) : '—'}</div>`);
        parts.push(`<div><strong>Drive Direct:</strong> ${driveDirect && gClientId ? 'ready' : '—'}</div>`);
        if (driveDirect && gClientId) {
          parts.push(`<div><strong>Folder:</strong> ${driveFolderIdCandidate ? 'id:' + escapeHtml(driveFolderIdCandidate) : escapeHtml(DRIVE_FOLDER_DEFAULT_NAME)}</div>`);
        }
        return parts.join('');
      }

      async function onSnapClick(){
        if (!batchItems.length || currentIndex < 0) return;
        const btn = document.getElementById('btnSnap');
        if (!snapArmed){
          ensureSnapshotUi();
          if (!restoreFrameRect()) {
            const r = defaultFrameRect(); setFrameRect(r.left, r.top, r.width, r.height);
          }
          snapEls.overlay.style.display='block';
          snapEls.helper.style.display='block';
          btn.classList.add('armed'); btn.setAttribute('aria-label','Capture framed snapshot');
          snapArmed = true;
        } else {
          try{
            await ensureLibs();
            await waitForTilesReady(map, 12000);

            const it = (currentIndex>=0 && currentIndex<batchItems.length) ? batchItems[currentIndex] : null;
            if (!it) throw new Error('No focused route to capture.');
            const r = getFrameRect();

            // Composite capture (prevents blank maps)
            const canvas = await captureCanvas(r);
            drawBannerOntoCanvas(canvas, it, r);
            flashCropped(r);

            const dataUrl = canvas.toDataURL('image/png');
            if (dataUrl.length < 256) throw new Error('Empty image produced. Check CORS or waitForTilesReady.');
            lastPngDataUrl = dataUrl;

            const rawName = it.outName || `${safeName(it.driver)}_${safeName(it.day)}.png`;
            lastSuggestedName = ensurePngExt(rawName);
            showDock(lastPngDataUrl, lastSuggestedName);
            saveFrameRect();
          } catch (e) {
            showDock(null, null);
            dockEls.note.textContent = `Capture failed — ${String(e && e.message || e)}`;
          } finally {
            exitFraming();
          }
        }
      }

      async function onSaveClick(){
        const { title, note, link, saveBtn, name } = dockEls;
        if (!lastPngDataUrl) { note.textContent = 'No snapshot to save.'; return; }
        const ctx = currentIndex>=0 ? batchItems[currentIndex] : {};
        const outName = ensurePngExt((name.value || lastSuggestedName || 'snapshot.png').replace(/[^\w.-]+/g,'_'));

        link.innerHTML = ''; title.textContent = 'Saving…'; note.textContent = '';
        saveBtn.disabled = true;

        try {
          let success = false, reply = null;

          // 1) Drive Direct (if enabled)
          if (driveDirect && gClientId) {
            try {
              reply = await uploadToDriveDirect({
                dataUrl: lastPngDataUrl,
                name: outName,
                folderId: driveFolderIdCandidate || '',
                folderName: DRIVE_FOLDER_DEFAULT_NAME,
                clientId: gClientId
              });
              if (reply && reply.id) {
                success = true;
                const href = toDriveViewLink(reply);
                title.textContent = 'Saved ✓';
                link.innerHTML = href ? `<a href="${href}" target="_blank" rel="noopener">Open in Drive</a>` : '';
                note.innerHTML = `Saved to <b>${escapeHtml(DRIVE_FOLDER_DEFAULT_NAME)}</b> as <b>${escapeHtml(outName)}</b>.`;
              }
            } catch (e) {
              console.warn('Direct Drive upload failed; trying webhook next:', e);
            }
          }

          // 2) Webhook fallback (Apps Script doPost)
          if (!success && cbUrl) {
            const payload = {
              name: outName,
              day: ctx.day, driver: ctx.driver, routeName: ctx.name,
              pngBase64: lastPngDataUrl,
              folderId: driveFolderIdCandidate || '',
              folderName: DRIVE_FOLDER_DEFAULT_NAME,
            };
            reply = await saveViaWebhook(cbUrl, payload);
            if (reply && (reply.ok || reply.success)) {
              success = true;
              const href = toDriveViewLink(reply);
              title.textContent = 'Saved ✓';
              link.innerHTML = href ? `<a href="${href}" target="_blank" rel="noopener">Open in Drive</a>` : '';
              const folderPretty = reply.folder || DRIVE_FOLDER_DEFAULT_NAME;
              note.innerHTML = `Saved to <b>${escapeHtml(folderPretty)}</b> as <b>${escapeHtml(outName)}</b>.`;
            }
          }

          // 3) No target configured / unconfirmed
          if (!success) {
            title.textContent = 'Saved (unconfirmed)';
            note.innerHTML = cbUrl
              ? 'Upload returned no readable ACK. Ensure your webhook returns JSON (or JSON text).'
              : 'No upload target configured. Use Download or provide ?cb=… (webhook) or ?driveDirect=1&gClientId=…';
          }
        } catch (err) {
          title.textContent = 'Save failed';
          note.textContent = `Upload failed — you can still Download. (${String(err && err.message || err)})`;
        } finally {
          saveBtn.disabled = false;
        }
      }

      function onDownloadClick(){
        const { note, name, title } = dockEls;
        if (!lastPngDataUrl) { note.textContent = 'No snapshot to download.'; return; }
        const nm = ensurePngExt((name.value || lastSuggestedName || 'snapshot.png').replace(/[^\w.-]+/g,'_'));
        downloadFallback(lastPngDataUrl, nm);
        title.textContent = 'Downloaded ✓';
        note.textContent = `Downloaded as ${nm}.`;
      }

      function cancelFraming(){
        if (!snapArmed) return;
        exitFraming();
        showDock(null, null);
        dockEls.note.textContent = 'Framing cancelled.';
      }
      function exitFraming(){
        const btn = document.getElementById('btnSnap');
        if (btn) { btn.classList.remove('armed'); btn.setAttribute('aria-label','Frame & capture snapshot'); }
        if (snapEls) { snapEls.overlay.style.display='none'; snapEls.helper.style.display='none'; }
        snapArmed = false;
      }

      // ---------- framing helpers ----------
      function defaultFrameRect(){
        const vw = window.innerWidth, vh = window.innerHeight;
        const w = Math.round(vw*0.7), h = Math.round(vh*0.6);
        const l = Math.round((vw - w)/2), t = Math.round((vh - h)/2);
        return { left:l, top:t, width:w, height:h };
      }
      function setFrameRect(l,t,w,h){
        const f = snapEls.frame;
        Object.assign(f.style, {
          left: `${clamp(l,0,window.innerWidth-20)}px`,
          top: `${clamp(t,0,window.innerHeight-20)}px`,
          width: `${Math.max(40, Math.min(w, window.innerWidth))}px`,
          height:`${Math.max(40, Math.min(h, window.innerHeight))}px`
        });
      }
      function getFrameRect(){ const f = snapEls.frame.getBoundingClientRect(); return { left: Math.round(f.left), top: Math.round(f.top), width: Math.round(f.width), height: Math.round(f.height) }; }
      function saveFrameRect(){ try{ localStorage.setItem(LS_KEYS.frame, JSON.stringify(getFrameRect())); }catch{} }
      function restoreFrameRect(){ try{ const raw = localStorage.getItem(LS_KEYS.frame); if (!raw) return false; const p = JSON.parse(raw); if (['left','top','width','height'].every(k => typeof p[k]==='number')) { setFrameRect(p.left,p.top,p.width,p.height); return true; } }catch{} return false; }

      function bindFramingGestures(frameEl){
        let mode=null, start={x:0,y:0,l:0,t:0,w:0,h:0}, dir=null;
        const onDown = (e, d) => {
          e.preventDefault();
          mode = d ? 'resize' : 'move';
          dir = d || null;
          const r = frameEl.getBoundingClientRect();
          start = { x:e.clientX, y:e.clientY, l:r.left, t:r.top, w:r.width, h:r.height };
          window.addEventListener('pointermove', onMove, {passive:false});
          window.addEventListener('pointerup', onUp, {once:true});
        };
        const onMove = (e) => {
          e.preventDefault();
          const dx = e.clientX - start.x, dy = e.clientY - start.y;
          let l=start.l, t=start.t, w=start.w, h=start.h;
          if (mode==='move'){ l = clamp(start.l + dx, 0, window.innerWidth  - w); t = clamp(start.t + dy, 0, window.innerHeight - h); }
          else {
            if (dir.includes('e')) w = clamp(start.w + dx, 40, window.innerWidth);
            if (dir.includes('s')) h = clamp(start.h + dy, 40, window.innerHeight);
            if (dir.includes('w')) { l = clamp(start.l + dx, 0, start.l + start.w - 40); w = clamp(start.w - dx, 40, window.innerWidth); }
            if (dir.includes('n')) { t = clamp(start.t + dy, 0, start.t + start.h - 40); h = clamp(start.h - dy, 40, window.innerHeight); }
          }
          setFrameRect(l,t,w,h);
        };
        const onUp = () => { window.removeEventListener('pointermove', onMove); saveFrameRect(); };
        frameEl.addEventListener('pointerdown', (e)=>{ if (!e.target.classList.contains('handle')) onDown(e, null); });
        frameEl.querySelectorAll('.handle').forEach(h => h.addEventListener('pointerdown', (e)=> onDown(e, h.getAttribute('data-dir'))));
      }

      // ---------- capture helpers ----------
      async function waitForTilesReady(map, timeoutMs=10000){
        const tileLayers = [];
        map.eachLayer(l => { if (l instanceof L.TileLayer && typeof l._tilesToLoad === 'number') tileLayers.push(l); });

        const allReady = () => tileLayers.every(l => l._tilesToLoad === 0);
        if (allReady()) return;

        await new Promise((resolve) => {
          const t0 = performance.now();
          const tick = () => {
            if (allReady() || performance.now() - t0 > timeoutMs) return resolve();
            setTimeout(tick, 120);
          };
          tick();
        });
      }

      // Composite capture that avoids blank maps
      async function captureCanvas(r){
        await ensureLibs();

        const mapEl = document.getElementById('map');
        if (!mapEl) throw new Error('#map not found');

        const mapRect = mapEl.getBoundingClientRect();
        const ix = intersectRects(r, { left: mapRect.left, top: mapRect.top, width: mapRect.width, height: mapRect.height });
        const frameArea = Math.max(1, r.width * r.height);
        const mapShare = ix ? (ix.width * ix.height) / frameArea : 0;

        // If the frame overlaps the map even a bit, render the map via leaflet-image.
        if (ix && mapShare > 0.10) {
          if (!window.leafletImage) throw new Error('leaflet-image missing');
          const baseCanvas = await new Promise((resolve, reject) => leafletImage(map, (err, c) => err ? reject(err) : resolve(c)));

          const out = document.createElement('canvas');
          out.width = r.width; out.height = r.height;
          const ctx = out.getContext('2d');

          const sx = Math.max(0, Math.round(ix.left - mapRect.left));
          const sy = Math.max(0, Math.round(ix.top  - mapRect.top));
          const sw = Math.round(ix.width);
          const sh = Math.round(ix.height);
          const dx = Math.round(ix.left - r.left);
          const dy = Math.round(ix.top  - r.top);

          ctx.clearRect(0, 0, out.width, out.height);
          ctx.drawImage(baseCanvas, sx, sy, sw, sh, dx, dy, sw, sh);

          return out; // banner drawn later
        }

        // UI-only capture (no map region). Exclude the banner (we draw it ourselves) and dock/overlay
        if (!window.html2canvas) throw new Error('html2canvas missing');
        return await html2canvas(document.body, {
          useCORS: true,
          allowTaint: false,
          backgroundColor: null,
          x: r.left, y: r.top, width: r.width, height: r.height,
          windowWidth: document.documentElement.clientWidth,
          windowHeight: document.documentElement.clientHeight,
          scrollX: 0, scrollY: 0,
          imageTimeout: 15000,
          foreignObjectRendering: false,
          ignoreElements: (el) => {
            if (!el) return false;
            const isMap     = (el.id === 'map') || el.closest?.('#map');
            const isBanner  = (el.id === 'dispatchBanner') || el.closest?.('#dispatchBanner');
            const cls = el.classList || { contains: ()=>false };
            return isMap ||
                   isBanner ||
                   cls.contains('snap-overlay') ||
                   cls.contains('snap-dock') ||
                   !!el.closest?.('.snap-dock');
          }
        });
      }

      function flashCropped(r){
        if (!snapEls) return;
        const f = snapEls.flashCrop;
        Object.assign(f.style, { left:`${r.left}px`, top:`${r.top}px`, width:`${r.width}px`, height:`${r.height}px` });
        f.style.opacity = '0.65';
        setTimeout(()=>{ f.style.opacity = '0'; }, 125);
      }
      function intersectRects(a,b){
        const x1 = Math.max(a.left, b.left), y1 = Math.max(a.top,  b.top);
        const x2 = Math.min(a.left+a.width,  b.left+b.width);
        const y2 = Math.min(a.top +a.height, b.top +b.height);
        if (x2<=x1 || y2<=y1) return null;
        return { left:x1, top:y1, width:x2-x1, height:y2-y1 };
      }

      // =================================================================
      // Auto export (headless)
      // =================================================================
      async function runAutoExport(items){
        ['legend','drivers','status','error'].forEach(id=>{ const n=document.getElementById(id); if(n) n.style.display='none'; });

        for (let i=0;i<items.length;i++){
          const it = items[i];
          try{
            manualSelectedKeys = (it.keys||[]).map(normalizeKey);
            runtimeCustEnabled = true;
            await applySelection();
            await loadCustomersIfAny();
            if (selectionBounds) fitWithHints(selectionBounds, it?.view || null);
            await ensureLibs(); await waitForTilesReady(map, 12000);

            // Map-only snapshot for auto-export (banner drawn onto it)
            const canvas = await new Promise((resolve, reject) => leafletImage(map, (err, c) => (err ? reject(err) : resolve(c))));
            drawBannerOntoCanvas(canvas, it, {left:0,top:0,width:canvas.width,height:canvas.height});
            const png = canvas.toDataURL('image/png');
            const name = ensurePngExt(it.outName || `${safeName(it.driver)}_${safeName(it.day)}.png`);

            let reply = null;
            if (cbUrl) {
              reply = await saveViaWebhook(cbUrl, { name, day: it.day, driver: it.driver, routeName: it.name, pngBase64: png, folderId: driveFolderIdCandidate, folderName: DRIVE_FOLDER_DEFAULT_NAME });
            } else if (driveDirect && gClientId) {
              reply = await uploadToDriveDirect({ dataUrl: png, name, folderId: driveFolderIdCandidate, folderName: DRIVE_FOLDER_DEFAULT_NAME, clientId: gClientId });
            } else {
              downloadFallback(png, name);
            }
            if (reply) console.log('Saved:', toDriveViewLink(reply) || reply);
          }catch(e){ console.error('auto export item failed', e); }
        }
      }

      // =================================================================
      // Loaders / selection / customers
      // =================================================================
      async function loadLayerSet(arr, collector, addToMap = true) {
        for (const Lcfg of (arr || [])) {
          try {
            const url = Lcfg.url;
            const gj = await fetchJson(url);
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
          } catch (err) {
            // BETTER error surfacing
            console.error('[layer] load failed', Lcfg, err);
            warn(`Layer load failed for ${escapeHtml(Lcfg?.day || 'day')} — check ${escapeHtml(Lcfg?.url || '(missing URL)')} (CORS / 404?).`);
          }
        }
        updateDiagnostics();
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
              const rowsAA = parseCsvRows(await fetchText(selUrl));
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
        if (!noKeys) for (const k of selectedSet) {
          if (isSubquadrantKey(k)) { subqBasesSelected.add(baseKeyFrom(k)); const pq = basePlusQuad(k); if (pq) subqQuadsSelected.add(pq); }
          else if (isQuadrantKey(k)) quadBasesSelected.add(baseKeyFrom(k));
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
          if (cfg.behavior?.autoZoom && allBounds) map.fitBounds(allBounds.pad(0.1));
        } else {
          for (const entry of baseDayLayers)  for (const lyr of entry.features) { const k=lyr._routeKey; const visible = selectedSet.has(k) && !quadBasesSelected.has(k) && !subqBasesSelected.has(k); setFeatureVisible(entry, lyr, visible, visible); }
          for (const entry of quadDayLayers)  for (const lyr of entry.features) { const k=lyr._routeKey; const bq=basePlusQuad(k); const visible = selectedSet.has(k) && !subqQuadsSelected.has(bq); setFeatureVisible(entry, lyr, visible, visible); }
          for (const entry of subqDayLayers)  for (const lyr of entry.features) { const k=lyr._routeKey; const visible = selectedSet.has(k); setFeatureVisible(entry, lyr, visible, visible); }

          rebuildCoverageFromVisible();
          selectedMunicipalities = Array.from(new Set(selectedMunicipalities)).sort((a,b)=>a.toLowerCase().localeCompare(b.toLowerCase()));
          if (cfg.behavior?.autoZoom) { if (selectionBounds) map.fitBounds(selectionBounds.pad(0.1)); else if (allBounds) map.fitBounds(allBounds.pad(0.1)); }
          info(`Loaded ${selectedSet.size} selected key(s).`);
          if (selectedSet.size > 0 && coveragePolysSelected.length === 0) warn('Selected keys matched no polygons. Check cfg.fields.key vs GeoJSON.');
        }

        recolorAndRecountCustomers();
        const activeKeysOrderedLower = selectedOrderedKeys.filter(k => visibleSelectedKeysSet.has(k)).map(k => k.toLowerCase());
        updateLegend();
        setStatus(makeStatusLine(selectedMunicipalities, custWithinSel, custOutsideSel, activeKeysOrderedLower));
        await rebuildDriverOverlays();
        updateDiagnostics();

        // NEW: mark ready after successful selection rebuild
        setStatus('Ready.');
      }

      function rebuildCoverageFromVisible() {
        coveragePolysAll = [];
        for (const arr of allDaySets) for (const entry of arr) for (const lyr of entry.features) {
          if (entry.layer.hasLayer(lyr)) { coveragePolysAll.push({ feat: lyr._turfFeat, perDay: entry.perDay, layerRef: lyr }); lyr._custAny = 0; lyr._custSel = 0; }
        }
      }

      async function loadCustomersIfAny() {
        customerLayer.clearLayers();
        customerMarkers.length = 0;
        customerCount = 0;

        const custCfg = cfg.customers || {};
        const forceOff = (runtimeCustEnabled === false);

        if (forceOff || !custCfg.enabled || !custCfg.url) {
          custWithinSel = custOutsideSel = 0; resetDayCounts();
          updateLegend();
          setStatus(makeStatusLine(selectedMunicipalities, custWithinSel, custOutsideSel, []));
          renderDriversPanel(driverMeta, driverOverlays, true, driverSelectedCounts, custWithinSel);
          updateDiagnostics();
          return;
        }

        phase('Loading customers…');
        const text = await fetchText(custCfg.url);
        const rows = parseCsvRows(text);
        if (!rows.length) { custWithinSel = custOutsideSel = 0; resetDayCounts(); updateLegend(); setStatus(makeStatusLine(selectedMunicipalities, custWithinSel, custOutsideSel, [])); renderDriversPanel(driverMeta, driverOverlays, true, driverSelectedCounts, custWithinSel); updateDiagnostics(); return; }

        // Header finder uses '&&' fix
        const hdrIdx = findCustomerHeaderIndex(rows, custCfg.schema || { coords: 'Verified Coordinates', note: 'Order Note' });
        if (hdrIdx === -1) { warn('Customers CSV: header not found.'); custWithinSel = custOutsideSel = 0; resetDayCounts(); updateLegend(); setStatus(makeStatusLine(selectedMunicipalities, custWithinSel, custOutsideSel, [])); renderDriversPanel(driverMeta, driverOverlays, true, driverSelectedCounts, custWithinSel); updateDiagnostics(); return; }

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

        if (!(window.turf && turf.booleanPointInPolygon)) { try { await ensureLibs(); } catch {} }
        recolorAndRecountCustomers();

        const activeKeysOrderedLower = selectedOrderedKeys.filter(k => visibleSelectedKeysSet.has(k)).map(k => k.toLowerCase());
        updateLegend();
        setStatus(makeStatusLine(selectedMunicipalities, custWithinSel, custOutsideSel, activeKeysOrderedLower));
        renderDriversPanel(driverMeta, driverOverlays, true, driverSelectedCounts, custWithinSel);
        updateDiagnostics();
      }

      function resetDayCounts(){ custByDayInSel.Wednesday=0; custByDayInSel.Thursday=0; custByDayInSel.Friday=0; custByDayInSel.Saturday=0; }

      function recolorAndRecountCustomers() {
        let inSel = 0, outSel = 0;
        resetDayCounts();
        for (const rec of coveragePolysAll) { if (rec.layerRef) { rec.layerRef._custAny = 0; rec.layerRef._custSel = 0; } }

        const turfOn = (typeof turf !== 'undefined');
        const cst = cfg.style?.customers || {};
        let outStyle = outsideHighlight
          ? { radius: (cst.radius || 9), color: '#d32f2f', weight: (cst.weightPx || 2), opacity: 0.95, fillColor: '#ffcdd2', fillOpacity: 0.95 }
          : { radius: (cst.radius || 9), color: '#7a7a7a', weight: (cst.weightPx || 2), opacity: 0.8,  fillColor: '#c7c7c7', fillOpacity: 0.6 };

        const onlySelectedCustomers = manualMode && (currentIndex >= 0) && !outsideHighlight;

        for (const rec of customerMarkers) {
          let show = true, style = { ...outStyle }, insideSel = false, selDay = null;

          if (turfOn) {
            const pt = turf.point([rec.lng, rec.lat]);
            for (let j = 0; j < coveragePolysAll.length; j++) {
              if (turf.booleanPointInPolygon(pt, coveragePolysAll[j].feat)) { const lyr = coveragePolysAll[j].layerRef; if (lyr) lyr._custAny += 1; break; }
            }
            for (let k = 0; k < coveragePolysSelected.length; k++) {
              if (turf.booleanPointInPolygon(pt, coveragePolysSelected[k].feat)) {
                insideSel = true; selDay = (coveragePolysSelected[k].feat.properties.day || '').trim();
                const pd = coveragePolysSelected[k].perDay || {};
                style = { radius:(cst.radius||9), color: pd.stroke || (cst.stroke || '#111'), weight:(cst.weightPx||2), opacity:(cst.opacity ?? 0.95), fillColor: pd.fill || (cst.fill || '#ffffff'), fillOpacity:(cst.fillOpacity ?? 0.95) };
                const lyr = coveragePolysSelected[k].layerRef; if (lyr) lyr._custSel += 1; break;
              }
            }
          }

          if (onlySelectedCustomers && !insideSel) show = false;

          if (insideSel) { inSel++; if (selDay && custByDayInSel[selDay] != null) custByDayInSel[selDay] += 1; }
          else { outSel++; }

          if (show && !rec.visible) { customerLayer.addLayer(rec.marker); rec.visible = true; }
          else if (!show && rec.visible) { customerLayer.removeLayer(rec.marker); rec.visible = false; }

          if (show) rec.marker.setStyle(style);
        }

        custWithinSel = inSel; custOutsideSel = outSel;
        driverSelectedCounts = computeDriverCounts();
        renderDriversPanel(driverMeta, driverOverlays, true, driverSelectedCounts, custWithinSel);
        updateLegend();
      }

      function computeDriverCounts() {
        const out = {};
        for (const rec of coveragePolysSelected) {
          const lyr = rec.layerRef; if (!lyr) continue;
          const drv = lookupDriverForKey(lyr._routeKey); if (!drv) continue;
          out[drv] = (out[drv] || 0) + (lyr._custSel || 0);
        }
        return out;
      }

      async function rebuildDriverOverlays() {
        Object.values(driverOverlays).forEach(rec => { try { map.removeLayer(rec.group); } catch {} });
        for (const k of Object.keys(driverOverlays)) delete driverOverlays[k];

        const byDriver = new Map();
        coveragePolysSelected.forEach(rec => {
          const key = rec.layerRef && rec.layerRef._routeKey; if (!key) return;
          const drv = lookupDriverForKey(key); if (!drv) return;
          if (!byDriver.has(drv)) byDriver.set(drv, []); byDriver.get(drv).push(rec.layerRef._turfFeat);
        });

        if (byDriver.size === 0) { renderDriversPanel(driverMeta, {}, false, driverSelectedCounts, custWithinSel); return; }

        byDriver.forEach((features, name) => {
          const meta = (driverMeta.find(d => (d.name || '').toLowerCase() === (name || '').toLowerCase())
                      || { name, color: colorFromName(name) });
          const color = meta.color || '#888';

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
            const b = outline.getBounds(); const center = b?.getCenter?.();
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

      function lookupDriverForKey(key) {
        if (!key) return null;
        if (activeAssignMap[key]) return activeAssignMap[key];
        const m = String(key).match(/^([WTFS]\d+)/);
        if (m && activeAssignMap[m[1]]) return activeAssignMap[m[1]];
        return null;
      }

      // ---------- Popups/focus ----------
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

      // ---------- Drivers & Legend ----------
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
          const safe = escapeHtml(name || '');
          const isPresent = !!overlays?.[name];
          const onAttr = isPresent && defaultOn ? 'checked' : '';
          const count = typeof countsMap[name] === 'number' ? countsMap[name] : 0;
          const frac = `${count}/${totalSelected || 0}`;
          return `<div class="row" style="display:flex;align-items:center;gap:8px;margin:4px 0">
            <input type="checkbox" data-driver="${safe}" aria-label="Toggle ${safe}" ${onAttr} ${isPresent ? '' : 'disabled'}>
            <span class="swatch" style="width:16px;height:16px;border-radius:3px;border:2px solid ${col};background:${col};box-sizing:border-box"></span>
            <div>${safe}</div>
            <div class="counts" style="margin-left:auto;opacity:.8;font-variant-numeric:tabular-nums">${frac}</div>
          </div>`;
        }).join('');

        el.innerHTML = `<h4 style="margin:0 0 6px 0;font-size:14px">Drivers</h4>${rows}`;
        el.querySelectorAll('input[type="checkbox"]').forEach(cb => {
          const name = cb.getAttribute('data-driver');
          if (!overlays?.[name]) return;
          toggleDriverOverlay(name, !!cb.checked);
          cb.addEventListener('change', (e) => toggleDriverOverlay(name, e.target.checked));
        });
      }
      function toggleDriverOverlay(name, on) { const rec = driverOverlays[name]; if (!rec) return; if (on) { rec.group.addTo(map); try { rec.labelMarker?.openTooltip(); } catch{} } else { map.removeLayer(rec.group); } }

      function renderLegend(cfg, legendCounts, custIn, custOut, outsideToggle) {
        const el = document.getElementById('legend'); if (!el) return;
        // NEW: gate 0/0 display until features are known
        const totalKnown = Array.isArray(cfg.layers) && cfg.layers.length && totalFeatureCount() > 0;
        const rowsHtml = (cfg.layers || []).map(Lcfg => {
          const st = (cfg.style?.perDay?.[Lcfg.day]) || {};
          const c  = (legendCounts && legendCounts[Lcfg.day]) ? legendCounts[Lcfg.day] : { selected: 0, total: 0 };
          const frac = totalKnown ? `${c.selected}/${c.total}` : '—';
          return `<div class="row" style="display:flex;align-items:center;gap:8px;margin:4px 0">
            <span class="swatch" style="width:16px;height:16px;border-radius:3px;border:2px solid ${st.stroke};background:${st.fill};box-sizing:border-box"></span>
            <div>${Lcfg.day}</div>
            <div class="counts" style="margin-left:auto;opacity:.8;font-variant-numeric:tabular-nums">${frac}</div>
          </div>`;
        }).join('');
        const custBlock = `<div class="row" style="margin-top:6px;border-top:1px solid #eee;padding-top:6px;display:flex;gap:8px;align-items:center">
            <div>Customers within selection:</div><div class="counts" style="margin-left:auto;opacity:.8;font-variant-numeric:tabular-nums">${custIn ?? 0}</div>
          </div>
          <div class="row" style="display:flex;gap:8px;align-items:center">
            <div>Customers outside selection:</div><div class="counts" style="margin-left:auto;opacity:.8;font-variant-numeric:tabular-nums">${custOut ?? 0}</div>
          </div>`;
        const toggle = `<div class="row" style="margin-top:4px;display:flex;gap:8px;align-items:center">
            <input type="checkbox" id="toggleOutside" ${outsideToggle ? 'checked' : ''} aria-label="Highlight outside customers">
            <div>Highlight outside customers</div>
          </div>`;
        el.innerHTML = `<h4 style="margin:0 0 6px 0;font-size:14px">Layers</h4>${rowsHtml}${custBlock}${toggle}`;

        const tgl = el.querySelector('#toggleOutside');
        if (tgl) {
          tgl.addEventListener('change', (e)=>{
            outsideHighlight = !!e.target.checked;
            try { localStorage.setItem(LS_KEYS.outside, outsideHighlight ? '1' : '0'); } catch {}
            recolorAndRecountCustomers();
          });
        }
      }

      // Put this directly after renderLegend(...) inside the IIFE
      function updateLegend() {
        // Build per-day counts from what's currently visible (coveragePolysAll)
        // and what's part of the active selection (coveragePolysSelected).
        // Works across base/quadrant/subquadrant because it only looks at what's
        // actually on the map right now.
        const byDay = {};
        const bump = (day, field) => {
          const d = day || '';
          if (!byDay[d]) byDay[d] = { selected: 0, total: 0 };
          byDay[d][field] += 1;
        };

        // total = visible features currently on the map
        for (const rec of coveragePolysAll) {
          bump(rec?.feat?.properties?.day, 'total');
        }
        // selected = features that are part of the active selection
        for (const rec of coveragePolysSelected) {
          bump(rec?.feat?.properties?.day, 'selected');
        }

        // Push the numbers into the legend
        renderLegend(cfg, byDay, custWithinSel, custOutsideSel, outsideHighlight);
      }

      // =================================================================
      // Helpers
      // =================================================================
      function totalFeatureCount(){
        try { return [...baseDayLayers, ...quadDayLayers, ...subqDayLayers].reduce((acc,e)=> acc + (e.features?.length || 0), 0); }
        catch { return 0; }
      }
      function fitWithHints(bounds, hints){
        try{
          const padPct = Math.max(0, Math.min(0.25, (hints && hints.padPct) || 0.10));
          map.fitBounds(bounds.pad(padPct));
          const minZ = (hints && Number(hints.minZoom)) || 7;
          const maxZ = (hints && Number(hints.maxZoom)) || 11;
          const z = map.getZoom();
          if (z > maxZ) map.setZoom(maxZ);
          if (z < minZ) map.setZoom(minZ);
        }catch(e){}
      }
      function ensureMinHeight(el){ try{ const h = parseInt(getComputedStyle(el).height, 10); if (!isFinite(h) || h < 40) { const fix = document.createElement('style'); fix.textContent = `#map{height:100vh}`; document.head.appendChild(fix); } }catch{} }
      function parseLatLng(s) { const m = String(s||'').trim().match(/(-?\d+(?:\.\d+)?)\s*[, ]\s*(-?\d+(?:\.\d+)?)/); if (!m) return null; const lat = parseFloat(m[1]), lng = parseFloat(m[2]); if (!isFinite(lat) || !isFinite(lng)) return null; if (lat < -90 || lat > 90 || lng < -180 || lng > 180) return null; return { lat, lng }; }
      function smartTitleCase(s) { if (!s) return ''; s = s.toLowerCase(); const parts = s.split(/([\/-])/g); const small = new Set(['and','or','of','the','a','an','in','on','at','by','for','to','de','la','le','el','du','von','van','di','da','del']); const fix = (w,first)=>!first && small.has(w) ? w : (w==='st'||w==='st.'?'St.':w==='mt'||w==='mt.'?'Mt.': w.charAt(0).toUpperCase()+w.slice(1)); let idx=0; for (let i=0;i<parts.length;i++){ if (parts[i]=='/'||parts[i]=='-') continue; parts[i]=parts[i].split(/\s+/).map(tok=>fix(tok, idx++===0)).join(' ');} return parts.map(p=>p==='/'?'/':(p==='-'?'-':p)).join('').replace(/\s+/g, ' ').trim(); }
      function escapeHtml(s) { return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;'); }
      function clamp(v,min,max){ return Math.max(min, Math.min(max, v)); }
      const safeName = s => String(s||'').replace(/[^\w.-]+/g,'_');
      const ensurePngExt = s => /\.(png)$/i.test(s) ? s : (s.replace(/\.[a-z0-9]+$/i,'') + '.png');

      // ✔ NEW: deterministic name → color helper (needed by driver overlays & panel)
      function colorFromName(name, opts = {}) {
        const s = String(name || '');
        // fast FNV-1a hash
        let h = 2166136261 >>> 0;
        for (let i = 0; i < s.length; i++) {
          h ^= s.charCodeAt(i);
          h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0;
        }
        const hue = h % 360;
        const sat = opts.sat ?? 70;     // tweak if you want softer colors
        const light = opts.light ?? 45; // tweak for lighter/darker strokes
        // Use classic HSL syntax for broad browser support
        return `hsl(${hue}, ${sat}%, ${light}%)`;
      }

      // Webhook helper (text/plain to avoid preflight); robust reply parsing
      async function saveViaWebhook(url, payload){
        const res = await fetch(url, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify(payload)
        });
        if (!res.ok) throw new Error(`Webhook HTTP ${res.status}`);
        const ct = (res.headers.get('content-type') || '').toLowerCase();
        if (ct.includes('application/json')) return await res.json();
        const txt = await res.text();
        try { return JSON.parse(txt); } catch { return null; }
      }

      function toDriveViewLink(reply){
        if (!reply || typeof reply !== 'object') return '';
        const id  = reply.id || reply.fileId || '';
        const url = reply.webViewLink || reply.fileUrl || reply.alternateLink || reply.openUrl || '';
        if (url) return url;
        if (id)  return `https://drive.google.com/file/d/${id}/view`;
        return '';
      }

      // Safer folder id extraction
      function extractDriveFolderId(s){
        if (!s) return '';
        const str = String(s);
        const m = str.match(/(?:folders\/|id=)([-\w]{25,})/);
        if (m) return m[1];
        const idOnly = str.match(/^[-\w]{25,}$/);
        return idOnly ? idOnly[0] : '';
      }

      function showLabel(lyr, text) { if (lyr.getTooltip()) lyr.unbindTooltip(); lyr.bindTooltip(text, { permanent: true, direction: 'center', className: 'lbl' }); }
      function hideLabel(lyr) { if (lyr.getTooltip()) lyr.unbindTooltip(); }
      function dimFill(perDay, cfg) { const base = perDay.fillOpacity ?? 0.8; const factor = cfg.style?.dimmed?.fillFactor ?? 0.3; return Math.max(0.08, base * factor); }

      function phase(msg){ setStatus(`⏳ ${msg}`); }
      function info(msg){ setStatus(`ℹ️ ${msg}`); }
      function warn(msg){ showTopError('Warning', msg); }
      function setStatus(msg) { const n = document.getElementById('status'); if (n) n.textContent = msg || ''; }
      function makeStatusLine(selMunis, inCount, outCount, activeKeysLowerArray) {
        const muniList = (Array.isArray(selMunis) && selMunis.length) ? selMunis.join(', ') : '—';
        const keysTxt = (Array.isArray(activeKeysLowerArray) && activeKeysLowerArray.length) ? activeKeysLowerArray.join(', ') : '—';
        return `Customers (in/out): ${inCount}/${outCount} • Municipalities: ${muniList} • active zone keys: ${keysTxt}`;
      }

      // ---------- data/CSV helpers ----------
      function parseBatchItems(b64){
        if (!b64) return [];
        try {
          const json = atob(String(b64).replace(/-/g,'+').replace(/_/g,'/'));
          const arr = JSON.parse(json);
          return Array.isArray(arr) ? arr : [];
        } catch { return []; }
      }
      function splitKeys(s) { return String(s||'').split(/[;,/|]/).map(x => x.trim()).filter(Boolean); }
      function unionAllKeys(items){ const set = new Set(); (items || []).forEach(it => (it.keys || []).forEach(k => set.add(normalizeKey(k)))); return Array.from(set); }
      function normalizeKey(s) { s = String(s || '').trim().toUpperCase(); const m = s.match(/^([WTFS])0*(\d+)(_.+)?$/); return m ? (m[1] + String(parseInt(m[2], 10)) + (m[3] || '')) : s; }
      function baseKeyFrom(key) { const m = String(key||'').toUpperCase().match(/^([WTFS]\d+)/); return m ? m[1] : String(key||'').toUpperCase(); }
      function quadParts(key) { const m = String(key || '').toUpperCase().match(/_(NE|NW|SE|SW)(?:_(TL|TR|LL|LR))?$/); return m ? { quad: m[1], sub: m[2] || null } : null; }
      function basePlusQuad(key) { const p = quadParts(key); return p ? (baseKeyFrom(key) + '_' + p.quad) : null; }
      function isSubquadrantKey(k) { const p = quadParts(k); return !!(p && p.sub); }
      function isQuadrantKey(k)    { const p = quadParts(k); return !!(p && !p.sub); }

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
        const idx = (name) => Array.isArray(hdrRow)
          ? hdrRow.findIndex(h => (h||'').toLowerCase() === name)
          : -1;
        return { coords: idx(wantCoords), note: idx(wantNote) };
      }
      function extractAssignmentsFromCsv(rows, knownNamesSet) {
        const like = (s) => String(s||'').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
        const dayWords = new Set(['mon','monday','tue','tues','tuesday','wed','weds','wednesday','thu','thur','thurs','thursday','fri','friday','sat','saturday','sun','sunday']);
        const looksLikeName = (s) => { const v = like(s); if (!v) return false; if (dayWords.has(v)) return false; if (knownNamesSet && knownNamesSet.size) return knownNamesSet.has(v); return /^[a-z][a-z .'\-]{1,30}$/.test(v) && !/\d/.test(v); };
        const isKeysHeader = (s) => { const v = like(s); return v.includes('key') || v.includes('zone') || v.includes('route'); };
        const isDriverHeader = (s) => { const v = like(s); return v.includes('driver') || v === 'name' || v.includes('assigned'); };

        let headerIndex=-1, driverCol=-1, keysCol=-1;
        for (let i=0; i<Math.min(rows.length, 50); i++) {
          const row = rows[i] || [];
          let d=-1, k=-1;
          for (let c=0;c<row.length;c++) {
            const cell=(row[c]||'');
            if (d === -1 && isDriverHeader(cell)) d = c;
            if (k === -1 && isKeysHeader(cell))   k = c;
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
            if (!looksLikeName(dn)) continue;
            splitKeys(ks).map(normalizeKey).forEach(k => { out[k] = dn; });
          }
          return out;
        }

        // fallback: name at start + keys anywhere on the row
        for (let r = 0; r < Math.min(rows.length, 40); r++) {
          const row = rows[r] || [];
          if (!row.length) continue;
          let dn = '';
          for (let c=0;c<row.length;c++) {
            const cell = (row[c]||'').trim(); if (!cell) continue;
            if (looksLikeName(cell)) { dn = cell; break; }
          }
          if (!dn) continue;
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

      // ---------- diagnostics ----------
      function updateDiagnostics() {
        if (!diagMode) return;
        let box = document.getElementById('diag');
        if (!box) { box = document.createElement('div'); box.id = 'diag'; document.body.appendChild(box); }
        const featureCounts = (arr) => (arr||[]).reduce((a,e)=>a+(e.features?.length||0),0);
        box.style.display = 'block';
        box.innerHTML = [
          `cfg: ${escapeHtml(cfgUrl)}`,
          `features: base=${featureCounts(baseDayLayers)} quad=${featureCounts(quadDayLayers)} subq=${featureCounts(subqDayLayers)}`,
          `batch routes: ${batchItems.length}`,
          `selected keys: ${selectedOrderedKeys.length} → matched polys: ${coveragePolysSelected.length}`,
          `customers: total=${customerCount} in=${custWithinSel} out=${custOutsideSel}`
        ].map(x=>`<div>${x}</div>`).join('');
      }

      // ---------- shell ----------
      function ensureUiShell() {
        const need = [{id:'legend'}, {id:'drivers'}, {id:'status'}, {id:'error', className:'top-error'}];
        for (const spec of need) if (!document.getElementById(spec.id)) { const n = document.createElement('div'); n.id = spec.id; if (spec.className) n.className = spec.className; document.body.appendChild(n); }
      }
      function ensureMapRoot() {
        if (!document.getElementById('map')) {
          const m = document.createElement('div');
          m.id = 'map';
          m.style.height = '100vh';
          document.body.insertBefore(m, document.body.firstChild || null);
        }
      }
      function injectBaseCss(){
        const css = document.createElement('style');
        css.textContent = `
          .top-error{position:fixed;left:12px;right:12px;top:12px;z-index:10000;background:#ffebee;color:#b71c1c;border:1px solid #ffcdd2;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,.08);padding:10px;display:none;font:600 13px system-ui}
          #diag{position:fixed;right:12px;top:12px;z-index:9000;background:rgba(255,255,255,.95);border-radius:8px;box-shadow:0 6px 18px rgba(0,0,0,.12);padding:10px 12px;font:600 12px system-ui;line-height:1.35;display:none;max-width:40ch}
          .lbl{background:rgba(255,255,255,.8);padding:2px 4px;border-radius:4px;border:1px solid rgba(0,0,0,.1);color:#111;font:600 12px system-ui}
        `;
        document.head.appendChild(css);
      }
      function injectToolbarCss(){
        const css = document.createElement('style');
        css.textContent = `
          .route-toolbar{position:absolute;top:10px;left:50%;transform:translateX(-50%);z-index:9000;display:flex;gap:8px}
          .route-toolbar button{background:#111;color:#fff;border:none;border-radius:6px;padding:8px 12px;font:600 14px system-ui;cursor:pointer;opacity:.95;transition:background .15s ease}
          .route-toolbar button:hover{opacity:1}
          .route-toolbar button.armed{background:#c62828}
          .dispatch-banner{position:fixed;left:50%;top:78vh;transform:translateX(-50%);z-index:8500;cursor:grab;background:rgba(255,255,255,.92);backdrop-filter:saturate(120%) blur(2px);border-radius:12px;box-shadow:0 8px 22px rgba(0,0,0,.14);padding:10px 14px;max-width:min(80vw,1100px);display:none}
          .dispatch-banner.dragging{cursor:grabbing; user-select:none}
          .dispatch-banner.visible{display:block}
          .dispatch-banner .row{font:500 14px system-ui;color:#111;line-height:1.35;margin:2px 0;word-wrap:break-word;overflow-wrap:break-word}
          .dispatch-banner .row.stats{white-space:nowrap;overflow-x:auto;overflow-y:hidden;-webkit-overflow-scrolling:touch}
          .dispatch-banner .row strong{font-weight:700}
          .dispatch-banner .row.meta strong{font-weight:700}
          .snap-overlay{position:fixed;inset:0;z-index:9400;pointer-events:none}
          .snap-overlay .helper{position:fixed;left:50%;top:8px;transform:translateX(-50%);background:#111;color:#fff;border-radius:8px;padding:6px 10px;font:600 12px system-ui;pointer-events:auto}
          .snap-overlay .frame{position:fixed;border:2px dashed #c62828;border-radius:10px;box-shadow:0 0 0 9999px rgba(0,0,0,.25);cursor:move;pointer-events:auto}
          .snap-overlay .handle{position:absolute;width:12px;height:12px;background:#fff;border:2px solid #c62828;border-radius:50%}
          .snap-overlay .handle.n{left:50%;top:-8px;transform:translate(-50%,-50%)}
          .snap-overlay .handle.s{left:50%;bottom:-8px;transform:translate(-50%,-50%)}
          .snap-overlay .handle.e{right:-8px;top:50%;transform:translate(50%,-50%)}
          .snap-overlay .handle.w{left:-8px;top:50%;transform:translate(-50%,-50%)}
          .snap-overlay .handle.ne{right:-8px;top:-8px;transform:translate(50%,-50%)}
          .snap-overlay .handle.nw{left:-8px;top:-8px;transform:translate(-50%,-50%)}
          .snap-overlay .handle.se{right:-8px;bottom:-8px;transform:translate(50%,50%)}
          .snap-overlay .handle.sw{left:-8px;bottom:-8px;transform:translate(-50%,50%)}
          .snap-flash-crop{position:fixed;border:2px solid #4caf50;border-radius:10px;pointer-events:none;left:0;top:0;width:0;height:0;opacity:0;transition:opacity .18s ease}
          .snap-dock{position:fixed;left:50%;top:50%;transform:translate(-50%,-50%);z-index:9500;min-width:280px;width:min(90vw,1100px);max-width:1100px;background:#fff;border:1px solid #e8e8e8;border-radius:14px;box-shadow:0 16px 44px rgba(0,0,0,.18);display:none}
          .snap-dock .head{display:flex;align-items:center;gap:12px;padding:10px 12px;border-bottom:1px solid #eee;cursor:move}
          .snap-dock .head .spacer{flex:1}
          .snap-dock .head .x{background:transparent;border:none;font:700 16px system-ui;cursor:pointer}
          .snap-dock .body{display:flex;gap:12px; padding:12px}
          .snap-dock .body .row{margin:6px 0}
          .snap-dock .preview{display:block;max-width:64vw; max-height:70vh;object-fit:contain;background:#f6f7f8;border-radius:10px}
          .snap-dock input[type="text"]{width:100%;padding:6px 8px;border:1px solid #ddd;border-radius:6px;font:600 13px system-ui}
          .snap-dock .btn{background:#111;color:#fff;border:none;border-radius:8px;padding:8px 10px;font:700 13px system-ui;cursor:pointer;margin-right:6px}
          .snap-dock .btn.secondary{background:#f2f3f6;color:#222}
          .snap-dock .link a{font:700 13px system-ui}
          .snap-dock .note{font:600 12px system-ui;color:#555;margin-top:4px}
        `;
        document.head.appendChild(css);
      }

      // ---------- UI Drag helpers ----------
      function makeBannerDraggable(el){
        let dragging=false, dx=0, dy=0;
        const onDown = (e)=>{ dragging=true; el.classList.add('dragging'); const r=el.getBoundingClientRect(); dx=e.clientX - r.left; dy=e.clientY - r.top; e.preventDefault(); };
        const onMove = (e)=>{ if(!dragging) return; const x = clamp(e.clientX - dx, 8, window.innerWidth - el.offsetWidth - 8); const y = clamp(e.clientY - dy, 8, window.innerHeight - el.offsetHeight - 8); el.style.left = `${x}px`; el.style.top = `${y}px`; el.style.transform='none'; };
        const onUp   = ()=>{ if(!dragging) return; dragging=false; el.classList.remove('dragging'); saveBannerPos(el); };
        el.addEventListener('pointerdown', onDown);
        window.addEventListener('pointermove', onMove);
        window.addEventListener('pointerup', onUp);
      }
      function saveBannerPos(el){
        try {
          const r = el.getBoundingClientRect();
          localStorage.setItem(LS_KEYS.banner, JSON.stringify({left:r.left, top:r.top}));
        }catch{}
      }
      function restoreBannerPos(el){
        try{
          const raw = localStorage.getItem(LS_KEYS.banner);
          if (!raw) return;
          const p = JSON.parse(raw);
          if (typeof p.left==='number' && typeof p.top==='number'){
            el.style.left = `${p.left}px`;
            el.style.top  = `${p.top}px`;
            el.style.transform='none';
          }
        }catch{}
      }
      function makeDockDraggable(dock, handle){
        let dragging=false, dx=0, dy=0;
        handle.addEventListener('pointerdown', (e)=>{
          dragging=true; const r=dock.getBoundingClientRect(); dx=e.clientX - r.left; dy=e.clientY - r.top; e.preventDefault();
          document.body.style.userSelect='none';
        });
        window.addEventListener('pointermove', (e)=>{
          if(!dragging) return;
          const x = clamp(e.clientX - dx, 8, window.innerWidth - dock.offsetWidth - 8);
          const y = clamp(e.clientY - dy, 8, window.innerHeight - dock.offsetHeight - 8);
          dock.style.left = `${x}px`;
          dock.style.top  = `${y}px`;
          dock.style.transform='none';
        });
        window.addEventListener('pointerup', ()=>{
          if(!dragging) return;
          dragging=false; document.body.style.userSelect='';
          saveDockPos(dock);
        });
      }
      function saveDockPos(el){
        try {
          const r = el.getBoundingClientRect();
          localStorage.setItem(LS_KEYS.dock, JSON.stringify({left:r.left, top:r.top}));
        }catch{}
      }
      function restoreDockPos(el){
        try{
          const raw = localStorage.getItem(LS_KEYS.dock);
          if (!raw) return;
          const p = JSON.parse(raw);
          if (typeof p.left==='number' && typeof p.top==='number'){
            el.style.left = `${p.left}px`;
            el.style.top  = `${p.top}px`;
            el.style.transform='none';
          }
        }catch{}
      }

      // ---------- libs loader ----------
      async function ensureLibs(){
        const load = (src, attrs={}) => new Promise((resolve, reject) => {
          const s = document.createElement('script');
          s.src = src; s.async = true; Object.entries(attrs).forEach(([k,v])=>s.setAttribute(k,v));
          s.onload = resolve; s.onerror = () => reject(new Error(`Failed to load ${src}`));
          document.head.appendChild(s);
        });

        if (!window.html2canvas) {
          await load('https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js');
        }
        if (!window.leafletImage) {
          await load('https://unpkg.com/leaflet-image@0.0.4/leaflet-image.js');
        }
        if (!(window.turf && turf.booleanPointInPolygon)) {
          await load('https://cdn.jsdelivr.net/npm/@turf/turf@6.5.0/turf.min.js');
        }
        // Optional boundary mask if you use it:
        // if (!(L.TileLayer && L.TileLayer.boundaryCanvas)) {
        //   await load('https://unpkg.com/leaflet-boundary-canvas@2.0.1/dist/leaflet-boundary-canvas.min.js');
        // }
      }

      // ---------- Drive Direct upload (GIS) ----------
      const driveAuth = { ready:false, tokenClient:null, accessToken:'', expiresAt:0, loading:false };
      async function ensureGIS(){
        if (driveAuth.ready) return;
        if (driveAuth.loading) { while(!driveAuth.ready) await new Promise(r=>setTimeout(r,50)); return; }
        driveAuth.loading = true;
        await new Promise((resolve, reject) => {
          const s = document.createElement('script');
          s.src = 'https://accounts.google.com/gsi/client';
          s.async = true;
          s.onload = () => resolve();
          s.onerror = () => reject(new Error('Failed to load Google Identity Services'));
          document.head.appendChild(s);
        });
        driveAuth.ready = true; driveAuth.loading = false;
      }
      function nowSec(){ return Math.floor(Date.now()/1000); }
      async function getAccessToken(clientId, forcePrompt=false){
        await ensureGIS();
        return new Promise((resolve, reject)=>{
          const scope = 'https://www.googleapis.com/auth/drive.file';
          if (!driveAuth.tokenClient) {
            driveAuth.tokenClient = google.accounts.oauth2.initTokenClient({
              client_id: clientId,
              scope,
              prompt: '',
              callback: (resp) => {
                if (resp && resp.access_token) {
                  driveAuth.accessToken = resp.access_token;
                  driveAuth.expiresAt = nowSec() + (resp.expires_in ? Math.floor(resp.expires_in*0.9) : 3200);
                  resolve(driveAuth.accessToken);
                } else { reject(new Error('No access token')); }
              }
            });
          }
          const needNew = !driveAuth.accessToken || nowSec() >= driveAuth.expiresAt || forcePrompt;
          try {
            driveAuth.tokenClient.requestAccessToken({ prompt: needNew ? 'consent' : '' });
          } catch (e) {
            try { driveAuth.tokenClient.requestAccessToken({ prompt: 'consent' }); }
            catch (err) { reject(err); }
          }
        });
      }
      function dataUrlToBlob(dataUrl){
        const [meta, b64] = String(dataUrl).split(',');
        const mime = /data:(.*?);base64/.exec(meta)?.[1] || 'application/octet-stream';
        const bin = atob(b64 || '');
        const len = bin.length;
        const bytes = new Uint8Array(len);
        for (let i=0; i<len; i++) bytes[i] = bin.charCodeAt(i);
        return new Blob([bytes], {type: mime});
      }
      async function ensureDriveFolder(accessToken, folderId, folderName){
        if (folderId) return folderId;
        const cacheKey = LS_KEYS.driveFolderCachePrefix + (folderName || DRIVE_FOLDER_DEFAULT_NAME);
        try {
          const cached = localStorage.getItem(cacheKey);
          if (cached) return cached;
        } catch {}

        // Create a folder (listing existing may not show results with drive.file scope)
        const meta = {
          name: folderName || DRIVE_FOLDER_DEFAULT_NAME,
          mimeType: 'application/vnd.google-apps.folder'
        };
        const res = await fetch('https://www.googleapis.com/drive/v3/files?fields=id,webViewLink', {
          method: 'POST',
          headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
          body: JSON.stringify(meta)
        });
        if (!res.ok) throw new Error(`Drive folder create failed HTTP ${res.status}`);
        const js = await res.json();
        const id = js.id;
        try { localStorage.setItem(cacheKey, id); } catch {}
        return id;
      }
      async function uploadToDriveDirect(opts){
        const { dataUrl, name, folderId, folderName, clientId } = opts || {};
        if (!dataUrl || !clientId) throw new Error('Drive Direct not configured.');
        const token = await getAccessToken(clientId);

        // Ensure folder
        const parentId = await ensureDriveFolder(token, folderId, folderName);

        // Build multipart/related
        const metadata = {
          name: name || 'snapshot.png',
          parents: parentId ? [parentId] : undefined
        };
        const boundary = '-------314159265358979323846';
        const delimiter = `\r\n--${boundary}\r\n`;
        const closeDelim = `\r\n--${boundary}--`;
        const pngBlob = dataUrlToBlob(dataUrl);

        const body = new Blob(
          [
            delimiter,
            'Content-Type: application/json; charset=UTF-8\r\n\r\n',
            JSON.stringify(metadata),
            delimiter,
            'Content-Type: image/png\r\n\r\n',
            pngBlob,
            closeDelim
          ],
          { type: `multipart/related; boundary=${boundary}` }
        );

        const res = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink', {
          method: 'POST',
          headers: { Authorization: `Bearer ${token}` },
          body
        });
        if (!res.ok) throw new Error(`Drive upload failed HTTP ${res.status}`);
        const js = await res.json();
        return js; // { id, webViewLink }
      }

      function downloadFallback(dataUrl, name){
        const a = document.createElement('a');
        a.href = dataUrl; a.download = name || 'snapshot.png';
        document.body.appendChild(a); a.click(); a.remove();
      }

      // finally, update diagnostics once at boot end
      updateDiagnostics();
    })();
  }

  // FIX B (Option 1): Top error UI helper (safe; no innerHTML, no external deps)
  function showTopError(title, msg){
    const n = document.getElementById('error'); if (!n) return;
    n.style.display='block';
    n.textContent = (title ? (title + ': ') : '') + (msg || '');
    setTimeout(()=>{ n.style.display='none'; }, 5000);
  }

  // Kick it off
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start);
  else start();
})();
