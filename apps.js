<!-- apps.js (vNext, banner-pos accurate + robust stats + save UX) -->
<script>
(async function () {
  // ---------- error surfacing ----------
  window.addEventListener('error', (e) => showTopError('Script error', (e && e.message) ? e.message : String(e)));
  window.addEventListener('unhandledrejection', (e) => showTopError('Promise error', (e?.reason?.message) || String(e?.reason || e)));

  // ---------- URL / config ----------
  const qs = new URLSearchParams(location.search);
  const cfgUrl = qs.get('cfg') || './config/app.config.json';
  const parseJSON = (s, fb=null) => { try { return JSON.parse(s); } catch { return fb; } };

  const driverMetaParam = parseJSON(qs.get('driverMeta') || '[]', []);
  const assignMapParam  = parseJSON(qs.get('assignMap')  || '{}', {});
  const cbUrl           = qs.get('cb') || '';
  const batchParam      = qs.get('batch'); // websafe b64 JSON
  const manualMode      = (qs.get('manual') === '1') || (!!batchParam && qs.get('auto') !== '1');

  // Upload params
  const driveFolderRaw  = qs.get('driveFolder') || qs.get('driveFolderId') || qs.get('driveFolderUrl') || (window.appUpload?.driveFolder) || '';
  const DRIVE_FOLDER_FALLBACK = '1ZzGt5_Kmgf1IUdiUdf_kN88lOscL792a';  // "Snapshots"
  const driveFolderId   = extractDriveFolderId(driveFolderRaw) || DRIVE_FOLDER_FALLBACK;
  const driveDirect     = qs.get('driveDirect') === '1';
  const gClientId       = qs.get('gClientId') || '';

  // ---------- runtime state ----------
  let batchItems = parseBatchItems(batchParam); // [{name, day, driver, keys[], stats, outName, view?}]
  let currentIndex = -1;     // -1 = overview
  let statsVisible = false;
  let manualSelectedKeys = null;
  let runtimeCustEnabled = null;

  let activeAssignMap = { ...assignMapParam };
  let driverMeta      = Array.isArray(driverMetaParam) ? [...driverMetaParam] : [];

  // --- Snapshot state ---
  let snapArmed = false;
  let snapEls = null;            // { overlay, helper, frame, flashCrop }
  let dockEls = null;            // { dock, img, name, saveBtn, dlBtn, exitBtn, closeBtn, title, note, link }
  let lastPngDataUrl = null;
  let lastSuggestedName = null;

  // ---------- map init ----------
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

  // ---------- collections / state ----------
  const baseDayLayers = [], quadDayLayers = [], subqDayLayers = [], allDaySets=[baseDayLayers,quadDayLayers,subqDayLayers];
  let allBounds=null, selectionBounds=null, hasSelection=false;

  const boundaryFeatures = [];
  const customerLayer = L.layerGroup().addTo(map);
  const customerMarkers = [];
  let customerCount=0, custWithinSel=0, custOutsideSel=0;
  const custByDayInSel = { Wednesday:0, Thursday:0, Friday:0, Saturday:0 };

  let currentFocus=null;
  let coveragePolysAll=[], coveragePolysSelected=[];
  let selectedMunicipalities=[];
  let driverSelectedCounts={}, driverOverlays={};

  let selectedOrderedKeys = [];
  let visibleSelectedKeysSet = new Set();

  renderLegend(cfg, null, custWithinSel, custOutsideSel, false);
  renderDriversPanel([], {}, false, {}, 0);

  // ---------- layers ----------
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

  // ---------- initial selection + customers ----------
  if (manualMode && batchItems.length) {
    manualSelectedKeys = unionAllKeys(batchItems);
    runtimeCustEnabled = false; // overview: customers off
  }
  await applySelection();
  await loadCustomersIfAny();

  // ---------- UI controls ----------
  if (manualMode && batchItems.length) {
    injectManualUI();
    ensureSnapshotUi();
    ensureDockUi();
    await zoomToOverview();
  } else if (!manualMode && batchItems.length) {
    await runAutoExport(batchItems);
  }

  // =================================================================
  // UI: Manual toolbar + Banner + Snapshot frame + Snapshot panel
  // =================================================================
  function injectManualUI(){
    const css = document.createElement('style');
    css.textContent = `
      .route-toolbar{position:absolute;top:10px;left:50%;transform:translateX(-50%);z-index:9000;display:flex;gap:8px}
      .route-toolbar button{background:#111;color:#fff;border:none;border-radius:6px;padding:8px 12px;font:600 14px system-ui;cursor:pointer;opacity:.95;transition:background .15s ease}
      .route-toolbar button:hover{opacity:1}
      .route-toolbar button.armed{background:#c62828}

      /* draggable banner */
      .dispatch-banner{position:fixed;left:50%;top:78vh;transform:translateX(-50%);z-index:8500;cursor:grab;
        background:rgba(255,255,255,.92);backdrop-filter:saturate(120%) blur(2px);
        border-radius:12px;box-shadow:0 8px 22px rgba(0,0,0,.14);
        padding:10px 14px;max-width:min(80vw,1100px);display:none}
      .dispatch-banner.dragging{cursor:grabbing; user-select:none}
      .dispatch-banner.visible{display:block}
      .dispatch-banner .row{font:500 14px system-ui;color:#111;line-height:1.35;margin:2px 0;word-wrap:break-word;overflow-wrap:break-word}
      .dispatch-banner .row strong{font-weight:700}

      /* snapshot framing UI */
      .snap-overlay{position:fixed;inset:0;z-index:9500;pointer-events:none}
      .snap-overlay .helper{position:absolute;top:10px;right:12px;pointer-events:auto;
        background:rgba(255,255,255,.95);border-radius:10px;box-shadow:0 6px 18px rgba(0,0,0,.12);
        padding:8px 10px;font:600 13px system-ui;color:#111;display:none}
      .snap-overlay .frame{position:absolute;border:2px dashed #111;background:rgba(0,0,0,.03);
        pointer-events:auto;border-radius:8px;box-shadow:inset 0 0 0 1px rgba(0,0,0,.06)}
      .snap-overlay .handle{position:absolute;width:14px;height:14px;background:#111;border-radius:3px}
      .snap-overlay .handle::after{content:'';position:absolute;inset:2px;border:2px solid #fff;border-radius:2px;opacity:.75}
      .snap-overlay .handle.nw{left:-7px;top:-7px} .snap-overlay .handle.ne{right:-7px;top:-7px}
      .snap-overlay .handle.sw{left:-7px;bottom:-7px} .snap-overlay .handle.se{right:-7px;bottom:-7px}
      .snap-overlay .handle.n{top:-7px;left:calc(50% - 7px)} .snap-overlay .handle.s{bottom:-7px;left:calc(50% - 7px)}
      .snap-overlay .handle.w{left:-7px;top:calc(50% - 7px)} .snap-overlay .handle.e{right:-7px;top:calc(50% - 7px)}

      /* cropped flash */
      .snap-flash-crop{position:fixed;background:#000;opacity:0;pointer-events:none;z-index:9600;transition:opacity .125s ease;border-radius:8px}

      /* persistent snapshot panel */
      .snap-dock{position:fixed;left:12px;bottom:12px;z-index:9605;background:rgba(17,17,17,.98);color:#fff;
        border-radius:12px;box-shadow:0 8px 24px rgba(0,0,0,.3);width:min(90vw,380px);display:none}
      .snap-dock .head{display:flex;align-items:center;gap:8px;padding:10px 12px;border-bottom:1px solid rgba(255,255,255,.12);cursor:grab}
      .snap-dock .title{font:600 14px system-ui}
      .snap-dock .spacer{flex:1}
      .snap-dock .x{background:transparent;border:0;color:#fff;font:700 16px system-ui;cursor:pointer;opacity:.9}
      .snap-dock .body{padding:10px 12px;display:grid;gap:8px}
      .snap-dock .preview{width:100%;max-height:42vh;object-fit:contain;border-radius:8px;background:#111}
      .snap-dock .row{display:flex;gap:8px;align-items:center}
      .snap-dock input[name="snapName"]{flex:1;min-width:0;background:#222;border:1px solid #444;color:#fff;border-radius:8px;padding:6px 8px;font:600 13px system-ui}
      .snap-dock .btn{background:#1e88e5;color:#fff;border:0;border-radius:8px;padding:8px 10px;font:700 13px system-ui;cursor:pointer}
      .snap-dock .btn.secondary{background:#424242}
      .snap-dock .btn:disabled{opacity:.5;cursor:not-allowed}
      .snap-dock .note{font:12px/1.3 system-ui;color:#cfd8dc}
      .snap-dock .link{font:12px/1.3 system-ui}
    `;
    document.head.appendChild(css);

    const bar = document.createElement('div');
    bar.className = 'route-toolbar';
    bar.innerHTML = `
      <button id="btnPrev" title="Previous (overview ↔ routes)" aria-label="Previous route">◀ Prev</button>
      <button id="btnNext" title="Next (overview ↔ routes)" aria-label="Next route">Next ▶</button>
      <button id="btnStats" title="Toggle dispatch banner" aria-label="Toggle stats">Stats</button>
      <button id="btnSnap"  title="Frame & capture snapshot" aria-label="Snapshot">📸 Snap</button>
    `;
    document.body.appendChild(bar);

    // Banner
    const banner = document.createElement('div');
    banner.id = 'dispatchBanner';
    banner.className = 'dispatch-banner';
    banner.innerHTML = `<div class="row r1"></div><div class="row r2"></div><div class="row r3"></div>`;
    document.body.appendChild(banner);
    restoreBannerPos(banner);
    makeBannerDraggable(banner);

    // Wire buttons
    document.getElementById('btnPrev').addEventListener('click', async () => { await stepRouteCycle(-1); });
    document.getElementById('btnNext').addEventListener('click', async () => { await stepRouteCycle(+1); });
    document.getElementById('btnStats').addEventListener('click', () => toggleStats());
    document.getElementById('btnSnap').addEventListener('click', onSnapClick);

    // keyboard
    window.addEventListener('keydown', async (e)=>{
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
  }

  // ---------- Stats text (robust + interprets sums) ----------
  function stringifyStat(val){
    if (val == null) return '';
    if (typeof val === 'number') return String(val);
    if (typeof val === 'string') return val.trim();
    if (Array.isArray(val)) return val.map(v => stringifyStat(v)).filter(Boolean).join(' | ');
    if (typeof val === 'object') {
      const entries = Object.entries(val).filter(([k,v]) => v != null && String(v).trim?.() !== '');
      if (!entries.length) return '';
      const pretty = entries.map(([k,v]) => `${k}: ${stringifyStat(v)}`).join(' | ');
      const valuesOnly = entries.map(([,v]) => stringifyStat(v)).join(' | ');
      return (pretty.length <= 64 ? pretty : valuesOnly).trim();
    }
    return String(val);
  }
  function trySumExpression(s){
    if (!s) return null;
    const t = String(s).trim();
    // support "12+1", "12 + 1", "12.5+1.25"
    if (!/[+\d]/.test(t)) return null;
    const parts = t.replace(/[^0-9.+-]/g,'').split('+').map(x=>x.trim()).filter(Boolean);
    if (parts.length < 2) return null;
    let sum = 0, any=false;
    for (const p of parts){
      const n = Number(p);
      if (Number.isFinite(n)) { sum += n; any = true; }
      else return null;
    }
    return any ? String(sum) : null;
  }
  function normalizeQty(s){
    const str = stringifyStat(s);
    const summed = trySumExpression(str);
    return summed || (str || '—');
  }
  function findByKeywords(obj, ...words){
    if (!obj) return '';
    const want = words.map(w => String(w).toLowerCase());
    const direct = scanObject(obj, (k,v)=> want.every(w => k.includes(w)));
    if (direct) return direct;
    // loose aliases
    const aliases = {
      base: ['basebox','baseboxes','base','regular'],
      custom: ['custom','customs','special','custom_boxes','customBoxes','customsText'],
      addon: ['add','addon','addons','add-ons','add_ons','extra','extras','addonsText']
    };
    const pool = new Set(aliases[words[0]] || []);
    const aliased = scanObject(obj, (k,v)=> [...pool].some(a => k.includes(a)));
    return aliased || '';
  }
  function scanObject(obj, pred){
    // depth-2 scan
    for (const [k,v] of Object.entries(obj||{})){
      const lk = k.toLowerCase();
      if (pred(lk, v)) { const s = normalizeQty(v); if (s) return s; }
      if (v && typeof v === 'object' && !Array.isArray(v)) {
        for (const [k2,v2] of Object.entries(v)){
          const lk2 = k2.toLowerCase();
          if (pred(lk2, v2)) { const s = normalizeQty(v2); if (s) return s; }
        }
      }
    }
    return '';
  }
  function getStatsTexts(statsObj) {
    const s = statsObj || {};
    const baseTxt = normalizeQty(findByKeywords(s, 'base', 'box') || s.baseBoxesText || s.base || s.baseBoxes);
    const custTxt = normalizeQty(findByKeywords(s, 'custom')      || s.customsText || s.customs);
    const addTxt  = normalizeQty(findByKeywords(s, 'addon')       || s.addOnsText  || s.addOns || s.add_ons);
    return { baseTxt, custTxt, addTxt };
  }

  function renderDispatchBanner(it){
    const banner = document.getElementById('dispatchBanner'); if (!banner) return;
    if (!it) { banner.classList.remove('visible'); return; }
    const s = it?.stats || {};
    const { baseTxt, custTxt, addTxt } = getStatsTexts(s);
    const row1 = `<strong>Day:</strong> ${escapeHtml(it.day||'')} &nbsp; - &nbsp; <strong>Driver:</strong> ${escapeHtml(it.driver||'')} &nbsp; - &nbsp; <strong>Route Name:</strong> ${escapeHtml(it.name||'')}`;
    const row2 = `<strong>Deliveries:</strong> ${escapeHtml(String(s.deliveries ?? 0))} &nbsp; - &nbsp; <strong>Apts:</strong> ${escapeHtml(String(s.apartments ?? 0))} &nbsp; - &nbsp; <strong>Base boxes:</strong> ${escapeHtml(baseTxt)}`;
    const row3 = `<strong>Customs:</strong> ${escapeHtml(custTxt)} &nbsp; - &nbsp; <strong>Add-ons:</strong> ${escapeHtml(addTxt)}`;
    banner.querySelector('.r1').innerHTML = row1;
    banner.querySelector('.r2').innerHTML = row2;
    banner.querySelector('.r3').innerHTML = row3;
    if (statsVisible) banner.classList.add('visible');
  }

  // Canvas banner (for saved PNG) — mirrors the DOM banner position & width
  function drawBannerOntoCanvas(canvas, it, frameRect){
    if (!it || !statsVisible) return;
    const domBanner = document.getElementById('dispatchBanner');
    if (!domBanner || !domBanner.classList.contains('visible')) return;

    const s = it?.stats || {};
    const { baseTxt, custTxt, addTxt } = getStatsTexts(s);

    const row1 = `Day: ${it.day||''}  -  Driver: ${it.driver||''}  -  Route Name: ${it.name||''}`;
    const row2 = `Deliveries: ${String(s.deliveries ?? 0)}  -  Apts: ${String(s.apartments ?? 0)}  -  Base boxes: ${baseTxt}`;
    const row3 = `Customs: ${custTxt}  -  Add-ons: ${addTxt}`;

    const bRect = domBanner.getBoundingClientRect();
    // position inside the cropped canvas
    let x = Math.round(bRect.left - frameRect.left);
    let y = Math.round(bRect.top  - frameRect.top);
    let boxW = Math.max(120, Math.min(canvas.width, Math.round(bRect.width)));

    // clamp within canvas
    if (x > canvas.width || y > canvas.height || (x + 10) < 0 || (y + 10) < 0) return;
    x = Math.max(0, Math.min(x, canvas.width - 10));
    y = Math.max(0, Math.min(y, canvas.height - 10));

    const ctx = canvas.getContext('2d');
    const pad = Math.max(8, Math.round(boxW * 0.04));
    const lineH = Math.max(16, Math.round(Math.min(canvas.width, canvas.height) * 0.027));

    const fBold = `700 ${Math.round(lineH*0.95)}px system-ui, -apple-system, Segoe UI, Roboto, sans-serif`;
    const fNorm = `600 ${Math.round(lineH*0.9)}px system-ui, -apple-system, Segoe UI, Roboto, sans-serif`;

    const wrap = (text, font) => {
      ctx.font = font;
      const words = String(text||'').split(/\s+/);
      let cur='', out=[];
      for (const w of words){
        const test = cur ? cur + ' ' + w : w;
        if (ctx.measureText(test).width <= boxW - pad*2) cur = test;
        else { if (cur) out.push(cur); cur = w; }
      }
      if (cur) out.push(cur);
      return out;
    };

    const w1 = wrap(row1, fBold);
    const w2 = wrap(row2, fNorm);
    const w3 = wrap(row3, fNorm);
    const lines = [...w1, ...w2, ...w3];
    const boxH = pad*2 + lines.length * lineH + Math.round(lineH*0.2);

    ctx.fillStyle = 'rgba(255,255,255,0.92)'; roundRect(ctx, x, y, boxW, boxH, 12).fill();

    let yy = y + pad + lineH; ctx.fillStyle='#111';
    ctx.font = fBold; for (const ln of w1) { ctx.fillText(ln, x+pad, yy); yy += lineH; }
    ctx.font = fNorm; for (const ln of w2) { ctx.fillText(ln, x+pad, yy); yy += lineH; }
    for (const ln of w3) { ctx.fillText(ln, x+pad, yy); yy += lineH; }
  }
  function roundRect(ctx, x, y, w, h, r) { ctx.beginPath(); ctx.moveTo(x + r, y); ctx.arcTo(x + w, y, x + w, y + h, r); ctx.arcTo(x + w, y + h, x, y + h, r); ctx.arcTo(x, y + h, x, y, r); ctx.arcTo(x, y, x + w, y, r); ctx.closePath(); return ctx; }

  // ---------- Zoom helpers ----------
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

  async function stepRouteCycle(delta){
    const N = batchItems.length; if (!N) return;

    if (delta > 0) {
      if (currentIndex === -1) currentIndex = 0;
      else if (currentIndex === N - 1) currentIndex = -1;
      else currentIndex += 1;
    } else {
      if (currentIndex === -1) currentIndex = N - 1;
      else if (currentIndex === 0) currentIndex = -1;
      else currentIndex -= 1;
    }

    if (currentIndex === -1) { await zoomToOverview(); return; }

    const it = batchItems[currentIndex];
    manualSelectedKeys = (it.keys || []).map(normalizeKey);
    runtimeCustEnabled = true;
    await applySelection();
    await loadCustomersIfAny();
    if (selectionBounds) fitWithHints(selectionBounds, (it && it.view) || null);
    renderDispatchBanner(it);
    updateButtons();
  }

  async function zoomToOverview(){
    currentIndex = -1;
    manualSelectedKeys = unionAllKeys(batchItems);
    runtimeCustEnabled = false; // hide customers
    statsVisible = false; renderDispatchBanner(null);
    await applySelection();
    await loadCustomersIfAny();
    if (selectionBounds) fitWithHints(selectionBounds, null);
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

  // =================================================================
  // SNAPSHOT: frame → capture → draw banner at DOM pos → preview panel
  // =================================================================
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

    const flashCrop = document.createElement('div'); // cropped flash
    flashCrop.className='snap-flash-crop';
    document.body.appendChild(flashCrop);

    snapEls = { overlay, helper: overlay.querySelector('.helper'), frame: overlay.querySelector('.frame'), flashCrop };
    bindFramingGestures(snapEls.frame);
  }

  function defaultFrameRect(){
    const vw = window.innerWidth, vh = window.innerHeight;
    const w = Math.round(vw*0.7), h = Math.round(vh*0.6);
    const l = Math.round((vw - w)/2), t = Math.round((vh - h)/2);
    return { left:l, top:t, width:w, height:h };
  }

  const LS_KEYS = {
    frame: 'dispatchViewer.snapFrameRect',
    banner:'dispatchViewer.bannerPos',
    dock:  'dispatchViewer.snapDockPos'
  };

  function setFrameRect(l,t,w,h){
    const f = snapEls.frame;
    Object.assign(f.style, {
      left: `${clamp(l,0,window.innerWidth-20)}px`,
      top: `${clamp(t,0,window.innerHeight-20)}px`,
      width: `${Math.max(40, Math.min(w, window.innerWidth))}px`,
      height:`${Math.max(40, Math.min(h, window.innerHeight))}px`
    });
  }
  function getFrameRect(){
    const f = snapEls.frame.getBoundingClientRect();
    return { left: Math.round(f.left), top: Math.round(f.top), width: Math.round(f.width), height: Math.round(f.height) };
  }
  function saveFrameRect(){
    const r = getFrameRect();
    try{ localStorage.setItem(LS_KEYS.frame, JSON.stringify(r)); }catch{}
  }
  function restoreFrameRect(){
    try{
      const raw = localStorage.getItem(LS_KEYS.frame);
      if (!raw) return false;
      const p = JSON.parse(raw);
      if (typeof p.left==='number' && typeof p.top==='number' && typeof p.width==='number' && typeof p.height==='number') {
        setFrameRect(p.left, p.top, p.width, p.height);
        return true;
      }
    }catch{}
    return false;
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
        const it = (currentIndex>=0 && currentIndex<batchItems.length) ? batchItems[currentIndex] : null;
        if (!it) throw new Error('No focused route to capture.');
        const r = getFrameRect();
        const canvas = await captureCanvas(r);   // DOM-first with Leaflet fallback
        drawBannerOntoCanvas(canvas, it, r);     // draw exactly where the user positioned it
        flashCropped(r);
        lastPngDataUrl = canvas.toDataURL('image/png');
        const rawName = it.outName || `${safeName(it.driver)}_${safeName(it.day)}.png`;
        lastSuggestedName = ensurePngExt(rawName);
        showDock(lastPngDataUrl, lastSuggestedName, it);
        saveFrameRect();
      } catch (e) {
        showDock(null, null, null, `Capture failed — ${escapeHtml(e && e.message || e)}`);
      } finally {
        exitFraming();
      }
    }
  }

  function cancelFraming(){
    if (!snapArmed) return;
    exitFraming();
    showDock(null, null, null, 'Framing cancelled.');
  }
  function exitFraming(){
    const btn = document.getElementById('btnSnap');
    if (btn) { btn.classList.remove('armed'); btn.setAttribute('aria-label','Frame & capture snapshot'); }
    if (snapEls) { snapEls.overlay.style.display='none'; snapEls.helper.style.display='none'; }
    snapArmed = false;
  }

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

  async function captureCanvas(r){
    // DOM-first, exclude banner & dock & overlay
    if (window.html2canvas) {
      try{
        return await html2canvas(document.body, {
          useCORS: true,
          backgroundColor: null,
          x: r.left, y: r.top, width: r.width, height: r.height,
          windowWidth: document.documentElement.clientWidth,
          windowHeight: document.documentElement.clientHeight,
          scrollX: 0, scrollY: 0,
          ignoreElements: (el) => {
            if (!el) return false;
            if (el.id === 'dispatchBanner' || el.closest?.('#dispatchBanner')) return true;
            const cls = el.classList || { contains:()=>false };
            return cls.contains('dispatch-banner') || cls.contains('snap-overlay') || cls.contains('snap-dock') || !!el.closest?.('.snap-dock');
          }
        });
      } catch (e) { console.warn('html2canvas failed; trying Leaflet fallback:', e); }
    }
    // Fallback: map-only
    if (!window.leafletImage) throw new Error('leaflet-image not loaded (fallback unavailable).');
    const mapRect = document.getElementById('map').getBoundingClientRect();
    const ix = intersectRects(r, {left:mapRect.left, top:mapRect.top, width:mapRect.width, height:mapRect.height});
    if (!ix || ix.width<=0 || ix.height<=0) throw new Error('Frame does not overlap the map; fallback capture cannot proceed.');
    const baseCanvas = await new Promise((resolve, reject)=>leafletImage(map, (err,c)=>err?reject(err):resolve(c)));
    const sx = Math.max(0, Math.round(ix.left - mapRect.left));
    const sy = Math.max(0, Math.round(ix.top  - mapRect.top));
    const sw = Math.round(ix.width), sh = Math.round(ix.height);
    const out = document.createElement('canvas'); out.width = sw; out.height = sh;
    out.getContext('2d').drawImage(baseCanvas, sx, sy, sw, sh, 0, 0, sw, sh);
    return out;
  }

  function flashCropped(r){
    if (!snapEls) return;
    const f = snapEls.flashCrop;
    Object.assign(f.style, { left:`${r.left}px`, top:`${r.top}px`, width:`${r.width}px`, height:`${r.height}px` });
    f.style.opacity = '0.65';
    setTimeout(()=>{ f.style.opacity = '0'; }, 125);
  }

  function intersectRects(a,b){
    const x1 = Math.max(a.left, b.left);
    const y1 = Math.max(a.top,  b.top);
    const x2 = Math.min(a.left+a.width,  b.left+b.width);
    const y2 = Math.min(a.top +a.height, b.top +b.height);
    if (x2<=x1 || y2<=y1) return null;
    return { left:x1, top:y1, width:x2-x1, height:y2-y1 };
  }

  // ---------- Snapshot Dock ----------
  function ensureDockUi(){
    if (dockEls) return;
    const dock = document.createElement('div');
    dock.className = 'snap-dock';
    dock.innerHTML = `
      <div class="head" id="snapDockHead" aria-grabbed="false">
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
      </div>
    `;
    document.body.appendChild(dock);

    const img = dock.querySelector('#snapDockImg');
    const name = dock.querySelector('#snapDockName');
    const saveBtn = dock.querySelector('#snapDockSave');
    const dlBtn = dock.querySelector('#snapDockDownload');
    const exitBtn = dock.querySelector('#snapDockExit');
    const closeBtn = dock.querySelector('#snapDockClose');
    const title = dock.querySelector('.title');
    const note = dock.querySelector('#snapDockNote');
    const link = dock.querySelector('#snapDockLink');

    makeDockDraggable(dock, dock.querySelector('#snapDockHead')); restoreDockPos(dock);

    saveBtn.addEventListener('click', async ()=>{
      if (!lastPngDataUrl) { note.textContent = 'No snapshot to save.'; return; }
      const ctx = currentIndex>=0 ? batchItems[currentIndex] : {};
      const outName = ensurePngExt((name.value || lastSuggestedName || 'snapshot.png').replace(/[^\w.-]+/g,'_'));
      const payload = { name: outName, day: ctx.day, driver: ctx.driver, routeName: ctx.name, pngBase64: lastPngDataUrl, folderId: driveFolderId };

      link.innerHTML = ''; title.textContent = 'Saving…'; note.textContent = '';
      saveBtn.disabled = true;

      try {
        let success = false, reply = null;

        // 1) Optional direct-to-Drive (OAuth client-side)
        if (driveDirect && gClientId) {
          try {
            reply = await uploadToDriveDirect({ dataUrl: lastPngDataUrl, name: outName, folderId: driveFolderId, clientId: gClientId });
            if (reply && reply.id) {
              success = true;
              const href = reply.webViewLink || `https://drive.google.com/file/d/${reply.id}/view`;
              title.textContent = 'Saved ✓';
              link.innerHTML = `<a href="${href}" target="_blank" rel="noopener">Open in Drive</a>`;
              note.innerHTML = `Saved to <b>Snapshots</b> as <b>${escapeHtml(outName)}</b>.`;
            }
          } catch (e) {
            console.warn('Direct Drive upload failed, will try webhook if available:', e);
          }
        }

        // 2) Webhook path
        if (!success && cbUrl) {
          reply = await saveViaWebhook(cbUrl, payload);
          if (reply && reply.ok) {
            success = true;
            const href = reply.webViewLink || reply.fileUrl || (reply.fileId ? `https://drive.google.com/file/d/${reply.fileId}/view` : '');
            title.textContent = 'Saved ✓';
            link.innerHTML = href ? `<a href="${href}" target="_blank" rel="noopener">Open in Drive</a>` : '';
            const folderName = reply.folder || 'Snapshots';
            note.innerHTML = `Saved to <b>${escapeHtml(folderName)}</b> as <b>${escapeHtml(outName)}</b>.`;
          }
        }

        if (!success) {
          title.textContent = 'Saved (unconfirmed)';
          const folderHref = `https://drive.google.com/drive/folders/${driveFolderId}`;
          note.innerHTML = cbUrl
            ? 'Save completed but no server ACK was readable. If this keeps happening, enable CORS on the webhook to return JSON.'
            : `No upload endpoint configured. Use Download or add <code>?cb=…</code>.`;
          link.innerHTML = `<a href="${folderHref}" target="_blank" rel="noopener">Open Snapshots folder</a>`;
        }
      } catch (err) {
        title.textContent = 'Save failed';
        note.textContent = `Upload failed — you can still Download. (${String(err && err.message || err)})`;
      } finally {
        saveBtn.disabled = false;
      }
    });

    dlBtn.addEventListener('click', ()=>{
      if (!lastPngDataUrl) { note.textContent = 'No snapshot to download.'; return; }
      const nm = ensurePngExt((name.value || lastSuggestedName || 'snapshot.png').replace(/[^\w.-]+/g,'_'));
      downloadFallback(lastPngDataUrl, nm);
      title.textContent = 'Downloaded ✓';
      note.textContent = `Downloaded as ${nm}.`;
    });

    const closeDock = ()=>{ dock.style.display='none'; lastPngDataUrl=null; lastSuggestedName=null; note.textContent=''; link.textContent=''; };
    exitBtn.addEventListener('click', closeDock);
    closeBtn.addEventListener('click', closeDock);

    dockEls = { dock, img, name, saveBtn, dlBtn, exitBtn, closeBtn, title, note, link };
  }

  function showDock(pngDataUrl, suggestedName){
    ensureDockUi();
    const { dock, img, name, title, note, link } = dockEls;
    title.textContent = 'Snapshot'; note.textContent = ''; link.textContent='';
    dock.style.display = 'block';
    if (pngDataUrl) img.src = pngDataUrl;
    if (suggestedName) name.value = suggestedName;
    saveDockPos(dock);
  }

  function makeDockDraggable(el, handle){
    let dragging=false, start={x:0,y:0, left:0, top:0};
    const onDown=(e)=>{ dragging=true; const r=el.getBoundingClientRect(); start={x:e.clientX,y:e.clientY,left:r.left,top:r.top}; window.addEventListener('pointermove',onMove,{passive:false}); window.addEventListener('pointerup',onUp,{once:true}); };
    const onMove=(e)=>{ if(!dragging) return; e.preventDefault(); const dx=e.clientX-start.x, dy=e.clientY-start.y; const L = clamp(start.left+dx, 0, window.innerWidth - el.offsetWidth); const T = clamp(start.top +dy, 0, window.innerHeight- el.offsetHeight); el.style.left = `${L}px`; el.style.top = `${T}px`; el.style.right='auto'; el.style.bottom='auto'; };
    const onUp=()=>{ dragging=false; saveDockPos(el); window.removeEventListener('pointermove',onMove); };
    handle.addEventListener('pointerdown', onDown);
  }
  function saveDockPos(el){ const r = el.getBoundingClientRect(); try{ localStorage.setItem(LS_KEYS.dock, JSON.stringify({ left:r.left, top:r.top })); }catch{} }
  function restoreDockPos(el){
    try{ const raw = localStorage.getItem(LS_KEYS.dock); if (!raw) return;
      const p = JSON.parse(raw);
      if (typeof p.left==='number' && typeof p.top==='number') { el.style.left = `${clamp(p.left,0,window.innerWidth - el.offsetWidth)}px`; el.style.top = `${clamp(p.top, 0,window.innerHeight- el.offsetHeight)}px`; el.style.right='auto'; el.style.bottom='auto'; }
    }catch{}
  }

  // ---------- banner drag helpers ----------
  function makeBannerDraggable(el){
    let dragging=false, start={x:0,y:0, left:0, top:0};
    const onDown=(e)=>{ dragging=true; el.classList.add('dragging'); const r=el.getBoundingClientRect(); start={x:e.clientX,y:e.clientY,left:r.left,top:r.top}; window.addEventListener('pointermove',onMove,{passive:false}); window.addEventListener('pointerup',onUp,{once:true}); };
    const onMove=(e)=>{ if(!dragging) return; e.preventDefault(); const dx=e.clientX-start.x, dy=e.clientY-start.y; const L = clamp(start.left+dx, 0, window.innerWidth - el.offsetWidth); const T = clamp(start.top +dy, 0, window.innerHeight- el.offsetHeight); el.style.left = `${L}px`; el.style.top = `${T}px`; el.style.transform=''; el.style.right='auto'; el.style.bottom='auto'; };
    const onUp=()=>{ dragging=false; el.classList.remove('dragging'); saveBannerPos(el); window.removeEventListener('pointermove',onMove); };
    el.addEventListener('pointerdown', onDown);
  }
  function saveBannerPos(el){
    const r = el.getBoundingClientRect();
    try{ localStorage.setItem(LS_KEYS.banner, JSON.stringify({ left:r.left, top:r.top })); }catch{}
  }
  function restoreBannerPos(el){
    try{
      const raw = localStorage.getItem(LS_KEYS.banner);
      if (!raw) return;
      const p = JSON.parse(raw);
      if (typeof p.left === 'number' && typeof p.top === 'number') {
        el.style.left = `${clamp(p.left,0,window.innerWidth - el.offsetWidth)}px`;
        el.style.top  = `${clamp(p.top, 0,window.innerHeight- el.offsetHeight)}px`;
        el.style.transform = '';
      }
    }catch{}
  }

  // ---------- helpers ----------
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

  function splitKeys(s) { return String(s||'').split(/[;,/|]/).map(x => x.trim()).filter(Boolean); }
  function unionAllKeys(items){ const set = new Set(); (items || []).forEach(it => (it.keys || []).forEach(k => set.add(normalizeKey(k)))); return Array.from(set); }

  function normalizeKey(s) {
    s = String(s || '').trim().toUpperCase();
    const m = s.match(/^([WTFS])0*(\d+)(_.+)?$/);
    if (m) return m[1] + String(parseInt(m[2], 10)) + (m[3] || '');
    return s;
  }
  function baseKeyFrom(key) { const m = String(key||'').toUpperCase().match(/^([WTFS]\d+)/); return m ? m[1] : String(key||'').toUpperCase(); }
  function quadParts(key) { const m = String(key || '').toUpperCase().match(/_(NE|NW|SE|SW)(?:_(TL|TR|LL|LR))?$/); return m ? { quad: m[1], sub: m[2] || null } : null; }
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

  function escapeHtml(s) { return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;'); }

  function smartTitleCase(s) {
    if (!s) return '';
    s = s.toLowerCase();
    const parts = s.split(/([\/-])/g);
    const small = new Set(['and','or','of','the','a','an','in','on','at','by','for','to','de','la','le','el','du','von','van','di','da','del']);
    const fixWord = (w, isFirst) => {
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

  function showTopError(title, msg) { const el = document.getElementById('error'); if (!el) return; el.style.display = 'block'; el.innerHTML = `<strong>${escapeHtml(title)}</strong><br>${escapeHtml(String(msg || ''))}`; }
  function renderError(msg) { showTopError('Load error', msg); }
  function phase(msg){ setStatus(`⏳ ${msg}`); }
  function info(msg){ setStatus(`ℹ️ ${msg}`); }
  function warn(msg){ showTopError('Warning', msg); }

  function setStatus(msg) { const n = document.getElementById('status'); if (n) n.textContent = msg || ''; }
  function makeStatusLine(selMunis, inCount, outCount, activeKeysLowerArray) {
    const muniList = (Array.isArray(selMunis) && selMunis.length) ? selMunis.join(', ') : '—';
    const keysTxt = (Array.isArray(activeKeysLowerArray) && activeKeysLowerArray.length) ? activeKeysLowerArray.join(', ') : '—';
    return `Customers (in/out): ${inCount}/${outCount} • Municipalities: ${muniList} • active zone keys: ${keysTxt}`;
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

  function colorFromName(name) { let h=0; const s=65, l=50; const str = String(name||''); for (let i=0;i<str.length;i++) h = (h*31 + str.charCodeAt(i)) % 360; return `hsl(${h}, ${s}%, ${l}%)`; }
  function clamp(v,min,max){ return Math.max(min, Math.min(max, v)); }
  const safeName = s => String(s||'').replace(/[^\w.-]+/g,'_');
  const ensurePngExt = s => /\.(png)$/i.test(s) ? s : (s.replace(/\.[a-z0-9]+$/i,'') + '.png');

  function downloadFallback(dataUrl, name){
    try{ const a = document.createElement('a'); a.href = dataUrl; a.download = name || 'snapshot.png'; document.body.appendChild(a); a.click(); setTimeout(()=>a.remove(), 0); }
    catch(e){ console.warn('downloadFallback failed:', e); }
  }

  // Load libs if missing
  async function ensureLibs(){
    const loadScript = (src) => new Promise((resolve, reject) => {
      if ([...document.scripts].some(s => s.src && s.src.includes(src))) return resolve();
      const el = document.createElement('script'); el.src = src; el.onload = resolve; el.onerror = reject; document.head.appendChild(el);
    });
    if (!window.html2canvas) await loadScript('https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js');
    if (!window.leafletImage) await loadScript('https://unpkg.com/leaflet-image/leaflet-image.js');
  }

  // -----------------------------------------------------------------
  // Upload helpers
  // -----------------------------------------------------------------
  async function saveViaWebhook(url, payload){
    try{
      const res = await fetch(url, { method:'POST', mode:'cors', headers:{'Content-Type':'application/json'}, body: JSON.stringify(payload) });
      const j = await res.json().catch(()=>null);
      if (j && (j.ok || j.success)) return { ok:true, ...j };
    } catch (e) {
      console.warn('CORS JSON upload failed; will try no-cors fallback', e);
    }
    try{
      await fetch(url, { method:'POST', mode:'no-cors', headers:{'Content-Type':'application/json'}, body: JSON.stringify(payload) });
      return { ok:false };
    } catch (e) { throw e; }
  }
  function extractDriveFolderId(str){
    if (!str) return '';
    const m = String(str).match(/\/folders\/([A-Za-z0-9_-]{10,})/);
    if (m) return m[1];
    if (/^[A-Za-z0-9_-]{10,}$/.test(str)) return str;
    return '';
  }

  // Optional direct-to-Drive client side
  async function uploadToDriveDirect({ dataUrl, name, folderId, clientId }){
    await ensureGoogleApis();
    const token = await getDriveToken(clientId);
    const blob = dataURLtoBlob(dataUrl);
    const metadata = { name, parents: folderId ? [folderId] : undefined, mimeType: 'image/png' };
    const boundary = '-------314159265358979323846';
    const delimiter = `\r\n--${boundary}\r\n`;
    const closeDelim = `\r\n--${boundary}--`;

    const body =
      delimiter + 'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
      JSON.stringify(metadata) +
      delimiter + 'Content-Type: image/png\r\n\r\n';

    const bodyUint8 = new Uint8Array(body.length + blob.size + closeDelim.length);
    let offset = 0;
    for (let i=0;i<body.length;i++) bodyUint8[offset++] = body.charCodeAt(i);
    const blobArray = await blob.arrayBuffer();
    bodyUint8.set(new Uint8Array(blobArray), offset); offset += blob.size;
    for (let i=0;i<closeDelim.length;i++) bodyUint8[offset++] = closeDelim.charCodeAt(i);

    const resp = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,webViewLink,parents', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'multipart/related; boundary=' + boundary },
      body: bodyUint8
    });
    if (!resp.ok) throw new Error(`Drive upload ${resp.status}`);
    return resp.json();
  }
  function dataURLtoBlob(dataUrl){
    const [meta, b64] = dataUrl.split(',');
    const mime = (meta.match(/data:([^;]+)/) || [])[1] || 'application/octet-stream';
    const bin = atob(b64);
    const len = bin.length; const arr = new Uint8Array(len);
    for (let i=0;i<len;i++) arr[i] = bin.charCodeAt(i);
    return new Blob([arr], { type: mime });
  }
  async function ensureGoogleApis(){
    const load = src => new Promise((res,rej)=>{
      if ([...document.scripts].some(s => s.src && s.src.includes(src))) return res();
      const el = document.createElement('script'); el.src = src; el.onload = res; el.onerror = rej; document.head.appendChild(el);
    });
    if (!window.google?.accounts?.oauth2) await load('https://accounts.google.com/gsi/client');
  }
  function getDriveToken(clientId){
    return new Promise((resolve, reject)=>{
      try{
        const tokenClient = google.accounts.oauth2.initTokenClient({
          client_id: clientId,
          scope: 'https://www.googleapis.com/auth/drive.file',
          callback: (t) => t && t.access_token ? resolve(t.access_token) : reject(new Error('No access token'))
        });
        tokenClient.requestAccessToken({ prompt: '' });
      }catch(e){ reject(e); }
    });
  }

  // =================================================================
  // Optional: legacy auto-export (only if auto=1 in URL)
  // =================================================================
  async function runAutoExport(items){
    const ids=['legend','drivers','status','error'];
    ids.forEach(id=>{ const n=document.getElementById(id); if(n) n.style.display='none'; });

    for (let i=0;i<items.length;i++){
      const it = items[i];
      try{
        manualSelectedKeys = (it.keys||[]).map(normalizeKey);
        runtimeCustEnabled = true;
        await applySelection();
        await loadCustomersIfAny();
        if (selectionBounds) fitWithHints(selectionBounds, (it && it.view) || null);
        await sleep(350);

        if (!window.leafletImage) await ensureLibs();
        const canvas = await new Promise((resolve, reject) => leafletImage(map, (err, c) => (err ? reject(err) : resolve(c))));
        drawBannerOntoCanvas(canvas, it, {left:0,top:0}); // entire map canvas
        const png = canvas.toDataURL('image/png');
        const name = ensurePngExt(it.outName || `${safeName(it.driver)}_${safeName(it.day)}.png`);
        if (cbUrl) {
          await saveViaWebhook(cbUrl, { name, day: it.day, driver: it.driver, routeName: it.name, pngBase64: png, folderId: driveFolderId });
        } else { downloadFallback(png, name); }
      }catch(e){ console.error('auto export item failed', e); }
    }
  }

  // =================================================================
  // LOADERS / SELECTION / CUSTOMERS — (unchanged)
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
    customerLayer.clearLayers();
    customerMarkers.length = 0;
    customerCount = 0;

    const custCfg = cfg.customers || {};
    const forceOff = (runtimeCustEnabled === false);

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
    const outStyle = { radius: cst.radius || 9, color: '#7a7a7a', weight: cst.weightPx || 2, opacity: 0.8, fillColor: '#c7c7c7', fillOpacity: 0.6 };

    const onlySelectedCustomers = manualMode && (currentIndex >= 0);

    for (const rec of customerMarkers) {
      let show = true, style = { ...outStyle }, insideSel = false, selDay = null;

      if (turfOn) {
        const pt = turf.point([rec.lng, rec.lat]);
        for (let j = 0; j < coveragePolysAll.length; j++) {
          if (turf.booleanPointInPolygon(pt, coveragePolysAll[j].feat)) {
            const lyr = coveragePolysAll[j].layerRef; if (lyr) lyr._custAny += 1; break;
          }
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
        const b = outline.getBounds(); const center = b && b.getCenter ? b.getCenter() : null;
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

  // ---------- POPUP / focus ----------
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

  // ---------- Drivers & Legend panels ----------
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
  function toggleDriverOverlay(name, on) { const rec = driverOverlays[name]; if (!rec) return; if (on) { rec.group.addTo(map); try { if (rec.labelMarker) rec.labelMarker.openTooltip(); } catch{} } else { map.removeLayer(rec.group); } }

  function renderLegend(cfg, legendCounts, custIn, custOut, outsideToggle) {
    const el = document.getElementById('legend'); if (!el) return;
    const rowsHtml = (cfg.layers || []).map(Lcfg => {
      const st = (cfg.style?.perDay?.[Lcfg.day]) || {};
      const c  = (legendCounts && legendCounts[Lcfg.day]) ? legendCounts[Lcfg.day] : { selected: 0, total: 0 };
      const frac = (c.total > 0) ? `${c.selected}/${c.total}` : '0/0';
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
  }

})().catch(e => {
  const el = document.getElementById('error');
  if (el) { el.style.display = 'block'; el.innerHTML = `<strong>Load error</strong><br>${String(e && e.message ? e.message : e)}`; }
  console.error(e);
});

// --------- small helpers kept outside IIFE for reuse in tests ----------
function headerIndexMap(hdrRow, schema) {
  const wantCoords = ((schema && schema.coords) || 'Verified Coordinates').toLowerCase();
  const wantNote   = ((schema && schema.note)   || 'Order Note').toLowerCase();
  const idx = (name) => Array.isArray(hdrRow)
    ? hdrRow.findIndex(h => (h || '').toLowerCase() === name)
    : -1;
  return { coords: idx(wantCoords), note: idx(wantNote) };
}
</script>
