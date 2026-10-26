/* Fuel Log — single-page app. Sign-in via MSAL (Entra ID), data via Microsoft Graph into a SharePoint site. */
(() => {
  "use strict";
  const CFG = window.FUEL_CONFIG || {};
  const SCOPES = ["User.Read", "Sites.Manage.All"];
  const LIST_NAME = "Fuel Log";
  const PHOTO_ROOT = "Food Photos";
  const PROGRESS_ROOT = "Progress Photos";
  const POSES = ["front", "side", "back"];

  // ---------- Meal plan (kcal / protein are rough plan estimates) ----------
  const MEALS = {
    b:   { slot: "Breakfast", name: "Green smoothie", what: "Lactose-free milk, whey isolate, spinach, berries, oats, chia", kcal: 450, protein: 40 },
    lA:  { slot: "Lunch", name: "Chicken rice bowl", what: "Cooked chicken, half rice pouch, leaves, cucumber, tomatoes", kcal: 480, protein: 45 },
    lB:  { slot: "Lunch", name: "Tuna jacket", what: "Jacket potato, tuna, light mayo, side salad", kcal: 420, protein: 35 },
    d1:  { slot: "Dinner", name: "Salmon + Tenderstem", what: "Air fryer 180°C, 12 min, half rice pouch", kcal: 550, protein: 40 },
    d2:  { slot: "Dinner", name: "Paprika chicken + beans", what: "Air fryer 200°C, 15 min, jacket potato", kcal: 520, protein: 48 },
    d3:  { slot: "Dinner", name: "Burgers + sweet potato fries", what: "5% beef, no bun, 200°C, 15 min, salad", kcal: 560, protein: 45 },
    dAny:{ slot: "Dinner", name: "Your pick", what: "Salmon, paprika chicken or burgers" }
  };
  const LUNCH  = ["lA", "lA", "lB", "lA", "lB", "lA", "lB"];   // by getDay(), 0 = Sunday
  const DINNER = ["dAny", "d1", "d2", "d3", "d1", "d2", "d3"];
  const OPTIONS = { b: ["b"], l: ["lA", "lB"], d: ["d1", "d2", "d3"] };
  const SLOT_FIELD = { Breakfast: "b", Lunch: "l", Dinner: "d", Snack: null };
  const FIELD_SLOT = { b: "Breakfast", l: "Lunch", d: "Dinner" };
  const ICONS = {
    tick: '<svg viewBox="0 0 24 24"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>',
    b: '<svg viewBox="0 0 24 24" style="stroke-width:1.8"><path d="M7 4h10l-1.2 15a2 2 0 0 1-2 1.8h-3.6a2 2 0 0 1-2-1.8z"/><path d="M7.5 9h9"/></svg>',
    l: '<svg viewBox="0 0 24 24" style="stroke-width:1.8"><path d="M4 12a8 8 0 0 0 16 0z"/><path d="M9 8c0-2 2-2 2-4M13 8c0-2 2-2 2-4"/></svg>',
    d: '<svg viewBox="0 0 24 24" style="stroke-width:1.8"><circle cx="12" cy="13" r="6"/><path d="M3 5v6a2 2 0 0 0 2 2M5 5v16M21 5c-2 0-3 2-3 5s1 3 3 3v8"/></svg>',
    s: '<svg viewBox="0 0 24 24" style="stroke-width:1.8"><path d="M12 7c-3-3-8-1-7 4 1 6 4 9 7 7 3 2 6-1 7-7 1-5-4-7-7-4z"/><path d="M12 7c0-2 1-3 3-4"/></svg>',
    busy: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="8" style="stroke-dasharray:12 38"><animateTransform attributeName="transform" type="rotate" from="0 12 12" to="360 12 12" dur="1s" repeatCount="indefinite"/></circle></svg>'
  };

  // ---------- Helpers ----------
  const $ = id => document.getElementById(id);
  const pad = n => String(n).padStart(2, "0");
  const isoDate = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const today = () => new Date();
  const todayKey = () => isoDate(today());
  const nowTime = (d = new Date()) => `${pad(d.getHours())}:${pad(d.getMinutes())}`;
  const fmt = n => Math.round(n).toLocaleString("en-GB");
  const blank = () => ({ b: false, l: false, d: false, gut: null, move: false, weight: null, waist: null, foods: [] });
  const totals = r => (r.foods || []).reduce((a, f) => ({ kcal: a.kcal + (f.calories || 0), protein: a.protein + (f.protein_g || 0) }), { kcal: 0, protein: 0 });
  const store = {
    get(k, f) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : f; } catch (e) { return f; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} }
  };
  const shortDate = d => new Date(d).toLocaleDateString("en-GB", { day: "numeric", month: "short" });

  let days = store.get("fuel.days", {});
  let itemIds = store.get("fuel.itemIds", {});
  let pending = new Set(store.get("fuel.pending", []));
  let ids = store.get("fuel.ids", null);
  let targets = store.get("fuel.targets", { kcal: 1700, protein: 150 });
  let todaysPhotos = [];
  let progress = [];
  let revealed = false;
  let range = 7;
  const save = () => { store.set("fuel.days", days); store.set("fuel.itemIds", itemIds); store.set("fuel.pending", [...pending]); };
  const getDay = k => Object.assign(blank(), days[k] || {});
  const complete = r => !!(r && r.b && r.l && r.d);
  const setStatus = t => { $("status").textContent = t; };

  // ---------- Auth ----------
  let pca, account;
  async function initAuth() {
    pca = new msal.PublicClientApplication({
      auth: { clientId: CFG.clientId, authority: `https://login.microsoftonline.com/${CFG.tenantId}`, redirectUri: window.location.origin, navigateToLoginRequestUrl: false },
      cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false }
    });
    const res = await pca.handleRedirectPromise();
    account = (res && res.account) || pca.getAllAccounts()[0] || null;
    if (account) pca.setActiveAccount(account);
    return account;
  }
  function signIn() { pca.loginRedirect({ scopes: SCOPES, prompt: "select_account" }); }
  async function token() {
    try { return (await pca.acquireTokenSilent({ scopes: SCOPES, account })).accessToken; }
    catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) await pca.acquireTokenRedirect({ scopes: SCOPES, account });
      throw e;
    }
  }

  // ---------- Graph ----------
  async function graph(path, opts = {}) {
    const t = await token();
    const res = await fetch(path.startsWith("http") ? path : "https://graph.microsoft.com/v1.0" + path, {
      ...opts,
      headers: { Authorization: "Bearer " + t, ...(opts.body && !(opts.body instanceof Blob) ? { "Content-Type": "application/json" } : {}), ...(opts.headers || {}) }
    });
    if (res.status === 204) return null;
    const txt = await res.text();
    const json = txt ? JSON.parse(txt) : null;
    if (!res.ok) { const err = new Error((json && json.error && json.error.message) || res.statusText); err.status = res.status; throw err; }
    return json;
  }

  async function ensureSetup() {
    if (ids && ids.siteId && ids.listId && ids.v === 3) return ids;
    const u = new URL(CFG.siteUrl);
    const site = await graph(`/sites/${u.hostname}:${u.pathname.replace(/\/$/, "")}`);
    const lists = await graph(`/sites/${site.id}/lists?$select=id,displayName&$top=200`);
    let list = (lists.value || []).find(l => l.displayName === LIST_NAME);
    if (!list) {
      list = await graph(`/sites/${site.id}/lists`, {
        method: "POST",
        body: JSON.stringify({
          displayName: LIST_NAME, list: { template: "genericList" },
          columns: [
            { name: "Breakfast", boolean: {} }, { name: "Lunch", boolean: {} }, { name: "Dinner", boolean: {} },
            { name: "Gut", text: {} }, { name: "Moved", boolean: {} }, { name: "WeightKg", number: { decimalPlaces: "one" } }
          ]
        })
      });
    }
    const cols = await graph(`/sites/${site.id}/lists/${list.id}/columns?$select=name&$top=200`);
    const have = new Set((cols.value || []).map(c => c.name));
    const wanted = [
      { name: "FoodsJson", text: { allowMultipleLines: true, linesForEditing: 6 } },
      { name: "Calories", number: { decimalPlaces: "none" } },
      { name: "ProteinG", number: { decimalPlaces: "none" } },
      { name: "WaistCm", number: { decimalPlaces: "one" } }
    ];
    for (const c of wanted) if (!have.has(c.name)) await graph(`/sites/${site.id}/lists/${list.id}/columns`, { method: "POST", body: JSON.stringify(c) });
    ids = { siteId: site.id, listId: list.id, v: 3 };
    store.set("fuel.ids", ids);
    return ids;
  }

  async function pull() {
    const { siteId, listId } = await ensureSetup();
    let url = `/sites/${siteId}/lists/${listId}/items?$expand=fields($select=Title,Breakfast,Lunch,Dinner,Gut,Moved,WeightKg,WaistCm,FoodsJson)&$top=500`;
    while (url) {
      const page = await graph(url);
      for (const it of page.value || []) {
        const f = it.fields || {};
        if (!f.Title || pending.has(f.Title)) continue;
        itemIds[f.Title] = it.id;
        let foods = [];
        try { foods = f.FoodsJson ? JSON.parse(f.FoodsJson) : []; } catch (e) {}
        days[f.Title] = {
          b: !!f.Breakfast, l: !!f.Lunch, d: !!f.Dinner, gut: f.Gut || null, move: !!f.Moved,
          weight: typeof f.WeightKg === "number" ? f.WeightKg : null, waist: typeof f.WaistCm === "number" ? f.WaistCm : null, foods
        };
      }
      url = page["@odata.nextLink"] || null;
    }
    save();
  }

  let flushing = null, flushTimer = null;
  function scheduleFlush() { clearTimeout(flushTimer); flushTimer = setTimeout(flush, 700); }
  async function flush() {
    if (flushing) return flushing;
    flushing = (async () => {
      try {
        const { siteId, listId } = await ensureSetup();
        for (const key of [...pending]) {
          const r = getDay(key), t = totals(r);
          const fields = { Title: key, Breakfast: r.b, Lunch: r.l, Dinner: r.d, Gut: r.gut || "", Moved: r.move,
            FoodsJson: JSON.stringify(r.foods || []), Calories: t.kcal, ProteinG: t.protein };
          if (typeof r.weight === "number") fields.WeightKg = r.weight;
          if (typeof r.waist === "number") fields.WaistCm = r.waist;
          if (itemIds[key]) await graph(`/sites/${siteId}/lists/${listId}/items/${itemIds[key]}/fields`, { method: "PATCH", body: JSON.stringify(fields) });
          else itemIds[key] = (await graph(`/sites/${siteId}/lists/${listId}/items`, { method: "POST", body: JSON.stringify({ fields }) })).id;
          pending.delete(key); save();
        }
        setStatus(navigator.onLine ? "Synced" : "Offline · saved on this phone");
      } catch (e) {
        setStatus("Couldn't sync just now · saved on this phone");
      } finally { flushing = null; }
    })();
    return flushing;
  }

  // ---------- State ----------
  function change(patch, key = todayKey()) {
    days[key] = Object.assign(getDay(key), patch);
    pending.add(key); save(); renderAll(); scheduleFlush();
  }
  const foodFor = name => (getDay(todayKey()).foods || []).find(f => f.file === name);
  const foodForField = field => (getDay(todayKey()).foods || []).find(f => f.key === field);
  const keyFor = (slot, time) => SLOT_FIELD[slot] || `snack-${(time || "").replace(":", "")}-${Date.now() % 100000}`;
  function upsertFood(rec, extraPatch = {}) {
    const foods = (getDay(todayKey()).foods || []).filter(f => f.key !== rec.key && !(rec.file && f.file === rec.file));
    foods.push(rec);
    const field = SLOT_FIELD[rec.slot];
    change({ ...(field ? { [field]: true } : {}), ...extraPatch, foods });
  }

  // ---------- AI ----------
  const blobToBase64 = blob => new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(String(r.result).split(",")[1]);
    r.onerror = () => reject(r.error);
    r.readAsDataURL(blob);
  });
  async function callAI(payload) {
    const res = await fetch("/api/analyze", {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: "Bearer " + await token() },
      body: JSON.stringify(payload)
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || "Couldn't get an estimate.");
    return data;
  }
  const announce = (name, data) => showToast(`${name}: ~${fmt(data.calories)} kcal, ${data.protein_g}g protein${data.note ? ". " + data.note : ""}`, null, 7000);

  async function analysePhoto(entry, blob) {
    entry.analysing = true; renderAll();
    try {
      const data = await callAI({ image: await blobToBase64(blob), slot: entry.slot });
      const name = data.items.map(i => i.name).slice(0, 3).join(", ") || "Photo";
      upsertFood({ key: entry.key, slot: entry.slot, time: entry.time, source: "photo", file: entry.name, name, base: { kcal: data.calories, protein: data.protein_g }, portion: 1, ...data });
      announce(name, data);
    } catch (e) {
      entry.aiError = e.message;
      showToast(/set up|ANTHROPIC/.test(e.message) ? e.message : "Photo saved, but I couldn't estimate it. Tap the meal to type it instead.", null, 5000);
    } finally { entry.analysing = false; renderAll(); }
  }
  async function logText(field, text) {
    const slot = FIELD_SLOT[field] || "Snack", time = nowTime();
    const key = field || keyFor("Snack", time);
    upsertFood({ key, slot, time, source: "text", name: text, calories: 0, protein_g: 0, flags: [], items: [], pendingAI: true });
    try {
      const data = await callAI({ text, slot });
      upsertFood({ key, slot, time, source: "text", name: text, base: { kcal: data.calories, protein: data.protein_g }, portion: 1, ...data });
      announce(text, data);
    } catch (e) {
      upsertFood({ key, slot, time, source: "text", name: text, calories: 0, protein_g: 0, flags: [], items: [] });
      showToast(`Logged "${text}", but I couldn't estimate it right now.`);
    }
  }
  function logPlan(field, mk) {
    const m = MEALS[mk];
    upsertFood({ key: field, slot: m.slot, time: nowTime(), source: "plan", name: m.name, base: { kcal: m.kcal, protein: m.protein }, portion: 1,
      calories: m.kcal, protein_g: m.protein, flags: [], items: [], note: "" });
  }
  function setPortion(key, p) {
    const foods = (getDay(todayKey()).foods || []).map(f => {
      if (f.key !== key) return f;
      const base = f.base || { kcal: f.calories, protein: f.protein_g };
      return { ...f, base, portion: p, calories: Math.round(base.kcal * p), protein_g: Math.round(base.protein * p) };
    });
    change({ foods });
  }
  function removeFood(key) {
    const foods = (getDay(todayKey()).foods || []).filter(f => f.key !== key);
    const patch = { foods };
    if (FIELD_SLOT[key]) patch[key] = false;
    change(patch);
  }

  // ---------- Photos ----------
  function slotForNow(d = new Date()) {
    const h = d.getHours() + d.getMinutes() / 60;
    if (h < 11) return "Breakfast";
    if (h < 15) return "Lunch";
    if (h < 17.5) return "Snack";
    return "Dinner";
  }
  function shrink(file, max = 1600) {
    return new Promise((resolve, reject) => {
      const url = URL.createObjectURL(file);
      const img = new Image();
      img.onload = () => {
        const s = Math.min(1, max / Math.max(img.naturalWidth, img.naturalHeight));
        const c = document.createElement("canvas");
        c.width = Math.round(img.naturalWidth * s); c.height = Math.round(img.naturalHeight * s);
        c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
        URL.revokeObjectURL(url);
        c.toBlob(b => b ? resolve(b) : reject(new Error("Couldn't read that photo")), "image/jpeg", 0.8);
      };
      img.onerror = () => { URL.revokeObjectURL(url); reject(new Error("Couldn't read that photo")); };
      img.src = url;
    });
  }
  const encPath = p => p.split("/").map(encodeURIComponent).join("/");

  async function loadPhotos() {
    try {
      const { siteId } = await ensureSetup();
      const res = await graph(`/sites/${siteId}/drive/root:/${encPath(PHOTO_ROOT + "/" + todayKey())}:/children?$expand=thumbnails&$select=id,name`);
      const remote = (res.value || []).map(f => {
        const m = f.name.match(/(Breakfast|Lunch|Dinner|Snack)/);
        const rec = foodFor(f.name);
        return { id: f.id, name: f.name, slot: m ? m[1] : "Snack", key: rec ? rec.key : null, thumb: f.thumbnails && f.thumbnails[0] && (f.thumbnails[0].medium || f.thumbnails[0].small || {}).url };
      });
      const busy = todaysPhotos.filter(p => p.local && !remote.some(r => r.name === p.name));
      todaysPhotos = remote.concat(busy).sort((a, b) => a.name.localeCompare(b.name));
    } catch (e) { if (e.status !== 404) console.warn(e); }
    renderAll();
  }

  let lastUpload = null, toastTimer = null;
  async function onPhoto(file, slotOverride) {
    if (!file) return;
    const now = new Date();
    const taken = file.lastModified && isoDate(new Date(file.lastModified)) === todayKey() ? new Date(file.lastModified) : now;
    const slot = slotOverride || slotForNow(taken);
    const time = nowTime(taken);
    const entry = { id: null, name: `${time.replace(":", "")}-${slot}.jpg`, thumb: URL.createObjectURL(file), local: true, slot, time };
    entry.key = keyFor(slot, time);
    todaysPhotos.push(entry); lastUpload = entry;
    const field = SLOT_FIELD[slot];
    if (field && !getDay(todayKey())[field]) change({ [field]: true }); else renderAll();
    showToast(slotOverride ? `Added to ${slot.toLowerCase()}. Working out calories…` : `Logged as ${slot.toLowerCase()}. Wrong? Tap to move it.`, slotOverride ? null : slot);
    let blob;
    try { blob = await shrink(file); } catch (e) { showToast(e.message); return; }
    try {
      const { siteId } = await ensureSetup();
      const res = await graph(`/sites/${siteId}/drive/root:/${encPath(PHOTO_ROOT + "/" + todayKey() + "/" + entry.name)}:/content?@microsoft.graph.conflictBehavior=rename`, {
        method: "PUT", body: blob, headers: { "Content-Type": "image/jpeg" }
      });
      entry.id = res.id; entry.name = res.name;
      if (!entry.name.includes(entry.slot)) {
        entry.name = entry.name.replace(/(Breakfast|Lunch|Dinner|Snack)/, entry.slot);
        await graph(`/sites/${siteId}/drive/items/${entry.id}`, { method: "PATCH", body: JSON.stringify({ name: entry.name }) }).catch(() => {});
      }
    } catch (e) {
      showToast("Photo didn't upload. Check your signal and try again.");
      todaysPhotos = todaysPhotos.filter(p => p !== entry); renderAll();
      return;
    }
    analysePhoto(entry, blob);
  }
  async function reslot(newSlot) {
    const entry = lastUpload;
    if (!entry || entry.slot === newSlot) return;
    const oldField = SLOT_FIELD[entry.slot], newField = SLOT_FIELD[newSlot];
    const oldName = entry.name, oldKey = entry.key;
    entry.slot = newSlot; entry.name = entry.name.replace(/(Breakfast|Lunch|Dinner|Snack)/, newSlot);
    entry.key = keyFor(newSlot, entry.time);
    let foods = getDay(todayKey()).foods || [];
    const rec = foods.find(f => f.key === oldKey || f.file === oldName);
    foods = foods.filter(f => f !== rec && (!newField || f.key !== newField));
    if (rec) foods.push({ ...rec, key: entry.key, slot: newSlot, file: entry.name });
    const patch = { foods };
    if (oldField && !foods.some(f => f.key === oldField)) patch[oldField] = false;
    if (newField) patch[newField] = true;
    change(patch);
    showToast(`Moved to ${newSlot.toLowerCase()}.`);
    if (entry.id) { try { const { siteId } = await ensureSetup(); await graph(`/sites/${siteId}/drive/items/${entry.id}`, { method: "PATCH", body: JSON.stringify({ name: entry.name }) }); } catch (e) {} }
  }
  function showToast(text, slot, ms) {
    $("toastText").textContent = text;
    const slots = $("toastSlots");
    slots.hidden = !slot;
    slots.querySelectorAll("button").forEach(b => b.classList.toggle("on", b.dataset.slot === slot));
    $("toast").hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { $("toast").hidden = true; }, ms || (slot ? 6000 : 2800));
  }

  // ---------- Progress photos ----------
  async function loadProgress() {
    try {
      const { siteId } = await ensureSetup();
      let url = `/sites/${siteId}/drive/root:/${encPath(PROGRESS_ROOT)}:/children?$expand=thumbnails&$select=id,name&$top=500`;
      const out = [];
      while (url) {
        const res = await graph(url);
        for (const f of res.value || []) {
          const m = f.name.match(/^(\d{4}-\d{2}-\d{2})-(front|side|back)/);
          const th = f.thumbnails && f.thumbnails[0];
          if (m && th) out.push({ date: m[1], pose: m[2], id: f.id, url: (th.large || th.medium || {}).url, thumb: (th.medium || th.small || {}).url });
        }
        url = res["@odata.nextLink"] || null;
      }
      progress = out.sort((a, b) => a.date.localeCompare(b.date));
    } catch (e) { if (e.status !== 404) console.warn(e); }
    renderBody();
  }
  async function uploadProgress(file, pose) {
    if (!file) return;
    showToast(`Saving ${pose} photo…`);
    try {
      const blob = await shrink(file, 2000);
      const { siteId } = await ensureSetup();
      await graph(`/sites/${siteId}/drive/root:/${encPath(PROGRESS_ROOT + "/" + todayKey() + "-" + pose + ".jpg")}:/content`, { method: "PUT", body: blob, headers: { "Content-Type": "image/jpeg" } });
      showToast(`${pose[0].toUpperCase() + pose.slice(1)} photo saved.`);
      await loadProgress();
    } catch (e) { showToast("Photo didn't save. Try again."); }
  }

  // ---------- Series helpers ----------
  function lastN(n) {
    const out = [];
    for (let i = n - 1; i >= 0; i--) { const d = today(); d.setDate(d.getDate() - i); out.push(isoDate(d)); }
    return out;
  }
  function series(field) { return Object.keys(days).sort().filter(k => typeof days[k][field] === "number").map(k => [k, days[k][field]]); }
  function streak() {
    let n = 0; const d = today();
    if (!complete(days[isoDate(d)])) d.setDate(d.getDate() - 1);
    while (complete(days[isoDate(d)])) { n++; d.setDate(d.getDate() - 1); }
    return n;
  }

  // ---------- Render: Today ----------
  function renderToday() {
    const now = today(), t = getDay(todayKey()), tot = totals(t);
    const hour = now.getHours();
    $("date").textContent = now.toLocaleDateString("en-GB", { weekday: "long", day: "numeric", month: "long" });
    $("headline").textContent = hour < 12 ? "Morning, George" : hour < 18 ? "Afternoon, George" : "Evening, George";

    // ring
    const left = targets.kcal - tot.kcal;
    const pct = Math.min(1, tot.kcal / targets.kcal);
    const ring = $("ringFill");
    ring.style.strokeDashoffset = String(326.7 * (1 - pct));
    ring.classList.toggle("over", left < 0);
    $("kcalLeft").textContent = fmt(Math.abs(left));
    $("kcalLeftLabel").textContent = left < 0 ? "kcal over" : "kcal left";
    $("kcalEaten").textContent = fmt(tot.kcal);
    $("kcalTarget").textContent = fmt(targets.kcal);
    $("proteinText").textContent = `${fmt(tot.protein)} / ${targets.protein} g`;
    $("proteinFill").style.width = Math.min(100, tot.protein / targets.protein * 100) + "%";

    // meals
    const wrap = $("meals"); wrap.textContent = "";
    const rows = [["b", "b"], ["l", LUNCH[now.getDay()]], ["d", DINNER[now.getDay()]]];
    rows.forEach(([field, mk]) => {
      const m = MEALS[mk], rec = foodForField(field);
      const busy = todaysPhotos.some(p => p.key === field && p.analysing) || (rec && rec.pendingAI);
      const done = !!t[field];
      const btn = document.createElement("button");
      btn.className = "meal" + (done ? " done" : "") + (busy ? " busy" : "");
      btn.innerHTML = `<span class="m-ic">${busy ? ICONS.busy : done ? ICONS.tick : ICONS[field]}</span><span class="m-txt"><span class="m-slot"></span><span class="m-name"></span></span><span class="m-kcal"></span>`;
      btn.querySelector(".m-slot").textContent = done ? m.slot : `${m.slot} · plan`;
      btn.querySelector(".m-name").textContent = busy ? "Working out calories…" : done ? (rec ? rec.name : "Logged") : m.name;
      const k = btn.querySelector(".m-kcal");
      if (done && rec && rec.calories) {
        k.innerHTML = `<b></b><span></span>`;
        k.querySelector("b").textContent = `${fmt(rec.calories)}`;
        k.querySelector("span").textContent = `kcal · ${rec.protein_g}g P`;
      } else if (!done) { k.innerHTML = '<span class="m-add">Log</span>'; }
      btn.onclick = () => openMealSheet(field, mk);
      wrap.appendChild(btn);
    });
    // snacks row
    const snacks = (t.foods || []).filter(f => String(f.key).startsWith("snack"));
    const sk = snacks.reduce((a, f) => a + (f.calories || 0), 0);
    const sb = document.createElement("button");
    sb.className = "meal" + (snacks.length ? " done" : "");
    sb.innerHTML = `<span class="m-ic">${snacks.length ? ICONS.tick : ICONS.s}</span><span class="m-txt"><span class="m-slot">Snacks & extras</span><span class="m-name"></span></span><span class="m-kcal"></span>`;
    sb.querySelector(".m-name").textContent = snacks.length ? snacks.map(s => s.name).join(", ") : "Anything else today";
    sb.querySelector(".m-kcal").innerHTML = snacks.length ? `<b>${fmt(sk)}</b><span>kcal</span>` : '<span class="m-add">Add</span>';
    sb.onclick = () => openMealSheet(null);
    wrap.appendChild(sb);

    const s = streak();
    $("streakText").textContent = s ? `${s}-day streak` : "";

    // photos
    $("photoSection").hidden = !todaysPhotos.length;
    $("photoCount").textContent = todaysPhotos.length ? `${todaysPhotos.length} today` : "";
    const box = $("photos"); box.textContent = "";
    todaysPhotos.forEach(p => {
      const el = document.createElement("div"); el.className = "photo";
      const img = document.createElement("img"); img.alt = ""; img.loading = "lazy"; if (p.thumb) img.src = p.thumb;
      const est = document.createElement("em"), sl = document.createElement("span");
      const food = foodFor(p.name);
      est.textContent = food ? `~${fmt(food.calories)} kcal` : p.analysing ? "Estimating…" : "";
      sl.textContent = (p.name.match(/(Breakfast|Lunch|Dinner|Snack)/) || [])[1] || p.slot || "";
      el.append(img, est, sl); box.appendChild(el);
    });

    // quiet pills
    const gp = $("pillGut");
    gp.className = "pill" + (t.gut ? " set " + t.gut : "");
    $("gutText").textContent = t.gut ? `Gut: ${t.gut}` : "Gut check";
    $("pillMove").className = "pill" + (t.move ? " set" : "");
    $("moveText").textContent = t.move ? "Workout done" : "Workout";
    const ws = series("weight");
    $("pillWeight").className = "pill" + (typeof t.weight === "number" ? " set" : "");
    $("weightText").textContent = typeof t.weight === "number" ? `${t.weight.toFixed(1)} kg` : now.getDay() === 1 ? "Weigh-in day" : ws.length ? `Weight · ${ws[ws.length - 1][1].toFixed(1)}` : "Weight";
  }

  // ---------- Charts ----------
  const NS = "http://www.w3.org/2000/svg";
  function svgEl(tag, attrs, text) { const e = document.createElementNS(NS, tag); for (const k in attrs) e.setAttribute(k, attrs[k]); if (text != null) e.textContent = text; return e; }
  function barChart(svg, keys, values, target, h) {
    svg.textContent = "";
    const W = 320, top = 10, bottom = h - 18, n = keys.length;
    const max = Math.max(target * 1.25, ...values, 1);
    const y = v => bottom - (v / max) * (bottom - top);
    const bw = (W / n) * (n > 14 ? 0.7 : 0.62);
    if (!values.some(v => v > 0)) { svg.appendChild(svgEl("text", { x: W / 2, y: h / 2, "text-anchor": "middle", class: "empty-msg" }, "Log some food to see this fill up")); }
    values.forEach((v, i) => {
      const cx = (i + .5) * (W / n);
      const hgt = v > 0 ? Math.max(3, bottom - y(v)) : 3;
      svg.appendChild(svgEl("rect", { x: cx - bw / 2, y: bottom - hgt, width: bw, height: hgt, rx: Math.min(4, bw / 2), class: "bar" + (v <= 0 ? " empty" : v > target * 1.08 ? " over" : "") }));
      if (n <= 7 || i % Math.ceil(n / 6) === 0 || i === n - 1) {
        const d = new Date(keys[i]);
        svg.appendChild(svgEl("text", { x: cx, y: h - 4, "text-anchor": "middle", class: "lbl" }, n <= 7 ? d.toLocaleDateString("en-GB", { weekday: "narrow" }) : d.getDate()));
      }
    });
    const ty = y(target);
    svg.appendChild(svgEl("line", { x1: 0, x2: W, y1: ty, y2: ty, class: "target" }));
    svg.appendChild(svgEl("text", { x: W, y: ty - 4, "text-anchor": "end", class: "lbl" }, `target ${fmt(target)}`));
  }
  function lineChart(svg, pts, h, unit) {
    svg.textContent = "";
    const W = 320, top = 12, bottom = h - 18;
    if (pts.length < 2) { svg.appendChild(svgEl("text", { x: W / 2, y: h / 2, "text-anchor": "middle", class: "empty-msg" }, pts.length ? "One more entry and you'll see a trend" : `No ${unit === "kg" ? "weigh-ins" : "measurements"} yet`)); return; }
    const vals = pts.map(p => p[1]);
    const min = Math.min(...vals) - 0.5, max = Math.max(...vals) + 0.5;
    const t0 = new Date(pts[0][0]).getTime(), t1 = new Date(pts[pts.length - 1][0]).getTime();
    const x = d => 8 + ((new Date(d).getTime() - t0) / Math.max(1, t1 - t0)) * (W - 16);
    const y = v => bottom - ((v - min) / (max - min)) * (bottom - top);
    [min + (max - min) * .25, min + (max - min) * .75].forEach(g => {
      svg.appendChild(svgEl("line", { x1: 0, x2: W, y1: y(g), y2: y(g), class: "grid" }));
      svg.appendChild(svgEl("text", { x: 0, y: y(g) - 3, class: "lbl" }, g.toFixed(1)));
    });
    const d = pts.map((p, i) => `${i ? "L" : "M"}${x(p[0]).toFixed(1)},${y(p[1]).toFixed(1)}`).join(" ");
    svg.appendChild(svgEl("path", { d: `${d} L${x(pts[pts.length - 1][0])},${bottom} L${x(pts[0][0])},${bottom}Z`, class: "area" }));
    svg.appendChild(svgEl("path", { d, class: "line" }));
    pts.forEach((p, i) => svg.appendChild(svgEl("circle", { cx: x(p[0]), cy: y(p[1]), r: i === pts.length - 1 ? 4 : 2.5, class: "pt" + (i === pts.length - 1 ? " last" : "") })));
    svg.appendChild(svgEl("text", { x: 8, y: h - 4, class: "lbl" }, shortDate(pts[0][0])));
    svg.appendChild(svgEl("text", { x: W - 8, y: h - 4, "text-anchor": "end", class: "lbl" }, shortDate(pts[pts.length - 1][0])));
  }

  // ---------- Render: Trends ----------
  function renderTrends() {
    const keys = lastN(range);
    const kc = keys.map(k => totals(getDay(k)).kcal), pr = keys.map(k => totals(getDay(k)).protein);
    const logged = kc.filter(v => v > 0);
    barChart($("chartCal"), keys, kc, targets.kcal, 140);
    barChart($("chartPro"), keys, pr, targets.protein, 120);
    const avg = logged.length ? logged.reduce((a, b) => a + b, 0) / logged.length : 0;
    $("calAvg").textContent = logged.length ? `${fmt(avg)} avg` : "–";
    const under = logged.filter(v => v <= targets.kcal * 1.05).length;
    $("calNote").textContent = logged.length ? `${under} of ${logged.length} logged days at or under target.${logged.length < keys.length ? " Days with nothing logged are left out of the average." : ""}` : "";
    const plog = pr.filter(v => v > 0);
    $("proAvg").textContent = plog.length ? `${fmt(plog.reduce((a, b) => a + b, 0) / plog.length)}g avg` : "–";

    const wAll = series("weight");
    const wpts = wAll.filter(([k]) => k >= lastN(Math.max(range, 30))[0]);
    lineChart($("chartWeight"), wpts, 120, "kg");
    if (wAll.length >= 2) {
      const diff = wAll[wAll.length - 1][1] - wAll[0][1];
      $("wDelta").textContent = `${diff <= 0 ? "−" : "+"}${Math.abs(diff).toFixed(1)} kg`;
      const weeks = Math.max(1, (new Date(wAll[wAll.length - 1][0]) - new Date(wAll[0][0])) / 604800000);
      $("wNote").textContent = `Since ${shortDate(wAll[0][0])}, about ${(Math.abs(diff) / weeks).toFixed(1)} kg a week ${diff <= 0 ? "down" : "up"}. Aim for 0.4–0.8 kg a week.`;
    } else { $("wDelta").textContent = "–"; $("wNote").textContent = "Weigh in once a week, same morning, after the loo."; }

    // gut
    const gs = $("gutStrip"); gs.textContent = ""; gs.style.setProperty("--n", Math.min(range, 15));
    const gutKeys = lastN(Math.min(range, 30));
    gutKeys.forEach(k => { const v = getDay(k).gut; const c = document.createElement("div"); c.className = "gs" + (v ? " " + v : ""); c.title = `${shortDate(k)}: ${v || "not logged"}`; gs.appendChild(c); });
    const gl = gutKeys.map(k => getDay(k).gut).filter(Boolean);
    $("gutKpi").textContent = gl.length ? `${gl.filter(v => v === "good").length}/${gl.length} good` : "–";
    $("gutInsight").textContent = gutInsight();

    // consistency
    const cs = $("consStrip"); cs.textContent = ""; cs.style.setProperty("--n", Math.min(range, 15));
    let full = 0;
    gutKeys.forEach(k => { const r = getDay(k); const n = ["b", "l", "d"].filter(f => r[f]).length; if (n === 3) full++; const c = document.createElement("div"); c.className = "cs" + (n ? " p" + n : ""); c.title = `${shortDate(k)}: ${n}/3 meals`; cs.appendChild(c); });
    $("consKpi").textContent = `${full}/${gutKeys.length} days`;
  }
  function gutInsight() {
    const all = Object.keys(days).sort();
    const bad = [], good = [];
    all.forEach(k => {
      const g = days[k].gut; if (!g) return;
      const prev = new Date(k); prev.setDate(prev.getDate() - 1);
      const p = days[isoDate(prev)]; if (!p || !(p.foods || []).length) return;
      const flags = new Set(); (p.foods || []).forEach(f => (f.flags || []).forEach(x => flags.add(x)));
      (g === "good" ? good : bad).push(flags);
    });
    if (bad.length + good.length < 4) return "Log your gut most mornings and your food the day before. After a week or so I'll point out what seems to set it off.";
    if (!bad.length) return "No rough mornings recently. Whatever you're doing, it's working.";
    const counts = {};
    bad.forEach(s => s.forEach(f => { counts[f] = (counts[f] || 0) + 1; }));
    const ranked = Object.entries(counts).map(([f, c]) => {
      const gr = good.filter(s => s.has(f)).length / Math.max(1, good.length);
      return [f, c, c / bad.length - gr];
    }).filter(r => r[2] > 0.2).sort((a, b) => b[2] - a[2]);
    if (!ranked.length) return `${bad.length} rough or meh mornings, with no clear food pattern yet. Stress and sleep can matter as much as food.`;
    const [f, c] = ranked[0];
    return `Rough or meh mornings followed days with ${f} ${c} of ${bad.length} times, more than your good days. Worth testing a week without it.`;
  }

  // ---------- Render: Body ----------
  function renderBody() {
    const box = $("progressThumbs"); box.textContent = "";
    const latest = {}; progress.forEach(p => { latest[p.pose] = p; });
    POSES.forEach(pose => {
      const p = latest[pose];
      const el = document.createElement("div"); el.className = "pthumb" + (revealed ? "" : " private");
      if (p && p.thumb) { const img = document.createElement("img"); img.alt = `${pose} progress photo`; img.src = p.thumb; el.appendChild(img); }
      else { const e = document.createElement("div"); e.className = "pempty"; e.textContent = "None yet"; el.appendChild(e); }
      const s = document.createElement("span"); s.textContent = pose + (p ? " · " + shortDate(p.date) : ""); el.appendChild(s);
      box.appendChild(el);
    });
    const last = progress.length ? progress[progress.length - 1].date : null;
    const since = last ? Math.floor((new Date(todayKey()) - new Date(last)) / 86400000) : null;
    $("progressHint").textContent = !last ? "Front, side and back. Same spot, same light, once a week." : since >= 7 ? `Last set ${since} days ago, so this week's are due.` : `Last set ${since === 0 ? "today" : since + " days ago"}. Next in ${7 - since} days.`;
    $("compareBtn").disabled = new Set(progress.map(p => p.date)).size < 2;
    $("revealBtn").textContent = revealed ? "Hide photos" : "Show photos";

    const w = series("weight"), wa = series("waist");
    $("bodyWeightV").textContent = w.length ? `${w[w.length - 1][1].toFixed(1)} kg` : "–";
    $("bodyWeightS").textContent = w.length ? `Logged ${shortDate(w[w.length - 1][0])}` : "Tap to log";
    $("bodyWaistV").textContent = wa.length ? `${wa[wa.length - 1][1].toFixed(1)} cm` : "–";
    $("bodyWaistS").textContent = wa.length ? `Logged ${shortDate(wa[wa.length - 1][0])}` : "Tap to log";
    lineChart($("chartWaist"), wa.slice(-20), 110, "cm");
    $("waistDelta").textContent = wa.length >= 2 ? `${wa[wa.length - 1][1] - wa[0][1] <= 0 ? "−" : "+"}${Math.abs(wa[wa.length - 1][1] - wa[0][1]).toFixed(1)} cm` : "–";
  }

  let current = "today";
  function renderAll() {
    renderToday();
    if (current === "trends") renderTrends();
    if (current === "body") renderBody();
  }
  function show(view) {
    current = view;
    ["today", "trends", "body"].forEach(v => { $("view-" + v).hidden = v !== view; });
    document.querySelectorAll(".tab[data-view]").forEach(t => { if (t.dataset.view === view) t.setAttribute("aria-current", "page"); else t.removeAttribute("aria-current"); });
    renderAll();
    window.scrollTo(0, 0);
  }

  // ---------- Sheets ----------
  const SHEETS = ["logSheet", "mealSheet", "gutSheet", "moveSheet", "numSheet", "compareSheet", "settingsSheet"];
  function openSheet(id) { SHEETS.forEach(s => { $(s).hidden = s !== id; }); $("scrim").hidden = false; }
  function closeSheets() { SHEETS.forEach(s => { $(s).hidden = true; }); $("scrim").hidden = true; }

  let sheetField = null;
  function openMealSheet(field, plannedKey) {
    sheetField = field;
    const rec = field ? foodForField(field) : null;
    $("mealTitle").textContent = field ? (rec ? FIELD_SLOT[field] : `What did you have for ${FIELD_SLOT[field].toLowerCase()}?`) : "Add a snack or extra";
    const lg = $("mealLogged"); lg.textContent = ""; lg.hidden = !rec;
    if (rec) {
      const b = document.createElement("b"); b.textContent = rec.name;
      const meta = document.createElement("div"); meta.className = "ln-meta";
      meta.textContent = rec.calories ? `~${fmt(rec.calories)} kcal · ${rec.protein_g}g protein${rec.flags && rec.flags.length ? " · contains " + rec.flags.join(", ") : ""}` : "No estimate yet";
      lg.append(b, meta);
      if (rec.note) { const n = document.createElement("div"); n.className = "ln-note"; n.textContent = rec.note; lg.appendChild(n); }
      if (rec.calories) {
        const row = document.createElement("div"); row.className = "portion";
        row.innerHTML = '<span class="muted small">Portion</span><div class="seg"><button data-p="0.5">½</button><button data-p="1">1×</button><button data-p="1.5">1½</button><button data-p="2">2×</button></div>';
        row.querySelectorAll("button").forEach(btn => {
          btn.classList.toggle("on", Number(btn.dataset.p) === (rec.portion || 1));
          btn.onclick = () => { setPortion(rec.key, Number(btn.dataset.p)); openMealSheet(field, plannedKey); };
        });
        lg.appendChild(row);
      }
    }
    const box = $("mealOpts"); box.textContent = "";
    (field ? OPTIONS[field] : []).forEach(mk => {
      const m = MEALS[mk];
      const b = document.createElement("button");
      b.className = "opt" + (rec && rec.source === "plan" && rec.name === m.name ? " on" : "");
      b.innerHTML = "<span><b></b><small></small></span>";
      b.querySelector("b").textContent = m.name;
      b.querySelector("small").textContent = `~${m.kcal} kcal · ${m.protein}g protein`;
      if (mk === plannedKey) { const s = document.createElement("span"); s.className = "badge"; s.textContent = "Today's plan"; b.appendChild(s); }
      b.onclick = () => { logPlan(field, mk); closeSheets(); showToast(`${m.name} logged: ~${m.kcal} kcal.`); };
      box.appendChild(b);
    });
    box.hidden = !field;
    $("mealText").value = "";
    $("mealText").placeholder = field ? (rec ? "Change it: type what you had…" : "Something else…") : "e.g. banana and a protein bar";
    $("mealNone").hidden = !rec;
    openSheet("mealSheet");
  }

  let numKind = "weight", numVal = 84;
  function openNum(kind) {
    numKind = kind;
    const t = getDay(todayKey()), s = series(kind);
    numVal = typeof t[kind] === "number" ? t[kind] : s.length ? s[s.length - 1][1] : kind === "weight" ? 84 : 90;
    $("numTitle").textContent = kind === "weight" ? "Log weight" : "Log waist";
    $("numUnit").textContent = kind === "weight" ? "kg" : "cm";
    $("numHint").textContent = kind === "weight" ? "Same time each week, first thing, after the loo." : "At your belly button, relaxed, tape level.";
    const steps = kind === "weight" ? [-1, -0.1, 0.1, 1] : [-1, -0.5, 0.5, 1];
    const box = $("numSteps"); box.textContent = "";
    steps.forEach(st => {
      const b = document.createElement("button"); b.textContent = (st > 0 ? "+" : "−") + Math.abs(st);
      b.onclick = () => { numVal = Math.max(30, Math.min(250, Math.round((numVal + st) * 10) / 10)); $("numValue").textContent = numVal.toFixed(1); };
      box.appendChild(b);
    });
    $("numValue").textContent = numVal.toFixed(1);
    openSheet("numSheet");
  }
  function openGut() {
    const g = getDay(todayKey()).gut;
    document.querySelectorAll("#gutSheet .chip").forEach(c => c.classList.toggle("on", c.dataset.v === g));
    openSheet("gutSheet");
  }
  function openCompare() {
    const dates = [...new Set(progress.map(p => p.date))];
    if (dates.length < 2) return;
    const fill = (sel, val) => { sel.textContent = ""; dates.forEach(d => { const o = document.createElement("option"); o.value = d; o.textContent = shortDate(d); sel.appendChild(o); }); sel.value = val; };
    fill($("cmpA"), dates[0]); fill($("cmpB"), dates[dates.length - 1]);
    $("cmpPose").value = "front";
    drawCompare();
    openSheet("compareSheet");
  }
  function drawCompare() {
    const pose = $("cmpPose").value;
    [["cmpImgA", $("cmpA").value], ["cmpImgB", $("cmpB").value]].forEach(([id, d]) => {
      const p = progress.find(x => x.date === d && x.pose === pose), img = $(id);
      if (p && p.url) { img.src = p.url; img.hidden = false; img.nextElementSibling.hidden = true; }
      else { img.hidden = true; img.nextElementSibling.hidden = false; }
    });
    const near = (d, field) => { const s = series(field).filter(([k]) => k <= d); return s.length ? s[s.length - 1][1] : null; };
    const bits = [];
    const wa = near($("cmpA").value, "weight"), wb = near($("cmpB").value, "weight");
    if (wa != null && wb != null) bits.push(`Weight ${wb - wa <= 0 ? "−" : "+"}${Math.abs(wb - wa).toFixed(1)} kg`);
    const xa = near($("cmpA").value, "waist"), xb = near($("cmpB").value, "waist");
    if (xa != null && xb != null) bits.push(`waist ${xb - xa <= 0 ? "−" : "+"}${Math.abs(xb - xa).toFixed(1)} cm`);
    $("cmpStats").textContent = bits.join(", ");
  }

  // ---------- Wire up ----------
  function wire() {
    document.querySelectorAll(".tab[data-view]").forEach(t => t.onclick = () => show(t.dataset.view));
    $("plusBtn").onclick = () => openSheet("logSheet");
    $("scrim").onclick = closeSheets;
    document.querySelectorAll("[data-close]").forEach(b => b.onclick = closeSheets);
    $("openSettings").onclick = () => {
      $("setKcal").value = targets.kcal; $("setPro").value = targets.protein;
      $("whoami").textContent = account ? `Signed in as ${account.username}` : "";
      openSheet("settingsSheet");
    };
    $("setSave").onclick = () => {
      const k = parseInt($("setKcal").value, 10), p = parseInt($("setPro").value, 10);
      if (k >= 1000 && k <= 4000) targets.kcal = k;
      if (p >= 50 && p <= 300) targets.protein = p;
      store.set("fuel.targets", targets); closeSheets(); renderAll(); showToast("Targets saved.");
    };
    $("signOut").onclick = () => pca.logoutRedirect({ account });

    const pick = (id, slotFn) => $(id).addEventListener("change", e => { const f = e.target.files && e.target.files[0]; e.target.value = ""; if (f) { const s = slotFn ? slotFn() : undefined; closeSheets(); onPhoto(f, s); } });
    pick("camera"); pick("lmCamera"); pick("lmGallery");
    const sheetSlot = () => sheetField ? FIELD_SLOT[sheetField] : "Snack";
    pick("mealCam", sheetSlot); pick("mealLib", sheetSlot);

    document.querySelectorAll("#logSheet [data-act]").forEach(b => b.onclick = () => {
      const a = b.dataset.act;
      if (a === "type") openMealSheet(null);
      if (a === "gut") openGut();
      if (a === "move") openSheet("moveSheet");
      if (a === "weight" || a === "waist") openNum(a);
      if (a === "body") { closeSheets(); show("body"); }
    });
    $("mealForm").addEventListener("submit", e => {
      e.preventDefault();
      const text = $("mealText").value.trim();
      if (!text) { $("mealText").focus(); return; }
      const f = sheetField; closeSheets(); logText(f, text.slice(0, 120));
    });
    $("mealNone").onclick = () => {
      const rec = sheetField ? foodForField(sheetField) : null;
      closeSheets();
      if (rec) { removeFood(rec.key); showToast(`${FIELD_SLOT[sheetField] || "Item"} removed.`); }
    };
    document.querySelectorAll("#gutSheet .chip").forEach(c => c.onclick = () => {
      const v = getDay(todayKey()).gut === c.dataset.v ? null : c.dataset.v;
      change({ gut: v }); closeSheets(); if (v) showToast(v === "good" ? "Nice. Logged." : "Logged. I'll look for patterns.");
    });
    $("moveDone").onclick = () => { change({ move: true }); closeSheets(); showToast("Workout logged. Good work."); };
    $("moveUndo").onclick = () => { change({ move: false }); closeSheets(); };
    $("numSave").onclick = () => { change({ [numKind]: numVal }); closeSheets(); showToast(`Saved ${numVal.toFixed(1)} ${numKind === "weight" ? "kg" : "cm"}.`); };
    $("pillGut").onclick = openGut;
    $("pillMove").onclick = () => openSheet("moveSheet");
    $("pillWeight").onclick = () => openNum("weight");
    $("bodyWeight").onclick = () => openNum("weight");
    $("bodyWaist").onclick = () => openNum("waist");

    POSES.forEach(pose => $("pp-" + pose).addEventListener("change", e => { const f = e.target.files && e.target.files[0]; e.target.value = ""; uploadProgress(f, pose); }));
    $("revealBtn").onclick = () => { revealed = !revealed; renderBody(); };
    $("compareBtn").onclick = openCompare;
    ["cmpA", "cmpB", "cmpPose"].forEach(id => $(id).onchange = drawCompare);

    document.querySelectorAll("#rangeSeg button").forEach(b => b.onclick = () => {
      range = Number(b.dataset.range);
      document.querySelectorAll("#rangeSeg button").forEach(x => x.setAttribute("aria-selected", String(x === b)));
      renderTrends();
    });
    $("toastSlots").querySelectorAll("button").forEach(b => b.onclick = () => reslot(b.dataset.slot));
    window.addEventListener("online", () => { if (pending.size) flush(); });
    document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") { renderAll(); refresh(); } });
  }

  let lastRefresh = 0;
  async function refresh() {
    if (Date.now() - lastRefresh < 30000) return;
    lastRefresh = Date.now();
    try { if (pending.size) await flush(); await pull(); renderAll(); loadPhotos(); loadProgress(); setStatus("Synced"); }
    catch (e) { setStatus(navigator.onLine ? "Couldn't reach SharePoint · saved on this phone" : "Offline · saved on this phone"); }
  }

  async function start() {
    if ("serviceWorker" in navigator) navigator.serviceWorker.register("/sw.js").catch(() => {});
    if (!CFG.clientId || !CFG.tenantId || !CFG.siteUrl) {
      $("gate").hidden = false; $("signIn").hidden = true;
      $("gateText").textContent = "Setup isn't finished: config.js needs the client ID, tenant ID and SharePoint site URL.";
      return;
    }
    let acct = null;
    try { acct = await initAuth(); } catch (e) { console.error(e); }
    if (!acct) { $("gate").hidden = false; $("signIn").onclick = signIn; return; }
    $("app").hidden = false;
    wire(); renderAll(); renderBody();
    // Deep links from notifications: #gut, #weight, #waist, #move, #snap, #body, #trends, #b, #l, #d
    const h = location.hash.replace("#", "");
    if (h === "trends" || h === "body") show(h);
    else if (h === "gut") openGut();
    else if (h === "move") openSheet("moveSheet");
    else if (h === "weight" || h === "waist") openNum(h);
    else if (["b", "l", "d"].includes(h)) openMealSheet(h, h === "b" ? "b" : h === "l" ? LUNCH[today().getDay()] : DINNER[today().getDay()]);
    else if (h === "log") openSheet("logSheet");
    if (h) history.replaceState(null, "", "/");
    setStatus("Syncing…");
    refresh();
  }
  start();
})();
