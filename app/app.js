/* Fuel Log — single-page app. Sign-in via MSAL (Entra ID), data via Microsoft Graph into a SharePoint site. */
(() => {
  "use strict";
  const CFG = window.FUEL_CONFIG || {};
  const SCOPES = ["User.Read", "Sites.Manage.All"];
  const LIST_NAME = "Fuel Log";
  const PHOTO_ROOT = "Food Photos";

  // ---------- Meal plan ----------
  const MEALS = {
    b:   { slot: "Breakfast", name: "Green smoothie", what: "Lactose-free milk, whey isolate, spinach, berries, 30g oats, chia" },
    lA:  { slot: "Lunch", name: "Chicken rice bowl", what: "Cooked chicken, half rice pouch, leaves, cucumber, tomatoes" },
    lB:  { slot: "Lunch", name: "Tuna jacket", what: "Microwave jacket, tuna, 1 tbsp light mayo, side salad" },
    d1:  { slot: "Dinner", name: "Salmon + Tenderstem", what: "Air fryer 180°C, 12 min. Half rice pouch" },
    d2:  { slot: "Dinner", name: "Paprika chicken + beans", what: "Air fryer 200°C, 15 min. Jacket potato" },
    d3:  { slot: "Dinner", name: "Burgers + sweet potato fries", what: "5% beef, no bun. Air fryer 200°C, 15 min. Salad" },
    dAny:{ slot: "Dinner", name: "Your pick", what: "Salmon, paprika chicken or burgers" }
  };
  const LUNCH  = ["lA", "lA", "lB", "lA", "lB", "lA", "lB"];      // indexed by getDay(), 0 = Sunday
  const DINNER = ["dAny", "d1", "d2", "d3", "d1", "d2", "d3"];
  const SLOT_FIELD = { Breakfast: "b", Lunch: "l", Dinner: "d", Snack: null };

  // ---------- Helpers ----------
  const $ = id => document.getElementById(id);
  const pad = n => String(n).padStart(2, "0");
  const isoDate = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const today = () => new Date();
  const todayKey = () => isoDate(today());
  const blank = () => ({ b: false, l: false, d: false, gut: null, move: false, weight: null, foods: [] });
  const totals = r => (r.foods || []).reduce((a, f) => ({ kcal: a.kcal + (f.calories || 0), protein: a.protein + (f.protein_g || 0) }), { kcal: 0, protein: 0 });
  const store = {
    get(k, f) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : f; } catch (e) { return f; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} }
  };

  let days = store.get("fuel.days", {});          // "YYYY-MM-DD" -> {b,l,d,gut,move,weight}
  let itemIds = store.get("fuel.itemIds", {});    // "YYYY-MM-DD" -> SharePoint list item id
  let pending = new Set(store.get("fuel.pending", []));
  let ids = store.get("fuel.ids", null);          // {siteId, listId}
  let todaysPhotos = [];
  const save = () => { store.set("fuel.days", days); store.set("fuel.itemIds", itemIds); store.set("fuel.pending", [...pending]); };
  const getDay = k => Object.assign(blank(), days[k] || {});
  const complete = r => !!(r && r.b && r.l && r.d);
  const setStatus = t => { $("status").textContent = t; };

  // ---------- Auth ----------
  let pca, account;
  async function initAuth() {
    pca = new msal.PublicClientApplication({
      auth: {
        clientId: CFG.clientId,
        authority: `https://login.microsoftonline.com/${CFG.tenantId}`,
        redirectUri: window.location.origin,
        navigateToLoginRequestUrl: false
      },
      cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false }
    });
    const res = await pca.handleRedirectPromise();
    account = (res && res.account) || pca.getAllAccounts()[0] || null;
    if (account) pca.setActiveAccount(account);
    return account;
  }
  function signIn() { pca.loginRedirect({ scopes: SCOPES, prompt: "select_account" }); }
  async function token() {
    try {
      const r = await pca.acquireTokenSilent({ scopes: SCOPES, account });
      return r.accessToken;
    } catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) { await pca.acquireTokenRedirect({ scopes: SCOPES, account }); }
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
    if (ids && ids.siteId && ids.listId && ids.v === 2) return ids;
    const u = new URL(CFG.siteUrl);
    const site = await graph(`/sites/${u.hostname}:${u.pathname.replace(/\/$/, "")}`);
    const lists = await graph(`/sites/${site.id}/lists?$select=id,displayName&$top=200`);
    let list = (lists.value || []).find(l => l.displayName === LIST_NAME);
    if (!list) {
      list = await graph(`/sites/${site.id}/lists`, {
        method: "POST",
        body: JSON.stringify({
          displayName: LIST_NAME,
          list: { template: "genericList" },
          columns: [
            { name: "Breakfast", boolean: {} },
            { name: "Lunch", boolean: {} },
            { name: "Dinner", boolean: {} },
            { name: "Gut", text: {} },
            { name: "Moved", boolean: {} },
            { name: "WeightKg", number: { decimalPlaces: "one" } }
          ]
        })
      });
    }
    // v2 columns: AI food estimates and day totals
    const cols = await graph(`/sites/${site.id}/lists/${list.id}/columns?$select=name&$top=200`);
    const have = new Set((cols.value || []).map(c => c.name));
    const wanted = [
      { name: "FoodsJson", text: { allowMultipleLines: true, linesForEditing: 6 } },
      { name: "Calories", number: { decimalPlaces: "none" } },
      { name: "ProteinG", number: { decimalPlaces: "none" } }
    ];
    for (const c of wanted) if (!have.has(c.name)) await graph(`/sites/${site.id}/lists/${list.id}/columns`, { method: "POST", body: JSON.stringify(c) });
    ids = { siteId: site.id, listId: list.id, v: 2 };
    store.set("fuel.ids", ids);
    return ids;
  }

  async function pull() {
    const { siteId, listId } = await ensureSetup();
    let url = `/sites/${siteId}/lists/${listId}/items?$expand=fields($select=Title,Breakfast,Lunch,Dinner,Gut,Moved,WeightKg,FoodsJson)&$top=500`;
    while (url) {
      const page = await graph(url);
      for (const it of page.value || []) {
        const f = it.fields || {};
        if (!f.Title || pending.has(f.Title)) continue;   // local unsynced edits win
        itemIds[f.Title] = it.id;
        let foods = [];
        try { foods = f.FoodsJson ? JSON.parse(f.FoodsJson) : []; } catch (e) {}
        days[f.Title] = { b: !!f.Breakfast, l: !!f.Lunch, d: !!f.Dinner, gut: f.Gut || null, move: !!f.Moved, weight: typeof f.WeightKg === "number" ? f.WeightKg : null, foods };
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
          const r = getDay(key);
          const fields = { Title: key, Breakfast: r.b, Lunch: r.l, Dinner: r.d, Gut: r.gut || "", Moved: r.move };
          if (typeof r.weight === "number") fields.WeightKg = r.weight;
          if (r.foods && r.foods.length) {
            const t = totals(r);
            fields.FoodsJson = JSON.stringify(r.foods); fields.Calories = t.kcal; fields.ProteinG = t.protein;
          }
          if (itemIds[key]) {
            await graph(`/sites/${siteId}/lists/${listId}/items/${itemIds[key]}/fields`, { method: "PATCH", body: JSON.stringify(fields) });
          } else {
            const created = await graph(`/sites/${siteId}/lists/${listId}/items`, { method: "POST", body: JSON.stringify({ fields }) });
            itemIds[key] = created.id;
          }
          pending.delete(key); save();
        }
        setStatus(navigator.onLine ? "Synced to SharePoint." : "Offline. Saved on this phone.");
      } catch (e) {
        setStatus("Couldn't sync just now. Saved on this phone and will retry.");
      } finally { flushing = null; }
    })();
    return flushing;
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
  const photoPath = (date, name) => `${PHOTO_ROOT}/${date}/${name}`;
  const encPath = p => p.split("/").map(encodeURIComponent).join("/");

  async function loadPhotos() {
    try {
      const { siteId } = await ensureSetup();
      const res = await graph(`/sites/${siteId}/drive/root:/${encPath(PHOTO_ROOT + "/" + todayKey())}:/children?$expand=thumbnails&$select=id,name&$orderby=name`);
      todaysPhotos = (res.value || []).map(f => ({ id: f.id, name: f.name, thumb: f.thumbnails && f.thumbnails[0] && (f.thumbnails[0].medium || f.thumbnails[0].small || {}).url }));
    } catch (e) {
      if (e.status !== 404) console.warn(e);
      todaysPhotos = todaysPhotos.filter(p => p.local);
    }
    renderPhotos();
  }

  const blobToBase64 = blob => new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(String(r.result).split(",")[1]);
    r.onerror = () => reject(r.error);
    r.readAsDataURL(blob);
  });
  const foodFor = name => (getDay(todayKey()).foods || []).find(f => f.file === name);

  async function analyse(entry, blob) {
    entry.analysing = true; renderPhotos();
    try {
      const res = await fetch("/api/analyze", {
        method: "POST",
        headers: { "Content-Type": "application/json", Authorization: "Bearer " + await token() },
        body: JSON.stringify({ image: await blobToBase64(blob), slot: entry.slot })
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(data.error || "Couldn't analyse that photo.");
      const food = { file: entry.name, slot: entry.slot, time: entry.time, ...data };
      const foods = (getDay(todayKey()).foods || []).filter(f => f.file !== entry.name).concat(food);
      change({ foods });
      const what = data.items.map(i => i.name).slice(0, 3).join(", ");
      if (what) showToast(`${what} · ~${data.calories} kcal, ${data.protein_g}g protein${data.note ? ". " + data.note : ""}`, null, 7000);
    } catch (e) {
      entry.aiError = e.message;
    } finally { entry.analysing = false; renderPhotos(); }
  }

  let lastUpload = null, toastTimer = null;
  async function onPhoto(file) {
    if (!file) return;
    const now = new Date();
    const slot = slotForNow(now);
    const local = URL.createObjectURL(file);
    const entry = { id: null, name: `${pad(now.getHours())}${pad(now.getMinutes())}-${slot}.jpg`, thumb: local, local: true, slot, time: `${pad(now.getHours())}:${pad(now.getMinutes())}` };
    todaysPhotos.push(entry); lastUpload = entry; renderPhotos();
    const field = SLOT_FIELD[slot];
    if (field && !getDay(todayKey())[field]) change({ [field]: true });
    showToast(`Logged as ${slot}. Wrong? Tap to change.`, slot);
    let blob;
    try { blob = await shrink(file); } catch (e) { showToast(e.message); return; }
    try {
      const { siteId } = await ensureSetup();
      const res = await graph(`/sites/${siteId}/drive/root:/${encPath(photoPath(todayKey(), entry.name))}:/content?@microsoft.graph.conflictBehavior=rename`, {
        method: "PUT", body: blob, headers: { "Content-Type": "image/jpeg" }
      });
      entry.id = res.id; entry.name = res.name;
      if (!entry.name.includes(entry.slot)) {   // re-slotted while uploading
        entry.name = entry.name.replace(/(Breakfast|Lunch|Dinner|Snack)/, entry.slot);
        await graph(`/sites/${siteId}/drive/items/${entry.id}`, { method: "PATCH", body: JSON.stringify({ name: entry.name }) }).catch(() => {});
      }
    } catch (e) {
      showToast("Photo didn't upload. Check your signal and snap again.");
      todaysPhotos = todaysPhotos.filter(p => p !== entry); renderPhotos();
      return;
    }
    analyse(entry, blob);
  }
  async function reslot(newSlot) {
    const entry = lastUpload || todaysPhotos[todaysPhotos.length - 1];
    if (!entry) return;
    const oldField = SLOT_FIELD[entry.slot], newField = SLOT_FIELD[newSlot];
    const patch = {};
    if (oldField && oldField !== newField) patch[oldField] = false;
    if (newField) patch[newField] = true;
    const oldName = entry.name;
    entry.slot = newSlot; entry.name = entry.name.replace(/(Breakfast|Lunch|Dinner|Snack)/, newSlot);
    const foods = (getDay(todayKey()).foods || []).map(f => f.file === oldName ? { ...f, file: entry.name, slot: newSlot } : f);
    patch.foods = foods;
    change(patch); renderPhotos();
    showToast(`Moved to ${newSlot}.`);
    if (entry.id) {
      try { const { siteId } = await ensureSetup(); await graph(`/sites/${siteId}/drive/items/${entry.id}`, { method: "PATCH", body: JSON.stringify({ name: entry.name }) }); } catch (e) {}
    }
  }
  function showToast(text, slot, ms) {
    $("toastText").textContent = text;
    const slots = $("toastSlots");
    slots.hidden = !slot;
    slots.querySelectorAll("button").forEach(b => b.classList.toggle("on", b.dataset.slot === slot));
    $("toast").hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => { $("toast").hidden = true; }, ms || (slot ? 6000 : 2500));
  }

  // ---------- State changes ----------
  function change(patch) {
    const k = todayKey();
    days[k] = Object.assign(getDay(k), patch);
    pending.add(k); save(); render(); scheduleFlush();
  }

  // ---------- Render ----------
  function streak() {
    let n = 0; const d = today();
    if (!complete(days[isoDate(d)])) d.setDate(d.getDate() - 1);
    while (complete(days[isoDate(d)])) { n++; d.setDate(d.getDate() - 1); }
    return n;
  }
  function weights() { return Object.keys(days).sort().filter(k => typeof days[k].weight === "number").map(k => [k, days[k].weight]); }

  function render() {
    const now = today(), k = todayKey(), t = getDay(k);
    $("date").textContent = now.toLocaleDateString("en-GB", { weekday: "long", day: "numeric", month: "long" });

    const wrap = $("meals"); wrap.textContent = "";
    [["b", "b"], ["l", LUNCH[now.getDay()]], ["d", DINNER[now.getDay()]]].forEach(([field, mk]) => {
      const m = MEALS[mk];
      const btn = document.createElement("button");
      btn.className = "meal"; btn.id = "meal-" + field;
      btn.setAttribute("aria-pressed", t[field] ? "true" : "false");
      btn.innerHTML = '<div class="txt"><div class="slot"></div><div class="name"></div><div class="what"></div></div><div class="tick"><svg viewBox="0 0 24 24"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg></div>';
      btn.querySelector(".slot").textContent = m.slot;
      btn.querySelector(".name").textContent = m.name;
      btn.querySelector(".what").textContent = m.what;
      btn.onclick = () => change({ [field]: !getDay(todayKey())[field] });
      wrap.appendChild(btn);
    });
    const left = ["b", "l", "d"].filter(f => !t[f]).length;
    $("headline").textContent = left === 0 ? "All done today" : left === 3 ? "Three ticks today" : `${left} to go`;
    $("doneMsg").hidden = left !== 0;

    document.querySelectorAll("#gut .chip").forEach(c => c.setAttribute("aria-pressed", t.gut === c.dataset.v ? "true" : "false"));
    $("move").setAttribute("aria-pressed", t.move ? "true" : "false");
    $("streak").textContent = streak();

    // Weight
    const ws = weights();
    const isMonday = now.getDay() === 1;
    $("nudge").hidden = !(isMonday && typeof t.weight !== "number");
    $("weightCard").classList.toggle("due", isMonday && typeof t.weight !== "number");
    if (ws.length) {
      const last = ws[ws.length - 1][1], first = ws[0][1], diff = last - first;
      $("wLast").textContent = `${last.toFixed(1)} kg`;
      $("wTrend").textContent = ws.length > 1 ? `${diff <= 0 ? "Down" : "Up"} ${Math.abs(diff).toFixed(1)} kg since ${new Date(ws[0][0]).toLocaleDateString("en-GB", { day: "numeric", month: "short" })}` : "First weigh-in logged";
    } else { $("wLast").textContent = "Add weight"; $("wTrend").textContent = "Tap to log"; }
    drawSpark(ws.slice(-12).map(w => w[1]));

    // 14-day strip
    const strip = $("strip"); strip.textContent = "";
    for (let i = 13; i >= 0; i--) {
      const d = today(); d.setDate(d.getDate() - i);
      const r = days[isoDate(d)];
      const n = r ? ["b", "l", "d"].filter(f => r[f]).length : 0;
      const el = document.createElement("div"); el.className = "day";
      el.innerHTML = `<div class="dot ${n === 3 ? "full" : n ? "part" : ""} ${i === 0 ? "today" : ""}"></div><div class="gutmark ${r && r.gut ? r.gut : ""}"></div><div class="dl">${d.getDate()}</div>`;
      strip.appendChild(el);
    }
    renderPhotos();
  }
  function drawSpark(vals) {
    const svg = $("spark");
    if (vals.length < 2) { svg.innerHTML = ""; return; }
    const min = Math.min(...vals) - 0.3, max = Math.max(...vals) + 0.3;
    const x = i => 2 + i * (116 / (vals.length - 1));
    const y = v => 4 + (1 - (v - min) / (max - min)) * 32;
    const pts = vals.map((v, i) => `${x(i).toFixed(1)},${y(v).toFixed(1)}`);
    svg.innerHTML = `<path class="area" d="M${pts[0]} L${pts.join(" L")} L${x(vals.length - 1)},40 L2,40Z"/><path class="line" d="M${pts.join(" L")}"/><circle cx="${x(vals.length - 1)}" cy="${y(vals[vals.length - 1])}" r="3"/>`;
  }
  function renderPhotos() {
    const box = $("photos");
    box.querySelectorAll(".photo").forEach(n => n.remove());
    $("noPhotos").hidden = todaysPhotos.length > 0;
    $("photoCount").textContent = todaysPhotos.length ? `${todaysPhotos.length} today` : "";
    todaysPhotos.forEach(p => {
      const el = document.createElement("div"); el.className = "photo";
      const img = document.createElement("img"); img.alt = ""; img.loading = "lazy"; if (p.thumb) img.src = p.thumb;
      const s = document.createElement("span");
      const m = p.name.match(/(Breakfast|Lunch|Dinner|Snack)/); s.textContent = (m ? m[1] : p.slot || "");
      const est = document.createElement("em");
      const food = foodFor(p.name);
      if (food) {
        est.textContent = `~${food.calories} kcal · ${food.protein_g}g P`;
        if (food.flags && food.flags.length) {
          const fl = document.createElement("b"); fl.className = "flags"; fl.textContent = food.flags.join(", "); el.append(img, s, est, fl);
        } else el.append(img, s, est);
      } else {
        est.textContent = p.analysing ? "Spotting food…" : p.aiError ? "No estimate" : "";
        el.append(img, s, est);
      }
      box.appendChild(el);
    });
    const t = totals(getDay(todayKey()));
    $("dayTotal").hidden = !t.kcal;
    $("dayTotal").textContent = `So far today: ~${t.kcal.toLocaleString("en-GB")} kcal · ${t.protein}g protein`;
  }

  // ---------- Weight sheet ----------
  let wVal = 84.0;
  function openSheet() {
    const ws = weights(); const t = getDay(todayKey());
    wVal = typeof t.weight === "number" ? t.weight : ws.length ? ws[ws.length - 1][1] : 84.0;
    $("wValue").textContent = wVal.toFixed(1);
    $("scrim").hidden = false; $("sheet").hidden = false;
  }
  function closeSheet() { $("scrim").hidden = true; $("sheet").hidden = true; }

  // ---------- Wire up ----------
  function wire() {
    document.querySelectorAll("#gut .chip").forEach(c => c.onclick = () => change({ gut: getDay(todayKey()).gut === c.dataset.v ? null : c.dataset.v }));
    $("move").onclick = () => change({ move: !getDay(todayKey()).move });
    $("weightCard").onclick = openSheet;
    $("scrim").onclick = closeSheet; $("wCancel").onclick = closeSheet;
    document.querySelectorAll(".steppers button").forEach(b => b.onclick = () => {
      wVal = Math.max(30, Math.min(250, Math.round((wVal + parseFloat(b.dataset.step)) * 10) / 10));
      $("wValue").textContent = wVal.toFixed(1);
    });
    $("wSave").onclick = () => { change({ weight: wVal }); closeSheet(); showToast(`Saved ${wVal.toFixed(1)} kg.`); };
    $("camera").addEventListener("change", e => { const f = e.target.files && e.target.files[0]; e.target.value = ""; onPhoto(f); });
    $("toastSlots").querySelectorAll("button").forEach(b => b.onclick = () => reslot(b.dataset.slot));
    window.addEventListener("online", () => { if (pending.size) flush(); });
    document.addEventListener("visibilitychange", () => {
      if (document.visibilityState === "visible") { render(); refresh(); }
    });
  }

  let lastRefresh = 0;
  async function refresh() {
    if (Date.now() - lastRefresh < 30000) return;
    lastRefresh = Date.now();
    try { if (pending.size) await flush(); await pull(); render(); loadPhotos(); setStatus("Synced to SharePoint."); }
    catch (e) { setStatus(navigator.onLine ? "Couldn't reach SharePoint. Saved on this phone." : "Offline. Saved on this phone."); }
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
    if (!acct) {
      $("gate").hidden = false;
      $("signIn").onclick = signIn;
      return;
    }
    $("app").hidden = false; $("fab").hidden = false;
    wire(); render(); renderPhotos();
    setStatus("Syncing…");
    refresh();
  }
  start();
})();
