/* Fuel — single-page app. Sign-in via MSAL (Entra ID), data via Microsoft Graph into a SharePoint site. */
(() => {
  "use strict";
  const CFG = window.FUEL_CONFIG || {};
  const SCOPES = ["User.Read", "Sites.Manage.All"];
  const LIST_NAME = "Fuel";
  const OLD_LIST_NAMES = ["Fuel Log"];
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
  let activeKey = null;                       // null = today; "YYYY-MM-DD" when backdating
  const dayKey = () => activeKey || todayKey();
  const isToday = () => dayKey() === todayKey();
  const nowTime = (d = new Date()) => `${pad(d.getHours())}:${pad(d.getMinutes())}`;
  const fmt = n => Math.round(n).toLocaleString("en-GB");
  const blank = () => ({ b: false, l: false, d: false, gut: null, mood: null, energy: null, sleep: null, move: false, weight: null, waist: null, foods: [] });
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
    if (ids && ids.siteId && ids.listId && ids.v === 5) return ids;
    const u = new URL(CFG.siteUrl);
    const site = await graph(`/sites/${u.hostname}:${u.pathname.replace(/\/$/, "")}`);
    const lists = await graph(`/sites/${site.id}/lists?$select=id,displayName&$top=200`);
    let list = (lists.value || []).find(l => l.displayName === LIST_NAME) || (lists.value || []).find(l => OLD_LIST_NAMES.includes(l.displayName));
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
      { name: "WaistCm", number: { decimalPlaces: "one" } },
      { name: "Mood", number: { decimalPlaces: "none" } },
      { name: "Energy", text: {} },
      { name: "SleepQ", text: {} }
    ];
    for (const c of wanted) if (!have.has(c.name)) await graph(`/sites/${site.id}/lists/${list.id}/columns`, { method: "POST", body: JSON.stringify(c) });
    // Push reminders: device subscriptions + app config (the public VAPID key)
    const ensureList = async (name, columns) => {
      let l = (lists.value || []).find(x => x.displayName === name);
      if (!l) l = await graph(`/sites/${site.id}/lists`, { method: "POST", body: JSON.stringify({ displayName: name, list: { template: "genericList" }, columns }) });
      return l.id;
    };
    const pushListId = await ensureList("Fuel Push", [{ name: "Subscription", text: { allowMultipleLines: true } }, { name: "Enabled", boolean: {} }]);
    const configListId = await ensureList("Fuel Config", [{ name: "Value", text: { allowMultipleLines: true } }]);
    ids = { siteId: site.id, listId: list.id, pushListId, configListId, v: 5 };
    store.set("fuel.ids", ids);
    return ids;
  }

  async function pull() {
    const { siteId, listId } = await ensureSetup();
    let url = `/sites/${siteId}/lists/${listId}/items?$expand=fields($select=Title,Breakfast,Lunch,Dinner,Gut,Moved,WeightKg,WaistCm,FoodsJson,Mood,Energy,SleepQ)&$top=500`;
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
          mood: typeof f.Mood === "number" ? f.Mood : null, energy: f.Energy || null, sleep: f.SleepQ || null,
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
            FoodsJson: JSON.stringify(r.foods || []), Calories: t.kcal, ProteinG: t.protein,
            Energy: r.energy || "", SleepQ: r.sleep || "" };
          if (typeof r.mood === "number") fields.Mood = r.mood;
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
  function change(patch, key = dayKey()) {
    days[key] = Object.assign(getDay(key), patch);
    pending.add(key); save(); renderAll(); scheduleFlush();
  }
  const foodFor = name => (getDay(dayKey()).foods || []).find(f => f.file === name);
  const foodForField = field => (getDay(dayKey()).foods || []).find(f => f.key === field);
  const keyFor = (slot, time) => SLOT_FIELD[slot] || `snack-${(time || "").replace(":", "")}-${Date.now() % 100000}`;
  function upsertFood(rec, extraPatch = {}) {
    const foods = (getDay(dayKey()).foods || []).filter(f => f.key !== rec.key && !(rec.file && f.file === rec.file));
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
      headers: { "Content-Type": "application/json", "X-Fuel-Token": await token() },
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
  function slotFromText(text) {
    const t = text.toLowerCase();
    if (/\bbreakfast\b/.test(t)) return "Breakfast";
    if (/\blunch\b/.test(t)) return "Lunch";
    if (/\b(dinner|tea|supper)\b/.test(t)) return "Dinner";
    if (/\bsnack\b/.test(t)) return "Snack";
    return null;
  }
  async function logText(field, text, opts = {}) {
    // From a meal's own sheet: that meal. From "Type food": what the words say, else the time of day.
    const slot = field ? FIELD_SLOT[field] : (opts.slot || slotFromText(text) || slotForNow());
    const f = SLOT_FIELD[slot];
    const time = opts.time || nowTime();
    const key = opts.key || f || keyFor("Snack", time);
    upsertFood({ key, slot, time, source: "text", name: text, calories: 0, protein_g: 0, flags: [], items: [], pendingAI: true });
    try {
      const data = await callAI({ text, slot });
      upsertFood({ key, slot, time, source: "text", name: text, base: { kcal: data.calories, protein: data.protein_g }, portion: 1, ...data });
      if (!opts.quiet) announce(text, data);
    } catch (e) {
      upsertFood({ key, slot, time, source: "text", name: text, calories: 0, protein_g: 0, flags: [], items: [], aiError: e.message });
      if (!opts.quiet) showToast(`Saved "${text}", but couldn't estimate it: ${e.message}`, null, 6000);
    }
  }
  // Re-estimate anything saved without calories (e.g. while offline or before a fix), and move
  // typed items that clearly name a meal into that meal.
  let repairing = false;
  async function repairFoods() {
    if (repairing) return; repairing = true;
    try {
      const t = getDay(dayKey());
      for (const f of (t.foods || []).filter(x => x.source === "text" && !x.calories && !x.pendingAI)) {
        const named = slotFromText(f.name);
        if (named && named !== f.slot) {
          const foods = (getDay(dayKey()).foods || []).filter(x => x.key !== f.key);
          change({ foods });
          await logText(SLOT_FIELD[named] || null, f.name, { slot: named, time: f.time, quiet: true });
        } else {
          await logText(SLOT_FIELD[f.slot] || null, f.name, { slot: f.slot, time: f.time, key: f.key, quiet: true });
        }
      }
    } finally { repairing = false; }
  }
  function logPlan(field, mk) {
    const m = MEALS[mk];
    upsertFood({ key: field, slot: m.slot, time: nowTime(), source: "plan", name: m.name, base: { kcal: m.kcal, protein: m.protein }, portion: 1,
      calories: m.kcal, protein_g: m.protein, flags: [], items: [], note: "" });
  }
  function setPortion(key, p) {
    const foods = (getDay(dayKey()).foods || []).map(f => {
      if (f.key !== key) return f;
      const base = f.base || { kcal: f.calories, protein: f.protein_g };
      return { ...f, base, portion: p, calories: Math.round(base.kcal * p), protein_g: Math.round(base.protein * p) };
    });
    change({ foods });
  }
  const foodsInSlot = (slot, r = getDay(dayKey())) => (r.foods || []).filter(f => (f.slot || "Snack") === slot);
  function addFood(rec) {   // adds alongside anything already logged for that meal
    const field = SLOT_FIELD[rec.slot];
    const existing = foodsInSlot(rec.slot);
    rec.key = field ? (existing.length ? `${field}+${Date.now() % 1000000}` : field) : keyFor("Snack", rec.time);
    const foods = (getDay(dayKey()).foods || []).concat(rec);
    change({ ...(field ? { [field]: true } : {}), foods });
  }
  function removeFood(key) {
    const all = getDay(dayKey()).foods || [];
    const gone = all.find(f => f.key === key);
    const foods = all.filter(f => f.key !== key);
    const patch = { foods };
    const field = gone && SLOT_FIELD[gone.slot];
    if (field && !foods.some(f => f.slot === gone.slot)) patch[field] = false;
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
      const res = await graph(`/sites/${siteId}/drive/root:/${encPath(PHOTO_ROOT + "/" + dayKey())}:/children?$expand=thumbnails&$select=id,name`);
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
    const taken = file.lastModified && isoDate(new Date(file.lastModified)) === dayKey() ? new Date(file.lastModified) : now;
    const slot = slotOverride || slotForNow(taken);
    const time = nowTime(taken);
    const entry = { id: null, name: `${time.replace(":", "")}-${slot}.jpg`, thumb: URL.createObjectURL(file), local: true, slot, time };
    entry.key = keyFor(slot, time);
    todaysPhotos.push(entry); lastUpload = entry;
    const field = SLOT_FIELD[slot];
    if (field && !getDay(dayKey())[field]) change({ [field]: true }); else renderAll();
    showToast(slotOverride ? `Added to ${slot.toLowerCase()}. Working out calories…` : `Logged as ${slot.toLowerCase()}. Wrong? Tap to move it.`, slotOverride ? null : slot);
    let blob;
    try { blob = await shrink(file); } catch (e) { showToast(e.message); return; }
    try {
      const { siteId } = await ensureSetup();
      const res = await graph(`/sites/${siteId}/drive/root:/${encPath(PHOTO_ROOT + "/" + dayKey() + "/" + entry.name)}:/content?@microsoft.graph.conflictBehavior=rename`, {
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
    let foods = getDay(dayKey()).foods || [];
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
      await graph(`/sites/${siteId}/drive/root:/${encPath(PROGRESS_ROOT + "/" + dayKey() + "-" + pose + ".jpg")}:/content`, { method: "PUT", body: blob, headers: { "Content-Type": "image/jpeg" } });
      showToast(`${pose[0].toUpperCase() + pose.slice(1)} photo saved.`);
      await loadProgress();
    } catch (e) { showToast("Photo didn't save. Try again."); }
  }

  // ---------- Barcode scanning ----------
  let reader = null, scanning = false, scanSlot = null, product = null, prodGrams = 100, prodSlot = "Snack";
  async function openScanner(slot) {
    scanSlot = slot || null;
    closeSheets();
    $("scanner").hidden = false; $("scanCode").value = "";
    $("scanMsg").textContent = "Point at the barcode";
    for (let i = 0; i < 30 && !window.ZXing; i++) await new Promise(r => setTimeout(r, 100));
    if (!window.ZXing) { $("scanMsg").textContent = "Scanner didn't load. Type the numbers instead."; return; }
    try {
      const hints = new Map();
      const F = ZXing.BarcodeFormat;
      hints.set(ZXing.DecodeHintType.POSSIBLE_FORMATS, [F.EAN_13, F.EAN_8, F.UPC_A, F.UPC_E]);
      reader = new ZXing.BrowserMultiFormatReader(hints);
      scanning = true;
      await reader.decodeFromConstraints({ video: { facingMode: { ideal: "environment" } } }, $("scanVideo"), res => {
        if (res && scanning) { scanning = false; if (navigator.vibrate) navigator.vibrate(40); stopScanner(); lookup(res.getText()); }
      });
    } catch (e) {
      $("scanMsg").textContent = "Couldn't open the camera. Allow camera access in Settings, or type the numbers below.";
    }
  }
  function stopScanner() {
    scanning = false;
    try { if (reader) reader.reset(); } catch (e) {}
    reader = null; $("scanner").hidden = true;
  }
  function flagsFor(p) {
    const f = new Set();
    const al = (p.allergens_tags || []).join(" ") + " " + (p.traces_tags || []).join(" ");
    const ing = String(p.ingredients_text_en || p.ingredients_text || "").toLowerCase();
    if (/en:milk/.test(al) || /\b(milk|whey|lactose|cream|butter|cheese|yogurt|yoghurt)\b/.test(ing)) f.add("dairy");
    if (/\bonion/.test(ing)) f.add("onion");
    if (/\bgarlic/.test(ing)) f.add("garlic");
    if (/\b(beans?|lentils?|chickpeas?)\b/.test(ing)) f.add("beans");
    if (/(sucralose|aspartame|acesulfame|sorbitol|maltitol|xylitol|erythritol|stevia|sweetener)/.test(ing)) f.add("sweetener");
    return [...f];
  }
  async function lookup(code) {
    code = String(code).replace(/\D/g, "");
    if (code.length < 8) { showToast("That barcode looks too short. Try again."); return; }
    showToast("Looking it up…");
    try {
      const res = await fetch(`https://world.openfoodfacts.org/api/v2/product/${code}.json?fields=product_name,product_name_en,brands,nutriments,serving_size,serving_quantity,allergens_tags,traces_tags,ingredients_text,ingredients_text_en,image_front_small_url,quantity`);
      const data = await res.json();
      const p = data && data.status === 1 && data.product;
      const n = p && p.nutriments || {};
      const kcal100 = n["energy-kcal_100g"] ?? (n.energy_100g ? n.energy_100g / 4.184 : null);
      if (!p || kcal100 == null) {
        showToast("Not in the food database yet. Snap a photo of it or type it instead.", null, 5000);
        openMealSheet(scanSlot ? SLOT_FIELD[scanSlot] : null, null);
        return;
      }
      product = {
        code, name: p.product_name_en || p.product_name || "Unnamed product", brand: (p.brands || "").split(",")[0].trim(),
        img: p.image_front_small_url || "", kcal100: Number(kcal100), pro100: Number(n.proteins_100g || 0),
        servingQ: Number(p.serving_quantity) || null, servingLabel: p.serving_size || "", pack: p.quantity || "", flags: flagsFor(p)
      };
      prodGrams = product.servingQ || 100;
      prodSlot = scanSlot || slotForNow();
      $("toast").hidden = true;
      openProduct();
    } catch (e) { showToast("Couldn't reach the food database. Check your signal."); }
  }
  function prodNumbers(g) { return { kcal: Math.round(product.kcal100 * g / 100), protein: Math.round(product.pro100 * g / 10) / 10 }; }
  function openProduct() {
    $("prodName").textContent = product.name;
    $("prodBrand").textContent = [product.brand, product.pack].filter(Boolean).join(" · ");
    $("prodImg").src = product.img || "";
    $("prodFlags").hidden = !product.flags.length;
    $("prodFlags").textContent = product.flags.length ? `Heads up: contains ${product.flags.join(", ")}` : "";
    const box = $("prodAmounts"); box.textContent = "";
    const amounts = [];
    if (product.servingQ) amounts.push([product.servingQ, `1 serving${product.servingLabel ? " (" + product.servingLabel + ")" : ""}`], [product.servingQ * 2, "2 servings"]);
    amounts.push([100, "100 g"]);
    amounts.forEach(([g, label]) => {
      const b = document.createElement("button"); b.className = "opt" + (Math.abs(g - prodGrams) < .01 ? " on" : "");
      const v = prodNumbers(g);
      b.innerHTML = "<span><b></b><small></small></span>";
      b.querySelector("b").textContent = label;
      b.querySelector("small").textContent = `${fmt(v.kcal)} kcal · ${v.protein}g protein`;
      b.onclick = () => { prodGrams = g; openProduct(); };
      box.appendChild(b);
    });
    $("prodGrams").value = Math.round(prodGrams);
    const v = prodNumbers(prodGrams);
    $("prodTotal").innerHTML = `${fmt(v.kcal)} kcal<small></small>`;
    $("prodTotal").querySelector("small").textContent = `${v.protein}g protein · ${Math.round(prodGrams)} g`;
    document.querySelectorAll("#prodSlot button").forEach(b => b.classList.toggle("on", b.dataset.slot === prodSlot));
    openSheet("productSheet");
  }
  function logProduct() {
    const v = prodNumbers(prodGrams);
    addFood({ slot: prodSlot, time: nowTime(), source: "barcode", code: product.code, grams: Math.round(prodGrams),
      name: product.name, calories: v.kcal, protein_g: Math.round(v.protein), flags: product.flags, items: [], note: "" });
    closeSheets();
    showToast(`${product.name}: ${fmt(v.kcal)} kcal added to ${prodSlot.toLowerCase()}.`);
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
    const now = today(), t = getDay(dayKey()), tot = totals(t);
    const hour = now.getHours();
    const ad = new Date(dayKey() + "T12:00:00");
    $("date").textContent = isToday() ? "Today, " + ad.toLocaleDateString("en-GB", { day: "numeric", month: "long" }) : ad.toLocaleDateString("en-GB", { weekday: "long", day: "numeric", month: "long" });
    document.querySelector(".daypick").classList.toggle("past", !isToday());
    $("dayNext").disabled = isToday();
    $("headline").textContent = "Food";

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
    const rows = [["b", "b"], ["l", LUNCH[ad.getDay()]], ["d", DINNER[ad.getDay()]]];
    rows.forEach(([field, mk]) => {
      const m = MEALS[mk], recs = foodsInSlot(m.slot, t);
      const rec = recs.length ? { name: recs.map(x => x.name).join(", "), calories: recs.reduce((a, x) => a + (x.calories || 0), 0), protein_g: recs.reduce((a, x) => a + (x.protein_g || 0), 0) } : null;
      const busy = todaysPhotos.some(p => p.slot === m.slot && p.analysing) || recs.some(x => x.pendingAI);
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
    const snacks = foodsInSlot("Snack", t);
    const sk = snacks.reduce((a, f) => a + (f.calories || 0), 0);
    const sb = document.createElement("button");
    sb.className = "meal" + (snacks.length ? " done" : "");
    sb.innerHTML = `<span class="m-ic">${snacks.length ? ICONS.tick : ICONS.s}</span><span class="m-txt"><span class="m-slot">Snacks & extras</span><span class="m-name"></span></span><span class="m-kcal"></span>`;
    sb.querySelector(".m-name").textContent = snacks.length ? snacks.map(s => s.name).join(", ") : (isToday() ? "Anything else today" : "Anything else");
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

  // ---------- Moss, sunlight and the garden ----------
  const SUN_TARGET = 10;
  function sunFor(r) {
    if (!r) return 0;
    let s = 0;
    ["gut", "mood", "energy", "sleep"].forEach(k => { if (r[k] != null && r[k] !== "") s += 0.5; });
    ["b", "l", "d"].forEach(k => { if (r[k]) s += 2; });
    if (r.move) s += 2;
    if (typeof r.weight === "number" || typeof r.waist === "number") s += 1;
    if ((r.foods || []).some(f => f.slot === "Snack")) s += 0.5;
    return s;
  }
  const fullDays = () => Object.keys(days).filter(k => k < todayKey() && sunFor(days[k]) >= SUN_TARGET).length + (sunFor(getDay(todayKey())) >= SUN_TARGET ? 1 : 0);
  const GARDEN = [
    [1, "pebbles", "a little pile of pebbles"], [2, "mushroom", "a red mushroom"], [3, "daisy", "a daisy"], [5, "fern", "a curly fern"],
    [7, "lantern", "a tiny lantern"], [10, "butterfly", "a butterfly friend"], [14, "path", "a stepping-stone path"], [18, "tulips", "a row of tulips"],
    [21, "bench", "a little bench"], [25, "pond", "a pond"], [30, "tree", "a young tree"], [40, "bird", "a robin who visits"]
  ];
  const STORIES = [
    "Moss followed a beetle to the bottom of the garden and back. It took all night. Moss says it was worth it.",
    "Moss climbed the fence to look at the stars, counted to forty-two, then fell asleep in a flowerpot.",
    "Moss found a dewdrop the size of a pea and spent the night admiring it.",
    "Moss had tea with a snail. The snail was late, but very polite.",
    "Moss tried to catch the moon in a puddle. Nearly managed it.",
    "Moss went looking for the source of a mysterious hum. It was a bee. A very sleepy bee.",
    "Moss rearranged the pebbles three times and decided the first way was best.",
    "Moss practised a little dance under the lantern. No one saw. Probably.",
    "Moss followed a trail of crumbs and found a ladybird having a picnic.",
    "Moss listened to the rain on a leaf and wrote a very short song about it."
  ];
  const hash = str => { let h = 0; for (const c of str) h = (h * 31 + c.charCodeAt(0)) >>> 0; return h; };
  function mossMood() {
    const h = today().getHours();
    const t = getDay(todayKey());
    if (sunFor(t) >= SUN_TARGET) return "glow";
    if (h < 9 && !t.mood) return "sleepy";
    if (h >= 21) return "cosy";
    return "happy";
  }
  function mossSays() {
    const t = getDay(todayKey()), h = today().getHours(), sun = sunFor(t);
    const pickOne = arr => arr[hash(todayKey() + h) % arr.length];
    if (sun >= SUN_TARGET) return pickOne(["Full of sunshine. I'm going exploring tonight!", "Look at us. Proper glowing today.", "That's a full sun day. Garden's going to love this."]);
    if (t.gut === "rough") return "Rough gut today? Let's go gentle. Small meals, plenty of water.";
    if (h < 10 && t.mood == null) return pickOne(["Morning. How are we feeling?", "Sun's coming up. Quick check-in?", "Morning! Tap me a check-in and I'll wake up properly."]);
    if (h < 12) return pickOne(["One small thing at a time. That's the whole trick.", "Smoothie first, world second.", "You're doing fine. Honestly."]);
    if (h < 17) return pickOne(["Afternoon. Drink some water, I'll wait.", "Halfway there. What's one tiny thing next?", "Lunch logged? I'm nosy about lunch."]);
    if (h < 21) return pickOne(["Evening. Dinner, then feet up.", "Nearly done for the day. Proud of you.", "Twenty minutes of moving and I get a new plant. Just saying."]);
    return pickOne(["Cosy time. Phone down soon?", "Big day. Rest up.", "Night night. I'll tell you what I find."]);
  }
  function mossSvg(mood, items) {
    const has = id => items.includes(id);
    const eyes = mood === "sleepy" ? '<path d="M178 168q8 6 16 0M226 168q8 6 16 0" stroke="#1c2a20" stroke-width="3.5" fill="none" stroke-linecap="round"/>'
      : mood === "cosy" ? '<path d="M178 170q8-6 16 0M226 170q8-6 16 0" stroke="#1c2a20" stroke-width="3.5" fill="none" stroke-linecap="round"/>'
      : mood === "glow" ? '<path d="M178 170q8-9 16 0M226 170q8-9 16 0" stroke="#1c2a20" stroke-width="3.5" fill="none" stroke-linecap="round"/>'
      : '<ellipse cx="186" cy="167" rx="7" ry="8.5" fill="#1c2a20"/><ellipse cx="234" cy="167" rx="7" ry="8.5" fill="#1c2a20"/><circle cx="188.5" cy="164" r="2.4" fill="#fff"/><circle cx="236.5" cy="164" r="2.4" fill="#fff"/>';
    const mouth = mood === "sleepy" ? '<ellipse cx="210" cy="188" rx="4" ry="3" fill="#1c2a20"/>' : '<path d="M202 184q8 8 16 0" stroke="#1c2a20" stroke-width="3" fill="none" stroke-linecap="round"/>';
    const glow = mood === "glow" ? '<circle cx="210" cy="170" r="92" fill="var(--sun)" opacity=".18"/>' : "";
    const zzz = mood === "sleepy" ? '<text x="268" y="118" font-family="var(--display)" font-weight="700" font-size="18" fill="var(--muted)">z</text><text x="282" y="100" font-family="var(--display)" font-weight="700" font-size="13" fill="var(--muted)">z</text>' : "";
    const it = (id, svg) => has(id) ? `<g class="g-item">${svg}</g>` : "";
    return `<svg viewBox="0 0 420 262" role="img" aria-label="Moss, sitting in its pot${items.length ? " in a garden with " + items.length + " things" : ""}">
      <ellipse cx="210" cy="252" rx="200" ry="14" fill="var(--track)"/>
      ${it("pond", '<ellipse cx="350" cy="240" rx="52" ry="12" fill="#8cc4d8" opacity=".85"/><ellipse cx="338" cy="237" rx="14" ry="3" fill="#fff" opacity=".6"/>')}
      ${it("tree", '<rect x="34" y="120" width="12" height="122" rx="5" fill="var(--soil)"/><circle cx="40" cy="104" r="34" fill="var(--moss-dark)"/><circle cx="62" cy="124" r="22" fill="var(--moss)"/><circle cx="20" cy="124" r="20" fill="var(--moss)"/>')}
      ${it("bird", '<g transform="translate(46 70)"><ellipse cx="0" cy="0" rx="11" ry="9" fill="#7a5a43"/><circle cx="8" cy="-6" r="6" fill="#7a5a43"/><ellipse cx="2" cy="3" rx="6" ry="5" fill="#d9663b"/><path d="M14 -6l5 1.5-5 1.5z" fill="var(--sun)"/><circle cx="9" cy="-7" r="1.3" fill="#1c2a20"/></g>')}
      ${it("path", '<ellipse cx="280" cy="250" rx="13" ry="5" fill="var(--line)"/><ellipse cx="306" cy="246" rx="11" ry="4" fill="var(--line)"/><ellipse cx="254" cy="252" rx="11" ry="4" fill="var(--line)"/>')}
      ${it("bench", '<rect x="300" y="196" width="70" height="8" rx="3" fill="var(--soil)"/><rect x="300" y="180" width="70" height="7" rx="3" fill="var(--soil)"/><rect x="306" y="204" width="5" height="36" fill="var(--soil)"/><rect x="359" y="204" width="5" height="36" fill="var(--soil)"/>')}
      ${it("tulips", '<g fill="var(--moss-dark)"><rect x="96" y="214" width="3" height="30"/><rect x="110" y="208" width="3" height="36"/><rect x="124" y="216" width="3" height="28"/></g><path d="M91 214q6-14 12 0q-6 6-12 0z" fill="#e0607e"/><path d="M105 208q6-14 12 0q-6 6-12 0z" fill="var(--sun)"/><path d="M119 216q6-14 12 0q-6 6-12 0z" fill="#e0607e"/>')}
      ${it("fern", '<path d="M330 244q-2-40 22-58M330 244q10-30 38-34M330 244q-16-28-40-30" stroke="var(--moss)" stroke-width="5" fill="none" stroke-linecap="round"/>')}
      ${it("lantern", '<rect x="96" y="150" width="3" height="94" fill="var(--soil)"/><rect x="88" y="138" width="19" height="22" rx="4" fill="var(--sun)" opacity=".9"/><circle cx="97" cy="149" r="16" fill="var(--sun)" opacity=".18"/>')}
      ${it("mushroom", '<rect x="148" y="226" width="9" height="18" rx="4" fill="#f3e9da"/><path d="M137 230q15-24 31 0z" fill="#d9483b"/><circle cx="146" cy="223" r="2.4" fill="#fff"/><circle cx="156" cy="221" r="2" fill="#fff"/>')}
      ${it("daisy", '<rect x="282" y="214" width="3" height="30" fill="var(--moss-dark)"/><g fill="#fff"><circle cx="283" cy="206" r="5"/><circle cx="291" cy="212" r="5"/><circle cx="275" cy="212" r="5"/><circle cx="279" cy="220" r="5"/><circle cx="287" cy="220" r="5"/></g><circle cx="283" cy="213" r="4" fill="var(--sun)"/>')}
      ${it("pebbles", '<ellipse cx="128" cy="246" rx="9" ry="5" fill="#a9b0a8"/><ellipse cx="141" cy="248" rx="6" ry="4" fill="#c2c7c0"/>')}
      ${it("butterfly", '<g transform="translate(330 120)"><path d="M0 0q-16-16-18 2q2 12 18-2z" fill="#e0607e"/><path d="M0 0q16-16 18 2q-2 12-18-2z" fill="#f08aa2"/><rect x="-1.5" y="-6" width="3" height="14" rx="1.5" fill="#1c2a20"/></g>')}
      ${glow}
      <g class="moss-body">
        <path d="M150 196h120l-12 52a8 8 0 0 1-8 6h-80a8 8 0 0 1-8-6z" fill="var(--clay)"/>
        <rect x="142" y="188" width="136" height="18" rx="6" fill="var(--clay-dark)"/>
        <path d="M210 92c46 0 66 34 66 66 0 22-12 34-24 38h-84c-12-4-24-16-24-38 0-32 20-66 66-66z" fill="var(--moss)"/>
        <circle cx="174" cy="126" r="6" fill="var(--moss-dark)" opacity=".6"/><circle cx="250" cy="138" r="5" fill="var(--moss-dark)" opacity=".6"/><circle cx="232" cy="112" r="4" fill="var(--moss-dark)" opacity=".5"/>
        <path d="M210 94c-2-14 2-26 12-32" stroke="var(--moss-dark)" stroke-width="4" fill="none" stroke-linecap="round"/>
        <path d="M220 64c10-8 24-6 30 2c-10 8-24 8-30-2z" fill="var(--moss-dark)"/>
        <path d="M212 76c-10-10-24-8-28 0c8 8 22 8 28 0z" fill="var(--moss)"/>
        ${eyes}
        <ellipse cx="170" cy="182" rx="9" ry="5" fill="#e98b8b" opacity=".45"/><ellipse cx="250" cy="182" rx="9" ry="5" fill="#e98b8b" opacity=".45"/>
        ${mouth}
      </g>
      ${zzz}
    </svg>`;
  }
  function upNext() {
    const t = getDay(todayKey()), now = today(), h = now.getHours() + now.getMinutes() / 60;
    const din = DINNER[now.getDay()], lun = LUNCH[now.getDay()];
    if (t.mood == null && h < 12) return { title: "Morning check-in", text: "Gut, mood, energy, sleep. Four taps, all optional.", btn: "Check in", go: openCheckin };
    if (now.getDay() === 1 && t.weight == null && h < 12) return { title: "Weigh-in day", text: "Scales first thing, after the loo. Waist too if you've got the tape.", btn: "Log weight", go: () => openNum("weight") };
    if (!t.b && h < 11.5) return { title: "Breakfast", text: `${MEALS.b.name}, or log whatever you had.`, btn: "Log breakfast", go: () => openMealSheet("b", "b") };
    if (!t.l && h >= 11.5 && h < 16) return { title: "Lunch", text: `Plan says ${MEALS[lun].name.toLowerCase()}. Snap it, scan it or tap it.`, btn: "Log lunch", go: () => openMealSheet("l", lun) };
    if (!t.move && h >= 16.5 && h < 20.5) return { title: "20-minute workout", text: "Press-ups, squats, lunges, band rows, plank. Moss gets 2 sunlight.", btn: "I've done it", go: () => openSheet("moveSheet") };
    if (!t.d && h >= 17.5 && h < 22) return { title: "Dinner", text: `Plan says ${din === "dAny" ? "your pick" : MEALS[din].name.toLowerCase()}.`, btn: "Log dinner", go: () => openMealSheet("d", din) };
    if (t.mood == null) return { title: "Quick check-in", text: "Missed it this morning? It still counts.", btn: "Check in", go: openCheckin };
    if (h >= 20.5) return { title: "That's the day", text: "Have a look at your wins, then phone down.", btn: "See trends", go: () => show("trends") };
    return { title: "All caught up", text: "Nothing needs doing right now. Moss is sunbathing.", btn: "Log something", go: () => openSheet("logSheet") };
  }
  function winsList() {
    const t = getDay(todayKey()), out = [];
    if (t.mood != null || t.gut) out.push(["Morning check-in", 2]);
    ["b", "l", "d"].forEach(f => { if (t[f]) { const recs = foodsInSlot(FIELD_SLOT[f], t); out.push([`${FIELD_SLOT[f]}: ${recs.length ? recs.map(r => r.name).join(", ") : "logged"}`, 2]); } });
    foodsInSlot("Snack", t).forEach(r => out.push([`Snack: ${r.name}`, 0]));
    if (t.move) out.push(["20-minute workout", 2]);
    if (typeof t.weight === "number") out.push([`Weighed in: ${t.weight.toFixed(1)} kg`, 1]);
    if (typeof t.waist === "number") out.push([`Waist: ${t.waist.toFixed(1)} cm`, 0]);
    return out;
  }
  function renderHome() {
    const now = today(), t = getDay(todayKey()), h = now.getHours();
    $("homeDate").textContent = now.toLocaleDateString("en-GB", { weekday: "long", day: "numeric", month: "long" });
    $("homeHello").textContent = h < 12 ? "Morning, George" : h < 18 ? "Afternoon, George" : "Evening, George";
    const sun = sunFor(t);
    $("sunCount").textContent = Number.isInteger(sun) ? sun : sun.toFixed(1);
    $("sunFill").style.width = Math.min(100, sun / SUN_TARGET * 100) + "%";
    $("sunNote").textContent = sun >= SUN_TARGET ? "Full sun! Moss goes exploring tonight." : `${(SUN_TARGET - sun).toFixed(sun % 1 ? 1 : 0)} more sunlight and Moss goes exploring tonight.`;
    const full = fullDays();
    const unlocked = GARDEN.filter(g => full >= g[0]).map(g => g[1]);
    $("garden").innerHTML = mossSvg(mossMood(), unlocked);
    $("mossSays").textContent = mossSays();
    const nextG = GARDEN.find(g => full < g[0]);
    $("gardenCount").textContent = `${unlocked.length} of ${GARDEN.length} found`;
    $("gardenNext").textContent = nextG ? `${nextG[0] - full} more full-sun day${nextG[0] - full === 1 ? "" : "s"} until Moss finds ${nextG[2]}. Missed days never take anything away.` : "Moss has found everything. The garden's complete.";
    // overnight story (morning only)
    const y = new Date(now); y.setDate(y.getDate() - 1);
    const yk = isoDate(y), ysun = sunFor(days[yk]);
    if (h < 14 && days[yk]) {
      $("storyCard").hidden = false;
      if (ysun >= SUN_TARGET) {
        const found = GARDEN.find(g => g[0] === Object.keys(days).filter(k => k <= yk && sunFor(days[k]) >= SUN_TARGET).length);
        $("storyText").textContent = STORIES[hash(yk) % STORIES.length] + (found ? ` And look: Moss brought back ${found[2]}!` : "");
      } else {
        $("storyText").textContent = "Moss stayed in last night, had a cup of tea and an early one. Fresh start today.";
      }
    } else $("storyCard").hidden = true;
    const u = upNext();
    $("upTitle").textContent = u.title; $("upText").textContent = u.text; $("upBtn").textContent = u.btn;
    $("upBtn").onclick = () => { if (!isToday()) setActive(null); u.go(); };
    const w = winsList(), ul = $("wins"); ul.textContent = "";
    $("winCount").textContent = w.length ? `${w.length} so far` : "";
    if (!w.length) { const li = document.createElement("li"); li.className = "empty"; li.textContent = "Nothing yet, and that's fine. One tiny thing gets this going."; ul.appendChild(li); }
    w.forEach(([txt, s]) => {
      const li = document.createElement("li");
      li.innerHTML = `<span class="w-ic">${ICONS.tick}</span><span class="w-t"></span>${s ? `<span class="w-sun">+${s} sun</span>` : ""}`;
      li.querySelector(".w-t").textContent = txt; ul.appendChild(li);
    });
  }

  // ---------- Check-in ----------
  function openCheckin() {
    const t = getDay(dayKey());
    document.querySelectorAll("#checkinSheet .ci-opts").forEach(g => {
      const f = g.dataset.field;
      g.querySelectorAll("button").forEach(b => b.classList.toggle("on", String(t[f]) === b.dataset.v));
    });
    openSheet("checkinSheet");
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

    // mood
    const msn = $("moodStrip"); msn.textContent = ""; msn.style.setProperty("--n", Math.min(range, 15));
    const moods = gutKeys.map(k => getDay(k).mood);
    gutKeys.forEach((k, i) => { const v = moods[i]; const c = document.createElement("div"); c.className = "ms" + (v ? " m" + v : ""); c.title = `${shortDate(k)}: ${v ? "mood " + v + "/5" : "not logged"}`; msn.appendChild(c); });
    const ml = moods.filter(v => v);
    $("moodKpi").textContent = ml.length ? `${(ml.reduce((a, b) => a + b, 0) / ml.length).toFixed(1)}/5 avg` : "–";
    $("moodInsight").textContent = moodInsight();
    // consistency
    const cs = $("consStrip"); cs.textContent = ""; cs.style.setProperty("--n", Math.min(range, 15));
    let full = 0;
    gutKeys.forEach(k => { const r = getDay(k); const n = ["b", "l", "d"].filter(f => r[f]).length; if (n === 3) full++; const c = document.createElement("div"); c.className = "cs" + (n ? " p" + n : ""); c.title = `${shortDate(k)}: ${n}/3 meals`; cs.appendChild(c); });
    $("consKpi").textContent = `${full}/${gutKeys.length} days`;
  }
  function moodInsight() {
    const ks = Object.keys(days).filter(k => days[k].mood);
    if (ks.length < 5) return "Check in most mornings and I'll show you what lifts your mood: sleep, workouts, eating well.";
    const avg = arr => arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : null;
    const prevMove = ks.map(k => { const p = new Date(k); p.setDate(p.getDate() - 1); return [days[k].mood, !!(days[isoDate(p)] || {}).move]; });
    const a = avg(prevMove.filter(x => x[1]).map(x => x[0])), b = avg(prevMove.filter(x => !x[1]).map(x => x[0]));
    const sl = avg(ks.filter(k => days[k].sleep === "great").map(k => days[k].mood)), sp = avg(ks.filter(k => days[k].sleep === "poor").map(k => days[k].mood));
    if (sl != null && sp != null && sl - sp >= 0.8) return `Mood averages ${sl.toFixed(1)} after great sleep and ${sp.toFixed(1)} after poor sleep. Bedtime is doing a lot of work.`;
    if (a != null && b != null && a - b >= 0.5) return `Mood is ${(a - b).toFixed(1)} points higher the day after a workout. Worth protecting those 20 minutes.`;
    return "No strong pattern yet. Keep checking in and it'll show up.";
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
    const since = last ? Math.floor((new Date(dayKey()) - new Date(last)) / 86400000) : null;
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

  let current = "home";
  function renderAll() {
    renderToday();
    if (current === "home") renderHome();
    if (current === "trends") renderTrends();
    if (current === "body") renderBody();
  }
  function show(view) {
    current = view;
    ["home", "today", "trends", "body"].forEach(v => { $("view-" + v).hidden = v !== view; });
    document.querySelectorAll(".tab[data-view]").forEach(t => { if (t.dataset.view === view) t.setAttribute("aria-current", "page"); else t.removeAttribute("aria-current"); });
    renderAll();
    window.scrollTo(0, 0);
  }

  // ---------- Push reminders ----------
  const b64url = buf => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  const fromB64url = s => { const p = s.replace(/-/g, "+").replace(/_/g, "/"); const bin = atob(p + "===".slice((p.length + 3) % 4)); return Uint8Array.from(bin, c => c.charCodeAt(0)); };
  async function getConfig(key) {
    const { siteId, configListId } = await ensureSetup();
    const res = await graph(`/sites/${siteId}/lists/${configListId}/items?$expand=fields($select=Title,Value)&$top=50`);
    const it = (res.value || []).find(i => i.fields && i.fields.Title === key);
    return it ? it.fields.Value : null;
  }
  async function setConfig(key, value) {
    const { siteId, configListId } = await ensureSetup();
    await graph(`/sites/${siteId}/lists/${configListId}/items`, { method: "POST", body: JSON.stringify({ fields: { Title: key, Value: value } }) });
  }
  function pushSupported() { return "serviceWorker" in navigator && "PushManager" in window && "Notification" in window; }
  async function pushStatus() {
    if (!pushSupported()) return "unsupported";
    const reg = await navigator.serviceWorker.ready;
    const sub = await reg.pushManager.getSubscription();
    return sub ? "on" : Notification.permission === "denied" ? "denied" : "off";
  }
  async function renderPushStatus() {
    const st = await pushStatus().catch(() => "unsupported");
    const standalone = window.matchMedia("(display-mode: standalone)").matches || navigator.standalone;
    $("pushBtn").hidden = st === "on" || st === "unsupported";
    $("pushOff").hidden = st !== "on";
    $("pushState").textContent = st === "on" ? "Reminders are on for this phone." : st === "denied" ? "Notifications are blocked. Turn them on in iPhone Settings → Fuel." : st === "unsupported" ? (standalone ? "This device can't receive web notifications." : "Open Fuel from your home screen icon to turn on reminders.") : "Reminders are off on this phone.";
  }
  async function enablePush() {
    try {
      const perm = await Notification.requestPermission();
      if (perm !== "granted") { renderPushStatus(); return; }
      let pub = await getConfig("vapidPublicKey");
      let newPrivate = null;
      if (!pub) {
        const kp = await crypto.subtle.generateKey({ name: "ECDSA", namedCurve: "P-256" }, true, ["sign", "verify"]);
        pub = b64url(await crypto.subtle.exportKey("raw", kp.publicKey));
        newPrivate = (await crypto.subtle.exportKey("jwk", kp.privateKey)).d;
        await setConfig("vapidPublicKey", pub);
      }
      const reg = await navigator.serviceWorker.ready;
      const sub = await reg.pushManager.subscribe({ userVisibleOnly: true, applicationServerKey: fromB64url(pub) });
      const { siteId, pushListId } = await ensureSetup();
      await graph(`/sites/${siteId}/lists/${pushListId}/items`, { method: "POST", body: JSON.stringify({ fields: { Title: (navigator.userAgent.match(/iPhone|iPad|Android|Mac|Windows/) || ["Device"])[0] + " " + todayKey(), Subscription: JSON.stringify(sub), Enabled: true } }) });
      if (newPrivate) {
        $("vapidKey").value = newPrivate; $("vapidBox").hidden = false;
      } else showToast("Reminders on. Moss will nudge you.");
      renderPushStatus();
    } catch (e) { showToast("Couldn't turn on reminders: " + e.message, null, 6000); renderPushStatus(); }
  }
  async function disablePush() {
    try {
      const reg = await navigator.serviceWorker.ready;
      const sub = await reg.pushManager.getSubscription();
      if (sub) await sub.unsubscribe();
      showToast("Reminders off on this phone.");
    } catch (e) {}
    renderPushStatus();
  }

  // ---------- Backdating ----------
  function setActive(key) {
    if (!key || key >= todayKey()) key = null;
    activeKey = key;
    const past = !!key;
    const label = past ? new Date(key + "T12:00:00").toLocaleDateString("en-GB", { weekday: "long", day: "numeric", month: "short" }) : "Today";
    $("backdate").hidden = !past;
    $("backdateText").textContent = `Logging for ${label}`;
    $("logDayText").textContent = label;
    document.querySelector(".log-day").classList.toggle("past", past);
    $("dayInput").value = dayKey(); $("logDayInput").value = dayKey();
    $("dayInput").max = todayKey(); $("logDayInput").max = todayKey();
    todaysPhotos = [];
    renderAll(); loadPhotos();
  }
  const shiftDay = n => { const d = new Date(dayKey() + "T12:00:00"); d.setDate(d.getDate() + n); setActive(isoDate(d)); };

  // ---------- Sheets ----------
  const SHEETS = ["checkinSheet", "logSheet", "mealSheet", "gutSheet", "moveSheet", "numSheet", "compareSheet", "settingsSheet", "productSheet"];
  function openSheet(id) { SHEETS.forEach(s => { $(s).hidden = s !== id; }); $("scrim").hidden = false; }
  function closeSheets() { SHEETS.forEach(s => { $(s).hidden = true; }); $("scrim").hidden = true; }

  let sheetField = null;
  function openMealSheet(field, plannedKey) {
    sheetField = field;
    const rec = field ? foodForField(field) : null;
    $("mealTitle").textContent = field ? (rec ? FIELD_SLOT[field] : `What did you have for ${FIELD_SLOT[field].toLowerCase()}?`) : "Snacks & extras";
    const slotName = field ? FIELD_SLOT[field] : "Snack";
    const items = foodsInSlot(slotName);
    const lg = $("mealLogged"); lg.textContent = ""; lg.hidden = !items.length;
    items.forEach(it => {
      const wrapI = document.createElement("div");
      const row = document.createElement("div"); row.className = "logged-item";
      const b = document.createElement("b"); b.textContent = it.name;
      const x = document.createElement("button"); x.className = "x"; x.textContent = "Remove";
      x.onclick = () => { removeFood(it.key); openMealSheet(field, plannedKey); };
      row.append(b, x); wrapI.appendChild(row);
      const meta = document.createElement("div"); meta.className = "ln-meta";
      meta.textContent = it.pendingAI ? "Estimating…" : it.calories ? `~${fmt(it.calories)} kcal · ${it.protein_g}g protein${it.flags && it.flags.length ? " · contains " + it.flags.join(", ") : ""}` : "No estimate yet";
      wrapI.appendChild(meta);
      if (!it.calories && !it.pendingAI && it.source === "text") {
        const rb = document.createElement("button"); rb.className = "btn ghost sm"; rb.textContent = "Estimate calories";
        rb.onclick = async () => { closeSheets(); await logText(SLOT_FIELD[it.slot] || null, it.name, { slot: it.slot, time: it.time, key: it.key }); };
        wrapI.appendChild(rb);
      }
      if (it.note) { const n = document.createElement("div"); n.className = "ln-note"; n.textContent = it.note; wrapI.appendChild(n); }
      if (it.calories && it.source !== "barcode") {
        const pr = document.createElement("div"); pr.className = "portion";
        pr.innerHTML = '<span class="muted small">Portion</span><div class="seg"><button data-p="0.5">½</button><button data-p="1">1×</button><button data-p="1.5">1½</button><button data-p="2">2×</button></div>';
        pr.querySelectorAll("button").forEach(btn => {
          btn.classList.toggle("on", Number(btn.dataset.p) === (it.portion || 1));
          btn.onclick = () => { setPortion(it.key, Number(btn.dataset.p)); openMealSheet(field, plannedKey); };
        });
        wrapI.appendChild(pr);
      }
      lg.appendChild(wrapI);
    });
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
    $("mealNone").hidden = true;
    openSheet("mealSheet");
  }

  let numKind = "weight", numVal = 84;
  function openNum(kind) {
    numKind = kind;
    const t = getDay(dayKey()), s = series(kind);
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
    const g = getDay(dayKey()).gut;
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
    $("dayPrev").onclick = () => shiftDay(-1);
    $("dayNext").onclick = () => shiftDay(1);
    $("dayInput").addEventListener("change", e => setActive(e.target.value));
    $("logDayInput").addEventListener("change", e => setActive(e.target.value));
    $("backdate").onclick = () => { setActive(null); showToast("Back to today."); };
    $("scrim").onclick = closeSheets;
    document.querySelectorAll("[data-close]").forEach(b => b.onclick = closeSheets);
    $("openSettings").onclick = () => {
      $("setKcal").value = targets.kcal; $("setPro").value = targets.protein;
      $("whoami").textContent = account ? `Signed in as ${account.username}` : "";
      $("vapidBox").hidden = true; renderPushStatus();
      openSheet("settingsSheet");
    };
    $("setSave").onclick = () => {
      const k = parseInt($("setKcal").value, 10), p = parseInt($("setPro").value, 10);
      if (k >= 1000 && k <= 4000) targets.kcal = k;
      if (p >= 50 && p <= 300) targets.protein = p;
      store.set("fuel.targets", targets); closeSheets(); renderAll(); showToast("Targets saved.");
    };
    $("pushBtn").onclick = enablePush;
    $("pushOff").onclick = disablePush;
    $("vapidCopy").onclick = async () => { try { await navigator.clipboard.writeText($("vapidKey").value); showToast("Key copied. Paste it into GitHub."); } catch (e) { $("vapidKey").select(); } };
    $("signOut").onclick = () => pca.logoutRedirect({ account });

    const pick = (id, slotFn) => $(id).addEventListener("change", e => { const f = e.target.files && e.target.files[0]; e.target.value = ""; if (f) { const s = slotFn ? slotFn() : undefined; closeSheets(); onPhoto(f, s); } });
    pick("camera"); pick("camera2"); pick("lmCamera"); pick("lmGallery");
    const sheetSlot = () => sheetField ? FIELD_SLOT[sheetField] : "Snack";
    pick("mealCam", sheetSlot); pick("mealLib", sheetSlot);

    document.querySelectorAll("#logSheet [data-act]").forEach(b => b.onclick = () => {
      const a = b.dataset.act;
      if (a === "checkin") openCheckin();
      if (a === "type") openMealSheet(null);
      if (a === "scan") openScanner(null);
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
    $("mealScan").onclick = () => openScanner(sheetField ? FIELD_SLOT[sheetField] : "Snack");
    $("scanClose").onclick = stopScanner;
    $("scanManual").addEventListener("submit", e => { e.preventDefault(); const c = $("scanCode").value; if (c) { stopScanner(); lookup(c); } });
    $("prodGrams").addEventListener("input", () => {
      const g = parseFloat($("prodGrams").value);
      if (g > 0 && g < 5000) { prodGrams = g; const v = prodNumbers(g); $("prodTotal").firstChild.textContent = `${fmt(v.kcal)} kcal`; $("prodTotal").querySelector("small").textContent = `${v.protein}g protein · ${Math.round(g)} g`; document.querySelectorAll("#prodAmounts .opt").forEach(o => o.classList.remove("on")); }
    });
    document.querySelectorAll("#prodSlot button").forEach(b => b.onclick = () => { prodSlot = b.dataset.slot; document.querySelectorAll("#prodSlot button").forEach(x => x.classList.toggle("on", x === b)); });
    $("prodLog").onclick = logProduct;
    $("prodRescan").onclick = () => openScanner(prodSlot);
    $("mealNone").onclick = () => {
      const rec = sheetField ? foodForField(sheetField) : null;
      closeSheets();
      if (rec) { removeFood(rec.key); showToast(`${FIELD_SLOT[sheetField] || "Item"} removed.`); }
    };
    document.querySelectorAll("#gutSheet .chip").forEach(c => c.onclick = () => {
      const v = getDay(dayKey()).gut === c.dataset.v ? null : c.dataset.v;
      change({ gut: v }); closeSheets(); if (v) showToast(v === "good" ? "Nice. Logged." : "Logged. I'll look for patterns.");
    });
    document.querySelectorAll("#checkinSheet .ci-opts").forEach(g => g.querySelectorAll("button").forEach(b => b.onclick = () => {
      const f = g.dataset.field, cur = getDay(dayKey())[f];
      const v = f === "mood" ? Number(b.dataset.v) : b.dataset.v;
      change({ [f]: String(cur) === b.dataset.v ? null : v });
      g.querySelectorAll("button").forEach(x => x.classList.toggle("on", x === b && String(cur) !== b.dataset.v));
    }));
    $("ciDone").onclick = () => {
      const t = getDay(dayKey()); closeSheets();
      showToast(t.gut === "rough" || t.mood <= 2 || t.energy === "low" ? "Logged. Gentle day then. One small thing at a time." : "Checked in. Moss is awake now.");
    };
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
    try { if (pending.size) await flush(); await pull(); renderAll(); loadPhotos(); loadProgress(); setStatus("Synced"); repairFoods(); }
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
    wire(); setActive(null); show("home"); renderBody();
    // Deep links from notifications: #gut, #weight, #waist, #move, #snap, #body, #trends, #b, #l, #d
    const h = location.hash.replace("#", "");
    if (h === "trends" || h === "body" || h === "home") show(h);
    else if (h === "food") show("today");
    else if (h === "checkin") openCheckin();
    else if (h === "gut") openGut();
    else if (h === "move") openSheet("moveSheet");
    else if (h === "weight" || h === "waist") openNum(h);
    else if (["b", "l", "d"].includes(h)) openMealSheet(h, h === "b" ? "b" : h === "l" ? LUNCH[today().getDay()] : DINNER[today().getDay()]);
    else if (h === "log") openSheet("logSheet");
    else if (h === "scan") openScanner(null);
    else if (h === "reminders") $("openSettings").click();
    if (h) history.replaceState(null, "", "/");
    setStatus("Syncing…");
    refresh();
  }
  start();
})();
