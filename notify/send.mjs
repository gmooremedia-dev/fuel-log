// Fuel reminders: runs every 30 minutes on GitHub Actions. Reads today's row from the Fuel Log list
// (read-only, app-only Graph) and sends a push to George's phone only when something is still missing.
import webpush from "web-push";
import fs from "node:fs";

const { TENANT_ID, CLIENT_ID, CLIENT_SECRET, VAPID_PRIVATE_KEY, SITE_URL, MODE = "scheduled" } = process.env;
const STATE = "notify/.state/sent.json";
const TZ = "Europe/London";

const need = ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", "VAPID_PRIVATE_KEY", "SITE_URL"].filter(k => !process.env[k]);
if (need.length) { console.log(`Reminders not set up yet; missing: ${need.join(", ")}`); process.exit(0); }

// ---------- time in London ----------
const now = new Date();
const parts = Object.fromEntries(new Intl.DateTimeFormat("en-GB", { timeZone: TZ, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", weekday: "short", hour12: false }).formatToParts(now).map(p => [p.type, p.value]));
const date = `${parts.year}-${parts.month}-${parts.day}`;
const mins = Number(parts.hour) * 60 + Number(parts.minute);
const dow = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].indexOf(parts.weekday);
const hash = s => { let h = 0; for (const c of s) h = (h * 31 + c.charCodeAt(0)) >>> 0; return h; };
const pick = (arr, salt) => arr[hash(date + salt) % arr.length];

// ---------- Graph ----------
async function token() {
  const r = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: CLIENT_ID, client_secret: CLIENT_SECRET, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" })
  });
  const j = await r.json();
  if (!r.ok) throw new Error(`Token failed: ${j.error_description || r.status}`);
  return j.access_token;
}
const tk = await token();
const graph = async path => {
  const r = await fetch(path.startsWith("http") ? path : "https://graph.microsoft.com/v1.0" + path, { headers: { Authorization: "Bearer " + tk, Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } });
  const j = await r.json();
  if (!r.ok) throw new Error(`Graph ${r.status}: ${j.error && j.error.message}`);
  return j;
};
const u = new URL(SITE_URL);
const site = await graph(`/sites/${u.hostname}:${u.pathname.replace(/\/$/, "")}`);
const lists = (await graph(`/sites/${site.id}/lists?$select=id,displayName&$top=200`)).value;
const listId = name => (lists.find(l => l.displayName === name) || {}).id;
const items = async (name, select) => {
  const id = listId(name); if (!id) return [];
  return (await graph(`/sites/${site.id}/lists/${id}/items?$expand=fields($select=${select})&$top=500`)).value.map(i => i.fields);
};

const config = await items("Fuel Config", "Title,Value");
const vapidPublic = (config.find(c => c.Title === "vapidPublicKey") || {}).Value;
if (!vapidPublic) { console.log("No VAPID public key yet: turn on reminders in the app first."); process.exit(0); }
webpush.setVapidDetails("mailto:george@shirestone.uk", vapidPublic, VAPID_PRIVATE_KEY);

const subs = (await items("Fuel Push", "Title,Subscription,Enabled")).filter(s => s.Enabled !== false && s.Subscription);
if (!subs.length) { console.log("No devices subscribed."); process.exit(0); }

const days = await items("Fuel Log", "Title,Breakfast,Lunch,Dinner,Gut,Moved,WeightKg,Mood,Calories,ProteinG,FoodsJson");
const t = days.find(d => d.Title === date) || {};
const foods = (() => { try { return JSON.parse(t.FoodsJson || "[]"); } catch { return []; } })();

// ---------- what to send ----------
const hm = s => { const [h, m] = s.split(":").map(Number); return h * 60 + m; };
const plan = { 0: "your pick", 1: "salmon and Tenderstem", 2: "paprika chicken", 3: "burgers and sweet potato fries", 4: "salmon and Tenderstem", 5: "paprika chicken", 6: "burgers and sweet potato fries" };
const lunch = { 0: "a chicken rice bowl", 1: "a chicken rice bowl", 2: "a tuna jacket", 3: "a chicken rice bowl", 4: "a tuna jacket", 5: "a chicken rice bowl", 6: "a tuna jacket" };
const done = ["Breakfast", "Lunch", "Dinner"].filter(k => t[k]).length + (t.Moved ? 1 : 0) + (t.Mood != null ? 1 : 0) + (t.WeightKg != null ? 1 : 0);

const REMINDERS = [
  { id: "checkin", at: "07:30", when: () => t.Mood == null && !t.Gut, url: "/#checkin",
    bodies: ["Morning. How's the gut? Good, meh or rough, one tap.", "Sun's up. Quick check-in and I'll wake up properly.", "Morning! Four taps: gut, mood, energy, sleep. All optional."] },
  { id: "weigh", at: "07:45", days: [1], when: () => t.WeightKg == null, url: "/#weight",
    bodies: ["Weigh-in day. Scales, after the loo, two minutes tops.", "Monday weigh-in. Then a waist measurement if the tape's handy."] },
  { id: "breakfast", at: "09:00", when: () => !t.Breakfast, url: "/#b",
    bodies: ["Smoothie time? Moss is thirsty just thinking about it.", "Breakfast logged yet? Smoothie or whatever you had, one tap.", "Fuel first, emails second. What did you have?"] },
  { id: "water", at: "11:00", when: () => true, url: "/#home",
    bodies: ["Glass of water and a stretch. I'll wait.", "Quick one: water. That's it, that's the reminder.", "Shoulders down, sip of water, carry on."] },
  { id: "lunch", at: "13:30", when: () => !t.Lunch, url: "/#l",
    bodies: [`Lunch o'clock. Plan says ${lunch[dow]}. Snap it or tap it.`, "Lunch logged? Snap it, scan it or type it.", "I'm nosy about lunch. What did you have?"] },
  { id: "snack", at: "16:00", when: () => true, url: "/#log",
    bodies: ["Afternoon slump check: snack, water or a walk?", "Had a snack? Snap it so it counts.", "Four o'clock. Water, then one tiny thing."] },
  { id: "prep", at: "17:30", days: [0], when: () => true, url: "/#home",
    bodies: ["Sunday reset: smoothie bags, batch cook, Ocado basket. Pick one.", "Ten minutes of prep now saves you all week. Smoothie bags first?"] },
  { id: "workout", at: "18:00", when: () => !t.Moved, url: "/#move",
    bodies: ["20 minutes of moving and Moss gets two sunlight.", "Press-ups, squats, lunges, rows, plank. Three rounds. Go on.", "Workout's the last big sunlight of the day. You in?"] },
  { id: "dinner", at: "19:30", when: () => !t.Dinner, url: "/#d",
    bodies: [`Dinner logged yet? Plan says ${plan[dow]}.`, "Dinner time. Snap it before you eat it.", "What's for dinner? One tap and you're done."] },
  { id: "wrap", at: "21:00", when: () => true, url: "/#home",
    bodies: [done ? `You did ${done} thing${done === 1 ? "" : "s"} today${t.Calories ? `, ~${t.Calories} kcal logged` : ""}. Have a look, then phone down.` : "Quiet day? That's allowed. Tomorrow's a fresh pot of sunshine."] }
];

// ---------- state: what's already been sent today ----------
let state = {};
try { state = JSON.parse(fs.readFileSync(STATE, "utf8")); } catch {}
for (const k of Object.keys(state)) if (k !== date) delete state[k];
state[date] = state[date] || [];

let due = [];
if (MODE === "test") due = [{ id: "test", url: "/#home", bodies: ["Hello from Moss! Reminders are working."] }];
else if (mins >= hm("21:45") || mins < hm("07:00")) console.log("Quiet hours.");
else due = REMINDERS.filter(r => (!r.days || r.days.includes(dow)) && mins >= hm(r.at) && mins < hm(r.at) + 90 && !state[date].includes(r.id) && r.when());

// Back-off: only the most recent due reminder goes out in one run, so a late run never sends a pile.
due = due.slice(-1);
for (const r of due) {
  const payload = JSON.stringify({ title: "Moss", body: pick(r.bodies, r.id), url: r.url, tag: r.id });
  for (const s of subs) {
    try { await webpush.sendNotification(JSON.parse(s.Subscription), payload); console.log(`Sent ${r.id} to ${s.Title}`); }
    catch (e) { console.log(`Failed ${r.id} to ${s.Title}: ${e.statusCode || e.message}`); }
  }
  if (r.id !== "test") state[date].push(r.id);
}
// mark skipped-because-done reminders too, so they don't fire later in the window
for (const r of REMINDERS) if ((!r.days || r.days.includes(dow)) && mins >= hm(r.at) && !r.when() && !state[date].includes(r.id)) state[date].push(r.id);
fs.mkdirSync("notify/.state", { recursive: true });
fs.writeFileSync(STATE, JSON.stringify(state));
console.log(`London ${date} ${parts.hour}:${parts.minute}; sent today: ${state[date].join(", ") || "none"}`);
