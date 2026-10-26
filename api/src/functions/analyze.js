// POST /api/analyze  — estimate what's in a food photo with Claude.
// Body: { image: "<base64 jpeg>", slot: "Lunch" }  or  { text: "chicken wrap and crisps", slot: "Lunch" }
// Header: Authorization: Bearer <Microsoft Graph token from the app's sign-in>
// App settings: ANTHROPIC_API_KEY (required), ALLOWED_USERS (comma-separated UPNs), FOOD_MODEL (optional)
const { app } = require("@azure/functions");

const MODEL = process.env.FOOD_MODEL || "claude-sonnet-5-5";
const ALLOWED = (process.env.ALLOWED_USERS || "george@shirestone.uk").toLowerCase().split(",").map(s => s.trim()).filter(Boolean);

const PROMPT = `You are estimating the contents of a meal photo for a personal food log.
The person is lactose intolerant and has IBS-like symptoms, so flag foods that commonly trigger those.
Return ONLY a JSON object, no other text, in exactly this shape:
{
  "items": [ { "name": "short food name", "portion": "e.g. 1 fillet (~130g)" } ],
  "calories": integer estimate for the whole plate,
  "protein_g": integer estimate,
  "flags": [ zero or more of "dairy", "onion", "garlic", "beans", "high-fat", "fried", "spicy", "sweetener", "alcohol" ],
  "confidence": "low" | "medium" | "high",
  "note": "one short, plain sentence: the most useful observation (e.g. protein looks light, or sauce may hide dairy). No lecturing."
}
If it isn't food, return {"items":[],"calories":0,"protein_g":0,"flags":[],"confidence":"low","note":"No food spotted in this photo."}`;

async function whoIs(authHeader) {
  if (!authHeader || !authHeader.startsWith("Bearer ")) return null;
  const r = await fetch("https://graph.microsoft.com/v1.0/me?$select=userPrincipalName", { headers: { Authorization: authHeader } });
  if (!r.ok) return null;
  const me = await r.json();
  return (me.userPrincipalName || "").toLowerCase();
}

app.http("analyze", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    const upn = await whoIs(request.headers.get("authorization"));
    if (!upn || !ALLOWED.includes(upn)) return { status: 401, jsonBody: { error: "Sign in with an allowed Shirestone account." } };
    if (!process.env.ANTHROPIC_API_KEY) return { status: 503, jsonBody: { error: "Food spotting isn't set up yet: add ANTHROPIC_API_KEY in the Static Web App's environment variables." } };

    let body;
    try { body = await request.json(); } catch { return { status: 400, jsonBody: { error: "Expected JSON." } }; }
    const image = typeof body.image === "string" ? body.image.replace(/^data:image\/\w+;base64,/, "") : "";
    const text = typeof body.text === "string" ? body.text.trim().slice(0, 200) : "";
    if (!image && !text) return { status: 400, jsonBody: { error: "Send a photo or a description." } };
    if (image.length > 6_000_000) return { status: 400, jsonBody: { error: "That photo is too big." } };
    const slot = String(body.slot || "unknown").slice(0, 20);
    const content = image
      ? [
          { type: "image", source: { type: "base64", media_type: "image/jpeg", data: image } },
          { type: "text", text: `${PROMPT}\n\nMeal slot: ${slot}.` }
        ]
      : [{ type: "text", text: `${PROMPT}\n\nThere is no photo. Estimate from this description of what they ate, assuming typical UK portions: "${text.replace(/"/g, "'")}"\nMeal slot: ${slot}.` }];

    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "x-api-key": process.env.ANTHROPIC_API_KEY,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json"
      },
      body: JSON.stringify({
        model: MODEL,
        max_tokens: 600,
        messages: [{ role: "user", content }]
      })
    });
    if (!res.ok) {
      context.warn(`Anthropic API ${res.status}: ${(await res.text()).slice(0, 300)}`);
      return { status: 502, jsonBody: { error: res.status === 429 ? "AI is busy, try again in a minute." : "Couldn't analyse that photo." } };
    }
    const data = await res.json();
    const answer = (data.content || []).filter(c => c.type === "text").map(c => c.text).join("");
    const match = answer.match(/\{[\s\S]*\}/);
    let result;
    try { result = JSON.parse(match ? match[0] : answer); } catch { return { status: 502, jsonBody: { error: "Couldn't read the AI's answer." } }; }
    const clean = {
      items: Array.isArray(result.items) ? result.items.slice(0, 12).map(i => ({ name: String(i.name || "").slice(0, 60), portion: String(i.portion || "").slice(0, 60) })) : [],
      calories: Math.max(0, Math.round(Number(result.calories) || 0)),
      protein_g: Math.max(0, Math.round(Number(result.protein_g) || 0)),
      flags: Array.isArray(result.flags) ? result.flags.map(String).slice(0, 9) : [],
      confidence: ["low", "medium", "high"].includes(result.confidence) ? result.confidence : "low",
      note: String(result.note || "").slice(0, 200)
    };
    return { status: 200, jsonBody: clean };
  }
});
