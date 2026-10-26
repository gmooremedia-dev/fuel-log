# Fuel Log

George's meal, gut, workout and weight tracker. A phone-first web app (PWA) hosted on Azure Static Web Apps, signing in with Shirestone Entra ID and saving everything to SharePoint.

## What's already set up

| Piece | Where | Value |
|---|---|---|
| App registration | Entra → App registrations → Fuel Log | Client ID `c11feb51-df62-491d-917c-fa03c228e29c` |
| Tenant | Shirestone | `c4e0279c-d022-4e37-b9d7-6a4676219aeb` |
| Permissions | Microsoft Graph, delegated | `User.Read`, `Sites.Manage.All` (no admin consent needed) |
| Redirect URIs (SPA) | | `https://fuel.shirestone.uk`, `https://blue-water-08e10e110.4.azurestaticapps.net` |
| SharePoint site | | `https://shirestone581.sharepoint.com/sites/fuel` |
| Static Web App | Azure → rg-fuel → swa-fuel (Free plan) | `https://blue-water-08e10e110.4.azurestaticapps.net` |

The app creates its own **Fuel Log** list and **Food Photos** folder on the site the first time you sign in.

## Deploys (GitHub)

Every push to `main` deploys automatically via `.github/workflows/deploy.yml`.
One-time setup: GitHub repo → Settings → Secrets and variables → Actions → New repository secret
`AZURE_STATIC_WEB_APPS_API_TOKEN` = the token from swa-fuel → Overview → Manage deployment token.

## AI food spotting

`api/src/functions/analyze.js` runs as a managed function inside the Static Web App (free).
It checks the caller is signed in as an allowed Shirestone account, then asks Claude to estimate the plate.
Set these in swa-fuel → Settings → Environment variables (Production):

| Name | Value |
|---|---|
| `ANTHROPIC_API_KEY` | your key from console.anthropic.com |
| `ALLOWED_USERS` | `george@shirestone.uk` |
| `FOOD_MODEL` | optional, defaults to `claude-sonnet-5-5` |

Estimates are saved to the day's row in the Fuel Log list (`FoodsJson`, `Calories`, `ProteinG`).

## Manual deploy (fallback, from your PC)

You need Node.js (LTS) installed.

1. Azure portal → swa-fuel → Overview → **Manage deployment token** → copy it.
2. Unzip this folder, open a terminal in it, and run:

```
npx @azure/static-web-apps-cli deploy ./app --deployment-token YOUR_TOKEN --env production
```

3. Open the site on your iPhone in Safari → Share → **Add to Home Screen**.

Redeploy the same way whenever the app changes.

## Custom domain (fuel.shirestone.uk)

1. swa-fuel → Custom domains → Add → Custom domain on other DNS → `fuel.shirestone.uk`
2. At your DNS provider, add a CNAME: `fuel` → `blue-water-08e10e110.4.azurestaticapps.net`
3. Back in Azure, Validate. The SSL certificate is issued automatically.

## Where the data lives

- **Fuel Log list:** one row per day. Title = date (YYYY-MM-DD), Breakfast / Lunch / Dinner / Moved (yes/no), Gut (good / meh / rough), WeightKg.
- **Documents → Food Photos → YYYY-MM-DD:** photos named like `1312-Lunch.jpg`, shrunk to 1600px before upload.

## Changing the meal plan

Edit the `MEALS`, `LUNCH` and `DINNER` tables at the top of `app/app.js`, then redeploy.

## Files

- `app/index.html`, `app/styles.css`, `app/app.js` – the app
- `app/config.js` – client ID, tenant ID, site URL (not secrets)
- `app/msal-browser.min.js` – Microsoft sign-in library (v2.38.3)
- `app/sw.js`, `app/manifest.webmanifest`, `app/icons/` – home-screen app bits
- `app/staticwebapp.config.json` – hosting config
