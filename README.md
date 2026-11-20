# Japanese Text Analyzer

A installable web app (PWA) that lets you photograph Japanese text, crop a
rectangular region, and get an AI-powered breakdown tailored to your JLPT level:

- Recognized text and translation with the photographed line layout preserved,
  plus an on-card furigana toggle for the recognized Japanese
- A list of the most important words (with reading/romaji/meaning)
- Expandable kanji decomposition below each multi-kanji vocabulary word
- Explanation of the relevant grammar points
- One-tap copy controls for each result section
- Automatic textbook-exercise detection with answer fields, saved drafts, and
  AI correction after submission
- English/Japanese interface captions, selectable in Settings and remembered
  on each device

The explanation depth adapts to the **JLPT level** you pick on the main screen
(N5 → N1): more advanced levels skip explanations of vocabulary/grammar you'd
already be expected to know.

## How it works / architecture

- **Pure static site** — `index.html` + `css/` + `js/i18n.js` + `js/app.js`.
  No backend server.
- **Interface localization**: Settings lets you switch app captions between
  English and Japanese without reloading or losing the current analysis or
  exercise answers. This affects interface captions only: recognized content
  remains Japanese, and translations and learning explanations always remain
  English.
- **Camera / photo**: uses a plain `<input type="file" capture="environment">`,
  which opens the camera (or gallery) on both Android phones and tablets.
- **Rotate + crop**: lightweight preview canvases (at most 1600 pixels on
  their longest side) let you rotate and select a rectangle. Cropping uses
  the source image, not the preview, so the selected text retains its original
  resolution by default. Only the cropped region is sent to the model.
- **Image memory**: photos load through temporary blob URLs rather than
  full-photo base64 strings. Replacing a photo releases the previous image
  and canvas buffers; the temporary export canvas is also released after
  cropping. Settings offers optional 4096- or 2048-pixel longest-side limits,
  applied to the next selected photo without enlarging small images. Limits
  reduce working-image memory but can affect recognition of small kanji.
  **Original** is the default. Decoding a large original and exporting a large
  crop still require memory; these changes cannot prevent Android from
  killing the browser while the external camera is open. Photos are not
  automatically saved across page restarts (except the existing sign-in crop
  recovery described below).
- **AI backend**: your own **Azure OpenAI** resource (GPT vision model). The
  app calls the Chat Completions REST API directly from the browser and uses
  strict JSON Schema output so analysis responses remain machine-readable.
- **Exercise mode**: Auto-detect distinguishes normal passages from textbook
  exercises. You can force Reading text or Textbook exercise mode before
  analyzing. Detected questions become multiple-choice or typed-answer fields;
  answers stay in browser storage and are sent for correction only when you
  tap **Check answers**. Pattern-practice exercises retain the model sentence
  and present each new phrase combination as a sentence-building prompt.
- **Authentication**: this Azure environment enforces "no API keys" (local
  auth disabled by policy), so the app signs you in with your **Microsoft
  Entra ID** account via [MSAL.js](https://github.com/AzureAD/microsoft-authentication-library-for-js)
  and calls Azure OpenAI with your personal delegated token. No secret is
  ever stored — only your Azure OpenAI endpoint, deployment name, tenant ID
  and app (client) ID are kept in the browser's `localStorage`. Authentication
  uses a full-page redirect rather than nested popups, which is more reliable
  in installed mobile PWAs. The app checks token readiness when it starts and
  resumes; if sign-in expires after an image is selected, the cropped image is
  kept briefly in IndexedDB and restored after authentication.
- **Installable (PWA)**: `manifest.webmanifest` + `sw.js` let you "Add to Home
  screen" on both the Samsung S24 and the Lenovo tablet, so it behaves like an
  app icon, full-screen, works from any browser without an app-store install.

## Dedicated Azure backend

| Resource | Value |
|---|---|
| Foundry resource | `foundry-japanese-text-analyzer-a8fbd8` (resource group `rg-japanese-text-analyzer`, region germanywestcentral) |
| Foundry project | `japanese-text-analyzer` (display name "Japanese Text Analyzer") |
| Inference endpoint | `https://foundry-japanese-text-analyzer-a8fbd8.openai.azure.com` |
| Model deployment | `japanese-text-analyzer` → `gpt-5.4-mini` (vision-capable) |
| Entra ID App registration | "Japanese Text Analyzer" (client ID pre-filled in Settings) |
| Redirect URIs registered | Dedicated `auth.html` callback URLs for GitHub Pages and local development, plus the original app URLs |

These values are already pre-filled as defaults in Settings — you normally
don't need to touch them. If you ever redeploy to a different URL, add that
URL as a redirect URI on the app registration (Entra admin center → App
registrations → Japanese Text Analyzer → Authentication), otherwise sign-in
will fail with a redirect URI mismatch.

This backend is dedicated to the live app, not shared with temporary demos.
Resource tags identify the application, its purpose, and this repository.
API-key authentication is disabled; inference access uses the **Cognitive
Services OpenAI User** role scoped to this resource. Public network access
is required by the current direct-from-browser architecture.

The previous `sommerlernplan-ai-a8fbd8e1` backend was deleted on 19 October
2026. When loading saved settings for that endpoint and the standard
`japanese-text-analyzer` deployment, the app automatically substitutes the
new endpoint. Other settings and custom endpoints/deployments are left
unchanged. Tapping **Save** also persists the replacement to browser storage.

### Recreating the backend

`infra/main.bicep` defines the Foundry resource, project, model deployment,
and resource-scoped inference role. From PowerShell, select the intended
subscription and supply the Entra **user object ID**, not the application
client ID:

```powershell
az account set --subscription <subscription-id>
az group create --name rg-japanese-text-analyzer --location germanywestcentral --tags application=japanese-text-analyzer purpose=live-app-backend
az deployment group create --resource-group rg-japanese-text-analyzer --template-file .\infra\main.bicep --parameters inferenceUserObjectId=<user-object-id>
```

The model uses consumption-based `GlobalStandard` capacity of 30 (30,000
tokens/minute), not prepaid provisioned throughput. The resource is located
in Germany West Central, but Global Standard inference can process requests
globally. The Foundry project does not change the app's existing Chat
Completions API or create an agent.

## First-time setup on a device

1. Open the app URL (GitHub Pages link) in Chrome (Android/S24) or any modern
   browser (Lenovo tablet).
2. Tap **Sign in** at the top of the app → sign in with your Microsoft account
   that has access to the Azure subscription. The first time, you may see a
   permission consent screen for "Microsoft Cognitive Services" — accept it.
3. Tap **Save**.
4. (Optional) Tap the browser menu → **Install app** / **Add to Home screen**
   so it opens full-screen like a native app.
5. Pick your JLPT level and analysis mode, take/choose a photo, drag to select
   the text region, and tap **Analyze**.

## Local development

Because Entra ID sign-in (MSAL) requires `http(s)://`, you can't just double
click `index.html` (file:// origin won't work for sign-in). Serve it locally:

```powershell
# from the JapaneseApp folder
python -m http.server 5500
# then open http://localhost:5500/
```

(`http://localhost:5500/` is already registered as a redirect URI.)

Run the settings-migration regression tests with Node.js (no packages required):

```powershell
node --test .\tests\settings.test.cjs
```

Browser image tests cover preview allocations, original-resolution crops,
rotation, resize, resolution settings, image replacement, and sign-in crop
recovery. They require Python with Playwright and Pillow, plus Microsoft Edge:

```powershell
python .\tests\image-memory.browser.py
```

## Deploying / updating

The app is a static site published via **GitHub Pages** from this repo's
`main` branch (root). Just commit and push changes — GitHub Pages rebuilds
automatically within a minute or two.

## Cost note

Only the cropped snippet of the photo (not the full photo) is sent per
analysis, using a small/cheap GPT model deployment, to keep Azure OpenAI
token costs low.
