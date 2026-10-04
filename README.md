# Japanese Text Analyzer

A installable web app (PWA) that lets you photograph Japanese text, crop a
rectangular region, and get an AI-powered breakdown tailored to your JLPT level:

- Recognized text and translation with the photographed line layout preserved,
  plus an on-card furigana toggle for the recognized Japanese
- A list of the most important words (with reading/romaji/meaning)
- Expandable kanji decomposition below each multi-kanji vocabulary word
- Explanation of the relevant grammar points
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
- **Rotate + crop**: the photo is drawn to a `<canvas>`; rotate it left or
  right as needed, then drag (mouse or touch) to select a rectangle. Only the
  cropped region is sent to the model (cheaper, more private, more accurate
  OCR than the whole photo).
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
  in installed mobile PWAs.
- **Installable (PWA)**: `manifest.webmanifest` + `sw.js` let you "Add to Home
  screen" on both the Samsung S24 and the Lenovo tablet, so it behaves like an
  app icon, full-screen, works from any browser without an app-store install.

## Azure resources already provisioned for you

| Resource | Value |
|---|---|
| Azure OpenAI resource | `sommerlernplan-ai-a8fbd8e1` (resource group `rg-sommerlernplan-2026`, region germanywestcentral) |
| Model deployment | `japanese-text-analyzer` → `gpt-5.4-mini` (vision-capable) |
| Entra ID App registration | "Japanese Text Analyzer" (client ID pre-filled in Settings) |
| Redirect URIs registered | Dedicated `auth.html` callback URLs for GitHub Pages and local development, plus the original app URLs |

These values are already pre-filled as defaults in Settings — you normally
don't need to touch them. If you ever redeploy to a different URL, add that
URL as a redirect URI on the app registration (Entra admin center → App
registrations → Japanese Text Analyzer → Authentication), otherwise sign-in
will fail with a redirect URI mismatch.

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

## Deploying / updating

The app is a static site published via **GitHub Pages** from this repo's
`main` branch (root). Just commit and push changes — GitHub Pages rebuilds
automatically within a minute or two.

## Cost note

Only the cropped snippet of the photo (not the full photo) is sent per
analysis, using a small/cheap GPT model deployment, to keep Azure OpenAI
token costs low.
