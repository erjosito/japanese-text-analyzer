// Japanese Text Analyzer - client-side app logic
// Uses MSAL.js (Entra ID) to get a token for the user's own Azure OpenAI resource,
// calls it directly from the browser (no backend), and renders a structured analysis.

const LS_KEY = 'jta_settings_v1';

const DEFAULT_SETTINGS = {
  endpoint: 'https://sommerlernplan-ai-a8fbd8e1.openai.azure.com',
  deployment: 'japanese-text-analyzer',
  apiVersion: '2024-10-21',
  tenantId: '5ad00b69-0386-4c74-8adc-ac7a28649f34',
  clientId: 'b9852c8b-060f-4915-9850-9d185e19b3e5',
  furigana: true,
  romaji: false,
};

function loadSettings() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) return { ...DEFAULT_SETTINGS };
    return { ...DEFAULT_SETTINGS, ...JSON.parse(raw) };
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

function saveSettings(s) {
  localStorage.setItem(LS_KEY, JSON.stringify(s));
}

let settings = loadSettings();

// ---------- DOM refs ----------
const els = {
  btnSettings: document.getElementById('btn-settings'),
  btnAccount: document.getElementById('btn-account'),
  accountIndicator: document.getElementById('account-indicator'),
  accountLabel: document.getElementById('account-label'),
  settingsModal: document.getElementById('settings-modal'),
  btnCloseSettings: document.getElementById('btn-close-settings'),
  btnSaveSettings: document.getElementById('btn-save-settings'),
  cfgEndpoint: document.getElementById('cfg-endpoint'),
  cfgDeployment: document.getElementById('cfg-deployment'),
  cfgApiVersion: document.getElementById('cfg-apiversion'),
  cfgTenant: document.getElementById('cfg-tenant'),
  cfgClientId: document.getElementById('cfg-clientid'),
  cfgFurigana: document.getElementById('cfg-furigana'),
  cfgRomaji: document.getElementById('cfg-romaji'),
  btnSignin: document.getElementById('btn-signin'),
  btnSignout: document.getElementById('btn-signout'),
  accountStatus: document.getElementById('account-status'),
  levelSelect: document.getElementById('level-select'),
  analysisModeSelect: document.getElementById('analysis-mode-select'),
  captureEmpty: document.getElementById('capture-empty'),
  captureEditor: document.getElementById('capture-editor'),
  btnPickImage: document.getElementById('btn-pick-image'),
  fileInput: document.getElementById('file-input'),
  photoCanvas: document.getElementById('photo-canvas'),
  selectionCanvas: document.getElementById('selection-canvas'),
  cropStage: document.getElementById('crop-stage'),
  btnRetake: document.getElementById('btn-retake'),
  btnRotateLeft: document.getElementById('btn-rotate-left'),
  btnRotateRight: document.getElementById('btn-rotate-right'),
  btnResetCrop: document.getElementById('btn-reset-crop'),
  btnAnalyze: document.getElementById('btn-analyze'),
  resultArea: document.getElementById('result-area'),
  toast: document.getElementById('toast'),
};

function toast(msg, ms = 3000) {
  els.toast.textContent = msg;
  els.toast.classList.remove('hidden');
  clearTimeout(toast._t);
  toast._t = setTimeout(() => els.toast.classList.add('hidden'), ms);
}

// ---------- Settings modal ----------
function openSettingsModal() {
  els.cfgEndpoint.value = settings.endpoint || '';
  els.cfgDeployment.value = settings.deployment || '';
  els.cfgApiVersion.value = settings.apiVersion || '';
  els.cfgTenant.value = settings.tenantId || '';
  els.cfgClientId.value = settings.clientId || '';
  els.cfgFurigana.checked = !!settings.furigana;
  els.cfgRomaji.checked = !!settings.romaji;
  refreshAccountStatus();
  els.settingsModal.showModal();
}

els.btnSettings.addEventListener('click', openSettingsModal);
els.btnAccount.addEventListener('click', openSettingsModal);
els.btnCloseSettings.addEventListener('click', () => els.settingsModal.close());

els.btnSaveSettings.addEventListener('click', (e) => {
  settings.endpoint = els.cfgEndpoint.value.trim().replace(/\/+$/, '');
  settings.deployment = els.cfgDeployment.value.trim();
  settings.apiVersion = els.cfgApiVersion.value.trim() || DEFAULT_SETTINGS.apiVersion;
  settings.tenantId = els.cfgTenant.value.trim();
  settings.clientId = els.cfgClientId.value.trim();
  settings.furigana = els.cfgFurigana.checked;
  settings.romaji = els.cfgRomaji.checked;
  saveSettings(settings);
  toast('Settings saved');
  initMsal(); // re-init in case tenant/client changed
});

// ---------- MSAL auth ----------
let msalApp = null;
let msalReady = Promise.resolve();
let activeAccount = null;

function initMsal() {
  if (!settings.tenantId || !settings.clientId) {
    msalApp = null;
    msalReady = Promise.resolve();
    return msalReady;
  }
  if (typeof msal === 'undefined') {
    msalApp = null;
    msalReady = Promise.reject(new Error('The Microsoft sign-in library failed to load. Refresh the app and try again.'));
    return msalReady;
  }
  const config = {
    auth: {
      clientId: settings.clientId,
      authority: `https://login.microsoftonline.com/${settings.tenantId}`,
      redirectUri: new URL('./', window.location.href).href,
      navigateToLoginRequestUrl: false,
    },
    cache: { cacheLocation: 'localStorage' },
  };
  msalApp = new msal.PublicClientApplication(config);
  msalReady = msalApp.initialize().then(() => msalApp.handleRedirectPromise()).then((response) => {
    if (response?.account) {
      activeAccount = response.account;
    }
    const accounts = msalApp.getAllAccounts();
    if (!activeAccount && accounts.length > 0) {
      activeAccount = accounts[0];
    }
    refreshAccountStatus();
    if (localStorage.getItem('jta_auth_pending') === '1') {
      localStorage.removeItem('jta_auth_pending');
      openSettingsModal();
      if (activeAccount) toast('Signed in');
    }
  });
  return msalReady;
}

function setAccountDisplay(account) {
  if (account) {
    els.accountStatus.textContent = `Signed in as ${account.username}`;
    els.accountLabel.textContent = account.name || account.username;
    els.accountIndicator.classList.add('authenticated');
    els.btnAccount.title = `Signed in as ${account.username}`;
  } else {
    els.accountStatus.textContent = 'Not signed in';
    els.accountLabel.textContent = 'Sign in';
    els.accountIndicator.classList.remove('authenticated');
    els.btnAccount.title = 'Sign in';
  }
}

function refreshAccountStatus() {
  setAccountDisplay(activeAccount);
}

function clearStaleMsalInteraction() {
  sessionStorage.removeItem('msal.interaction.status');
}

els.btnSignin.addEventListener('click', async () => {
  els.btnSignin.disabled = true;
  els.accountStatus.classList.remove('auth-error');
  els.accountStatus.textContent = 'Redirecting to Microsoft sign-in...';
  try {
    if (!msalApp) initMsal();
    if (!msalApp) { toast('Fill in Tenant ID and Client ID first'); return; }
    await msalReady;
    clearStaleMsalInteraction();
    localStorage.setItem('jta_auth_pending', '1');
    await msalApp.loginRedirect({
      scopes: ['https://cognitiveservices.azure.com/user_impersonation'],
    });
  } catch (err) {
    localStorage.removeItem('jta_auth_pending');
    console.error(err);
    els.accountStatus.classList.add('auth-error');
    els.accountStatus.textContent = 'Sign-in failed: ' + (err.message || err);
  } finally {
    els.btnSignin.disabled = false;
  }
});

els.btnSignout.addEventListener('click', async () => {
  if (!msalApp || !activeAccount) return;
  await msalApp.logoutRedirect({
    account: activeAccount,
    postLogoutRedirectUri: new URL('./', window.location.href).href,
  });
});

async function getAccessToken() {
  if (!msalApp) initMsal();
  if (!msalApp) throw new Error('Sign-in not configured. Open Settings and fill Tenant/Client ID.');
  await msalReady;
  const request = {
    scopes: ['https://cognitiveservices.azure.com/user_impersonation'],
    account: activeAccount || msalApp.getAllAccounts()[0],
  };
  if (!request.account) {
    clearStaleMsalInteraction();
    localStorage.setItem('jta_auth_pending', '1');
    await msalApp.loginRedirect(request);
    throw new Error('Redirecting to Microsoft sign-in. Tap Analyze again after returning.');
  }
  try {
    const result = await msalApp.acquireTokenSilent(request);
    return result.accessToken;
  } catch (err) {
    clearStaleMsalInteraction();
    await msalApp.acquireTokenRedirect(request);
    throw new Error('Additional sign-in is required. Tap Analyze again after returning.');
  }
}

// ---------- Image capture + crop ----------
let img = new Image();
let imgLoaded = false;
let selection = null; // {x,y,w,h} in canvas pixel coords (natural image resolution)
let dragStart = null;
let scaleFactor = 1; // displayed size / natural size
let rotationDegrees = 0;
let currentAnalysis = null;
let currentExerciseDraftKey = null;

els.btnPickImage.addEventListener('click', () => els.fileInput.click());
els.btnRetake.addEventListener('click', () => els.fileInput.click());

els.fileInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (ev) => {
    img = new Image();
    img.onload = () => {
      imgLoaded = true;
      rotationDegrees = 0;
      els.captureEmpty.classList.add('hidden');
      els.captureEditor.classList.remove('hidden');
      setupCanvas();
      els.resultArea.classList.add('hidden');
      els.resultArea.innerHTML = '';
    };
    img.src = ev.target.result;
  };
  reader.readAsDataURL(file);
  els.fileInput.value = '';
});

function setupCanvas() {
  const maxW = els.cropStage.clientWidth || 360;
  const sourceW = img.naturalWidth;
  const sourceH = img.naturalHeight;
  const swapsDimensions = rotationDegrees % 180 !== 0;
  const canvasW = swapsDimensions ? sourceH : sourceW;
  const canvasH = swapsDimensions ? sourceW : sourceH;
  scaleFactor = Math.min(1, maxW / canvasW);
  const dispW = Math.round(canvasW * scaleFactor);
  const dispH = Math.round(canvasH * scaleFactor);

  [els.photoCanvas, els.selectionCanvas].forEach((c) => {
    c.width = canvasW;
    c.height = canvasH;
    c.style.width = dispW + 'px';
    c.style.height = dispH + 'px';
  });
  els.cropStage.style.height = dispH + 'px';

  const ctx = els.photoCanvas.getContext('2d');
  ctx.clearRect(0, 0, canvasW, canvasH);
  ctx.save();
  ctx.translate(canvasW / 2, canvasH / 2);
  ctx.rotate(rotationDegrees * Math.PI / 180);
  ctx.drawImage(img, -sourceW / 2, -sourceH / 2, sourceW, sourceH);
  ctx.restore();

  selection = null;
  drawSelectionOverlay();
  els.btnAnalyze.disabled = true;
}

function rotateImage(deltaDegrees) {
  if (!imgLoaded) return;
  rotationDegrees = (rotationDegrees + deltaDegrees + 360) % 360;
  setupCanvas();
  toast(deltaDegrees < 0 ? 'Rotated left' : 'Rotated right', 1200);
}

els.btnRotateLeft.addEventListener('click', () => rotateImage(-90));
els.btnRotateRight.addEventListener('click', () => rotateImage(90));

function canvasPosFromEvent(e) {
  const rect = els.selectionCanvas.getBoundingClientRect();
  const clientX = e.touches ? e.touches[0].clientX : e.clientX;
  const clientY = e.touches ? e.touches[0].clientY : e.clientY;
  const xDisp = clientX - rect.left;
  const yDisp = clientY - rect.top;
  // convert displayed coords to natural-resolution canvas coords
  const x = xDisp * (els.selectionCanvas.width / rect.width);
  const y = yDisp * (els.selectionCanvas.height / rect.height);
  return { x, y };
}

function startDrag(e) {
  if (!imgLoaded) return;
  e.preventDefault();
  dragStart = canvasPosFromEvent(e);
  selection = { x: dragStart.x, y: dragStart.y, w: 0, h: 0 };
}

function moveDrag(e) {
  if (!dragStart) return;
  e.preventDefault();
  const pos = canvasPosFromEvent(e);
  selection = {
    x: Math.min(dragStart.x, pos.x),
    y: Math.min(dragStart.y, pos.y),
    w: Math.abs(pos.x - dragStart.x),
    h: Math.abs(pos.y - dragStart.y),
  };
  drawSelectionOverlay();
}

function endDrag() {
  dragStart = null;
  els.btnAnalyze.disabled = !(selection && selection.w > 10 && selection.h > 10);
}

els.selectionCanvas.addEventListener('mousedown', startDrag);
els.selectionCanvas.addEventListener('mousemove', moveDrag);
window.addEventListener('mouseup', endDrag);
els.selectionCanvas.addEventListener('touchstart', startDrag, { passive: false });
els.selectionCanvas.addEventListener('touchmove', moveDrag, { passive: false });
els.selectionCanvas.addEventListener('touchend', endDrag);

function drawSelectionOverlay() {
  const ctx = els.selectionCanvas.getContext('2d');
  ctx.clearRect(0, 0, els.selectionCanvas.width, els.selectionCanvas.height);
  if (!selection) return;
  ctx.fillStyle = 'rgba(0,0,0,0.45)';
  ctx.fillRect(0, 0, els.selectionCanvas.width, els.selectionCanvas.height);
  ctx.clearRect(selection.x, selection.y, selection.w, selection.h);
  ctx.strokeStyle = '#e0544e';
  ctx.lineWidth = Math.max(2, 3 / scaleFactor);
  ctx.strokeRect(selection.x, selection.y, selection.w, selection.h);
}

els.btnResetCrop.addEventListener('click', () => {
  selection = null;
  drawSelectionOverlay();
  els.btnAnalyze.disabled = true;
});

function getCroppedDataUrl() {
  const w = Math.round(selection.w);
  const h = Math.round(selection.h);
  const out = document.createElement('canvas');
  out.width = w;
  out.height = h;
  out.getContext('2d').drawImage(
    els.photoCanvas,
    Math.round(selection.x), Math.round(selection.y), w, h,
    0, 0, w, h
  );
  return out.toDataURL('image/jpeg', 0.92);
}

// ---------- Analysis ----------
const LEVEL_LABELS = {
  N5: 'JLPT N5 (absolute beginner)',
  N4: 'JLPT N4 (beginner)',
  N3: 'JLPT N3 (intermediate)',
  N2: 'JLPT N2 (upper intermediate)',
  N1: 'JLPT N1 (advanced)',
};

function getAnalysisResponseFormat() {
  const stringProperty = { type: 'string' };
  return {
    type: 'json_schema',
    json_schema: {
      name: 'japanese_text_analysis',
      strict: true,
      schema: {
        type: 'object',
        additionalProperties: false,
        required: ['contentType', 'exercise', 'sourceText', 'sourceWithFurigana', 'translation', 'words', 'kanjiBreakdown', 'grammar'],
        properties: {
          contentType: { type: 'string', enum: ['reading', 'exercise'] },
          exercise: {
            type: 'object',
            additionalProperties: false,
            required: ['instructions', 'exerciseType', 'detectionConfidence', 'uncertainty', 'questions'],
            properties: {
              instructions: stringProperty,
              exerciseType: {
                type: 'string',
                enum: ['none', 'multiple_choice', 'fill_blank', 'conjugation', 'reorder', 'translation', 'reading_comprehension', 'free_text', 'mixed'],
              },
              detectionConfidence: { type: 'number' },
              uncertainty: stringProperty,
              questions: {
                type: 'array',
                items: {
                  type: 'object',
                  additionalProperties: false,
                  required: ['id', 'prompt', 'answerType', 'choices', 'context', 'confidence'],
                  properties: {
                    id: stringProperty,
                    prompt: stringProperty,
                    answerType: { type: 'string', enum: ['multiple_choice', 'text'] },
                    choices: { type: 'array', items: stringProperty },
                    context: stringProperty,
                    confidence: { type: 'number' },
                  },
                },
              },
            },
          },
          sourceText: stringProperty,
          sourceWithFurigana: stringProperty,
          translation: stringProperty,
          words: {
            type: 'array',
            items: {
              type: 'object',
              additionalProperties: false,
              required: ['word', 'reading', 'romaji', 'meaning', 'partOfSpeech'],
              properties: {
                word: stringProperty,
                reading: stringProperty,
                romaji: stringProperty,
                meaning: stringProperty,
                partOfSpeech: stringProperty,
              },
            },
          },
          kanjiBreakdown: {
            type: 'array',
            items: {
              type: 'object',
              additionalProperties: false,
              required: ['word', 'kanji'],
              properties: {
                word: stringProperty,
                kanji: {
                  type: 'array',
                  items: {
                    type: 'object',
                    additionalProperties: false,
                    required: ['char', 'onyomi', 'kunyomi', 'meaning', 'note'],
                    properties: {
                      char: stringProperty,
                      onyomi: stringProperty,
                      kunyomi: stringProperty,
                      meaning: stringProperty,
                      note: stringProperty,
                    },
                  },
                },
              },
            },
          },
          grammar: {
            type: 'array',
            items: {
              type: 'object',
              additionalProperties: false,
              required: ['pattern', 'excerpt', 'explanation'],
              properties: {
                pattern: stringProperty,
                excerpt: stringProperty,
                explanation: stringProperty,
              },
            },
          },
        },
      },
    },
  };
}

function getCorrectionResponseFormat() {
  const stringProperty = { type: 'string' };
  return {
    type: 'json_schema',
    json_schema: {
      name: 'japanese_exercise_correction',
      strict: true,
      schema: {
        type: 'object',
        additionalProperties: false,
        required: ['summary', 'results'],
        properties: {
          summary: stringProperty,
          results: {
            type: 'array',
            items: {
              type: 'object',
              additionalProperties: false,
              required: ['id', 'status', 'correctedAnswer', 'explanation', 'alternatives', 'confidence'],
              properties: {
                id: stringProperty,
                status: {
                  type: 'string',
                  enum: ['correct', 'partially_correct', 'incorrect', 'cannot_determine'],
                },
                correctedAnswer: stringProperty,
                explanation: stringProperty,
                alternatives: { type: 'array', items: stringProperty },
                confidence: { type: 'number' },
              },
            },
          },
        },
      },
    },
  };
}

function buildPrompt(level, requestedMode) {
  const levelLabel = LEVEL_LABELS[level] || level;
  const modeInstruction = {
    auto: 'Determine whether the image is ordinary reading text or a textbook exercise.',
    reading: 'Treat the image as ordinary reading text, even if it resembles an exercise.',
    exercise: 'Treat the image as a textbook exercise and extract its questions.',
  }[requestedMode] || 'Determine whether the image is ordinary reading text or a textbook exercise.';
  return `You are a Japanese language tutor analyzing a photo of Japanese text for a learner at ${levelLabel}.
${modeInstruction}
Respond with ONLY a single valid JSON object (no markdown fences, no extra commentary) matching exactly this shape:
{
  "contentType": "reading or exercise",
  "exercise": {
    "instructions": "exercise instructions, or an empty string for reading text",
    "exerciseType": "none, multiple_choice, fill_blank, conjugation, reorder, translation, reading_comprehension, free_text, or mixed",
    "detectionConfidence": 0.0,
    "uncertainty": "what is unclear about the exercise extraction, or an empty string",
    "questions": [
      {
        "id": "stable short ID such as q1",
        "prompt": "question text exactly enough for the learner to answer",
        "answerType": "multiple_choice or text",
        "choices": ["choice text when present"],
        "context": "passage or local context required to answer, or an empty string",
        "confidence": 0.0
      }
    ]
  },
  "sourceText": "the Japanese text exactly as read from the image, preserving its original line breaks with \\n",
  "sourceWithFurigana": "the same text and line breaks, annotating kanji words as {surface|hiragana reading}, for example {日本語|にほんご}",
  "translation": "a natural, fluent English translation preserving the same line-by-line structure with \\n",
  "words": [
    {"word": "...", "reading": "... (hiragana reading)", "romaji": "...", "meaning": "short English meaning", "partOfSpeech": "..."}
  ],
  "kanjiBreakdown": [
    {"word": "a word from the text made of two or more kanji", "kanji": [
      {"char": "single kanji character", "onyomi": "...", "kunyomi": "...", "meaning": "core meaning(s)", "note": "how this kanji contributes to the word's meaning"}
    ]}
  ],
  "grammar": [
    {"pattern": "grammar point / structure name", "excerpt": "the exact phrase from the text showing it", "explanation": "clear explanation of the rule and why it's used here"}
  ]
}

Rules for tailoring to the learner's level (${levelLabel}):
- For ordinary reading text, set "contentType" to "reading", "exerciseType" to "none", and "questions" to an empty array.
- For a textbook exercise, set "contentType" to "exercise" and extract every visible question in reading order. Never include, infer, or reveal correct answers in the extraction response.
- Use "multiple_choice" only when visible choices exist; otherwise use "text". Reorder, fill-in, conjugation, translation, comprehension, and free-response questions all use a text answer field in this first version.
- If classification or question extraction is uncertain, explain that briefly in "uncertainty" and lower the relevant confidence. Do not invent missing text.
- Preserve the visible line breaks from the image in "sourceText". Encode line breaks as \\n inside the JSON string.
- In "sourceWithFurigana", reproduce the exact same text and line breaks as "sourceText", adding readings only with {kanji-containing surface text|hiragana reading}. Do not change or omit any source characters outside those annotations.
- Preserve corresponding line breaks in "translation", translating each source line in the same order so the two blocks are easy to compare.
- "words": only include words that are genuinely useful/important to learn for someone at this level. Skip words that are trivially basic for this level (e.g. for N2/N1 learners, skip elementary particles or very common N5 vocabulary already assumed known). For N5 learners, include most content words since everything is new.
- "kanjiBreakdown": only include entries for words composed of two or more kanji characters, decomposed into their individual kanji. Every "kanjiBreakdown" word must also appear exactly in "words". Skip this decomposition for kanji that would already be well known at the learner's level (e.g. do not decompose extremely common kanji for N1 learners); focus on kanji at or above their current level.
- "grammar": only explain grammar points that are at or above the learner's current level (i.e. things they likely do NOT already know). Do not explain grammar that is more basic than their level.
- If the text is very short or simple, it is fine for "words", "kanjiBreakdown", or "grammar" to be empty arrays.
- All explanations should be written in English.
- Output strictly valid JSON, with no trailing commas.`;
}

function setLoading(isLoading) {
  if (isLoading) {
    els.resultArea.classList.remove('hidden');
    els.resultArea.innerHTML = '<div class="loading">Analyzing text… ⏳</div>';
  }
}

function showError(err) {
  els.resultArea.classList.remove('hidden');
  els.resultArea.innerHTML = `<div class="error-box">Analysis failed:\n${(err && err.message) || err}</div>`;
}

function parseModelJson(raw) {
  let text = raw.trim();
  // strip markdown code fences if present
  text = text.replace(/^```(?:json)?/i, '').replace(/```$/, '').trim();
  try {
    return JSON.parse(text);
  } catch (err) {
    throw new Error(
      `Azure returned a complete response that was not valid JSON (${text.length} characters). `
      + `Please retry the analysis; if it happens repeatedly, select a smaller text region. `
      + `Technical detail: ${err.message}`
    );
  }
}

function parseAnalysisResponse(data) {
  const choice = data.choices?.[0];
  if (!choice) {
    throw new Error('Azure returned no analysis result. Please retry.');
  }
  if (choice.finish_reason === 'length') {
    throw new Error(
      'The analysis was too long and Azure cut it off before completion. '
      + 'Select a smaller text region or choose a higher JLPT level so fewer basic words and grammar points are explained.'
    );
  }
  if (choice.finish_reason === 'content_filter') {
    throw new Error('Azure content filtering stopped the analysis. Try a smaller region containing only the Japanese exercise or passage.');
  }
  if (choice.finish_reason && choice.finish_reason !== 'stop') {
    throw new Error(`Azure stopped the analysis unexpectedly (${choice.finish_reason}). Please retry.`);
  }
  if (choice.message?.refusal) {
    throw new Error(`Azure declined to analyze this image: ${choice.message.refusal}`);
  }
  const content = choice.message?.content;
  if (!content) {
    throw new Error('Azure returned an empty analysis. Please retry.');
  }
  return parseModelJson(content);
}

function parseCorrectionResponse(data) {
  const choice = data.choices?.[0];
  if (!choice) throw new Error('Azure returned no correction result. Please retry.');
  if (choice.finish_reason === 'length') {
    throw new Error('The correction was too long and Azure cut it off. Try submitting fewer questions at once.');
  }
  if (choice.finish_reason === 'content_filter') {
    throw new Error('Azure content filtering stopped the correction request.');
  }
  if (choice.finish_reason && choice.finish_reason !== 'stop') {
    throw new Error(`Azure stopped the correction unexpectedly (${choice.finish_reason}). Please retry.`);
  }
  if (choice.message?.refusal) {
    throw new Error(`Azure declined to correct this exercise: ${choice.message.refusal}`);
  }
  if (!choice.message?.content) throw new Error('Azure returned an empty correction. Please retry.');
  return parseModelJson(choice.message.content);
}

async function analyze() {
  if (!settings.endpoint || !settings.deployment) {
    toast('Open Settings and configure the Azure OpenAI endpoint/deployment first');
    openSettingsModal();
    return;
  }
  setLoading(true);
  try {
    const token = await getAccessToken();
    const dataUrl = getCroppedDataUrl();
    const level = els.levelSelect.value;
    const prompt = buildPrompt(level, els.analysisModeSelect.value);

    const url = `${settings.endpoint}/openai/deployments/${encodeURIComponent(settings.deployment)}/chat/completions?api-version=${encodeURIComponent(settings.apiVersion)}`;
    const body = {
      messages: [
        {
          role: 'user',
          content: [
            { type: 'text', text: prompt },
            { type: 'image_url', image_url: { url: dataUrl } },
          ],
        },
      ],
      max_completion_tokens: 6000,
      response_format: getAnalysisResponseFormat(),
    };

    const resp = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${token}`,
      },
      body: JSON.stringify(body),
    });

    if (!resp.ok) {
      const errText = await resp.text();
      throw new Error(`HTTP ${resp.status}: ${errText}`);
    }
    const data = await resp.json();
    const parsed = parseAnalysisResponse(data);
    renderResult(parsed);
  } catch (err) {
    console.error(err);
    showError(err);
  }
}

els.btnAnalyze.addEventListener('click', analyze);

function esc(s) {
  const d = document.createElement('div');
  d.textContent = s == null ? '' : String(s);
  return d.innerHTML;
}

function renderFurigana(text) {
  const source = String(text || '');
  const annotation = /\{([^{}|\n]+)\|([^{}|\n]+)\}/g;
  let html = '';
  let lastIndex = 0;
  let match;
  while ((match = annotation.exec(source)) !== null) {
    html += esc(source.slice(lastIndex, match.index));
    html += `<ruby>${esc(match[1])}<rt>${esc(match[2])}</rt></ruby>`;
    lastIndex = annotation.lastIndex;
  }
  return html + esc(source.slice(lastIndex));
}

function renderKanjiDetails(breakdowns) {
  if (!breakdowns.length) return '';
  let html = `<details class="kanji-details"><summary>Kanji breakdown</summary>`;
  for (const kb of breakdowns) {
    html += `<div>`;
    for (const k of kb.kanji || []) {
      html += `<span class="kanji-chip"><span class="kanji-char">${esc(k.char)}</span>`
        + `<span class="kanji-detail">${esc(k.onyomi || '')}${k.kunyomi ? ' / ' + esc(k.kunyomi) : ''} — ${esc(k.meaning || '')}</span></span>`;
    }
    const notes = (kb.kanji || []).filter(k => k.note).map(k => `${esc(k.char)}: ${esc(k.note)}`);
    if (notes.length) html += `<div class="muted">${notes.join(' · ')}</div>`;
    html += `</div>`;
  }
  return html + `</details>`;
}

function exerciseDraftKey(analysis) {
  let hash = 2166136261;
  const text = `${analysis.sourceText}|${analysis.exercise.questions.map((q) => q.id).join('|')}`;
  for (let i = 0; i < text.length; i += 1) {
    hash ^= text.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return `jta_exercise_draft_${(hash >>> 0).toString(16)}`;
}

function loadExerciseDraft(key) {
  try {
    return JSON.parse(localStorage.getItem(key) || '{}');
  } catch {
    return {};
  }
}

function collectExerciseAnswers() {
  if (!currentAnalysis?.exercise?.questions) return {};
  const answers = {};
  currentAnalysis.exercise.questions.forEach((question, index) => {
    if (question.answerType === 'multiple_choice') {
      const selected = document.querySelector(`input[name="exercise-q-${index}"]:checked`);
      answers[question.id] = selected ? selected.value : '';
    } else {
      answers[question.id] = document.querySelector(`[data-exercise-answer="${index}"]`)?.value.trim() || '';
    }
  });
  return answers;
}

function saveExerciseDraft() {
  if (!currentExerciseDraftKey) return;
  localStorage.setItem(currentExerciseDraftKey, JSON.stringify(collectExerciseAnswers()));
}

function renderExerciseCard(analysis) {
  const exercise = analysis.exercise;
  if (analysis.contentType !== 'exercise' || !exercise) return '';

  currentExerciseDraftKey = exerciseDraftKey(analysis);
  const draft = loadExerciseDraft(currentExerciseDraftKey);
  const confidence = Math.round(Math.max(0, Math.min(1, exercise.detectionConfidence || 0)) * 100);
  let html = `<div class="result-card exercise-card"><h3>Exercise</h3>`;
  if (exercise.instructions) html += `<div class="jp-text text-block">${esc(exercise.instructions)}</div>`;
  html += `<div class="exercise-meta">${esc(exercise.exerciseType.replaceAll('_', ' '))} · detection confidence ${confidence}%`;
  if (exercise.uncertainty) html += `<br />Uncertainty: ${esc(exercise.uncertainty)}`;
  html += `</div><div id="exercise-feedback-summary"></div>`;
  if (!exercise.questions.length) {
    return html + `<div class="error-box">This looks like an exercise, but no answerable questions could be extracted. Try selecting a tighter region or choose Textbook exercise mode.</div></div>`;
  }

  exercise.questions.forEach((question, index) => {
    const savedAnswer = draft[question.id] || '';
    html += `<div class="exercise-question">`;
    html += `<div class="exercise-prompt"><strong>${index + 1}.</strong> ${esc(question.prompt)}</div>`;
    if (question.context) html += `<div class="muted text-block">${esc(question.context)}</div>`;
    if (question.answerType === 'multiple_choice' && question.choices.length) {
      question.choices.forEach((choice, choiceIndex) => {
        const checked = savedAnswer === choice ? ' checked' : '';
        html += `<label class="exercise-choice"><input type="radio" name="exercise-q-${index}" value="${choiceIndex}" data-choice-index="${choiceIndex}"${checked} /> <span>${esc(choice)}</span></label>`;
      });
    } else {
      html += `<textarea class="exercise-answer" data-exercise-answer="${index}" placeholder="Enter your answer">${esc(savedAnswer)}</textarea>`;
    }
    html += `<div class="exercise-feedback-slot" id="exercise-feedback-${index}"></div></div>`;
  });
  html += `<div class="exercise-actions"><button id="btn-submit-exercise" class="primary-btn">Check answers</button></div></div>`;
  return html;
}

function wireExerciseControls() {
  const exercise = currentAnalysis?.exercise;
  if (currentAnalysis?.contentType !== 'exercise' || !exercise?.questions?.length) return;
  exercise.questions.forEach((question, index) => {
    if (question.answerType === 'multiple_choice') {
      document.querySelectorAll(`input[name="exercise-q-${index}"]`).forEach((input) => {
        const choiceIndex = Number(input.dataset.choiceIndex);
        input.value = question.choices[choiceIndex] || '';
        input.addEventListener('change', saveExerciseDraft);
      });
    } else {
      document.querySelector(`[data-exercise-answer="${index}"]`)?.addEventListener('input', saveExerciseDraft);
    }
  });
  document.getElementById('btn-submit-exercise')?.addEventListener('click', gradeExercise);
}

function buildCorrectionPrompt(analysis, answers, level) {
  const exerciseData = {
    sourceText: analysis.sourceText,
    instructions: analysis.exercise.instructions,
    exerciseType: analysis.exercise.exerciseType,
    questions: analysis.exercise.questions.map((question) => ({
      id: question.id,
      prompt: question.prompt,
      context: question.context,
      choices: question.choices,
      learnerAnswer: answers[question.id] || '',
    })),
  };
  return `You are a careful Japanese tutor correcting a learner at ${LEVEL_LABELS[level] || level}.
Assess the learner's submitted answers to the exercise data below.

Rules:
- Do not require exact string matching when multiple Japanese answers are valid.
- For translation or composition, use a rubric based on meaning, grammar, naturalness, and whether the prompt was fulfilled.
- Use "cannot_determine" if the image did not contain enough information to know the textbook's intended answer.
- Give concise, constructive explanations appropriate for the learner's level.
- Include a corrected or suggested answer for incorrect or partially correct responses.
- List genuinely acceptable alternatives, not merely paraphrases.
- Keep each result ID exactly equal to its question ID.

Exercise data:
${JSON.stringify(exerciseData)}`;
}

function renderExerciseFeedback(feedback) {
  const summary = document.getElementById('exercise-feedback-summary');
  if (summary) {
    summary.innerHTML = `<div class="exercise-feedback"><strong>Overall feedback</strong><div>${esc(feedback.summary)}</div></div>`;
  }
  const byId = new Map((feedback.results || []).map((result) => [result.id, result]));
  currentAnalysis.exercise.questions.forEach((question, index) => {
    const result = byId.get(question.id);
    const slot = document.getElementById(`exercise-feedback-${index}`);
    if (!slot || !result) return;
    const label = result.status.replaceAll('_', ' ');
    const confidence = Math.round(Math.max(0, Math.min(1, result.confidence || 0)) * 100);
    let html = `<div class="exercise-feedback ${esc(result.status)}"><div class="feedback-status">${esc(label)}</div>`;
    if (result.correctedAnswer) html += `<div><strong>Suggested answer:</strong> ${esc(result.correctedAnswer)}</div>`;
    html += `<div>${esc(result.explanation)}</div>`;
    if (result.alternatives?.length) html += `<div class="muted">Also acceptable: ${result.alternatives.map(esc).join(' · ')}</div>`;
    html += `<div class="muted">Confidence: ${confidence}%</div></div>`;
    slot.innerHTML = html;
  });
}

async function gradeExercise() {
  const button = document.getElementById('btn-submit-exercise');
  const answers = collectExerciseAnswers();
  if (!Object.values(answers).some(Boolean)) {
    toast('Enter at least one answer before checking.');
    return;
  }
  saveExerciseDraft();
  button.disabled = true;
  button.textContent = 'Checking...';
  try {
    const token = await getAccessToken();
    const url = `${settings.endpoint}/openai/deployments/${encodeURIComponent(settings.deployment)}/chat/completions?api-version=${encodeURIComponent(settings.apiVersion)}`;
    const body = {
      messages: [{
        role: 'user',
        content: buildCorrectionPrompt(currentAnalysis, answers, els.levelSelect.value),
      }],
      max_completion_tokens: 3000,
      response_format: getCorrectionResponseFormat(),
    };
    const resp = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${token}`,
      },
      body: JSON.stringify(body),
    });
    if (!resp.ok) {
      const errText = await resp.text();
      throw new Error(`HTTP ${resp.status}: ${errText}`);
    }
    renderExerciseFeedback(parseCorrectionResponse(await resp.json()));
  } catch (err) {
    console.error(err);
    const summary = document.getElementById('exercise-feedback-summary');
    if (summary) summary.innerHTML = `<div class="error-box">Correction failed:\n${esc(err.message || err)}</div>`;
  } finally {
    button.disabled = false;
    button.textContent = 'Check answers';
  }
}

function renderResult(r) {
  currentAnalysis = r;
  currentExerciseDraftKey = null;
  const showFurigana = settings.furigana;
  const showRomaji = settings.romaji;
  let html = '';

  if (r.sourceText) {
    const hasFurigana = !!r.sourceWithFurigana;
    html += `<div class="result-card"><div class="result-card-header"><h3>Recognized text</h3>`;
    if (hasFurigana) {
      html += `<label class="furigana-toggle"><input type="checkbox" id="result-furigana-toggle"${showFurigana ? ' checked' : ''} /> Furigana</label>`;
    }
    html += `</div><div id="source-plain" class="jp-text text-block${hasFurigana && showFurigana ? ' hidden' : ''}">${esc(r.sourceText)}</div>`;
    if (hasFurigana) {
      html += `<div id="source-furigana" class="jp-text text-block${showFurigana ? '' : ' hidden'}">${renderFurigana(r.sourceWithFurigana)}</div>`;
    }
    html += `</div>`;
  }

  html += renderExerciseCard(r);

  html += `<div class="result-card"><h3>Translation</h3><div class="text-block">${esc(r.translation || '—')}</div></div>`;

  if (r.words && r.words.length) {
    const breakdownByWord = new Map();
    for (const kb of r.kanjiBreakdown || []) {
      const entries = breakdownByWord.get(kb.word) || [];
      entries.push(kb);
      breakdownByWord.set(kb.word, entries);
    }
    html += `<div class="result-card"><h3>Important words</h3><ul class="word-list">`;
    for (const w of r.words) {
      html += `<li><span class="word-main">${esc(w.word)}</span>`;
      if (showFurigana && w.reading) html += `<span class="word-reading">${esc(w.reading)}</span>`;
      if (showRomaji && w.romaji) html += `<span class="word-reading">[${esc(w.romaji)}]</span>`;
      html += `<div class="word-meaning">${esc(w.meaning || '')}${w.partOfSpeech ? ` <span class="muted">(${esc(w.partOfSpeech)})</span>` : ''}</div>`;
      html += renderKanjiDetails(breakdownByWord.get(w.word) || []);
      html += `</li>`;
    }
    html += `</ul></div>`;
  }

  if (r.grammar && r.grammar.length) {
    html += `<div class="result-card"><h3>Grammar</h3>`;
    for (const g of r.grammar) {
      html += `<div class="grammar-item"><div class="pattern">${esc(g.pattern)}</div>`;
        if (g.excerpt) html += `<div class="jp-text">${esc(g.excerpt)}</div>`;
      html += `<div>${esc(g.explanation || '')}</div></div>`;
    }
    html += `</div>`;
  }

  els.resultArea.innerHTML = html;
  els.resultArea.classList.remove('hidden');

  const furiganaToggle = document.getElementById('result-furigana-toggle');
  if (furiganaToggle) {
    furiganaToggle.addEventListener('change', () => {
      document.getElementById('source-plain').classList.toggle('hidden', furiganaToggle.checked);
      document.getElementById('source-furigana').classList.toggle('hidden', !furiganaToggle.checked);
    });
  }
  wireExerciseControls();
}

// ---------- init ----------
els.levelSelect.value = localStorage.getItem('jta_level') || 'N3';
els.levelSelect.addEventListener('change', () => localStorage.setItem('jta_level', els.levelSelect.value));
els.analysisModeSelect.value = localStorage.getItem('jta_analysis_mode') || 'auto';
els.analysisModeSelect.addEventListener('change', () => localStorage.setItem('jta_analysis_mode', els.analysisModeSelect.value));

initMsal();

if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  });
}

window.addEventListener('resize', () => { if (imgLoaded) setupCanvas(); });
