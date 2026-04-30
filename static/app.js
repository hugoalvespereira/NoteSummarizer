const MAX_UPLOAD_FILES = 5;
const MAX_UPLOAD_MB = 80;
const THEME_STORAGE_KEY = "notes-summarizer-theme";

const state = {
  sessionId: null,
  filename: null,
  presentations: [],
  providers: [],
  providerStatuses: {},
  provider: "codex",
  allowSavedApiKeys: true,
  themeOverride: storedThemeOverride(),
  resultComparison: [],
  resultExpanded: false,
  codexDetailsExpanded: false,
  promptExpanded: false,
  summarizeRunId: 0,
  aiWaitTimers: {},
};

const els = {
  status: document.querySelector("#status-pill"),
  shell: document.querySelector("#app-shell"),
  themeToggle: document.querySelector("#theme-toggle"),
  themeToggleIcon: document.querySelector("#theme-toggle-icon"),
  themeToggleLabel: document.querySelector("#theme-toggle-label"),
  toast: document.querySelector("#toast"),
  uploadView: document.querySelector("#upload-view"),
  setupView: document.querySelector("#setup-view"),
  resultView: document.querySelector("#result-view"),
  dropZone: document.querySelector("#drop-zone"),
  fileInput: document.querySelector("#file-input"),
  presentationLabel: document.querySelector("#presentation-label"),
  deckName: document.querySelector("#deck-name"),
  slideCount: document.querySelector("#slide-count"),
  noteCount: document.querySelector("#note-count"),
  providerGroup: document.querySelector("#provider-group"),
  codexAuthPanel: document.querySelector("#codex-auth-panel"),
  codexAuthButton: document.querySelector("#codex-auth-button"),
  codexDetailsButton: document.querySelector("#codex-details-button"),
  codexAuthStatus: document.querySelector("#codex-auth-status"),
  codexAccountDetails: document.querySelector("#codex-account-details"),
  codexAuthCode: document.querySelector("#codex-auth-code"),
  codexAuthLink: document.querySelector("#codex-auth-link"),
  keyField: document.querySelector("#key-field"),
  keyLabel: document.querySelector("#key-label"),
  keyInput: document.querySelector("#key-input"),
  providerSetupCard: document.querySelector("#provider-setup-card"),
  providerSetupTitle: document.querySelector("#provider-setup-title"),
  providerSetupText: document.querySelector("#provider-setup-text"),
  providerKeyStatus: document.querySelector("#provider-key-status"),
  providerKeyLink: document.querySelector("#provider-key-link"),
  providerGuideLink: document.querySelector("#provider-guide-link"),
  saveKeyButton: document.querySelector("#save-key-button"),
  forgetKeyButton: document.querySelector("#forget-key-button"),
  modelInput: document.querySelector("#model-input"),
  promptPanel: document.querySelector("#prompt-panel"),
  promptEditButton: document.querySelector("#prompt-edit-button"),
  promptInput: document.querySelector("#prompt-input"),
  summaryProgress: document.querySelector("#summary-progress"),
  summaryProgressTitle: document.querySelector("#summary-progress-title"),
  summaryProgressDetail: document.querySelector("#summary-progress-detail"),
  summaryProgressBar: document.querySelector("#summary-progress-bar"),
  cancelButton: document.querySelector("#cancel-button"),
  cancelButtonSecondary: document.querySelector("#cancel-button-secondary"),
  summarizeButton: document.querySelector("#summarize-button"),
  resultHomeButton: document.querySelector("#result-home-button"),
  resultTitle: document.querySelector("#result-title"),
  downloadLink: document.querySelector("#download-link"),
  warnings: document.querySelector("#warnings"),
  comparison: document.querySelector("#comparison"),
  showAllButton: document.querySelector("#show-all-button"),
  showAllLabel: document.querySelector("#show-all-label"),
  scrollCue: document.querySelector("#scroll-cue"),
};

const systemTheme = window.matchMedia("(prefers-color-scheme: dark)");
let codexAuthPoll = null;
let comparisonSizingFrame = null;

function currentSystemTheme() {
  return systemTheme.matches ? "dark" : "light";
}

function activeTheme() {
  return state.themeOverride || currentSystemTheme();
}

function storedThemeOverride() {
  try {
    const stored = window.localStorage.getItem(THEME_STORAGE_KEY);
    return stored === "dark" || stored === "light" ? stored : null;
  } catch {
    return null;
  }
}

function applyTheme(theme) {
  document.documentElement.dataset.theme = theme;
  const nextTheme = theme === "dark" ? "light" : "dark";
  els.themeToggleIcon.textContent = theme === "dark" ? "☀" : "☾";
  els.themeToggle.setAttribute("aria-label", `Switch to ${nextTheme} mode`);
  els.themeToggle.title = `Switch to ${nextTheme} mode`;
  els.themeToggleLabel.textContent = `Switch to ${nextTheme} mode`;
}

function toggleTheme() {
  state.themeOverride = activeTheme() === "dark" ? "light" : "dark";
  try {
    window.localStorage.setItem(THEME_STORAGE_KEY, state.themeOverride);
  } catch {
    // Theme still changes for this page when storage is unavailable.
  }
  applyTheme(state.themeOverride);
}

applyTheme(activeTheme());

systemTheme.addEventListener("change", () => {
  if (!state.themeOverride) {
    applyTheme(currentSystemTheme());
  }
});

function setStatus(label) {
  els.status.textContent = label;
}

function showToast(message) {
  els.toast.textContent = message;
  els.toast.classList.remove("hidden");
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => els.toast.classList.add("hidden"), 5200);
}

function setView(view) {
  els.shell.dataset.view = view;
  els.uploadView.classList.toggle("hidden", view !== "upload");
  els.setupView.classList.toggle("hidden", view !== "setup");
  els.resultView.classList.toggle("hidden", view !== "result");
  if (view === "result") {
    scheduleComparisonSizing();
  }
  scheduleScrollCueUpdate();
}

function clearCodexAuthPoll() {
  if (codexAuthPoll) {
    window.clearInterval(codexAuthPoll);
    codexAuthPoll = null;
  }
}

function startCodexAuthPoll() {
  clearCodexAuthPoll();
  codexAuthPoll = window.setInterval(refreshCodexAuthStatus, 2200);
}

function renderCodexAccountDetails(account = {}) {
  const rows = [];
  const signedIn = [account.name, account.email].filter(Boolean).join(" · ");
  if (signedIn) rows.push(["Signed in", signedIn]);

  if (account.plan) {
    const plan = account.subscriptionActiveUntil
      ? `${account.plan} · active until ${formatDateTime(account.subscriptionActiveUntil)}`
      : account.plan;
    rows.push(["Plan", plan]);
  }

  if (account.organization?.title) {
    const organization = account.organization.role
      ? `${account.organization.title} · ${account.organization.role}`
      : account.organization.title;
    rows.push(["Organization", organization]);
  }

  if (account.authMode) {
    const login = account.authProvider ? `${account.authMode} via ${account.authProvider}` : account.authMode;
    rows.push(["Login", login]);
  }

  if (account.accountId) rows.push(["Account ID", account.accountId]);
  if (account.lastRefresh) rows.push(["Last refreshed", formatDateTime(account.lastRefresh)]);

  els.codexAccountDetails.innerHTML = rows
    .map(([label, value]) => `<div><dt>${escapeHtml(label)}</dt><dd>${escapeHtml(value)}</dd></div>`)
    .join("");
  els.codexAccountDetails.dataset.hasDetails = rows.length ? "true" : "false";
  return rows.length;
}

function formatDateTime(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return value;
  return new Intl.DateTimeFormat(undefined, {
    dateStyle: "medium",
    timeStyle: "short",
  }).format(date);
}

function codexConnectedMessage(account = {}) {
  if (account.name) return `Connected as ${account.name}.`;
  if (account.email) return `Connected as ${account.email}.`;
  if (account.authMode) return `Connected through ${account.authMode}.`;
  return "OpenAI account connected.";
}

function updateCodexAuthUi(data) {
  const status = data.status || "disconnected";
  els.codexAuthPanel.dataset.authStatus = status;
  els.codexAuthPanel.dataset.expanded = "false";
  els.codexAuthCode.classList.add("hidden");
  els.codexAuthLink.classList.add("hidden");
  els.codexAccountDetails.classList.add("hidden");
  els.codexDetailsButton.classList.add("hidden");
  els.codexAuthButton.disabled = false;

  if (status === "connected") {
    clearCodexAuthPoll();
    els.codexAuthStatus.textContent = codexConnectedMessage(data.account);
    const detailCount = renderCodexAccountDetails(data.account);
    els.codexDetailsButton.classList.toggle("hidden", detailCount === 0);
    setCodexDetailsExpanded(state.codexDetailsExpanded && detailCount > 0);
    els.codexAuthButton.textContent = "Reconnect";
    return;
  }

  if (status === "pending") {
    els.codexAuthStatus.textContent = "Enter this one-time code in the OpenAI popup, then return here.";
    els.codexAuthCode.textContent = data.userCode || "";
    els.codexAuthCode.classList.remove("hidden");
    if (data.authUrl) {
      els.codexAuthLink.href = data.authUrl;
      els.codexAuthLink.classList.remove("hidden");
    }
    els.codexAuthButton.textContent = "Open popup";
    startCodexAuthPoll();
    return;
  }

  if (status === "error") {
    clearCodexAuthPoll();
    els.codexAuthStatus.textContent = data.error || "OpenAI account connection failed.";
    els.codexAuthButton.textContent = "Try again";
    return;
  }

  clearCodexAuthPoll();
  els.codexAuthStatus.textContent = "Connect with OpenAI login to summarize using your account.";
  els.codexAuthButton.textContent = "Connect";
}

function setCodexDetailsExpanded(expanded) {
  const canExpand = els.codexAccountDetails.dataset.hasDetails === "true";
  state.codexDetailsExpanded = Boolean(expanded && canExpand);
  els.codexAuthPanel.dataset.expanded = state.codexDetailsExpanded ? "true" : "false";
  els.codexAccountDetails.classList.toggle("hidden", !state.codexDetailsExpanded);
  els.codexDetailsButton.textContent = state.codexDetailsExpanded ? "Hide" : "Details";
  els.codexDetailsButton.setAttribute("aria-expanded", String(state.codexDetailsExpanded));
  scheduleScrollCueUpdate();
}

function toggleCodexDetails() {
  setCodexDetailsExpanded(!state.codexDetailsExpanded);
}

async function refreshCodexAuthStatus() {
  if (state.provider !== "codex") return;
  try {
    const response = await fetch("/api/codex-oauth/status");
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Could not check OpenAI login status.");
    updateCodexAuthUi(data);
  } catch (error) {
    updateCodexAuthUi({ status: "error", error: error.message });
  }
}

function writeOAuthPopup(popup, data) {
  if (!popup) return;
  if (data.status === "connected") {
    const account = data.account || {};
    const accountLine = account.name || account.email || account.authMode || "OpenAI account";
    const planLine = account.plan ? `${account.plan} plan` : "Connected through OpenAI login";
    popup.document.open();
    popup.document.write(`<!doctype html>
      <html>
        <head>
          <title>OpenAI OAuth</title>
          <meta name="viewport" content="width=device-width, initial-scale=1" />
          <style>
            :root { color-scheme: light dark; }
            body {
              min-height: 100vh;
              margin: 0;
              display: grid;
              place-items: center;
              background: Canvas;
              color: CanvasText;
              font: 16px system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
            }
            main {
              width: min(420px, calc(100% - 40px));
              display: grid;
              gap: 12px;
              text-align: center;
            }
            h1 { margin: 0; font-size: 28px; }
            p { margin: 0; color: color-mix(in srgb, CanvasText 72%, transparent); line-height: 1.45; }
          </style>
        </head>
        <body>
          <main>
            <h1>OpenAI account connected</h1>
            <p>${escapeHtml(accountLine)}</p>
            <p>${escapeHtml(planLine)}</p>
          </main>
        </body>
      </html>`);
    popup.document.close();
    window.setTimeout(() => popup.close(), 1200);
    return;
  }

  const authUrl = data.authUrl || "https://auth.openai.com/codex/device";
  popup.location.href = authUrl;
}

async function startCodexOAuth() {
  const popup = window.open("", "openai-codex-oauth", "width=520,height=680");
  if (popup) {
    popup.document.write("<p style=\"font:16px system-ui;padding:24px\">Starting OpenAI sign in...</p>");
  }

  els.codexAuthButton.disabled = true;
  els.codexAuthStatus.textContent = "Starting OpenAI account sign in...";
  try {
    const response = await fetch("/api/codex-oauth/start", { method: "POST" });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Could not start OpenAI login.");
    writeOAuthPopup(popup, data);
    updateCodexAuthUi(data);
  } catch (error) {
    if (popup) popup.close();
    updateCodexAuthUi({ status: "error", error: error.message });
  }
}

function setBusy(isBusy, label = "Working") {
  els.summarizeButton.disabled = isBusy;
  els.cancelButton.disabled = isBusy;
  els.cancelButtonSecondary.disabled = isBusy;
  els.shell.setAttribute("aria-busy", isBusy ? "true" : "false");
  if (isBusy) setStatus(label);
}

function setSummarizeBusy(isBusy) {
  els.summarizeButton.classList.toggle("is-loading", isBusy);
  els.summarizeButton.textContent = isBusy ? "Summarizing" : "Summarize";
  els.codexAuthButton.disabled = isBusy;
  els.codexDetailsButton.disabled = isBusy;
  els.keyInput.disabled = isBusy;
  els.saveKeyButton.disabled = isBusy || !els.keyInput.value.trim();
  els.forgetKeyButton.disabled = isBusy;
  els.modelInput.disabled = isBusy;
  els.promptEditButton.disabled = isBusy;
  els.promptInput.disabled = isBusy;
  els.providerGroup.querySelectorAll("input").forEach((input) => {
    input.disabled = isBusy;
  });
}

async function analyzeFiles(files) {
  const selected = Array.from(files || []).filter(Boolean);
  if (!selected.length) return;
  if (selected.length > MAX_UPLOAD_FILES) {
    showToast(`Upload up to ${MAX_UPLOAD_FILES} .pptx files at once.`);
    els.fileInput.value = "";
    return;
  }
  const unsupported = selected.find((file) => !isPowerPointFile(file));
  if (unsupported) {
    showToast("Only .pptx files are supported.");
    els.fileInput.value = "";
    return;
  }

  const form = new FormData();
  selected.forEach((file) => form.append("file", file));

  setBusy(true, "Reading");
  try {
    const response = await fetch("/api/analyze", { method: "POST", body: form });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Upload failed.");
    state.sessionId = data.sessionId;
    state.filename = data.filename;
    state.presentations = normalizePresentations(data);
    state.allowSavedApiKeys = data.allowSavedApiKeys ?? true;
    renderSetup(data);
    setView("setup");
    setStatus("Loaded");
  } catch (error) {
    showToast(error.message);
    setStatus("Ready");
  } finally {
    setBusy(false);
    els.fileInput.value = "";
  }
}

function isPowerPointFile(file) {
  return /\.pptx$/i.test(file.name || "");
}

function normalizePresentations(data) {
  if (Array.isArray(data.presentations) && data.presentations.length) {
    return data.presentations;
  }
  return [
    {
      id: "1",
      filename: data.filename,
      slideCount: data.slideCount,
      noteCount: data.noteCount,
      noteSlideCount: data.noteSlideCount,
      writableNoteCount: data.writableNoteCount,
    },
  ];
}

function renderSetup(data) {
  const presentations = state.presentations.length ? state.presentations : normalizePresentations(data);
  const presentationCount = presentations.length;
  const slideCount = sumPresentations(presentations, "slideCount");
  const noteCount = sumPresentations(presentations, "noteCount");
  els.deckName.textContent = presentationCount === 1 ? presentations[0].filename : `${presentationCount} presentations selected`;
  els.presentationLabel.textContent = presentationCount === 1 ? "Presentation" : "Presentations";
  els.slideCount.textContent = slideCount;
  els.noteCount.textContent = noteCount;
  els.promptInput.value = data.prompt;
  setPromptExpanded(false);
  state.codexDetailsExpanded = false;
  state.providers = data.providers || [];
  state.providerStatuses = providerStatusesFromList(state.providers);
  state.provider = data.provider || "codex";
  renderProviders();
  selectProvider(state.provider);
}

function providerStatusesFromList(providers) {
  return Object.fromEntries((providers || []).map((provider) => [provider.id, provider.keyStatus || {}]));
}

function sumPresentations(presentations, key) {
  return presentations.reduce((total, presentation) => total + Number(presentation[key] || 0), 0);
}

function renderProviders() {
  els.providerGroup.innerHTML = "";
  state.providers.forEach((provider) => {
    const label = document.createElement("label");
    const speedLabel = provider.speedLabel || (provider.id === "codex" ? "slower" : "faster");
    const speedClass = `is-${speedLabel.toLowerCase().replaceAll(/[^a-z0-9]+/g, "-")}`;
    const keyStatus = provider.keyStatus || state.providerStatuses[provider.id] || {};
    const keyBadge = provider.requiresKey
      ? `<small class="provider-config ${keyStatus.configured ? "is-configured" : "is-missing"}">${escapeHtml(keyStatus.configured ? keyStatus.source : "setup")}</small>`
      : "";
    label.innerHTML = `
      <input type="radio" name="provider" value="${escapeHtml(provider.id)}" />
      <span class="provider-option">
        <span class="provider-name">${escapeHtml(provider.shortLabel)}</span>
        <small class="provider-speed ${speedClass}">${escapeHtml(speedLabel)}</small>
        ${keyBadge}
      </span>
    `;
    const input = label.querySelector("input");
    input.checked = provider.id === state.provider;
    input.addEventListener("change", () => selectProvider(provider.id));
    els.providerGroup.appendChild(label);
  });
}

function selectProvider(providerId) {
  const provider = state.providers.find((item) => item.id === providerId) || state.providers[0];
  if (!provider) return;
  state.provider = provider.id;
  els.providerGroup.querySelectorAll("input").forEach((input) => {
    input.checked = input.value === provider.id;
  });
  els.keyField.classList.toggle("hidden", !provider.requiresKey);
  els.codexAuthPanel.classList.toggle("hidden", provider.id !== "codex");
  els.keyLabel.textContent = provider.keyLabel || "API key";
  const keyStatus = state.providerStatuses[provider.id] || provider.keyStatus || {};
  els.keyInput.placeholder = keyStatus.configured
    ? "Optional: insert a different API key for this run"
    : "Insert your API key here";
  els.keyInput.value = "";
  els.modelInput.innerHTML = "";
  provider.models.forEach((model) => {
    const option = document.createElement("option");
    const description = provider.modelDescriptions?.[model];
    option.value = model;
    option.textContent = description ? `${model} — ${description}` : model;
    els.modelInput.appendChild(option);
  });
  els.modelInput.value = provider.defaultModel;
  if (provider.id === "codex") {
    refreshCodexAuthStatus();
  } else {
    clearCodexAuthPoll();
  }
  renderProviderSetupCard(provider);
}

function renderProviderSetupCard(provider = selectedProvider()) {
  if (!provider || !provider.requiresKey) {
    els.providerSetupCard.classList.add("hidden");
    return;
  }

  const keyStatus = state.providerStatuses[provider.id] || provider.keyStatus || {};
  const isGemini = provider.id === "gemini";
  els.providerSetupTitle.textContent = isGemini ? "Need a free Gemini key?" : `${provider.shortLabel} setup`;
  els.providerSetupText.textContent = isGemini
    ? "Create one in Google AI Studio, paste it here once, and save it for future summaries."
    : provider.setupSummary || "Paste an API key once and save it for future summaries.";
  els.providerKeyStatus.textContent = keyStatus.label || "Needs setup";
  els.providerKeyStatus.dataset.status = keyStatus.configured ? "configured" : "missing";
  els.providerKeyLink.href = provider.keyUrl || provider.docsUrl || "/provider-setup.html";
  els.providerKeyLink.textContent = isGemini ? "Get Gemini key" : "Get key";
  els.providerGuideLink.href = `/provider-setup.html#${encodeURIComponent(provider.id)}`;

  const canSave = Boolean(state.allowSavedApiKeys && keyStatus.canSave !== false);
  els.saveKeyButton.disabled = !canSave || !els.keyInput.value.trim();
  els.saveKeyButton.classList.toggle("hidden", !canSave);
  els.forgetKeyButton.disabled = !canSave || !keyStatus.saved;
  els.forgetKeyButton.classList.toggle("hidden", !canSave || !keyStatus.saved);
  els.providerSetupCard.classList.remove("hidden");
  scheduleScrollCueUpdate();
}

function selectedProvider() {
  return state.providers.find((item) => item.id === state.provider);
}

async function refreshProviderStatuses() {
  const response = await fetch("/api/provider-status", { cache: "no-store" });
  const data = await response.json();
  if (!response.ok) throw new Error(data.error || "Could not refresh provider status.");
  state.allowSavedApiKeys = data.allowSavedApiKeys ?? state.allowSavedApiKeys;
  state.providerStatuses = data.statuses || providerStatusesFromList(data.providers || []);
  if (Array.isArray(data.providers) && data.providers.length) {
    state.providers = data.providers;
  } else {
    state.providers = state.providers.map((provider) => ({
      ...provider,
      keyStatus: state.providerStatuses[provider.id] || provider.keyStatus,
    }));
  }
  renderProviders();
  selectProvider(state.provider);
}

async function saveProviderKey() {
  const provider = selectedProvider();
  if (!provider || !provider.requiresKey) return;
  const apiKey = els.keyInput.value.trim();
  if (!apiKey) {
    showToast("Paste an API key before saving.");
    return;
  }

  els.saveKeyButton.disabled = true;
  try {
    const response = await fetch("/api/provider-key", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ provider: provider.id, apiKey }),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Could not save API key.");
    state.allowSavedApiKeys = data.allowSavedApiKeys ?? state.allowSavedApiKeys;
    state.providerStatuses = data.statuses || providerStatusesFromList(data.providers || []);
    if (Array.isArray(data.providers)) state.providers = data.providers;
    els.keyInput.value = "";
    renderProviders();
    selectProvider(provider.id);
    showToast(`${provider.shortLabel} key saved on this server.`);
  } catch (error) {
    showToast(error.message);
    renderProviderSetupCard(provider);
  }
}

async function forgetProviderKey() {
  const provider = selectedProvider();
  if (!provider || !provider.requiresKey) return;

  els.forgetKeyButton.disabled = true;
  try {
    const response = await fetch(`/api/provider-key/${encodeURIComponent(provider.id)}`, { method: "DELETE" });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Could not forget API key.");
    state.allowSavedApiKeys = data.allowSavedApiKeys ?? state.allowSavedApiKeys;
    state.providerStatuses = data.statuses || providerStatusesFromList(data.providers || []);
    if (Array.isArray(data.providers)) state.providers = data.providers;
    renderProviders();
    selectProvider(provider.id);
    showToast(`${provider.shortLabel} saved key removed.`);
  } catch (error) {
    showToast(error.message);
    renderProviderSetupCard(provider);
  }
}

async function summarize() {
  if (!state.sessionId) return;
  const runId = state.summarizeRunId + 1;
  state.summarizeRunId = runId;
  const allPresentations = state.presentations.length ? state.presentations : [{ id: "1", filename: state.filename }];
  const presentations = allPresentations.filter((presentation) => Number(presentation.writableNoteCount ?? presentation.noteCount ?? 1) > 0);
  const skippedResults = allPresentations
    .filter((presentation) => Number(presentation.writableNoteCount ?? presentation.noteCount ?? 1) <= 0)
    .map((presentation) => ({
      filename: presentation.filename,
      error: "No writable speaker notes found.",
    }));
  if (!presentations.length) {
    showToast("No writable speaker notes were found in the selected presentations.");
    return;
  }

  setBusy(true, "Summarizing");
  setSummarizeBusy(true);
  showSummaryProgress(presentations.length);
  const results = [...skippedResults];
  try {
    if (state.provider === "codex") {
      const codexResults = await summarizeSequentially(presentations, runId);
      results.push(...codexResults);
    } else {
      const apiResults = await summarizeConcurrently(presentations, runId);
      results.push(...apiResults);
    }

    if (runId !== state.summarizeRunId) return;
    const successes = results.filter((result) => !result.error);
    if (!successes.length) {
      throw new Error(results[0]?.error || "Summarization failed.");
    }
    renderResultBatch(results);
    setView("result");
    setStatus("Done");
  } catch (error) {
    showToast(error.message);
    setStatus("Loaded");
  } finally {
    if (runId === state.summarizeRunId) {
      setSummarizeBusy(false);
      setBusy(false);
      hideSummaryProgress();
    }
  }
}

async function summarizeSequentially(presentations, runId) {
  const results = [];
  for (let index = 0; index < presentations.length; index += 1) {
    if (runId !== state.summarizeRunId) return results;
    const presentation = presentations[index];
    setSummaryProgress(index, presentations.length, `Starting ${presentation.filename}`, `${index} of ${presentations.length} complete`);
    results.push(
      await summarizePresentation(presentation, (job) => {
        if (runId !== state.summarizeRunId) return;
        setPresentationJobProgress(index, presentations.length, presentation, job);
      }),
    );
    setSummaryProgress(index + 1, presentations.length, `Summarized ${presentation.filename}`, `${index + 1} of ${presentations.length} complete`);
  }
  return results;
}

async function summarizeConcurrently(presentations, runId) {
  let completed = 0;
  const progressByPresentation = new Map(presentations.map((presentation) => [presentation.id, 0]));
  setSummaryProgress(0, presentations.length, "Summarizing presentations", `0 of ${presentations.length} complete`);
  const jobs = presentations.map(async (presentation) => {
    const result = await summarizePresentation(presentation, (job) => {
      if (runId !== state.summarizeRunId) return;
      progressByPresentation.set(presentation.id, jobProgressFraction(presentation, job));
      const completedUnits = Array.from(progressByPresentation.values()).reduce((total, value) => total + value, 0);
      const title = job.title ? `${job.title}: ${presentation.filename}` : `Summarizing ${presentation.filename}`;
      const detail = `${Math.floor(completedUnits)} of ${presentations.length} complete · ${job.detail || "Working"}`;
      setSummaryProgress(completedUnits, presentations.length, title, detail);
    });
    completed += 1;
    progressByPresentation.set(presentation.id, 1);
    setSummaryProgress(
      Array.from(progressByPresentation.values()).reduce((total, value) => total + value, 0),
      presentations.length,
      `Summarized ${presentation.filename}`,
      `${completed} of ${presentations.length} complete`,
    );
    return result;
  });
  return Promise.all(jobs);
}

async function summarizePresentation(presentation, onProgress) {
  try {
    const response = await fetch("/api/summarize/start", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(summarizePayload(presentation)),
    });
    const data = await response.json();
    if (!response.ok) {
      return {
        filename: presentation.filename,
        error: data.error || "Summarization failed.",
      };
    }
    if (!data.jobId) {
      throw new Error("Summarization job did not start.");
    }
    return await pollSummarizeJob(data.jobId, presentation, onProgress);
  } catch (error) {
    return {
      filename: presentation.filename,
      error: error.message || "Summarization failed.",
    };
  }
}

function summarizePayload(presentation) {
  return {
    sessionId: state.sessionId,
    presentationId: presentation.id,
    provider: state.provider,
    model: els.modelInput.value.trim(),
    apiKey: els.keyInput.value.trim(),
    prompt: els.promptInput.value.trim(),
  };
}

async function pollSummarizeJob(jobId, presentation, onProgress) {
  while (true) {
    const response = await fetch(`/api/summarize/progress/${encodeURIComponent(jobId)}`, { cache: "no-store" });
    const job = await response.json();
    if (!response.ok) {
      return {
        filename: presentation.filename,
        error: job.error || "Could not read summarization progress.",
      };
    }

    if (onProgress) onProgress(job);
    if (job.status === "complete") {
      return job.result;
    }
    if (job.status === "error") {
      return {
        filename: presentation.filename,
        error: job.error || job.detail || "Summarization failed.",
      };
    }
    await delay(650);
  }
}

function setPresentationJobProgress(index, total, presentation, job) {
  const jobFraction = jobProgressFraction(presentation, job);
  const title = job.title ? `${job.title}: ${presentation.filename}` : `Summarizing ${presentation.filename}`;
  const detail = `${index} of ${total} complete · ${job.detail || "Working"}`;
  setSummaryProgress(index + jobFraction, total, title, detail);
}

function jobProgressFraction(presentation, job) {
  const baseFraction = clamp(Number(job.percent || 0) / 100, 0, 1);
  const timerKey = `${state.summarizeRunId}:${presentation.id}`;
  const isWaitingForAi =
    job.status === "running" &&
    /^Sending to\s+/i.test(job.title || "") &&
    baseFraction >= 0.3 &&
    baseFraction < 0.66;

  if (!isWaitingForAi) {
    delete state.aiWaitTimers[timerKey];
    return baseFraction;
  }

  const now = performance.now();
  if (!state.aiWaitTimers[timerKey]) {
    state.aiWaitTimers[timerKey] = {
      startedAt: now,
      baseFraction,
    };
  }

  const timer = state.aiWaitTimers[timerKey];
  const elapsed = now - timer.startedAt;
  const capFraction = 0.65;
  const easedFraction = capFraction - (capFraction - timer.baseFraction) * Math.exp(-elapsed / 18000);
  return clamp(Math.max(baseFraction, easedFraction), baseFraction, capFraction - 0.002);
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function delay(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

function showSummaryProgress(total) {
  els.summaryProgress.classList.remove("hidden");
  setSummaryProgress(0, total, "Preparing", `0 of ${total} complete`);
}

function setSummaryProgress(completed, total, title, detail) {
  const percent = total > 0 ? Math.round((completed / total) * 100) : 0;
  els.summaryProgressTitle.textContent = title;
  els.summaryProgressDetail.textContent = detail;
  els.summaryProgressBar.style.width = `${percent}%`;
  els.summaryProgressBar.parentElement.setAttribute("aria-label", `${percent}% complete`);
  els.summaryProgressBar.parentElement.setAttribute("aria-valuenow", String(percent));
  scheduleScrollCueUpdate();
}

function hideSummaryProgress() {
  els.summaryProgress.classList.add("hidden");
  els.summaryProgressBar.style.width = "0%";
  state.aiWaitTimers = {};
}

async function cancel() {
  const sessionId = state.sessionId;
  resetAppState();
  if (sessionId) {
    fetch("/api/cancel", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId }),
    });
  }
}

function goHome() {
  resetAppState();
}

function resetAppState() {
  state.summarizeRunId += 1;
  state.sessionId = null;
  state.filename = null;
  state.presentations = [];
  state.providers = [];
  state.providerStatuses = {};
  state.provider = "codex";
  state.resultComparison = [];
  state.resultExpanded = false;
  state.codexDetailsExpanded = false;
  state.aiWaitTimers = {};
  setPromptExpanded(false);
  els.keyInput.value = "";
  els.comparison.innerHTML = "";
  els.showAllButton.classList.add("hidden");
  els.downloadLink.textContent = "Download";
  els.warnings.innerHTML = "";
  els.warnings.classList.add("hidden");
  hideSummaryProgress();
  setView("upload");
  setStatus("Ready");
}

function renderResultBatch(results) {
  const successes = results.filter((result) => !result.error);
  const failures = results.filter((result) => result.error);
  const updatedCount = successes.reduce((total, result) => total + Number(result.updatedCount || 0), 0);
  const isBatch = successes.length > 1 || state.presentations.length > 1;
  const presentationWord = successes.length === 1 ? "presentation" : "presentations";
  els.resultTitle.textContent = isBatch
    ? `${updatedCount} slide notes updated across ${successes.length} ${presentationWord}`
    : `${updatedCount} slide notes updated`;
  els.downloadLink.href = isBatch ? `/api/download/${state.sessionId}` : successes[0].downloadUrl;
  els.downloadLink.textContent = isBatch ? "Download all" : "Download";
  els.downloadLink.download = "";

  const warningLines = [];
  successes.forEach((result) => {
    (result.warnings || []).forEach((warning) => {
      warningLines.push(`${result.filename}: ${warning}`);
    });
  });
  failures.forEach((failure) => {
    warningLines.push(`${failure.filename}: ${failure.error}`);
  });

  if (warningLines.length) {
    els.warnings.innerHTML = warningLines.map(escapeHtml).join("<br />");
    els.warnings.classList.remove("hidden");
  } else {
    els.warnings.classList.add("hidden");
  }

  state.resultComparison = successes.flatMap((result) =>
    (result.comparison || []).map((slide) => ({
      ...slide,
      filename: result.filename,
      showFilename: isBatch,
    })),
  );
  state.resultExpanded = false;
  renderComparisonRows();
}

function renderComparisonRows() {
  els.comparison.innerHTML = "";
  state.resultComparison.forEach((slide) => {
    const pair = document.createElement("article");
    pair.className = "slide-pair";
    const deckLabel = slide.showFilename ? `<p class="deck-label">${escapeHtml(slide.filename)}</p>` : "";
    pair.innerHTML = `
      <div class="note-panel">
        ${deckLabel}
        <h3><span class="slide-label">Slide ${slide.number}</span> original</h3>
        <pre></pre>
      </div>
      <div class="note-panel">
        ${deckLabel}
        <h3><span class="slide-label">Slide ${slide.number}</span> summarized</h3>
        <pre></pre>
      </div>
    `;
    const [original, summarized] = pair.querySelectorAll("pre");
    original.textContent = slide.originalNotes || "";
    summarized.textContent = slide.summarizedNotes || "";
    els.comparison.appendChild(pair);
  });
  scheduleComparisonSizing();
}

function scheduleComparisonSizing() {
  window.cancelAnimationFrame(comparisonSizingFrame);
  comparisonSizingFrame = window.requestAnimationFrame(updateComparisonVisibility);
}

function updateComparisonVisibility() {
  comparisonSizingFrame = null;
  if (els.shell.dataset.view !== "result") return;

  const rows = Array.from(els.comparison.querySelectorAll(".slide-pair"));
  rows.forEach((row) => row.classList.remove("is-hidden"));
  els.showAllButton.classList.add("hidden");
}

function comparisonRowsThatFit(rows) {
  const viewportBottom = window.innerHeight || document.documentElement.clientHeight;
  const reserveForButton = rows.length > 1 ? 58 : 0;
  const bottomPadding = 24;
  const fitLine = viewportBottom - reserveForButton - bottomPadding;
  let visibleCount = 0;

  for (const row of rows) {
    if (row.getBoundingClientRect().bottom <= fitLine) {
      visibleCount += 1;
    } else {
      break;
    }
  }

  return Math.max(1, visibleCount);
}

function expandComparison() {
  state.resultExpanded = true;
  updateComparisonVisibility();
}

function setPromptExpanded(expanded) {
  state.promptExpanded = Boolean(expanded);
  els.promptPanel.classList.toggle("is-expanded", state.promptExpanded);
  els.promptInput.classList.toggle("hidden", !state.promptExpanded);
  els.promptEditButton.textContent = state.promptExpanded ? "Done" : "Edit";
  els.promptEditButton.setAttribute("aria-expanded", String(state.promptExpanded));
  scheduleScrollCueUpdate();
}

function togglePromptEditor() {
  setPromptExpanded(!state.promptExpanded);
  if (state.promptExpanded) {
    els.promptInput.focus();
  }
}

function scheduleScrollCueUpdate() {
  window.requestAnimationFrame(updateScrollCue);
}

function updateScrollCue() {
  const canScrollMore = window.scrollY + window.innerHeight < document.documentElement.scrollHeight - 18;
  els.shell.classList.toggle("can-scroll", canScrollMore);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

els.fileInput.addEventListener("change", (event) => {
  analyzeFiles(event.target.files);
});

["dragenter", "dragover"].forEach((name) => {
  els.dropZone.addEventListener(name, (event) => {
    event.preventDefault();
    els.dropZone.classList.add("dragover");
  });
});

["dragleave", "drop"].forEach((name) => {
  els.dropZone.addEventListener(name, (event) => {
    event.preventDefault();
    els.dropZone.classList.remove("dragover");
  });
});

els.dropZone.addEventListener("drop", (event) => {
  analyzeFiles(event.dataTransfer.files);
});

els.summarizeButton.addEventListener("click", summarize);
els.cancelButton.addEventListener("click", cancel);
els.cancelButtonSecondary.addEventListener("click", cancel);
els.resultHomeButton.addEventListener("click", goHome);
els.showAllButton.addEventListener("click", expandComparison);
els.themeToggle.addEventListener("click", toggleTheme);
els.codexAuthButton.addEventListener("click", startCodexOAuth);
els.codexDetailsButton.addEventListener("click", toggleCodexDetails);
els.keyInput.addEventListener("input", () => renderProviderSetupCard());
els.saveKeyButton.addEventListener("click", saveProviderKey);
els.forgetKeyButton.addEventListener("click", forgetProviderKey);
els.promptEditButton.addEventListener("click", togglePromptEditor);
window.addEventListener("resize", scheduleComparisonSizing);
window.addEventListener("resize", scheduleScrollCueUpdate);
window.addEventListener("scroll", updateScrollCue, { passive: true });
scheduleScrollCueUpdate();
