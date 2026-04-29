const state = {
  sessionId: null,
  filename: null,
  providers: [],
  provider: "codex",
};

const els = {
  status: document.querySelector("#status-pill"),
  toast: document.querySelector("#toast"),
  uploadView: document.querySelector("#upload-view"),
  setupView: document.querySelector("#setup-view"),
  resultView: document.querySelector("#result-view"),
  dropZone: document.querySelector("#drop-zone"),
  fileInput: document.querySelector("#file-input"),
  deckName: document.querySelector("#deck-name"),
  slideCount: document.querySelector("#slide-count"),
  noteCount: document.querySelector("#note-count"),
  writableCount: document.querySelector("#writable-count"),
  providerGroup: document.querySelector("#provider-group"),
  keyField: document.querySelector("#key-field"),
  keyLabel: document.querySelector("#key-label"),
  keyInput: document.querySelector("#key-input"),
  modelInput: document.querySelector("#model-input"),
  modelOptions: document.querySelector("#model-options"),
  promptInput: document.querySelector("#prompt-input"),
  cancelButton: document.querySelector("#cancel-button"),
  cancelButtonSecondary: document.querySelector("#cancel-button-secondary"),
  summarizeButton: document.querySelector("#summarize-button"),
  resultTitle: document.querySelector("#result-title"),
  downloadLink: document.querySelector("#download-link"),
  warnings: document.querySelector("#warnings"),
  comparison: document.querySelector("#comparison"),
};

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
  els.uploadView.classList.toggle("hidden", view !== "upload");
  els.setupView.classList.toggle("hidden", view !== "setup");
  els.resultView.classList.toggle("hidden", view !== "result");
}

function setBusy(isBusy, label = "Working") {
  els.summarizeButton.disabled = isBusy;
  els.cancelButton.disabled = isBusy;
  els.cancelButtonSecondary.disabled = isBusy;
  setStatus(isBusy ? label : "Ready");
}

async function analyzeFile(file) {
  if (!file) return;
  const form = new FormData();
  form.append("file", file);

  setBusy(true, "Reading");
  try {
    const response = await fetch("/api/analyze", { method: "POST", body: form });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Upload failed.");
    state.sessionId = data.sessionId;
    state.filename = data.filename;
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

function renderSetup(data) {
  els.deckName.textContent = data.filename;
  els.slideCount.textContent = data.slideCount;
  els.noteCount.textContent = data.noteCount;
  els.writableCount.textContent = data.writableNoteCount;
  els.promptInput.value = data.prompt;
  state.providers = data.providers || [];
  state.provider = data.provider || "codex";
  renderProviders();
  selectProvider(state.provider);
}

function renderProviders() {
  els.providerGroup.innerHTML = "";
  state.providers.forEach((provider) => {
    const label = document.createElement("label");
    label.innerHTML = `
      <input type="radio" name="provider" value="${escapeHtml(provider.id)}" />
      <span>${escapeHtml(provider.shortLabel)}</span>
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
  els.keyLabel.textContent = provider.keyLabel || "API key";
  els.keyInput.placeholder = provider.envKey ? `Optional if ${provider.envKey} is set` : "";
  els.keyInput.value = "";
  els.modelInput.value = provider.defaultModel;
  els.modelOptions.innerHTML = "";
  provider.models.forEach((model) => {
    const option = document.createElement("option");
    option.value = model;
    els.modelOptions.appendChild(option);
  });
}

async function summarize() {
  if (!state.sessionId) return;
  setBusy(true, "Summarizing");
  try {
    const response = await fetch("/api/summarize", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        sessionId: state.sessionId,
        provider: state.provider,
        model: els.modelInput.value.trim(),
        apiKey: els.keyInput.value.trim(),
        prompt: els.promptInput.value.trim(),
      }),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Summarization failed.");
    renderResult(data);
    setView("result");
    setStatus("Done");
  } catch (error) {
    showToast(error.message);
    setStatus("Loaded");
  } finally {
    setBusy(false);
  }
}

async function cancel() {
  const sessionId = state.sessionId;
  state.sessionId = null;
  state.filename = null;
  state.providers = [];
  state.provider = "codex";
  els.keyInput.value = "";
  if (sessionId) {
    fetch("/api/cancel", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sessionId }),
    });
  }
  els.comparison.innerHTML = "";
  els.warnings.innerHTML = "";
  els.warnings.classList.add("hidden");
  setView("upload");
  setStatus("Ready");
}

function renderResult(data) {
  els.resultTitle.textContent = `${data.updatedCount} slide notes updated`;
  els.downloadLink.href = data.downloadUrl;
  els.downloadLink.download = "";

  if (data.warnings && data.warnings.length) {
    els.warnings.innerHTML = data.warnings.map(escapeHtml).join("<br />");
    els.warnings.classList.remove("hidden");
  } else {
    els.warnings.classList.add("hidden");
  }

  els.comparison.innerHTML = "";
  data.comparison.forEach((slide) => {
    const pair = document.createElement("article");
    pair.className = "slide-pair";
    pair.innerHTML = `
      <div class="note-panel">
        <h3><span class="slide-label">Slide ${slide.number}</span> original</h3>
        <pre></pre>
      </div>
      <div class="note-panel">
        <h3><span class="slide-label">Slide ${slide.number}</span> summarized</h3>
        <pre></pre>
      </div>
    `;
    const [original, summarized] = pair.querySelectorAll("pre");
    original.textContent = slide.originalNotes || "";
    summarized.textContent = slide.summarizedNotes || "";
    els.comparison.appendChild(pair);
  });
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
  analyzeFile(event.target.files[0]);
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
  analyzeFile(event.dataTransfer.files[0]);
});

els.summarizeButton.addEventListener("click", summarize);
els.cancelButton.addEventListener("click", cancel);
els.cancelButtonSecondary.addEventListener("click", cancel);
