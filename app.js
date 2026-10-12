"use strict";

const STORAGE_KEY = "jigyasa-ai-demo-v1";
const ACCEPTED_EXTENSIONS = new Set([
  "pdf", "pptx", "docx", "txt", "mp4", "mov", "webm", "opus", "ogg", "mp3", "m4a", "wav", "aac",
]);
const pageLabels = {
  overview: "Overview",
  study: "Study with AI",
  quiz: "Quiz arena",
  library: "My library",
  progress: "My progress",
  history: "History",
  settings: "Settings",
};
const starterQuestions = [
  { question: "Which structure is the basic unit of life?", options: ["An atom", "A cell", "A tissue", "An organ"], answer: 1, topic: "Biology" },
  { question: "What is the main role of mitochondria?", options: ["Store genetic information", "Make proteins", "Release usable energy", "Control water loss"], answer: 2, topic: "Biology" },
  { question: "Which molecule carries hereditary information?", options: ["DNA", "Glucose", "Water", "Chlorophyll"], answer: 0, topic: "Biology" },
  { question: "What do plant cells use to capture light energy?", options: ["Ribosomes", "Chlorophyll", "Keratin", "Hemoglobin"], answer: 1, topic: "Biology" },
  { question: "What is the movement of water through a selectively permeable membrane called?", options: ["Diffusion", "Respiration", "Osmosis", "Digestion"], answer: 2, topic: "Biology" },
];
const topicQuestion = {
  "Cell biology": starterQuestions,
  "Photosynthesis": [
    { question: "Which pigment absorbs light for photosynthesis?", options: ["Melanin", "Chlorophyll", "Keratin", "Hemoglobin"], answer: 1, topic: "Photosynthesis" },
    { question: "Which gas do plants take in during photosynthesis?", options: ["Oxygen", "Nitrogen", "Carbon dioxide", "Hydrogen"], answer: 2, topic: "Photosynthesis" },
    { question: "Where does photosynthesis mainly take place?", options: ["Mitochondria", "Chloroplasts", "Nucleus", "Ribosomes"], answer: 1, topic: "Photosynthesis" },
    { question: "What sugar is produced during photosynthesis?", options: ["Glucose", "Lactose", "Fructose only", "Starch"], answer: 0, topic: "Photosynthesis" },
    { question: "Which is released as a by-product of photosynthesis?", options: ["Carbon dioxide", "Oxygen", "Methane", "Nitrogen"], answer: 1, topic: "Photosynthesis" },
  ],
  "Human body": [
    { question: "Which organ pumps blood around the body?", options: ["Lungs", "Liver", "Heart", "Kidneys"], answer: 2, topic: "Human body" },
    { question: "What is the largest organ of the human body?", options: ["Skin", "Liver", "Brain", "Lung"], answer: 0, topic: "Human body" },
    { question: "Which blood cells help the body fight infection?", options: ["Red blood cells", "White blood cells", "Platelets", "Plasma"], answer: 1, topic: "Human body" },
    { question: "Where does most nutrient absorption happen?", options: ["Stomach", "Large intestine", "Small intestine", "Esophagus"], answer: 2, topic: "Human body" },
    { question: "How many chambers does a human heart have?", options: ["Two", "Three", "Four", "Five"], answer: 2, topic: "Human body" },
  ],
  "World geography": [
    { question: "Which is the largest ocean on Earth?", options: ["Atlantic", "Indian", "Arctic", "Pacific"], answer: 3, topic: "World geography" },
    { question: "What is the capital of Japan?", options: ["Kyoto", "Tokyo", "Osaka", "Hiroshima"], answer: 1, topic: "World geography" },
    { question: "Which continent is the Sahara Desert in?", options: ["Asia", "South America", "Africa", "Australia"], answer: 2, topic: "World geography" },
    { question: "Which imaginary line divides Earth into Northern and Southern Hemispheres?", options: ["Prime Meridian", "Equator", "Tropic of Cancer", "International Date Line"], answer: 1, topic: "World geography" },
    { question: "Which is the smallest continent by land area?", options: ["Europe", "Antarctica", "South America", "Australia"], answer: 3, topic: "World geography" },
  ],
};
const icons = {
  upload: '<svg viewBox="0 0 24 24"><path d="M12 16V4m0 0L7 9m5-5 5 5"/><path d="M20 16.5v2A1.5 1.5 0 0 1 18.5 20h-13A1.5 1.5 0 0 1 4 18.5v-2"/></svg>',
  send: '<svg viewBox="0 0 24 24"><path d="m21 3-7.5 18-3.7-7.8L2 9.5 21 3Z"/><path d="M21 3 9.8 13.2"/></svg>',
  attach: '<svg viewBox="0 0 24 24"><path d="m8 12.5 7.1-7.1a3.2 3.2 0 0 1 4.5 4.5l-9.2 9.2a5 5 0 0 1-7.1-7.1l9.2-9.2"/></svg>',
  sparkle: '<svg viewBox="0 0 24 24"><path d="m12 3 1.8 6.2L20 11l-6.2 1.8L12 19l-1.8-6.2L4 11l6.2-1.8L12 3Z"/><path d="m19 15 .8 2.2L22 18l-2.2.8L19 21l-.8-2.2L16 18l2.2-.8L19 15Z"/></svg>',
};

function emptyState() {
  return { files: [], history: [], quizScores: [], chatMessages: [] };
}

function loadState(storageKey = STORAGE_KEY) {
  try {
    const saved = localStorage.getItem(storageKey);
    if (!saved) return emptyState();
    const parsed = JSON.parse(saved);
    return {
      files: Array.isArray(parsed.files) ? parsed.files : [],
      history: Array.isArray(parsed.history) ? parsed.history : [],
      quizScores: Array.isArray(parsed.quizScores) ? parsed.quizScores : [],
      chatMessages: Array.isArray(parsed.chatMessages) ? parsed.chatMessages : [],
    };
  } catch (error) {
    console.error("Could not load Jigyasa AI demo state.", error);
    return emptyState();
  }
}

let state = loadState();
let activeStorageKey = STORAGE_KEY;
let currentPage = "overview";
let quizSession = null;
let pendingFiles = [];
let serverDocuments = null;
let backendHealth = null;
let backendOnline = false;
let authenticatedUser = null;
let authMode = "login";
let toastTimer = 0;
let cloudSyncQueue = Promise.resolve();
let cloudStateReady = false;
const pageContent = document.getElementById("page-content");
const fileInput = document.getElementById("file-input");

function saveState() {
  try {
    localStorage.setItem(activeStorageKey, JSON.stringify(state));
  } catch (error) {
    console.error("Could not save Jigyasa AI demo state.", error);
    showToast("Could not save this update in your browser storage.");
  }
  if (!backendHealth?.auth_required || !authenticatedUser || !backendOnline) return Promise.resolve(true);
  if (!cloudStateReady) {
    showToast("Your account history has not loaded yet; this update is saved only on this device.");
    return Promise.resolve(false);
  }
  const snapshot = {
    history: state.history,
    quizScores: state.quizScores,
    chatMessages: state.chatMessages,
  };
  const pendingSync = cloudSyncQueue.then(async () => {
    const response = await fetch("/api/learning-state", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(snapshot),
    });
    if (!response.ok) {
      const result = await response.json();
      throw new Error(result.detail || `Learning history sync failed (HTTP ${response.status}).`);
    }
  });
  const syncResult = pendingSync.then(() => true).catch((error) => {
    console.error("Could not sync learning history to the account.", error);
    showToast("This update is saved on this device, but could not sync to your account.");
    return false;
  });
  cloudSyncQueue = syncResult.then(() => undefined);
  return syncResult;
}

async function restoreCloudLearningState() {
  cloudStateReady = false;
  const response = await fetch("/api/learning-state");
  const result = await response.json();
  if (!response.ok) throw new Error(result.detail || `Learning history could not be loaded (HTTP ${response.status}).`);
  if (result.has_cloud_state) {
    state = { ...emptyState(), ...result.state, files: [] };
  } else {
    state = { ...state, files: [] };
    const saveResponse = await fetch("/api/learning-state", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        history: state.history,
        quizScores: state.quizScores,
        chatMessages: state.chatMessages,
      }),
    });
    if (!saveResponse.ok) {
      const saveResult = await saveResponse.json();
      throw new Error(saveResult.detail || `This device's learning history could not be migrated (HTTP ${saveResponse.status}).`);
    }
  }
  try {
    localStorage.setItem(activeStorageKey, JSON.stringify(state));
  } catch (error) {
    console.error("Could not cache synced learning history in this browser.", error);
    showToast("Your account history loaded, but could not be cached on this device.");
  }
  cloudStateReady = true;
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  })[character]);
}

function formatInlineMarkdown(value) {
  const protectedContent = [];
  const protect = (html) => {
    const token = `\u0000${protectedContent.length}\u0000`;
    protectedContent.push(html);
    return token;
  };
  const withCode = value.replace(/`([^`\n]+)`/g, (_, code) => {
    return protect(`<code>${code}</code>`);
  });
  const withMath = withCode.replace(/\\\[([\s\S]*?)\\\]|\$\$([\s\S]*?)\$\$|\$([^$\n]+)\$|\\\((.+?)\\\)/g, (_, displayBracket, displayDollar, inlineDollar, inlineParen) => {
    const expression = displayBracket ?? displayDollar ?? inlineDollar ?? inlineParen;
    return protect(`<span class="math">${renderMathExpression(expression)}</span>`);
  });
  return withMath
    .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
    .replace(/__(.+?)__/g, "<strong>$1</strong>")
    .replace(/(^|[^*])\*([^*\n]+)\*(?!\*)/g, "$1<em>$2</em>")
    .replace(/(^|[^_])_([^_\n]+)_(?!_)/g, "$1<em>$2</em>")
    .replace(/\u0000(\d+)\u0000/g, (_, index) => protectedContent[Number(index)]);
}

function renderMathExpression(expression) {
  const symbols = {
    alpha: "α", beta: "β", gamma: "γ", delta: "δ", epsilon: "ε", varepsilon: "ε",
    zeta: "ζ", eta: "η", theta: "θ", vartheta: "θ", iota: "ι", kappa: "κ",
    lambda: "λ", mu: "μ", nu: "ν", xi: "ξ", pi: "π", rho: "ρ", sigma: "σ",
    tau: "τ", upsilon: "υ", phi: "φ", varphi: "φ", chi: "χ", psi: "ψ", omega: "ω",
    Gamma: "Γ", Delta: "Δ", Theta: "Θ", Lambda: "Λ", Xi: "Ξ", Pi: "Π",
    Sigma: "Σ", Phi: "Φ", Psi: "Ψ", Omega: "Ω",
    to: "→", rightarrow: "→", longrightarrow: "⟶", leftarrow: "←", gets: "←",
    leftrightarrow: "↔", mapsto: "↦", implies: "⇒", iff: "⇔",
    times: "×", cdot: "·", pm: "±", mp: "∓", le: "≤", leq: "≤", ge: "≥", geq: "≥",
    neq: "≠", ne: "≠", approx: "≈", equiv: "≡", in: "∈", notin: "∉",
    subset: "⊂", subseteq: "⊆", supset: "⊃", supseteq: "⊇", cup: "∪", cap: "∩",
    emptyset: "∅", infinity: "∞", partial: "∂", nabla: "∇", forall: "∀",
    exists: "∃", sum: "∑", prod: "∏", int: "∫",
  };
  const subSupContent = (content) => content
    .replace(/\\([A-Za-z]+)/g, (match, command) => symbols[command] ?? match)
    .replace(/[{}]/g, "");

  return expression
    .replace(/\\(?:left|right)\b/g, "")
    .replace(/\\(?:,|;|!|:|quad|qquad)\s*/g, " ")
    .replace(/\\text\{([^{}]*)\}/g, "$1")
    .replace(/\\frac\{([^{}]+)\}\{([^{}]+)\}/g, '<span class="math-fraction"><span>$1</span><span>$2</span></span>')
    .replace(/\\sqrt\{([^{}]+)\}/g, "√($1)")
    .replace(/_\{([^{}]+)\}|_([A-Za-z0-9+-])/g, (_, group, character) => `<sub>${subSupContent(group ?? character)}</sub>`)
    .replace(/\^\{([^{}]+)\}|\^([A-Za-z0-9+-])/g, (_, group, character) => `<sup>${subSupContent(group ?? character)}</sup>`)
    .replace(/\\([A-Za-z]+)/g, (match, command) => symbols[command] ?? match)
    .replace(/[{}]/g, "")
    .replace(/\\ /g, " ")
    .trim();
}

function renderMarkdown(text) {
  const lines = escapeHtml(text).replace(/\r\n?/g, "\n").split("\n");
  const output = [];
  let paragraph = [];
  let listType = null;
  let codeLines = null;

  const flushParagraph = () => {
    if (paragraph.length) {
      output.push(`<p>${paragraph.map(formatInlineMarkdown).join("<br>")}</p>`);
      paragraph = [];
    }
  };
  const closeList = () => {
    if (listType) {
      output.push(`</${listType}>`);
      listType = null;
    }
  };

  for (const line of lines) {
    if (codeLines) {
      if (/^\s*```/.test(line)) {
        output.push(`<pre><code>${codeLines.join("\n")}</code></pre>`);
        codeLines = null;
      } else {
        codeLines.push(line);
      }
      continue;
    }
    if (/^\s*```/.test(line)) {
      flushParagraph();
      closeList();
      codeLines = [];
      continue;
    }
    if (!line.trim()) {
      flushParagraph();
      closeList();
      continue;
    }
    if (/^\s*(?:---+|___+|\*\*\*+)\s*$/.test(line)) {
      flushParagraph();
      closeList();
      continue;
    }

    const heading = line.match(/^\s{0,3}(#{1,6})\s+(.+?)\s*#*\s*$/);
    if (heading) {
      flushParagraph();
      closeList();
      const level = heading[1].length <= 2 ? 3 : 4;
      output.push(`<h${level}>${formatInlineMarkdown(heading[2])}</h${level}>`);
      continue;
    }

    const listItem = line.match(/^\s*(?:([-*+])|(\d+)[.)])\s+(.+)$/);
    if (listItem) {
      flushParagraph();
      const nextListType = listItem[2] ? "ol" : "ul";
      if (listType !== nextListType) {
        closeList();
        output.push(`<${nextListType}>`);
        listType = nextListType;
      }
      output.push(`<li>${formatInlineMarkdown(listItem[3])}</li>`);
      continue;
    }
    closeList();
    paragraph.push(line);
  }

  if (codeLines) output.push(`<pre><code>${codeLines.join("\n")}</code></pre>`);
  flushParagraph();
  closeList();
  return output.join("");
}

function formatBytes(bytes) {
  if (!bytes) return "0 KB";
  const units = ["B", "KB", "MB", "GB"];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${(bytes / (1024 ** index)).toFixed(index === 0 ? 0 : 1)} ${units[index]}`;
}

function formatDate(timestamp, options = { month: "short", day: "numeric" }) {
  return new Date(timestamp).toLocaleDateString(undefined, options);
}

function showToast(message) {
  const region = document.getElementById("toast-region");
  const toast = document.createElement("div");
  toast.className = "toast";
  toast.textContent = message;
  region.append(toast);
  window.setTimeout(() => toast.remove(), 4200);
}

function setPage(page) {
  if (!pageLabels[page]) return;
  currentPage = page;
  document.getElementById("page-breadcrumb").textContent = pageLabels[page];
  document.querySelectorAll(".nav-item[data-page]").forEach((button) => {
    button.classList.toggle("active", button.dataset.page === page);
  });
  document.getElementById("sidebar").classList.remove("open");
  document.getElementById("mobile-menu").setAttribute("aria-expanded", "false");
  renderPage();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

function updateCounts() {
  document.getElementById("library-count").textContent = activeFiles().length;
  document.getElementById("streak-count").textContent = state.quizScores.length > 0 ? "1" : "0";
}

function activeFiles() {
  return serverDocuments ?? state.files;
}

async function refreshServerDocuments() {
  const response = await fetch("/api/documents");
  if (!response.ok) throw new Error(`The local library could not be loaded (HTTP ${response.status}).`);
  const documents = await response.json();
  if (!Array.isArray(documents)) throw new Error("The local library returned an unexpected response.");
  serverDocuments = documents.map((document) => ({
    id: document.id,
    name: document.name,
    ext: document.ext,
    size: document.size,
    addedAt: document.added_at,
    server: true,
  }));
  pendingFiles = pendingFiles
    .map((pending) => serverDocuments.find((document) => document.id === pending.id))
    .filter(Boolean);
}

async function initializeBackend() {
  if (window.location.protocol !== "http:" && window.location.protocol !== "https:") {
    document.querySelector(".app-shell").classList.add("ready");
    return;
  }
  try {
    const response = await fetch("/api/health");
    if (!response.ok) throw new Error(`The local backend returned HTTP ${response.status}.`);
    backendHealth = await response.json();
    if (backendHealth.auth_required) {
      await consumeEmailConfirmation();
      const authResponse = await fetch("/api/auth/me");
      if (authResponse.status === 401) {
        showAuthGate();
        return;
      }
      const profile = await authResponse.json();
      if (!authResponse.ok || !profile.authenticated) {
        throw new Error(profile.detail || "Could not verify your account session.");
      }
      authenticatedUser = profile;
      updateProfileDisplay();
      activeStorageKey = `${STORAGE_KEY}:${profile.id}`;
      state = loadState(activeStorageKey);
    } else {
      activeStorageKey = STORAGE_KEY;
      state = loadState(activeStorageKey);
    }
    await refreshServerDocuments();
    backendOnline = true;
    if (backendHealth.auth_required) {
      try {
        await restoreCloudLearningState();
      } catch (error) {
        console.error("Could not load or migrate account learning history.", error);
        showToast(error.message || "Learning history could not be synced. It remains on this device.");
      }
    }
    document.getElementById("auth-gate").hidden = true;
    document.querySelector(".app-shell").classList.add("ready");
    renderPage();
  } catch (error) {
    backendOnline = false;
    console.error("Could not connect to the local Jigyasa backend.", error);
    if (backendHealth?.auth_required) {
      showAuthGate(error.message || "The secure account service is unavailable. Please try again.");
    } else {
      document.querySelector(".app-shell").classList.add("ready");
      showToast("Can't reach the Jigyasa backend. Start the project with run.ps1 and refresh this page.");
    }
  }
}

async function consumeEmailConfirmation() {
  const callback = new URLSearchParams(window.location.hash.slice(1));
  if (!callback.has("access_token") && !callback.has("refresh_token")) return;
  const accessToken = callback.get("access_token");
  const refreshToken = callback.get("refresh_token");
  const callbackType = callback.get("type");
  window.history.replaceState(null, "", `${window.location.pathname}${window.location.search}`);
  if (callbackType !== "signup" || !accessToken || !refreshToken) {
    throw new Error("That email confirmation link is incomplete or expired. Sign in or request a fresh confirmation email.");
  }
  const response = await fetch("/api/auth/confirm", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ access_token: accessToken, refresh_token: refreshToken }),
  });
  const result = await response.json();
  if (!response.ok) throw new Error(result.detail || "Could not verify the email confirmation.");
}

function showAuthGate(message = "") {
  document.querySelector(".app-shell").classList.remove("ready");
  const gate = document.getElementById("auth-gate");
  gate.hidden = false;
  document.getElementById("auth-error").textContent = message;
  updateAuthForm();
}

function updateAuthForm() {
  const signingUp = authMode === "signup";
  document.getElementById("auth-title").textContent = signingUp ? "Create your account" : "Welcome back";
  document.getElementById("auth-copy").textContent = signingUp
    ? "Create a private space for your study materials and learning history."
    : "Sign in to access your study materials and learning history.";
  document.getElementById("auth-password").autocomplete = signingUp ? "new-password" : "current-password";
  document.getElementById("auth-submit").textContent = signingUp ? "Create account" : "Sign in";
  document.getElementById("auth-toggle-copy").innerHTML = signingUp
    ? 'Already have an account? <button class="text-link" type="button" data-auth-toggle>Sign in</button>'
    : 'New to Jigyasa AI? <button class="text-link" type="button" data-auth-toggle>Create an account</button>';
}

function updateProfileDisplay() {
  const email = authenticatedUser?.email || "Student";
  const initial = email.slice(0, 1).toUpperCase();
  document.querySelector(".profile-info strong").textContent = email;
  document.querySelector(".profile-info span").textContent = backendHealth?.auth_required ? "Private learning space" : "Your learning space";
  document.querySelector(".profile-card .avatar").textContent = initial;
  document.querySelector(".top-avatar").textContent = initial;
  document.querySelector(".top-avatar").setAttribute("aria-label", `${email} profile`);
}

async function submitAuthentication(event) {
  event.preventDefault();
  const submit = document.getElementById("auth-submit");
  const errorRegion = document.getElementById("auth-error");
  const email = document.getElementById("auth-email").value.trim();
  const password = document.getElementById("auth-password").value;
  submit.disabled = true;
  errorRegion.textContent = "";
  try {
    const response = await fetch(`/api/auth/${authMode}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ email, password }),
    });
    const result = await response.json();
    if (!response.ok) throw new Error(result.detail || "Could not complete sign-in.");
    if (result.email_confirmation_required) {
      authMode = "login";
      updateAuthForm();
      errorRegion.textContent = "Check your email for a confirmation link, then sign in.";
      return;
    }
    const profile = await (await fetch("/api/auth/me")).json();
    if (!profile.authenticated || !profile.id) throw new Error("Your account session could not be verified. Please sign in again.");
    authenticatedUser = profile;
    updateProfileDisplay();
    activeStorageKey = `${STORAGE_KEY}:${profile.id}`;
    state = loadState(activeStorageKey);
    await refreshServerDocuments();
    backendOnline = true;
    try {
      await restoreCloudLearningState();
    } catch (error) {
      console.error("Could not load or migrate account learning history.", error);
      showToast(error.message || "Learning history could not be synced. It remains on this device.");
    }
    document.getElementById("auth-gate").hidden = true;
    document.querySelector(".app-shell").classList.add("ready");
    renderPage();
  } catch (error) {
    console.error("Account authentication failed.", error);
    errorRegion.textContent = error.message || "Could not complete sign-in. Please try again.";
  } finally {
    submit.disabled = false;
  }
}

function viewHeader(title, subtitle, action = "") {
  return `<div class="view-header"><div><p class="eyebrow">JIGYASA AI</p><h1 class="page-title">${title}</h1><p class="page-subtitle">${subtitle}</p></div>${action ? `<div class="view-header-actions">${action}</div>` : ""}</div>`;
}

function getQuizAverage() {
  if (state.quizScores.length === 0) return 0;
  return Math.round(state.quizScores.reduce((sum, item) => sum + item.score / item.total, 0) / state.quizScores.length * 100);
}

function renderDashboard() {
  const today = new Date();
  const dateLabel = today.toLocaleDateString(undefined, { weekday: "long", month: "long", day: "numeric" }).toUpperCase();
  const files = activeFiles();
  const recentFiles = [...files].slice(-3).reverse();
  const recentMarkup = recentFiles.length
    ? recentFiles.map((file) => `<div class="recent-item"><div class="recent-file-icon ${escapeHtml(file.ext)}">${escapeHtml(file.ext.toUpperCase().slice(0, 4))}</div><div class="recent-copy"><strong>${escapeHtml(file.name)}</strong><span>Added ${formatDate(file.addedAt)}</span></div><button class="recent-open" aria-label="Open study chat" data-action="open-study">↗</button></div>`).join("")
    : `<p class="empty-inline">Your study materials will show up here. Add a file to get started.</p>`;
  const quizAverage = getQuizAverage();
  const stats = [
    ["✦", "lilac", files.length, "Study materials"],
    ["◷", "mint", state.history.filter((item) => item.type === "chat").length, "Questions explored"],
    ["✓", "peach", state.quizScores.length, "Quizzes completed"],
    ["↗", "yellow", state.quizScores.length ? `${quizAverage}%` : "—", "Average quiz score"],
  ];
  const statMarkup = stats.map(([icon, color, value, label]) => `<div class="stat-card"><div class="stat-icon ${color}">${icon}</div><div><div class="stat-value">${value}</div><div class="stat-label">${label}</div></div></div>`).join("");
  const bars = [24, 43, 31, 59, 37, 77, 47].map((height, index) => {
    const activity = state.quizScores.length ? Math.min(88, 12 + state.quizScores.length * (index % 2 ? 18 : 12)) : height;
    return `<div class="chart-column"><div class="chart-bars"><span class="chart-bar" style="height:${Math.max(4, activity * .52)}%"></span><span class="chart-bar primary" style="height:${activity}%"></span></div><span>${["M", "T", "W", "T", "F", "S", "S"][index]}</span></div>`;
  }).join("");
  return `
    <div class="greeting-row"><div><p class="eyebrow">${dateLabel}</p><h1 class="page-title">Hey there! <span class="wave">✦</span></h1><p class="page-subtitle">A little curiosity goes a long way. What shall we learn today?</p></div><div class="date-label"><i class="date-dot"></i> Your learning space</div></div>
    <section class="hero-card">
      <div class="hero-copy"><div class="hero-kicker"><span>✦</span> YOUR AI STUDY COMPANION</div><h2 class="hero-title">Big questions.<br />Clear answers.</h2><p class="hero-description">Bring your notes, slides or lessons. Make learning feel a little more you.</p><div class="hero-actions"><button class="button" data-page="study">${icons.sparkle} Start learning</button><button class="button button-quiet" data-page="quiz">Try a quick quiz <span aria-hidden="true">→</span></button></div></div>
      <div class="hero-illustration" aria-hidden="true"><div class="orbit"></div><div class="orbit orbit-two"></div><span class="floating-star star-one">✦</span><span class="floating-star star-two">✧</span><span class="hero-person"></span><div class="book-stack"><span class="book book-one"></span><span class="book book-two"></span><span class="book book-three"></span><span class="book book-four"></span></div><span class="plant">✿</span></div>
    </section>
    <div class="section-heading"><h2>Your learning snapshot</h2><button class="text-link" data-page="progress">See progress <span>→</span></button></div>
    <div class="stats-grid">${statMarkup}</div>
    <div class="dashboard-lower">
      <section class="panel"><div class="panel-head"><h2>Your study rhythm</h2><span>Last 7 days</span></div><div class="activity-chart">${bars}</div></section>
      <section class="panel"><div class="panel-head"><h2>Recently added</h2><button class="text-link" data-page="library">View library <span>→</span></button></div><div class="recent-list">${recentMarkup}</div></section>
    </div>
    <div class="section-heading"><h2>Pick up where you like</h2></div>
    <div class="quick-grid">
      <button class="quick-card" data-page="study"><span class="stat-icon lilac">${icons.sparkle}</span><span><h3>Ask your study buddy</h3><p>Explore a tricky idea, one question at a time.</p></span><span class="quick-arrow">→</span></button>
      <button class="quick-card" data-page="quiz"><span class="stat-icon mint">✓</span><span><h3>Take a little quiz</h3><p>Practice a topic and see what sticks.</p></span><span class="quick-arrow">→</span></button>
      <button class="quick-card" data-action="upload">${`<span class="stat-icon peach">${icons.upload}</span>`}<span><h3>Add your materials</h3><p>Keep your notes in one handy place.</p></span><span class="quick-arrow">→</span></button>
    </div>`;
}

function chatMessagesMarkup() {
  return state.chatMessages.map((message) => {
    const content = message.role === "assistant" ? renderMarkdown(message.text) : escapeHtml(message.text);
    return `<div class="message ${message.role === "user" ? "user" : "assistant"}">${content}</div>`;
  }).join("");
}

function renderStudy() {
  const files = activeFiles();
  const attached = pendingFiles.map((file, index) => `<span class="attached-chip">📎 ${escapeHtml(file.name)} <button data-remove-pending="${index}" aria-label="Remove ${escapeHtml(file.name)}">×</button></span>`).join("");
  const storageLabel = backendHealth?.auth_required ? "your private account" : "this computer";
  const statusMessage = backendOnline
    ? `<div class="demo-banner"><span>ⓘ</span><span>${backendHealth?.gemini_configured
      ? `<strong>${backendHealth.auth_required ? "Secure study backend is connected." : "Local study backend is connected."}</strong> Selected study materials are sent to Gemini when you ask a question. Audio and videos are temporarily uploaded to Gemini for analysis; Jigyasa requests their deletion after each answer.`
      : `<strong>Document reading is connected.</strong> Gemini is not configured, so AI answers are unavailable. Files are stored in ${storageLabel}.`}</span></div>`
    : `<div class="demo-banner"><span>ⓘ</span><span><strong>Preview mode:</strong> Open this app using run.ps1 to enable document reading. Files opened directly in a browser are only saved as demo names.</span></div>`;
  const fileChoices = files.length
    ? `<div class="recent-list">${files.slice().reverse().map((file) => `<div class="recent-item"><div class="recent-file-icon ${escapeHtml(file.ext)}">${escapeHtml(file.ext.toUpperCase().slice(0, 3))}</div><div class="recent-copy"><strong>${escapeHtml(file.name)}</strong><span>${formatBytes(file.size)}</span></div><button class="topic-chip" data-select-document="${escapeHtml(file.id)}">${pendingFiles.some((item) => item.id === file.id) ? "Selected" : "Use"}</button></div>`).join("")}</div>`
    : `<p class="empty-inline">Add a supported document to begin a material-based study session.</p>`;
  const topics = ["Cell biology", "Photosynthesis", "Human body", "World geography"];
  return `${viewHeader("Study with AI", "Bring a question, a topic, or your own study material.")}
    ${statusMessage}
    <div class="study-layout" style="margin-top:14px">
      <section class="panel chat-panel">
        <div class="chat-topline"><div class="ai-orb">✦</div><div class="chat-heading"><strong>Your study buddy</strong><span><i class="online-dot"></i> ${backendHealth?.gemini_configured ? `${backendHealth.auth_required ? "Secure" : "Local"} backend · Gemini connected` : backendOnline ? `${backendHealth?.auth_required ? "Secure" : "Local"} backend · Gemini not configured` : "Demo mode · local backend not running"}</span></div><div class="chat-context">${files.length} material${files.length === 1 ? "" : "s"} added</div></div>
        <div class="chat-feed" id="chat-feed"><p class="chat-time">A FRESH START</p><div class="chat-welcome"><strong>Hi! I’m your study buddy.</strong><br />${backendOnline ? backendHealth?.gemini_configured ? "Select the documents you want to study, then ask your question. I’ll explain it using the material you chose." : "Select the documents you want to study, then ask a question. Gemini answers will be enabled after its Jigyasa-only key is configured." : "This is a preview. Start the local app with run.ps1 to read documents and connect the study chat."}</div>${chatMessagesMarkup()}</div>
        <div class="attached-files" id="attached-files">${attached}</div>
        <form class="chat-compose" id="chat-form"><div class="compose-box"><textarea id="chat-input" rows="1" placeholder="Ask a question about your selected materials..." aria-label="Type your question"></textarea><div class="compose-tools"><button type="button" class="compose-tool" data-action="upload" aria-label="Attach study material">${icons.attach}</button><button class="send-button" type="button" data-action="send-chat" aria-label="Send question">${icons.send}</button></div></div><p class="compose-note">${backendHealth?.daily_ai_question_limit ? `Up to ${backendHealth.daily_ai_question_limit} Gemini questions per day; resets at 00:00 UTC.` : "Questions are answered only from the selected material when Gemini is configured."}</p></form>
      </section>
      <aside class="study-side">
        <section class="panel side-card"><div class="side-card-head"><h3>Your materials</h3><span>${files.length} added</span></div><label class="upload-zone" id="upload-zone" for="file-input"><span class="upload-icon">↑</span><strong>Drop a file or browse</strong><span>PDF, PPTX, DOCX, TXT, MP4, MOV, WebM<br />OPUS/OGG, MP3, M4A, WAV or AAC<br />Up to 25 MB per file</span></label><p class="upload-footnote">Files are stored in ${storageLabel}${backendHealth?.auth_required ? " (100 MB and 25 files per account)" : ""}. Audio and video files are sent to Gemini only when selected for a question. Legacy PPT and scanned image-only PDFs are not supported.</p></section>
        <section class="panel side-card"><div class="side-card-head"><h3>Select study materials</h3><button class="text-link" data-page="library">See all →</button></div>${fileChoices}</section>
        <section class="panel side-card"><div class="side-card-head"><h3>Try a topic</h3><span>Quick prompts</span></div><div class="topic-chips">${topics.map((topic) => `<button class="topic-chip" data-topic="${escapeHtml(topic)}">${escapeHtml(topic)}</button>`).join("")}</div></section>
      </aside>
    </div>`;
}

function getFileExt(file) {
  return file.name.includes(".") ? file.name.split(".").pop().toLowerCase() : "file";
}

function renderLibrary() {
  const action = `<button class="button button-primary" data-action="upload">${icons.upload} Add materials</button>`;
  const files = [...activeFiles()].reverse();
  const content = files.length
    ? `<div class="file-list">${files.map((file) => `<div class="file-card"><div class="file-kind ${escapeHtml(file.ext)}">${escapeHtml(file.ext.toUpperCase().slice(0, 4))}</div><div class="file-details"><strong>${escapeHtml(file.name)}</strong><span>${formatBytes(file.size)} · Added ${formatDate(file.addedAt, { month: "short", day: "numeric", year: "numeric" })}</span></div><span class="file-status"><i></i> ${file.server ? backendHealth?.auth_required ? "Private cloud" : "On this computer" : "Demo only"}</span><button class="file-remove" data-remove-file="${escapeHtml(file.id)}" aria-label="Remove ${escapeHtml(file.name)}">×</button></div>`).join("")}</div>`
    : `<div class="panel empty-state"><div class="empty-illustration">↑</div><h3>Your library is ready for its first note</h3><p>Add documents, videos, or audio such as WhatsApp OPUS voice notes. Media files are analyzed by Gemini when selected for a question.</p><button class="button button-primary" data-action="upload">${icons.upload} Choose a file</button></div>`;
  return `${viewHeader("My library", "A home for your study materials.", action)}<div class="demo-banner" style="margin-bottom:14px"><span>ⓘ</span><span>PDF/PPTX/DOCX/TXT files are stored in ${backendHealth?.auth_required ? "your private cloud account" : "this computer"} and text-extracted. Audio and video files are stored privately and sent to Gemini for analysis when selected; deletion is requested after each answer. Scanned PDFs are not supported.</span></div>${content}`;
}

function getSelectedTopic() {
  return document.getElementById("quiz-topic")?.value || "Cell biology";
}

function startQuiz() {
  const selected = getSelectedTopic();
  const questionSet = topicQuestion[selected] || starterQuestions.map((question) => ({ ...question, topic: selected }));
  const requestedCount = Number(document.getElementById("quiz-count")?.value || 5);
  quizSession = {
    topic: selected,
    questions: questionSet.slice(0, requestedCount),
    index: 0,
    score: 0,
    selected: null,
    submitted: false,
  };
  renderQuiz();
}

function renderQuizCard() {
  const session = quizSession;
  if (!session) return `<div class="panel empty-state"><div class="empty-illustration">✦</div><h3>Ready when you are</h3><p>Choose a topic and start a quick quiz to check what you already know.</p></div>`;
  if (session.index >= session.questions.length) {
    const percentage = Math.round(session.score / session.questions.length * 100);
    return `<div class="panel quiz-summary"><div class="score-ring"><strong>${percentage}%</strong></div><h2>${percentage >= 80 ? "Brilliant work!" : percentage >= 50 ? "Nice effort!" : "Good start!"}</h2><p>You got <strong>${session.score} out of ${session.questions.length}</strong> right in ${escapeHtml(session.topic)}.<br />Every attempt helps you learn.</p><div class="breakdown-row"><span>Correct answers</span><strong>${session.score}</strong></div><div class="breakdown-row"><span>To revisit</span><strong>${session.questions.length - session.score}</strong></div><div style="display:flex;justify-content:center;gap:8px;margin-top:14px"><button class="button button-outline" data-action="quiz-again">Try again</button><button class="button button-primary" data-page="progress">See my progress →</button></div></div>`;
  }
  const question = session.questions[session.index];
  const letters = ["A", "B", "C", "D"];
  const options = question.options.map((option, index) => {
    let outcome = "";
    if (session.submitted && index === question.answer) outcome = " correct";
    else if (session.submitted && index === session.selected) outcome = " incorrect";
    else if (index === session.selected) outcome = " selected";
    return `<button class="answer-option${outcome}" data-answer="${index}" ${session.submitted ? "disabled" : ""}><span class="answer-letter">${letters[index]}</span>${escapeHtml(option)}</button>`;
  }).join("");
  let feedback = "";
  if (session.submitted) {
    feedback = session.selected === question.answer
      ? "That's right — lovely work! ✦"
      : `Not quite this time. The right answer is ${escapeHtml(question.options[question.answer])}.`;
  } else if (session.selected !== null) {
    feedback = "Choose check answer when you're ready.";
  }
  return `<div class="panel quiz-card"><div class="quiz-progress-row"><span>QUESTION ${session.index + 1} OF ${session.questions.length}</span><span>${session.score} correct so far</span></div><div class="progress-track"><div class="progress-fill" style="width:${session.index / session.questions.length * 100}%"></div></div><h2 class="question-heading">${escapeHtml(question.question)}</h2><div class="option-list">${options}</div><div class="quiz-feedback">${feedback}</div><div class="quiz-footer"><span class="topic-chip">${escapeHtml(question.topic)}</span>${session.submitted ? `<button class="button button-primary" data-action="quiz-next">${session.index === session.questions.length - 1 ? "See my results" : "Next question →"}</button>` : `<button class="button button-primary" data-action="quiz-submit" ${session.selected === null ? "disabled" : ""}>Check answer</button>`}</div></div>`;
}

function renderQuiz() {
  const session = quizSession;
  const questions = session?.questions || [];
  const sideContent = session && session.index >= session.questions.length
    ? `<div class="panel side-card"><div class="side-card-head"><h3>This attempt</h3></div><div class="breakdown-row"><span>Topic</span><strong>${escapeHtml(session.topic)}</strong></div><div class="breakdown-row"><span>Score</span><strong>${session.score}/${session.questions.length}</strong></div></div>`
    : `<div class="panel side-card"><div class="side-card-head"><h3>Little study tip</h3><span>✦</span></div><p class="page-subtitle">Try explaining the answer in your own words after each question. It helps the idea stick.</p></div><div class="panel side-card"><div class="side-card-head"><h3>Recent attempts</h3><button class="text-link" data-page="history">History →</button></div>${state.quizScores.slice(-3).reverse().map((item) => `<div class="topic-result"><div class="topic-result-icon">✓</div><div><strong>${escapeHtml(item.topic)}</strong><span>${item.score}/${item.total} correct · ${formatDate(item.at)}</span></div></div>`).join("") || `<p class="empty-inline">Your quiz attempts will show up here.</p>`}</div>`;
  pageContent.innerHTML = `${viewHeader("Quiz arena", "A relaxed space to practise, check in, and grow.")}
    <section class="quiz-hero"><p class="eyebrow">PRACTICE AT YOUR PACE</p><h2>Curiosity looks good on you.</h2><p>Pick a topic, take a quick quiz, and notice what you might want to revisit. These sample questions are just a preview.</p></section>
    <div class="quiz-control-row"><div class="field-group"><label for="quiz-topic">Choose a topic</label><select id="quiz-topic"><option>Cell biology</option><option>Photosynthesis</option><option>Human body</option><option>World geography</option></select></div><div class="field-group"><label for="quiz-count">Questions</label><select id="quiz-count"><option value="5">5 questions</option><option value="3">3 questions</option></select></div><button class="button button-primary" data-action="start-quiz">Start a quick quiz&nbsp; →</button><span class="topic-chip">✦ Sample quiz · no AI connected</span></div>
    <div class="quiz-layout"><div id="quiz-card">${renderQuizCard()}</div><aside class="study-side">${sideContent}</aside></div>`;
}

function renderHistory() {
  const items = [...state.history].reverse();
  const content = items.length
    ? `<div class="history-list">${items.map((item) => `<div class="history-card"><div class="history-icon">${item.type === "quiz" ? "✓" : "✦"}</div><div class="history-card-copy"><strong>${escapeHtml(item.title)}</strong><span>${item.type === "quiz" ? "Quiz" : "Study chat"} · ${formatDate(item.at, { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" })}</span></div>${item.type === "quiz" ? `<span class="history-result">${item.score}/${item.total}</span>` : `<span class="topic-chip">Demo</span>`}</div>`).join("")}</div>`
    : `<div class="panel empty-state"><div class="empty-illustration">◷</div><h3>Your learning journey starts here</h3><p>Questions you explore and quizzes you complete will appear here${backendHealth?.auth_required ? ", saved privately to your account" : ", saved in this browser"}.</p><button class="button button-primary" data-page="study">${icons.sparkle} Start learning</button></div>`;
  return `${viewHeader("Your history", "A little trail of what you've explored so far.")}${content}`;
}

function renderProgress() {
  const average = getQuizAverage();
  const quizCount = state.quizScores.length;
  const totalQuestions = state.quizScores.reduce((sum, item) => sum + item.total, 0);
  const right = state.quizScores.reduce((sum, item) => sum + item.score, 0);
  const topics = Object.entries(state.quizScores.reduce((counts, attempt) => {
    if (!counts[attempt.topic]) counts[attempt.topic] = { score: 0, total: 0 };
    counts[attempt.topic].score += attempt.score;
    counts[attempt.topic].total += attempt.total;
    return counts;
  }, {}));
  const strengths = topics.filter(([, result]) => result.score / result.total >= .7);
  const review = topics.filter(([, result]) => result.score / result.total < .7);
  const topicList = (items, strong) => items.length
    ? items.map(([name, result]) => `<span class="topic-tag ${strong ? "strong" : ""}">${escapeHtml(name)} · ${Math.round(result.score / result.total * 100)}%</span>`).join("")
    : `<p class="empty-inline">${quizCount ? "No topics here yet." : "Take a quiz to start spotting patterns."}</p>`;
  return `${viewHeader("My progress", "Notice the little wins and choose what to practise next.")}
    <div class="demo-banner" style="margin-bottom:14px"><span>ⓘ</span><span>This is a simple summary of your sample quiz attempts in this browser, not an AI assessment of your learning.</span></div>
    <div class="progress-grid"><section class="panel"><div class="panel-head"><h2>Your overall snapshot</h2><span>${quizCount} quiz${quizCount === 1 ? "" : "zes"}</span></div><div class="progress-score"><div class="stat-icon mint" style="width:47px;height:47px;font-size:19px">↗</div><div style="flex:1"><div class="progress-number">${quizCount ? `${average}%` : "—"}</div><div class="progress-note">Average across completed quizzes</div><div class="progress-meter"><span style="width:${average}%"></span></div></div></div>    <div class="recent-list"><div class="breakdown-row"><span>Questions answered</span><strong>${totalQuestions}</strong></div><div class="breakdown-row"><span>Correct answers</span><strong>${right}</strong></div><div class="breakdown-row"><span>Study materials added</span><strong>${activeFiles().length}</strong></div></div></section>
      <section class="panel"><div class="panel-head"><h2>Quiz activity</h2><span>Recent first</span></div><div class="recent-list">${state.quizScores.slice(-5).reverse().map((attempt) => `<div class="topic-result"><div class="topic-result-icon">${Math.round(attempt.score / attempt.total * 100)}%</div><div><strong>${escapeHtml(attempt.topic)}</strong><span>${attempt.score}/${attempt.total} correct · ${formatDate(attempt.at)}</span></div></div>`).join("") || `<p class="empty-inline">No quiz attempts yet. A quick practice quiz is a good place to start.</p>`}</div></section>
      <section class="panel topic-box"><h3>Feeling confident in</h3><div>${topicList(strengths, true)}</div></section><section class="panel topic-box"><h3>Worth another look</h3><div>${topicList(review, false)}</div></section></div>`;
}

function renderSettings() {
  if (backendHealth?.auth_required) {
    return `${viewHeader("Settings", "Manage your Jigyasa AI account.")}<section class="panel settings-card"><h2>Private study space</h2><p>Signed in as <strong>${escapeHtml(authenticatedUser?.email || "")}</strong>. Study materials and learning history are stored privately in your account and are not visible to other students. Chat, quiz scores, and activity sync across devices and stay until you clear them.</p><p>When you ask about audio or video, the selected file is sent to Gemini for analysis and Jigyasa requests deletion of the temporary Gemini file after the answer.</p><button class="button button-outline" data-action="clear-account-history">Clear synced learning history</button> <button class="button button-outline" data-action="logout">Sign out</button></section>`;
  }
  return `${viewHeader("Settings", "A few details about this local build.")}<section class="panel settings-card"><h2>Jigyasa AI local setup</h2><p>Uploaded PDF, PPTX, DOCX, TXT, video and audio files are stored on this computer. Text is extracted locally, while selected media is sent to Gemini for analysis and deletion is requested after each answer. Quiz attempts, chat history, and progress are saved only in this browser in local mode.</p><p>Gemini is not connected until a separate Jigyasa API key is added to the local <code>.env</code> file. Never put a key in browser code or send it in chat. This setup does not read or change the Anvesha or Ulekh project files.</p><p>Scanned-PDF OCR and cloud accounts are available only in the hosted configuration.</p><button class="button button-outline" data-action="clear-demo">Clear browser demo data</button></section>`;
}

function renderPage() {
  updateCounts();
  if (currentPage === "quiz") {
    renderQuiz();
    return;
  }
  const renderers = {
    overview: renderDashboard,
    study: renderStudy,
    library: renderLibrary,
    progress: renderProgress,
    history: renderHistory,
    settings: renderSettings,
  };
  pageContent.innerHTML = renderers[currentPage]();
}

function openFilePicker() {
  fileInput.click();
}

async function addFiles(fileList) {
  const files = Array.from(fileList || []);
  if (!files.length) return;
  const valid = [];
  const rejected = [];
  for (const file of files) {
    const extension = getFileExt(file);
    if (!ACCEPTED_EXTENSIONS.has(extension)) {
      rejected.push(file.name);
      continue;
    }
    if (file.size > 25 * 1024 * 1024) {
      rejected.push(`${file.name} (over 25 MB)`);
      continue;
    }
    valid.push(file);
  }
  if (valid.length && backendOnline) {
    let added = 0;
    for (const file of valid) {
      const formData = new FormData();
      formData.append("file", file);
      try {
        const response = await fetch("/api/documents", { method: "POST", body: formData });
        const result = await response.json();
        if (!response.ok) {
          rejected.push(`${file.name}: ${result.detail || `upload failed (HTTP ${response.status})`}`);
          continue;
        }
        added += 1;
        const storedDocument = {
          id: result.id,
          name: result.name,
          ext: result.ext,
          size: result.size,
          addedAt: result.added_at,
          server: true,
        };
        pendingFiles.push(storedDocument);
      } catch (error) {
        console.error(`Could not upload ${file.name}.`, error);
        rejected.push(`${file.name}: could not reach the local backend`);
      }
    }
    if (added) {
      try {
        await refreshServerDocuments();
        showToast(`${added} study material${added === 1 ? "" : "s"} uploaded to your library.`);
      } catch (error) {
        console.error("Could not refresh the local document library.", error);
        showToast(error.message || "The document was uploaded, but the library did not refresh.");
      }
    }
  } else if (valid.length) {
    state.files.push(...valid.map((file) => ({
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
      name: file.name,
      size: file.size,
      ext: getFileExt(file),
      addedAt: Date.now(),
    })));
    saveState();
    showToast(`${valid.length} file${valid.length === 1 ? "" : "s"} added to your demo library. They were not uploaded or read.`);
  }
  if (rejected.length) showToast(`Couldn't add: ${rejected.join(", ")}. Use PDF, PPTX, DOCX, TXT, MP4, MOV, WebM, OPUS, OGG, MP3, M4A, WAV, or AAC up to 25 MB.`);
  renderPage();
}

function recordHistory(item) {
  state.history.push({ ...item, at: Date.now() });
  state.history = state.history.slice(-80);
}

async function sendChatMessage(text) {
  const message = text.trim();
  if (!message) return;
  if (backendOnline) {
    if (pendingFiles.length === 0) {
      showToast("Select at least one study material before asking a document question.");
      return;
    }
    const input = document.getElementById("chat-input");
    const sendButton = document.querySelector('[data-action="send-chat"]');
    if (sendButton) sendButton.disabled = true;
    try {
      const response = await fetch("/api/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          question: message,
          document_ids: pendingFiles.map((file) => file.id),
        }),
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.detail || `Study chat failed (HTTP ${response.status}).`);
      state.chatMessages.push({ role: "user", text: message });
      state.chatMessages.push({
        role: "assistant",
        text: `${result.answer}\n\n**Sources:** ${result.sources.map((source) => source.name).join(", ")}`,
      });
      state.chatMessages = state.chatMessages.slice(-30);
      recordHistory({ type: "chat", title: message.slice(0, 75) });
      saveState();
      renderPage();
      const feed = document.getElementById("chat-feed");
      feed.scrollTop = feed.scrollHeight;
    } catch (error) {
      console.error("Document question could not be answered.", error);
      showToast(error.message || "The question could not be answered.");
      if (input) input.value = message;
    } finally {
      if (sendButton) sendButton.disabled = false;
    }
    return;
  }
  state.chatMessages.push({ role: "user", text: message });
  state.chatMessages.push({
    role: "assistant",
    text: "This is a preview, so I can't read your materials or generate a real AI answer yet. The next build step is connecting a secure AI backend. For now, you can try the sample quizzes to explore the learning flow.",
  });
  state.chatMessages = state.chatMessages.slice(-30);
  recordHistory({ type: "chat", title: message.slice(0, 75) });
  saveState();
  renderPage();
  const feed = document.getElementById("chat-feed");
  feed.scrollTop = feed.scrollHeight;
}

function saveQuizResult() {
  if (!quizSession || !quizSession.questions.length) return;
  const attempt = {
    topic: quizSession.topic,
    score: quizSession.score,
    total: quizSession.questions.length,
    at: Date.now(),
  };
  state.quizScores.push(attempt);
  state.quizScores = state.quizScores.slice(-500);
  recordHistory({ type: "quiz", title: `${attempt.topic} quiz`, score: attempt.score, total: attempt.total });
  saveState();
}

document.addEventListener("click", (event) => {
  const target = event.target.closest("[data-page], [data-action], [data-topic], [data-answer], [data-select-document], [data-remove-pending], [data-remove-file], [data-auth-toggle]");
  if (!target) return;
  if (target.hasAttribute("data-auth-toggle")) {
    authMode = authMode === "login" ? "signup" : "login";
    document.getElementById("auth-error").textContent = "";
    updateAuthForm();
  } else if (target.dataset.page) {
    event.preventDefault();
    setPage(target.dataset.page);
  } else if (target.dataset.action === "upload") {
    openFilePicker();
  } else if (target.dataset.action === "open-study") {
    setPage("study");
  } else if (target.dataset.action === "send-chat") {
    const input = document.getElementById("chat-input");
    if (!input) return;
    sendChatMessage(input.value);
  } else if (target.dataset.selectDocument) {
    const file = activeFiles().find((document) => document.id === target.dataset.selectDocument);
    if (!file) return;
    const selectedIndex = pendingFiles.findIndex((document) => document.id === file.id);
    if (selectedIndex >= 0) pendingFiles.splice(selectedIndex, 1);
    else if (pendingFiles.length >= 5) {
      showToast("Choose up to five study materials for one question.");
      return;
    } else pendingFiles.push(file);
    renderPage();
  } else if (target.dataset.action === "start-quiz") {
    startQuiz();
  } else if (target.dataset.action === "quiz-submit") {
    if (quizSession?.selected === null || quizSession?.selected === undefined) return;
    const question = quizSession.questions[quizSession.index];
    quizSession.submitted = true;
    if (quizSession.selected === question.answer) quizSession.score += 1;
    renderQuiz();
  } else if (target.dataset.action === "quiz-next") {
    if (!quizSession) return;
    quizSession.index += 1;
    quizSession.selected = null;
    quizSession.submitted = false;
    if (quizSession.index >= quizSession.questions.length) saveQuizResult();
    renderQuiz();
  } else if (target.dataset.action === "quiz-again") {
    startQuiz();
  } else if (target.dataset.action === "clear-demo") {
    if (!window.confirm("Clear your demo files, chats, history, and quiz scores from this browser?")) return;
    state = { files: [], history: [], quizScores: [], chatMessages: [] };
    quizSession = null;
    saveState();
    showToast("Demo data cleared from this browser.");
    renderPage();
  } else if (target.dataset.action === "clear-account-history") {
    if (!window.confirm("Clear your chat, quiz scores, and activity history from your account on all devices? Study materials in your private account will not be removed.")) return;
    state = { ...state, history: [], quizScores: [], chatMessages: [] };
    quizSession = null;
    saveState().then((synced) => {
      if (synced) showToast("Your learning history was cleared from your account.");
    });
    renderPage();
  } else if (target.dataset.action === "logout") {
    cloudSyncQueue.then(() => fetch("/api/auth/logout", { method: "POST" }))
      .then(async (response) => {
        if (!response.ok) throw new Error("Could not sign out. Please try again.");
        authenticatedUser = null;
        cloudStateReady = false;
        pendingFiles = [];
        serverDocuments = null;
        state = emptyState();
        showAuthGate();
      })
      .catch((error) => {
        console.error("Could not sign out.", error);
        showToast(error.message || "Could not sign out.");
      });
  } else if (target.dataset.topic) {
    const topic = target.dataset.topic;
    setPage("quiz");
    document.getElementById("quiz-topic").value = topicQuestion[topic] ? topic : "Cell biology";
    startQuiz();
  } else if (target.dataset.answer !== undefined) {
    if (!quizSession || quizSession.submitted) return;
    quizSession.selected = Number(target.dataset.answer);
    renderQuiz();
  } else if (target.dataset.removePending !== undefined) {
    pendingFiles.splice(Number(target.dataset.removePending), 1);
    renderPage();
  } else if (target.dataset.removeFile) {
    if (backendOnline) {
      fetch(`/api/documents/${encodeURIComponent(target.dataset.removeFile)}`, {
        method: "DELETE",
        headers: { Accept: "application/json" },
      })
        .then(async (response) => {
          if (!response.ok) {
            const result = await response.json();
            throw new Error(result.detail || `Delete failed (HTTP ${response.status}).`);
          }
          await refreshServerDocuments();
          renderPage();
          showToast("Study material removed from this computer.");
        })
        .catch((error) => {
          console.error("Could not remove study material.", error);
          showToast(error.message || "Could not remove study material.");
        });
    } else {
      state.files = state.files.filter((file) => file.id !== target.dataset.removeFile);
      saveState();
      renderPage();
      showToast("File removed from your demo library.");
    }
  }
});

document.addEventListener("keydown", (event) => {
  if (event.target.id !== "chat-input" || event.key !== "Enter" || event.shiftKey) return;
  event.preventDefault();
  sendChatMessage(event.target.value);
});

fileInput.addEventListener("change", () => {
  addFiles(fileInput.files);
  fileInput.value = "";
});

document.getElementById("auth-form").addEventListener("submit", submitAuthentication);

document.addEventListener("change", (event) => {
  if (event.target.id === "quiz-topic" && quizSession) quizSession = null;
});

document.addEventListener("dragover", (event) => {
  if (!event.target.closest("#upload-zone")) return;
  event.preventDefault();
  document.getElementById("upload-zone")?.classList.add("dragover");
});

document.addEventListener("dragleave", (event) => {
  if (!event.target.closest("#upload-zone")) return;
  document.getElementById("upload-zone")?.classList.remove("dragover");
});

document.addEventListener("drop", (event) => {
  if (!event.target.closest("#upload-zone")) return;
  event.preventDefault();
  document.getElementById("upload-zone")?.classList.remove("dragover");
  addFiles(event.dataTransfer.files);
});

document.getElementById("mobile-menu").addEventListener("click", (event) => {
  const sidebar = document.getElementById("sidebar");
  const isOpen = sidebar.classList.toggle("open");
  event.currentTarget.setAttribute("aria-expanded", String(isOpen));
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    document.getElementById("sidebar").classList.remove("open");
    document.getElementById("mobile-menu").setAttribute("aria-expanded", "false");
  }
});

updateCounts();
renderPage();
initializeBackend();
