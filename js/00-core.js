// redirectUri is auto-detected from the current page URL
// IMPORTANT: this exact URL must be added as a Redirect URI in Entra ID.
const TENANT_STORAGE_KEY = "intuneTenantConfigsV1";
const THEME_STORAGE_KEY = "intuneTheme";
const PANEL_STATE_STORAGE_KEY = "intunePanelStateV1";

const redirectUri = window.location.origin + window.location.pathname;

// DOM elements – globalnie
const tenantNameInput = document.getElementById("tenant-name");
const tenantTenantInput = document.getElementById("tenant-tenant");
const tenantClientIdInput = document.getElementById("tenant-client-id");
const tenantSaveButton = document.getElementById("tenant-save-button");
const tenantSelect = document.getElementById("tenant-select");
const tenantExportButton = document.getElementById("tenant-export-button");
const tenantImportButton = document.getElementById("tenant-import-button");
const tenantImportInput = document.getElementById("tenant-import-input");
const tenantDeleteButton = document.getElementById("tenant-delete-button");
const themeToggleBtn = document.getElementById("theme-toggle");
const helpButton = document.getElementById("help-button");


let tenantConfigs = [];
let currentTenantConfig = null;
let msalInstance = null;

// panel state (collapsed/expanded) persisted in localStorage
let panelState = {};

function loadPanelStateFromStorage() {
  try {
    const raw = localStorage.getItem(PANEL_STATE_STORAGE_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === "object") return parsed;
    return {};
  } catch (e) {
    console.warn("Failed to load panel state:", e);
    return {};
  }
}

function savePanelStateToStorage() {
  try {
    localStorage.setItem(PANEL_STATE_STORAGE_KEY, JSON.stringify(panelState));
  } catch (e) {
    console.warn("Failed to save panel state:", e);
  }
}

function loadTenantsFromStorage() {
  try {
    const raw = localStorage.getItem(TENANT_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) return parsed;
    return [];
  } catch (e) {
    console.warn("Failed to load tenants from localStorage:", e);
    return [];
  }
}

function saveTenantsToStorage() {
  try {
    localStorage.setItem(TENANT_STORAGE_KEY, JSON.stringify(tenantConfigs));
  } catch (e) {
    console.warn("Failed to save tenants to localStorage:", e);
  }
}

function renderTenantSelect() {
  tenantSelect.innerHTML = '<option value="">-- none --</option>';
  tenantConfigs.forEach((cfg, index) => {
    const opt = document.createElement("option");
    opt.value = String(index);
    opt.textContent = cfg.name + " (" + cfg.tenant + ")";
    tenantSelect.appendChild(opt);
  });

  if (currentTenantConfig) {
    const idx = tenantConfigs.findIndex(
      c => c.name === currentTenantConfig.name &&
           c.tenant === currentTenantConfig.tenant &&
           c.clientId === currentTenantConfig.clientId
    );
    if (idx >= 0) {
      tenantSelect.value = String(idx);
    }
  }
}

function initMsalForCurrentTenant() {
  const statusEl = document.getElementById("status");
  if (!currentTenantConfig) {
    msalInstance = null;
    if (statusEl) statusEl.textContent = "Status: select a tenant and sign in.";
    return;
  }

  const msalConfig = {
    auth: {
      clientId: currentTenantConfig.clientId,
      authority: "https://login.microsoftonline.com/" + currentTenantConfig.tenant,
      redirectUri: redirectUri
    },
    cache: {
      cacheLocation: "sessionStorage",
      storeAuthStateInCookie: false
    }
  };

  msalInstance = new msal.PublicClientApplication(msalConfig);
  setStatus("Selected tenant: " + currentTenantConfig.name + " (" + currentTenantConfig.tenant + ") – not signed in.");
}

tenantSaveButton.addEventListener("click", () => {
  const name = tenantNameInput.value.trim();
  const tenant = tenantTenantInput.value.trim();
  const clientId = tenantClientIdInput.value.trim();

  if (!name || !tenant || !clientId) {
    setMessage("Please fill in Name, Tenant and Client ID.", "error");
    return;
  }

  const existingIndex = tenantConfigs.findIndex(c => c.name === name);
  const cfg = { name, tenant, clientId };

  if (existingIndex >= 0) {
    tenantConfigs[existingIndex] = cfg;
  } else {
    tenantConfigs.push(cfg);
  }

  currentTenantConfig = cfg;
  saveTenantsToStorage();
  renderTenantSelect();
  initMsalForCurrentTenant();
  setMessage("Tenant saved: " + name, "success");
});

tenantSelect.addEventListener("change", () => {
  const val = tenantSelect.value;
  if (val === "") {
    currentTenantConfig = null;
    initMsalForCurrentTenant();
    return;
  }
  const idx = parseInt(val, 10);
  if (Number.isNaN(idx) || !tenantConfigs[idx]) {
    currentTenantConfig = null;
    initMsalForCurrentTenant();
    return;
  }
  currentTenantConfig = tenantConfigs[idx];
  initMsalForCurrentTenant();
  setMessage("Tenant selected: " + currentTenantConfig.name, "info");
});

tenantExportButton.addEventListener("click", () => {
  try {
    const dataStr = JSON.stringify(tenantConfigs, null, 2);
    const blob = new Blob([dataStr], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "intune-tenants.json";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    setMessage("Tenant configuration exported to intune-tenants.json.", "success");
  } catch (e) {
    console.error("JSON export error:", e);
    setMessage("Failed to export tenant configuration.", "error");
  }
});

tenantImportButton.addEventListener("click", () => {
  tenantImportInput.click();
});

tenantImportInput.addEventListener("change", (event) => {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const text = e.target.result;
      const parsed = JSON.parse(text);

      if (!Array.isArray(parsed)) {
        throw new Error("JSON file does not contain an array.");
      }

      const imported = parsed.filter(
        x =>
          x &&
          typeof x.name === "string" &&
          typeof x.tenant === "string" &&
          typeof x.clientId === "string"
      );

      tenantConfigs = imported;
      saveTenantsToStorage();

      currentTenantConfig = tenantConfigs[0] || null;
      renderTenantSelect();
      initMsalForCurrentTenant();

      setMessage("Imported " + tenantConfigs.length + " tenants from JSON (list replaced).", "success");
    } catch (err) {
      console.error("JSON import error:", err);
      setMessage("Failed to import tenant JSON file. Check the format.", "error");
    } finally {
      tenantImportInput.value = "";
    }
  };
  reader.readAsText(file);
});

tenantDeleteButton.addEventListener("click", () => {
  const val = tenantSelect.value;
  if (val === "") {
    setMessage("Select a tenant to delete.", "error");
    return;
  }
  const idx = parseInt(val, 10);
  if (Number.isNaN(idx) || !tenantConfigs[idx]) {
    setMessage("Cannot delete – invalid tenant selection.", "error");
    return;
  }
  const removed = tenantConfigs.splice(idx, 1)[0];
  saveTenantsToStorage();

  if (tenantConfigs.length > 0) {
    currentTenantConfig = tenantConfigs[0];
  } else {
    currentTenantConfig = null;
  }

  renderTenantSelect();
  initMsalForCurrentTenant();
  setMessage("Tenant deleted: " + removed.name, "success");
});

function applyTheme(theme) {
  const body = document.body;
  if (theme === "light") {
    body.classList.add("light-theme");
  } else {
    body.classList.remove("light-theme");
    theme = "dark";
  }
  try {
    localStorage.setItem(THEME_STORAGE_KEY, theme);
  } catch (e) {
    console.warn("Failed to save theme:", e);
  }
  if (themeToggleBtn) {
    themeToggleBtn.textContent =
      theme === "light" ? "Theme: Light ☀️" : "Theme: Dark 🌙";
  }
}

function toggleTheme() {
  const isLight = document.body.classList.contains("light-theme");
  applyTheme(isLight ? "dark" : "light");
}

if (themeToggleBtn) {
  themeToggleBtn.addEventListener("click", toggleTheme);
}

