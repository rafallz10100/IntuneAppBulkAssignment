// --- Help / About modal ---
let _helpModalWired = false;
let _helpLastFocus = null;

function buildHelpHtml() {
  const redirect = redirectUri || (window.location.origin + window.location.pathname);

  return `
    <h4>What is this app?</h4>
    <ul>
      <li>A small, single-page web tool (pure HTML/JS) for bulk adding/removing <b>Intune mobile app</b> assignments.</li>
      <li><b>Multi-tenant</b>: you can store multiple tenant configurations (Tenant ID / domain + Client ID) and switch between them.</li>
      <li>All calls go directly from your browser to <b>Microsoft Graph</b>. There is no backend server.</li>
      <li>Tenant configs are stored locally in <code>localStorage</code>; sign-in tokens are stored in <code>sessionStorage</code>.</li>
    </ul>

    <h4>Entra ID requirements (App Registration)</h4>
    <ul>
      <li>Create an <b>App registration</b> in Entra ID (single-tenant or multi-tenant—your choice).</li>
      <li>Add the <b>Single-page application (SPA)</b> platform and set the Redirect URI to:</li>
      <li><code>${escapeHtml(redirect)}</code></li>
      <li>Copy the <b>Application (client) ID</b> and paste it into the <i>Client ID</i> field in this app.</li>
      <li>If needed, grant <b>Admin consent</b> for the permissions below.</li>
    </ul>

    <h4>Required permissions (Microsoft Graph — Delegated)</h4>
    <ul>
      <li><code>User.Read</code> — basic info about the signed-in user.</li>
      <li><code>Group.Read.All</code> — read groups for assignment targeting and lookups.</li>
      <li><code>DeviceManagementApps.ReadWrite.All</code> — read Intune apps and add/remove assignments.</li>
      <li><code>DeviceManagementConfiguration.Read.All</code> — read assignment filters.</li>
    </ul>

    <h4>Security notes</h4>
    <ul>
      <li>Your account still needs the right Intune/Entra roles (for example, <i>Intune Administrator</i>).</li>
      <li>If you can’t see filters or groups, it’s usually missing Graph permissions/consent (<code>DeviceManagementConfiguration.Read.All</code> and/or <code>Group.Read.All</code>).</li>
      <li>The Redirect URI must <b>exactly</b> match the URL where you host this page.</li>
    </ul>
  `;
}

function wireHelpModal() {
  if (_helpModalWired) return;
  _helpModalWired = true;

  const overlay = document.getElementById("help-overlay");
  const btnClose = document.getElementById("help-close");

  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) closeHelpModal();
  });

  btnClose.addEventListener("click", closeHelpModal);

  document.addEventListener("keydown", (e) => {
    if (overlay.classList.contains("hidden")) return;
    if (e.key === "Escape") {
      e.preventDefault();
      closeHelpModal();
    }
  });
}

function openHelpModal() {
  wireHelpModal();

  const overlay = document.getElementById("help-overlay");
  const bodyEl = document.getElementById("help-body");
  const btnClose = document.getElementById("help-close");

  if (bodyEl) bodyEl.innerHTML = buildHelpHtml();

  _helpLastFocus = document.activeElement;
  overlay.classList.remove("hidden");

  // Focus a safe default
  setTimeout(() => btnClose && btnClose.focus(), 0);
}

function closeHelpModal() {
  const overlay = document.getElementById("help-overlay");
  overlay.classList.add("hidden");

  if (_helpLastFocus && typeof _helpLastFocus.focus === "function") {
    try { _helpLastFocus.focus(); } catch (e) {}
  }
  _helpLastFocus = null;
}

if (helpButton) {
  helpButton.addEventListener("click", openHelpModal);
}


const statusEl = document.getElementById("status");
const messageBarEl = document.getElementById("message-bar");
const outputEl = document.getElementById("output");
const loginButton = document.getElementById("login-button");
const logoutButton = document.getElementById("logout-button");
const appsTableWrapper = document.getElementById("apps-table-wrapper");
const appsSummaryEl = document.getElementById("apps-summary");


// --- Assignment context menu (PPM) ---
let ctxMenuEl = null;
let ctxRemoveBtn = null;
let ctxState = { appId: "", assignmentId: "", label: "" };

function ensureContextMenu() {
  if (!ctxMenuEl) ctxMenuEl = document.getElementById("context-menu");
  if (!ctxRemoveBtn) ctxRemoveBtn = document.getElementById("ctx-remove-assignment");

  if (!ctxMenuEl) {
    ctxMenuEl = document.createElement("div");
    ctxMenuEl.id = "context-menu";
    ctxMenuEl.className = "ctx-menu hidden";
    ctxMenuEl.setAttribute("role", "menu");
    ctxMenuEl.setAttribute("aria-label", "Assignment actions");

    ctxRemoveBtn = document.createElement("button");
    ctxRemoveBtn.id = "ctx-remove-assignment";
    ctxRemoveBtn.className = "ctx-menu-item danger";
    ctxRemoveBtn.type = "button";
    ctxRemoveBtn.textContent = "Remove assignment";

    ctxMenuEl.appendChild(ctxRemoveBtn);
    document.body.appendChild(ctxMenuEl);
  }

  if (ctxMenuEl && ctxRemoveBtn && !ctxRemoveBtn.__wired) {
    ctxRemoveBtn.__wired = true;
    ctxRemoveBtn.addEventListener("click", async (e) => {
      e.preventDefault();

      const appId = ctxState.appId;
      const assignmentId = ctxState.assignmentId;
      const label = ctxState.label;

      hideContextMenu();

      if (!appId || !assignmentId) return;

      const appName = getAppNameById(appId);
      const ok = await showConfirmModal({
        title: "Remove assignment?",
        message: `App: ${appName}\n\n${label}`,
        okText: "Remove",
        cancelText: "Cancel"
      });
      if (!ok) return;

      await removeSingleAssignment(appId, assignmentId);
    });
  }
}

function hideContextMenu() {
  if (!ctxMenuEl) ctxMenuEl = document.getElementById("context-menu");
  if (ctxMenuEl) ctxMenuEl.classList.add("hidden");
  ctxState = { appId: "", assignmentId: "", label: "" };
}


let _confirmModalWired = false;

function wireConfirmModal() {
  if (_confirmModalWired) return;
  _confirmModalWired = true;

  const overlay = document.getElementById("confirm-overlay");
  const btnOk = document.getElementById("confirm-ok");
  const btnCancel = document.getElementById("confirm-cancel");

  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) closeConfirmModal(false);
  });

  btnOk.addEventListener("click", () => closeConfirmModal(true));
  btnCancel.addEventListener("click", () => closeConfirmModal(false));

  document.addEventListener("keydown", (e) => {
    if (overlay.classList.contains("hidden")) return;
    if (e.key === "Escape") {
      e.preventDefault();
      closeConfirmModal(false);
    }
    if (e.key === "Enter") {
      // Enter confirms (but avoid interfering with textarea)
      const tag = (document.activeElement && document.activeElement.tagName || "").toLowerCase();
      if (tag !== "textarea" && tag !== "input") {
        e.preventDefault();
        closeConfirmModal(true);
      }
    }
  });
}

let _confirmResolve = null;
let _confirmLastFocus = null;

function showConfirmModal({ title = "Confirm", message = "", okText = "OK", cancelText = "Cancel" } = {}) {
  wireConfirmModal();

  const overlay = document.getElementById("confirm-overlay");
  const titleEl = document.getElementById("confirm-title");
  const msgEl = document.getElementById("confirm-message");
  const btnOk = document.getElementById("confirm-ok");
  const btnCancel = document.getElementById("confirm-cancel");

  titleEl.textContent = title;
  msgEl.textContent = message;
  btnOk.textContent = okText;
  btnCancel.textContent = cancelText;

  _confirmLastFocus = document.activeElement;
  overlay.classList.remove("hidden");

  // Focus a safe default
  setTimeout(() => btnCancel.focus(), 0);

  return new Promise((resolve) => {
    _confirmResolve = resolve;
  });
}

function closeConfirmModal(result) {
  const overlay = document.getElementById("confirm-overlay");
  overlay.classList.add("hidden");

  const resolve = _confirmResolve;
  _confirmResolve = null;
  if (typeof resolve === "function") resolve(!!result);

  if (_confirmLastFocus && typeof _confirmLastFocus.focus === "function") {
    try { _confirmLastFocus.focus(); } catch (e) {}
  }
  _confirmLastFocus = null;
}



function showContextMenu(x, y, state) {
  ensureContextMenu();
  ctxState = state;

  ctxMenuEl.style.left = x + "px";
  ctxMenuEl.style.top = y + "px";
  ctxMenuEl.classList.remove("hidden");

  // Keep inside viewport
  requestAnimationFrame(() => {
    const r = ctxMenuEl.getBoundingClientRect();
    let left = x;
    let top = y;

    if (left + r.width > window.innerWidth - 8) left = Math.max(8, window.innerWidth - r.width - 8);
    if (top + r.height > window.innerHeight - 8) top = Math.max(8, window.innerHeight - r.height - 8);

    ctxMenuEl.style.left = left + "px";
    ctxMenuEl.style.top = top + "px";
  });
}

async function removeSingleAssignment(appId, assignmentId) {
  const account = getActiveAccount();
  if (!account || !msalInstance) {
    setStatus("Not signed in.");
    setMessage("Sign in first.", "error");
    return;
  }

  setUiBusy(true);
  setStatus("Removing assignment…");
  setMessage("", "info");

  try {
    let tokenResponse;
    try {
      tokenResponse = await msalInstance.acquireTokenSilent({ ...loginRequest, account });
    } catch (e) {
      tokenResponse = await msalInstance.acquireTokenPopup({ ...loginRequest, account });
    }

    const accessToken = tokenResponse.accessToken;
    const url =
      `https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps/${encodeURIComponent(appId)}/assignments/${encodeURIComponent(assignmentId)}`;

    const res = await fetch(url, {
      method: "DELETE",
      headers: { "Authorization": `Bearer ${accessToken}` }
    });

    const body = await res.text();

    if (!res.ok) {
      console.error("Error deleting assignment", assignmentId, "for app", appId, res.status, body);
      logOutput(`Error deleting assignment for app ${getAppNameById(appId)} (assignmentId=${assignmentId}) (${res.status}):\n${body}`);
      setMessage("Failed to remove assignment (see Output).", "error");
      setStatus("Error removing assignment.");
      return;
    }

    // Update local cache + re-render
    if (cachedAssignmentsMap && cachedAssignmentsMap[appId]) {
      cachedAssignmentsMap[appId] = cachedAssignmentsMap[appId].filter(a => a && a.id !== assignmentId);
    }
    renderApps(cachedApps || [], cachedAssignmentsMap || {}, cachedGroupMap || {});
    setMessage("Assignment removed.", "success");
    setStatus("Assignment removed.");
    logOutput(`Removed assignmentId=${assignmentId} from app ${getAppNameById(appId)}.`);
  } catch (e) {
    console.error("Exception deleting assignment", assignmentId, "for app", appId, e);
    logOutput(`Exception while deleting assignment for app ${getAppNameById(appId)} (assignmentId=${assignmentId}):\n${String(e)}`);
    setMessage("Failed to remove assignment (see Output).", "error");
    setStatus("Error removing assignment.");
  } finally {
    setUiBusy(false);
  }
}

function wireAssignmentContextMenu() {
  ensureContextMenu();

  document.addEventListener("click", () => hideContextMenu());
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") hideContextMenu(); });
  document.addEventListener("scroll", () => hideContextMenu(), true);
  window.addEventListener("resize", () => hideContextMenu());

  if (appsTableWrapper && !appsTableWrapper.__ctxWired) {
    appsTableWrapper.__ctxWired = true;
    appsTableWrapper.addEventListener("contextmenu", (e) => {
      const el = e.target.closest(".assignment-item");
      if (!el) return;

      const assignmentId = el.dataset.assignmentId || "";
      const appId = el.dataset.appId || "";
      if (!assignmentId || !appId) return; // only for real assignments (with id)

      e.preventDefault();
      showContextMenu(e.clientX, e.clientY, {
        appId,
        assignmentId,
        label: (el.textContent || "").trim()
      });
    });
  }
}

wireAssignmentContextMenu();

const bulkTargetTypeSelect = document.getElementById("bulk-target-type");
const bulkGroupNameWrapper = document.getElementById("bulk-group-name-wrapper");
const bulkAssignButton = document.getElementById("bulk-assign-button");
const bulkDeleteButton = document.getElementById("bulk-delete-button");
const bulkIntentSelect = document.getElementById("bulk-intent");
const bulkGroupNameInput = document.getElementById("bulk-group-name");


const bulkGroupModeWrapper = document.getElementById("bulk-group-mode-wrapper");
const bulkGroupModeSelect = document.getElementById("bulk-group-mode");
// Assignment filter controls (Intune Assignment Filters)
const bulkAssignmentFilterModeWrapper = document.getElementById("bulk-assignment-filter-mode-wrapper");
const bulkAssignmentFilterModeSelect = document.getElementById("bulk-assignment-filter-mode");
const bulkAssignmentFilterNameWrapper = document.getElementById("bulk-assignment-filter-name-wrapper");
const bulkAssignmentFilterNameInput = document.getElementById("bulk-assignment-filter-name");

// Suggestions cache: displayName -> id
const assignmentFilterSuggestionCache = {};
// Keep id -> displayName mapping for table/export/logging
let cachedAssignmentFilterMap = {};



