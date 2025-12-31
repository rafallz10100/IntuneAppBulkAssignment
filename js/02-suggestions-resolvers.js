// --- Assignment filter cache helper (id/name + case-insensitive) ---
function upsertAssignmentFilterIntoCaches(item) {
  if (!item || !item.id) return;
  const name = String(item.displayName || "").trim();

  // id -> displayName
  cachedAssignmentFilterMap[item.id] = name || cachedAssignmentFilterMap[item.id] || item.id;

  // displayName -> id (also store lowercase key so "name" lookup is case-insensitive)
  if (name) {
    assignmentFilterSuggestionCache[name] = item.id;
    assignmentFilterSuggestionCache[name.toLowerCase()] = item.id;
  }
}

// --- Group name suggestions (typeahead via Graph) ---
const groupNameSuggestionCache = {}; // displayName -> id (for quick selection)

function escapeODataString(value) {
  return String(value || "").replace(/'/g, "''");
}


function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function debounce(fn, delayMs) {
  let t = null;
  return (...args) => {
    if (t) clearTimeout(t);
    t = setTimeout(() => fn(...args), delayMs);
  };
}

async function getAccessTokenForGraph() {
  if (!msalInstance) throw new Error("MSAL not initialized.");
  const account = getActiveAccount();
  if (!account) throw new Error("Not signed in.");
  try {
    const tokenResponse = await msalInstance.acquireTokenSilent({
      ...loginRequest,
      account
    });
    return tokenResponse.accessToken;
  } catch (e) {
    const tokenResponse = await msalInstance.acquireTokenPopup({
      ...loginRequest,
      account
    });
    return tokenResponse.accessToken;
  }
}

async function queryGroupsForSuggestions(prefix, limit = 20, accessTokenOverride = null) {
  const q = (prefix || "").trim();
  if (q.length < 2) return [];

  const token = (accessTokenOverride ? accessTokenOverride : await getAccessTokenForGraph());

  const filter = `startswith(displayName,'${escapeODataString(q)}')`;
  const url =
    `https://graph.microsoft.com/v1.0/groups` +
    `?$select=id,displayName&$top=${limit}&$filter=${encodeURIComponent(filter)}`;

  const res = await fetch(url, {
    method: "GET",
    headers: {
      "Authorization": `Bearer ${token}`,
      "Accept": "application/json"
    }
  });

  const text = await res.text();
  let json = null;
  try { json = JSON.parse(text); } catch (e) {}

  if (!res.ok) {
    logOutput("Group suggestions error:\\n" + text);
    return [];
  }

  return ((json && json.value) || [])
    .filter(x => x && x.id && x.displayName)
    .map(x => ({ id: x.id, displayName: x.displayName }));
}

function updateGroupDatalist(groups) {
  const dl = document.getElementById("group-name-suggestions");
  if (!dl) return;
  dl.innerHTML = "";

  (groups || []).slice(0, 30).forEach(g => {
    groupNameSuggestionCache[g.displayName] = g.id;
    
    // keep id -> name mapping for reports/table/export
    cachedGroupMap[g.id] = g.displayName;
const opt = document.createElement("option");
    opt.value = g.displayName;
    dl.appendChild(opt);
  });
}

const debouncedSuggestGroups = debounce(async () => {
  try {
    const q = bulkGroupNameInput.value;
    if (!q || q.trim().length < 2) {
      updateGroupDatalist([]);
      return;
    }
    const items = await queryGroupsForSuggestions(q.trim(), 20);
    updateGroupDatalist(items);
  } catch (e) {
    // ignore when not signed in / token not available
  }
}, 250);

if (bulkGroupNameInput) {
  bulkGroupNameInput.addEventListener("input", () => {
    if (bulkTargetTypeSelect && bulkTargetTypeSelect.value !== "group") return;
    debouncedSuggestGroups();
  });
}
// --- end group suggestions ---

// --- Assignment filter suggestions (typeahead) ---
// Uses: GET /deviceManagement/assignmentFilters (v1.0 preferred; fallback to beta)
let cachedAssignmentFiltersList = [];
let lastAssignmentFilterLookupError = null;
let assignmentFilterPermissionErrorShown = false;

function looksLikeGuid(value) {
  return /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(String(value || "").trim());
}

async function fetchAllAssignmentFilters(accessTokenOverride) {
  if (cachedAssignmentFiltersList && cachedAssignmentFiltersList.length) return cachedAssignmentFiltersList;

  lastAssignmentFilterLookupError = null;

  const token = await getAccessTokenForGraph();

  logOutput('Loading assignment filters list…');

  // Some tenants behave oddly between v1.0 and beta for assignmentFilters.
// In practice, beta is the safest first try, then fall back to v1.0.
  const baseUrls = [
    { label: "beta", url: "https://graph.microsoft.com/beta/deviceManagement/assignmentFilters?$select=id,displayName&$top=100" },
    { label: "v1.0", url: "https://graph.microsoft.com/v1.0/deviceManagement/assignmentFilters?$select=id,displayName&$top=100" }
  ];

  for (const base of baseUrls) {
    let nextUrl = base.url;
    let all = [];
    let lastRes = null;
    let lastText = "";

    while (nextUrl) {
      const res = await fetch(nextUrl, {
        method: "GET",
        headers: {
          "Authorization": `Bearer ${token}`,
          "Accept": "application/json"
        }
      });

      lastRes = res;
      lastText = await res.text();

      let json = null;
      try { json = lastText ? JSON.parse(lastText) : null; } catch (e) { json = null; }

      if (!res.ok) break;

      const items = ((json && json.value) || [])
        .filter(x => x && x.id && x.displayName)
        .map(x => ({ id: x.id, displayName: x.displayName }));

      all.push(...items);

      nextUrl = (json && json["@odata.nextLink"]) ? json["@odata.nextLink"] : null;
    }

    // Successful response (even if empty)
    if (lastRes && lastRes.ok) {
      // De-dup by id
      const byId = {};
      for (const it of all) byId[it.id] = it;
      const list = Object.values(byId);

      if (list.length > 0) {
        cachedAssignmentFiltersList = list;
        (cachedAssignmentFiltersList || []).forEach(f => upsertAssignmentFilterIntoCaches(f));
        logOutput(`Assignment filters loaded from ${base.label}: ${cachedAssignmentFiltersList.length}`);
        return cachedAssignmentFiltersList;
      }

      // Empty list on this endpoint → try next (e.g., beta)
      logOutput(`Assignment filters list is empty on ${base.label} – trying next endpoint...`);
      continue;
    }

    // Error on this endpoint
    if (lastRes) {
      lastAssignmentFilterLookupError = { status: lastRes.status, url: base.url, body: lastText };
      logOutput(`Assignment filters request failed (${lastRes.status}) for ${base.url}:\n` + lastText);

      // If this API version isn't available, try the next one
      if (lastRes.status === 404) continue;

      // Other errors: stop early
      break;
    }
  }

  return [];
}

async function suggestAssignmentFilters(prefix, limit = 20) {
  const q = (prefix || "").trim().toLowerCase();
  if (q.length < 2) return [];

  const all = await fetchAllAssignmentFilters();
  return (all || [])
    .filter(f => (f.displayName || "").toLowerCase().includes(q))
    .slice(0, limit);
}

function updateAssignmentFilterDatalist(filters) {
  const dl = document.getElementById("assignment-filter-suggestions");
  if (!dl) return;
  dl.innerHTML = "";

  (filters || []).slice(0, 30).forEach(f => {
    upsertAssignmentFilterIntoCaches(f);
    const opt = document.createElement("option");
    opt.value = f.displayName;
    dl.appendChild(opt);
  });
}

const debouncedSuggestAssignmentFilters = debounce(async () => {
  try {
    if (!bulkAssignmentFilterModeSelect || bulkAssignmentFilterModeSelect.value === "none") {
      updateAssignmentFilterDatalist([]);
      return;
    }

    const q = bulkAssignmentFilterNameInput ? bulkAssignmentFilterNameInput.value : "";
    if (!q || q.trim().length < 2) {
      updateAssignmentFilterDatalist([]);
      return;
    }

    const items = await suggestAssignmentFilters(q.trim(), 20);
    updateAssignmentFilterDatalist(items);

    if (!items.length && lastAssignmentFilterLookupError && (lastAssignmentFilterLookupError.status === 401 || lastAssignmentFilterLookupError.status === 403)) {
      if (!assignmentFilterPermissionErrorShown) {
        assignmentFilterPermissionErrorShown = true;
        setMessage(
          "Couldn't read the assignment filter list (HTTP " + lastAssignmentFilterLookupError.status + "). " +
          "Most often this means you're missing permission/admin consent for DeviceManagementConfiguration.Read.All. " +
          "You can also paste the filter ID (GUID) instead of the name.",
          "error"
        );
      }
    }
  } catch (e) {
    // Keep quiet in UI, but log for debugging
    logOutput("Assignment filter suggestions exception:\n" + String(e));
}
}, 250);

if (bulkAssignmentFilterNameInput) {
  bulkAssignmentFilterNameInput.addEventListener("input", () => {
    debouncedSuggestAssignmentFilters();
  });
}
if (bulkAssignmentFilterModeSelect) {
  bulkAssignmentFilterModeSelect.addEventListener("change", () => {
    updateAssignmentFilterDatalist([]);
    if (bulkAssignmentFilterNameInput && bulkAssignmentFilterModeSelect.value === "none") {
      bulkAssignmentFilterNameInput.value = "";
    }
    updateAssignmentFilterControlsVisibility();
  });
}
// --- end assignment filter suggestions ---
function clearAssignmentFilterCaches() {
  cachedAssignmentFiltersList = [];
  lastAssignmentFilterLookupError = null;
  assignmentFilterPermissionErrorShown = false;

  // clear suggestion cache (const object)
  try {
    for (const k of Object.keys(assignmentFilterSuggestionCache || {})) delete assignmentFilterSuggestionCache[k];
  } catch (e) {}

  // reset id->name map
  cachedAssignmentFilterMap = {};

  const dl = document.getElementById("assignment-filter-suggestions");
  if (dl) dl.innerHTML = "";
}



const clearLogButton = document.getElementById("clear-log-button");
const platformFilterSelect = document.getElementById("platform-filter");
const exportXlsxButton = document.getElementById("export-xlsx-button");
const nameFilterInput = document.getElementById("name-filter");

const loginRequest = {
  scopes: [
    "User.Read",
    "DeviceManagementApps.ReadWrite.All",
    "DeviceManagementConfiguration.Read.All",
    "Group.Read.All"
  ]
};

const graphAppsBaseUrl =
  "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?$top=50";

let cachedApps = [];
let cachedAssignmentsMap = {};
let cachedGroupMap = {};

// --- Group ID -> displayName resolver (for reports/table) ---
const pendingGroupNameLookups = new Set();

async function fetchGroupNamesByIds(accessToken, ids) {
  const unique = Array.from(new Set((ids || []).filter(Boolean)));
  const results = {};

  const chunkSize = 100;
  for (let i = 0; i < unique.length; i += chunkSize) {
    const chunk = unique.slice(i, i + chunkSize);
    const url = "https://graph.microsoft.com/v1.0/directoryObjects/getByIds";
    const body = {
      ids: chunk,
      types: ["group"]
    };

    const res = await fetch(url, {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "application/json",
        "Content-Type": "application/json"
      },
      body: JSON.stringify(body)
    });

    const text = await res.text();
    let json = null;
    try { json = text ? JSON.parse(text) : null; } catch (e) {}

    if (!res.ok) {
      logOutput("Group name resolve error:\\n" + text);
      continue;
    }

    const arr = (json && json.value) ? json.value : [];
    for (const obj of arr) {
      if (obj && obj.id && obj.displayName) {
        results[obj.id] = obj.displayName;
      }
    }
  }

  return results;
}

function collectGroupIdsFromAssignmentsMap(assignmentsMap) {
  const ids = [];
  for (const asgs of Object.values(assignmentsMap || {})) {
    for (const a of (asgs || [])) {
      const t = (a && a.target) ? a.target : null;
      const tt = ((t && t["@odata.type"]) || "").toLowerCase();
      if (tt.includes("groupassignmenttarget") || tt.includes("exclusiongroupassignmenttarget")) {
        const gid = t.groupId;
        if (gid) ids.push(gid);
      }
    }
  }
  return ids;
}

async function resolveMissingGroupNamesForAssignments(accessToken, assignmentsMap) {
  const allIds = collectGroupIdsFromAssignmentsMap(assignmentsMap);
  const missing = allIds.filter(id => !cachedGroupMap[id]);
  if (!missing.length) return;

  const resolved = await fetchGroupNamesByIds(accessToken, missing);
  let changed = false;
  for (const [id, name] of Object.entries(resolved)) {
    if (!cachedGroupMap[id]) {
      cachedGroupMap[id] = name;
      changed = true;
    }
  }
  if (changed) {
    renderApps(cachedApps, cachedAssignmentsMap, cachedGroupMap);
  }
}

// Called during rendering if we encounter unknown groupId – schedules background resolve and re-render.
function scheduleGroupNameResolve(groupId) {
  if (!groupId) return;
  if (cachedGroupMap[groupId]) return;
  if (pendingGroupNameLookups.has(groupId)) return;

  pendingGroupNameLookups.add(groupId);

  (async () => {
    try {
      const token = await getAccessTokenForGraph();
      const resolved = await fetchGroupNamesByIds(token, [groupId]);
      if (resolved[groupId]) {
        cachedGroupMap[groupId] = resolved[groupId];
        renderApps(cachedApps, cachedAssignmentsMap, cachedGroupMap);
      }
    } catch (e) {
      // ignore
    } finally {
      pendingGroupNameLookups.delete(groupId);
    }
  })();
}
// --- end resolver ---
let bulkConfirmState = null;
let bulkDeleteConfirmState = null;
let suppressNextRefreshMessage = false;
let uiBusy = false;

function applyDisabledState() {
  const disabled = uiBusy;
  [loginButton, logoutButton, bulkAssignButton, bulkDeleteButton, exportXlsxButton].forEach(btn => {
    if (btn) btn.disabled = disabled;
  });
  document.querySelectorAll("input[type='checkbox']").forEach(cb => {
    cb.disabled = disabled;
  });
}

function setUiBusy(busy) {
  uiBusy = busy;
  applyDisabledState();
}

