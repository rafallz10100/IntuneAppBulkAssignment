async function handleBulkAssignClick() {
  const selectedAppIds = getSelectedAppIds();
  const intent = bulkIntentSelect.value;
  const targetType = bulkTargetTypeSelect.value;
  const groupName = bulkGroupNameInput.value.trim();

  
  const groupMode = (targetType === "group" && bulkGroupModeSelect) ? bulkGroupModeSelect.value : "include";

  const filterMode = bulkAssignmentFilterModeSelect ? bulkAssignmentFilterModeSelect.value : "none";
  const filterName = bulkAssignmentFilterNameInput ? bulkAssignmentFilterNameInput.value.trim() : "";

  if (!selectedAppIds.length) {
    setMessage("Select at least one app first.", "error");
    return;
  }

  const summaryFilter = (filterMode && filterMode !== "none")
    ? ` (filter: ${filterMode} ${filterName || "(no filter name)"})`
    : "";

  const summaryTarget =
    (targetType === "group"
      ? ((groupMode === "exclude" ? "EXCLUDE: " : "") + (groupName || "(no group name)"))
      : (targetType === "allDevices" ? "All devices" : "All users")) + summaryFilter;


  const newState = {
    intent,
    targetType,
    groupName,
    groupMode,
    filterMode,
    filterName,
    count: selectedAppIds.length
  };

  const prevStateJson = bulkConfirmState ? JSON.stringify(bulkConfirmState) : null;
  const newStateJson = JSON.stringify(newState);

  if (!bulkConfirmState || prevStateJson !== newStateJson) {
    bulkConfirmState = newState;
    setMessage(
      `Prepared to ADD assignments: "${intent}" → "${summaryTarget}" ` +
      `for ${selectedAppIds.length} apps. Click "Add assignment" again to confirm.`,
      "info"
    );
    return;
  }

  bulkConfirmState = null;

  try {
    const account = getActiveAccount();
    if (!account) {
      setMessage("Sign in and load the app list first.", "error");
      return;
    }

    setUiBusy(true);

    const tokenResponse = await msalInstance.acquireTokenSilent({
      ...loginRequest,
      account
    });

    const accessToken = tokenResponse.accessToken;

    let target;
    let targetLabel;

    if (targetType === "allDevices") {
      target = { "@odata.type": "#microsoft.graph.allDevicesAssignmentTarget" };
      targetLabel = "All devices";
    } else if (targetType === "allUsers") {
      target = { "@odata.type": "#microsoft.graph.allLicensedUsersAssignmentTarget" };
      targetLabel = "All users";
    } else {
      if (!groupName) {
        setMessage("Enter a group name.", "error");
        setUiBusy(false);
        return;
      }

      let groupId = null;

      // Use suggestions cache first (if user picked from dropdown)
      if (groupNameSuggestionCache && groupNameSuggestionCache[groupName]) {
        groupId = groupNameSuggestionCache[groupName];
      
        
        if (groupId) cachedGroupMap[groupId] = groupName;
if (groupId) cachedGroupMap[groupId] = groupName;
}

      if (!groupId) {
        for (const [gid, gname] of Object.entries(cachedGroupMap)) {
if ((gname || "").toLowerCase() === groupName.toLowerCase()) {
          groupId = gid;
          break;
        }
        }
      }

      if (!groupId) {
        const resolved = await resolveGroupByName(groupName, accessToken);
        if (!resolved) {
          setMessage("Group not found: " + groupName, "error");
          setUiBusy(false);
          return;
        }
        groupId = resolved.id;
        cachedGroupMap[groupId] = resolved.displayName || groupName;
      }

      target = {
        "@odata.type": groupMode === "exclude"
          ? "#microsoft.graph.exclusionGroupAssignmentTarget"
          : "#microsoft.graph.groupAssignmentTarget",
        groupId
      };
      targetLabel = cachedGroupMap[groupId] || groupName;
    }

    // Apply assignment filter (optional)
    // NOTE: Intune does NOT support assignment filters on exclusion assignments.
    // If user selected Exclude group, block using filters to avoid Graph BadRequest.
    if (target && String(target["@odata.type"] || "").toLowerCase().includes("exclusiongroupassignmenttarget")) {
      const normalizedFilterModeGuard = normalizeFilterMode(filterMode);
      if (normalizedFilterModeGuard !== "none") {
        setMessage(
          "Exclude group does not support Assignment Filters (Intune/Graph). " +
          "Set 'No filter' or use Include + filter (invert the logic).",
          "error"
        );
        setUiBusy(false);
        return;
      }
    }

    const normalizedFilterMode = normalizeFilterMode(filterMode);
    if (normalizedFilterMode !== "none") {
      if (!filterName) {
        setMessage("Enter an assignment filter name (or set filter mode to 'No filter').", "error");
        setUiBusy(false);
        return;
      }

      let filterId = null;

      // From suggestions cache (exact match from datalist)
      if (assignmentFilterSuggestionCache && assignmentFilterSuggestionCache[filterName]) {
        filterId = assignmentFilterSuggestionCache[filterName];
      }

      // Case-insensitive match from cached map
      if (!filterId) {
        const wanted = filterName.toLowerCase();
        for (const [fid, fname] of Object.entries(cachedAssignmentFilterMap || {})) {
          if ((fname || "").toLowerCase() === wanted) {
            filterId = fid;
            break;
          }
        }
      }

      if (!filterId) {
        const resolved = await resolveAssignmentFilterByName(filterName, accessToken);
        if (!resolved) {
          if (lastAssignmentFilterLookupError && (lastAssignmentFilterLookupError.status === 401 || lastAssignmentFilterLookupError.status === 403)) {
            setMessage(
              "Couldn't read assignment filters (HTTP " + lastAssignmentFilterLookupError.status + "). " +
              "Check you have permission/admin consent for DeviceManagementConfiguration.Read.All. " +
              "Tip: you can paste the filter ID (GUID) instead of the name.",
              "error"
            );
          } else {
            setMessage("Assignment filter not found: " + filterName + ". Tip: you can paste the filter ID (GUID).", "error");
          }
          setUiBusy(false);
          return;
        }
        if (resolved.ambiguous) {
          const list = (resolved.matches || []).map(x => x.displayName).filter(Boolean).slice(0, 8);
          setMessage(
            "Filter name is ambiguous. Please refine it or pick from the suggestion list. Matches: " +
              (list.length ? list.join(" · ") : "(no names)"),
            "error"
          );
          setUiBusy(false);
          return;
        }

        filterId = resolved.id;
        cachedAssignmentFilterMap[filterId] = resolved.displayName || filterName;
      }

      target.deviceAndAppManagementAssignmentFilterId = filterId;
      target.deviceAndAppManagementAssignmentFilterType = normalizedFilterMode;

      const fname = cachedAssignmentFilterMap[filterId] || filterName;
      targetLabel = `${targetLabel} (filter: ${normalizedFilterMode} ${fname})`;
    }

    await bulkAssignToApps(accessToken, intent, targetType, target, targetLabel);
  } catch (err) {
    console.error("Error during bulk assignment:", err);
    setStatus("Error during bulk assignment – see console.");
    setMessage("An error occurred during bulk assignment.", "error");
    logOutput("Error during bulk assignment:\n" + String(err));
  } finally {
    setUiBusy(false);
  }
}

async function handleBulkDeleteClick() {
  const selectedAppIds = getSelectedAppIds();
  const intent = bulkIntentSelect.value;
  const targetType = bulkTargetTypeSelect.value;
  const groupName = bulkGroupNameInput.value.trim();

  
  const groupMode = (targetType === "group" && bulkGroupModeSelect) ? bulkGroupModeSelect.value : "include";

  const filterMode = bulkAssignmentFilterModeSelect ? bulkAssignmentFilterModeSelect.value : "none";
  const filterName = bulkAssignmentFilterNameInput ? bulkAssignmentFilterNameInput.value.trim() : "";

  if (!selectedAppIds.length) {
    setMessage("Select at least one app first.", "error");
    return;
  }

  const summaryFilter = (filterMode && filterMode !== "none")
    ? ` (filter: ${filterMode} ${filterName || "(no filter name)"})`
    : "";

  const summaryTarget =
    (targetType === "group"
      ? ((groupMode === "exclude" ? "EXCLUDE: " : "") + (groupName || "(no group name)"))
      : (targetType === "allDevices" ? "All devices" : "All users")) + summaryFilter;


  const newState = {
    intent,
    targetType,
    groupName,
    groupMode,
    filterMode,
    filterName,
    count: selectedAppIds.length
  };

  const prevStateJson = bulkDeleteConfirmState ? JSON.stringify(bulkDeleteConfirmState) : null;
  const newStateJson = JSON.stringify(newState);

  if (!bulkDeleteConfirmState || prevStateJson !== newStateJson) {
    bulkDeleteConfirmState = newState;
    setMessage(
      `Prepared to REMOVE assignments: "${intent}" → "${summaryTarget}" ` +
      `for ${selectedAppIds.length} apps. Click "Remove assignment" again to confirm.`,
      "info"
    );
    return;
  }

  bulkDeleteConfirmState = null;

  try {
    const account = getActiveAccount();
    if (!account) {
      setMessage("Sign in and load the app list first.", "error");
      return;
    }

    setUiBusy(true);

    const tokenResponse = await msalInstance.acquireTokenSilent({
      ...loginRequest,
      account
    });

    const accessToken = tokenResponse.accessToken;

    let target;
    let targetLabel;

    if (targetType === "allDevices") {
      target = { "@odata.type": "#microsoft.graph.allDevicesAssignmentTarget" };
      targetLabel = "All devices";
    } else if (targetType === "allUsers") {
      target = { "@odata.type": "#microsoft.graph.allLicensedUsersAssignmentTarget" };
      targetLabel = "All users";
    } else {
      if (!groupName) {
        setMessage("Enter a group name.", "error");
        setUiBusy(false);
        return;
      }

      let groupId = null;

      // Use suggestions cache first (if user picked from dropdown)
      if (groupNameSuggestionCache && groupNameSuggestionCache[groupName]) {
        groupId = groupNameSuggestionCache[groupName];
      
        if (groupId) cachedGroupMap[groupId] = groupName;
}

      if (!groupId) {
        for (const [gid, gname] of Object.entries(cachedGroupMap)) {
if ((gname || "").toLowerCase() === groupName.toLowerCase()) {
          groupId = gid;
          break;
        }
        }
      }

      if (!groupId) {
        const resolved = await resolveGroupByName(groupName, accessToken);
        if (!resolved) {
          setMessage("Group not found: " + groupName, "error");
          setUiBusy(false);
          return;
        }
        groupId = resolved.id;
        cachedGroupMap[groupId] = resolved.displayName || groupName;
      }

      target = {
        "@odata.type": groupMode === "exclude"
          ? "#microsoft.graph.exclusionGroupAssignmentTarget"
          : "#microsoft.graph.groupAssignmentTarget",
        groupId
      };
      targetLabel = cachedGroupMap[groupId] || groupName;
    }

    // Apply assignment filter (optional)
    const normalizedFilterMode = normalizeFilterMode(filterMode);
    if (normalizedFilterMode !== "none") {
      if (!filterName) {
        setMessage("Enter an assignment filter name (or set filter mode to 'No filter').", "error");
        setUiBusy(false);
        return;
      }

      let filterId = null;

      if (assignmentFilterSuggestionCache && assignmentFilterSuggestionCache[filterName]) {
        filterId = assignmentFilterSuggestionCache[filterName];
      }

      if (!filterId) {
        const wanted = filterName.toLowerCase();
        for (const [fid, fname] of Object.entries(cachedAssignmentFilterMap || {})) {
          if ((fname || "").toLowerCase() === wanted) {
            filterId = fid;
            break;
          }
        }
      }

      if (!filterId) {
        const resolved = await resolveAssignmentFilterByName(filterName, accessToken);
        if (!resolved) {
          if (lastAssignmentFilterLookupError && (lastAssignmentFilterLookupError.status === 401 || lastAssignmentFilterLookupError.status === 403)) {
            setMessage(
              "Couldn't read assignment filters (HTTP " + lastAssignmentFilterLookupError.status + "). " +
              "Check you have permission/admin consent for DeviceManagementConfiguration.Read.All. " +
              "Tip: you can paste the filter ID (GUID) instead of the name.",
              "error"
            );
          } else {
            setMessage("Assignment filter not found: " + filterName + ". Tip: you can paste the filter ID (GUID).", "error");
          }
          setUiBusy(false);
          return;
        }
        if (resolved.ambiguous) {
          const list = (resolved.matches || []).map(x => x.displayName).filter(Boolean).slice(0, 8);
          setMessage(
            "Filter name is ambiguous. Please refine it or pick from the suggestion list. Matches: " +
              (list.length ? list.join(" · ") : "(no names)"),
            "error"
          );
          setUiBusy(false);
          return;
        }

        filterId = resolved.id;
        cachedAssignmentFilterMap[filterId] = resolved.displayName || filterName;
      }

      target.deviceAndAppManagementAssignmentFilterId = filterId;
      target.deviceAndAppManagementAssignmentFilterType = normalizedFilterMode;

      const fname = cachedAssignmentFilterMap[filterId] || filterName;
      targetLabel = `${targetLabel} (filter: ${normalizedFilterMode} ${fname})`;
    }

    await bulkDeleteAssignments(accessToken, intent, targetType, target, targetLabel);
  } catch (err) {
    console.error("Error while removing assignments:", err);
    setStatus("Error while removing assignments – see console.");
    setMessage("An error occurred while removing assignments.", "error");
    logOutput("Error while removing assignments:\n" + String(err));
  } finally {
    setUiBusy(false);
  }
}

async function signOut() {
  const account = getActiveAccount();
  if (!account || !msalInstance) {
    setStatus("No active user.");
    setMessage("There is no active user to sign out.", "info");
    return;
  }

  try {
    await msalInstance.logoutPopup({ account });
    setStatus("Signed out.");
    setMessage("User has been signed out.", "success");
    logOutput("User has been signed out.");
    if (appsTableWrapper) {
      appsTableWrapper.innerHTML = "<em>No data – select a tenant and click “Sign in &amp; load apps”.</em>";
    }
    if (appsSummaryEl) {
      appsSummaryEl.textContent = "";
    }
    cachedApps = [];
    cachedAssignmentsMap = {};
    cachedGroupMap = {};
    clearAssignmentFilterCaches();
  } catch (err) {
    console.error("Error during sign-out:", err);
    setStatus("Error during sign-out – see console.");
    setMessage("An error occurred during sign-out.", "error");
    logOutput("Error during sign-out:\n" + String(err));
  }
}

async function handleLoginClick() {
  if (!currentTenantConfig) {
    setMessage("Configure and select a tenant first.", "error");
    return;
  }
  try {
    setUiBusy(true);
    await signInAndLoadApps();
  } finally {
    setUiBusy(false);
  }
}

async function handleLogoutClick() {
  try {
    setUiBusy(true);
    await signOut();
  } finally {
    setUiBusy(false);
  }
}

async function exportAssignmentsToExcel() {
  if (!cachedApps.length) {
    setMessage("Nothing to export – load apps first.", "error");
    return;
  }

  let appsToExport = cachedApps;

  if (platformFilterSelect) {
    const filterVal = platformFilterSelect.value || "all";
    if (filterVal !== "all") {
      appsToExport = appsToExport.filter(app => getAppPlatformKey(app) === filterVal);
    }
  }

  if (nameFilterInput) {
    const nameValue = (nameFilterInput.value || "").trim().toLowerCase();
    if (nameValue) {
      appsToExport = appsToExport.filter(app => {
        const n = (app.displayName || "").toLowerCase();
        const publisher = (app.publisher || app.developer || "").toLowerCase();
        return n.includes(nameValue) || publisher.includes(nameValue);
      });
    }
  }

  if (!appsToExport.length) {
    setMessage("No apps match current filters. Nothing to export.", "error");
    return;
  }

  // Ensure assignment filter names are available for export (best effort)
  try {
    let needFilterLoad = false;
    for (const app of appsToExport) {
      const appAssignments = cachedAssignmentsMap[app.id] || [];
      for (const a of appAssignments) {
        const t = (a && a.target) ? a.target : {};
        const fid = t.deviceAndAppManagementAssignmentFilterId || t.assignmentFilterId || "";
        if (fid && !cachedAssignmentFilterMap[fid]) { needFilterLoad = true; break; }
      }
      if (needFilterLoad) break;
    }
    if (needFilterLoad) {
      await fetchAllAssignmentFilters();
    }
  } catch (e) {
    logOutput("Warning: failed to preload assignment filters for export: " + String(e));
  }

  const rows = [];
  const tenantName = currentTenantConfig ? currentTenantConfig.name : "";
  const tenantId = currentTenantConfig ? currentTenantConfig.tenant : "";

  rows.push([
    "Tenant name",
    "Tenant",
    "App name",
    "Platform",
    "App type",
    "Publisher",
    "Intent",
    "Target type",
    "Target",
    "Assignment filter",
    "Filter mode",
    "Filter Id",
    "Assignment Id",
    "App Id"
  ]);

  for (const app of appsToExport) {
    const appAssignments = cachedAssignmentsMap[app.id] || [];
    const platformLabel = getAppPlatformLabel(app);
    const appType = parseAppType(app["@odata.type"]);
    const publisher = app.publisher || app.developer || "";
    const appName = app.displayName || "(no name)";

    if (!appAssignments.length) {
      rows.push([
        tenantName,
        tenantId,
        appName,
        platformLabel,
        appType,
        publisher,
        "(none)",
        "",
        "",
        "",
        "",
        "",
        "",
        app.id
      ]);
      continue;
    }

    for (const assignment of appAssignments) {
      let intent = (assignment.intent || "").toLowerCase();
      if (intent === "availablewithoutenrollment") intent = "available";

      const label = getAssignmentLabel(assignment, cachedGroupMap);

      const target = assignment.target || {};
      const t = (target["@odata.type"] || "").toLowerCase();
      let targetTypeLabel = "Other";
      let targetName = label;

      if (t.includes("alldevicesassignmenttarget")) {
        targetTypeLabel = "All devices";
        targetName = "";
      } else if (t.includes("alllicensedusersassignmenttarget") || t.includes("allusersassignmenttarget")) {
        targetTypeLabel = "All users";
        targetName = "";
      } else if (t.includes("exclusiongroupassignmenttarget")) {
        targetTypeLabel = "Group (exclude)";
        const gid = target.groupId;
        targetName = gid ? (cachedGroupMap[gid] || gid) : label;
      } else if (t.includes("groupassignmenttarget")) {
        targetTypeLabel = "Group";
        const gid = target.groupId;
        targetName = gid ? (cachedGroupMap[gid] || gid) : label;
      }

      const filterId = target.deviceAndAppManagementAssignmentFilterId || target.assignmentFilterId || "";
      const filterMode = normalizeFilterMode(target.deviceAndAppManagementAssignmentFilterType || target.assignmentFilterType);
      const filterName = filterId ? (cachedAssignmentFilterMap[filterId] || filterId) : "";

      rows.push([
        tenantName,
        tenantId,
        appName,
        platformLabel,
        appType,
        publisher,
        intent || "",
        targetTypeLabel,
        targetName,
        filterName,
        (filterMode !== "none") ? filterMode : "",
        filterId,
        assignment.id || "",
        app.id
      ]);
}
  }

  try {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, "Assignments");

    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });

    const now = new Date();
    const yyyy = now.getFullYear();
    const mm = String(now.getMonth() + 1).padStart(2, "0");
    const dd = String(now.getDate()).padStart(2, "0");
    const safeTenantName = (tenantName || "tenant").replace(/[^a-z0-9\-]+/gi, "_");
    const fileName = `intune-app-assignments_${safeTenantName}_${yyyy}-${mm}-${dd}.xlsx`;

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    setMessage(`Exported ${rows.length - 1} rows to ${fileName}.`, "success");
    logOutput(`Exported assignments to Excel (${fileName}) with ${rows.length - 1} data rows.`);
  } catch (e) {
    console.error("Excel export error:", e);
    setMessage("Failed to export to Excel.", "error");
    logOutput("Excel export error:\n" + String(e));
  }
}

// Event listeners
if (loginButton) loginButton.addEventListener("click", handleLoginClick);
if (logoutButton) logoutButton.addEventListener("click", handleLogoutClick);
if (bulkAssignButton) bulkAssignButton.addEventListener("click", handleBulkAssignClick);
if (bulkDeleteButton) bulkDeleteButton.addEventListener("click", handleBulkDeleteClick);


function updateGroupControlsVisibility() {
  const val = bulkTargetTypeSelect ? bulkTargetTypeSelect.value : "allDevices";
  const isGroup = val === "group";

  if (bulkGroupNameWrapper) {
    bulkGroupNameWrapper.style.display = isGroup ? "block" : "none";
  }
  if (bulkGroupModeWrapper) {
    bulkGroupModeWrapper.style.display = isGroup ? "block" : "none";
  }
  if (bulkGroupNameInput) {
    bulkGroupNameInput.disabled = !isGroup;
    if (!isGroup) bulkGroupNameInput.value = "";
  }
  if (bulkGroupModeSelect) {
    bulkGroupModeSelect.disabled = !isGroup;
    if (!isGroup) bulkGroupModeSelect.value = "include";
  }
  if (!isGroup) {
    const dl = document.getElementById("group-name-suggestions");
    if (dl) dl.innerHTML = "";
  }
}


function updateAssignmentFilterControlsVisibility() {
  const mode = bulkAssignmentFilterModeSelect ? bulkAssignmentFilterModeSelect.value : "none";

  const isGroup = bulkTargetTypeSelect && bulkTargetTypeSelect.value === "group";
  const isExcludeGroup = isGroup && bulkGroupModeSelect && bulkGroupModeSelect.value === "exclude";

  // Exclusion assignments do not support Assignment Filters – force No filter.
  if (bulkAssignmentFilterModeSelect) {
    bulkAssignmentFilterModeSelect.disabled = isExcludeGroup;
    if (isExcludeGroup && bulkAssignmentFilterModeSelect.value !== "none") {
      bulkAssignmentFilterModeSelect.value = "none";
    }
  }

  const effectiveMode = bulkAssignmentFilterModeSelect ? bulkAssignmentFilterModeSelect.value : mode;
  const hasFilter = effectiveMode && effectiveMode !== "none";

  if (bulkAssignmentFilterNameWrapper) {
    bulkAssignmentFilterNameWrapper.style.display = hasFilter ? "block" : "none";
  }
  if (bulkAssignmentFilterNameInput) {
    bulkAssignmentFilterNameInput.disabled = !hasFilter || isExcludeGroup;
    if (!hasFilter) bulkAssignmentFilterNameInput.value = "";
  }

  // Optional hint in UI
  const hintEl = document.getElementById("bulk-assignment-filter-hint");
  if (hintEl) {
    hintEl.style.display = isExcludeGroup ? "block" : "none";
  }

  if (!hasFilter) {
    const dl = document.getElementById("assignment-filter-suggestions");
    if (dl) dl.innerHTML = "";
  }

}


if (bulkTargetTypeSelect) {
  bulkTargetTypeSelect.addEventListener("change", () => {
    updateGroupControlsVisibility();
  });
}

if (bulkGroupModeSelect) {
  bulkGroupModeSelect.addEventListener("change", () => {
    updateAssignmentFilterControlsVisibility();
  });
}

// Initial UI state
updateGroupControlsVisibility();
updateAssignmentFilterControlsVisibility();
if (clearLogButton) {
  clearLogButton.addEventListener("click", () => {
    clearOutput();
    setMessage("Log cleared.", "info");
  });
}

if (platformFilterSelect) {
  platformFilterSelect.addEventListener("change", () => {
    if (!cachedApps.length) return;
    renderApps(cachedApps, cachedAssignmentsMap, cachedGroupMap);
  });
}

if (nameFilterInput) {
  nameFilterInput.addEventListener("input", () => {
    if (!cachedApps.length) return;
    renderApps(cachedApps, cachedAssignmentsMap, cachedGroupMap);
  });
}

if (exportXlsxButton) {
  exportXlsxButton.addEventListener("click", exportAssignmentsToExcel);
}

function setupCollapsiblePanel(toggleId, panelId, initialCollapsed = false) {
  const toggleEl = document.getElementById(toggleId);
  const panelEl = document.getElementById(panelId);
  if (!toggleEl || !panelEl) return;

  const iconEl = toggleEl.querySelector(".toggle-icon");

  // read state from panelState (if present)
  let collapsed = panelState.hasOwnProperty(panelId)
    ? !!panelState[panelId]
    : initialCollapsed;

  function applyState(c) {
    panelEl.style.display = c ? "none" : "block";
    toggleEl.classList.toggle("collapsed", c);
    if (iconEl) {
      iconEl.textContent = c ? "▸" : "▾";
    }
  }

  applyState(collapsed);

  toggleEl.addEventListener("click", () => {
    collapsed = !collapsed;
    applyState(collapsed);
    panelState[panelId] = collapsed;
    savePanelStateToStorage();
  });
}

window.addEventListener("DOMContentLoaded", () => {
  try {
    const savedTheme = localStorage.getItem(THEME_STORAGE_KEY) || "dark";
    applyTheme(savedTheme);
  } catch {
    applyTheme("dark");
  }

  // restore panel state
  panelState = loadPanelStateFromStorage();

  tenantConfigs = loadTenantsFromStorage();
  if (tenantConfigs.length > 0) {
    currentTenantConfig = tenantConfigs[0];
    initMsalForCurrentTenant();
  } else {
    setStatus("Status: configure at least one tenant above.");
  }
  renderTenantSelect();
  applyDisabledState();

  setupCollapsiblePanel("tenant-panel-toggle", "tenant-panel", false);
  setupCollapsiblePanel("auth-panel-toggle", "auth-panel", false);
  setupCollapsiblePanel("bulk-panel-toggle", "bulk-panel", false);
  setupCollapsiblePanel("filter-panel-toggle", "filter-panel", false);
  setupCollapsiblePanel("apps-panel-toggle", "apps-panel", false);
  setupCollapsiblePanel("log-panel-toggle", "log-panel", false);
});
