async function fetchAllApps(accessToken) {
  let url = graphAppsBaseUrl;
  let allApps = [];
  let page = 1;

  while (url) {
    setStatus(`Loading apps (page ${page})…`);

    const res = await fetch(url, {
      method: "GET",
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "application/json"
      }
    });

    const text = await res.text();
    let json = null;
    try {
      json = JSON.parse(text);
    } catch (e) {}

    if (!res.ok) {
      setStatus("Graph API error while loading apps (HTTP " + res.status + ")");
      logOutput("Graph error (apps):\n" + text);
      throw new Error("Graph error when fetching apps");
    }

    const pageApps = (json && json.value) || [];
    allApps = allApps.concat(pageApps);

    logOutput(`Apps page ${page} loaded (${pageApps.length} apps).`);
    url = json["@odata.nextLink"] || null;
    page += 1;

    if (allApps.length > 500) {
      url = null;
    }
  }

  return allApps;
}

async function fetchAssignmentsForApp(appId, accessToken) {
  const url =
    `https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/${encodeURIComponent(appId)}/assignments` +
    `?$select=id,intent,target`;

  const res = await fetch(url, {
    method: "GET",
    headers: {
      "Authorization": `Bearer ${accessToken}`,
      "Accept": "application/json"
    }
  });

  const text = await res.text();
  let json = null;
  try {
    json = JSON.parse(text);
  } catch (e) {}

  if (!res.ok) {
    console.warn("Error fetching assignments for appId", appId, "status", res.status, text);
    logOutput(`Error fetching assignments for ${appId}:\n${text}`);
    return [];
  }

  return (json && json.value) || [];
}



async function refreshAssignmentsForApps(appIds, accessToken) {
  const unique = Array.from(new Set(appIds || [])).filter(Boolean);
  for (const appId of unique) {
    try {
      cachedAssignmentsMap[appId] = await fetchAssignmentsForApp(appId, accessToken);
    } catch (e) {
      console.warn("Failed to refresh assignments for app", appId, e);
      logOutput("Failed to refresh assignments for app " + getAppNameById(appId) + ":\\n" + String(e));
      try {
    const token = await getAccessTokenForGraph();
    await resolveMissingGroupNamesForAssignments(token, cachedAssignmentsMap);
  } catch (e) {}
}
    try {
    const token = await getAccessTokenForGraph();
    await resolveMissingGroupNamesForAssignments(token, cachedAssignmentsMap);
  } catch (e) {}
}
  try {
    const token = await getAccessTokenForGraph();
    await resolveMissingGroupNamesForAssignments(token, cachedAssignmentsMap);
  } catch (e) {}
}


async function fetchGroupsForIds(groupIds, accessToken) {
  const uniqueIds = [...new Set(groupIds.filter(Boolean))];
  const groupMap = {};
  let idx = 0;

  for (const gid of uniqueIds) {
    idx += 1;
    setStatus(`Loading group names (${idx}/${uniqueIds.length})…`);

    const url = `https://graph.microsoft.com/v1.0/groups/${encodeURIComponent(gid)}?$select=displayName`;
    const res = await fetch(url, {
      method: "GET",
      headers: {
        "Authorization": `Bearer ${accessToken}`,
        "Accept": "application/json"
      }
    });

    const text = await res.text();
    let json = null;
    try {
      json = JSON.parse(text);
    } catch (e) {}

    if (res.ok && json) {
      groupMap[gid] = json.displayName || gid;
    } else {
      console.warn("Error fetching group", gid, "status", res.status, text);
      logOutput("Error fetching group " + gid + ":\n" + text);
      groupMap[gid] = gid;
    }
  }

  return groupMap;
}

async function signInAndLoadApps() {
  if (!currentTenantConfig || !msalInstance) {
    setMessage("Configure and select a tenant first.", "error");
    return;
  }

  setStatus("Signing in…");
  if (!suppressNextRefreshMessage) {
    setMessage("", "info");
  }
  logOutput("Starting sign-in and app loading…");

  try {
    let account = getActiveAccount();

    if (!account) {
      const loginResult = await msalInstance.loginPopup(loginRequest);
      account = loginResult.account;
      msalInstance.setActiveAccount(account);
      setStatus("Signed in as: " + (account.username || account.name || "unknown user"));
      logOutput("Signed in as: " + (account.username || account.name || "unknown user"));
    } else {
      setStatus("User already signed in: " + (account.username || account.name));
      logOutput("User already signed in: " + (account.username || account.name));
    }

    const tokenRequest = { ...loginRequest, account };

    setStatus("Acquiring access token…");
    logOutput("Acquiring access token…");

    let tokenResponse;
    try {
      tokenResponse = await msalInstance.acquireTokenSilent(tokenRequest);
    } catch (silentError) {
      console.warn("acquireTokenSilent error, falling back to popup:", silentError);
      logOutput("acquireTokenSilent error:\n" + silentError + "\nTrying acquireTokenPopup…");
      tokenResponse = await msalInstance.acquireTokenPopup(tokenRequest);
    }

    const accessToken = tokenResponse.accessToken;

    // Assignment filters are tenant-scoped; clear caches on refresh
    clearAssignmentFilterCaches();

    // Preload assignment filters so the table can show filter name + include/exclude
    try {
      await fetchAllAssignmentFilters(accessToken);
    } catch (e) {
      console.warn('Failed to preload assignment filters', e);
      logOutput('Failed to preload assignment filters:\n' + String(e));
    }

    const apps = await fetchAllApps(accessToken);
    if (!apps || apps.length === 0) {
      cachedApps = [];
      cachedAssignmentsMap = {};
      cachedGroupMap = {};
      renderApps([], {}, {});
      if (!suppressNextRefreshMessage) {
        setMessage("No apps returned from Intune.", "info");
      }
      suppressNextRefreshMessage = false;
      logOutput("No apps returned from Intune.");
      return;
    }

    const assignmentsMap = {};
    const allGroupIds = new Set();
    let idx = 0;

    for (const app of apps) {
      idx += 1;
      setStatus(`Loading assignments (${idx}/${apps.length})…`);

      const assignments = await fetchAssignmentsForApp(app.id, accessToken);
      assignmentsMap[app.id] = assignments;

      for (const a of assignments) {
        const gid = extractGroupIdFromAssignment(a);
        if (gid) {
          allGroupIds.add(gid);
        }
      }
    }

    const groupMap = await fetchGroupsForIds([...allGroupIds], accessToken);

    cachedApps = apps;
    cachedAssignmentsMap = assignmentsMap;
    cachedGroupMap = groupMap;

    setStatus("Done – apps, assignments and group names loaded.");
    if (!suppressNextRefreshMessage) {
      setMessage("Data refreshed.", "success");
    }
    suppressNextRefreshMessage = false;
    logOutput("Apps, assignments and group names loaded successfully.");
    renderApps(apps, assignmentsMap, groupMap);

  } catch (err) {
    console.error("Error in signInAndLoadApps:", err);
    setStatus("Error – see console for details.");
    if (!suppressNextRefreshMessage) {
      setMessage("An error occurred while fetching data from Graph API.", "error");
    }
    suppressNextRefreshMessage = false;
    logOutput("Error in signInAndLoadApps:\n" + String(err));
  }
}

function getSelectedAppIds() {
  const checkboxes = document.querySelectorAll(".app-select:checked");
  return Array.from(checkboxes).map(cb => cb.getAttribute("data-app-id"));
}

async function resolveGroupByName(name, accessToken) {
  const trimmed = name.trim();
  if (!trimmed) return null;

  const escaped = trimmed.replace(/'/g, "''");
  const filter = "displayName eq '" + escaped + "'";
  const url =
    "https://graph.microsoft.com/v1.0/groups?$filter=" +
    encodeURIComponent(filter) +
    "&$select=id,displayName&$top=1";

  const res = await fetch(url, {
    method: "GET",
    headers: {
      "Authorization": `Bearer ${accessToken}`,
      "Accept": "application/json"
    }
  });

  const text = await res.text();
  let json = null;
  try {
    json = JSON.parse(text);
  } catch (e) {}

  if (!res.ok) {
    console.warn("Error searching group by name", trimmed, res.status, text);
    logOutput("Error searching group by name '" + trimmed + "':\n" + text);
    return null;
  }

  const items = (json && json.value) || [];
  if (!items.length) return null;
  return items[0];
}

async function resolveAssignmentFilterByName(name, accessToken) {
  const trimmed = (name || "").trim();
  if (!trimmed) logOutput(`Assignment filter not found by name locally: "${trimmed}"`);
  return null;

  // Allow pasting the filter ID (GUID) directly
  if (looksLikeGuid(trimmed)) {
    upsertAssignmentFilterIntoCaches({ id: trimmed, displayName: cachedAssignmentFilterMap[trimmed] || trimmed });
    return { id: trimmed, displayName: cachedAssignmentFilterMap[trimmed] || trimmed };
  }

  const wanted = trimmed.toLowerCase();

  // 1) Fast path: cache exact (case-insensitive)
  const direct = assignmentFilterSuggestionCache[trimmed] || assignmentFilterSuggestionCache[wanted];
  if (direct) {
    return { id: direct, displayName: cachedAssignmentFilterMap[direct] || trimmed };
  }

  // 2) Try to load full list (best effort)
  try { await fetchAllAssignmentFilters(); } catch (e) {}

  // 2a) Exact match in id->name map
  for (const [fid, fname] of Object.entries(cachedAssignmentFilterMap || {})) {
    if ((fname || "").toLowerCase() === wanted) return { id: fid, displayName: fname };
  }

  // 2b) Partial match in cached list (contains)
  const localMatches = (cachedAssignmentFiltersList || [])
    .filter(f => (f.displayName || "").toLowerCase().includes(wanted));

  if (localMatches.length === 1) {
    upsertAssignmentFilterIntoCaches(localMatches[0]);
    return localMatches[0];
  }
  if (localMatches.length > 1) {
    return { ambiguous: true, matches: localMatches.slice(0, 10) };
  }
  // No server-side $filter search here: many tenants (and especially some proxies) return
  // confusing OData errors for this endpoint when filtering from a browser. We rely on full list + local match.
  logOutput(`Assignment filter not found by name locally: "${trimmed}"`);
  return null;
}


function normalizeFilterMode(val) {
  const v = String(val || "").toLowerCase();
  if (v === "include" || v === "exclude") return v;
  return "none";
}
function getFilterKeyFromTarget(target) {
  const t = target || {};
  const ft = normalizeFilterMode(t.deviceAndAppManagementAssignmentFilterType);
  const fid = t.deviceAndAppManagementAssignmentFilterId || "";
  if (ft !== "none" && fid) return `filter:${ft}:${fid}`;
  return "filter:none";
}

async function bulkAssignToApps(accessToken, intent, targetType, target, targetLabel) {
  const selectedAppIds = getSelectedAppIds();
  const intentLower = intent.toLowerCase();

  let baseKey = null;
  if (targetType === "allDevices") baseKey = "allDevices";
  else if (targetType === "allUsers") baseKey = "allUsers";
  else {
    const tt = ((target && target["@odata.type"]) || "").toLowerCase();
    if (tt.includes("exclusiongroupassignmenttarget")) baseKey = "excludeGroup:" + (target.groupId || "");
    else baseKey = "group:" + (target.groupId || "");
  }

  const targetKey = baseKey + "|" + getFilterKeyFromTarget(target);

  const toAssignIds = [];
  const conflictIds = [];

  for (const appId of selectedAppIds) {
    const existingAssignments = cachedAssignmentsMap[appId] || [];
    let conflict = false;

    for (const a of existingAssignments) {
      const existingKey = getTargetKeyFromAssignment(a);
      if (!existingKey || existingKey !== targetKey) continue;

      let existingIntent = (a.intent || "").toLowerCase();
      if (existingIntent === "availablewithoutenrollment") {
        existingIntent = "available";
      }

      if (existingIntent && existingIntent !== intentLower) {
        conflict = true;
        break;
      }
    }

    if (conflict) conflictIds.push(appId);
    else toAssignIds.push(appId);
  }

  if (conflictIds.length > 0) {
    const names = conflictIds.slice(0, 5).map(getAppNameById).join(", ");
    const more =
      conflictIds.length > 5 ? ` and ${conflictIds.length - 5} more…` : "";
    setMessage(
      `Skipped ${conflictIds.length} apps due to conflicts (same target "${targetLabel}", different intents). ` +
      `Examples: ${names}${more}.`,
      "error"
    );
    logOutput(
      `Conflicts detected for ${conflictIds.length} apps (target "${targetLabel}", intent "${intent}").`
    );
  }

  if (!toAssignIds.length) {
    setStatus("Bulk assignment stopped – all selected apps have conflicts.");
    logOutput("Bulk assignment stopped: all selected apps have conflicts.");
    return;
  }

  setStatus("Running bulk assignment…");
  logOutput(
    `Running bulk assignment: intent "${intent}", target "${targetLabel}" for ${toAssignIds.length} apps.`
  );
  let successCount = 0;
  let errorCount = 0;

  for (const appId of toAssignIds) {
    const url =
      `https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps/${encodeURIComponent(appId)}/assignments`;

    const body = {
      "@odata.type": "#microsoft.graph.mobileAppAssignment",
      intent: intent,
      target: target
    };

    try {
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

      if (!res.ok) {
        console.error("Assignment error for appId", appId, res.status, text);
        errorCount++;
        logOutput("Assignment error for app " + getAppNameById(appId) + ":\n" + text);
      } else {
        successCount++;
        logOutput("Assignment OK for app " + getAppNameById(appId) + ".");
      
        // Optimistic cache update so table refreshes immediately
        if (!cachedAssignmentsMap[appId]) cachedAssignmentsMap[appId] = [];
        cachedAssignmentsMap[appId].push({
          id: "local-" + Date.now() + "-" + Math.random().toString(16).slice(2),
          intent: intent,
          target: target
        });
        renderApps(cachedApps, cachedAssignmentsMap, cachedGroupMap);
}
    } catch (e) {
      console.error("Exception while assigning appId", appId, e);
      errorCount++;
      logOutput("Exception while assigning app " + getAppNameById(appId) + ":\n" + String(e));
    }
  }

  if (errorCount === 0) {
    setMessage(
      `Bulk assignment finished. New "${intent}" → "${targetLabel}" assignments created for ${successCount} apps.`,
      "success"
    );
  } else {
    setMessage(
      `Bulk assignment finished. Success: ${successCount}, errors: ${errorCount}. See log below.`,
      "error"
    );
  }
  setStatus("Bulk assignment completed.");
  logOutput(
    `Bulk assignment completed. Success: ${successCount}, errors: ${errorCount}.`
  );

    // Refresh ONLY assignments for affected apps (no full reload)
  await refreshAssignmentsForApps(toAssignIds, accessToken);
  renderApps(cachedApps, cachedAssignmentsMap, cachedGroupMap);
}

async function bulkDeleteAssignments(accessToken, intent, targetType, target, targetLabel) {
  const selectedAppIds = getSelectedAppIds();
  const intentLower = intent.toLowerCase();

  let baseKey = null;
  if (targetType === "allDevices") baseKey = "allDevices";
  else if (targetType === "allUsers") baseKey = "allUsers";
  else {
    const tt = ((target && target["@odata.type"]) || "").toLowerCase();
    if (tt.includes("exclusiongroupassignmenttarget")) baseKey = "excludeGroup:" + (target.groupId || "");
    else baseKey = "group:" + (target.groupId || "");
  }

  const targetKey = baseKey + "|" + getFilterKeyFromTarget(target);

  setStatus("Removing assignments…");
  logOutput(
    `Running bulk removal: intent "${intent}", target "${targetLabel}" for ${selectedAppIds.length} apps.`
  );

  let successCount = 0;
  let errorCount = 0;
  let noMatchCount = 0;

  for (const appId of selectedAppIds) {
    const existingAssignments = cachedAssignmentsMap[appId] || [];

    const toDelete = existingAssignments.filter(a => {
      const existingKey = getTargetKeyFromAssignment(a);
      if (!existingKey || existingKey !== targetKey) return false;

      let existingIntent = (a.intent || "").toLowerCase();
      if (existingIntent === "availablewithoutenrollment") {
        existingIntent = "available";
      }

      return existingIntent === intentLower;
    });

    if (!toDelete.length) {
      noMatchCount++;
      logOutput(
        `No matching "${intent}" assignment for "${targetLabel}" on app ${getAppNameById(appId)}.`
      );
      continue;
    }

    for (const assignment of toDelete) {
      if (!assignment.id) continue;

      const url =
        `https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps/${encodeURIComponent(appId)}/assignments/${encodeURIComponent(assignment.id)}`;

      try {
        const res = await fetch(url, {
          method: "DELETE",
          headers: {
            "Authorization": `Bearer ${accessToken}`
          }
        });

        const text = await res.text();

        if (!res.ok) {
          console.error("Error deleting assignment", assignment.id, "for appId", appId, res.status, text);
          errorCount++;
          logOutput(
            `Error deleting assignment for app ${getAppNameById(appId)} (assignmentId=${assignment.id}):\n` + text
          );
        } else {
          successCount++;
          logOutput(
            `Assignment removed for app ${getAppNameById(appId)} (assignmentId=${assignment.id}).`
          );
        
          // Update local cache so table refreshes immediately
          removeAssignmentFromCache(appId, { id: assignment.id, intent: intent, target: target });
          renderApps(cachedApps, cachedAssignmentsMap, cachedGroupMap);
}
      } catch (e) {
        console.error("Exception while deleting assignment", assignment.id, "for appId", appId, e);
        errorCount++;
        logOutput(
          `Exception while deleting assignment for app ${getAppNameById(appId)} (assignmentId=${assignment.id}):\n` +
          String(e)
        );
      }
    }
  }

  let msg = `Removing assignments "${intent}" → "${targetLabel}" finished. ` +
            `Removed: ${successCount}, errors: ${errorCount}.`;
  if (noMatchCount > 0) {
    msg += ` For ${noMatchCount} apps no such assignment was found.`;
  }

  setStatus("Assignment removal finished.");
  setMessage(msg, errorCount === 0 ? "success" : "error");
  logOutput(
    `Bulk removal completed. Removed: ${successCount}, errors: ${errorCount}, no-match apps: ${noMatchCount}.`
  );

    // Refresh ONLY assignments for affected apps (no full reload)
  await refreshAssignmentsForApps(selectedAppIds, accessToken);
  renderApps(cachedApps, cachedAssignmentsMap, cachedGroupMap);
}

