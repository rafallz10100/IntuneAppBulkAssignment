function setStatus(text) {
  if (!statusEl) return;
  statusEl.textContent = typeof text === "string" && text.startsWith("Status:")
    ? text
    : "Status: " + text;
}

function clearOutput() {
  if (!outputEl) return;
  outputEl.textContent = "";
}

function logOutput(text) {
  if (!outputEl || !text) return;
  const ts = new Date().toLocaleTimeString();
  const prefix = outputEl.textContent ? "\n\n" : "";
  outputEl.textContent += `${prefix}[${ts}] ${text}`;
  outputEl.scrollTop = outputEl.scrollHeight;
}

function setMessage(text, type = "info") {
  if (!messageBarEl) return;
  if (!text) {
    messageBarEl.style.display = "none";
    messageBarEl.textContent = "";
    messageBarEl.className = "";
    return;
  }
  messageBarEl.textContent = text;
  messageBarEl.className = "";
  if (type === "error") messageBarEl.classList.add("msg-error");
  else if (type === "success") messageBarEl.classList.add("msg-success");
  else messageBarEl.classList.add("msg-info");
  messageBarEl.style.display = "block";
}

function getActiveAccount() {
  if (!msalInstance) return null;
  const accounts = msalInstance.getAllAccounts();
  return accounts[0] || null;
}

function parseAppType(odataType) {
  if (!odataType) return "-";
  const parts = odataType.split(".");
  return parts[parts.length - 1] || odataType;
}

function getAppPlatformKey(app) {
  const t = (app["@odata.type"] || "").toLowerCase();

  if (
    t.includes("androidstoreapp") ||
    t.includes("androidlobapp") ||
    t.includes("managedandroidlobapp") ||
    t.includes("androidmanagedstoreapp") ||
    t.includes("android")
  ) {
    return "android";
  }

  if (
    t.includes("win32lobapp") ||
    t.includes("windowsstoreapp") ||
    t.includes("windows") ||
    t.includes("win")
  ) {
    return "windows";
  }

  if (
    t.includes("iosstoreapp") ||
    t.includes("ioslobapp") ||
    t.includes("iosvppapp") ||
    t.includes("ios")
  ) {
    return "ios";
  }

  if (t.includes("macos") || t.includes("mac")) {
    return "macos";
  }

  return "other";
}

function getAppPlatformLabel(app) {
  const key = getAppPlatformKey(app);
  switch (key) {
    case "android": return "Android";
    case "windows": return "Windows";
    case "ios":     return "iOS";
    case "macos":   return "macOS";
    default:        return "Other";
  }
}

function getAppNameById(appId) {
  const app = cachedApps.find(a => a.id === appId);
  return app ? (app.displayName || app.id) : appId;
}

function extractGroupIdFromAssignment(assignment) {
  if (!assignment || !assignment.target) return null;
  const target = assignment.target;
  const t = (target["@odata.type"] || "").toLowerCase();

  if (t.includes("groupassignmenttarget") || t.includes("exclusiongroupassignmenttarget")) {
    return target.groupId || null;
  }
  return null;
}

function getTargetKeyFromAssignment(assignment) {
  if (!assignment || !assignment.target) return null;
  const target = assignment.target;
  const t = (target["@odata.type"] || "").toLowerCase();

  let baseKey = null;

  if (t.includes("alldevicesassignmenttarget")) baseKey = "allDevices";
  else if (t.includes("alllicensedusersassignmenttarget") || t.includes("allusersassignmenttarget")) baseKey = "allUsers";
  else if (t.includes("exclusiongroupassignmenttarget")) {
    const gid = target.groupId;
    if (!gid) return null;
    baseKey = "excludeGroup:" + gid;
  } else if (t.includes("groupassignmenttarget")) {
    const gid = target.groupId;
    if (!gid) return null;
    baseKey = "group:" + gid;
  }

  if (!baseKey) return null;
  return baseKey + "|" + getFilterKeyFromTarget(target);
}




function sameTarget(aTarget, bTarget) {
  const a = aTarget || {};
  const b = bTarget || {};

  const aType = (a["@odata.type"] || "").toLowerCase();
  const bType = (b["@odata.type"] || "").toLowerCase();
  if (aType !== bType) return false;

  const aFilter = getFilterKeyFromTarget(a);
  const bFilter = getFilterKeyFromTarget(b);
  if (aFilter !== bFilter) return false;

  if (aType.includes("group")) {
    return a.groupId === b.groupId;
  }
  return true;
}


function removeAssignmentFromCache(appId, assignment) {
  if (!cachedAssignmentsMap[appId]) return;
  const aid = assignment && assignment.id ? String(assignment.id) : null;
  const intent = ((assignment && assignment.intent) || "").toLowerCase();
  const target = (assignment && assignment.target) || null;

  cachedAssignmentsMap[appId] = (cachedAssignmentsMap[appId] || []).filter(a => {
    if (!a) return false;

    if (aid && a.id && String(a.id) === aid) return false;

    if (intent) {
      const aIntent = (a.intent || "").toLowerCase();
      if (aIntent !== intent) return true;
      if (target && sameTarget(a.target, target)) return false;
    }

    return true;
  });
}


function getAssignmentLabel(assignment, groupMap) {
  if (!assignment || !assignment.target) return "(no target)";
  const target = assignment.target;
  const t = (target["@odata.type"] || "").toLowerCase();

  // Assignment filter info (optional)
  const filterId = target.deviceAndAppManagementAssignmentFilterId || target.assignmentFilterId || null;
  const filterTypeRaw = target.deviceAndAppManagementAssignmentFilterType || target.assignmentFilterType || "";
  const filterType = String(filterTypeRaw || "").toLowerCase().trim();
  let filterSuffix = "";
  if (filterId) {
    const filterName = cachedAssignmentFilterMap[filterId] || filterId;
    const typeLabel = (filterType && filterType !== "none") ? filterType : "";
    filterSuffix = typeLabel ? ` [Filter: ${filterName} · ${typeLabel}]` : ` [Filter: ${filterName}]`;
  } else if (filterType && filterType !== "none") {
    filterSuffix = ` [Filter: ${filterType}]`;
  }

  if (t.includes("alldevicesassignmenttarget")) {
    return "All devices" + filterSuffix;
  }
  if (t.includes("alllicensedusersassignmenttarget") || t.includes("allusersassignmenttarget")) {
    return "All users" + filterSuffix;
  }
  if (t.includes("exclusiongroupassignmenttarget")) {
    const gid = target.groupId;
    const name = gid ? (groupMap[gid] || gid) : "(groupId)";
    return "EXCLUDE: " + name + filterSuffix;
  }
  if (t.includes("groupassignmenttarget")) {
    const gid = target.groupId;
    const name = gid ? (groupMap[gid] || gid) : "(groupId)";
    return name + filterSuffix;
  }

  return "(other target)" + filterSuffix;
}

function renderApps(apps, assignmentsMap, groupMap) {
  if (!appsTableWrapper || !appsSummaryEl) return;

  if (!apps || apps.length === 0) {
    appsTableWrapper.innerHTML = "<em>No apps returned by Graph.</em>";
    appsSummaryEl.textContent = "";
    applyDisabledState();
    return;
  }

  let filteredApps = apps;

  if (platformFilterSelect) {
    const filterVal = platformFilterSelect.value || "all";
    if (filterVal !== "all") {
      filteredApps = filteredApps.filter(app => getAppPlatformKey(app) === filterVal);
    }
  }

  if (nameFilterInput) {
    const nameValue = (nameFilterInput.value || "").trim().toLowerCase();
    if (nameValue) {
      filteredApps = filteredApps.filter(app => {
        const n = (app.displayName || "").toLowerCase();
        const publisher = (app.publisher || app.developer || "").toLowerCase();
        return n.includes(nameValue) || publisher.includes(nameValue);
      });
    }
  }

  if (!filteredApps.length) {
    appsTableWrapper.innerHTML = "<em>No apps for selected filters.</em>";
    appsSummaryEl.textContent = `Loaded ${apps.length} apps in total, 0 match current filters.`;
    applyDisabledState();
    return;
  }

  const rowsHtml = filteredApps.map(app => {
    const appAssignments = assignmentsMap[app.id] || [];

    const buckets = { required: [], available: [], uninstall: [] };

    for (const a of appAssignments) {
      const intentRaw = (a.intent || "").toLowerCase();
      let intent = intentRaw;
      if (intentRaw === "availablewithoutenrollment") {
        intent = "available";
      }

      const label = getAssignmentLabel(a, groupMap);
      const item = { label, assignmentId: a.id || "" };

      if (intent === "required") {
        buckets.required.push(item);
      } else if (intent === "available") {
        buckets.available.push(item);
      } else if (intent === "uninstall") {
        buckets.uninstall.push(item);
      }
    }

    const hasRequired = buckets.required.length > 0;
    const hasAvailable = buckets.available.length > 0;
    const hasUninstall = buckets.uninstall.length > 0;

    const appType = parseAppType(app["@odata.type"]);
    const publisher = app.publisher || app.developer || "";
    const platformLabel = getAppPlatformLabel(app);

    const pillHtml = on => `<div class="pill ${on ? "on" : ""}">${on ? "✔" : ""}</div>`;
    const listHtml = arr =>
      arr.length
        ? `<div class="assignment-list">${arr
            .map(item => {
              const label = (typeof item === "string") ? item : (item.label || "");
              const assignmentId = (typeof item === "string") ? "" : (item.assignmentId || item.id || "");
              const safe = escapeHtml(label);
              const aidAttr = assignmentId ? ` data-assignment-id="${escapeHtml(assignmentId)}"` : "";
              return `<div class="assignment-item" data-app-id="${app.id}"${aidAttr} title="${safe}">${safe}</div>`;
            })
            .join("")}</div>`
        : `<div class="subtle">–</div>`;

    return `
      <tr>
        <td>
          <input type="checkbox" class="app-select" data-app-id="${app.id}" />
        </td>
        <td>
          <div>${app.displayName || "(no name)"}</div>
          <div class="app-type">
            ${platformLabel ? platformLabel + " · " : ""}${appType}
            ${publisher ? " · " + publisher : ""}
          </div>
        </td>
        <td>
          ${pillHtml(hasRequired)}
          ${listHtml(buckets.required)}
        </td>
        <td>
          ${pillHtml(hasAvailable)}
          ${listHtml(buckets.available)}
        </td>
        <td>
          ${pillHtml(hasUninstall)}
          ${listHtml(buckets.uninstall)}
        </td>
        <td>
          <span class="subtle">
            Assignments: ${appAssignments.length}
          </span>
        </td>
      </tr>
    `;
  }).join("");

  appsTableWrapper.innerHTML = `
    <div class="table-container">
      <div class="table-scroll">
        <table>
          <thead>
            <tr>
              <th style="width:2rem;">
                <input type="checkbox" id="select-all-apps" />
              </th>
              <th>App</th>
              <th>Required (groups / All)</th>
              <th>Available (groups / All)</th>
              <th>Uninstall (groups / All)</th>
              <th>Info</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
        </table>
      </div>
    </div>
  `;

  const total = apps.length;
  const filteredTotal = filteredApps.length;
  const withAssignments = filteredApps.filter(a => (assignmentsMap[a.id] || []).length > 0).length;
  appsSummaryEl.textContent =
    `Loaded ${total} apps in total, ${filteredTotal} match filters; ` +
    `${withAssignments} of filtered apps have at least one assignment.`;

  const selectAll = document.getElementById("select-all-apps");
  if (selectAll) {
    selectAll.addEventListener("change", () => {
      const checked = selectAll.checked;
      document.querySelectorAll(".app-select").forEach(cb => {
        cb.checked = checked;
      });
    });
  }

  applyDisabledState();
}

