# Intune Apps Bulk Assignment (Multi-Tenant)

[![Latest release](https://img.shields.io/github/v/release/<OWNER>/<REPO>?display_name=tag&sort=semver)](https://github.com/<OWNER>/<REPO>/releases)
[![Downloads](https://img.shields.io/github/downloads/<OWNER>/<REPO>/total)](https://github.com/<OWNER>/<REPO>/releases)
[![Issues](https://img.shields.io/github/issues/<OWNER>/<REPO>)](https://github.com/<OWNER>/<REPO>/issues)
[![License](https://img.shields.io/github/license/<OWNER>/<REPO>)](https://github.com/<OWNER>/<REPO>/blob/main/LICENSE)

A lightweight **single-page web app** (pure HTML/JS) for **bulk adding and removing Intune app assignments** — including **assignment filters**, **multi-tenant profiles**, and **export of current assignments to Excel (.xlsx)**. ⚡

> ✅ Runs 100% in the browser and talks directly to Microsoft Graph (no backend).

* * *

## Table of contents
- ✨ Features
- Requirements
- Installation
- Quick start
- 📦 Export to Excel
- 🔐 Security notes
- ⚠️ Known limitations
- 🧰 Troubleshooting
- 🤝 Contributing
- 🐞 Reporting issues & feedback
- 📌 Project status
- 📄 License / Disclaimer

* * *

## ✨ Features
- ✅ **Bulk assignment**: add the same assignment (intent + target) to many apps in one go (with a 2-click confirmation).
- ✅ **Bulk removal**: remove a selected assignment from multiple apps (also with confirmation).
- 🎯 **Targets supported**:
  - All devices
  - All users
  - **Group (include / exclude)**
- 🧩 **Assignment filters** (include / exclude) with name suggestions + support for pasting filter GUID.
- 🧠 **Conflict detection** (e.g., same target but different intent) and skipping problematic apps with a clear message.
- 📦 **Excel export (.xlsx)** of apps + assignments, with filtering by platform and app name.
- 🏢 **Multi-tenant**: store multiple tenant profiles (Tenant + Client ID), quickly switch between them, export/import profiles as JSON.
- 🖱️ **Remove a single assignment** from the table via **right-click / context menu**.
- 🌐 **No backend** — everything runs in the browser using Microsoft Graph.

* * *

## Requirements
- 🌍 A modern browser (Chrome / Edge recommended).
- 🔑 Microsoft Intune access and sufficient roles/permissions (e.g., Intune Administrator), depending on your org policies.
- 🆔 **Microsoft Entra ID App Registration** (SPA) with delegated Microsoft Graph permissions.

* * *

## Installation
This is a **static** app — host the files as a web page.

1. 📁 Copy the repository files to any static hosting (GitHub Pages / IIS / Nginx / Azure Storage Static Website, etc.).
2. 🔁 Add the hosting URL as a **Redirect URI** in Entra ID (see below).
3. ✅ Open the app in your browser.

* * *

## Entra ID setup (App Registration)
The app derives Redirect URI from the current page URL (`window.location.origin + window.location.pathname`).
**That exact URL must be registered as a Redirect URI** in Entra ID. ⚠️

1. Microsoft Entra ID → **App registrations** → **New registration**
2. Go to **Authentication** → **Add a platform** → **Single-page application (SPA)**
3. Add the Redirect URI (exact hosting URL)
4. Copy **Application (client) ID** (you’ll enter it in the app UI)
5. Add the delegated Microsoft Graph permissions below (Admin consent may be required)

### Required Microsoft Graph permissions (Delegated)
- `User.Read`
- `Group.Read.All`
- `DeviceManagementApps.ReadWrite.All`
- `DeviceManagementConfiguration.Read.All`

* * *

## Quick start
1. 🏢 In **Tenant configuration**, add a profile (Name, Tenant ID/domain, Client ID) and save.
   - (Optional) export/import tenant profiles as JSON.
2. 🔐 Select the tenant and click **Sign in & load apps**.
3. 🔎 (Optional) set **Filters** (platform / name search) — affects the table and export.
4. ✅ Select apps in the table.
5. 🧰 In **Bulk assignment / removal**:
   - choose **Intent** (Required / Available / Uninstall)
   - choose **Target** (All devices / All users / Group)
   - for **Group**: choose include/exclude and type the group name (with suggestions)
   - (optional) set **Assignment filter** (include/exclude + filter name)
6. ▶️ Click **Add assignment**:
   - 1st click shows a summary
   - 2nd click executes the change  
   Same flow for **Remove assignment**.
7. 🖱️ Remove a single assignment from the table using **right-click → Remove assignment**.

* * *

## 📦 Export to Excel
Use **Export to Excel (apps & assignments)** in the Filters section.

Export includes (among others):
- Tenant name / Tenant
- App name / Platform / App type / Publisher
- Intent / Target type / Target
- Assignment filter (name) / Filter mode / Filter Id
- Assignment Id / App Id

The file name is generated like:
`intune-app-assignments_<tenant>_<YYYY-MM-DD>.xlsx`

* * *

## 🔐 Security notes
- ✅ No backend — requests go directly from your browser to Microsoft Graph.
- 💾 Tenant profiles are stored locally in `localStorage`.
- 🧾 Auth tokens are stored in `sessionStorage`.

* * *

## ⚠️ Known limitations
- 🚫 **Assignment filters are not supported for “Exclude group”** (Graph/Intune behavior) — the app blocks that combination to prevent `BadRequest`.
- 📄 App list loads in pages, but the tool may **stop after ~500 apps** as a safety limit.
- 🧪 Reading assignments may use the **beta** endpoint for Intune `mobileApps` assignments.

* * *

## 🧰 Troubleshooting
### ❌ AADSTS50011 / redirect_uri_mismatch
- Ensure the Redirect URI in Entra ID matches the **exact** hosting URL (including path).

### ❌ 403 / missing groups / missing filters
- Usually missing consent for `DeviceManagementConfiguration.Read.All` and/or `Group.Read.All`.

### ❌ Filter not found by name
- Try a more exact name or paste the filter GUID directly.

### 🔎 Need details?
- Check the **Raw log** panel for Graph requests and error payloads.

* * *

## 🤝 Contributing
Contributions are welcome! 🛠️

1. Check existing issues and open a new one if needed.
2. Fork the repo and create a feature branch.
3. Commit changes with clear messages.
4. Open a Pull Request describing:
   - what was changed,
   - why it is useful,
   - how it was tested.

* * *

## 🐞 Reporting issues & feedback
Bug reports and feature requests are very welcome. 💬

Open an issue here:
- https://github.com/<OWNER>/<REPO>/issues

When reporting a bug, please include:
- steps to reproduce,
- expected vs. actual behavior,
- a sanitized snippet from **Raw log**,
- whether it happens in one tenant or multiple tenants.

* * *

## 📌 Project status
Actively maintained. 🚀  
Check **Releases** for the latest version, change log and downloads:
- https://github.com/<OWNER>/<REPO>/releases

* * *

## 📄 License / Disclaimer
This tool is not a Microsoft product and is not affiliated with Microsoft.  
Use at your own risk — always test in a non-production environment first.

> Replace `<OWNER>/<REPO>` in badges and links with your GitHub repository path.
