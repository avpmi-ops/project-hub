# AVPMi Unified Project Hub — Cowork Briefing
**Date:** March 20, 2026
**From:** Closed Projects Cowork session
**Goal:** Wire the PMO Hub to live SharePoint data, merge it with the Closed Projects dashboard, and deploy the unified hub to SharePoint for the full team.

---

## What's Already Built and Working

### 1. Closed Projects Dashboard (Live ✅)
- **URL:** https://avpmi-ops.github.io/closed-projects-dashboard/
- **Source:** `avpmi-ops/closed-projects-dashboard` GitHub repo (`index.html`)
- **Reads from:** `AVPMi Closed Projects` SharePoint list (175 items as of today)
- **Auth:** MSAL browser auth — same Azure AD app below
- **Status:** Fully working. "Connect Live Data" button authenticates and loads live data.

### 2. PMO Hub (Prototype — needs wiring ⚠️)
- **File:** `C:/Users/RobertEdgar/Documents/avpmi-pmo-dashboard.html`
- **Reads from:** Inline static data (prototype)
- **Needs:** `DataService.load()` wired to `AVPMi Project Intake` SharePoint list via MSAL
- **Architecture is ready:** The `DataService` object already has `load()`, `getAll()`, `getActive()`, `getClosed()` methods. Just replace the inline data with a live SharePoint fetch.

---

## Critical Infrastructure (DO NOT GUESS THESE)

### SharePoint Site
```
Base URL: https://avpm-my.sharepoint.com/personal/robert_edgar_avpmi_com
```
Both lists live on **Rob's personal SharePoint site** (not avpmi.sharepoint.com).

### Azure AD App (MSAL)
```
Client ID:    85de8e7e-29c5-402b-8604-8483b569d8e5
Authority:    https://login.microsoftonline.com/avpmi.com
Tenant ID:    35339fc5-294f-41d6-8b9f-02a954bfd668
Scope:        https://avpm-my.sharepoint.com/AllSites.Read
```
This is a delegated (user) permission — no admin consent needed. Already working.

### MSAL Config Block (copy exactly)
```javascript
const MSAL_CONFIG = {
  clientId: "85de8e7e-29c5-402b-8604-8483b569d8e5",
  authority: "https://login.microsoftonline.com/avpmi.com",
  redirectUri: window.location.href.split(/[?#]/)[0],
  cacheLocation: "sessionStorage"
};
const SP_SCOPES = ["https://avpm-my.sharepoint.com/AllSites.Read"];
const SP_BASE   = "https://avpm-my.sharepoint.com/personal/robert_edgar_avpmi_com";
```

---

## SharePoint List Field Names (Verified — Use These Exactly)

### `AVPMi Project Intake` (Active Projects — 43 items)
These are the **internal field names** to use in `$select`:
```
Title                   — Project name
Project_Number          — QB / project number
Project_Type            — Choice: Swyft Valet, AVPMi Valet, Self-Parking, LAZ, Access Control, Unified, Custom, Premium, Bell Desk
Stage                   — Project stage (Design, Development, Kickoff, etc.)
Size                    — XL, L, M, S
Contract_Value          — Currency
Assigned_PM             — Project Manager name
Legal_Billing_Name      — Legal entity for billing
Client_Operator         — Client / operator name (use this, NOT Client_Name)
Sales_Rep               — Sales rep name
Location_Name           — Property/site name
Street_Address
City
State
ZIP
Country
PandaDoc_Number         — e.g. P-4790
QB_Proposal_Number
Account_ID
Contract_Signed_Date    — DateTime
Payment_Model
Go_Live_Date            — DateTime
Teams_Channel_Link      — URL to Teams channel
Planner_Link            — URL to Planner board
Planner_Progress        — Text
SKU_Summary
Notes
```

### `AVPMi Closed Projects` (Historical — 175 items)
```
Title                   — Project name (NOT Project_Title)
PandaDoc_Number
Project_Type            — Same choices as above
Client_Operator         — Client name (NOT Client_Name)
Legal_Billing_Name
Sales_Rep
Contract_Value          — Currency
State
Contract_Signed_Date    — DateTime
SKU_Summary
Teams_Channel_Link
```
⚠️ **The Closed Projects dashboard previously had a bug where it used `Project_Title` and `Client_Name` — those fields DO NOT EXIST. Always use `Title` and `Client_Operator`.**

---

## Task 1: Wire PMO Hub DataService to Live SharePoint

Replace the inline data in `DataService.load()` with this pattern (mirrors what works in the Closed Projects dashboard):

```javascript
const DataService = {
  _data: [],
  _msalInstance: null,

  async _getToken() {
    if (!this._msalInstance) {
      this._msalInstance = new msal.PublicClientApplication({
        auth: { clientId: MSAL_CONFIG.clientId, authority: MSAL_CONFIG.authority, redirectUri: MSAL_CONFIG.redirectUri },
        cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false }
      });
      await this._msalInstance.initialize();
    }
    const account = this._msalInstance.getAllAccounts()[0];
    try {
      const t = await this._msalInstance.acquireTokenSilent({ scopes: SP_SCOPES, account });
      return t.accessToken;
    } catch(e) {
      const t = await this._msalInstance.acquireTokenPopup({ scopes: SP_SCOPES });
      return t.accessToken;
    }
  },

  async load() {
    const token = await this._getToken();
    const fields = [
      "Title","Project_Number","Project_Type","Stage","Size","Contract_Value",
      "Assigned_PM","Client_Operator","Legal_Billing_Name","Sales_Rep",
      "State","City","Contract_Signed_Date","PandaDoc_Number","Payment_Model",
      "Go_Live_Date","Teams_Channel_Link","Planner_Progress","SKU_Summary","Notes"
    ].join(",");
    const headers = { Authorization: "Bearer " + token, Accept: "application/json;odata=nometadata" };
    let url = `${SP_BASE}/_api/web/lists/getbytitle('AVPMi Project Intake')/items?$select=${fields}&$top=500`;
    let allItems = [];
    while (url) {
      const resp = await fetch(url, { headers });
      if (!resp.ok) {
        const err = await resp.json().catch(() => ({}));
        throw new Error(`SharePoint ${resp.status}: ${err?.['odata.error']?.message?.value || resp.statusText}`);
      }
      const data = await resp.json();
      allItems = allItems.concat(data.value || []);
      url = data['odata.nextLink'] || null;
    }
    // Normalize fields to match dashboard expectations
    this._data = allItems.map(item => ({
      ...item,
      Client_Name: item.Client_Operator,  // backward compat alias
      Contract_Value: parseFloat(item.Contract_Value) || 0
    }));
  },

  getAll()    { return this._data || []; },
  getActive() { return this.getAll().filter(p => p.Stage !== 'Close' && p.Stage !== 'Closed'); },
  getClosed() { return this.getAll().filter(p => p.Stage === 'Close' || p.Stage === 'Closed'); }
};
```

Then update the init call at the bottom of the page:
```javascript
// Replace: await DataService.load();  (which currently loads static data)
// With a proper try/catch and loading indicator:
async function init() {
  document.getElementById('refresh-date').textContent = 'Loading...';
  try {
    await DataService.load();
    const projects = DataService.getAll();
    // ... rest of render calls
    document.getElementById('refresh-date').textContent = new Date().toLocaleString('en-US', {month:'short',day:'numeric',hour:'2-digit',minute:'2-digit'});
  } catch(e) {
    console.error(e);
    document.getElementById('refresh-date').textContent = 'Error loading — ' + e.message.substring(0, 60);
  }
}
init();
```

---

## Task 2: Add "Closed Projects" Section

The PMO hub already has a top nav bar with: Dashboard | Intake List | PowerApps (coming soon) | Reports (coming soon).

Add a "Closed Projects" nav link that navigates to the Closed Projects dashboard:
```html
<a href="https://avpmi-ops.github.io/closed-projects-dashboard/" target="_blank" title="Closed Projects Archive">
  <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
    <polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/>
  </svg>
  <span>Closed Projects</span>
</a>
```

**OR (preferred for a fully merged single page):**

Add a 5th tab to the tab bar: `<button onclick="switchTab('closed',this)">📦 Closed Projects</button>`

Then wire the "Closed Projects" tab to also load data from `AVPMi Closed Projects` using the same MSAL token — same `_getToken()` method, just a different list name and `$select` fields.

---

## Task 3: Deploy to SharePoint (Team Access)

### Step 1 — Upload the merged HTML file to SharePoint Site Assets
1. Open SharePoint: `https://avpm-my.sharepoint.com/personal/robert_edgar_avpmi_com`
2. Go to **Site Contents → Site Assets**
3. Upload the merged HTML file as `avpmi-hub.html`
4. The URL will be: `https://avpm-my.sharepoint.com/personal/robert_edgar_avpmi_com/SiteAssets/avpmi-hub.html`

### Step 2 — Register the SharePoint URL as a Redirect URI in Azure AD
1. Go to `portal.azure.com` → **Azure Active Directory** → **App registrations**
2. Open **AVPMi Dashboard** (Client ID: `85de8e7e-29c5-402b-8604-8483b569d8e5`)
3. **Authentication** → Under **Single-page application** → **Add URI**
4. Add: `https://avpm-my.sharepoint.com/personal/robert_edgar_avpmi_com/SiteAssets/avpmi-hub.html`
5. Save

### Step 3 — Share the SharePoint site with the team
For team members to access the lists (Intake + Closed Projects), Rob must share his personal site with them:
1. Go to the SharePoint site: `https://avpm-my.sharepoint.com/personal/robert_edgar_avpmi_com`
2. **Settings (gear)** → **Site permissions** → **Share site**
3. Add each team member — grant **Read** access minimum (PMs may want **Edit**)
4. Send them the hub URL: `https://avpm-my.sharepoint.com/personal/robert_edgar_avpmi_com/SiteAssets/avpmi-hub.html`

### Step 4 — Pin it in Teams (optional but highly recommended)
1. Open the relevant Teams team/channel
2. **+** (Add a tab) → **Website**
3. Name: `AVPMi Hub`, URL: `https://avpm-my.sharepoint.com/personal/robert_edgar_avpmi_com/SiteAssets/avpmi-hub.html`
4. Team members auto-authenticate via their M365 session — no extra login needed

---

## Architecture Summary

```
AVPMi Unified Hub (avpmi-hub.html)
├── Top Nav: Dashboard | Intake List | Closed Projects | Reports
├── Tab: Overview          → DataService → AVPMi Project Intake list
├── Tab: All Projects      → DataService → AVPMi Project Intake list
├── Tab: Financial         → DataService → AVPMi Project Intake list
├── Tab: Team Workload     → DataService → AVPMi Project Intake list
└── Tab: Closed Projects   → ClosedDataService → AVPMi Closed Projects list

Both DataServices use the SAME MSAL token (one auth popup for both lists).

Hosted at:
  SharePoint: https://avpm-my.sharepoint.com/personal/robert_edgar_avpmi_com/SiteAssets/avpmi-hub.html
  Teams Tab: Pinned in relevant Teams channel
  GitHub (Closed Projects only, existing): https://avpmi-ops.github.io/closed-projects-dashboard/
```

---

## Key Gotchas

1. **Never use `$orderby` in SharePoint REST API** — returns 400 if column isn't indexed. Sort client-side instead.
2. **Field names are case-sensitive** — `Title` not `title`, `Client_Operator` not `client_operator`
3. **The Intake list has `Client_Operator` for the client name** — NOT `Client_Name` (doesn't exist)
4. **The Closed Projects list has `Title` for the project title** — NOT `Project_Title` (doesn't exist)
5. **Redirect URIs must be exact** — include or exclude trailing slash consistently; `window.location.href.split(/[?#]/)[0]` handles this dynamically
6. **Team members need READ access** to the personal SharePoint site — Rob must explicitly share it
7. **Admin consent is NOT needed** for this setup — delegated `AllSites.Read` works without it

---

## Files Reference

| File | Location | Status |
|------|----------|--------|
| Closed Projects Dashboard | `avpmi-ops/closed-projects-dashboard` (GitHub Pages) | ✅ Live |
| PMO Hub (prototype) | `C:/Users/RobertEdgar/Documents/avpmi-pmo-dashboard.html` | ⚠️ Static data |
| Merged Hub (to create) | `avpmi-hub.html` → upload to SharePoint Site Assets | 🔲 To build |
