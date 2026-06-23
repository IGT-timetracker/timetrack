# TimeTrack PWA — Setup & Deployment Guide

## What this is
A Progressive Web App (PWA) that employees install from any browser.
It reads and writes directly to your **Attendance.xlsx** on SharePoint via Microsoft Graph API.

---

## Option 1 — Host on GitHub Pages (Free, 5 minutes)

1. Create a free account at https://github.com
2. Create a new repository named `timetrack`
3. Upload all files from this folder
4. Go to **Settings → Pages → Source: main branch → / (root)**
5. Your app URL will be: `https://yourusername.github.io/timetrack`
6. Share this URL with employees

---

## Option 2 — Host on SharePoint itself

1. In SharePoint, go to **Site Contents → Site Pages** (or any document library)
2. Upload all files from this folder into a folder called `timetrack`
3. Share the URL: `https://igtplc.sharepoint.com/sites/APACManufacturingOperationsTeam/SitePages/timetrack/index.html`

---

## Option 3 — Host on Azure Static Web Apps (Free tier)

1. Create an Azure Static Web Apps resource
2. Connect your GitHub repo
3. Deploy automatically on every push

---

## How employees install it as a desktop app

### On Windows (Edge or Chrome):
1. Open the app URL in Edge or Chrome
2. Sign in with their Microsoft work account
3. Click the **install icon** in the address bar (or the "Install app" banner)
4. The app appears on their desktop and taskbar — works like a native app

### On Mac (Chrome):
1. Open the URL in Chrome
2. Click ⋮ menu → **Cast, save and share → Install page as app**

---

## First-time admin setup

1. Sign in and go to **Admin → Settings**
2. Confirm the SharePoint settings:
   - **Site name:** `APACManufacturingOperationsTeam`
   - **File path:** `General/ATTENDANCE/Attendance.xlsx`
3. Click **Test SharePoint connection** to verify
4. Click **Save settings**
5. Add your employees under the Employees section

---

## How the SharePoint Excel sync works

The app maintains 3 sheets in `Attendance.xlsx`:

| Sheet | Contents |
|-------|----------|
| `Employees` | Employee list with standard hours/areas |
| `Attendance` | All clock in/out records |
| `Settings` | App config (company name, recipient, etc.) |

- Every clock in/out **automatically syncs** to SharePoint
- On app open, it **pulls the latest data** from SharePoint
- Use **Admin → Push to SharePoint** for a manual sync
- Works **offline** — syncs when back online

---

## SharePoint file details
- **URL:** https://igtplc.sharepoint.com/sites/APACManufacturingOperationsTeam/Shared%20Documents/General/ATTENDANCE/Attendance.xlsx
- All employees need at least **read access** to the SharePoint site
- The app uses their own Microsoft login — no shared passwords

---

## Files in this package
```
timetrack-pwa/
├── index.html      ← Main app
├── app.js          ← All app logic + SharePoint API
├── manifest.json   ← PWA install config
├── sw.js           ← Service worker (offline support)
└── icons/
    ├── icon-192.png
    └── icon-512.png
```
