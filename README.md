# PTC Gauge Cluster (GitHub Pages)
A car-style gauge dashboard that shows **counts only** (no message content).

## What it shows
- OUTLOOK = Inbox unread count
- SLACK / HUBSPOT / MONDAY = unread counts in Outlook folders that receive notification emails
- TOTAL = sum of the four gauges

## Setup (10 minutes)
### 1) Create Outlook folders (under Inbox)
Create these folders (or change names in app.js):
- PTC - Slack Alerts
- PTC - HubSpot Alerts
- PTC - Monday Alerts

### 2) Route notifications into those folders
Enable email notifications in each platform (Slack/HubSpot/Monday) and create Outlook rules that move those emails into the matching folder.

### 3) Create Entra (Azure AD) App Registration (SPA)
In Microsoft Entra admin center:
- App registrations -> New registration
- Add platform -> Single-page application (SPA)
- Redirect URI: https://YOURUSER.github.io/YOURREPO/

API Permissions:
- Microsoft Graph -> Delegated -> Mail.ReadBasic
- Grant admin consent if your org requires it

### 4) Configure app.js
Open `app.js` and set:
- tenantId
- clientId

Optionally adjust:
- folder names
- gauge max/redline thresholds
- app links

### 5) Enable GitHub Pages
Repo Settings -> Pages:
- Source: Deploy from branch
- Branch: main / root
Save, then open the provided URL.

## Notes
- This site reads ONLY folder metadata (unread counts). It does not fetch message bodies.
- TRIP RESET stores a baseline in localStorage per signed-in user.
