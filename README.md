# Daily OOF — Outlook Add-in

An Outlook Add-in for managing Exchange Online Out of Office (OOF) auto-reply messages directly from Outlook. Built with React, TypeScript, Fluent UI, and Microsoft Graph API.

> **Migrated from:** [DailyOOF (PowerShell GUI)](https://github.com/MicrosoftAzureAaron/DailyOOF)

---

## Features

| Tab | Description |
|---|---|
| **Actions** | Connect via Microsoft Graph, enable scheduled auto-reply, set vacation OOF with a return date, view current status. |
| **Config** | Set your full name, role, office hours, work days (with presets), and signature options. Syncs across devices via Office roaming settings. |
| **Templates** | Load built-in HTML templates (Normal, Vacation, Sick, Holiday), edit raw HTML, live preview, apply as internal/external/both. Import your current online message. |
| **Current** | View your live OOF message rendered in an embedded preview. |

### Key Improvements Over the PowerShell Version

- **No module installs** — Uses Microsoft Graph REST API instead of ExchangeOnlineManagement
- **No scheduled task needed** — Graph API handles OOF scheduling natively
- **Cross-platform** — Works in Outlook desktop (Windows/Mac), Outlook on the web, and mobile
- **SSO** — Users are already signed into Outlook, minimal auth friction
- **Roaming settings** — Config syncs across devices automatically

---

## Prerequisites

- **Node.js 18+** and **npm**
- An **Azure AD / Entra ID** app registration with these delegated permissions:
  - `User.Read`
  - `MailboxSettings.ReadWrite`

---

## Install (End Users)

The add-in is hosted on GitHub Pages — no build steps required.

### 1. Register an Entra ID App

> Your tenant admin may need to do this once. All users in the org can then use the add-in.

1. Go to [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. Click **New registration**
   - Name: `Daily OOF Outlook Add-in`
   - Supported account types: **Accounts in any organizational directory**
   - Redirect URI: **Single-page application (SPA)** → `https://microsoftazureaaron.github.io`
3. After creation, copy the **Application (client) ID**
4. Go to **API permissions** → Add:
   - `User.Read` (delegated)
   - `MailboxSettings.ReadWrite` (delegated)
5. Click **Grant admin consent** (or have your admin do this)

### 2. Sideload the Add-in

#### Outlook on the Web
1. Download [`manifest.xml`](https://raw.githubusercontent.com/MicrosoftAzureAaron/DailyOOF-Outlook/main/manifest.xml) from this repo
2. Go to **Outlook on the web** → **Settings** (gear icon) → **View all Outlook settings**
3. Go to **Mail** → **Customize actions** → **Get add-ins** → **My add-ins**
4. Click **Add a custom add-in** → **Add from file...**
5. Select the downloaded `manifest.xml`

#### Org-Wide Deployment (Admin)
1. Go to [Microsoft 365 admin center → Integrated Apps](https://admin.microsoft.com/Adminportal/Home#/Settings/IntegratedApps)
2. Click **Upload custom apps** → **Office Add-in** → **Upload manifest file**
3. Upload the `manifest.xml` and assign to users/groups

---

## Development Setup

If you want to run the add-in locally or contribute to the project:

### 1. Install Node.js

Download from [nodejs.org](https://nodejs.org/) or:

```powershell
winget install OpenJS.NodeJS.LTS
```

### 2. Clone and Install

```bash
git clone https://github.com/MicrosoftAzureAaron/DailyOOF-Outlook
cd DailyOOF-Outlook
npm install
```

### 3. Configure the Client ID

Open `src/taskpane/services/authService.ts` and replace:

```typescript
clientId: "YOUR_CLIENT_ID_HERE",
```

with your Entra ID Application (client) ID.

### 4. Start Development Server

```bash
npm run dev
```

This starts a local HTTPS server at `https://localhost:3000`.

> **Note:** For local development, the `manifest.xml` URLs need to point to `https://localhost:3000` instead of GitHub Pages. You can use the dev manifest or temporarily modify the URLs.

### 5. Sideload for Testing

#### Outlook on the Web
1. Go to **Outlook on the web** → **Settings** → **Get add-ins** → **My add-ins**
2. Click **Add a custom add-in** → **Add from file...**
3. Select the `manifest.xml` file from this project

#### Outlook Desktop (Windows)
```bash
npm run start
```

This uses the Office tooling to sideload the manifest automatically.

---

## Project Structure

```
DailyOOF-Outlook/
├── manifest.xml                    # Outlook add-in manifest
├── package.json                    # Dependencies and scripts
├── webpack.config.js               # Build configuration
├── tsconfig.json                   # TypeScript configuration
├── assets/                         # Add-in icons (replace with your own)
├── templates/                      # HTML email templates
│   ├── normal_oof.html
│   ├── vacation_oof.html
│   ├── sick_oof.html
│   └── holiday_oof.html
└── src/
    ├── commands/
    │   └── commands.ts             # Ribbon button handlers (future)
    └── taskpane/
        ├── index.html              # HTML entry point
        ├── index.tsx               # React entry point
        ├── App.tsx                 # Main app with tab navigation
        ├── styles/
        │   └── App.css             # Styling
        ├── types/
        │   └── index.ts            # TypeScript types and constants
        ├── services/
        │   ├── authService.ts      # MSAL authentication + SSO
        │   ├── graphService.ts     # Microsoft Graph API calls
        │   ├── configService.ts    # Office roaming settings persistence
        │   └── templateService.ts  # Template loading and placeholder resolution
        └── components/
            ├── QuickActions.tsx     # Connect, schedule, vacation, status
            ├── Configuration.tsx   # Identity, hours, days, signature options
            ├── MessageTemplates.tsx # Template editor with preview
            └── CurrentOOF.tsx      # Live OOF message viewer
```

---

## Template Placeholders

| Placeholder | Replaced with |
|---|---|
| `[ROLE]` | Your configured role / job title |
| `[RETURN DATE]` | The return date selected in the vacation date picker |
| `[SIGNATURE]` | Auto-generated signature: name, office hours, timezone, work days, email |

---

## Deployment

The add-in is automatically built and deployed to **GitHub Pages** on every push to `main` via GitHub Actions.

- **Live URL:** `https://microsoftazureaaron.github.io/DailyOOF-Outlook/`
- **Workflow:** [`.github/workflows/deploy.yml`](.github/workflows/deploy.yml)

To build manually:

```bash
npm run build
# Output is in /dist
```

Then update all URLs in `manifest.xml` to your production URL and distribute via the Microsoft 365 admin center.

---

## License

MIT
