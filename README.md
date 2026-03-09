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

## Setup

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

### 3. Register an Entra ID App

1. Go to [Azure Portal → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. Click **New registration**
   - Name: `Daily OOF Outlook Add-in`
   - Supported account types: **Accounts in any organizational directory**
   - Redirect URI: **Single-page application (SPA)** → `https://localhost:3000`
3. After creation, copy the **Application (client) ID**
4. Go to **API permissions** → Add:
   - `User.Read` (delegated)
   - `MailboxSettings.ReadWrite` (delegated)
5. Click **Grant admin consent** (or have your admin do this)

### 4. Configure the Client ID

Open `src/taskpane/services/authService.ts` and replace:

```typescript
clientId: "YOUR_CLIENT_ID_HERE",
```

with your Application (client) ID from step 3.

### 5. Start Development Server

```bash
npm run dev
```

This starts a local HTTPS server at `https://localhost:3000`.

### 6. Sideload the Add-in

#### Outlook on the Web
1. Go to **Outlook on the web** → **Settings** (gear icon) → **View all Outlook settings**
2. Go to **Mail** → **Customize actions** → **Get add-ins** → **My add-ins**
3. Click **Add a custom add-in** → **Add from file...**
4. Select the `manifest.xml` file from this project

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

For production, host the built files on **Azure Static Web Apps** (free tier) or similar:

```bash
npm run build
# Deploy the contents of /dist to your hosting service
```

Then update all `https://localhost:3000` URLs in `manifest.xml` to your production URL and distribute via the Microsoft 365 admin center.

---

## License

MIT
