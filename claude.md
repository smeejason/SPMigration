# SP Migration Planner - Project Context
## What This Project Is
A **browser-based SPA** (TypeScript + Vite) that helps plan and orchestrate SharePoint Online migrations. Users sign in with their M365 account, upload a TreeSize report, and visually map file system nodes to SharePoint sites and document libraries. If target sites don't exist, the dashboard can create them.
## Architecture
- **Frontend**: Static SPA — TypeScript + Vite, vanilla TS with DOM manipulation (no framework)
- **Hosting**: Azure Static Web Apps (free tier)
- **Backend**: None — all calls go directly to Microsoft Graph API from the browser
- **Auth**: Delegated (MSAL.js + PKCE) — user signs in via browser, no secrets in code
- **Data Storage**: SharePoint list (`MigrationProjects`) on a dedicated site (`sites/SPMigrationPlanner`)
## Key Dependencies
- `@azure/msal-browser` — Delegated auth (PKCE flow)
- `@microsoft/microsoft-graph-client` — Graph API calls for SharePoint
- `exceljs` — Parse TreeSize Excel exports (browser)
- `papaparse` — Parse TreeSize CSV exports (browser)
- Dev: `vite`, `typescript`
## Directory Structure
```
src/
  main.ts                    # App entry point
  auth/
    msalConfig.ts            # MSAL browser configuration
    authService.ts           # Sign-in/out, token acquisition
  graph/
    graphClient.ts           # Graph API wrapper (sites, drives)
    projectService.ts        # CRUD for MigrationProjects list
  parsers/
    treeSizeParser.ts        # TreeSize CSV/Excel parser
    fileDetector.ts          # Auto-detect CSV vs Excel
  ui/
    components/
      authPanel.ts           # Sign-in/sign-out UI
      projectList.ts         # Projects dashboard (after login)
      projectForm.ts         # Create/edit project form
      uploadPanel.ts         # File upload drop zone
      treeView.ts            # TreeSize data as interactive tree
      mappingPanel.ts        # Drag/drop or select mapping UI
      siteCreator.ts         # Create new sites form
      summaryPanel.ts        # Migration plan summary/export
    app.ts                   # Main app shell / layout
  state/
    store.ts                 # Simple observable state management
  types/
    index.ts                 # Shared TypeScript interfaces
```
## Data Model
Projects are stored in a SharePoint list (`MigrationProjects`) with columns:
- `Title` — Project name
- `Description` — Notes
- `Status` — Choice: Planning, In Progress, Completed, On Hold
- `Owners` — Person (multi-value)
- `ProjectData` — JSON blob for tree data, mappings, and settings
## User Flow
1. User signs in via Microsoft (MSAL popup/redirect)
2. Projects page — shows migration projects the user owns
3. User selects or creates a project
4. Inside a project: Upload TreeSize report -> Map folders to SharePoint sites/libraries -> Create new sites if needed -> Review summary and export plan
## Graph API Permissions (Delegated)
- `Sites.ReadWrite.All` — Manage SharePoint sites and list items
- `Sites.Manage.All` — Create new site collections
- `User.Read` — Basic profile
- `People.Read` — Search users for project owners
## Build Order
1. Vite + TS scaffolding
2. Types/interfaces
3. MSAL config + auth service
4. Auth panel UI
5. Graph client wrapper
6. Project service (CRUD for list)
7. Projects page UI
8. TreeSize CSV parser
9. TreeSize Excel parser
10. Upload panel + tree view
11. Site search + library listing
12. Mapping panel UI
13. Site creation via Graph
14. Site creator UI
15. Summary panel + export
16. Save project state to SharePoint
17. Azure Static Web Apps deploy
## Commands
- `npm run dev` — Start Vite dev server on localhost:5173
- `npm run build` — Production build (static output)
## Notes
- No backend/server — everything runs client-side
- App Registration must be set up manually in Azure portal (SPA redirect, delegated permissions, no client secret)
- The full implementation plan is in `plan.md`
