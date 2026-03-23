# SharePoint Migration Planner — Implementation Plan
## Overview
A **browser-based dashboard** (static SPA) that helps plan and orchestrate SharePoint Online migrations. The user signs in with their M365 account, uploads a TreeSize report, and visually maps file system nodes to SharePoint sites and document libraries. If target sites don't exist, the dashboard can create them.
**Architecture**: Static SPA (TypeScript + Vite) → Azure Static Web Apps (free) → Microsoft Graph API
**Auth**: Delegated (MSAL.js + PKCE) — user signs in via browser, no secrets in code
**Hosting**: Azure Static Web Apps free tier, deployed in the client's Azure tenant
---
## User Flow
1. User opens the app → redirected to Microsoft sign-in (MSAL popup/redirect)
2. After login → **Projects page** — shows all migration projects the user owns
3. User selects an existing project or creates a new one
4. Inside a project → the existing workflow (upload TreeSize, map to SharePoint, etc.)
---
## App Data Storage — SharePoint Site & List
### SharePoint Site
A dedicated SharePoint site hosts the app's data. Initially created manually by an admin; auto-provisioning can be added later.
- **Site URL**: `https://{tenant}.sharepoint.com/sites/SPMigrationPlanner`
- **Purpose**: Central store for all migration project metadata
### MigrationProjects List Schema
| Column | Type | Description |
|---|---|---|
| `Title` | Single line of text (built-in) | Project name |
| `Description` | Multiple lines of text | Project description / notes |
| `Status` | Choice | `Planning`, `In Progress`, `Completed`, `On Hold` |
| `Owners` | Person or Group (multi-value) | Users who own this project |
| `ProjectData` | Multiple lines of text (plain) | JSON blob for flexible/extensible data (source paths, target URLs, mappings, settings, etc.) |
**Why a JSON column?** Avoids constant list schema changes as features evolve. The app reads/writes structured JSON while SharePoint just stores it as text. Structured columns (`Title`, `Status`, `Owners`) remain queryable/filterable at the list level.
### Manual Setup Instructions (Phase 1)
1. Create a SharePoint site: `sites/SPMigrationPlanner` (Team site or Communication site)
2. Create a list called `MigrationProjects`
3. Add columns: `Description` (multi-line text), `Status` (choice with values above), `Owners` (person, allow multiple), `ProjectData` (multi-line plain text)
4. Grant access to users who will use the app
### Graph API Operations for Projects
| Method | Graph API Call | Purpose |
|---|---|---|
| `getProjects()` | `GET /sites/{siteId}/lists/{listId}/items?$filter=...&$expand=fields` | Fetch projects where current user is an owner |
| `getProject(id)` | `GET /sites/{siteId}/lists/{listId}/items/{id}?$expand=fields` | Fetch single project |
| `createProject(data)` | `POST /sites/{siteId}/lists/{listId}/items` | Create new project |
| `updateProject(id, data)` | `PATCH /sites/{siteId}/lists/{listId}/items/{id}` | Update project fields (including JSON blob) |
| `deleteProject(id)` | `DELETE /sites/{siteId}/lists/{listId}/items/{id}` | Delete a project |
---
## Phase 1: Project Scaffolding
### 1.1 Initialize Vite + TypeScript project
```bash
npm create vite@latest sp-migration-planner -- --template vanilla-ts
```
- Vite for dev server + production build (outputs static files)
- TypeScript for type safety
- No framework initially — vanilla TS with DOM manipulation (lightweight)
- Can add a lightweight UI library later if needed (e.g., Lit, Preact)
### 1.2 Directory structure
```
sp-migration-planner/
├── src/
│   ├── main.ts                    # App entry point
│   ├── auth/
│   │   ├── msalConfig.ts          # MSAL browser configuration
│   │   └── authService.ts         # Sign-in/out, token acquisition
│   ├── graph/
│   │   ├── graphClient.ts         # Graph API wrapper (sites, drives)
│   │   └── projectService.ts      # CRUD for MigrationProjects list
│   ├── parsers/
│   │   ├── treeSizeParser.ts      # TreeSize CSV/Excel parser
│   │   └── fileDetector.ts        # Auto-detect CSV vs Excel
│   ├── ui/
│   │   ├── components/
│   │   │   ├── authPanel.ts       # Sign-in/sign-out UI
│   │   │   ├── projectList.ts     # Projects dashboard (after login)
│   │   │   ├── projectForm.ts     # Create/edit project form
│   │   │   ├── uploadPanel.ts     # File upload drop zone
│   │   │   ├── treeView.ts        # TreeSize data as interactive tree
│   │   │   ├── mappingPanel.ts    # Drag/drop or select mapping UI
│   │   │   ├── siteCreator.ts     # Create new sites form
│   │   │   └── summaryPanel.ts    # Migration plan summary/export
│   │   └── app.ts                 # Main app shell / layout
│   ├── state/
│   │   └── store.ts               # Simple state management
│   └── types/
│       └── index.ts               # Shared TypeScript interfaces
├── index.html                     # Single HTML entry point
├── public/
│   └── favicon.svg
├── package.json
├── tsconfig.json
├── vite.config.ts
└── staticwebapp.config.json       # Azure Static Web Apps config
```
### 1.3 Dependencies
| Package | Purpose |
|---|---|
| `@azure/msal-browser` | Delegated auth (PKCE flow) in the browser |
| `@microsoft/microsoft-graph-client` | Graph API calls for SharePoint |
| `exceljs` | Parse TreeSize Excel exports (works in browser) |
| `papaparse` | Parse TreeSize CSV exports (browser-native, lightweight) |
Dev dependencies: `vite`, `typescript`
> **Note**: No backend dependencies. No Express, no dotenv, no winston. Everything runs client-side.
---
## Phase 2: Authentication (MSAL.js Delegated)
### 2.1 App Registration Setup (manual step, documented in README)
In the client's Azure portal:
1. Register a new app (e.g., "SP Migration Planner")
2. Set **SPA** redirect URI: `http://localhost:5173` (dev) + production URL
3. Add **delegated** permissions:
   - `Sites.ReadWrite.All` — browse and manage SharePoint sites (includes reading/writing list items in the app's MigrationProjects list)
   - `Sites.Manage.All` — create new site collections (if needed)
   - `User.Read` — basic profile for the signed-in user
   - `People.Read` — search for users when adding project owners
4. **No client secret needed** — SPA uses PKCE
### 2.2 MSAL Configuration (`src/auth/msalConfig.ts`)
```typescript
// Configuration object — client ID and tenant ID are the only required values
{
  auth: {
    clientId: "<from-env-or-config>",
    authority: "https://login.microsoftonline.com/<tenant-id>",
    redirectUri: window.location.origin
  }
}
```
### 2.3 Auth Service (`src/auth/authService.ts`)
- `signIn()` — popup or redirect login
- `signOut()` — clear session
- `getToken(scopes)` — acquire token silently (with fallback to interactive)
- `isAuthenticated()` — check current auth state
- Expose the signed-in user's display name and tenant for the UI header
---
## Phase 3: TreeSize Report Parsing
### 3.1 TreeSize Export Format
TreeSize Pro/Free exports typically contain:
| Column | Description |
|---|---|
| `Path` | Full folder/file path (e.g., `\\server\share\Projects\2024`) |
| `Size` | Size in bytes or human-readable |
| `Files` | Number of files in folder |
| `Folders` | Number of subfolders |
| `% of Parent` | Percentage of parent folder size |
| `Last Change` | Last modified date |
### 3.2 Parser Implementation (`src/parsers/treeSizeParser.ts`)
- Accept uploaded `File` object from browser file input
- Auto-detect CSV (`.csv`) vs Excel (`.xlsx`) via extension
- **CSV**: Use `papaparse` — browser-native, streaming capable
- **Excel**: Use `exceljs` — load workbook from `ArrayBuffer`
- Normalize data into a typed `TreeNode[]` structure:
  ```typescript
  interface TreeNode {
    path: string;           // Original file system path
    name: string;           // Folder/file name
    depth: number;          // Nesting level
    sizeBytes: number;      // Size in bytes
    fileCount: number;      // Number of files
    folderCount: number;    // Number of subfolders
    lastModified?: Date;    // Last change date
    children: TreeNode[];   // Nested children (built from paths)
  }
  ```
- Build a **tree structure** from flat path data (split paths, nest by hierarchy)
### 3.3 Validation
- Check for required columns (`Path` at minimum)
- Handle different TreeSize versions/export formats gracefully
- Show warnings for unrecognized columns
---
## Phase 4: Dashboard UI
### 4.1 App Layout (`src/ui/app.ts`)
Two-level navigation: **Projects list** → **Project workspace**
**Projects Page** (after login):
```
┌─────────────────────────────────────────────────────┐
│  SP Migration Planner          [User Name] [Logout] │
├─────────────────────────────────────────────────────┤
│                                                     │
│  Your Migration Projects              [+ New Project]│
│                                                     │
│  ┌──────────────────────────────────────────────┐   │
│  │ 📋 Contoso File Share Migration    Planning  │   │
│  │    2 owners · 340 GB · Last updated Mar 5    │   │
│  ├──────────────────────────────────────────────┤   │
│  │ 📋 Marketing Archive Move        In Progress │   │
│  │    1 owner · 120 GB · Last updated Mar 3     │   │
│  └──────────────────────────────────────────────┘   │
│                                                     │
└─────────────────────────────────────────────────────┘
```
**Project Workspace** (inside a project):
```
┌─────────────────────────────────────────────────────┐
│  ← Projects   Contoso File Share     [User] [Logout]│
├─────────────────────────────────────────────────────┤
│  Upload  │  Map  │  Summary                         │
├────────────────────────┬────────────────────────────┤
│                        │                            │
│   [Active Panel]       │   [Context Panel]          │
│                        │                            │
└────────────────────────┴────────────────────────────┘
```
### 4.2 Login & Auth Panel
- "Sign in with Microsoft" button
- On successful sign-in → navigate to Projects page
- Tests Graph API connectivity (fetches root site)
- Status indicator: connected / disconnected
### 4.3 Projects Page (`src/ui/components/projectList.ts`)
- Fetches projects from `MigrationProjects` list where current user is in the `Owners` column
- Displays project cards with: name, status badge, owner count, summary stats (from JSON)
- **"+ New Project"** button → opens project creation form
- Click a project → enters the project workspace
- Project cards show status as color-coded badges
### 4.4 Project Form (`src/ui/components/projectForm.ts`)
- Fields: Project name, description, status (dropdown)
- Owners: people picker (uses Graph to search users)
- Creates/updates the list item via Graph API
- On create → navigates into the new project workspace
### 4.5 Upload TreeSize Report (Upload Panel — inside project)
- Drag-and-drop zone or file picker
- Accepts `.csv` and `.xlsx` files
- Parses on upload and shows preview:
  - Total folders/files detected
  - Total size
  - Root path
- Renders parsed data as an **interactive tree view** (expandable/collapsible)
- Parsed data is saved into the project's `ProjectData` JSON column
### 4.6 Map to SharePoint (Mapping Panel — inside project)
This is the core of the dashboard. Split view:
```
┌─ TreeSize Structure ──────────┬─ SharePoint Target ────────────┐
│                               │                                │
│  ▼ \\Server\Share             │                                │
│    ▼ Projects                 │  → contoso.sharepoint.com/     │
│      ▼ Engineering            │      sites/Engineering         │
│        📁 Docs (45 GB)  [Map]│      └─ Documents/Docs         │
│        📁 Data (12 GB)  [Map]│      └─ Documents/Data         │
│      ▼ Marketing              │    sites/Marketing             │
│        📁 Assets (8 GB) [Map]│      └─ Documents/Assets       │
│    ▼ Archive                  │  → (not mapped)                │
│                               │                                │
└───────────────────────────────┴────────────────────────────────┘
```
**Mapping workflow**:
1. User selects a TreeSize node (folder)
2. Right panel shows:
   - **Existing sites**: Fetched from Graph API (`GET /sites?search=*`), filterable
   - **Existing document libraries**: Fetched for the selected site
   - **"Create new site" option**: Opens site creation form
3. User picks a target site + document library + optional subfolder path
4. Mapping is saved in state and shown visually
### 4.7 Create New Sites (Site Creator — inside project)
When a target site doesn't exist:
- Form fields: Site name, URL suffix, description, template (Team site / Communication site)
- Calls Graph API: `POST /sites` (or SharePoint admin API)
- Shows creation status (pending / created / failed)
- Newly created site becomes available in the mapping dropdown
### 4.8 Summary & Export (Summary Panel — inside project)
- Table view of all mappings:
  | Source Path | Size | Files | Target Site | Target Library | Target Path | Status |
  |---|---|---|---|---|---|---|
  | `\\Server\Projects\Eng\Docs` | 45 GB | 1,200 | Engineering | Documents | /Docs | Ready |
- Totals: total data size, total files, number of sites to create
- **Export mapping plan** as CSV or JSON (for use by a future migration execution tool)
- Highlights unmapped nodes as warnings
---
## Phase 5: Graph API Integration
### 5.1 Graph Client (`src/graph/graphClient.ts`)
Wrapper methods using `@microsoft/microsoft-graph-client`:
| Method | Graph API Call | Purpose |
|---|---|---|
| `getCurrentUser()` | `GET /me` | Display signed-in user |
| `getRootSite()` | `GET /sites/root` | Verify connectivity + get tenant |
| `searchSites(query)` | `GET /sites?search={query}` | Search for existing sites |
| `getSiteDrives(siteId)` | `GET /sites/{id}/drives` | List document libraries |
| `getDriveContents(driveId, path)` | `GET /drives/{id}/root:/{path}:/children` | Browse library contents |
| `createTeamSite(name, alias)` | `POST /groups` (M365 group-connected) | Create team site |
| `createCommSite(name, url)` | SharePoint REST API or Graph beta | Create communication site |
### 5.2 Site Creation Details
- **Team sites**: Created via M365 Group (`POST /groups` with `groupTypes` including `Unified`)
- **Communication sites**: May require SharePoint-specific API (`/_api/SPSiteManager/create`) called via Graph proxy or direct REST
- Poll for provisioning completion before showing as "ready"
---
## Phase 6: State Management
### 6.1 Simple Store (`src/state/store.ts`)
No framework needed — a simple observable store pattern:
```typescript
interface AppState {
  auth: { user: User | null; isAuthenticated: boolean };
  currentProject: MigrationProject | null; // Active project from SharePoint list
  treeData: TreeNode | null;               // Parsed TreeSize data
  mappings: MigrationMapping[];            // Source → destination mappings
  sites: SharePointSite[];                 // Cached site list from Graph
  pendingSiteCreations: SiteRequest[];     // Sites queued for creation
}
```
- State changes trigger UI re-renders for affected panels
- **Primary persistence**: SharePoint `MigrationProjects` list — project data (tree, mappings, settings) saved as JSON in the `ProjectData` column
- **Session cache**: `sessionStorage` used as a working cache to avoid re-fetching on every interaction; synced back to SharePoint on save/navigation
---
## Phase 7: Hosting & Deployment
### 7.1 Azure Static Web Apps
- `staticwebapp.config.json` for SPA routing (fallback to `index.html`)
- Deploy via GitHub Actions (auto-configured by Azure) or Azure CLI
- Free tier includes: custom domain, SSL, global CDN, auth integration
- No backend, no Azure Functions needed for Phase 1
### 7.2 Local Development
- `npm run dev` — Vite dev server on `localhost:5173`
- Hot module reload for fast iteration
- App Registration redirect URI includes `http://localhost:5173` for local auth
---
## Build Order (Implementation Sequence)
| Step | What | Depends on | Deliverable |
|---|---|---|---|
| 1 | Vite + TS scaffolding, dependencies | — | Empty app builds and runs |
| 2 | Types / interfaces (incl. project types) | — | `types/index.ts` |
| 3 | MSAL config + auth service | Step 1 | Sign-in/out works |
| 4 | Auth panel UI (sign-in button) | Step 3 | User can authenticate |
| 5 | Graph client wrapper | Step 3 | Can call Graph API |
| 6 | Project service (CRUD for MigrationProjects list) | Step 5 | Can read/write projects via Graph |
| 7 | Projects page UI (list + create/edit form) | Step 6 | User sees their projects after login |
| 8 | TreeSize CSV parser | Step 2 | Parses CSV to TreeNode[] |
| 9 | TreeSize Excel parser | Step 2 | Parses Excel to TreeNode[] |
| 10 | Upload panel + tree view UI (inside project) | Step 7, 8, 9 | User can upload and see tree |
| 11 | Site search + library listing (Graph) | Step 5 | Can browse existing SharePoint sites |
| 12 | Mapping panel UI | Step 10, 11 | User can map nodes to sites |
| 13 | Site creation (Graph) | Step 5 | Can create new sites |
| 14 | Site creator UI | Step 13 | User can create sites from dashboard |
| 15 | Summary panel + export | Step 12 | User can review and export plan |
| 16 | Save project state to SharePoint (ProjectData JSON) | Step 6, 12 | All project data persists in SharePoint |
| 17 | Azure Static Web Apps config + deploy | Step 15 | Deployed to Azure |
---
## Out of Scope (Future Phases)
- **Migration execution** — actually moving files (Phase 2 of the larger project)
- **Progress tracking** — real-time migration progress dashboard
- **Permission mapping** — mapping NTFS permissions to SharePoint permissions
- **Metadata mapping** — preserving file metadata during migration
- **Incremental sync** — detecting changes and migrating only deltas
- **Multi-tenant support** — currently single-tenant per deployment
- **Azure Function backend** — only needed if app-only auth is required later
