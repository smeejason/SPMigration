import { Client } from '@microsoft/microsoft-graph-client'
import { getToken } from '../auth/authService'
import type {
  AppUser,
  SharePointSite,
  SharePointDrive,
  SiteRequest,
  GraphSite,
  GraphDrive,
  GraphUser,
  MigrationMapping,
} from '../types'

// ─── Graph client factory ─────────────────────────────────────────────────────

function createClient(): Client {
  return Client.initWithMiddleware({
    authProvider: {
      getAccessToken: () => getToken(),
    },
  })
}

function client(): Client {
  return createClient()
}

// ─── User ─────────────────────────────────────────────────────────────────────

export async function getCurrentUser(): Promise<AppUser> {
  const user = await client().api('/me').get() as GraphUser
  return {
    id: user.id,
    displayName: user.displayName,
    mail: user.mail ?? user.userPrincipalName,
    userPrincipalName: user.userPrincipalName,
  }
}

// ─── Root site (connectivity check) ──────────────────────────────────────────

export async function getRootSite(): Promise<SharePointSite> {
  const site = await client().api('/sites/root').get() as GraphSite | null
  if (!site) {
    throw new Error(
      'Could not reach SharePoint root site. Ensure admin consent has been granted for Sites.ReadWrite.All in your Azure App Registration.'
    )
  }
  return mapSite(site)
}

// ─── Sites ────────────────────────────────────────────────────────────────────

export async function searchSites(query: string = '*'): Promise<SharePointSite[]> {
  const response = await client()
    .api('/sites')
    .query({ search: query })
    .get() as { value: GraphSite[] }
  return (response.value ?? []).map(mapSite)
}

export async function getSiteById(siteId: string): Promise<SharePointSite> {
  const site = await client().api(`/sites/${siteId}`).get() as GraphSite
  return mapSite(site)
}

// ─── Drives (document libraries) ─────────────────────────────────────────────

export async function getSiteDrives(siteId: string): Promise<SharePointDrive[]> {
  const response = await client()
    .api(`/sites/${siteId}/drives`)
    .get() as { value: GraphDrive[] }
  return (response.value ?? []).map(mapDrive)
}

// ─── Users (people search for owners picker) ─────────────────────────────────

export async function searchUsers(query: string): Promise<AppUser[]> {
  if (!query.trim()) return []
  const response = await client()
    .api('/users')
    .filter(`startsWith(displayName,'${query}') or startsWith(userPrincipalName,'${query}')`)
    .select('id,displayName,mail,userPrincipalName')
    .top(10)
    .get() as { value: GraphUser[] }
  return (response.value ?? []).map((u) => ({
    id: u.id,
    displayName: u.displayName,
    mail: u.mail ?? u.userPrincipalName,
    userPrincipalName: u.userPrincipalName,
  }))
}

// ─── Site creation ────────────────────────────────────────────────────────────

/**
 * Create a Team site via an M365 Unified Group.
 * The SharePoint site is provisioned automatically by Microsoft 365.
 */
export async function createTeamSite(request: SiteRequest): Promise<string> {
  const group = await client().api('/groups').post({
    displayName: request.displayName,
    mailNickname: request.alias,
    description: request.description,
    groupTypes: ['Unified'],
    mailEnabled: true,
    securityEnabled: false,
    visibility: 'Private',
  }) as { id: string }
  return group.id
}

/**
 * Poll until the SharePoint site linked to the M365 group is provisioned.
 * Returns the site's Graph ID.
 */
export async function waitForGroupSite(groupId: string, maxWaitMs = 120_000): Promise<SharePointSite> {
  const deadline = Date.now() + maxWaitMs
  while (Date.now() < deadline) {
    try {
      const site = await client()
        .api(`/groups/${groupId}/sites/root`)
        .get() as GraphSite
      return mapSite(site)
    } catch {
      await delay(5_000)
    }
  }
  throw new Error('Timed out waiting for site to provision')
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function mapSite(s: GraphSite): SharePointSite {
  return {
    id: s.id,
    name: s.name,
    displayName: s.displayName ?? s.name,
    webUrl: s.webUrl,
    description: s.description,
  }
}

function mapDrive(d: GraphDrive): SharePointDrive {
  return {
    id: d.id,
    name: d.name,
    webUrl: d.webUrl,
    driveType: d.driveType,
  }
}

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms))
}

// ─── Drive file operations ────────────────────────────────────────────────────
//
// Used for per-project Excel/CSV upload history.
// Files are stored in: Documents/SPMigration/{projectTitle}_{projectId}/

function sanitizeSegment(s: string): string {
  // Remove characters that SharePoint/OneDrive disallows in names
  return s.replace(/["*:<>?/\\|#%]/g, '_').replace(/\.+$/, '').trim()
}

export async function getOrCreateProjectFolder(
  siteId: string,
  projectTitle: string,
  projectId: string
): Promise<string> {
  const folderName = `${sanitizeSegment(projectTitle).slice(0, 60)}_${projectId}`

  // 1. Fast path — folder already exists
  try {
    const item = await client()
      .api(`/sites/${siteId}/drive/root:/SPMigration/${folderName}:`)
      .get() as { id: string }
    return item.id
  } catch { /* will create below */ }

  // 2. Ensure SPMigration parent exists (ignore 409 conflict = already exists)
  try {
    await client()
      .api(`/sites/${siteId}/drive/root/children`)
      .post({ name: 'SPMigration', folder: {}, '@microsoft.graph.conflictBehavior': 'fail' })
  } catch { /* already exists — that's fine */ }

  // 3. Create project subfolder
  const result = await client()
    .api(`/sites/${siteId}/drive/root:/SPMigration:/children`)
    .post({ name: folderName, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' }) as { id: string }
  return result.id
}

export async function uploadFileToDrive(
  siteId: string,
  folderId: string,
  fileName: string,
  content: ArrayBuffer | string
): Promise<string> {
  const token = await getToken()
  const safeFileName = encodeURIComponent(sanitizeSegment(fileName))
  const contentType = typeof content === 'string' ? 'application/json' : 'application/octet-stream'

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${folderId}:/${safeFileName}:/content`,
    {
      method: 'PUT',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': contentType },
      body: content,
    }
  )
  if (!response.ok) {
    const text = await response.text().catch(() => '')
    throw new Error(`File upload failed (${response.status}): ${text}`)
  }
  const item = await response.json() as { id: string }
  return item.id
}

export async function downloadDriveItem(siteId: string, itemId: string): Promise<unknown> {
  const token = await getToken()
  // Graph returns a redirect to the actual file content — fetch follows it automatically
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${itemId}/content`,
    { headers: { Authorization: `Bearer ${token}` } }
  )
  if (!response.ok) {
    throw new Error(`File download failed (${response.status})`)
  }
  const text = await response.text()
  return JSON.parse(text)
}

// ─── Mappings file helpers ────────────────────────────────────────────────────
//
// Mappings are stored as {projectId}.mappings.json in the project SP folder
// rather than inline in the list item field, avoiding the ~63 KB column limit.
// sourceNode.children is stripped on write — the tree is already in .tree.json.

export function getProjectFolderName(projectTitle: string, projectId: string): string {
  return `${sanitizeSegment(projectTitle).slice(0, 60)}_${projectId}`
}

export async function saveMappingsFile(
  siteId: string,
  projectTitle: string,
  projectId: string,
  mappings: MigrationMapping[]
): Promise<void> {
  // Strip children from each sourceNode — they are large and already in .tree.json
  const slim = mappings.map((m) => ({
    ...m,
    sourceNode: { ...m.sourceNode, children: [] },
  }))

  const token = await getToken()
  const folderName = getProjectFolderName(projectTitle, projectId)
  const filePath = encodeURIComponent(`SPMigration/${folderName}/${projectId}.mappings.json`)

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${filePath}:/content`,
    {
      method: 'PUT',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(slim),
    }
  )
  if (!response.ok) {
    const text = await response.text().catch(() => '')
    throw new Error(`Mappings save failed (${response.status}): ${text}`)
  }
}

export async function loadMappingsFile(
  siteId: string,
  projectTitle: string,
  projectId: string
): Promise<MigrationMapping[] | null> {
  const token = await getToken()
  const folderName = getProjectFolderName(projectTitle, projectId)
  const filePath = encodeURIComponent(`SPMigration/${folderName}/${projectId}.mappings.json`)

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${filePath}:/content`,
    { headers: { Authorization: `Bearer ${token}` } }
  )
  if (response.status === 404) return null
  if (!response.ok) throw new Error(`Mappings load failed (${response.status})`)
  const text = await response.text()
  return JSON.parse(text) as MigrationMapping[]
}
