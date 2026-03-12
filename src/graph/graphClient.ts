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
  OneDriveUserMapping,
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

// ─── OneDrive (personal drives) ───────────────────────────────────────────────

/**
 * Get a user's OneDrive drive object (id + webUrl).
 * Returns null if the user has no OneDrive or the token lacks permission.
 */
export async function getUserDrive(userId: string): Promise<SharePointDrive | null> {
  try {
    const drive = await client().api(`/users/${userId}/drive`).get() as GraphDrive
    return mapDrive(drive)
  } catch {
    return null
  }
}

/**
 * Fetch a single user's profile by ID (for recovering UPN after a mapping is loaded).
 */
export async function getUserById(userId: string): Promise<AppUser | null> {
  try {
    const user = await client()
      .api(`/users/${userId}`)
      .select('id,displayName,mail,userPrincipalName')
      .get() as GraphUser
    return mapGraphUser(user)
  } catch {
    return null
  }
}

/**
 * Check whether the currently signed-in token can read a user's OneDrive root.
 * Returns 'accessible' | 'no-access' | 'no-drive' | 'error'.
 */
export async function checkUserDriveAccess(userId: string): Promise<'accessible' | 'no-access' | 'no-drive' | 'error'> {
  try {
    await client().api(`/users/${userId}/drive/root`).get()
    return 'accessible'
  } catch (err) {
    const status = (err as { statusCode?: number }).statusCode
    if (status === 404) return 'no-drive'
    if (status === 403 || status === 401) return 'no-access'
    return 'error'
  }
}

/**
 * Grant the migration account write access to a user's OneDrive.
 *
 * Approach: fetch the user's drive to get its SharePoint site ID, then POST
 * to /sites/{siteId}/permissions with role "write".  This works via
 * Files.ReadWrite.All (to read the drive) + Sites.Manage.All (to write
 * permissions) and does NOT require existing access to the drive root —
 * avoiding the circular dependency of the sharing-invite API.
 */
export async function grantUserDriveAccess(userId: string, migrationAccountEmail: string): Promise<void> {
  // Step 1: get the drive — Files.ReadWrite.All (admin-consented) allows this
  // regardless of whether the migration account already has access.
  const drive = await client()
    .api(`/users/${userId}/drive`)
    .select('id,sharepointIds')
    .get() as GraphDrive

  const siteId = drive.sharepointIds?.siteId
  if (!siteId) {
    throw new Error('Could not determine OneDrive site ID — OneDrive may not be provisioned for this user')
  }

  // Step 2: grant the migration account write access via the site permissions API.
  await client().api(`/sites/${siteId}/permissions`).post({
    roles: ['write'],
    grantedToIdentities: [
      { user: { email: migrationAccountEmail } },
    ],
  })
}

/**
 * Search for a user by display name using exact match first, then a broad search.
 * Returns matched user + status to distinguish single-match vs ambiguous vs not found.
 */
export async function findUserForOneDrive(displayName: string): Promise<{
  user: AppUser | null
  status: 'matched' | 'not-found' | 'ambiguous' | 'error'
  candidates: AppUser[]
}> {
  try {
    // Exact match
    const exact = await client()
      .api('/users')
      .filter(`displayName eq '${displayName.replace(/'/g, "''")}'`)
      .select('id,displayName,mail,userPrincipalName')
      .top(5)
      .get() as { value: GraphUser[] }

    const exactUsers = (exact.value ?? []).map(mapGraphUser)
    if (exactUsers.length === 1) return { user: exactUsers[0], status: 'matched', candidates: [] }
    if (exactUsers.length > 1) return { user: null, status: 'ambiguous', candidates: exactUsers }

    // Broad search using ConsistencyLevel + $search
    try {
      const broad = await client()
        .api('/users')
        .header('ConsistencyLevel', 'eventual')
        .query({ $search: `"displayName:${displayName}"`, $count: 'true' })
        .select('id,displayName,mail,userPrincipalName')
        .top(5)
        .get() as { value: GraphUser[] }

      const broadUsers = (broad.value ?? []).map(mapGraphUser)
      if (broadUsers.length === 0) return { user: null, status: 'not-found', candidates: [] }

      // Prefer case-insensitive exact display name match from broad results
      const nameMatch = broadUsers.find(
        (u) => u.displayName.toLowerCase() === displayName.toLowerCase()
      )
      if (nameMatch) return { user: nameMatch, status: 'matched', candidates: [] }
      if (broadUsers.length === 1) return { user: broadUsers[0], status: 'matched', candidates: [] }
      return { user: null, status: 'ambiguous', candidates: broadUsers }
    } catch {
      // Some tenants don't support $search — fall back to startsWith
      const prefix = displayName.split(' ')[0]
      const fallback = await client()
        .api('/users')
        .filter(`startsWith(displayName,'${prefix.replace(/'/g, "''")}')`)
        .select('id,displayName,mail,userPrincipalName')
        .top(10)
        .get() as { value: GraphUser[] }

      const fallbackUsers = (fallback.value ?? []).map(mapGraphUser)
      const match = fallbackUsers.find(
        (u) => u.displayName.toLowerCase() === displayName.toLowerCase()
      )
      if (match) return { user: match, status: 'matched', candidates: [] }
      if (fallbackUsers.length === 0) return { user: null, status: 'not-found', candidates: [] }
      return { user: null, status: 'ambiguous', candidates: fallbackUsers.slice(0, 5) }
    }
  } catch (err) {
    return { user: null, status: 'error', candidates: [] }
  }
}

function mapGraphUser(u: GraphUser): AppUser {
  return { id: u.id, displayName: u.displayName, mail: u.mail ?? u.userPrincipalName, userPrincipalName: u.userPrincipalName }
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

  // 2. Ensure SPMigration parent exists — GET first to avoid triggering a 409
  let spMigrationExists = false
  try {
    await client().api(`/sites/${siteId}/drive/root:/SPMigration:`).get()
    spMigrationExists = true
  } catch { /* doesn't exist yet */ }
  if (!spMigrationExists) {
    try {
      await client()
        .api(`/sites/${siteId}/drive/root/children`)
        .post({ name: 'SPMigration', folder: {} })
    } catch { /* race condition — another request just created it, continue */ }
  }

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

export async function saveOneDriveMappingsFile(
  siteId: string,
  projectTitle: string,
  projectId: string,
  mappings: OneDriveUserMapping[]
): Promise<void> {
  // Strip children from sourceNode to keep the file lean
  const slim = mappings.map((m) => ({ ...m, sourceNode: { ...m.sourceNode, children: [] } }))
  const token = await getToken()
  const folderName = getProjectFolderName(projectTitle, projectId)
  const filePath = encodeURIComponent(`SPMigration/${folderName}/${projectId}.odmappings.json`)
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
    throw new Error(`OneDrive mappings save failed (${response.status}): ${text}`)
  }
}

export async function loadOneDriveMappingsFile(
  siteId: string,
  projectTitle: string,
  projectId: string
): Promise<OneDriveUserMapping[] | null> {
  const token = await getToken()
  const folderName = getProjectFolderName(projectTitle, projectId)
  const filePath = encodeURIComponent(`SPMigration/${folderName}/${projectId}.odmappings.json`)
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${filePath}:/content`,
    { headers: { Authorization: `Bearer ${token}` } }
  )
  if (response.status === 404) return null
  if (!response.ok) throw new Error(`OneDrive mappings load failed (${response.status})`)
  return JSON.parse(await response.text()) as OneDriveUserMapping[]
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
