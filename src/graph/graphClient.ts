import { Client } from '@microsoft/microsoft-graph-client'
import { getToken } from '../auth/authService'
import type {
  AppUser,
  SharePointSite,
  SharePointDrive,
  SiteRequest,
  OrgSiteDesign,
  GraphSite,
  GraphDrive,
  GraphUser,
  GraphPerson,
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

/**
 * Returns the organisation's published SharePoint site designs via the SP REST API.
 * Out-of-box Microsoft templates are excluded so only custom org designs are shown.
 */
export async function getSiteDesigns(): Promise<OrgSiteDesign[]> {
  const { root: rootHost } = await getSharePointHosts()
  const spToken = await getToken([`https://${rootHost}/AllSites.FullControl`])
  const resp = await fetch(
    `https://${rootHost}/_api/Microsoft.SharePoint.Utilities.WebTemplateExtensions.SiteScriptUtility.GetSiteDesigns`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${spToken}`,
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json',
      },
    }
  )
  if (!resp.ok) throw new Error(`getSiteDesigns: ${resp.status} ${resp.statusText}`)
  const json = await resp.json() as { value: Array<{ Id: string; Title: string; Description?: string; WebTemplate: string; IsOutOfBoxTemplate: boolean }> }
  return (json.value ?? [])
    .filter(d => !d.IsOutOfBoxTemplate)
    .map(d => ({ id: d.Id, title: d.Title, description: d.Description, webTemplate: d.WebTemplate }))
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

  // Strategy 1: /users with startsWith (requires User.ReadBasic.All)
  try {
    const response = await client()
      .api('/users')
      .filter(`startsWith(displayName,'${query}') or startsWith(userPrincipalName,'${query}')`)
      .select('id,displayName,mail,userPrincipalName')
      .top(10)
      .get() as { value: GraphUser[] }
    return (response.value ?? []).map(mapGraphUser)
  } catch { /* fall through to People API */ }

  // Strategy 2: /me/people search (requires only People.Read)
  try {
    const response = await client()
      .api('/me/people')
      .query({ $search: `"${query}"` })
      .select('id,displayName,scoredEmailAddresses,userPrincipalName,personType')
      .top(10)
      .get() as { value: GraphPerson[] }
    return (response.value ?? [])
      .filter(p => p.personType?.subclass === 'OrganizationUser')
      .map(mapPersonToUser)
  } catch { return [] }
}

// ─── SharePoint personal site helpers ─────────────────────────────────────────

let _spHosts: { root: string; my: string; admin: string } | null = null

/**
 * Derives the tenant's SharePoint hostnames from /sites/root.
 * root = contoso.sharepoint.com  (used for token scope)
 * my   = contoso-my.sharepoint.com  (used for personal site REST calls)
 * Result is cached for the session.
 */
async function getSharePointHosts(): Promise<{ root: string; my: string; admin: string }> {
  if (_spHosts) return _spHosts
  const site = await client().api('/sites/root').select('webUrl').get() as { webUrl: string }
  const rootHost = new URL(site.webUrl).hostname          // contoso.sharepoint.com
  const tenantName = rootHost.split('.')[0]               // contoso
  _spHosts = {
    root: rootHost,                                       // contoso.sharepoint.com
    my: `${tenantName}-my.sharepoint.com`,               // contoso-my.sharepoint.com
    admin: `${tenantName}-admin.sharepoint.com`,         // contoso-admin.sharepoint.com
  }
  return _spHosts
}

/**
 * Converts a UPN to the path segment used in personal OneDrive site URLs.
 * e.g. "john.doe@contoso.com" → "john_doe_contoso_com"
 */
function formatUpnForPersonalSite(upn: string): string {
  return upn.replace(/[@.]/g, '_').toLowerCase()
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
 * Check whether the admin can access a user's OneDrive.
 * Returns 'accessible' | 'no-access' | 'no-drive' | 'error'.
 *
 * Graph delegated permissions often can't reach personal OneDrives even with
 * Files.ReadWrite.All.  When the Graph call returns 403 we fall back to the
 * SharePoint REST API with AllSites.FullControl to distinguish a missing drive
 * (no-drive) from an existing-but-inaccessible one (no-access).
 */
export async function checkUserDriveAccess(userId: string): Promise<'accessible' | 'no-access' | 'no-drive' | 'error'> {
  // Fast path — Graph API (works when Files.ReadWrite.All is sufficient)
  try {
    await client().api(`/users/${userId}/drive/root`).get()
    return 'accessible'
  } catch (err) {
    const status = (err as { statusCode?: number }).statusCode
    if (status === 404) return 'no-drive'
    if (status !== 403 && status !== 401) return 'error'
  }

  // Fallback — SharePoint REST API with AllSites.FullControl
  // Token scope MUST be the root SharePoint host (contoso.sharepoint.com), not
  // -my — Azure AD issues SharePoint tokens against the root resource even when
  // calling -my.sharepoint.com REST endpoints.
  try {
    const user = await getUserById(userId)
    if (!user?.userPrincipalName) return 'no-access'

    const { root: rootHost, my: myHost } = await getSharePointHosts()
    const sitePath = `/personal/${formatUpnForPersonalSite(user.userPrincipalName)}`
    const spToken = await getToken([`https://${rootHost}/AllSites.FullControl`])

    const resp = await fetch(`https://${myHost}${sitePath}/_api/web`, {
      headers: { Authorization: `Bearer ${spToken}`, Accept: 'application/json' },
    })

    if (resp.ok) return 'accessible'            // site exists, AllSites.FullControl confirmed
    if (resp.status === 404) return 'no-drive'  // personal site never provisioned
    return 'no-access'
  } catch {
    return 'no-access'
  }
}

/**
 * Grant the migration account write access to a user's OneDrive.
 *
 * Uses the SharePoint REST API with AllSites.FullControl — this works
 * regardless of whether the Graph drive endpoints are accessible, and does
 * not suffer from the delegated-permission limitations that cause 403s on
 * the Graph drive API.
 *
 * The migration account is added to the site's Owners group, giving it
 * full-control access needed to run migrations.
 */
export async function grantUserDriveAccess(userId: string, migrationAccountEmail: string): Promise<void> {
  // Step 1: get the user's UPN — required for User Profile lookup.
  const user = await getUserById(userId)
  if (!user?.userPrincipalName) {
    throw new Error('Could not retrieve UPN for user')
  }

  // Step 2: request a SharePoint token scoped to the ROOT host.
  const { root: rootHost, admin: adminHost } = await getSharePointHosts()
  const spToken = await getToken([`https://${rootHost}/AllSites.FullControl`])

  // Step 3: get the actual personal site URL from SharePoint User Profile Service.
  // This is the definitive URL as stored by SharePoint and handles all UPN edge
  // cases (B2B guests, federated users, special characters). An empty PersonalUrl
  // means the user has never provisioned their OneDrive.
  const encodedAccount = encodeURIComponent(`i:0#.f|membership|${user.userPrincipalName}`)
  const profileResp = await fetch(
    `https://${adminHost}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${encodedAccount}'`,
    { headers: { Authorization: `Bearer ${spToken}`, Accept: 'application/json;odata=nometadata' } },
  )
  if (!profileResp.ok) {
    const text = await profileResp.text().catch(() => '')
    throw new Error(`Could not load user profile (${profileResp.status}): ${text}`)
  }
  const profile = await profileResp.json() as { PersonalUrl?: string }
  const personalSiteUrl = profile.PersonalUrl?.replace(/\/$/, '')
  if (!personalSiteUrl) {
    throw new Error('OneDrive has not been provisioned for this user — they need to sign in to OneDrive at least once before access can be granted.')
  }

  // Step 4: use CSOM ProcessQuery to call Tenant.SetSiteAdmin.
  // This is the same mechanism PowerShell's Set-SPOUser uses internally and is
  // more reliable than the SharePoint REST OData endpoint which returns 404.
  // CSOM does not require a form digest — the bearer token is sufficient.
  const csomBody = `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="JavaScript Client" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="1" ObjectPathId="0" /><ObjectPath Id="3" ObjectPathId="2" /><Method Name="SetSiteAdmin" Id="4" ObjectPathId="2"><Parameters><Parameter Type="String">${personalSiteUrl}</Parameter><Parameter Type="String">i:0#.f|membership|${migrationAccountEmail}</Parameter><Parameter Type="Boolean">true</Parameter></Parameters></Method></Actions><ObjectPaths><StaticProperty Id="0" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /><Constructor Id="2" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`

  const csomResp = await fetch(`https://${adminHost}/_vti_bin/client.svc/ProcessQuery`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${spToken}`,
      'Content-Type': 'text/xml',
    },
    body: csomBody,
  })

  if (!csomResp.ok) {
    const text = await csomResp.text().catch(() => '')
    throw new Error(`Failed to grant access (${csomResp.status}): ${text}`)
  }

  // CSOM always returns HTTP 200; errors are embedded in the JSON response array.
  const csomJson = await csomResp.json() as unknown[]
  for (const item of csomJson) {
    if (item && typeof item === 'object' && 'ErrorInfo' in item) {
      const errorInfo = (item as { ErrorInfo: { ErrorMessage?: string } | null }).ErrorInfo
      if (errorInfo?.ErrorMessage) {
        throw new Error(`SetSiteAdmin failed: ${errorInfo.ErrorMessage}`)
      }
    }
  }
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
  const escaped = displayName.replace(/'/g, "''")
  const nameLower = displayName.toLowerCase()

  // ── Strategy 1: exact $filter (requires User.ReadBasic.All) ──────────────
  try {
    const res = await client()
      .api('/users')
      .filter(`displayName eq '${escaped}'`)
      .select('id,displayName,mail,userPrincipalName')
      .top(5)
      .get() as { value: GraphUser[] }

    const users = (res.value ?? []).map(mapGraphUser)
    if (users.length === 1) return { user: users[0], status: 'matched', candidates: [] }
    if (users.length > 1)  return { user: null, status: 'ambiguous', candidates: users }
    // 0 results — fall through to $search
  } catch { /* permission denied or unsupported — try next */ }

  // ── Strategy 2: $search with ConsistencyLevel ─────────────────────────────
  try {
    const res = await client()
      .api('/users')
      .header('ConsistencyLevel', 'eventual')
      .query({ $search: `"displayName:${displayName}"`, $count: 'true' })
      .select('id,displayName,mail,userPrincipalName')
      .top(5)
      .get() as { value: GraphUser[] }

    const users = (res.value ?? []).map(mapGraphUser)
    if (users.length === 0) return { user: null, status: 'not-found', candidates: [] }

    const exact = users.find(u => u.displayName.toLowerCase() === nameLower)
    if (exact) return { user: exact, status: 'matched', candidates: [] }
    if (users.length === 1) return { user: users[0], status: 'matched', candidates: [] }
    return { user: null, status: 'ambiguous', candidates: users }
  } catch { /* unsupported or permission denied — try next */ }

  // ── Strategy 3: startsWith $filter ───────────────────────────────────────
  try {
    const prefix = displayName.split(' ')[0]
    const res = await client()
      .api('/users')
      .filter(`startsWith(displayName,'${prefix.replace(/'/g, "''")}')`)
      .select('id,displayName,mail,userPrincipalName')
      .top(10)
      .get() as { value: GraphUser[] }

    const users = (res.value ?? []).map(mapGraphUser)
    const exact = users.find(u => u.displayName.toLowerCase() === nameLower)
    if (exact) return { user: exact, status: 'matched', candidates: [] }
    if (users.length === 0) return { user: null, status: 'not-found', candidates: [] }
    return { user: null, status: 'ambiguous', candidates: users.slice(0, 5) }
  } catch { /* fall through to People API */ }

  // ── Strategy 4: /me/people search (requires only People.Read) ────────────
  try {
    const res = await client()
      .api('/me/people')
      .query({ $search: `"${displayName}"` })
      .select('id,displayName,scoredEmailAddresses,userPrincipalName,personType')
      .top(10)
      .get() as { value: GraphPerson[] }

    const users = (res.value ?? [])
      .filter(p => p.personType?.subclass === 'OrganizationUser')
      .map(mapPersonToUser)
    const exact = users.find(u => u.displayName.toLowerCase() === nameLower)
    if (exact) return { user: exact, status: 'matched', candidates: [] }
    if (users.length === 0) return { user: null, status: 'not-found', candidates: [] }
    if (users.length === 1) return { user: users[0], status: 'matched', candidates: [] }
    return { user: null, status: 'ambiguous', candidates: users.slice(0, 5) }
  } catch { /* all strategies exhausted */ }

  return { user: null, status: 'error', candidates: [] }
}

function mapGraphUser(u: GraphUser): AppUser {
  return { id: u.id, displayName: u.displayName, mail: u.mail ?? u.userPrincipalName, userPrincipalName: u.userPrincipalName }
}

function mapPersonToUser(p: GraphPerson): AppUser {
  const mail = p.scoredEmailAddresses?.[0]?.address ?? p.userPrincipalName ?? ''
  return { id: p.id, displayName: p.displayName, mail, userPrincipalName: p.userPrincipalName ?? mail }
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

// ─── SharePoint item lookup ───────────────────────────────────────────────────

export interface SpDriveItemDetails {
  name: string
  createdBy: { user: { displayName: string } }
  createdDateTime: string
  lastModifiedBy: { user: { displayName: string } }
  lastModifiedDateTime: string
  listItem?: { fields?: { Title?: string } }
}

let _driveWebUrlCache: Record<string, string> = {}

export async function getDefaultDriveWebUrl(siteId: string): Promise<string> {
  if (_driveWebUrlCache[siteId]) return _driveWebUrlCache[siteId]
  const drive = await client().api(`/sites/${siteId}/drive`).select('webUrl').get() as { webUrl: string }
  _driveWebUrlCache[siteId] = drive.webUrl
  return drive.webUrl
}

export async function getSharePointItemByPath(siteId: string, driveRelativePath: string): Promise<SpDriveItemDetails> {
  const encoded = driveRelativePath.split('/').map(encodeURIComponent).join('/')
  return client()
    .api(`/sites/${siteId}/drive/root:/${encoded}?$select=name,createdBy,createdDateTime,lastModifiedBy,lastModifiedDateTime&$expand=listItem($expand=fields($select=Title))`)
    .get() as Promise<SpDriveItemDetails>
}

/**
 * Resolve any SharePoint or OneDrive URL directly to a drive item.
 * Uses the Graph /shares endpoint so it works regardless of which site/drive
 * the file lives in — no need to know siteId or driveId ahead of time.
 */
export async function resolveSharePointItemByUrl(rawOrEncodedUrl: string): Promise<SpDriveItemDetails> {
  let url: string
  try { url = decodeURIComponent(rawOrEncodedUrl) } catch { url = rawOrEncodedUrl }
  // Graph /shares requires base64url encoding with a 'u!' prefix
  const encoded = 'u!' + btoa(url).replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_')
  return client()
    .api(`/shares/${encoded}/driveItem`)
    .select('name,createdBy,createdDateTime,lastModifiedBy,lastModifiedDateTime')
    .expand('listItem($expand=fields($select=Title))')
    .get() as Promise<SpDriveItemDetails>
}

export async function deleteDriveItem(siteId: string, itemId: string): Promise<void> {
  const token = await getToken()
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${itemId}`,
    { method: 'DELETE', headers: { Authorization: `Bearer ${token}` } }
  )
  // 204 = success, 404 = already gone — both are acceptable
  if (!response.ok && response.status !== 404) {
    throw new Error(`Delete failed (${response.status})`)
  }
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
