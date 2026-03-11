import { Client } from '@microsoft/microsoft-graph-client'
import { getToken } from '../auth/authService'
import { downloadDriveItem, loadMappingsFile } from './graphClient'
import type { MigrationProject, ProjectData, ProjectStatus, ProjectType, GraphListItem, SharePointUser, TreeNode, MigrationMapping } from '../types'

// ─── Config ───────────────────────────────────────────────────────────────────

interface SPConfig {
  siteId: string
  listId: string
}

export function getSpConfig(): SPConfig {
  const win = window as Window & {
    __APP_CONFIG__?: { spSiteId?: string; spListId?: string }
  }
  return {
    siteId: win.__APP_CONFIG__?.spSiteId ?? (import.meta.env.VITE_SP_SITE_ID as string ?? ''),
    listId: win.__APP_CONFIG__?.spListId ?? (import.meta.env.VITE_SP_LIST_ID as string ?? ''),
  }
}

// ─── Client ───────────────────────────────────────────────────────────────────

function client(): Client {
  return Client.init({
    authProvider: async (done) => {
      try {
        done(null, await getToken())
      } catch (err) {
        done(err as Error, null)
      }
    },
  })
}

function listItemsUrl(): string {
  const { siteId, listId } = getSpConfig()
  return `/sites/${siteId}/lists/${listId}/items`
}

// ─── SP User resolution ───────────────────────────────────────────────────────
//
// SharePoint Person fields require the internal numeric SP user ID, not the AAD
// Object ID.  The only reliable way to get/create this via delegated auth is the
// SharePoint REST ensureUser endpoint, which we call directly with a SharePoint-
// scoped token.  The site webUrl (needed for the REST call) is fetched once and
// cached.

let _spWebUrl: string | null = null

async function getSpWebUrl(): Promise<string> {
  if (_spWebUrl) return _spWebUrl
  const { siteId } = getSpConfig()
  const site = await client()
    .api(`/sites/${siteId}`)
    .select('webUrl')
    .get() as { webUrl: string }
  _spWebUrl = site.webUrl
  return site.webUrl
}

async function resolveSpUserIds(emails: string[]): Promise<string[]> {
  if (emails.length === 0) return []
  try {
    const webUrl = await getSpWebUrl()
    const spHost = new URL(webUrl).origin  // e.g. https://tenant.sharepoint.com
    const spToken = await getToken([`${spHost}/.default`])
    const ids: string[] = []

    for (const email of emails) {
      if (!email) continue
      try {
        const res = await fetch(`${webUrl}/_api/web/ensureuser`, {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${spToken}`,
            'Content-Type': 'application/json;odata=verbose',
            Accept: 'application/json;odata=verbose',
          },
          body: JSON.stringify({ logonName: `i:0#.f|membership|${email}` }),
        })
        if (res.ok) {
          const data = await res.json() as { d: { Id: number } }
          ids.push(String(data.d.Id))
        }
      } catch {
        // Skip user if ensureUser fails for this email
      }
    }
    return ids
  } catch {
    // If we can't get the SP token or site URL, skip writing Owners field
    return []
  }
}

// ─── CRUD ─────────────────────────────────────────────────────────────────────

export async function getProjects(): Promise<MigrationProject[]> {
  const response = await client()
    .api(listItemsUrl())
    .expand('fields')
    .get() as { value: GraphListItem[] }
  return (response.value ?? []).map(mapItem)
}

export async function getProject(id: string): Promise<MigrationProject> {
  const item = await client()
    .api(`${listItemsUrl()}/${id}`)
    .expand('fields')
    .get() as GraphListItem
  return mapItem(item)
}

export async function createProject(
  data: Pick<MigrationProject, 'title' | 'description' | 'status' | 'type'> & { owners?: SharePointUser[] }
): Promise<MigrationProject> {
  const owners = data.owners ?? []

  const spFields: Record<string, unknown> = {
    Title: data.title,
    Description: data.description,
    Status: data.status,
    Type: data.type ?? 'SharePoint',
    ProjectData: JSON.stringify({}),
  }

  // Owners are stored exclusively in the SharePoint Owners person field
  const spUserIds = await resolveSpUserIds(owners.map((o) => o.email))
  if (spUserIds.length > 0) {
    spFields['OwnersLookupId@odata.type'] = 'Collection(Edm.String)'
    spFields['OwnersLookupId'] = spUserIds
  }

  const item = await client().api(listItemsUrl()).post({ fields: spFields }) as GraphListItem
  return mapItem(item)
}

export async function updateProject(
  id: string,
  fields: Partial<{
    title: string
    description: string
    status: ProjectStatus
    type: ProjectType
    projectData: ProjectData
    owners: SharePointUser[]
  }>
): Promise<void> {
  const spFields: Record<string, unknown> = {}
  if (fields.title !== undefined) spFields['Title'] = fields.title
  if (fields.description !== undefined) spFields['Description'] = fields.description
  if (fields.status !== undefined) spFields['Status'] = fields.status
  if (fields.type !== undefined) spFields['Type'] = fields.type
  if (fields.projectData !== undefined) spFields['ProjectData'] = JSON.stringify(fields.projectData)

  // Owners are stored exclusively in the SharePoint Owners person field
  if (fields.owners !== undefined) {
    const spUserIds = await resolveSpUserIds(fields.owners.map((o) => o.email))
    if (spUserIds.length > 0) {
      spFields['OwnersLookupId@odata.type'] = 'Collection(Edm.String)'
      spFields['OwnersLookupId'] = spUserIds
    }
  }

  await client()
    .api(`${listItemsUrl()}/${id}/fields`)
    .patch(spFields)
}

export async function deleteProject(id: string): Promise<void> {
  await client().api(`${listItemsUrl()}/${id}`).delete()
}

// ─── Tree loading ─────────────────────────────────────────────────────────────
//
// Resolves the active TreeNode for a project, handling both the new upload-file
// model (tree stored as .tree.json in SharePoint Documents) and the legacy model
// (tree embedded directly in the ProjectData JSON field).

export async function loadProjectTree(project: MigrationProject): Promise<TreeNode | null> {
  const { uploads, activeUploadId, treeData } = project.projectData

  if (uploads && uploads.length > 0) {
    const { siteId } = getSpConfig()
    const activeId = activeUploadId ?? uploads[uploads.length - 1].id
    const upload = uploads.find((u) => u.id === activeId) ?? uploads[uploads.length - 1]
    try {
      return (await downloadDriveItem(siteId, upload.treeItemId)) as TreeNode
    } catch (err) {
      console.warn('[ProjectService] Could not download tree file:', err)
      return null
    }
  }

  // Legacy: tree data was stored inline in the ProjectData field
  return treeData ?? null
}

// ─── Mappings loading ─────────────────────────────────────────────────────────
//
// For projects that have an SP upload folder, mappings are stored as a separate
// file to avoid hitting the SharePoint list item column size limit.
// Falls back to the inline ProjectData.mappings field for legacy projects.

export async function loadProjectMappings(project: MigrationProject): Promise<MigrationMapping[]> {
  const { uploads, mappings: inlineMappings } = project.projectData

  if (uploads && uploads.length > 0) {
    const { siteId } = getSpConfig()
    try {
      const fileMappings = await loadMappingsFile(siteId, project.title, project.id)
      if (fileMappings !== null) return fileMappings
    } catch (err) {
      console.warn('[ProjectService] Could not load mappings file, falling back to inline:', err)
    }
  }

  return inlineMappings ?? []
}

// ─── Mapping ──────────────────────────────────────────────────────────────────

interface SpLookupValue {
  LookupId?: number
  LookupValue?: string
  Email?: string
}

function mapItem(item: GraphListItem): MigrationProject {
  const f = item.fields
  let projectData: ProjectData = {}
  try {
    if (f.ProjectData && typeof f.ProjectData === 'string') {
      projectData = JSON.parse(f.ProjectData) as ProjectData
    }
  } catch {
    // Corrupt JSON — treat as empty
  }

  // Owners are read exclusively from the SharePoint Owners person field
  const owners: SharePointUser[] = Array.isArray(f.Owners)
    ? (f.Owners as SpLookupValue[]).map((u) => ({
        id: String(u.LookupId ?? ''),
        displayName: u.LookupValue ?? '',
        email: u.Email ?? '',
      }))
    : []

  return {
    id: item.id,
    title: f.Title ?? '',
    description: (f.Description as string | undefined) ?? '',
    status: ((f.Status as string | undefined) ?? 'Planning') as ProjectStatus,
    type: ((f.Type as string | undefined) ?? 'SharePoint') as ProjectType,
    owners,
    projectData,
    lastModified: f.Modified ? new Date(f.Modified as string) : undefined,
  }
}
