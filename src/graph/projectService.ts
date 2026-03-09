import { Client } from '@microsoft/microsoft-graph-client'
import { getToken } from '../auth/authService'
import type { MigrationProject, ProjectData, ProjectStatus, GraphListItem, SharePointUser } from '../types'

// ─── Config ───────────────────────────────────────────────────────────────────

interface SPConfig {
  siteId: string
  listId: string
}

function getSpConfig(): SPConfig {
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
// SharePoint Person fields require the internal SP list item ID from the site's
// "User Information List", not the AAD Object ID.  We query by email to resolve.
// If a user hasn't visited the SP site yet they won't be in the list; in that
// case we fall back to the claims-format identity string so the write still works.

async function resolveSpUserIds(emails: string[]): Promise<string[]> {
  const { siteId } = getSpConfig()
  const ids: string[] = []
  for (const email of emails) {
    if (!email) continue
    try {
      const res = await client()
        .api(`/sites/${siteId}/lists/User Information List/items`)
        .filter(`fields/EMail eq '${email}'`)
        .expand('fields($select=EMail)')
        .top(1)
        .get() as { value: Array<{ id: string }> }
      if (res.value.length > 0) {
        ids.push(res.value[0].id)
      } else {
        // User not yet in the SP site — use claims-format identity string
        ids.push(`i:0#.f|membership|${email}`)
      }
    } catch {
      ids.push(`i:0#.f|membership|${email}`)
    }
  }
  return ids
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
  data: Pick<MigrationProject, 'title' | 'description' | 'status'> & { owners?: SharePointUser[] }
): Promise<MigrationProject> {
  const owners = data.owners ?? []

  const spFields: Record<string, unknown> = {
    Title: data.title,
    Description: data.description,
    Status: data.status,
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
    projectData: ProjectData
    owners: SharePointUser[]
  }>
): Promise<void> {
  const spFields: Record<string, unknown> = {}
  if (fields.title !== undefined) spFields['Title'] = fields.title
  if (fields.description !== undefined) spFields['Description'] = fields.description
  if (fields.status !== undefined) spFields['Status'] = fields.status
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
    owners,
    projectData,
    lastModified: f.Modified ? new Date(f.Modified as string) : undefined,
  }
}
