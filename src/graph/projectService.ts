import { Client } from '@microsoft/microsoft-graph-client'
import { getToken } from '../auth/authService'
import type { MigrationProject, ProjectData, ProjectStatus, GraphListItem, SharePointUser } from '../types'

// ─── Config ───────────────────────────────────────────────────────────────────
// These are resolved once at runtime from the SharePoint site set up during
// manual admin setup (see plan.md — Manual Setup Instructions).
// Set VITE_SP_SITE_ID and VITE_SP_LIST_ID in .env.local for local dev.
// In production inject via window.__APP_CONFIG__.

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
  data: Pick<MigrationProject, 'title' | 'description' | 'status'> & { owner?: SharePointUser }
): Promise<MigrationProject> {
  const projectData: ProjectData = data.owner ? { owner: data.owner } : {}
  const body = {
    fields: {
      Title: data.title,
      Description: data.description,
      Status: data.status,
      ProjectData: JSON.stringify(projectData),
    },
  }
  const item = await client().api(listItemsUrl()).post(body) as GraphListItem
  return mapItem(item)
}

export async function updateProject(
  id: string,
  fields: Partial<{
    title: string
    description: string
    status: ProjectStatus
    projectData: ProjectData
  }>
): Promise<void> {
  const spFields: Record<string, unknown> = {}
  if (fields.title !== undefined) spFields['Title'] = fields.title
  if (fields.description !== undefined) spFields['Description'] = fields.description
  if (fields.status !== undefined) spFields['Status'] = fields.status
  if (fields.projectData !== undefined) spFields['ProjectData'] = JSON.stringify(fields.projectData)

  await client()
    .api(`${listItemsUrl()}/${id}/fields`)
    .patch(spFields)
}

export async function deleteProject(id: string): Promise<void> {
  await client().api(`${listItemsUrl()}/${id}`).delete()
}

// ─── Mapping ──────────────────────────────────────────────────────────────────

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

  return {
    id: item.id,
    title: f.Title ?? '',
    description: (f.Description as string | undefined) ?? '',
    status: ((f.Status as string | undefined) ?? 'Planning') as ProjectStatus,
    owners: projectData.owner ? [projectData.owner] : [],
    projectData,
    lastModified: f.Modified ? new Date(f.Modified as string) : undefined,
  }
}
