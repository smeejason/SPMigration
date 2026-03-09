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
