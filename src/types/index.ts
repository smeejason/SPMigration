// ─── Auth ────────────────────────────────────────────────────────────────────

export interface AppUser {
  id: string
  displayName: string
  mail: string
  userPrincipalName: string
}

// ─── Projects ────────────────────────────────────────────────────────────────

export type ProjectStatus = 'Planning' | 'In Progress' | 'Completed' | 'On Hold'

export interface MigrationProject {
  id: string                  // SharePoint list item ID
  title: string
  description: string
  status: ProjectStatus
  owners: SharePointUser[]
  projectData: ProjectData
  lastModified?: Date
}

export interface SharePointUser {
  id: string
  displayName: string
  email: string
}

// ─── ProjectData (JSON blob stored in SharePoint) ────────────────────────────

export interface ProjectData {
  sourcePaths?: string[]
  treeData?: TreeNode | null
  mappings?: MigrationMapping[]
  settings?: ProjectSettings
  lastSaved?: string          // ISO date string
}

export interface ProjectSettings {
  defaultLibrary?: string
  exportFormat?: 'csv' | 'json'
}

// ─── TreeSize ────────────────────────────────────────────────────────────────

export interface TreeNode {
  path: string
  name: string
  depth: number
  sizeBytes: number
  fileCount: number
  folderCount: number
  lastModified?: Date
  children: TreeNode[]
}

export interface ParsedTreeSizeRow {
  path: string
  sizeBytes: number
  fileCount: number
  folderCount: number
  percentOfParent?: number
  lastModified?: Date
}

// ─── SharePoint / Graph ──────────────────────────────────────────────────────

export interface SharePointSite {
  id: string
  name: string
  displayName: string
  webUrl: string
  description?: string
}

export interface SharePointDrive {
  id: string
  name: string
  webUrl: string
  driveType: string
}

// ─── Mappings ────────────────────────────────────────────────────────────────

export type MappingStatus = 'pending' | 'ready' | 'error'

export interface MigrationMapping {
  id: string
  sourceNode: TreeNode
  targetSite: SharePointSite | null
  targetDrive: SharePointDrive | null
  targetFolderPath: string
  status: MappingStatus
  notes?: string
}

// ─── Site Creation ───────────────────────────────────────────────────────────

export type SiteTemplate = 'team' | 'communication'
export type SiteCreationStatus = 'pending' | 'creating' | 'created' | 'failed'

export interface SiteRequest {
  id: string
  displayName: string
  alias: string             // URL suffix
  description: string
  template: SiteTemplate
  status: SiteCreationStatus
  createdSite?: SharePointSite
  error?: string
}

// ─── App State ───────────────────────────────────────────────────────────────

export interface AppState {
  auth: {
    user: AppUser | null
    isAuthenticated: boolean
  }
  projects: MigrationProject[]
  currentProject: MigrationProject | null
  treeData: TreeNode | null
  mappings: MigrationMapping[]
  sites: SharePointSite[]
  pendingSiteCreations: SiteRequest[]
  ui: {
    activeView: 'login' | 'projects' | 'project-upload' | 'project-map' | 'project-sites' | 'project-summary'
    loading: boolean
    error: string | null
  }
}

// ─── Graph API raw responses ─────────────────────────────────────────────────

export interface GraphSite {
  id: string
  name: string
  displayName: string
  webUrl: string
  description?: string
}

export interface GraphDrive {
  id: string
  name: string
  webUrl: string
  driveType: string
}

export interface GraphUser {
  id: string
  displayName: string
  mail: string
  userPrincipalName: string
}

export interface GraphListItem {
  id: string
  fields: {
    Title: string
    Description?: string
    Status?: string
    ProjectData?: string
    Owners?: GraphUser[]
    Modified?: string
    [key: string]: unknown
  }
}
