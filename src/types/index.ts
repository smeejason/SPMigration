// ─── Auth ────────────────────────────────────────────────────────────────────

export interface AppUser {
  id: string
  displayName: string
  mail: string
  userPrincipalName: string
}

// ─── Projects ────────────────────────────────────────────────────────────────

export type ProjectStatus = 'Planning' | 'In Progress' | 'Completed' | 'On Hold'
export type ProjectType = 'SharePoint' | 'OneDrive'

export interface MigrationProject {
  id: string                  // SharePoint list item ID
  title: string
  description: string
  status: ProjectStatus
  type: ProjectType
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

export interface ExcelUpload {
  id: string           // timestamp string used as unique ID, e.g. "1741234567890"
  fileName: string     // original filename shown in the UI
  uploadedAt: string   // ISO datetime string
  excelItemId: string  // Graph driveItem ID for the uploaded Excel/CSV file
  treeItemId: string   // Graph driveItem ID for the companion .tree.json file
  // Summary stats stored at upload time so history can display without reloading the tree
  rowCount?: number        // total node count across the whole tree
  topFolderName?: string   // name/path of the root folder
  fileCount?: number       // root node's total file count
  folderCount?: number     // root node's total folder count
  sizeBytes?: number       // root node's total size in bytes
}

export interface ProjectData {
  uploads?: ExcelUpload[]       // ordered oldest → newest; new upload model
  activeUploadId?: string       // which upload is currently active (defaults to last)
  treeData?: TreeNode | null    // LEGACY: pre-upload-history projects store tree here
  mappings?: MigrationMapping[]
  mappingCount?: number         // denormalized count kept in sync with the mappings file
  settings?: ProjectSettings
  lastSaved?: string            // ISO date string
  // OneDrive-specific
  autoMapSettings?: AutoMapSettings    // persisted level + account settings from Auto Map tab
  oneDriveMappingCount?: number        // denormalized count of auto-mapped OneDrive users
}

export interface ProjectSettings {
  defaultLibrary?: string
  exportFormat?: 'csv' | 'json'
}

// ─── TreeSize ────────────────────────────────────────────────────────────────

export interface TreeNode {
  path: string          // normalized internal key (forward slashes, no leading slash)
  originalPath: string  // raw path for display (preserves UNC prefix and backslashes)
  name: string
  depth: number
  sizeBytes: number
  fileCount: number
  folderCount: number
  lastModified?: Date
  lastAccessed?: Date
  children: TreeNode[]
}

export interface ParsedTreeSizeRow {
  path: string          // normalized
  originalPath: string  // raw from source file
  sizeBytes: number
  fileCount: number
  folderCount: number
  percentOfParent?: number
  lastModified?: Date
  lastAccessed?: Date
}

// ─── OneDrive Auto Map ────────────────────────────────────────────────────────

export type OneDriveMatchStatus = 'pending' | 'matched' | 'not-found' | 'ambiguous' | 'error'
export type OneDriveAccessStatus = 'unknown' | 'accessible' | 'granted' | 'no-access' | 'no-drive' | 'error'

export interface OneDriveUserMapping {
  id: string                       // = sourceNode.path (unique key)
  sourceNode: TreeNode             // the user folder in the tree
  folderName: string               // raw folder name, e.g. "MarisaBruan"
  resolvedDisplayName: string      // camelCase-split name, e.g. "Marisa Bruan"
  matchedUser: AppUser | null      // matched M365 user (null until resolved)
  matchStatus: OneDriveMatchStatus
  driveId: string                  // Graph drive ID (once resolved)
  driveWebUrl: string              // user's OneDrive root URL
  accessStatus: OneDriveAccessStatus
  targetFolderPath: string         // destination within their OneDrive, e.g. "Migration/Files"
  error?: string
}

export interface AutoMapSettings {
  selectedLevel: number        // tree depth (0-based) that holds user home-drive folders
  migrationAccount: string     // UPN of the account that will run the migration
  targetFolderPath: string     // folder path within each user's OneDrive (may be empty = root)
}



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

export interface PlannedSiteTarget {
  displayName: string
  alias: string
  description: string
  template: SiteTemplate
  libraryName: string
  folderPath: string
}

export interface MigrationMapping {
  id: string
  sourceNode: TreeNode
  targetSite: SharePointSite | null
  targetDrive: SharePointDrive | null
  targetFolderPath: string
  status: MappingStatus
  notes?: string
  plannedSite?: PlannedSiteTarget
  // OneDrive auto-map fields (optional — only set for Phase 1 results)
  matchStatus?: OneDriveMatchStatus
  accessStatus?: OneDriveAccessStatus
  resolvedDisplayName?: string
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
  oneDriveMappings: OneDriveUserMapping[]
  sites: SharePointSite[]
  pendingSiteCreations: SiteRequest[]
  ui: {
    activeView: 'login' | 'projects' | 'project-upload' | 'project-automap' | 'project-map' | 'project-sites' | 'project-summary'
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
  sharepointIds?: {
    siteId: string
    siteUrl: string
  }
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
    ProjectType?: string
    ProjectData?: string
    Owners?: GraphUser[]
    Modified?: string
    [key: string]: unknown
  }
}
