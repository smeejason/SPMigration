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
  siteTypes?: SiteType[]        // reusable site type presets defined on the Site Types tab
  settings?: ProjectSettings
  lastSaved?: string            // ISO date string
  // OneDrive-specific
  autoMapSettings?: AutoMapSettings    // persisted level + account settings from Auto Map tab
  oneDriveMappingCount?: number        // denormalized count of auto-mapped OneDrive users
  // Migration review
  resultUploads?: ResultUpload[]       // SPMT result ZIPs, ordered oldest → newest
  sharePointFeedEnabled?: boolean      // whether the SP live feed is shown on the Review tab
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
  owner?: string        // folder owner from TreeSize export; may be "Access is denied."
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
  owner?: string        // folder owner; may be "Access is denied." when unreadable
}

// ─── OneDrive Auto Map ────────────────────────────────────────────────────────

export type OneDriveMatchStatus = 'pending' | 'matched' | 'not-found' | 'ambiguous' | 'error' | 'cant-find'
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

/** @deprecated Use NewSiteConfig. Kept as alias so legacy persisted data still deserialises. */
export type PlannedSiteTarget = NewSiteConfig

export interface MigrationMapping {
  id: string
  sourceNode: TreeNode
  targetSite: SharePointSite | null
  targetDrive: SharePointDrive | null
  targetFolderPath: string
  status: MappingStatus
  notes?: string
  plannedSite?: NewSiteConfig
  // OneDrive auto-map fields (optional — only set for Phase 1 results)
  matchStatus?: OneDriveMatchStatus
  accessStatus?: OneDriveAccessStatus
  resolvedDisplayName?: string
}

// ─── Site Types ──────────────────────────────────────────────────────────────

export type SiteTemplate = 'team' | 'communication'

/** A lightweight user reference stored inside SiteType presets */
export interface UserRef {
  id: string
  displayName: string
  email: string
}

/** An org-published SharePoint site design (fetched from SP REST API) */
export interface OrgSiteDesign {
  id: string
  title: string
  description?: string
  webTemplate: '64' | '68' | string  // 64 = Team, 68 = Communication
}

/**
 * A reusable site configuration preset defined on the Site Types tab.
 * Used as a starting point when mapping a folder to a new SharePoint site.
 */
export interface SiteType {
  id: string
  name: string                        // e.g. "Department Site"
  template: SiteTemplate
  description?: string
  defaultLibrary?: string
  defaultSubfolder?: string
  siteDesignId?: string               // org site design to apply post-creation
  siteDesignName?: string
  createTeam?: boolean                // team sites only — provision a Teams team too
  owners: UserRef[]
  members: UserRef[]
}

/**
 * Configuration for a new SharePoint site to be created during migration.
 * Replaces PlannedSiteTarget — forward-compatible (same core fields).
 */
export interface NewSiteConfig {
  siteTypeId?: string                 // which SiteType template was used
  siteTypeName?: string
  displayName: string
  alias: string
  description?: string
  template: SiteTemplate
  libraryName?: string
  folderPath?: string
  siteDesignId?: string
  createTeam?: boolean
  owners: UserRef[]
  members: UserRef[]
}

// ─── Site Creation (legacy — kept for Review tab provisioning) ────────────────

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
  reviewData: ReviewData | null
  ui: {
    activeView: 'login' | 'projects' | 'project-upload' | 'project-automap' | 'project-map' | 'project-sites' | 'project-summary' | 'project-review'
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

export interface GraphPerson {
  id: string
  displayName: string
  scoredEmailAddresses?: Array<{ address: string }>
  userPrincipalName?: string
  personType?: { class: string; subclass: string }
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

// ─── Migration Review ─────────────────────────────────────────────────────────

export type MigrationResultStatus = 'Migrated' | 'Failed' | 'Skipped' | 'Partial'

export interface MigrationResultItem {
  source: string              // raw Source value from CSV (UNC path)
  destination: string         // raw Destination URL
  itemName: string            // Item name column
  itemType: 'File' | 'Folder' // Type column
  status: MigrationResultStatus
  resultCategory: string      // Result category column
  message: string             // Message column
  errorCode: string           // Error code (from ItemFailureReport, else '')
  fileSizeBytes: number       // Item size (bytes) column
  isRecycleBin: boolean       // true when Source includes '$RECYCLE.BIN'
  sourcePath: string          // normalized: forward slashes, UNC prefix stripped
}

export interface MigrationResultSummary {
  items: MigrationResultItem[]
  migratedCount: number
  failedCount: number
  skippedCount: number
  partialCount: number
  totalCount: number
}

export interface ResultUpload {
  id: string              // timestamp string used as unique key
  fileName: string        // original ZIP filename
  uploadedAt: string      // ISO datetime string
  zipItemId: string       // Graph driveItem ID of the stored raw ZIP
  summaryItemId: string   // Graph driveItem ID of the stored .result.json
  migratedCount: number
  failedCount: number
  skippedCount: number
  partialCount: number
  totalCount: number
}

export interface ReviewNode {
  path: string
  name: string
  depth: number
  children: ReviewNode[]
  migratedCount: number   // aggregated from all descendants
  failedCount: number
  skippedCount: number
  partialCount: number
  totalCount: number
}

export interface ReviewData {
  tree: ReviewNode
  items: MigrationResultItem[]   // flat list, kept for search and detail panels
  totals: {
    migrated: number
    failed: number
    skipped: number
    partial: number
    total: number
    failedRecycleBin: number
    skippedRecycleBin: number
  }
}
